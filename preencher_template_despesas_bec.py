# -*- coding: utf-8 -*-
import argparse
import re
import unicodedata
from copy import copy
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl import load_workbook

try:
    from pdf2image import convert_from_path
except Exception:
    convert_from_path = None

try:
    import pytesseract
except Exception:
    pytesseract = None


CATEGORIA_CUSTO = "2.3.3 - Apoios diretos a contratacao"
RUBRICA_REMUN = "632 - Remuneracoes do pessoal"
RUBRICA_SS = "635 - Encargos sobre remuneracoes"
RUBRICA_SEG = "636 - Seguros de acidentes no trabalho e doencas profissionais"
NIF_SS_FIXO = "505305500"
OCR_LANG = "por"
POPPLER_PATH = ""


def slug(txt: str) -> str:
    if txt is None:
        return ""
    t = unicodedata.normalize("NFKD", str(txt))
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = t.lower()
    t = re.sub(r"[^a-z0-9]+", " ", t).strip()
    return t


def ler_pdf_texto(pdf_path: Path) -> str:
    partes = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ""
            if texto:
                partes.append(texto)
    return "\n".join(partes)


def ler_pdf_texto_com_ocr(pdf_path: Path, dpi: int = 300, lang: str = "por") -> str:
    if convert_from_path is None or pytesseract is None:
        raise RuntimeError(
            "Fallback OCR indisponivel. Instale as dependencias: pip install pdf2image pytesseract"
        )

    try:
        kwargs = {"dpi": dpi}
        if POPPLER_PATH:
            kwargs["poppler_path"] = POPPLER_PATH
        imagens = convert_from_path(str(pdf_path), **kwargs)
    except Exception as exc:
        raise RuntimeError(
            "Nao foi possivel converter PDF para imagem. "
            "No Windows, instale o Poppler e adicione 'bin' ao PATH."
        ) from exc

    textos = []
    for img in imagens:
        txt = pytesseract.image_to_string(img, lang=lang) or ""
        if txt.strip():
            textos.append(txt)
    return "\n".join(textos)


def ler_pdf_texto_auto(pdf_path: Path, ocr_lang: str = OCR_LANG) -> str:
    texto = ler_pdf_texto(pdf_path)
    if texto.strip():
        return texto
    return ler_pdf_texto_com_ocr(pdf_path, lang=ocr_lang)


def normalizar_numero_pt(valor_str: str) -> float:
    limpo = re.sub(r"[^\d,.\-]", "", valor_str).replace(".", "").replace(",", ".")
    return float(limpo)


def normalizar_data(data_str: str) -> datetime:
    s = data_str.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"Nao foi possivel converter data: {data_str}")


def garantir_float(v) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    return normalizar_numero_pt(str(v))


def garantir_data(v) -> datetime:
    if isinstance(v, datetime):
        return v
    return normalizar_data(str(v))


def descobrir_mes_ref(nome_pasta: str) -> str:
    meses = {
        "janeiro": "01",
        "fevereiro": "02",
        "marco": "03",
        "marco": "03",
        "abril": "04",
        "maio": "05",
        "junho": "06",
        "julho": "07",
        "agosto": "08",
        "setembro": "09",
        "outubro": "10",
        "novembro": "11",
        "dezembro": "12",
    }
    base = slug(nome_pasta)
    for nome, numero in meses.items():
        if nome in base:
            return numero
    raise ValueError(f"Nao foi possivel inferir mes pela pasta: {nome_pasta}")


def extrair_colaboradores_recibo(recibo_pdf: Path) -> list[dict]:
    texto = ler_pdf_texto_auto(recibo_pdf, OCR_LANG)
    if not texto.strip():
        raise ValueError(
            f"O ficheiro '{recibo_pdf.name}' nao contem texto legivel para extracao."
        )

    padrao = re.compile(
        r"Nome[:\s]+(?P<nome>.+?)\n.*?(?:NIF|N\.?\s*Contribuinte)[:\s]+(?P<nif>\d{9}).*?"
        r"(?:Iliquido|Remuneracao Bruta|Vencimento Bruto|Total Bruto)[:\s]+(?P<bruto>[\d\.\s]+,\d{2})",
        re.IGNORECASE | re.DOTALL,
    )
    resultados = []
    for m in padrao.finditer(texto):
        resultados.append(
            {
                "nome": re.sub(r"\s+", " ", m.group("nome")).strip(" -"),
                "nif": m.group("nif"),
                "valor_bruto": normalizar_numero_pt(m.group("bruto")),
            }
        )
    if not resultados:
        raise ValueError("Nao foi possivel extrair colaboradores do recibo.")
    return resultados


def extrair_doc_pagamento_vencimento(extrato_pdf: Path, mes_ref: str) -> tuple[str, datetime]:
    texto = ler_pdf_texto(extrato_pdf)
    padrao_mes = re.compile(
        rf"(\d{{2}}/\d{{2}}/\d{{4}})\s+([A-Z]{{2}}\d+)\s+Processamento\s+Sal\S*\s+20\d{{2}}\.\s*{mes_ref}",
        re.IGNORECASE,
    )
    m = padrao_mes.search(texto)
    if not m:
        m = re.search(
            r"(\d{2}/\d{2}/\d{4})\s+([A-Z]{2}\d+)\s+Processamento\s+Sal\S*",
            texto,
            flags=re.IGNORECASE,
        )
    if not m:
        raise ValueError("Nao foi possivel extrair No/Data Doc Pagamento de vencimentos.")
    return m.group(2), normalizar_data(m.group(1))


def extrair_ss_folhas(ss_pdf: Path) -> tuple[float, str, datetime]:
    texto = ler_pdf_texto(ss_pdf)
    decl = re.search(r"Identificador\s+DR\s+(\d+)", texto, flags=re.IGNORECASE)
    if not decl:
        raise ValueError("Nao foi possivel extrair o No da Declaracao da SS.")

    total_valor = None
    # Caso mais direto: "Total de contribuicoes: 16941,86"
    total_direto = re.search(
        r"Total de contribui[a-zA-Z]+[:\s]+([\d\.\s]+,\d{2})",
        texto,
        flags=re.IGNORECASE,
    )
    if total_direto:
        total_valor = normalizar_numero_pt(total_direto.group(1))

    # Fallback: linha com dois valores (remuneracoes e contribuicoes) -> usar ultimo valor.
    if total_valor is None:
        linha_resumo = re.search(
            r"Total de Remunera[a-zA-Z]+/Contribui[a-zA-Z]+[^\n]*",
            texto,
            flags=re.IGNORECASE,
        )
        if linha_resumo:
            valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha_resumo.group(0))
            if valores:
                total_valor = normalizar_numero_pt(valores[-1])

    # Fallback final: procurar apos "Total de Contribui..." e apanhar o primeiro valor monetario.
    if total_valor is None:
        bloco = re.search(
            r"Total de Contribui[a-zA-Z]+[\s\S]{0,120}?(\d{1,3}(?:\.\d{3})*,\d{2})",
            texto,
            flags=re.IGNORECASE,
        )
        if bloco:
            total_valor = normalizar_numero_pt(bloco.group(1))

    if total_valor is None:
        raise ValueError("Nao foi possivel extrair o valor total da SS.")

    data = re.search(r"Data de entrega\s+(\d{4}-\d{2}-\d{2})", texto, flags=re.IGNORECASE)
    if not data:
        raise ValueError("Nao foi possivel extrair a data da declaracao SS.")

    return total_valor, decl.group(1), normalizar_data(data.group(1))


def extrair_data_valor_ss(extrato_ss_pdf: Path) -> datetime:
    texto = ler_pdf_texto(extrato_ss_pdf)
    dv = re.search(r"Data[-\s]*valor[:\s]+(\d{2}[/-]\d{2}[/-]\d{4})", texto, flags=re.IGNORECASE)
    if dv:
        return normalizar_data(dv.group(1))
    qualquer = re.search(r"(\d{2}/\d{2}/\d{4})\s+[A-Z]{2}\d+", texto)
    if not qualquer:
        raise ValueError("Nao foi possivel extrair Data Doc Pagamento da SS.")
    return normalizar_data(qualquer.group(1))


def extrair_seguro(fatura_pdf: Path, extrato_seg_pdf: Path) -> tuple[float, str, datetime, str, datetime]:
    texto_fat = ler_pdf_texto(fatura_pdf)
    numero_fat = None
    valor_fat = None
    data_fat = None

    if texto_fat.strip():
        numero = re.search(r"(FCT\d+|FT\s*\d+/\d+|Fatura\s*[Nn]?[o0]?\s*[:\-]?\s*[\w/-]+)", texto_fat)
        if numero:
            numero_fat = re.sub(r"\s+", " ", numero.group(1)).strip()
        total = re.search(r"(?:Total a pagar|Total)\s*([\d\.\s]+,\d{2})", texto_fat, flags=re.IGNORECASE)
        if total:
            valor_fat = normalizar_numero_pt(total.group(1))
        data = re.search(r"(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})", texto_fat)
        if data:
            data_fat = normalizar_data(data.group(1))

    texto_ext = ler_pdf_texto(extrato_seg_pdf)
    linha_fatura = re.search(
        r"(\d{2}/\d{2}/\d{4})\s+([A-Z]{2}\d+)\s+Fatura\s+(FCT\d+)\s+([0-9\.\s]+,\d{2})",
        texto_ext,
        flags=re.IGNORECASE,
    )
    if linha_fatura:
        data_pag = normalizar_data(linha_fatura.group(1))
        doc_pag = linha_fatura.group(2)
        if numero_fat is None:
            numero_fat = linha_fatura.group(3)
        if valor_fat is None:
            valor_fat = normalizar_numero_pt(linha_fatura.group(4))
    else:
        # fallback: encontra linha de movimento de seguro e usa o valor monetario da linha
        lp = re.search(
            r"(\d{2}/\d{2}/\d{4})\s+([A-Z]{2}\d+)\s+SEGURO DE ACIDENTES",
            texto_ext,
            flags=re.IGNORECASE,
        )
        if not lp:
            raise ValueError("Nao foi possivel confirmar pagamento do Seguro no extrato.")
        data_pag = normalizar_data(lp.group(1))
        doc_pag = lp.group(2)
        if valor_fat is None:
            linha_ini = max(0, lp.start() - 10)
            linha_fim = texto_ext.find("\n", lp.end())
            if linha_fim == -1:
                linha_fim = min(len(texto_ext), lp.end() + 120)
            linha = texto_ext[linha_ini:linha_fim]
            vals = re.findall(r"(\d{1,3}(?:[\.\s]\d{3})*,\d{2})", linha)
            if vals:
                valor_fat = normalizar_numero_pt(vals[-1])

    if valor_fat is None:
        raise ValueError("Nao foi possivel extrair o valor da fatura de Seguro AT.")
    if numero_fat is None:
        numero_fat = "Fatura Seguro AT"
    if data_fat is None:
        data_fat = data_pag

    return valor_fat, numero_fat, data_fat, doc_pag, data_pag


def encontrar_pasta_mes(base_dir: Path, nome_mes: str) -> Path:
    candidatas = [base_dir / nome_mes, base_dir / nome_mes.capitalize()]
    for c in candidatas:
        if c.exists() and c.is_dir():
            return c
    for d in base_dir.iterdir():
        if d.is_dir() and slug(nome_mes) in slug(d.name):
            return d
    raise FileNotFoundError(f"Nao encontrei a pasta do mes '{nome_mes}'.")


def mapear_ficheiros(pasta_mes: Path) -> dict:
    mapa = {}
    for f in pasta_mes.glob("*.pdf"):
        n = slug(f.name)
        if n.startswith("1 recibo"):
            mapa["recibo"] = f
        elif n.startswith("3 extrato") and "vencimento" in n:
            mapa["extrato_venc"] = f
        elif n.startswith("7 folhas"):
            mapa["folhas_ss"] = f
        elif n.startswith("9 extrato") and "ss" in n:
            mapa["extrato_ss"] = f
        elif n.startswith("10 fatura"):
            mapa["fatura_seg"] = f
        elif n.startswith("12 extrato") and "seguro" in n:
            mapa["extrato_seg"] = f

    obrigatorios = ["recibo", "extrato_venc", "folhas_ss", "extrato_ss", "fatura_seg", "extrato_seg"]
    falta = [k for k in obrigatorios if k not in mapa]
    if falta:
        raise FileNotFoundError(f"Ficheiros em falta: {', '.join(falta)}")
    return mapa


def construir_dataframe_linhas(
    pasta_mes: Path,
    colaboradores_override: list[dict] | None = None,
    campos_override: dict | None = None,
) -> pd.DataFrame:
    ficheiros = mapear_ficheiros(pasta_mes)
    campos_override = campos_override or {}
    mes_ref = descobrir_mes_ref(pasta_mes.name)
    ano_m = re.search(r"20\d{2}", pasta_mes.name)
    ano_ref = ano_m.group(0) if ano_m else str(datetime.now().year)

    colaboradores = colaboradores_override if colaboradores_override is not None else extrair_colaboradores_recibo(
        ficheiros["recibo"]
    )
    if campos_override.get("venc_doc_pagamento") and campos_override.get("venc_data_pagamento"):
        doc_pag_venc = str(campos_override["venc_doc_pagamento"])
        data_pag_venc = garantir_data(campos_override["venc_data_pagamento"])
    else:
        doc_pag_venc, data_pag_venc = extrair_doc_pagamento_vencimento(ficheiros["extrato_venc"], mes_ref)

    if (
        campos_override.get("ss_total") is not None
        and campos_override.get("ss_declaracao")
        and campos_override.get("ss_data_doc")
    ):
        total_ss = garantir_float(campos_override["ss_total"])
        no_decl_ss = str(campos_override["ss_declaracao"])
        data_doc_ss = garantir_data(campos_override["ss_data_doc"])
    else:
        total_ss, no_decl_ss, data_doc_ss = extrair_ss_folhas(ficheiros["folhas_ss"])

    if campos_override.get("ss_data_pagamento"):
        data_pag_ss = garantir_data(campos_override["ss_data_pagamento"])
    else:
        data_pag_ss = extrair_data_valor_ss(ficheiros["extrato_ss"])

    if (
        campos_override.get("seg_total") is not None
        and campos_override.get("seg_num_fatura")
        and campos_override.get("seg_data_doc")
        and campos_override.get("seg_doc_pagamento")
        and campos_override.get("seg_data_pagamento")
    ):
        total_seg = garantir_float(campos_override["seg_total"])
        no_fat_seg = str(campos_override["seg_num_fatura"])
        data_doc_seg = garantir_data(campos_override["seg_data_doc"])
        no_pag_seg = str(campos_override["seg_doc_pagamento"])
        data_pag_seg = garantir_data(campos_override["seg_data_pagamento"])
    else:
        total_seg, no_fat_seg, data_doc_seg, no_pag_seg, data_pag_seg = extrair_seguro(
            ficheiros["fatura_seg"], ficheiros["extrato_seg"]
        )

    linhas = []
    for c in colaboradores:
        linhas.append(
            {
                "categoria custo": CATEGORIA_CUSTO,
                "doc despesa": "Remuneracoes",
                "descricao": f"Recibo de vencimento {c['nome']} - {pasta_mes.name}",
                "data doc despesa": data_pag_venc,
                "n doc despesa": "Recibo Vencimento",
                "nif fornecedor": c["nif"],
                "nome fornecedor": c["nome"],
                "pais fornecedor": "Portugal",
                "total doc despesa": c["valor_bruto"],
                "mapa de investimentos": 1,
                "rubrica": RUBRICA_REMUN,
                "imputado doc despesa": c["valor_bruto"],
                "elegivel doc despesa": c["valor_bruto"],
                "doc pagamento": "Extrato Bancario",
                "n doc pagamento": doc_pag_venc,
                "data doc pagamento": data_pag_venc,
                "total doc pagamento": c["valor_bruto"],
                "imputado doc pagamento": c["valor_bruto"],
                "elegivel doc pagamento": c["valor_bruto"],
            }
        )

    linhas.append(
        {
            "categoria custo": CATEGORIA_CUSTO,
            "doc despesa": "Contribuicoes Seguranca Social",
            "descricao": f"TSU {pasta_mes.name}",
            "data doc despesa": data_doc_ss,
            "n doc despesa": f"DR {no_decl_ss}",
            "nif fornecedor": NIF_SS_FIXO,
            "nome fornecedor": "Seguranca Social, I.P.",
            "pais fornecedor": "Portugal",
            "total doc despesa": total_ss,
            "mapa de investimentos": 1,
            "rubrica": RUBRICA_SS,
            "imputado doc despesa": total_ss,
            "elegivel doc despesa": total_ss,
            "doc pagamento": "Extrato Bancario",
            "n doc pagamento": f"Pag SS {ano_ref}-{mes_ref}",
            "data doc pagamento": data_pag_ss,
            "total doc pagamento": total_ss,
            "imputado doc pagamento": total_ss,
            "elegivel doc pagamento": total_ss,
        }
    )

    linhas.append(
        {
            "categoria custo": CATEGORIA_CUSTO,
            "doc despesa": "Seguro de acidentes de Trabalho",
            "descricao": f"Seguro AT {pasta_mes.name}",
            "data doc despesa": data_doc_seg,
            "n doc despesa": no_fat_seg,
            "nif fornecedor": "",
            "nome fornecedor": "Seguro Acidentes Trabalho",
            "pais fornecedor": "Portugal",
            "total doc despesa": total_seg,
            "mapa de investimentos": 1,
            "rubrica": RUBRICA_SEG,
            "imputado doc despesa": total_seg,
            "elegivel doc despesa": total_seg,
            "doc pagamento": "Extrato Bancario",
            "n doc pagamento": no_pag_seg,
            "data doc pagamento": data_pag_seg,
            "total doc pagamento": total_seg,
            "imputado doc pagamento": total_seg,
            "elegivel doc pagamento": total_seg,
        }
    )

    return pd.DataFrame(linhas)


def copiar_estilo_linha(ws, linha_origem: int, linha_destino: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        c1 = ws.cell(row=linha_origem, column=col)
        c2 = ws.cell(row=linha_destino, column=col)
        c2._style = copy(c1._style)
        c2.number_format = c1.number_format
        c2.alignment = copy(c1.alignment)
        c2.font = copy(c1.font)
        c2.fill = copy(c1.fill)
        c2.border = copy(c1.border)
        c2.protection = copy(c1.protection)


def preencher_excel(template_excel: Path, saida_excel: Path, df_linhas: pd.DataFrame) -> None:
    wb = load_workbook(template_excel)
    ws = wb["Despesas"]

    cabecalho = [c.value for c in ws[1]]
    idx_col = {}
    for i, nome in enumerate(cabecalho, start=1):
        if nome is not None:
            idx_col[slug(str(nome))] = i

    ultima = ws.max_row
    ordem = ws.cell(row=ultima, column=1).value or 0

    for _, row in df_linhas.iterrows():
        nova = ws.max_row + 1
        copiar_estilo_linha(ws, ultima, nova, ws.max_column)
        ws.cell(row=nova, column=idx_col["ordem"], value=ordem + 1)
        ordem += 1
        for k, v in row.items():
            if k in idx_col:
                ws.cell(row=nova, column=idx_col[k], value=v)

    wb.save(saida_excel)


def main() -> None:
    global OCR_LANG, POPPLER_PATH
    parser = argparse.ArgumentParser(description="Preencher TemplateDespesasBEC com PDFs do mes.")
    parser.add_argument("--base-dir", default=".", help="Diretorio base com o template e pasta mensal.")
    parser.add_argument("--mes", default="maio", help="Nome da pasta do mes.")
    parser.add_argument("--template", default="TemplateDespesasBEC.xlsx", help="Ficheiro template.")
    parser.add_argument(
        "--output",
        default="TemplateDespesasBEC_preenchido_maio.xlsx",
        help="Ficheiro de saida.",
    )
    parser.add_argument(
        "--ocr-lang",
        default="por",
        help="Idioma OCR do Tesseract (ex: por, eng, por+eng).",
    )
    parser.add_argument(
        "--tesseract-cmd",
        default="",
        help="Caminho para tesseract.exe (opcional, para Windows).",
    )
    parser.add_argument(
        "--poppler-path",
        default="",
        help="Caminho para pasta bin do Poppler (opcional, para Windows).",
    )
    args = parser.parse_args()
    OCR_LANG = args.ocr_lang
    POPPLER_PATH = args.poppler_path

    if args.tesseract_cmd and pytesseract is not None:
        pytesseract.pytesseract.tesseract_cmd = args.tesseract_cmd

    base_dir = Path(args.base_dir).resolve()
    pasta_mes = encontrar_pasta_mes(base_dir, args.mes)
    template_excel = base_dir / args.template
    output_excel = base_dir / args.output

    if not template_excel.exists():
        raise FileNotFoundError(f"Template nao encontrado: {template_excel}")

    df = construir_dataframe_linhas(pasta_mes)
    preencher_excel(template_excel, output_excel, df)
    print(f"Concluido: {output_excel}")


if __name__ == "__main__":
    main()
