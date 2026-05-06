# -*- coding: utf-8 -*-
import argparse
import calendar
import re
import unicodedata
from copy import copy
from datetime import datetime, timedelta
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
NIF_GENERALI_FIXO = "500940231"
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


def _is_yellow_color(color) -> bool:
    if not color:
        return False
    try:
        # RGB normalizado [0..1]
        if isinstance(color, tuple) and len(color) >= 3:
            r, g, b = float(color[0]), float(color[1]), float(color[2])
            return r >= 0.75 and g >= 0.7 and b <= 0.45
        # CMYK [0..1] (amarelo costuma ter Y alto e C/M/K baixos)
        if isinstance(color, tuple) and len(color) == 4:
            c, m, y, k = [float(v) for v in color]
            return y >= 0.45 and c <= 0.35 and m <= 0.35 and k <= 0.35
    except Exception:
        return False
    return False


def extrair_valor_destacado_amarelo(pdf_path: Path) -> float | None:
    """
    Tenta ler o valor monetario na linha destacada a amarelo no PDF.
    Usa objetos graficos (rects) + palavras sobrepostas.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            candidatos: list[float] = []
            for page in pdf.pages:
                rects = page.rects or []
                words = page.extract_words() or []
                if not rects or not words:
                    continue

                for r in rects:
                    if not _is_yellow_color(r.get("non_stroking_color")):
                        continue

                    # Expande um pouco para apanhar a linha toda.
                    rx0 = float(r.get("x0", 0)) - 8
                    rx1 = float(r.get("x1", 0)) + 8
                    rtop = float(r.get("top", 0)) - 3
                    rbot = float(r.get("bottom", 0)) + 3

                    linha_words = []
                    for w in words:
                        wx0, wx1 = float(w["x0"]), float(w["x1"])
                        wtop, wbot = float(w["top"]), float(w["bottom"])
                        sobrepoe_x = wx1 >= rx0 and wx0 <= rx1
                        sobrepoe_y = wbot >= rtop and wtop <= rbot
                        if sobrepoe_x and sobrepoe_y:
                            linha_words.append(w)

                    if not linha_words:
                        continue

                    linha_words.sort(key=lambda w: (w["top"], w["x0"]))
                    linha_txt = " ".join(w["text"] for w in linha_words)
                    vals = re.findall(r"(\d{1,3}(?:[\.\s]\d{3})*,\d{2})", linha_txt)
                    for v in vals:
                        n = normalizar_numero_pt(v)
                        if 1 <= n <= 2000:
                            candidatos.append(n)

            if candidatos:
                # Em caso de varios matches amarelos, prefere o ultimo da leitura.
                return round(candidatos[-1], 2)
    except Exception:
        return None
    return None


def extrair_valor_destacado_amarelo_com_texto(pdf_path: Path, texto_alvo: str) -> float | None:
    """
    Extrai valor da linha destacada a amarelo que contenha o texto alvo.
    Ex.: "Joao Ferreira" no documento 12.
    """
    alvo_slug = slug(texto_alvo)
    if not alvo_slug:
        return None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            candidatos: list[float] = []
            for page in pdf.pages:
                rects = page.rects or []
                words = page.extract_words() or []
                if not rects or not words:
                    continue

                for r in rects:
                    if not _is_yellow_color(r.get("non_stroking_color")):
                        continue

                    rx0 = float(r.get("x0", 0)) - 8
                    rx1 = float(r.get("x1", 0)) + 8
                    rtop = float(r.get("top", 0)) - 3
                    rbot = float(r.get("bottom", 0)) + 3

                    linha_words = []
                    for w in words:
                        wx0, wx1 = float(w["x0"]), float(w["x1"])
                        wtop, wbot = float(w["top"]), float(w["bottom"])
                        sobrepoe_x = wx1 >= rx0 and wx0 <= rx1
                        sobrepoe_y = wbot >= rtop and wtop <= rbot
                        if sobrepoe_x and sobrepoe_y:
                            linha_words.append(w)

                    if not linha_words:
                        continue

                    linha_words.sort(key=lambda w: (w["top"], w["x0"]))
                    linha_txt = " ".join(w["text"] for w in linha_words)
                    if alvo_slug not in slug(linha_txt):
                        continue

                    vals = re.findall(r"(\d{1,3}(?:[\.\s]\d{3})*,\d{2})", linha_txt)
                    for v in vals:
                        n = normalizar_numero_pt(v)
                        if 1 <= n <= 2000:
                            candidatos.append(n)
            if candidatos:
                return round(candidatos[-1], 2)
    except Exception:
        return None
    return None


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


def nome_mes_pt(mes: int) -> str:
    nomes = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Marco",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }
    return nomes.get(mes, str(mes))


def inferir_ano_mes(pasta_mes: Path, ficheiros: dict) -> tuple[int, int]:
    # Prioridade: nome do recibo com formato 2025.05
    recibo_nome = ficheiros.get("recibo").name if ficheiros.get("recibo") else ""
    m = re.search(r"(20\d{2})[.\-/](\d{2})", recibo_nome)
    if m:
        return int(m.group(1)), int(m.group(2))

    # Fallback: pasta
    ano_m = re.search(r"20\d{2}", pasta_mes.name)
    mes_ref = descobrir_mes_ref(pasta_mes.name)
    ano = int(ano_m.group(0)) if ano_m else datetime.now().year
    return ano, int(mes_ref)


def ultimo_dia_util_mes(ano: int, mes: int) -> datetime:
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    dt = datetime(ano, mes, ultimo_dia)
    while dt.weekday() >= 5:  # sabado/domingo
        dt -= timedelta(days=1)
    return dt


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
    return deduplicar_colaboradores(resultados)


def deduplicar_colaboradores(colaboradores: list[dict]) -> list[dict]:
    unicos: dict[tuple[str, str], dict] = {}
    for c in colaboradores:
        chave = (str(c.get("nif", "")).strip(), slug(str(c.get("nome", ""))))
        if chave not in unicos:
            unicos[chave] = c
            continue
        # Em caso de OCR duplicado, mantem o maior valor bruto.
        if float(c.get("valor_bruto", 0) or 0) > float(unicos[chave].get("valor_bruto", 0) or 0):
            unicos[chave] = c
    return list(unicos.values())


def extrair_doc_pagamento_vencimento(extrato_pdf: Path, mes_ref: str) -> tuple[str, datetime]:
    texto = ler_pdf_texto(extrato_pdf)
    # Prioridade: referencia no formato "Banco - 2025/005"
    ref_banco = re.search(
        r"((?:Millenium|Milenium|Novo\s*Banco|Santander|CGD)\s*-\s*\d{1,4}/\d{4})",
        texto,
        flags=re.IGNORECASE,
    )
    data_qualquer = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
    if ref_banco and data_qualquer:
        return re.sub(r"\s+", " ", ref_banco.group(1)).strip(), normalizar_data(data_qualquer.group(1))

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
    texto_slug = slug(texto)
    decl = re.search(r"Identificador\s+DR\s+(\d+)", texto, flags=re.IGNORECASE)

    total_valor = None
    # Caso mais direto: "Total de contribuicoes: 16941,86"
    total_direto = re.search(
        r"Total de contribui[^\n:]*[:\s]+([\d\.\s]+,\d{2})",
        texto,
        flags=re.IGNORECASE,
    )
    if total_direto:
        total_valor = normalizar_numero_pt(total_direto.group(1))

    # Versao normalizada sem acentos/simbolos (mais robusta em PDFs com encoding degradado)
    if total_valor is None:
        total_direto_slug = re.search(
            r"total de contribu\w+\s+(\d{1,3}(?:\.\d{3})*,\d{2})",
            texto_slug,
            flags=re.IGNORECASE,
        )
        if total_direto_slug:
            total_valor = normalizar_numero_pt(total_direto_slug.group(1))

    # Fallback: linha com dois valores (remuneracoes e contribuicoes) -> usar ultimo valor.
    if total_valor is None:
        linha_resumo = re.search(
            r"Total de Remunera[^\n]*Contribui[^\n]*",
            texto,
            flags=re.IGNORECASE,
        )
        if linha_resumo:
            valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha_resumo.group(0))
            if valores:
                total_valor = normalizar_numero_pt(valores[-1])

    # Fallback adicional: linha de estabelecimento/ano-mes com dois montantes.
    if total_valor is None:
        linha_mes = re.search(
            r"\b20\d{2}[-/]\d{2}\b[^\n]*",
            texto,
            flags=re.IGNORECASE,
        )
        if linha_mes:
            valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha_mes.group(0))
            if len(valores) >= 2:
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

    # Ultimo fallback: no bloco "extrato de resumo", usa a linha com "contribui" e o ultimo montante.
    if total_valor is None:
        resumo_ini = texto_slug.find("extrato de resumo")
        if resumo_ini != -1:
            resumo = texto[resumo_ini : min(len(texto), resumo_ini + 2500)]
            linha_contrib = re.search(r"[^\n]*contribui[^\n]*", resumo, flags=re.IGNORECASE)
            if linha_contrib:
                valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha_contrib.group(0))
                if valores:
                    total_valor = normalizar_numero_pt(valores[-1])

    if total_valor is None:
        raise ValueError("Nao foi possivel extrair o valor total da SS.")

    data = re.search(r"Data de entrega\s+(\d{4}-\d{2}-\d{2})", texto, flags=re.IGNORECASE)
    if not data:
        raise ValueError("Nao foi possivel extrair a data da declaracao SS.")

    # Preferencia de No Doc Despesa no formato do exemplo: DMR MM/AAAA.
    ano_mes = re.search(r"Ano/M\S+s de refer\S+ncia[:\s]+(20\d{2})[/-](\d{2})", texto, flags=re.IGNORECASE)
    if not ano_mes:
        ano_mes = re.search(r"\b(20\d{2})[-/](\d{2})\b", texto)

    if ano_mes:
        doc_despesa = f"DMR {ano_mes.group(2)}/{ano_mes.group(1)}"
    elif decl:
        doc_despesa = f"DR {decl.group(1)}"
    else:
        doc_despesa = "DMR"

    return total_valor, doc_despesa, normalizar_data(data.group(1))


def extrair_data_valor_ss(extrato_ss_pdf: Path) -> datetime:
    texto = ler_pdf_texto(extrato_ss_pdf)
    dv = re.search(r"Data[-\s]*valor[:\s]+(\d{2}[/-]\d{2}[/-]\d{4})", texto, flags=re.IGNORECASE)
    if dv:
        return normalizar_data(dv.group(1))
    qualquer = re.search(r"(\d{2}/\d{2}/\d{4})\s+[A-Z]{2}\d+", texto)
    if not qualquer:
        raise ValueError("Nao foi possivel extrair Data Doc Pagamento da SS.")
    return normalizar_data(qualquer.group(1))


def extrair_doc_pagamento_ss(extrato_ss_pdf: Path) -> str:
    texto = ler_pdf_texto(extrato_ss_pdf)
    ref_banco = re.search(
        r"((?:Millenium|Milenium|Novo\s*Banco|Santander|CGD)\s*-\s*\d{1,4}/\d{4})",
        texto,
        flags=re.IGNORECASE,
    )
    if ref_banco:
        return re.sub(r"\s+", " ", ref_banco.group(1)).strip()
    m = re.search(r"\d{2}/\d{2}/\d{4}\s+([A-Za-z0-9][A-Za-z0-9/\-\. ]{2,40})", texto)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()
    return "Pagamento SS"


def extrair_imputado_ss_extrato(extrato_ss_pdf: Path, ano_ref: str | None = None, mes_ref: str | None = None) -> float | None:
    # Regra unica: valor da linha destacada a amarelo no documento 9.
    val_amarelo = extrair_valor_destacado_amarelo(extrato_ss_pdf)
    if val_amarelo is not None:
        return val_amarelo
    return None


def calcular_imputado_ss_11(ss_pdf: Path) -> float | None:
    texto = ler_pdf_texto(ss_pdf)
    # Captura linhas do tipo "Remuneracao X ... 123,45".
    valores = re.findall(r"Remunera\w+\s+[A-Z]\b[^\n]*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto, flags=re.IGNORECASE)
    if not valores:
        return None
    total_remun = sum(normalizar_numero_pt(v) for v in valores)
    return round(total_remun * 0.11, 2)


def extrair_imputado_seguro_colaborador(extrato_seg_pdf: Path, nome_colaborador: str) -> float | None:
    texto = ler_pdf_texto(extrato_seg_pdf)
    if not texto.strip() or not nome_colaborador.strip():
        return None

    # Fallback principal em texto corrido (quando o PDF perde quebras de linha).
    texto_slug = slug(texto)
    padrao_texto_corrido = re.search(
        r"afetacao\s+seg\w*\s+at\s+\d{2}\s+20\d{2}\s+joao\s+ferreira[\s\S]{0,60}?(\d{1,3}(?:[\.\s]\d{3})*,\d{2})",
        texto_slug,
        flags=re.IGNORECASE,
    )
    if padrao_texto_corrido:
        return normalizar_numero_pt(padrao_texto_corrido.group(1))

    # Regra hard: linha "Afetacao ... Joao Ferreira" no documento 12.
    alvo_nome = slug(nome_colaborador)
    for linha in texto.splitlines():
        s = slug(linha)
        if "afetacao" in s and "seg" in s and alvo_nome in s:
            vals = re.findall(r"(\d{1,3}(?:[\.\s]\d{3})*,\d{2})", linha)
            if vals:
                nums = [normalizar_numero_pt(v) for v in vals]
                candidatos = [n for n in nums if 1 <= n <= 2000]
                if candidatos:
                    return candidatos[-1]

    # Regra hard adicional para OCR degradado: apenas tokens joao + ferreira.
    for linha in texto.splitlines():
        s = slug(linha)
        if "afetacao" in s and "seg" in s and "joao" in s and "ferreira" in s:
            vals = re.findall(r"(\d{1,3}(?:[\.\s]\d{3})*,\d{2})", linha)
            if vals:
                nums = [normalizar_numero_pt(v) for v in vals]
                candidatos = [n for n in nums if 1 <= n <= 2000]
                if candidatos:
                    return candidatos[-1]

    alvo = slug(nome_colaborador)
    alvo_tokens = [t for t in alvo.split() if t]
    linhas = texto.splitlines()

    # Regra unica: linha de "Afetacao ... <nome>" no documento 12.
    for linha in linhas:
        s = slug(linha)
        if "afetacao" not in s or "seg" not in s:
            continue
        nome_bate = (alvo in s) if alvo else False
        if not nome_bate and alvo_tokens:
            nome_bate = all(tok in s for tok in alvo_tokens)
        if not nome_bate:
            continue
        vals = re.findall(r"(\d{1,3}(?:[\.\s]\d{3})*,\d{2})", linha)
        if vals:
            nums = [normalizar_numero_pt(v) for v in vals]
            candidatos = [n for n in nums if 1 <= n <= 2000]
            if candidatos:
                return candidatos[-1]
    return None


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
    ref_banco = re.search(
        r"((?:Millenium|Milenium|Novo\s*Banco|Santander|CGD)\s*-\s*\d{1,4}/\d{4})",
        texto_ext,
        flags=re.IGNORECASE,
    )
    linha_fatura = re.search(
        r"(\d{2}/\d{2}/\d{4})\s+([A-Z]{2}\d+)\s+Fatura\s+(FCT\d+)\s+([0-9\.\s]+,\d{2})",
        texto_ext,
        flags=re.IGNORECASE,
    )
    if linha_fatura:
        data_pag = normalizar_data(linha_fatura.group(1))
        doc_pag = re.sub(r"\s+", " ", ref_banco.group(1)).strip() if ref_banco else linha_fatura.group(2)
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
        doc_pag = re.sub(r"\s+", " ", ref_banco.group(1)).strip() if ref_banco else lp.group(2)
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
    ano_int, mes_int = inferir_ano_mes(pasta_mes, ficheiros)
    mes_ref = f"{mes_int:02d}"
    ano_ref = str(ano_int)
    mes_label = f"{nome_mes_pt(mes_int)} {ano_ref}"
    data_recibo = ultimo_dia_util_mes(ano_int, mes_int)

    colaboradores = (
        deduplicar_colaboradores(colaboradores_override)
        if colaboradores_override is not None
        else extrair_colaboradores_recibo(ficheiros["recibo"])
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
    if campos_override.get("ss_doc_pagamento"):
        no_pag_ss = str(campos_override["ss_doc_pagamento"])
    else:
        no_pag_ss = extrair_doc_pagamento_ss(ficheiros["extrato_ss"])

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

    # Imputado SS: regra de negocio -> tentar ler do doc 9 primeiro.
    imputado_ss = extrair_imputado_ss_extrato(ficheiros["extrato_ss"], ano_ref=ano_ref, mes_ref=mes_ref)
    if imputado_ss is None:
        if campos_override.get("ss_imputado") is not None:
            imputado_ss = garantir_float(campos_override["ss_imputado"])
        else:
            raise ValueError(
                "Nao foi possivel extrair 'SS imputado' do documento 9. "
                "Preenche o campo manual 'SS - Imputado'."
            )

    # Imputado Seguro: tentar extrair por nome no extrato (doc 12); fallback manual.
    if campos_override.get("seg_imputado") is not None:
        imputado_seg = garantir_float(campos_override["seg_imputado"])
    else:
        nome_ref = colaboradores[0]["nome"] if colaboradores else ""
        imputado_seg = extrair_imputado_seguro_colaborador(ficheiros["extrato_seg"], nome_ref)
        if imputado_seg is None:
            raise ValueError(
                "Nao foi possivel extrair 'Seguro imputado' (afetacao do colaborador) do documento 12. "
                "Preenche o campo manual 'Seguro - Imputado'."
            )

    linhas = []
    for c in colaboradores:
        bruto = float(c["valor_bruto"])
        subsidio_ref = float(c.get("valor_subsidio_refeicao", 0) or 0)
        imputado_remun = round(bruto - subsidio_ref, 2)
        linhas.append(
            {
                "categoria custo": CATEGORIA_CUSTO,
                "doc despesa": "Remuneracoes",
                "descricao": f"Recibo de vencimento {c['nome']} - {mes_label}",
                "data doc despesa": data_recibo,
                "n doc despesa": "Recibo Vencimento",
                "nif fornecedor": c["nif"],
                "nome fornecedor": c["nome"],
                "pais fornecedor": "Portugal",
                "total doc despesa": bruto,
                "mapa de investimentos": 1,
                "rubrica": RUBRICA_REMUN,
                "imputado doc despesa": imputado_remun,
                "elegivel doc despesa": imputado_remun,
                "doc pagamento": "Extrato Bancario",
                "n doc pagamento": doc_pag_venc,
                "data doc pagamento": data_pag_venc,
                "total doc pagamento": bruto,
                "imputado doc pagamento": imputado_remun,
                "elegivel doc pagamento": imputado_remun,
            }
        )

    linhas.append(
        {
            "categoria custo": CATEGORIA_CUSTO,
            "doc despesa": "Contribuicoes Seguranca Social",
            "descricao": f"TSU {mes_label}",
            "data doc despesa": data_doc_ss,
            "n doc despesa": f"DR {no_decl_ss}",
            "nif fornecedor": NIF_SS_FIXO,
            "nome fornecedor": "Seguranca Social, I.P.",
            "pais fornecedor": "Portugal",
            "total doc despesa": total_ss,
            "mapa de investimentos": 1,
            "rubrica": RUBRICA_SS,
            "imputado doc despesa": imputado_ss,
            "elegivel doc despesa": imputado_ss,
            "doc pagamento": "Extrato Bancario",
            "n doc pagamento": no_pag_ss,
            "data doc pagamento": data_pag_ss,
            "total doc pagamento": total_ss,
            "imputado doc pagamento": imputado_ss,
            "elegivel doc pagamento": imputado_ss,
        }
    )

    linhas.append(
        {
            "categoria custo": CATEGORIA_CUSTO,
            "doc despesa": "Seguro de acidentes de Trabalho",
            "descricao": f"Seguro AT {mes_label}",
            "data doc despesa": data_pag_seg,
            "n doc despesa": no_fat_seg,
            "nif fornecedor": NIF_GENERALI_FIXO,
            "nome fornecedor": "Seguro Acidentes Trabalho",
            "pais fornecedor": "Portugal",
            "total doc despesa": total_seg,
            "mapa de investimentos": 1,
            "rubrica": RUBRICA_SEG,
            "imputado doc despesa": imputado_seg,
            "elegivel doc despesa": imputado_seg,
            "doc pagamento": "Extrato Bancario",
            "n doc pagamento": no_pag_seg,
            "data doc pagamento": data_pag_seg,
            "total doc pagamento": total_seg,
            "imputado doc pagamento": imputado_seg,
            "elegivel doc pagamento": imputado_seg,
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
