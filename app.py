import re
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

import preencher_template_despesas_bec as core


def parse_colaboradores_manual(raw: str) -> list[dict]:
    """
    Formato esperado por linha:
    Nome;NIF;ValorBruto
    """
    resultados = []
    for idx, linha in enumerate(raw.splitlines(), start=1):
        l = linha.strip()
        if not l:
            continue
        partes = [p.strip() for p in l.split(";")]
        if len(partes) != 3:
            raise ValueError(f"Linha {idx} invalida. Usa: Nome;NIF;ValorBruto")
        nome, nif, bruto = partes
        if not (nif.isdigit() and len(nif) == 9):
            raise ValueError(f"Linha {idx}: NIF invalido ({nif}).")
        bruto_num = float(bruto.replace(".", "").replace(",", "."))
        resultados.append({"nome": nome, "nif": nif, "valor_bruto": bruto_num})
    if not resultados:
        raise ValueError("Nao foram encontrados colaboradores validos no texto manual.")
    return resultados


ORDEM_MESES = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}


def indice_ordenacao_pasta(p: Path) -> tuple[int, int, str]:
    s = core.slug(p.name)
    n = 99
    for nome, i in ORDEM_MESES.items():
        if nome in s:
            n = i
            break
    ano = 0
    m_ano = re.search(r"(20\d{2})", p.name)
    if m_ano:
        ano = int(m_ano.group(1))
    return (ano, n, p.name.lower())


def pastas_mes_a_partir_de_zip(arquivo: str, conteudo: bytes, area_trabalho: Path) -> list[Path]:
    stem = Path(arquivo).stem
    zip_path = area_trabalho / f"_upload_{stem}.zip"
    zip_path.write_bytes(conteudo)
    extract_root = area_trabalho / f"_ext_{stem}"
    extract_root.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(extract_root)
    finally:
        zip_path.unlink(missing_ok=True)

    subdirs = sorted([p for p in extract_root.iterdir() if p.is_dir()])
    pdfs_raiz = list(extract_root.glob("*.pdf"))

    def tem_pdfs(d: Path) -> bool:
        return any(d.glob("*.pdf"))

    if subdirs:
        out = [d for d in subdirs if tem_pdfs(d)]
        if out:
            return out

    if pdfs_raiz:
        dest = area_trabalho / stem
        dest.mkdir(parents=True, exist_ok=True)
        for f in pdfs_raiz:
            (dest / f.name).write_bytes(f.read_bytes())
        return [dest]

    if len(subdirs) == 1 and tem_pdfs(subdirs[0]):
        return [subdirs[0]]

    return []


st.set_page_config(page_title="Preenchimento Template BEC", layout="centered")
st.title("Preenchimento automatico do TemplateDespesasBEC")
st.write("Carrega o template e os PDFs (ou varios meses em ZIP) para gerar o Excel.")

modo = st.radio(
    "Modo",
    options=["Um mes (PDFs em lista)", "Varios meses (ZIP)"],
    horizontal=True,
    help="Com ZIPs podes processar maio, junho, julho, etc. de uma vez.",
)

st.markdown(
    """
Em cada pasta (ou ZIP), os PDFs devem seguir a nomenclatura:
- `1. Recibo [data].pdf`
- `3. Extrato contabilistico vencimento.pdf` (ou vencimento)
- `7. Folhas de Remuneracao SS.pdf`
- `9. Extrato contabilistico SS.pdf`
- `10. Fatura Seguro AT.pdf`
- `12. Extrato contabilistico Seguro.pdf`
"""
)

template_file = st.file_uploader("Template Excel", type=["xlsx"], accept_multiple_files=False)

pdf_files = None
zip_files = None
nome_mes = "maio"

if modo == "Um mes (PDFs em lista)":
    pdf_files = st.file_uploader("PDFs do mes", type=["pdf"], accept_multiple_files=True)
    nome_mes = st.text_input("Nome da pasta do mes (pasta virtual)", value="maio")
else:
    zip_files = st.file_uploader(
        "Ficheiros ZIP (um ou mais)",
        type=["zip"],
        accept_multiple_files=True,
        help="Opcao A: um ZIP com pastas `5. Maio`, `6. Junho`, etc. Opcao B: varios ZIPs `maio.zip`, `junho.zip` com os PDFs no topo.",
    )
    st.caption("Os meses sao ordenados por nome (Janeiro, Maio, Junho, etc.).")

ocr_lang = st.text_input("Idioma OCR (Tesseract)", value="por")
colab_manual = st.text_area(
    "Fallback manual de colaboradores (opcional)",
    value="",
    height=120,
    help="Se a leitura do recibo falhar, uma linha por colaborador: Nome;NIF;ValorBruto",
    placeholder="Joao Ferreira;205890326;813,46",
)

with st.expander("Fallback manual de campos (opcional)"):
    st.caption("Aplica-se a todos os meses nesta execucao, se preencheres os campos.")
    venc_doc_pagamento = st.text_input("Vencimentos - No Doc Pagamento", value="")
    venc_data_pagamento = st.text_input("Vencimentos - Data Doc Pagamento (dd/mm/aaaa)", value="")

    ss_total = st.text_input("SS - Total", value="")
    ss_declaracao = st.text_input("SS - No Declaracao (Identificador DR)", value="")
    ss_data_doc = st.text_input("SS - Data Doc Despesa (dd/mm/aaaa)", value="")
    ss_data_pagamento = st.text_input("SS - Data Doc Pagamento (dd/mm/aaaa)", value="")

    seg_total = st.text_input("Seguro - Total", value="")
    seg_num_fatura = st.text_input("Seguro - No Fatura", value="")
    seg_data_doc = st.text_input("Seguro - Data Doc Despesa (dd/mm/aaaa)", value="")
    seg_doc_pagamento = st.text_input("Seguro - No Doc Pagamento", value="")
    seg_data_pagamento = st.text_input("Seguro - Data Doc Pagamento (dd/mm/aaaa)", value="")


def montar_campos_override() -> dict:
    campos_override = {}
    if venc_doc_pagamento.strip() and venc_data_pagamento.strip():
        campos_override["venc_doc_pagamento"] = venc_doc_pagamento.strip()
        campos_override["venc_data_pagamento"] = venc_data_pagamento.strip()
    if ss_total.strip() and ss_declaracao.strip() and ss_data_doc.strip():
        campos_override["ss_total"] = ss_total.strip()
        campos_override["ss_declaracao"] = ss_declaracao.strip()
        campos_override["ss_data_doc"] = ss_data_doc.strip()
    if ss_data_pagamento.strip():
        campos_override["ss_data_pagamento"] = ss_data_pagamento.strip()
    if (
        seg_total.strip()
        and seg_num_fatura.strip()
        and seg_data_doc.strip()
        and seg_doc_pagamento.strip()
        and seg_data_pagamento.strip()
    ):
        campos_override["seg_total"] = seg_total.strip()
        campos_override["seg_num_fatura"] = seg_num_fatura.strip()
        campos_override["seg_data_doc"] = seg_data_doc.strip()
        campos_override["seg_doc_pagamento"] = seg_doc_pagamento.strip()
        campos_override["seg_data_pagamento"] = seg_data_pagamento.strip()
    return campos_override


if st.button("Processar e gerar Excel", type="primary"):
    if not template_file:
        st.error("Carrega primeiro o ficheiro template .xlsx.")
        st.stop()

    core.OCR_LANG = ocr_lang
    core.POPPLER_PATH = ""

    colaboradores_override = None
    if colab_manual.strip():
        try:
            colaboradores_override = parse_colaboradores_manual(colab_manual)
        except ValueError as ve:
            st.error(str(ve))
            st.stop()

    campos_override = montar_campos_override()

    with tempfile.TemporaryDirectory() as tmp:
        base = Path(tmp)
        template_path = base / "TemplateDespesasBEC.xlsx"
        template_path.write_bytes(template_file.getbuffer())

        pastas_processar: list[Path] = []

        if modo == "Um mes (PDFs em lista)":
            if not pdf_files:
                st.error("Carrega os PDFs do mes.")
                st.stop()
            pasta_mes = base / nome_mes
            pasta_mes.mkdir(parents=True, exist_ok=True)
            for f in pdf_files:
                (pasta_mes / f.name).write_bytes(f.getbuffer())
            pastas_processar = [pasta_mes]
        else:
            if not zip_files:
                st.error("Carrega um ou mais ficheiros ZIP.")
                st.stop()

            work_zip = base / "zip_work"
            work_zip.mkdir(parents=True, exist_ok=True)
            for f in zip_files:
                novas = pastas_mes_a_partir_de_zip(f.name, f.getbuffer().tobytes(), work_zip)
                pastas_processar.extend(novas)

            if not pastas_processar:
                st.error(
                    "Nenhuma pasta com PDFs encontrada dentro dos ZIPs. Verifica a estrutura das pastas."
                )
                st.stop()

            uniq = []
            visto = set()
            for p in pastas_processar:
                try:
                    k = str(p.resolve())
                except OSError:
                    k = str(p)
                if k not in visto:
                    visto.add(k)
                    uniq.append(p)
            pastas_processar = sorted(uniq, key=indice_ordenacao_pasta)

        meses_lista = ", ".join(p.name for p in pastas_processar)
        st.info(f"Meses a processar ({len(pastas_processar)}): {meses_lista}")

        saida_atual = template_path
        pasta_em_execucao = None
        try:
            for idx, pasta_mes in enumerate(pastas_processar):
                pasta_em_execucao = pasta_mes
                df = core.construir_dataframe_linhas(
                    pasta_mes,
                    colaboradores_override=colaboradores_override,
                    campos_override=campos_override or None,
                )
                out_path = base / f"step_{idx}.xlsx"
                core.preencher_excel(saida_atual, out_path, df)
                saida_atual = out_path

            resultado_final = saida_atual.read_bytes()
        except Exception as exc:
            nome_erro = pasta_em_execucao.name if pasta_em_execucao is not None else "?"
            st.error(f"Erro no processamento ({nome_erro}): {exc}")
            st.stop()

        st.success("Ficheiro gerado com sucesso.")
        st.download_button(
            label="Download do Excel preenchido",
            data=resultado_final,
            file_name="TemplateDespesasBEC_preenchido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
