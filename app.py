import tempfile
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


st.set_page_config(page_title="Preenchimento Template BEC", layout="centered")
st.title("Preenchimento automatico do TemplateDespesasBEC")
st.write("Carrega os PDFs do mes e o ficheiro template para gerar o Excel preenchido.")

st.markdown(
    """
Ficheiros esperados:
- `1. Recibo [data].pdf`
- `3. Extrato contabilistico vencimento.pdf`
- `7. Folhas de Remuneracao SS.pdf`
- `9. Extrato contabilistico SS.pdf`
- `10. Fatura Seguro AT.pdf`
- `12. Extrato contabilistico Seguro.pdf`
"""
)

template_file = st.file_uploader("Template Excel", type=["xlsx"], accept_multiple_files=False)
pdf_files = st.file_uploader("PDFs do mes", type=["pdf"], accept_multiple_files=True)
nome_mes = st.text_input("Nome da pasta do mes", value="maio")
ocr_lang = st.text_input("Idioma OCR (Tesseract)", value="por")
colab_manual = st.text_area(
    "Fallback manual de colaboradores (opcional)",
    value="",
    height=120,
    help="Se a leitura do recibo falhar, preenche uma linha por colaborador: Nome;NIF;ValorBruto",
    placeholder="Joao Ferreira;205890326;813,46",
)

if st.button("Processar e gerar Excel", type="primary"):
    if not template_file:
        st.error("Carrega primeiro o ficheiro template .xlsx.")
        st.stop()
    if not pdf_files:
        st.error("Carrega os PDFs do mes.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmp:
        base = Path(tmp)
        pasta_mes = base / nome_mes
        pasta_mes.mkdir(parents=True, exist_ok=True)

        template_path = base / "TemplateDespesasBEC.xlsx"
        template_path.write_bytes(template_file.getbuffer())

        for f in pdf_files:
            destino = pasta_mes / f.name
            destino.write_bytes(f.getbuffer())

        core.OCR_LANG = ocr_lang
        core.POPPLER_PATH = ""

        saida_path = base / "TemplateDespesasBEC_preenchido.xlsx"

        try:
            colaboradores_override = None
            if colab_manual.strip():
                colaboradores_override = parse_colaboradores_manual(colab_manual)
            df = core.construir_dataframe_linhas(pasta_mes, colaboradores_override=colaboradores_override)
            core.preencher_excel(template_path, saida_path, df)
        except Exception as exc:
            st.error(f"Erro no processamento: {exc}")
            st.stop()

        st.success("Ficheiro gerado com sucesso.")
        st.download_button(
            label="Download do Excel preenchido",
            data=saida_path.read_bytes(),
            file_name="TemplateDespesasBEC_preenchido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
