# App web para preencher TemplateDespesasBEC

Esta app permite carregar os PDFs de um mes e gerar automaticamente o `TemplateDespesasBEC.xlsx` preenchido.

## 1) Publicar no Streamlit Cloud

1. Colocar esta pasta num repositorio GitHub.
2. Abrir [Streamlit Cloud](https://share.streamlit.io/).
3. Criar app nova e escolher:
   - Branch: a tua branch principal
   - Main file path: `app.py`
4. Fazer deploy.

## 2) Dependencias

- Python packages em `requirements.txt`
- Pacotes de sistema em `packages.txt`:
  - `poppler-utils`
  - `tesseract-ocr`
  - `tesseract-ocr-por`

## 3) Como usar

1. Carregar o ficheiro template `.xlsx`.
2. Carregar os PDFs do mes:
   - `1. Recibo ...pdf`
   - `3. Extrato contabilistico vencimento...pdf`
   - `7. Folhas de Remuneracao SS.pdf`
   - `9. Extrato contabilistico SS.pdf`
   - `10. Fatura Seguro AT...pdf`
   - `12. Extrato contabilistico Seguro.pdf`
3. Confirmar nome do mes (ex: `maio`).
4. Clicar em **Processar e gerar Excel**.
5. Fazer download do ficheiro final.

## 4) Execucao local (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
```
