# App web para preencher TemplateDespesasBEC

Esta app permite carregar os PDFs de um mes (ou varios meses por ZIP) e gerar automaticamente o `TemplateDespesasBEC.xlsx` preenchido.

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
2. Escolher um modo:
   - `Um mes (PDFs em lista)`, ou
   - `Varios meses (ZIP)`.
3. Se for um mes, carregar os PDFs:
   - `1. Recibo ...pdf`
   - `3. Extrato contabilistico vencimento...pdf`
   - `7. Folhas de Remuneracao SS.pdf`
   - `9. Extrato contabilistico SS.pdf`
   - `10. Fatura Seguro AT...pdf`
   - `12. Extrato contabilistico Seguro.pdf`
4. Se for varios meses, carregar um ou mais ZIP:
   - opcao A: 1 ZIP com subpastas (`5. Maio`, `6. Junho`, `7. Julho`, ...)
   - opcao B: varios ZIPs (`maio.zip`, `junho.zip`, `julho.zip`) com PDFs na raiz
5. Clicar em **Processar e gerar Excel**.
6. Fazer download do ficheiro final.

## 4) Execucao local (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
```
