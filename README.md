# FinancePlus IDP - App 01_07

Web app Streamlit/Python per gestione intelligente documenti aziendali:

- upload cartella con sottocartelle;
- OCR PDF/immagini;
- lettura PDF, TXT, CSV, XLSX, DOCX;
- classificazione documentale;
- matching cliente per ragione sociale, P.IVA, codice fiscale e amministratore;
- archivio automatico per cliente/mese/tipologia;
- coda documenti da verificare;
- ricerca azienda;
- download documenti ZIP;
- report PDF riepilogativo.

## Avvio locale

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Su Mac/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy Streamlit Cloud

Caricare su GitHub:

- streamlit_app.py
- requirements.txt
- packages.txt
- .streamlit/config.toml
- README.md

Poi su Streamlit Cloud scegliere:

- Repository: il repository GitHub
- Branch: main
- Main file path: streamlit_app.py

## Nota OCR

Su PC Windows installare anche Tesseract OCR. Su Streamlit Cloud il file packages.txt installa tesseract-ocr e lingua italiana.
