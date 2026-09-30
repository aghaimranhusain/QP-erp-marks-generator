# QP → ERP Marks Upload Excel Generator

Streamlit app that converts a university question paper (**Word .docx**) into an ERP-ready marks upload Excel file.

## Features

- Upload question paper as **.docx**
- Auto-detect questions, marks, COs and Bloom levels
- Editable table to add/fix questions
- Generates Excel in the exact format required by university ERP systems
- Dropdowns for:
  - CO mapping: **Strong-H / Moderate-M / Weak-L**
  - Question Type & Optional: **Mandatory / Optional**
  - Bloom Level: **L1–L6**

## How to run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Push this repo to GitHub (public)
2. Go to https://share.streamlit.io
3. New app → select this repository → Main file: `app.py`
4. Deploy

## Usage

1. Save your question paper as `.docx`
2. Upload it in the app
3. Click **Extract Questions from Document**
4. Review / edit the question table
5. Generate and download the ERP Excel

## License

MIT
