# BA Excel Copilot 🧮

[![Live demo](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://ba-excel-copilot-basil.streamlit.app/)

**▶ Try it:** open the [live demo](https://ba-excel-copilot-basil.streamlit.app/), upload the [sample sales workbook](https://github.com/BasilThomas28/ba-excel-copilot/raw/main/sample_data/sample_sales.xlsx), and ask *"Total revenue for the East region"*.

A Streamlit web app that turns a plain-English business question into an **exact, copy-pasteable Excel formula** that uses your workbook's real sheet names and column letters. It can also plan a whole Business Analyst project. It runs on Google Gemini.

Built by [Basil Thomas](https://www.linkedin.com/in/basil-thomas-0408on28/), a Data Analyst working with SQL, Power BI and Excel.

## Why I built it
Analysts lose time translating questions like *"total revenue per region for Q3"* into the right `SUMIFS` with the right ranges. This app reads the structure of the workbook, so the formula it suggests matches the actual columns. It never calculates numbers itself. Excel does the maths, so the results are deterministic and auditable.

## Features
- **🧮 Formula Builder**: describe what you need and get one formula plus the exact cell to paste it into.
- **🗺️ Project Guide**: a 3-stage BA plan (Data prep → Analysis & formulas → Visualization & stakeholder story).
- **🎓 Tutor Mode**: a short flowchart that explains how each formula works, with a small worked example.
- **📊 Data quality pre-check**: row counts, blank cells and duplicate rows for every sheet.
- **🤖 Model fallback**: automatically tries the next Gemini model if one is busy or rate-limited.

## Tech stack
Python · Streamlit · pandas · openpyxl · Google GenAI SDK (Gemini)

## Quick start

1. **Get the code**
   ```bash
   git clone https://github.com/BasilThomas28/ba-excel-copilot.git
   cd ba-excel-copilot
   ```

2. **Install dependencies** (Python 3.10+)
   ```bash
   pip install -r requirements.txt
   ```

3. **Add your Gemini API key** (free at [Google AI Studio](https://aistudio.google.com)). Pick one option:
   - **`.streamlit/secrets.toml`** (copy `.streamlit/secrets.toml.example`):
     ```toml
     GEMINI_API_KEY = "your-api-key-here"
     ```
   - **`.env` file** (copy `.env.example`):
     ```
     GEMINI_API_KEY=your-api-key-here
     ```
   - **Environment variable** (PowerShell): `$env:GEMINI_API_KEY = "your-api-key-here"`

   `GOOGLE_API_KEY` is accepted too. Both `secrets.toml` and `.env` are git-ignored, so the key never gets committed.

4. **Run it**
   ```bash
   streamlit run ba_excel_copilot.py
   ```
   It opens at `http://localhost:8501`.

## How to use
1. Upload an `.xlsx` file with headers in row 1.
2. Check the data quality pre-check and, if you like, open **View Schema Details** to see the column letters and types.
3. Choose **Formula Builder** or **Project Guide** in the sidebar, then describe your question.
4. Click **Generate**, then paste the formula into Excel or download the result as text.

### Example
```
Mode: Formula Builder
Question: Total revenue for the "East" region

Paste into Sales!H2:
=SUMIFS(Sales!E2:E501, Sales!C2:C501, "East")
```

## Privacy: what is sent to Gemini
Only the workbook's **structure** is sent: sheet names, column headers, column letters, data types, the row count, and **up to 3 sample values per column**. All other rows stay on your computer. Avoid uploading files whose first rows contain confidential or personal data.

## Project structure
```
ba-excel-copilot/
├── ba_excel_copilot.py              # Streamlit app (single file)
├── requirements.txt                 # Python dependencies
├── .env.example                     # Template for a local .env key file
└── .streamlit/
    └── secrets.toml.example         # Template for Streamlit secrets
```

## Troubleshooting
| Problem | Fix |
|---|---|
| "API key not found" | Set `GEMINI_API_KEY` using one of the options in step 3, then restart Streamlit. |
| "Could not extract schema" | Make sure row 1 holds the headers and that the file is a real `.xlsx`. |
| "Gemini is busy" | You hit a rate limit. Wait a minute and try again, or pick another model in the sidebar. |
| Formula doesn't match my data | Check the column letters under **View Schema Details** and ask a more specific question. |

## Limitations
- `.xlsx` files only (not `.xls` or `.csv`).
- Formulas target Microsoft Excel, not Google Sheets.
- Needs an internet connection for the Gemini API.
