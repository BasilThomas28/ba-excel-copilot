# BA Excel Copilot 🧮

An intelligent Streamlit desktop application that generates exact, copy-pasteable Excel formulas and structures Business Analyst projects using Google Gemini AI.

## Features

- **🧮 Formula Builder** - Get exact Excel formulas for any business question with cell references
- **🗺️ Project Guide** - Receive 3-stage BA workflows: Data Prep → Analysis → Visualization
- **🎓 Tutor Mode** - Line-by-line formula explanations with function comparisons
- **📊 Data Quality Pre-Check** - Automatic detection of nulls, duplicates, and data issues
- **🔒 Privacy-First Schema Extraction** - Reads only first 3 rows for efficiency and privacy
- **✅ Deterministic Output** - Never calculates data; only generates formulas for Excel to execute

## Prerequisites

- Python 3.8 or higher
- Google Gemini API key (free tier available at [Google AI Studio](https://aistudio.google.com))

## Installation

1. Clone or download this project:
   ```bash
   cd BA_Copilot
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up Google API Key** (choose one):

   **Option A: Environment Variable**
   ```bash
   # Windows PowerShell
   $env:GOOGLE_API_KEY = "your-api-key-here"
   
   # Windows Command Prompt
   set GOOGLE_API_KEY=your-api-key-here
   
   # Mac/Linux
   export GOOGLE_API_KEY="your-api-key-here"
   ```

   **Option B: Streamlit Secrets** (persistent)
   - Create `.streamlit/secrets.toml` in the project directory:
     ```toml
     GOOGLE_API_KEY = "your-api-key-here"
     ```

## Running the App

```bash
streamlit run ba_excel_copilot.py
```

The app will open in your browser at `http://localhost:8501`

## How to Use

### 1. Upload an Excel File
- Click **"Choose an .xlsx file"** and select your spreadsheet
- The app extracts schema from the first 3 rows of each sheet

### 2. View Data Quality
- See automatic flagging of nulls, duplicates, and row counts
- Expand **"View Schema Details"** to see exact column names and data types

### 3. Choose a Mode

**Mode 1: 🧮 Formula Builder**
- Describe what formula you need (e.g., "Sum all sales by product")
- Get exact formula with target cell reference
- With Tutor Mode: See why each function was chosen

**Mode 2: 🗺️ Project Guide**
- Describe your BA scenario
- Get 3-stage checklist with exact formulas
- Learn structured BA methodology

### 4. Toggle 🎓 Tutor Mode
- **ON (default)**: Get explanations for each formula/step
- **OFF**: Get concise formulas only

## Example Workflows

### Example 1: Simple Formula
```
Mode: Formula Builder
Question: "Calculate total revenue from Sales sheet, columns A (Product) and D (Revenue)"

Output:
Paste into: Sheet1!E2
Formula: =SUM(D2:D1000)

Tutor explanation: 
- SUM adds all numeric values in a range
- D2:D1000 covers all data rows (row 1 is header)
```

### Example 2: Complex Analysis
```
Mode: Project Guide
Question: "I have customer data with purchase dates and amounts. Create a dashboard showing top 5 products by revenue over time."

Output:
Stage 1: Data Cleaning & Prep
- Ensure dates are formatted as dates (column B)
- Check for duplicate transactions

Stage 2: Analysis & Formulas
- Use PIVOT TABLE with rows=Product, values=SUM(Revenue)
- Filter top 5 using INDEX/MATCH

Stage 3: Visualization
- Recommend column chart showing product vs revenue
- Add slicer for date range analysis
```

## File Structure

```
BA_Copilot/
├── ba_excel_copilot.py      # Main Streamlit app (single file)
├── requirements.txt          # Python dependencies
├── README.md                # This file
└── .streamlit/
    └── secrets.toml         # (Optional) API key configuration
```

## Key Functions

### `extract_excel_schema(uploaded_file)`
- Reads first 3 rows from each sheet
- Extracts headers, data types, and sample values
- Converts column indices to Excel letters (A, B, C... AA, etc.)

### `data_quality_check(uploaded_file)`
- Counts nulls per column
- Detects duplicate rows
- Flags type inconsistencies

### `call_gemini_api(schema_text, user_question, mode, tutor_mode)`
- Sends schema + question to Google Gemini 2.5 Flash
- Generates formulas or BA workflows
- Includes tutor explanations if enabled

## Important Rules

1. **Deterministic Math Only** - AI generates formulas, Excel calculates
2. **Privacy First** - Only first 3 rows read; no data is sent to Google
3. **Real References** - Formulas use actual sheet names and column letters from your file
4. **No Hallucinations** - Tutor mode explains only what's in your data schema

## Troubleshooting

**Error: "Google API key not configured"**
- Set `GOOGLE_API_KEY` environment variable or add to `.streamlit/secrets.toml`
- Restart Streamlit: `streamlit run ba_excel_copilot.py`

**Error: "Could not extract schema from the Excel file"**
- Ensure row 1 contains headers
- Check file is not corrupted
- Try with a simpler Excel file first

**Error: Rate limit exceeded**
- You've hit Google's API rate limits
- Wait a few moments and try again
- Consider upgrading to Gemini API paid tier

**Formulas don't match my data**
- Check that your Excel file headers match what the schema shows
- Verify data types in the "View Schema Details" section
- Re-describe your question more specifically

## Technology Stack

- **Streamlit** - Web app framework for Python
- **Pandas** - Data reading and schema extraction
- **openpyxl** - Excel file handling
- **Google GenAI SDK** - Gemini 2.5 Flash model
- **Python-dotenv** - Environment variable management

## Limitations

- Works with `.xlsx` files only (not `.xls` or `.csv`)
- Schema extraction reads only first 3 rows (by design for privacy)
- Requires internet connection for Gemini API
- Formulas generated are for standard Excel (not Google Sheets, Power BI, etc.)

## Next Steps

1. ✅ Install and run the app
2. 📊 Upload a sample Excel file from your course
3. 🧮 Test Formula Builder with a simple question
4. 🗺️ Test Project Guide with a complex BA scenario
5. 🎓 Toggle Tutor Mode to learn formula reasoning

---

**Questions?** Review the code comments in `ba_excel_copilot.py` for detailed explanations of each function.
