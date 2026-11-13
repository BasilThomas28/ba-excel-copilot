import streamlit as st
import pandas as pd
import google.genai as genai
import os
import time
from io import BytesIO

# Model fallback list (in order of preference)
MODELS = ["gemini-3.8-flash", "gemini-3.7-flash", "gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]


def column_index_to_letter(index):
    """Convert 0-indexed column number to Excel letter (0→A, 1→B, 26→AA, etc.)."""
    result = ""
    index += 1
    while index > 0:
        index -= 1
        result = chr(65 + (index % 26)) + result
        index //= 26
    return result


def extract_excel_schema(uploaded_file):
    """
    Extract structural metadata from Excel file (first 3 rows only).
    Returns: Dict with sheet info, headers, dtypes, and sample values.
    """
    excel_file = pd.ExcelFile(uploaded_file)
    schema = {}

    for sheet_name in excel_file.sheet_names:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, nrows=3)
            if df.empty:
                continue

            sheet_info = {
                "sheet_name": sheet_name,
                "columns": []
            }

            for col_idx, col_name in enumerate(df.columns):
                col_letter = column_index_to_letter(col_idx)
                col_dtype = str(df[col_name].dtype)

                # Extract up to 3 non-null sample values
                samples = df[col_name].dropna().head(3).tolist()
                samples = [str(v) for v in samples]

                sheet_info["columns"].append({
                    "column_letter": col_letter,
                    "header_name": col_name,
                    "dtype": col_dtype,
                    "samples": samples
                })

            schema[sheet_name] = sheet_info

        except Exception as e:
            st.warning(f"Could not read sheet '{sheet_name}': {e}")

    return schema


def data_quality_check(uploaded_file):
    """
    Perform lightweight data quality checks on all sheets.
    Returns: Dict with quality metrics per sheet.
    """
    excel_file = pd.ExcelFile(uploaded_file)
    quality_report = {}

    for sheet_name in excel_file.sheet_names:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

            quality_report[sheet_name] = {
                "total_rows": len(df),
                "total_cols": len(df.columns),
                "null_counts": df.isnull().sum().to_dict(),
                "duplicate_rows": df.duplicated().sum()
            }
        except Exception as e:
            quality_report[sheet_name] = {"error": str(e)}

    return quality_report


def format_schema_for_prompt(schema):
    """Format extracted schema into a readable string for Gemini prompt."""
    schema_text = "## Excel Workbook Schema\n\n"

    for sheet_name, sheet_info in schema.items():
        schema_text += f"### Sheet: {sheet_name}\n"
        for col in sheet_info["columns"]:
            schema_text += f"- Column {col['column_letter']}: '{col['header_name']}' (Type: {col['dtype']}, Samples: {', '.join(col['samples'])})\n"
        schema_text += "\n"

    return schema_text


def call_gemini_api(schema_text, user_question, mode, tutor_mode, selected_model="Auto (fallback)"):
    """
    Call Google Gemini API with automatic fallback and retry logic.
    Returns: (response_text, model_used) or (None, None) on failure
    """
    # Load API key with priority: st.secrets > environment variable
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
    except (KeyError, FileNotFoundError):
        api_key = None

    if not api_key:
        api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        st.error("❌ API key not found. Add GEMINI_API_KEY to .streamlit/secrets.toml or set GEMINI_API_KEY environment variable.")
        return None, None

    client = genai.Client(api_key=api_key)

    if mode == "formula_builder":
        system_prompt = """You are an Excel formula expert for Business Analysts.
Use ONLY the exact sheet names and column letters from the schema below.
Generate ONLY deterministic Excel formulas. Never calculate data yourself.
Always specify the exact cell where the user should paste the formula (e.g., "Paste into Sheet1!D2").
Do NOT provide multiple options—give one best solution."""

        if tutor_mode:
            system_prompt += """\n\nAfter the formula, add "How it works:" followed by a simple flowchart:
- Use boxes with 2-6 plain words each (no jargon).
- Join boxes with → (for horizontal) or ↓ (for vertical steps).
- For complex formulas, use ↓ with one box per line, max 5 boxes.
- Add one example line with small numbers after the flowchart.
- Put flowchart in a ```text code block to keep spacing aligned.

EXAMPLE FORMAT:
Paste into Worksheet!H2:
```excel
=AVERAGE(Worksheet!C:C)
```
How it works:
```text
[Take ratings in column C] → [Add them up] → [Divide by count] → [Average]
Example: 4, 5, 3 → 12 ÷ 3 = 4
```"""

        full_prompt = f"{system_prompt}\n\n{schema_text}\n\nUser Question: {user_question}"

    elif mode == "project_guide":
        system_prompt = """You are a Business Analyst mentor. Create a 3-stage BA checklist:
1. **Data Cleaning & Prep** - Helper columns, NULL handling, deduplication
2. **Core Analysis & Formulas** - Exact Excel formulas or Pivot Table configs
3. **Visualization & Deliverable** - Chart recommendations and stakeholder framing

Use ONLY the exact sheet names and column letters from the schema.
Generate only deterministic formulas—no calculations."""

        if tutor_mode:
            system_prompt += """\n\nFor each formula, add a simple flowchart (not long explanations):
- Use boxes with 2-6 plain words each (no jargon).
- Join boxes with → (horizontal) or ↓ (vertical steps).
- Max 5 boxes per flowchart.
- Include one small example with numbers.
- Put flowchart in a ```text code block."""

        full_prompt = f"{system_prompt}\n\n{schema_text}\n\nBA Scenario: {user_question}"

    # Determine which models to try
    if selected_model == "Auto (fallback)":
        models_to_try = MODELS
    else:
        models_to_try = [selected_model]

    # Try each model in order
    for model in models_to_try:
        try:
            response = client.models.generate_content(model=model, contents=full_prompt)
            return response.text, model
        except Exception as e:
            error_str = str(e)
            # Retry on 503 (unavailable), 429 (rate limit), or 404 (model not available)
            if "503" in error_str or "UNAVAILABLE" in error_str or "429" in error_str or "RATE_LIMIT" in error_str or "404" in error_str or "NOT_FOUND" in error_str:
                time.sleep(2)
                continue
            # Stop on other errors (400, 403, etc.)
            else:
                st.error(f"❌ Gemini API Error: {e}")
                return None, None

    # All models failed with 503/429
    st.error("🔄 Gemini is busy right now. Please try again in a minute.")
    return None, None


def main():
    st.set_page_config(page_title="BA Excel Copilot", page_icon="🧮", layout="wide")
    st.title("🧮 BA Excel Copilot")
    st.markdown("Generate exact Excel formulas and structure BA projects with AI assistance.")

    # Sidebar Configuration
    with st.sidebar:
        st.header("⚙️ Configuration")

        mode = st.radio(
            "Select Mode:",
            options=["formula_builder", "project_guide"],
            format_func=lambda x: "🧮 Formula Builder" if x == "formula_builder" else "🗺️ Project Guide"
        )

        tutor_mode = st.checkbox("🎓 Tutor Mode (Explain formulas)", value=True)

        st.divider()

        model_options = ["Auto (fallback)"] + MODELS
        selected_model = st.selectbox("🤖 Model:", options=model_options, help="Auto tries each model in order if one is busy")

        st.divider()
        st.caption("**Tutor Mode:** Get flowchart explanations instead of long text.")

    # Main Content
    st.subheader("Step 1: Upload Excel File")
    uploaded_file = st.file_uploader("Choose an .xlsx file", type=["xlsx"])

    if uploaded_file is None:
        st.info("👆 Upload an Excel file to get started.")
        return

    # Extract Schema
    st.subheader("Step 2: File Analysis")
    with st.spinner("Extracting schema..."):
        schema = extract_excel_schema(uploaded_file)
        quality_report = data_quality_check(uploaded_file)

    if not schema:
        st.error("❌ Could not extract schema from the Excel file. Ensure it has headers in row 1.")
        return

    # Display Data Quality Pre-Check
    st.write("**Data Quality Pre-Check:**")
    quality_cols = st.columns(len(quality_report))

    for col, (sheet_name, metrics) in zip(quality_cols, quality_report.items()):
        with col:
            if "error" in metrics:
                st.warning(f"⚠️ {sheet_name}: Error reading sheet")
            else:
                st.metric(f"{sheet_name}", f"Rows: {metrics['total_rows']}")
                nulls_total = sum(metrics['null_counts'].values())
                dups = metrics['duplicate_rows']
                if nulls_total > 0 or dups > 0:
                    st.caption(f"⚠️ {nulls_total} nulls, {dups} duplicates")
                else:
                    st.caption("✅ No issues detected")

    # Display Schema (Expander)
    with st.expander("📊 View Schema Details"):
        schema_text = format_schema_for_prompt(schema)
        st.markdown(schema_text)

    # Step 3: User Input
    st.subheader(f"Step 3: {'Formula Request' if mode == 'formula_builder' else 'BA Scenario'}")

    if mode == "formula_builder":
        user_question = st.text_area(
            "Describe what formula you need (e.g., 'Sum all sales by product category'):",
            height=100,
            placeholder="E.g., Create a formula that counts unique customers from the Customer sheet..."
        )
    else:
        user_question = st.text_area(
            "Describe your BA scenario or business question:",
            height=100,
            placeholder="E.g., I have customer sales data. I need to analyze which products are trending..."
        )

    if st.button(f"Generate {'Formula' if mode == 'formula_builder' else 'BA Workflow'}", type="primary"):
        if not user_question.strip():
            st.warning("⚠️ Please enter a question or description.")
            return

        schema_text = format_schema_for_prompt(schema)

        with st.spinner("🤖 Calling Gemini AI..."):
            response, model_used = call_gemini_api(schema_text, user_question, mode, tutor_mode, selected_model)

        if response:
            st.subheader("✅ Result")
            st.markdown(response)

            if model_used:
                st.caption(f"📡 Generated with: {model_used}")

            # Offer download option
            if mode == "formula_builder":
                st.download_button(
                    label="📋 Copy as Text",
                    data=response,
                    file_name="formula_output.txt",
                    mime="text/plain"
                )
            else:
                st.download_button(
                    label="📋 Copy as Text",
                    data=response,
                    file_name="ba_workflow.txt",
                    mime="text/plain"
                )


if __name__ == "__main__":
    main()
