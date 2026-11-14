import streamlit as st
import pandas as pd
import google.genai as genai
import os
import time
from io import BytesIO
from dotenv import load_dotenv

# Load GEMINI_API_KEY / GOOGLE_API_KEY from a local .env file if present
load_dotenv()

# Accepted names for the Gemini API key (first one found wins)
API_KEY_NAMES = ("GEMINI_API_KEY", "GOOGLE_API_KEY")

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


@st.cache_data(show_spinner=False, max_entries=5)
def load_workbook(file_bytes):
    """
    Read every sheet of the workbook once and cache it, so clicking widgets
    doesn't re-read the file. Returns: (dict of sheet_name -> DataFrame, dict of sheet_name -> error).
    """
    sheets, errors = {}, {}
    excel_file = pd.ExcelFile(BytesIO(file_bytes))
    for sheet_name in excel_file.sheet_names:
        try:
            sheets[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name)
        except Exception as e:
            errors[sheet_name] = str(e)
    return sheets, errors


def extract_excel_schema(sheets):
    """
    Extract structural metadata from each sheet.
    Data types are inferred from the whole column, but only the first 3
    non-empty values per column are kept as samples (these are sent to Gemini).
    Returns: Dict with sheet info, headers, dtypes, row range and sample values.
    """
    schema = {}

    for sheet_name, df in sheets.items():
        if df.empty:
            continue

        sheet_info = {
            "sheet_name": sheet_name,
            "data_rows": len(df),
            # Row 1 is the header, so data runs from row 2 to row len(df)+1
            "last_row": len(df) + 1,
            "columns": []
        }

        for col_idx, col_name in enumerate(df.columns):
            samples = [str(v) for v in df[col_name].dropna().head(3).tolist()]

            sheet_info["columns"].append({
                "column_letter": column_index_to_letter(col_idx),
                "header_name": col_name,
                "dtype": str(df[col_name].dtype),
                "samples": samples
            })

        schema[sheet_name] = sheet_info

    return schema


def data_quality_check(sheets, errors):
    """
    Perform lightweight data quality checks on all sheets.
    Returns: Dict with quality metrics per sheet.
    """
    quality_report = {}

    for sheet_name, df in sheets.items():
        quality_report[sheet_name] = {
            "total_rows": len(df),
            "total_cols": len(df.columns),
            "null_counts": {str(k): int(v) for k, v in df.isnull().sum().items()},
            "duplicate_rows": int(df.duplicated().sum())
        }
    for sheet_name, err in errors.items():
        quality_report[sheet_name] = {"error": err}

    return quality_report


def format_schema_for_prompt(schema):
    """Format extracted schema into a readable string for Gemini prompt."""
    schema_text = "## Excel Workbook Schema\n\n"

    for sheet_name, sheet_info in schema.items():
        schema_text += f"### Sheet: {sheet_name}\n"
        schema_text += (
            f"Headers in row 1; data in rows 2-{sheet_info['last_row']} "
            f"({sheet_info['data_rows']} data rows)\n"
        )
        for col in sheet_info["columns"]:
            schema_text += f"- Column {col['column_letter']}: '{col['header_name']}' (Type: {col['dtype']}, Samples: {', '.join(col['samples'])})\n"
        schema_text += "\n"

    return schema_text


def get_api_key():
    """Find the Gemini key: Streamlit secrets first, then environment / .env."""
    for name in API_KEY_NAMES:
        try:
            value = st.secrets.get(name)
        except Exception:  # no secrets.toml present
            value = None
        if value:
            return value
    for name in API_KEY_NAMES:
        value = os.getenv(name)
        if value:
            return value
    return None


@st.cache_resource
def get_client(api_key):
    """Create the Gemini client once and reuse it across requests."""
    return genai.Client(api_key=api_key)


@st.cache_resource
def model_memory():
    """Shared memory of which model worked last and which ones don't exist."""
    return {"last_good": None, "unavailable": set()}


def call_gemini_api(schema_text, user_question, mode, tutor_mode, selected_model="Auto (fallback)"):
    """
    Call Google Gemini API with automatic fallback, streaming the answer.
    Returns: (generator of text chunks, model_used) or (None, None) on failure
    """
    api_key = get_api_key()
    if not api_key:
        st.error(
            "❌ API key not found. Add GEMINI_API_KEY to .streamlit/secrets.toml, "
            "a .env file, or an environment variable (GOOGLE_API_KEY also works)."
        )
        return None, None

    client = get_client(api_key)

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

    # Determine which models to try. In Auto mode, start with the model that
    # worked last time so we don't waste time on unavailable ones.
    if selected_model == "Auto (fallback)":
        memory = model_memory()
        last_good = memory.get("last_good")
        models_to_try = [m for m in MODELS if m not in memory["unavailable"]]
        if last_good in models_to_try:
            models_to_try.remove(last_good)
            models_to_try.insert(0, last_good)
    else:
        memory = None
        models_to_try = [selected_model]

    # Try each model in order
    for model in models_to_try:
        try:
            stream = iter(client.models.generate_content_stream(model=model, contents=full_prompt))
            # Errors (missing model, rate limit) surface on the first chunk
            first_chunk = next(stream, None)
        except Exception as e:
            error_str = str(e)
            if "404" in error_str or "NOT_FOUND" in error_str:
                # Model doesn't exist for this key: skip it now and on later requests
                if memory is not None:
                    memory["unavailable"].add(model)
                continue
            if "503" in error_str or "UNAVAILABLE" in error_str or "429" in error_str or "RESOURCE_EXHAUSTED" in error_str:
                # Busy or rate-limited: short pause, then try the next model
                time.sleep(0.5)
                continue
            # Stop on other errors (400, 403, etc.)
            st.error(f"❌ Gemini API Error: {e}")
            return None, None

        if memory is not None:
            memory["last_good"] = model

        def text_chunks(first=first_chunk, rest=stream):
            if first is not None and first.text:
                yield first.text
            for chunk in rest:
                if chunk.text:
                    yield chunk.text

        return text_chunks(), model

    # All models failed with 503/429/404
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
        st.session_state.pop("result", None)
        st.info("👆 Upload an Excel file to get started.")
        return

    # Extract Schema (workbook is read once and cached by file content)
    st.subheader("Step 2: File Analysis")
    with st.spinner("Extracting schema..."):
        try:
            sheets, sheet_errors = load_workbook(uploaded_file.getvalue())
        except Exception as e:
            st.error(f"❌ Could not open this file as an Excel workbook: {e}")
            return
        schema = extract_excel_schema(sheets)
        quality_report = data_quality_check(sheets, sheet_errors)

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

    streamed_now = False
    if st.button(f"Generate {'Formula' if mode == 'formula_builder' else 'BA Workflow'}", type="primary"):
        if not user_question.strip():
            st.warning("⚠️ Please enter a question or description.")
        else:
            schema_text = format_schema_for_prompt(schema)

            with st.spinner("🤖 Calling Gemini AI..."):
                chunks, model_used = call_gemini_api(schema_text, user_question, mode, tutor_mode, selected_model)

            if chunks:
                st.subheader("✅ Result")
                # Show the answer word by word as Gemini writes it
                try:
                    response = st.write_stream(chunks)
                except Exception as e:
                    st.error(f"❌ The answer was cut off: {e}")
                    response = None
                if isinstance(response, list):
                    response = "".join(str(part) for part in response)

                if response:
                    streamed_now = True
                    # Keep the result in session state so it stays on screen
                    # when the user clicks Download or changes another widget
                    st.session_state["result"] = {
                        "text": response,
                        "model": model_used,
                        "mode": mode,
                        "file": uploaded_file.name,
                    }

    result = st.session_state.get("result")
    if result and result["file"] == uploaded_file.name:
        if not streamed_now:  # already shown while streaming on this run
            st.subheader("✅ Result")
            st.markdown(result["text"])

        if result["model"]:
            st.caption(f"📡 Generated with: {result['model']}")

        st.download_button(
            label="📥 Download as text",
            data=result["text"],
            file_name="formula_output.txt" if result["mode"] == "formula_builder" else "ba_workflow.txt",
            mime="text/plain",
            on_click="ignore",
        )


if __name__ == "__main__":
    main()
