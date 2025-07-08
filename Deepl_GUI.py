import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import io
import deepl

# --- Page config ---
st.set_page_config(page_title="DeepL Excel Translator", layout="centered")
st.title("📄 DeepL Excel Translator (Side-by-Side In-Place)")

# --- API key ---
DEEPL_API_KEY = "aada8622-5380-f658-5142-0c006cc21976"

# --- Upload file ---
uploaded_file = st.file_uploader("📁 Upload your Excel file (.xlsx or .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    # Read row 2 as header using pandas to extract column names
    df_headers = pd.read_excel(uploaded_file, header=1, nrows=0, engine="openpyxl")  # header=1 → row 2
    column_names = list(df_headers.columns)

    # Define expected columns
    master_labels = ["Title - Master", "Product Description - Master", "Backend KW"]
    master_labels += [f"Bullet Point {i} - Master" for i in range(1, 10)]
    valid_columns = [col for col in column_names if col in master_labels]

    if not valid_columns:
        st.warning("No 'Master' columns found in row 2.")
    else:
        st.success("✅ Found the following Master Content columns:")
        st.write(valid_columns)

        # --- Language selection ---
        source_lang = st.text_input("🌍 Enter source language (e.g., FR)", "FR")
        target_lang = st.text_input("🌍 Enter target language (e.g., NL)", "NL")

        # Show buttons
        st.markdown("### ✏️ Choose columns to translate:")
        col_buttons = {}
        for col in valid_columns:
            col_buttons[col] = st.button(f"Translate: {col}")

        translate_all = st.button("🌍 Translate All Content")

        # Load workbook from uploaded file
        in_memory_file = io.BytesIO(uploaded_file.getbuffer())
        wb = openpyxl.load_workbook(in_memory_file)
        ws = wb.active  # assumes translation is on the first sheet

        def get_col_index(header_row, target_name):
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=header_row, column=col).value == target_name:
                    return col
            return None

        def translate_column(col_name):
            st.write(f"🔁 Translating: **{col_name}**")
            col_idx = get_col_index(2, col_name)  # headers are in row 2
            if not col_idx:
                st.error(f"Column '{col_name}' not found.")
                return

            translator = deepl.Translator(DEEPL_API_KEY)

            for row in range(3, ws.max_row + 1):  # content starts at row 3
                source_cell = ws.cell(row=row, column=col_idx)
                target_cell = ws.cell(row=row, column=col_idx + 1)

                if source_cell.value and not target_cell.value:
                    try:
                        result = translator.translate_text(str(source_cell.value), source_lang=source_lang, target_lang=target_lang)
                        target_cell.value = result.text
                    except Exception as e:
                        target_cell.value = f"ERROR: {e}"

            # Write translated header
            translated_header = f"{col_name} (translated)"
            ws.cell(row=2, column=col_idx + 1, value=translated_header)

        # Handle per-column buttons
        for col in valid_columns:
            if col_buttons[col]:
                translate_column(col)

        # Handle "Translate All"
        if translate_all:
            for col in valid_columns:
                translate_column(col)

        # Check if any translated headers exist
        has_translations = any("(translated)" in str(ws.cell(row=2, column=col).value) for col in range(1, ws.max_column + 1))

        # Prepare download
        if has_translations:
            output = io.BytesIO()
            wb.save(output)
            st.download_button(
                label="📥 Download Translated File",
                data=output.getvalue(),
                file_name="translated_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
