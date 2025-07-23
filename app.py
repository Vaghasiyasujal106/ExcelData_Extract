import streamlit as st
import pandas as pd
import json
from io import BytesIO

st.set_page_config(page_title=" Allotment Excel Extractor", layout="wide")
st.title("Allotment Excel to JSON Extractor")

uploaded_file = st.file_uploader("Upload Excel or CSV", type=["xlsx", "xls", "csv"])

def format_cell(val):
    if pd.isna(val) or str(val).strip().lower() in ["nan", "none"]:
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d-%m-%Y")
    return str(val).strip()

def is_potential_table_header(row):
    # Checks if row looks like a table header
    row_text = [format_cell(cell).lower() for cell in row]
    return (
        any("name" in cell for cell in row_text) and
        any("no" in cell for cell in row_text) and
        len([c for c in row_text if c]) >= 3
    )

def extract_excel(file):
    try:
        file.seek(0)
        filename = file.name.lower()

        if filename.endswith(".csv"):
            df = pd.read_csv(file, header=None, engine="python", on_bad_lines='skip')
        else:
            df = pd.read_excel(file, header=None)

        df.fillna("", inplace=True)

        header_data = {}
        for i, row in df.iterrows():
            row_vals = [format_cell(cell) for cell in row if str(cell).strip()]
            if len(row_vals) == 2:
                key, value = row_vals
                header_data[key] = value

        table_headers = []
        for i, row in df.iterrows():
            if is_potential_table_header(row):
                table_headers.append(i)

        all_tables = []
        for idx, start_row in enumerate(table_headers):
            end_row = table_headers[idx + 1] if idx + 1 < len(table_headers) else len(df)

            sub_df = df.iloc[start_row:end_row].copy()
            headers = [format_cell(h) for h in sub_df.iloc[0]]
            sub_df.columns = headers
            sub_df = sub_df[1:].reset_index(drop=True)

            table_data = []
            for _, row in sub_df.iterrows():
                row_dict = {}
                for col, val in row.items():
                    key = format_cell(col)
                    value = format_cell(val)
                    if key and value:
                        row_dict[key] = value
                if row_dict:
                    table_data.append(row_dict)

            if table_data:
                all_tables.append({
                    "Table Start Row": int(start_row),
                    "Records": table_data
                })

        return {
            "Header Information": header_data,
            "Allottee Tables": all_tables,
            "message": f" Extracted {len(all_tables)} table(s) from the file."
        }

    except Exception as e:
        return {"message": f" Error extracting data: {str(e)}"}

if uploaded_file:
    with st.spinner(" Extracting data from Excel..."):
        result = extract_excel(uploaded_file)

    st.success(" Extraction Complete!")
    st.subheader(" JSON Preview")
    st.json(result)

    json_bytes = BytesIO(json.dumps(result, indent=2).encode())
    st.download_button(" Download JSON", data=json_bytes, file_name="extracted_allotments.json", mime="application/json")
else:
    st.info(" Please upload your Excel or CSV file.")
