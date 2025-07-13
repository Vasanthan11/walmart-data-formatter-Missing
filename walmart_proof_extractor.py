import re
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

# Function to extract proof type
def detect_proof(page_name, path):
    combined = f"{page_name} {path}".lower()
    if "press" in combined:
        return "PRESS"
    elif "preview" in combined:
        return "PREVIEW"
    elif "cpr" in combined:
        return "CPR"
    elif "final" in combined:
        return "FINAL"
    return "Unknown"

# Function to clean page names
def clean_page_name(name):
    return name.replace("Proof", "").replace("_", " ").replace("-", " ").strip()

# Main parsing function
def extract_data(raw_text):
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    unwanted_keywords = ["unread", "confirm", "reduce"]
    cleaned_lines = [line for line in lines if not any(k in line.lower() for k in unwanted_keywords)]

    result = []
    skipped = []
    i = 0
    while i < len(cleaned_lines):
        try:
            line = cleaned_lines[i]
            next1 = cleaned_lines[i+1] if i+1 < len(cleaned_lines) else ""
            next2 = cleaned_lines[i+2] if i+2 < len(cleaned_lines) else ""

            # Case 1: Time-stamped line or assembler line
            if re.search(r'\d{1,2}[:.]\d{2}\s*[APMapm]+', line) or line.upper().startswith("D-"):
                assembler = line.split(',')[0].strip() if ',' in line else "Unknown"
                page = next1
                path = next2

                if path.startswith("/Volumes"):
                    page_clean = clean_page_name(page)
                    week_match = re.search(r'W[K\- ]*(\d+)', page_clean, re.IGNORECASE)
                    week = f"week-{week_match.group(1)}" if week_match else ""
                    proof = detect_proof(page_clean, path)
                    qc = "Direct Upload" if page.upper().startswith("D-") else "Hariharan"
                    result.append({
                        "Date": datetime.now().strftime("%d/%m/%Y"),
                        "Banner Name": "walmart",
                        "Week": week,
                        "Page Name": page_clean,
                        "Proof": proof,
                        "Language": "All zones",
                        "Page Assembler": assembler,
                        "QC": qc
                    })
                    i += 3
                    continue
                else:
                    skipped.append(f"{line}\n{page}\n{path}")
                    i += 3
                    continue

            # Case 2: CORP WK entries
            elif re.search(r'CORP\s*\[?WK\s*\d+', line, re.IGNORECASE):
                assembler = "Mohammed Siddik" if "Mohammed Siddik" in cleaned_lines[i - 1] else "Prasanth As"
                page = line
                path = next1
                if path.startswith("/Volumes"):
                    page_clean = clean_page_name(page)
                    week_match = re.search(r'W[K\- ]*(\d+)', page_clean, re.IGNORECASE)
                    week = f"week-{week_match.group(1)}" if week_match else ""
                    proof = detect_proof(page_clean, path)
                    result.append({
                        "Date": "25/06/2025",
                        "Banner Name": "walmart",
                        "Week": week,
                        "Page Name": page_clean,
                        "Proof": proof,
                        "Language": "All zones",
                        "Page Assembler": assembler,
                        "QC": "Hariharan"
                    })
                    i += 2
                    continue
                else:
                    skipped.append(f"{line}\n{path}")
                    i += 2
                    continue

            # Case 3: D- prefix without time
            elif line.upper().startswith("D-") and next1.startswith("/Volumes"):
                assembler = "Unknown"
                page_clean = clean_page_name(line)
                week_match = re.search(r'W[K\- ]*(\d+)', page_clean, re.IGNORECASE)
                week = f"week-{week_match.group(1)}" if week_match else ""
                proof = detect_proof(page_clean, next1)
                result.append({
                    "Date": datetime.now().strftime("%d/%m/%Y"),
                    "Banner Name": "walmart",
                    "Week": week,
                    "Page Name": page_clean,
                    "Proof": proof,
                    "Language": "All zones",
                    "Page Assembler": assembler,
                    "QC": "Direct Upload"
                })
                i += 2
                continue

            else:
                skipped.append(line)
                i += 1

        except Exception as e:
            skipped.append(f"Error at line {i}: {line} → {str(e)}")
            i += 1

    return pd.DataFrame(result), skipped

# Streamlit app UI
st.title("🛠 Walmart Proof Extractor + Skipped Entry Checker")
raw_text = st.text_area("📋 Paste the raw proof content:")

if st.button("🚀 Generate Excel"):
    if not raw_text.strip():
        st.warning("Please paste some content before processing.")
    else:
        df, skipped = extract_data(raw_text)

        if df.empty:
            st.error("No valid data extracted.")
        else:
            st.success(f"✅ {len(df)} entries extracted successfully.")
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Proof Data")
                worksheet = writer.sheets["Proof Data"]
                proof_col = df.columns.get_loc("Proof")
                options = ["CPR", "PRESS", "PREVIEW", "FINAL", "Unknown"]
                for row in range(1, len(df)+1):
                    worksheet.data_validation(row, proof_col, row, proof_col, {
                        'validate': 'list',
                        'source': options
                    })
            buffer.seek(0)
            st.download_button("📥 Download Excel", buffer, file_name="proof_data.xlsx")

        if skipped:
            with st.expander("⚠️ Skipped Entries"):
                st.write("These lines were not processed into the Excel file:")
                st.code("\n\n".join(skipped))
