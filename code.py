import os
import pandas as pd
from docx import Document
from difflib import SequenceMatcher
from collections import defaultdict


# ------------------------- Configuration -------------------------
INPUT_EXCEL = "input_data.xlsx"
OUTPUT_EXCEL = "output_comparison.xlsx"

SCRIPTS_FOLDER_PY = "scripts/python"
SCRIPTS_FOLDER_JAVA = "scripts/java"
SCRIPTS_FOLDER_WORD = "scripts/word"

# Mapping Excel columns to extracted fields
COLUMN_MAPPING = {
    "Excel_ID": "ID",
    "Excel_Objective": "Objective",
    "Excel_Author": "Author"
}

JAVA_KEY_MAP = {
    "id": "ID",
    "description": "Objective",
    "title": "Filename",
    "author": "Author"
}

WORD_KEY_MAP = JAVA_KEY_MAP


# ------------------------- Utility Functions -------------------------
def similar(a, b):
    return round(SequenceMatcher(None, a.lower(), b.lower()).ratio() * 100, 2)


def merge_excel_rows(df):
    """Merge rows in Excel with the same filename"""
    grouped = defaultdict(lambda: defaultdict(list))

    for _, row in df.iterrows():
        filename = str(row.get("Filename", "")).strip().lower()
        for col, val in row.items():
            if pd.notna(val):
                grouped[filename][col].append(str(val).strip())

    merged = []
    for fname, data in grouped.items():
        merged_row = {"Filename": fname}
        for col, values in data.items():
            merged_row[col] = "\n".join(sorted(set(values)))
        merged.append(merged_row)

    return pd.DataFrame(merged)


# ------------------------- Extractor: Python -------------------------
def extract_comment_header(file_path):
    data = {}
    with open(file_path, "r", encoding="utf-8") as f:
        current_key = None
        for line in f:
            if not line.strip().startswith("#"):
                break
            content = line.lstrip("#").strip()
            if ":" in content:
                key, value = content.split(":", 1)
                data[key.strip()] = value.strip()
                current_key = key.strip()
            elif current_key:
                data[current_key] += "\n" + content.strip()
    return data


# ------------------------- Extractor: Java -------------------------
def extract_java_header(file_path):
    data = {}
    with open(file_path, "r", encoding="utf-8") as f:
        header_started = False
        lines = []
        for line in f:
            if "public static String Header" in line:
                header_started = True
                continue
            if header_started:
                if '";' in line:
                    break
                lines.append(line.strip())

    full_header = "".join(lines).replace('"+', "").replace('"', "")
    for line in full_header.split("\\n"):
        if ":" in line:
            key, value = line.split(":", 1)
            data[JAVA_KEY_MAP.get(key.strip().lower(), key.strip())] = value.strip()
    return data


# ------------------------- Extractor: Word -------------------------
def extract_word_cases(file_path):
    doc = Document(file_path)
    tables = doc.tables
    if len(tables) < 2:
        return {}, [], ["Template or test cases not found"]

    metadata = {}
    for row in tables[0].rows:
        if len(row.cells) < 2:
            continue
        key = row.cells[0].text.strip()
        value = row.cells[1].text.strip()
        if key:
            metadata[key] = value

    expected_ids = set()
    for k, v in metadata.items():
        if "another id" in k.lower():
            expected_ids.update(p.strip() for p in v.replace(";", ",").split(",") if p.strip())

    cases = []
    found_ids = set()

    for table in tables[1:]:
        case_data = {}
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            key = row.cells[0].text.strip()
            value = row.cells[1].text.strip()
            if key:
                if key in case_data:
                    case_data[key] += "\n" + value
                else:
                    case_data[key] = value

        cid = case_data.get("Case ID") or case_data.get("case id")
        if cid:
            found_ids.add(cid.strip())
        cases.append(case_data)

    mismatches = []
    if expected_ids - found_ids:
        mismatches.append("IDs in template but not in cases: " + ", ".join(expected_ids - found_ids))
    if found_ids - expected_ids:
        mismatches.append("Case IDs not listed in template: " + ", ".join(found_ids - expected_ids))

    return metadata, cases, mismatches


def extract_word_header(file_path):
    metadata, cases, mismatches = extract_word_cases(file_path)

    combined_fields = {}
    for case in cases:
        for key, value in case.items():
            if key in combined_fields:
                combined_fields[key] += "\n" + value
            else:
                combined_fields[key] = value

    result = {k: v for k, v in metadata.items()}
    for k, v in combined_fields.items():
        if k not in result:
            result[k] = v
        else:
            result[k] += "\n" + v

    if mismatches:
        result["Word Template Issues"] = "; ".join(mismatches)

    return result


# ------------------------- Main Comparison Logic -------------------------
def main():
    df = pd.read_excel(INPUT_EXCEL)
    df = merge_excel_rows(df)
    df["Filename"] = df["Filename"].str.strip()

    all_files = []

    # --- Collect and parse all script files ---
    for folder, ext, extractor, filetype in [
        (SCRIPTS_FOLDER_PY, ".py", extract_comment_header, "Python"),
        (SCRIPTS_FOLDER_JAVA, ".java", extract_java_header, "Java"),
        (SCRIPTS_FOLDER_WORD, ".docx", extract_word_header, "Word"),
    ]:
        for fname in os.listdir(folder):
            if not fname.endswith(ext):
                continue
            path = os.path.join(folder, fname)
            data = extractor(path)
            data["Filename"] = fname
            data["FileType"] = filetype
            all_files.append(data)

    results = []

    excel_map = {
        str(row["Filename"]).strip().lower(): row
        for _, row in df.iterrows()
    }

    processed_files = set()

    for file_data in all_files:
        filename = file_data.get("Filename", "").strip()
        filename_lower = filename.lower()
        processed_files.add(filename_lower)

        excel_row = excel_map.get(filename_lower)
        entry = {
            "Filename": filename,
            "FileType": file_data.get("FileType", ""),
            "Title Match %": similar(filename, excel_row["Filename"]) if excel_row is not None else 0
        }

        file_remarks = []
        excel_remarks = []

        if excel_row:
            for excel_col, extracted_key in COLUMN_MAPPING.items():
                expected = str(excel_row.get(excel_col, "")).strip()
                actual = str(file_data.get(extracted_key, "")).strip()
                if expected and not actual:
                    file_remarks.append(f"{extracted_key}: missing")
                elif actual and not expected:
                    excel_remarks.append(f"{extracted_key}: missing")
                elif expected != actual:
                    file_remarks.append(f"{extracted_key}: mismatch")

            if file_remarks and not excel_remarks:
                entry["Error"] = "mismatch error"
            elif excel_remarks and not file_remarks:
                entry["Error"] = "missing errors"
            elif file_remarks and excel_remarks:
                entry["Error"] = "missing\nmismatch"
            else:
                entry["Error"] = ""
        else:
            entry["Error"] = "filename missing in excel"

        entry["Excel Remarks"] = "; ".join(excel_remarks) if excel_remarks else ""
        entry["File Remarks"] = "; ".join(file_remarks) if file_remarks else ""

        # Add all extracted data to output
        for k, v in file_data.items():
            if k not in entry:
                entry[k] = v

        results.append(entry)

    # --- Add files in Excel but not found on disk ---
    for filename in df["Filename"]:
        fname_lower = str(filename).strip().lower()
        if fname_lower not in processed_files:
            results.append({
                "Filename": filename,
                "Error": "file not found"
            })

    pd.DataFrame(results).to_excel(OUTPUT_EXCEL, index=False)
    print(f"✅ Comparison complete. Output saved to: {OUTPUT_EXCEL}")


# ------------------------- Entry Point -------------------------
if __name__ == "__main__":
    main()
