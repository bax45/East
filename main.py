import csv
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
import re
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR
OUTPUT_DIR = BASE_DIR / "output" if (BASE_DIR / "output").exists() else BASE_DIR
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def find_files():
    """Locates the Excel template and all raw PDU CSV files."""
    templates = list(DATA_DIR.glob("*template*.xlsx")) or list(
        BASE_DIR.glob("*template*.xlsx")
    )
    if not templates:
        raise FileNotFoundError(
            "Excel template not found in project or data directory!"
        )

    csv_files = list(DATA_DIR.glob("*.csv")) or list(BASE_DIR.glob("*.csv"))
    if not csv_files:
        raise FileNotFoundError(
            "No CSV files found in project or data directory!"
        )

    return templates[0], csv_files


def detect_floor_prefix(file_path: Path, lines: list[str]) -> str:
    """Extracts floor identifier (e.g., '2F_') from filename or top header lines."""
    m = re.search(r"PDU_(\d+)", file_path.stem, re.IGNORECASE)
    if m:
        return f"{m.group(1)}F_"

    for line in lines[:5]:
        m = re.search(r"PDU_(\d+)", line, re.IGNORECASE)
        if m:
            return f"{m.group(1)}F_"

    return ""


def load_template_knowledge(ws) -> dict:
    """Reads existing records from template to learn known (Rack_No., Rack_Name) pairs."""
    knowledge = {}
    for r in range(2, ws.max_row + 1):
        rack_no = ws.cell(row=r, column=1).value
        rack_name = ws.cell(row=r, column=2).value
        if rack_no:
            knowledge[str(rack_no).strip()] = (
                str(rack_name).strip() if rack_name else "Selectel"
            )
    return knowledge


def resolve_rack_with_floor(
    raw_rack: str, floor_prefix: str, template_racks: dict
) -> tuple[str, str]:
    """Resolves raw CSV rack code (e.g., '11A') with floor prefix ('2F_')

    to distinguish '2F_A11' from '3F_A11' based on template knowledge.
    """
    digits = "".join(re.findall(r"\d+", str(raw_rack)))
    letters = "".join(re.findall(r"[A-Za-z]+", str(raw_rack))).upper()

    candidates = [
        f"{floor_prefix}{letters}{digits}",
        f"{floor_prefix}{digits}{letters}",
    ]

    for cand in candidates:
        for t_rack_no, t_rack_name in template_racks.items():
            if cand.upper() == t_rack_no.upper():
                return t_rack_no, t_rack_name

    floor_num = "".join(re.findall(r"\d+", floor_prefix))
    for t_rack_no, t_rack_name in template_racks.items():
        if floor_num and not t_rack_no.startswith(f"{floor_num}F"):
            continue

        tail = t_rack_no.split("_")[-1]
        t_digits = "".join(re.findall(r"\d+", tail))
        t_letters = "".join(re.findall(r"[A-Za-z]+", tail)).upper()
        if digits == t_digits and letters in t_letters:
            return t_rack_no, t_rack_name

    best_match = None
    highest_ratio = 0.0
    for t_rack_no, t_rack_name in template_racks.items():
        ratio = SequenceMatcher(
            None, raw_rack.upper(), t_rack_no.upper()
        ).ratio()
        if ratio > highest_ratio and ratio >= 0.6:
            highest_ratio = ratio
            best_match = (t_rack_no, t_rack_name)

    if best_match:
        return best_match

    default_name = list(template_racks.values())[0] if template_racks else "Selectel"
    fallback_code = f"{floor_prefix}{letters}{digits}" if floor_prefix else raw_rack
    return fallback_code, default_name


def parse_pdu_csv(file_path: Path):
    """Parses semicolon-delimited CSV, handles '#' dates, and extracts PDU metrics."""
    records = []
    file_date = None

    with open(file_path, mode="r", encoding="utf-8-sig", errors="ignore") as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]

    floor_prefix = detect_floor_prefix(file_path, lines)

    # Extract date (keeps original date object so table constructor formats it)
    for line in lines[:10]:
        for cell in line.split(";"):
            cell = cell.strip()
            if re.match(r"^\d{2}\.\d{2}\.\d{4}$", cell):
                try:
                    file_date = datetime.strptime(cell, "%d.%m.%Y").date()
                except ValueError:
                    file_date = cell
                break
        if file_date:
            break

    reader = [line.split(";") for line in lines]
    for i, row in enumerate(reader):
        row_str = ";".join(row).lower()
        if "max" in row_str or "average" in row_str:
            headers = [c.strip() for c in row if c.strip()]
            if i + 1 < len(reader):
                val_row = [c.strip() for c in reader[i + 1] if c.strip()]

                for header, val_str in zip(headers, val_row):
                    m = re.search(
                        r"^([A-Za-z0-9_-]+)_(PDU[12])_(max|average)$",
                        header,
                        re.IGNORECASE,
                    )
                    if m:
                        raw_rack = m.group(1)
                        pdu_feed = m.group(2).upper()
                        metric = m.group(3).lower()

                        try:
                            val = float(val_str.replace(",", "."))
                        except (ValueError, AttributeError):
                            val = None

                        records.append(
                            {
                                "Raw_Rack": raw_rack,
                                "Floor_Prefix": floor_prefix,
                                "Date": file_date,
                                "PDU": pdu_feed,
                                "Metric": metric,
                                "Value": val,
                            }
                        )

    return pd.DataFrame(records)


def main():
    template_path, csv_files = find_files()
    output_path = OUTPUT_DIR / "2026_PowerUsageTable.xlsx"

    wb = openpyxl.load_workbook(template_path)
    ws = wb["YaIDC"] if "YaIDC" in wb.sheetnames else wb.active
    template_racks = load_template_knowledge(ws)

    all_dfs = []
    for csv_path in csv_files:
        df = parse_pdu_csv(csv_path)
        if not df.empty:
            all_dfs.append(df)

    if not all_dfs:
        print("[ERROR] No valid metrics extracted from CSV files.")
        return

    full_df = pd.concat(all_dfs, ignore_index=True)

    # Propagate valid date across paired feeds
    full_df["Date"] = full_df.groupby("Raw_Rack")["Date"].transform(
        lambda s: s.bfill().ffill()
    )

    # Disambiguate rack numbers and assign rack names based on floor + template
    matched_results = full_df.apply(
        lambda row: resolve_rack_with_floor(
            row["Raw_Rack"], row["Floor_Prefix"], template_racks
        ),
        axis=1,
    )
    full_df["Rack_No."] = [res[0] for res in matched_results]
    full_df["Rack_Name"] = [res[1] for res in matched_results]

    # Pivot so PDU1 and PDU2 align horizontally
    pivoted = full_df.pivot_table(
        index=["Rack_No.", "Rack_Name", "Date"],
        columns=["PDU", "Metric"],
        values="Value",
        aggfunc="first",
    ).reset_index()

    pivoted.columns = [
        f"{c[0]}_{c[1]}" if c[1] != "" else c[0] for c in pivoted.columns
    ]

    # --- TABLE INTEGRATION ---
    # Find the table constructor object in the worksheet
    target_table = None
    if ws.tables:
        target_table = list(ws.tables.values())[0]

    if target_table:
        min_col, min_row, max_col, max_row = range_boundaries(target_table.ref)
        start_row = max_row + 1
    else:
        min_col, min_row, max_col = 1, 1, 7
        start_row = ws.max_row + 1

    # Insert data rows directly into the table
    for idx, row in pivoted.iterrows():
        curr_row = start_row + idx

        ws.cell(row=curr_row, column=1, value=row.get("Rack_No."))
        ws.cell(row=curr_row, column=2, value=row.get("Rack_Name"))
        ws.cell(
            row=curr_row,
            column=3,
            value=row.get("PDU1_average") or row.get("PDU1_avg"),
        )
        ws.cell(row=curr_row, column=4, value=row.get("PDU1_max"))
        ws.cell(
            row=curr_row,
            column=5,
            value=row.get("PDU2_average") or row.get("PDU2_avg"),
        )
        ws.cell(row=curr_row, column=6, value=row.get("PDU2_max"))
        ws.cell(row=curr_row, column=7, value=row.get("Date"))

    # Update table range definition so Excel recognizes new rows as part of the table
    if target_table:
        new_max_row = start_row + len(pivoted) - 1
        new_ref = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{new_max_row}"
        )
        target_table.ref = new_ref

    try:
        wb.save(output_path)
        print(f"[SUCCESS] Data added inside Table at: {output_path.resolve()}")
    except PermissionError:
        print(
            f"[ERROR] Could not save to '{output_path.name}'. "
            "Please close Microsoft Excel and run the script again."
        )


if __name__ == "__main__":
    main()