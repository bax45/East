import csv
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
import re
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import pandas as pd
import logging


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
    """
    Reads existing records from template to learn known (Rack_No. -> Rack_Name) mappings.
    Preserves valid names and does not overwrite them with empty records.
    """
    knowledge = {}
    for r in range(2, ws.max_row + 1):
        cell_rack_no = ws.cell(row=r, column=1).value
        cell_rack_name = ws.cell(row=r, column=2).value
        
        if cell_rack_no is not None:
            rack_no_key = str(cell_rack_no).strip()
            rack_name_val = str(cell_rack_name).strip() if cell_rack_name is not None else ""
            
            # If we haven't seen this rack yet, register it
            if rack_no_key not in knowledge:
                knowledge[rack_no_key] = rack_name_val
            # If we find a real name in a later row, update an empty entry
            elif not knowledge[rack_no_key] and rack_name_val:
                knowledge[rack_no_key] = rack_name_val

    return knowledge




def resolve_rack_with_floor(
    raw_rack: str, floor_prefix: str, template_racks: dict
) -> tuple[str, str]:
    """Resolves raw CSV rack code (e.g.

    '11A') with floor prefix ('2F_') to match canonical template racks (e.g.
    '2F_A11').
    """
    if not raw_rack:
        return "", ""

    raw_str = str(raw_rack).strip()
    digits = "".join(re.findall(r"\d+", raw_str))
    letters = "".join(re.findall(r"[A-Za-z]+", raw_str)).upper()

    # Normalize floor prefix format (ensure it ends with '_' if present, e.g. '2F_')
    norm_prefix = str(floor_prefix).strip()
    if norm_prefix and not norm_prefix.endswith("_"):
        norm_prefix += "_"

    floor_num = "".join(re.findall(r"\d+", norm_prefix))

    # Candidate variations (e.g. '2F_A11', '2F_11A', '2F_A01')
    candidates = [
        f"{norm_prefix}{letters}{digits}",
        f"{norm_prefix}{digits}{letters}",
    ]
    # Handle possible zero-padded variations (e.g. 1 -> 01 or 01 -> 1)
    if digits:
        candidates.extend(
            [
                f"{norm_prefix}{letters}{digits.zfill(2)}",
                f"{norm_prefix}{letters}{int(digits)}",
            ]
        )

    # 1. Exact match against template
    for cand in candidates:
        for t_rack_no, t_rack_name in template_racks.items():
            if cand.upper() == str(t_rack_no).strip().upper():
                return t_rack_no, t_rack_name or ""

    # 2. Structural match constrained strictly to the SAME floor
    for t_rack_no, t_rack_name in template_racks.items():
        t_str = str(t_rack_no).strip().upper()

        # Enforce floor constraint
        if floor_num and not t_str.startswith(f"{floor_num}F"):
            continue

        tail = t_str.split("_")[-1]
        t_digits = "".join(re.findall(r"\d+", tail))
        t_letters = "".join(re.findall(r"[A-Za-z]+", tail))

        if letters and letters == t_letters:
            if digits and (
                digits == t_digits
                or int(digits or 0) == int(t_digits or -1)
            ):
                return t_rack_no, t_rack_name or ""

    # 3. Constrained fuzzy match (only within the same floor)
    best_match = None
    highest_ratio = 0.0

    for t_rack_no, t_rack_name in template_racks.items():
        t_str = str(t_rack_no).strip().upper()

        # Never fuzzy-match across different floors
        if floor_num and not t_str.startswith(f"{floor_num}F"):
            continue

        ratio = SequenceMatcher(
            None, raw_str.upper(), t_str
        ).ratio()
        if ratio > highest_ratio and ratio >= 0.7:
            highest_ratio = ratio
            best_match = (t_rack_no, t_rack_name or "")

    if best_match:
        return best_match

    # 4. Safe fallback: do not invent names or steal from other racks
    fallback_code = (
        f"{norm_prefix}{letters}{digits}" if norm_prefix else raw_str
    )
    return fallback_code, ""




def extract_date_flexible(text: str):
    """Detects and parses dates in various common formats (DD.MM.YYYY, YYYY-MM-DD,

    DD/MM/YYYY).
    """
    patterns = [
        (r"\b(\d{1,2}\.\d{1,2}\.\d{4})\b", "%d.%m.%Y"),
        (r"\b(\d{4}-\d{2}-\d{2})\b", "%Y-%m-%d"),
        (r"\b(\d{1,2}/\d{1,2}/\d{4})\b", "%d/%m/%Y"),
    ]
    for pattern, date_fmt in patterns:
        m = re.search(pattern, text)
        if m:
            try:
                return datetime.strptime(m.group(1), date_fmt).date()
            except ValueError:
                continue
    return None


def parse_pdu_csv(file_path: Path) -> pd.DataFrame:
    """Parses semicolon-delimited CSV, handles dates, preserves column alignment,

    and extracts PDU metrics dynamically without hardcoded limits.
    """
    records = []
    file_date = None

    with open(file_path, mode="r", encoding="utf-8-sig", errors="ignore") as f:
        lines = [line.rstrip("\r\n") for line in f if line.strip()]

    if not lines:
        return pd.DataFrame()

    floor_prefix = detect_floor_prefix(file_path, lines)

    # 1. Search first 15 lines for a date; fallback to parsing the filename
    for line in lines[:15]:
        detected = extract_date_flexible(line)
        if detected:
            file_date = detected
            break

    if not file_date:
        file_date = extract_date_flexible(file_path.stem)

    # 2. Parse table rows with strict 1:1 positional indexing
    grid = [line.split(";") for line in lines]

    for i, row in enumerate(grid):
        row_str = ";".join(row).lower()

        # Target header rows that contain measurement metric keywords
        if any(
            k in row_str
            for k in ["max", "average", "avg", "макс", "средн"]
        ):
            if i + 1 >= len(grid):
                break

            headers = row
            val_row = grid[i + 1]

            # Preserve exact index positioning
            min_len = min(len(headers), len(val_row))

            for idx in range(min_len):
                header = headers[idx].strip()
                val_str = val_row[idx].strip()

                if not header or not val_str:
                    continue

                # Matches: RackID + Delimiter + PDU (1, 2, A, B, etc.) + Delimiter + Metric
                m = re.search(
                    r"^([A-Za-z0-9_-]+)[_. -]+(PDU[_-]?[0-9A-Za-z]+|[AB])[_. -]+(max|average|avg|макс|средн.*)$",
                    header,
                    re.IGNORECASE,
                )

                if m:
                    raw_rack = m.group(1)
                    pdu_raw = m.group(2).upper().replace("_", "").replace("-", "")
                    metric_raw = m.group(3).lower()

                    # Normalize metric name
                    if any(x in metric_raw for x in ["avg", "средн"]):
                        metric = "average"
                    else:
                        metric = "max"

                    # Normalize numeric value
                    try:
                        clean_num = re.sub(r"[^\d.,-]", "", val_str).replace(",", ".")
                        val = float(clean_num) if clean_num else None
                    except ValueError:
                        val = None

                    records.append(
                        {
                            "Raw_Rack": raw_rack,
                            "Floor_Prefix": floor_prefix,
                            "Date": file_date,
                            "PDU": pdu_raw,
                            "Metric": metric,
                            "Value": val,
                        }
                    )

    return pd.DataFrame(records)



logger = logging.getLogger("PowerUsageApp")


def main():
    template_path, csv_files = find_files()

    wb = openpyxl.load_workbook(template_path)
    ws = wb["YaIDC"] if "YaIDC" in wb.sheetnames else wb.active
    template_racks = load_template_knowledge(ws)

    # 1. Parse all CSVs
    all_dfs = []
    for csv_path in csv_files:
        df = parse_pdu_csv(csv_path)
        if not df.empty:
            all_dfs.append(df)

    if not all_dfs:
        print("[ERROR] No valid metrics extracted from CSV files.")
        return

    full_df = pd.concat(all_dfs, ignore_index=True)

    # 2. Date propagation: Group by BOTH Floor_Prefix and Raw_Rack to prevent cross-floor leakage
    full_df["Date"] = full_df.groupby(["Floor_Prefix", "Raw_Rack"])["Date"].transform(
        lambda s: s.bfill().ffill()
    )

    # Fallback for any remaining unassigned dates: use the overall mode/first valid date
    if full_df["Date"].isna().any():
        valid_dates = full_df["Date"].dropna()
        if not valid_dates.empty:
            full_df["Date"] = full_df["Date"].fillna(valid_dates.iloc[0])

    # 3. Resolve rack numbers and names with floor disambiguation
    matched_results = full_df.apply(
        lambda row: resolve_rack_with_floor(
            row["Raw_Rack"], row["Floor_Prefix"], template_racks
        ),
        axis=1,
    )
    full_df["Rack_No."] = [res[0] for res in matched_results]
    full_df["Rack_Name"] = [res[1] for res in matched_results]

    # 4. Pivot so PDU feeds and metrics align into individual columns
    pivoted = full_df.pivot_table(
        index=["Rack_No.", "Rack_Name", "Date"],
        columns=["PDU", "Metric"],
        values="Value",
        aggfunc="first",
    ).reset_index()

    # Flatten multi-level columns (e.g. ('PDU1', 'max') -> 'PDU1_max')
    pivoted.columns = [
        f"{c[0]}_{c[1]}".strip("_") if c[1] != "" else c[0]
        for c in pivoted.columns
    ]

    # 5. Dynamically determine report year from data (No Hardcoding)
    report_year = 2026
    sample_date = full_df["Date"].dropna().iloc[0] if not full_df["Date"].dropna().empty else None
    if sample_date:
        if hasattr(sample_date, "year"):
            report_year = sample_date.year
        elif isinstance(sample_date, str) and len(sample_date) >= 4:
            report_year = sample_date[:4]

    output_dir = Path("output")
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{report_year}_PowerUsageTable.xlsx"

    # 6. --- DYNAMIC TABLE INTEGRATION ---
    target_table = list(ws.tables.values())[0] if ws.tables else None

    if target_table:
        min_col, min_row, max_col, max_row = range_boundaries(target_table.ref)
        header_row = min_row
    else:
        min_col, header_row, max_col = 1, 1, ws.max_column

    # Map column headers dynamically from the worksheet's header row
    col_name_to_idx = {}
    for col_idx in range(min_col, max_col + 1):
        cell_val = ws.cell(row=header_row, column=col_idx).value
        if cell_val:
            col_name_to_idx[str(cell_val).strip().lower()] = col_idx

    def get_col(*aliases):
        """Finds column index in worksheet matching any alias."""
        for alias in aliases:
            cleaned = alias.lower()
            if cleaned in col_name_to_idx:
                return col_name_to_idx[cleaned]
        return None

    c_rack_no   = get_col("rack_no.", "rack_no", "rack no", "rack", "стойка")
    c_rack_name = get_col("rack_name", "rack name", "name", "наименование")
    c_pdu1_avg  = get_col("pdu1_average", "pdu1_avg", "pdu1 average", "pdu 1 average")
    c_pdu1_max  = get_col("pdu1_max", "pdu 1 max")
    c_pdu2_avg  = get_col("pdu2_average", "pdu2_avg", "pdu2 average", "pdu 2 average")
    c_pdu2_max  = get_col("pdu2_max", "pdu 2 max")
    c_date      = get_col("date", "дата", "time", "timestamp")

    def get_val(df_row, *keys):
        """Safely extracts value without dropping 0.0."""
        for k in keys:
            if k in df_row and pd.notna(df_row[k]):
                return df_row[k]
        return None

    # Write data starting directly below the header row
    start_row = header_row + 1
    for idx, row in pivoted.iterrows():
        curr_row = start_row + idx

        if c_rack_no:
            ws.cell(row=curr_row, column=c_rack_no, value=row.get("Rack_No."))
        if c_rack_name:
            ws.cell(row=curr_row, column=c_rack_name, value=row.get("Rack_Name"))
        if c_pdu1_avg:
            ws.cell(row=curr_row, column=c_pdu1_avg, value=get_val(row, "PDU1_average", "PDU1_avg"))
        if c_pdu1_max:
            ws.cell(row=curr_row, column=c_pdu1_max, value=get_val(row, "PDU1_max"))
        if c_pdu2_avg:
            ws.cell(row=curr_row, column=c_pdu2_avg, value=get_val(row, "PDU2_average", "PDU2_avg"))
        if c_pdu2_max:
            ws.cell(row=curr_row, column=c_pdu2_max, value=get_val(row, "PDU2_max"))
        if c_date:
            ws.cell(row=curr_row, column=c_date, value=row.get("Date"))

    # 7. Expand table boundaries to encompass newly inserted rows
    new_max_row = start_row + len(pivoted) - 1
    if target_table:
        new_ref = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{new_max_row}"
        )
        target_table.ref = new_ref

    # 8. Save
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