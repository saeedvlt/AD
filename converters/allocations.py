import pandas as pd from openpyxl 
import load_workbook

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun",
"Jul","Aug","Sep","Oct","Nov","Dec"]

def _find_month_row(ws): for r in range(1, ws.max_row + 1): values =
[ws.cell(r, c).value for c in range(1, ws.max_column + 1)] if sum(v in
MONTHS for v in values) >= 6: return r return None

def _unpivot_standard_sheet(ws, sheet_name, allocation_type): month_row
= _find_month_row(ws) if month_row is None: return []

    month_cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(month_row, c).value
        if v in MONTHS:
            month_cols[c] = v

    first_month_col = min(month_cols)

    rows = []

    for r in range(month_row + 1, ws.max_row + 1):

        if all(ws.cell(r, c).value is None for c in range(1, ws.max_column + 1)):
            continue

        description = None

        for c in range(1, first_month_col):
            value = ws.cell(r, c).value
            if isinstance(value, str) and value.strip():
                description = value.strip()

        if not description:
            continue

        for col, month in month_cols.items():
            value = ws.cell(r, col).value
            if value is None:
                continue

            rows.append({
                "Allocation_Type": allocation_type,
                "Source_Sheet": sheet_name,
                "Description": description,
                "Month": month,
                "Value": value
            })

    return rows

def convert(uploaded_file):

    wb = load_workbook(uploaded_file, data_only=True)

    all_rows = []

    sheet_map = {
        "nads": "NADS",
        "corp": "CORP",
        "outside": "Outside Sales",
        "it": "IT",
        "dpf": "NA DPF US"
    }

    for sheet in wb.sheetnames:

        name = sheet.lower()

        if "allocation" not in name and "allocations" not in name:
            continue

        allocation_type = "Other"

        for key, value in sheet_map.items():
            if key in name:
                allocation_type = value
                break

        ws = wb[sheet]

        if allocation_type == "IT":
            # TODO:
            # IT Allocation has a different layout.
            # Add custom parser here when finalized.
            continue

        all_rows.extend(
            _unpivot_standard_sheet(
                ws,
                sheet,
                allocation_type
            )
        )

    return pd.DataFrame(all_rows)
