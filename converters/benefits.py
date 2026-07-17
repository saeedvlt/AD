import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import re



MONTH_MAP = {
    "Jan": "01 - Jan",
    "Feb": "02 - Feb",
    "Mar": "03 - Mar",
    "Apr": "04 - Apr",
    "May": "05 - May",
    "Jun": "06 - Jun",
    "Jul": "07 - Jul",
    "Aug": "08 - Aug",
    "Sep": "09 - Sep",
    "Oct": "10 - Oct",
    "Nov": "11 - Nov",
    "Dec": "12 - Dec"
}

META_COLS = {
    1: "Emp #",
    2: "Name",
    3: "Employee Code",
    4: "Date of Hire",
    5: "Site Location",
    6: "AOP File Assignment"
}


def split_name(name):

    if pd.isna(name):
        return "", ""

    name = str(name).strip()

    parts = re.split(
        r'(?<=[a-z])(?=[A-Z])',
        name,
        maxsplit=1
    )

    if len(parts) == 2:
        return parts[0], parts[1]

    return name, ""


def convert(uploaded_file):

    wb = load_workbook(
        uploaded_file,
        data_only=True
    )

    if "Benefits" in wb.sheetnames:
    	ws = wb["Benefits"]
    else:
    	raise ValueError("Worksheet 'Benefits' not found.")
    
    benefit_cols = {}

    current_benefit = None

    for idx, col in enumerate(range(7, ws.max_column + 1)):


        top = ws.cell(4, col).value
        low = ws.cell(5, col).value

        if top not in (None, ""):
            current_benefit = str(top).strip()

        if low is None:
            continue

        if isinstance(low, datetime):
            benefit_cols[col] = (
                current_benefit,
                low.strftime("%b")
            )
            continue

        low = str(low).strip()

        if low in (
            "Total",
            "CPP Calc",
            "Amount",
            "Period of Payment",
            ""
        ):
            continue

        if low[:3] in (
            "Jan",
            "Feb",
            "Mar",
            "Apr",
            "May",
            "Jun",
            "Jul",
            "Aug",
            "Sep",
            "Oct",
            "Nov",
            "Dec"
        ):

            benefit_cols[col] = (
                current_benefit,
                low[:3]
            )

    records = []

    for i, row in enumerate(range(7, ws.max_row + 1)):

        employee = {}

        empty = True

        for col, field in META_COLS.items():

            value = ws.cell(row, col).value

            employee[field] = value

            if value not in ("", None):
                empty = False

        if empty:
            continue

        for col, (benefit, month) in benefit_cols.items():

            value = ws.cell(row, col).value

            if value is None:
                value = 0

            rec = employee.copy()

            rec["Benefit"] = benefit
            rec["Month"] = MONTH_MAP.get(month, month)
            rec["Amount"] = value

            records.append(rec)

    df = pd.DataFrame(records)

    df[["Last Name", "First Name"]] = (
        df["Name"]
        .apply(lambda x: pd.Series(split_name(x)))
    )

    cols = list(df.columns)

    cols.remove("Last Name")
    cols.remove("First Name")

    idx = cols.index("Name") + 1

    cols = (
        cols[:idx]
        + ["Last Name", "First Name"]
        + cols[idx:]
    )

    df = df[cols]

    return df
    
