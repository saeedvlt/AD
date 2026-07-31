import pandas as pd

def convert(input_file): 
    
    """ Converts the Labour Summary’ sheet into a normalized database.

    "Output columns:
        Section
        Department
        Month
        Value """

    SHEET = "Labour Summary"
    MONTHS = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]

    raw = pd.read_excel(input_file, sheet_name=SHEET, header=None)

    # ----------------------------------------------------
    # Locate the month header row dynamically
    # ----------------------------------------------------
    month_row = None

    for i, row in raw.iterrows():
        vals = [str(v).strip() for v in row.tolist()]
        if {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}.issubset(set(vals)):
            month_row = i
            break

    if month_row is None:
        raise ValueError("Unable to locate month header row.")

    month_cols = {}

    for c in raw.columns:
        val = str(raw.iloc[month_row, c]).strip()
        if val in MONTHS:
            month_cols[c] = val

    records = []

    current_section = None

    skip_words = (
        "total",
        "grand total",
        "actual",
        "budget",
        "variance"
    )

    for r in range(month_row + 1, len(raw)):

        first = raw.iloc[r, 0]

        if pd.isna(first):
            continue

        first = str(first).strip()

        if first == "":
            continue

        # Section headers
        if first.endswith(":"):
            current_section = first[:-1].strip()
            continue

        if current_section is None:
            continue

        lower = first.lower()

        if any(word in lower for word in skip_words):
            continue

        department = first

        for col, month in month_cols.items():

            value = raw.iloc[r, col]

            if pd.isna(value):
                continue

            records.append({
                "Section": current_section,
                "Department": department,
                "Month": month,
                "Value": value
            })

    df = pd.DataFrame(records)

    return df
