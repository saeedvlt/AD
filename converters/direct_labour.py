#" labour_summary.py Converter for the "Salaries - Head Count Summary" sheet. "


import pandas as pd 

def convert(input_file):
    sheet = "Salaries - Head Count Summary"
    
    # Locate header row dynamically
    preview = pd.read_excel(input_file, sheet_name=sheet, header=None)

    header_row = None
    for i, row in preview.iterrows():
        if "Line Item" in row.astype(str).str.strip().tolist():
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not locate 'Line Item'.")

    df = pd.read_excel(
        input_file,
        sheet_name=sheet,
        header=header_row
    )

    df = df.rename(columns={df.columns[0]: "Line Item"})
    df = df[df["Line Item"].notna()]
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    # Stop before Headcount total
    stop = df[df["Line Item"].astype(str).str.strip().eq("Headcount")]
    if not stop.empty:
        df = df.iloc[:stop.index[0]]

    long_df = df.melt(
        id_vars=["Line Item"],
        var_name="Metric",
        value_name="Value"
    )

    long_df = long_df.dropna(subset=["Value"])

    months = [
        "Jan","Feb","Mar","Apr","May","Jun",
        "Jul","Aug","Sep","Oct","Nov","Dec"
    ]

    long_df["Month"] = long_df["Metric"].apply(
        lambda x: x if x in months else "Overall"
    )

    return long_df

