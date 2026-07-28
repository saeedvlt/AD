
# NOTE:
# I can't safely reconstruct your entire converter from scratch without
# risking breaking unrelated logic. Instead, replace ONLY the row-processing
# section in your convert() function with the logic below.

# Before the row loop:
material_seen = {
    "Fabrications": False,
    "Customer Material": False,
}

for row in range(1, ws.max_row + 1):

    label = clean_text(ws.cell(row, 1).value)
    normalized_label = label.casefold()

    # Detect section headers
    if normalized_label in SECTION_NAMES:
        current_section = SECTION_NAMES[normalized_label]
        continue

    # Ignore Gross Margin
    if current_section == "Gross Margin":
        continue

    if not label or current_section is None:
        continue

    # Skip calculated rows
    if is_derived_line(label):
        continue

    # Skip all Adjust to Actual rows
    if normalized_label == "adjust to actual":
        continue

    # Material section:
    # Keep only the SECOND occurrence of Fabrications and Customer Material
    if current_section == "Material":

        if label in ("Fabrications", "Customer Material"):

            if not material_seen[label]:
                material_seen[label] = True
                continue

    if not has_monthly_amount(ws, row, month_columns):
        continue

    for column, period in month_columns.items():

        value = ws.cell(row, column).value

        amount = float(value) if is_number(value) else 0.0

        records.append({
            "Location": location,
            "Territory": TERRITORIES[location],
            "Section": current_section,
            "Budget Category": budget_category(label),
            "Line Item": label,
            "Period": period,
            "Month": period.strftime("%m - %b"),
            "Amount": amount,
        })

# Also delete:
# - adjusted_line_item()
# - adjustment_target
# - previous_label
# - Adjusted Line Item column
# - any percentage code
