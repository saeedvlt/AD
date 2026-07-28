Based on the workbook structure, use the following logic:

1.  Remove:

    -   percent()
    -   add_percentages()
    -   return add_percentages(data)

    Replace with: return data

2.  Add a new column: “Adjustment Status”

3.  Material logic

Before looping rows:

    material_seen = {}

Inside the row loop, after determining current_section and before
appending records:

    status = "Standard"

    if current_section == "Material":
        if label == "Adjust to Actual":
            status = "Adjustment"
        elif label in BUDGET_CATEGORIES or label in (
            "Die Sets",
            "Cast Die Sets",
            "Stght Fwd Die Sets",
            "Machined Steel",
            "Bolster Plates",
            "Ground Steel",
            "Rough Steel",
            "Fabrications",
            "Components",
            "Customer Material",
        ):
            material_seen[label] = material_seen.get(label, 0) + 1
            status = "Original" if material_seen[label] == 1 else "Net"

    elif current_section in ("Labour", "Overhead"):
        status = "Adjustment" if label == "Adjust to Actual" else "Standard"

4.  Include in each record:

    “Adjustment Status”: status,

5.  Add “Adjustment Status” to the columns list.

This matches the workbook layout: Original rows Adjust to Actual Net
rows

for: - Die Sets - Plate - Fabrications - Customer Material

Labour and Overhead keep the Adjust to Actual rows tagged as
“Adjustment”.
