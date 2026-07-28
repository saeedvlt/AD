
# Replace the material adjustment logic with the following inside convert(),
# before the row loop:

pending_adjustment = {}

# ------------------------------------------------------------------
# Inside the row loop, immediately after:
# label = clean_text(...)
# normalized_label = label.casefold()
# ------------------------------------------------------------------

adjustment_status = "Standard"

if current_section == "Material":

    # Adjustment row
    if normalized_label == "adjust to actual":
        adjustment_status = "Adjustment"

        # remember the previous material item
        if previous_material in ("Fabrications", "Customer Material"):
            pending_adjustment[previous_material] = True

    elif label in ("Fabrications", "Customer Material"):

        if pending_adjustment.get(label, False):
            adjustment_status = "Net"
            pending_adjustment[label] = False
        else:
            adjustment_status = "Original"

        previous_material = label

# For every other Material row remember the label
elif current_section == "Material":
    previous_material = label


# ------------------------------------------------------------------
# Before the row loop initialize:
# ------------------------------------------------------------------

previous_material = None
pending_adjustment = {
    "Fabrications": False,
    "Customer Material": False,
}

# ------------------------------------------------------------------
# In records.append(...)
# ------------------------------------------------------------------

"Adjustment Status": adjustment_status,

# ------------------------------------------------------------------
# IMPORTANT
# ------------------------------------------------------------------
# Remove this line from is_derived_line():
#
# or normalized.startswith("adjust to actual")
#
# and return:
#
# return data
#
# instead of:
#
# return add_percentages(data)
