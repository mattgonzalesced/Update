# -*- coding: utf-8 -*-

import clr
clr.AddReference('System')

from Snippets._rotateutils import collect_data_for_rotation_or_orientation, orient_elements_group
from pyrevit import revit, DB, script

# Get the active document
doc = revit.doc

config = script.get_config("orientation_config")

# Determine if shift-click is being used
if __shiftclick__:
    adjust_tag_position = False  # Temporary override
    adjust_tag_angle = False    # Temporary override
    keep_model_orientation = False
else:
    adjust_tag_position = getattr(config, "tag_position", True)
    adjust_tag_angle = getattr(config, "tag_angle", True)
    keep_model_orientation = getattr(config, "tag_model_orientation", True)

# Step 1: Get the selected elements and filter out pinned ones
selection = revit.get_selection()
filtered_selection = [el for el in selection if isinstance(el, DB.FamilyInstance) and not el.Pinned]

# Step 2: Pre-collect all necessary data before starting the transaction
element_data = collect_data_for_rotation_or_orientation(doc, filtered_selection,adjust_tag_position)

# Step 3: Define the target orientation (Y-axis for 0 degrees)
target_orientation = DB.XYZ(0, -1, 0)

# Step 4: Orient Elements and Adjust Tags in a Single Transaction
with DB.Transaction(doc, "Orient Elements and Adjust Tags") as trans:
    trans.Start()

    for orientation_key, grouped_data in element_data.items():
        orient_elements_group(doc, grouped_data,
                              target_orientation,
                              adjust_tag_position,
                              adjust_tag_angle,
                              keep_model_orientation)
    trans.Commit()
