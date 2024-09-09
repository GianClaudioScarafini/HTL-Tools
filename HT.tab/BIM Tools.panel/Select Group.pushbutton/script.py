# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-09-02
# Version: 1.0.0
# Name: Select Element Group
# Description: First select an element belonging to a Group. By clicking on this tool a Group hosting this element will be selected.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

import clr
import os
# Import pyRevit modules
from pyrevit import revit, DB
# Imports from Revit API
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List  # Import List from .NET collections

# Get Revit application and active document
uidoc = revit.uidoc
doc = revit.doc

# Function to display message to the user
def show_message(title, message):
    TaskDialog.Show(title, message)

# Function to select groups based on the selected elements
def select_groups_from_elements(selected_elements):
    group_ids = set()
    
    for element in selected_elements:
        group_id = element.GroupId  # Try to get the group id of the element
        if group_id != DB.ElementId.InvalidElementId:
            group_ids.add(group_id)
    
    if not group_ids:
        show_message("Selection Error", "Selected element(s) do not belong to any group.")
        return
    
    if len(group_ids) > 1:
        show_message("Selection Error", "Selected elements belong to different groups.")
        return
    
    # Get the group element
    group_element = doc.GetElement(group_ids.pop())
    
    if group_element:
         # Create an ICollection[ElementId] to pass to SetElementIds
        element_ids = List[ElementId]([group_element.Id])
        uidoc.Selection.SetElementIds(element_ids)
        # show_message("Group Selected", f"Group '{group_element.Name}' has been selected.")

# Main execution
try:
    # Get selected elements
    selected_ids = uidoc.Selection.GetElementIds()
    selected_elements = [doc.GetElement(id) for id in selected_ids]

    # Check if no elements are selected
    if not selected_elements:
        # Allow user to select one element if none is initially selected
        show_message("Select Element", "Please select an element in the model.")
        selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select an element")
        selected_elements = [doc.GetElement(selected_ref.ElementId)]
    
    # Check again in case no element is selected after the prompt
    if not selected_elements:
        show_message("Error", "No element was selected.")
    else:
        # Process selection to identify the group
        select_groups_from_elements(selected_elements)

except Exception as e:
    # Display any unexpected errors to the user
    show_message("Error", str(e))
