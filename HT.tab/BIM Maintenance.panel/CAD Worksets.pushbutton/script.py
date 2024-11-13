# -*- coding: UTF-8 -*-
# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-08-26
# Version: 1.0.1
# Description: This tool checks if all CAD files (instance elements and Types) are set to the "Z-Linked CAD" workset and allows correction if they are not.
# Tested with: Revit 2022+
# Requirements: pyRevit add-in

from pyrevit import revit, DB, forms

# Get the current document
doc = revit.doc

# Define the workset name that should be used for CAD files
correct_workset_name = ["Z-Linked CAD","Z-Linked-CAD"]

# Collect all CAD instances (both linked and imported)
cad_instances = DB.FilteredElementCollector(doc)\
                  .OfClass(DB.ImportInstance)\
                  .WhereElementIsNotElementType()\
                  .ToElements()

# if there are no links there is no point to continue 
if not cad_instances:
    # https://docs.pyrevitlabs.io/reference/pyrevit/forms/#pyrevit.forms.alert
    forms.alert('No CAD links found in the project.', title="CAD Links Info", exitscript=True)

if not doc.IsWorkshared and doc.CanEnableWorksharing:
    forms.alert(
        'Current project is not workshared for collaboration.', 
        title="Worksharing Info", exitscript=True
    )

# Initialize lists to store CAD files with incorrect workset assignments
incorrect_workset_cad = []

# Iterate over CAD instances to check their workset assignments
for cad_instance in cad_instances:
    # Get the workset name for the CAD instance
    cad_workset = revit.query.get_element_workset(cad_instance).Name
    if not cad_workset.startswith("View"):
        # Get the CAD type (symbol) associated with the instance
        cad_type_id = cad_instance.GetTypeId()
        cad_type = doc.GetElement(cad_type_id)
        cad_type_workset = revit.query.get_element_workset(cad_type).Name

        # Get CAD file name
        cad_name = cad_instance.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()
        
        # Get the creator of the CAD instance
        cad_creator = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, cad_instance.Id).Creator

        # Check if the worksets are not set to the correct workset
        if cad_workset not in correct_workset_name or cad_type_workset not in  correct_workset_name:
            id = cad_instance.Id.ToString()
            # If either workset is incorrect, add it to the list
            if cad_workset not in correct_workset_name and cad_type_workset not in correct_workset_name:
                incorrect_workset_cad.append(
                    "{} - Instance & Type (Current: {}, {}) - Creator: {} - Id: {}".format(cad_name, cad_workset, cad_type_workset, cad_creator, id))
            elif cad_workset not in  correct_workset_name:
                incorrect_workset_cad.append(
                    "{} - Instance (Current: {}) - Creator: {} - Id: {}".format(cad_name, cad_workset, cad_creator, id))
            elif cad_type_workset not in correct_workset_name:
                incorrect_workset_cad.append(
                    "{} - Type (Current: {}) - Creator: {} - Id: {}".format(cad_name, cad_type_workset, cad_creator, id))

# Display the result
if not incorrect_workset_cad:
    forms.alert('All CAD files are on the correct workset: Z-Linked CAD', title="CAD Workset Check", exitscript=True)
else:
    # If there are CAD files with incorrect worksets, show a list
    selected_options = forms.SelectFromList.show(
        sorted(incorrect_workset_cad),
        title="Select CAD files on incorrect Worksets to correct:",
        width=900,
        button_name='Correct Workset for these selected CAD files',
        multiselect=True
    )

    if selected_options:
        # Get all worksets in the document
        worksets = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
        correct_workset_Id = None
        # Loop through all worksets
        for workset in worksets:
            # If the workset name matches the given name
            if workset.Name in correct_workset_name:
                # Return the workset Id
                correct_workset_Id = workset.Id.IntegerValue
        # Correct the worksets for the selected CAD files
        with revit.Transaction("Correct CAD Worksets"):
            # Create workset Z-Linked CAD if it doesn't exist:
            if not correct_workset_Id:
                new_CAD_workset = DB.Workset.Create(revit.doc, correct_workset_name[0])
                correct_workset_Id = new_CAD_workset.Id.IntegerValue
            for selected in selected_options:
                # Extract the CAD file name and workset type from the selected string
                cad_name = selected.split(" - ")[0]
                cad_type = "Instance" if "Instance" in selected else "Type"

                # Find the corresponding CAD instance by name
                for cad_instance in cad_instances:
                    instance_cad_name = cad_instance.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()

                    if instance_cad_name == cad_name:
                        # Set the correct workset for the CAD instance or type
                        if cad_type == "Instance":
                            workset_param = cad_instance.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                            workset_param.Set(correct_workset_Id)
                        elif cad_type == "Type":
                            cad_type_id = cad_instance.GetTypeId()
                            cad_type_element = doc.GetElement(cad_type_id)
                            workset_param = cad_type_element.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                            workset_param.Set(correct_workset_Id)

        forms.alert('Selected CAD files have been corrected to the Z-Linked CAD workset.', title="Correction Complete")
