# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-07-26
# Version: 1.0.1
# Description: This tool creates a Workset for each Revit Linked file in accordance with the HTL naming standard. It asks a user to include HTL originator code or not. It also moves existing links to corresponding Worksets if a link type or instance element is not placed correctly. For Revit 2023+ user will be asked at the end of the process if worksets with no RVT link replaced by a Workset with an updated name should be deleted. This unfortunately due to Revit API limitations can only be done to Editable Worksets.
# Tested with: Revit 2022+
# Requirements: pyRevit add-in
#
# Since 1.0.1 Workset Name and Mark Added. Error in startswith() corrected. Link prefix added as variable.

import re
# from sys import exit # to use exit() to terminate the script
# Import pyRevit modules
from pyrevit import revit, DB, script, forms

# Get the current document
doc = revit.doc
enable_worksharing = False
# Collects all Revit links in the project
revit_links = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()

# if there are no links there is no point to continue 
if not revit_links:
    # https://docs.pyrevitlabs.io/reference/pyrevit/forms/#pyrevit.forms.alert
    forms.alert('No Revit links found in the project.', title="Workset Info", exitscript=True)

if not doc.IsWorkshared and doc.CanEnableWorksharing:
    enable_worksharing = forms.alert(
        'Current project is not workshared for collaboration.\n\nWould you like to enable worksharing?', 
        title="Enable Worksharing?",
        cancel=True,
        ok = False,
        yes = True
    )
    if enable_worksharing:
        revit.doc.EnableWorksharing('Shared Levels and Grids', 'Workset1')

elif not doc.IsWorkshared and not doc.CanEnableWorksharing:
    forms.alert("This is not a worksahred project and it is not possible to enable worksharing.", title="Worksharing not possible", exitscript=True)

if not enable_worksharing:
    # Create a FilteredWorksetCollector to get all Worksets in the document
    worksets_collection = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
    # Create a dictionary with workset names as keys and worksets as values
    workset_dict = {workset.Name: workset for workset in worksets_collection}

all_rvt_link_names = []
new_workset_names = []
used_workset_names = []
linked_file_prefix = 'Z-Linked RVT-'

for link in revit_links:
    link_name = link.Name.split(".rvt")[0]
    all_rvt_link_names.append(link_name)
count = 0
# Query to add Originator to the workset link name
add_originator = forms.alert(
                'Include HTL in the workset name?\n\nTypical workset name for a an architectural link starts with Z-Linked RVT A-... followed by the originator code i.e. HTL.\n\nPress "Yes" if models of other architectural companies will be linked to this model.', 
                title="Include HTL?",
                cancel=True,
                ok = False,
                yes = True,
                no = True
            )
for link in revit_links:
    count += 1
    link_name = link.Name.split(".rvt")[0]
    link_workset = doc.GetWorksetTable().GetWorkset(link.WorksetId)
    link_type_id = link.GetTypeId()
    link_type = doc.GetElement(link_type_id)
    link_type_workset = doc.GetWorksetTable().GetWorkset(link_type.WorksetId)
    
    link_workset_name = link_workset.Name
    type_workset_name = link_type_workset.Name
    output = script.get_output()
    output.print_md( '**'+str(count)+'. Link: ' + link_name +'**' )
    output.print_md( '> Link Workset: ' + link_workset_name  )
    output.print_md( '> Link Type Workset: ' + type_workset_name  )

    # Extract parts from the file name
    # i.e. GSK-HTL-RE-ZZ-M3-A-0001.rvt
    match = re.match(r"(\w+)-(\w+)-(\w+)-(\w+)-(\w+)-(\w+)-(\d+)", link_name)
    # re.match: This function searches for a match at the beginning of the string (file_name). It returns a match object if the pattern is found, or None if no match is found.
    # (\w+): This captures one or more word characters (letters, digits, or underscores). The parentheses indicate a capturing group.
    # -: This matches the hyphen character literally.
    # (\d+): This captures one or more digits.
    #  The pattern consists of six groups, each separated by hyphens. These groups correspond to the different parts of the file name you want to extract.
    # Group 1: First part (e.g., "GSK")
    # Group 2: Second part - Originator (e.g., "HTL")
    # Group 3: Third part - Zone (e.g., "RE")
    # Group 4: Fourth part (e.g., "ZZ")
    # Group 5: Fifth part (e.g., "M3")
    # Group 6: Sixth part - Discipline (e.g., "A")
    # Group 7: Digits (e.g., "0001")

    if match:
        _, originator, zone, _, _, discipline, digits = match.groups()
        # match.groups(): This method returns a tuple containing all the captured groups from the regular expression match. In our case, it corresponds to the seven groups defined in the pattern.
        # _: This is a placeholder variable. It is used to ignore specific groups.
        # Add originator to the workset name
        if discipline == 'A' and not add_originator:
            originator = ''
        else:
            originator = '-' + originator
        # zone should not be used if "ZZ" or the same as the file name.
        file_name = doc.Title
        # Extract the third part from the file name
        groups = re.match(r"(\w+)-(\w+)-(\w+)", file_name)
        file_zone = ''
        if groups:
            file_zone = groups.group(3)
        if zone == 'ZZ' or zone == file_zone:
            zone = ''
            output.print_md( '> Zone is the same as the file name or ZZ. Skipping: ' + file_zone  )
        else:
            zone = '-' + zone
        instance_name = discipline + originator + zone
        workset_name = linked_file_prefix + instance_name

        similar_names = 0
        base_name = link_name.replace(digits, "").strip()
        output.print_md( '> Base name:' + base_name  )
        # Check how many links have the same base name. We removed last characters which usually are digits from 0001.
        for n in all_rvt_link_names:
            if n.startswith(base_name):
                similar_names += 1
        # Now we check if the workset name was already created if similar_names > 1
        if similar_names > 1:
            # only for more than 1 we need to add digits at the end.
            output.print_md( '> More than one link with the same base name. Adding digits at the end.'  )
            workset_name = workset_name + '-' + digits
    else:
        output.print_md( '> Link name does not match the naming standard. Adding whole name to workset name.'  )
        workset_name = linked_file_prefix + link_name
        instance_name = link_name
    output.print_md( '> New Workset name: ' + workset_name  )
    # Now we need to check if a workset with this name already exists for this link
    existing_workset = [] # with this link name
    if not enable_worksharing:
        for name in workset_dict.keys():
            # Link workset name must start with "Z-Linked RVT-XXX-XX-000X"
            if name.startswith(workset_name):
                output.print_md( "> Workset with this base name exists and should be used: "+workset_name )
                existing_workset.append(name)
    if len(existing_workset) == 0:
        # Workset needs to be created
        with revit.Transaction('Create Workset for linked model'):
            try:
                newWs = DB.Workset.Create(revit.doc, workset_name)
                worksetParam = \
                    link.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                worksetParam.Set(newWs.Id.IntegerValue)
                worksetTypeParam = \
                    link_type.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                worksetTypeParam.Set(newWs.Id.IntegerValue)
                # Sets link Name and MArk to make it the same as the link (this helps identify  the original link if it's duplicated)
                worksetName = \
                    link.Parameter[DB.BuiltInParameter.RVT_LINK_INSTANCE_NAME]
                worksetName.Set(instance_name)
                worksetMark = \
                    link.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
                worksetMark.Set(instance_name)
                output.print_md( '> New Workset created'  )
                new_workset_names.append(workset_name)
            except Exception as e:
                print('Workset: {} already exists\nError: {}'.format(workset_name,e))
    elif len(existing_workset) >= 1:
        # Workset/s already exists. For more than one first will be used.
        output.print_md( "> RVT link Type or instance Workset will be corrected if incorrect.")
        Ws = workset_dict[existing_workset[0]]
        if not existing_workset[0].startswith(link_workset_name):
            with revit.Transaction('Set correct Workset for linked model instance'):
                try:
                    worksetParam = \
                        link.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                        # DB.BuiltInParameter.ELEM_PARTITION_PARAM is Workset
                    worksetParam.Set(Ws.Id.IntegerValue)
                    # Sets link Name and MArk to make it the same as the link (this helps identify  the original link if it's duplicated)
                    worksetName = \
                        link.Parameter[DB.BuiltInParameter.RVT_LINK_INSTANCE_NAME]
                    worksetName.Set(instance_name)
                    worksetMark = \
                        link.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
                    worksetMark.Set(instance_name)
                    output.print_md( "> RVT link instance Workset was corrected.")
                except Exception as e:
                    print('Workset: {} could not be set to RVT link\nError: {}'.format(workset_name,e))
        if not existing_workset[0].startswith(type_workset_name):
            with revit.Transaction('Set correct Workset for linked model Type'):
                try:
                    worksetTypeParam = \
                        link_type.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                    worksetTypeParam.Set(Ws.Id.IntegerValue)
                    output.print_md( "> RVT link Type Workset was corrected.")
                except Exception as e:
                    print('Workset: {} could not be set to RVT link type\nError: {}'.format(workset_name,e))
        used_workset_names.append(existing_workset[0])
    else:
        # More than one workset with this new workset name beginning exists
        # First one will be used
        pass

unused_workset_names = []
if not enable_worksharing:
    for ws_name, ws in workset_dict.items():
        if ws_name.startswith('Z-Linked RVT') or ws_name.startswith('Z-Linked-RVT'):
            if ws_name not in used_workset_names and ws_name not in new_workset_names:
                unused_workset_names.append(ws_name)
 
if unused_workset_names:
    app = __revit__.Application
    ver = int(app.VersionNumber)
    unused_workset_names = ',\n'.join(unused_workset_names)
    if ver >=2023:
        delete_empty_link_worksets = forms.alert(
                'Delete empty Z-Linked... Worksets?\n'+unused_workset_names+'\n\n(Elements from these Worksets will be moved to a default Workset.)', 
                title="Enable Worksharing?",
                cancel=True,
                ok = False,
                yes = True
            )
        if delete_empty_link_worksets:
            delete_worksets = []
            deleted_worksets = []
            not_editable_worksets = []
            for ws in workset_dict.values():
                if ws.IsDefaultWorkset:
                    workset_id = ws.Id
                if ws.Name in unused_workset_names:
                    delete_worksets.append(ws)
            if delete_worksets:
                # Create DeleteWorksetSettings with DeleteWorksetOption.MoveElementsToWorkset not DeleteElements
                delete_settings = DB.DeleteWorksetSettings(DB.DeleteWorksetOption.MoveElementsToWorkset, workset_id)
                with revit.Transaction('Delete Unused Worksets'):
                    for ws in delete_worksets:
                        if ws.IsEditable:
                            try:
                                ws_table = doc.GetWorksetTable()
                                ws_table.DeleteWorkset(doc, ws.Id, delete_settings)
                                deleted_worksets.append(ws.Name)
                            except Exception as e:
                                print('Workset: {} could not be deleted\nError: {}'.format(ws.Name,e))
                        else:
                            not_editable_worksets.append(ws.Name)
                    if deleted_worksets:
                        deleted = 'Deleted Workset:\n'+',\n'.join(deleted_worksets)
                        if not_editable_worksets:
                            deleted += '\n\n'
                    else:
                        deleted = ''
                    if not_editable_worksets:
                        not_deleted = 'Make these worksets Editable and run the script again or delete manually:\n'+',\n'.join(not_editable_worksets)+'\n\n(It is not possible to delete Worksets using API id they are not Editable)'
                    else:
                        not_deleted = ''
                    if deleted or not_deleted:
                        forms.alert('{}{}'.format(deleted,not_deleted))
    else:
        forms.alert('It is not possible to delete Worksets using API in Revit 2022 and earlier.\n\nDelete following Worksets manually:\n{}'.format(unused_workset_names)) 
