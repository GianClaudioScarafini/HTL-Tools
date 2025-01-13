# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-07-26
# Version: 1.0.4
# Description: This tool creates a Workset for each Revit or IFC Linked file in accordance with the HTL naming standard. It asks a user to include HTL originator code and zone. It also moves existing links to corresponding Worksets if a link type or instance element is not placed correctly. For Revit 2023+ user will be asked at the end of the process if worksets with no RVT link replaced by a Workset with an updated name should be deleted. This unfortunately due to Revit API limitations can only be done to Editable Worksets. Due to Revit API limitations it is not possible to rename existing Worksets. Links from existing worksets are removed and these Worksets are deleted. This means filters or other settings may not work and should be checked. Script also adds the name of the workset except the prefix to the Name and Mark parameter of a linked model.
# Tested with: Revit 2022+
# Requirements: pyRevit add-in
#
# Since 1.0.1 Workset Name and Mark Added. Error in startswith() corrected. Link prefix added as variable.
# Since 1.0.2 Zone query added.
# Since 1.0.3 Links selector added.
# Since 1.0.4 IFC extension support added. "Mark" parameter amended only for Revit 2023+. Support for multiple instances of the same link added.

import re
# from sys import exit # to use exit() to terminate the script
# Import pyRevit modules
from pyrevit import revit, DB, script, forms

# Get the current document
doc = revit.doc
enable_worksharing = False
# Collects all Revit links in the project
revit_links_collection = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()
revit_links = [ x for x in revit_links_collection ]
revit_links.sort( key=lambda x: x.Name )

# Check Revit version
app = revit.doc.Application
ver = int(app.VersionNumber)

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

def startsWithZLinked(ws_name, extension):
    if ws_name.startswith('Z-Linked '+extension) or ws_name.startswith('Z-Linked-'+extension):
        return True
    else:
        return False

class AllRevitLinks(forms.TemplateListItem):
        @property
        def name(self):
            warning = ""
            workset = revit.query.get_element_workset(self.item).Name
            workset_type_id = self.item.GetTypeId()
            workset_type = doc.GetElement(workset_type_id)
            type_workset = revit.query.get_element_workset(workset_type).Name
            name = self.item.Name.split(" : location")[0]
            if ".rvt".lower() in name.lower():
                ext = "RVT"
            elif ".ifc".lower() in name.lower():
                ext = "IFC"
            else:
                ext = link.Name.split(' : ')[0].split('.')[-1].upper()
                    
            if not startsWithZLinked(workset, ext) or not startsWithZLinked(type_workset, ext):
                warning += " *** NOT Z-LINKED "+ext+" WORKSET!!!"
            if workset != type_workset:
                warning += " *** INSTANCE OR TYPE MISMATCH!!!"
            if warning:
                warning = " >>>"+ warning + " ***"
            
            return name + " | Workset: " + workset + " | Type workset: " + type_workset + warning
        def __bool__(self):
            return self.Name is not None

revit_links_selected = \
        forms.SelectFromList.show(
            [AllRevitLinks(x) for x in revit_links],
            title='Select linked Revit files',
            width=1070,
            button_name='Continue with selected',
            multiselect=True
        )

if revit_links_selected:
    
    all_rvt_link_names = []
    new_workset_dict = {}
    used_workset_names = []
    same_link_name_dict = {}
    linked_file_prefix = 'Z-Linked RVT-'
    extension = "RVT"

    if not enable_worksharing:
        # Create a FilteredWorksetCollector to get all Worksets in the document
        worksets_collection = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
        # Create a dictionary with workset names as keys and worksets as values
        workset_dict = {workset.Name: workset for workset in worksets_collection}

        for workset_name in workset_dict:
            used_workset_names.append(workset_name)

    for link in revit_links:
        if ".rvt" in link.Name.lower():
            link_name = link.Name.split(".rvt")[0]
        elif ".ifc" in link.Name.lower():
            link_name = link.Name.split(".ifc")[0]
        else:
            link_name = link.Name.split(".")[0]
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
    # Query to add Zone to the workset link name
    add_zone = forms.alert(
                    'Include Zone code in the workset name?\n\nPress "Yes" if model is split into multiple zones combined together.', 
                    title="Include Zones?",
                    cancel=True,
                    ok = False,
                    yes = True,
                    no = True
                )
    output = script.get_output()
    for link in revit_links_selected:
        count += 1
        if ".rvt" in link.Name.lower():
            link_name = link.Name.split(".rvt")[0]
            linked_file_prefix = 'Z-Linked RVT-'
            extension = "RVT"
        elif ".ifc" in link.Name.lower():
            link_name = link.Name.split(".ifc")[0]
            linked_file_prefix = 'Z-Linked IFC-'
            extension = "IFC"
        else:
            link_name = link.Name.split(".")[0]
            extension = link.Name.split(' : ')[0].split('.')[-1].upper()
            linked_file_prefix = 'Z-Linked '+extension+'-'
        if link_name not in same_link_name_dict:
            same_link_name_count = all_rvt_link_names.count(link_name)
            if same_link_name_count > 1:
                same_link_name_dict[link_name] = same_link_name_count
        link_workset = doc.GetWorksetTable().GetWorkset(link.WorksetId)
        link_type_id = link.GetTypeId()
        link_type = doc.GetElement(link_type_id)
        link_type_workset = doc.GetWorksetTable().GetWorkset(link_type.WorksetId)
        
        link_workset_name = link_workset.Name
        type_workset_name = link_type_workset.Name
        
        output.print_md( '**'+str(count)+'. Link: ' + link_name +'**' )
        output.print_md( '> Link Workset: ' + link_workset_name  )
        output.print_md( '> Link Type Workset: ' + type_workset_name  )

        # Extract parts from the file name
        # i.e. GSK-HTL-RE-ZZ-M3-A-0001.rvt
        match = re.match(r"(\w+)-(\w+)-(\w+)-(\w+)-(\w+)-(\w+)-(\d+)([\w\d\s]*)?", link_name)
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
        # Group 8: Optional underscore, dash or space (e.g., "-_ ") followed by some text

        if match:
            groups = list(match.groups())
            # Ensure the description is present
            if len(groups) < 8:
                groups.append("")
            _, originator, zone, _, _, discipline, digits, description = groups
            description = description if description is not None else ""
            output.print_md("> Originator: " + originator + " Zone: " +zone+ " Discipline: " + discipline+ " Digits: "+ digits + " Description: " + description)
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
            if zone == 'ZZ' or not add_zone:
                zone = ''
                output.print_md( '> Zone is ZZ or not requested. Skipping: ' + file_zone  )
            else:
                zone = '-' + zone
            instance_name = discipline + originator + zone
            
            similar_link_names = []
            if description:
                output.print_md( '> Description from the end removed: ' + description  )
                base_name = link_name.replace(description, "").strip()
                base_name = base_name.replace(digits, "").strip()
            else:
                base_name = link_name.replace(digits, "").strip()
            output.print_md( '> Base name: ' + base_name  )
            # Check how many links have the same base name. We removed last characters which usually are digits from 0001.
            # There will be always one (or more) the same as the link name is in the loop. Below removes links with the same name.
            rvt_link_names_except_link = [s for s in all_rvt_link_names if s != link_name]
            # Current model file name should be considered in naming new worksets. We add it to the list of link names.
            rvt_link_names_except_link.append(file_name)
            rvt_link_names_with_file_name = rvt_link_names_except_link
            for n in rvt_link_names_with_file_name:
                if n.startswith(base_name):
                    similar_link_names.append(n)

            def find_similar_part_names(desc, part_number, base_name, last_digit, similar_link_names):
                    if int(last_digit) > 1: # then we need to add number to the end anyway
                        if desc != "":
                            desc = desc + " " + last_digit
                        else:
                            desc = last_digit
                    else: # if 1 at the end we need to find if there are more
                    # find if there are two or more Internal model files.
                        similar_part_names = []
                        for p in similar_link_names:
                            if p.startswith(base_name + part_number):
                                similar_part_names.append(p)
                        if len(similar_part_names) > 0: # then there are many parts. We need to add additional number.
                            if desc != "":
                                desc = desc + " " + last_digit
                            else:
                                desc = last_digit
                    return desc
            # this gets last digit from digits
            last_digit = digits[-1]

            # Now we check if this links base name is used many times if similar_names > 0
            if len(similar_link_names) > 0:
                # only for more than 1 we need to add digits at the end.
                # It could be that teh file name is from the same model or two files are linked. In both situations we need to add digits at the end or description.
                output.print_md( '> More than one link with the same base name. Adding digits or description at the end.'  )
                # find what model part is he link
                # if there are more than 2 or more or model doesn't end with 1 adds that number to the end.
                if digits.startswith("1") and discipline == "A": # like ...100001 then this is Internal model
                    output.print_md( '> Internal Model detected with digits starting with 1.'  )
                    digits = find_similar_part_names("Internal", "1", base_name, last_digit, similar_link_names)
                elif digits.startswith("2") and discipline == "A": # like ...200001 then this is Internal model
                    output.print_md( '> Facade Model detected with digits starting with 2.'  )
                    digits = find_similar_part_names("Facade", "2", base_name, last_digit, similar_link_names)
                else:
                    output.print_md( '> Model with 0 or +3'  )
                    digits = find_similar_part_names("", digits[1], base_name, last_digit, similar_link_names)
                instance_name = instance_name + '-' + digits
            else:
                # if there are no other links with the base name and the model file name is not the same it may still be the Internal or the Facade model
                output.print_md('> This link has unique base name.')
                if digits.startswith("1"): # like ...100001 then this is Internal model
                    digits = find_similar_part_names("Internal", "1", base_name, last_digit, similar_link_names)
                    instance_name = instance_name + '-' + digits
                elif digits.startswith("2"): # like ...200001 then this is Internal model
                    digits = find_similar_part_names("Facade", "2", base_name, last_digit, similar_link_names)
                    instance_name = instance_name + '-' + digits
                else:
                    if int(last_digit) > 1: # then we need to add number to the end anyway
                        digits = digits[1] + last_digit
                        instance_name = instance_name + '-' + digits
                    else:
                        instance_name = instance_name
            workset_name = linked_file_prefix + instance_name
        else:
            output.print_md( '> Link name does not match the naming standard. Adding whole name to workset name.'  )
            workset_name = linked_file_prefix + link_name
            instance_name = link_name
        output.print_md( '> New Workset name: ' + workset_name  )
        # Now we need to check if a workset with this name already exists for this link
        existing_worksets = [] # with this link name
        if not enable_worksharing:
            for name in workset_dict.keys():
                # Link workset name must start with "Z-Linked RVT-X-XXX-XX-000X" (or IFC)
                if name.startswith(workset_name):
                    output.print_md( "> Workset with this base name exists and should be used: "+workset_name )
                    existing_worksets.append(name)

            # If two link instances are in the model one may already have been processed and a new workset was created. We need to check if the new Workset name is in new workset names.
            for name in new_workset_dict:
                if name.startswith(workset_name):
                    output.print_md( "> A new Workset with this base name already was created and should be used: "+workset_name )
                    existing_worksets.append(name)
        if same_link_name_count > 1:
            # we check what number dictionary holds for this link name. We decrease it with each instance so when it reaches 1 we know that this is the last instance. 
            instance_number = same_link_name_dict[link_name]
            instance_name += " (no." + str(instance_number)+")"
            output.print_md( '> Instance no.{} of this link.'.format(instance_number))
            same_link_name_dict.update({link_name:instance_number-1})
            
        if len(existing_worksets) == 0:
            # Workset needs to be created
            with revit.Transaction('Create Workset for linked model.'):
                try:
                    newWs = DB.Workset.Create(revit.doc, workset_name)
                    worksetParam = \
                        link.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                    worksetParam.Set(newWs.Id.IntegerValue)
                    worksetTypeParam = \
                        link_type.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                    worksetTypeParam.Set(newWs.Id.IntegerValue)
                    # Sets link Name and Mark to make it the same as the link (this helps identify  the original link if it's duplicated)
                    worksetName = \
                        link.Parameter[DB.BuiltInParameter.RVT_LINK_INSTANCE_NAME]
                    worksetName.Set(instance_name)
                    # "Mark" links parameter was added in Revit 2023
                    if ver >= 2023:
                        worksetMark = \
                            link.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
                        worksetMark.Set(instance_name)
                    output.print_md( '> New Workset created'  )
                    new_workset_dict[workset_name] = newWs
                except Exception as e:
                    print('> Creating or setting up workset: {} created error: {}'.format(workset_name,e))
        elif len(existing_worksets) >= 1:
            # Workset/s already exists. For more than one first will be used. If not in existing workset dict then it was created in this session and will be used for more then one link. 
            output.print_md( "> Link Type or instance Workset will be corrected if incorrect.")
            if existing_worksets[0] in workset_dict:
                Ws = workset_dict[existing_worksets[0]]
            else:
                Ws = new_workset_dict[existing_worksets[0]]
            if not existing_worksets[0].startswith(link_workset_name):
                with revit.Transaction('Set correct Workset for linked model instance'):
                    try:
                        worksetParam = \
                            link.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                            # DB.BuiltInParameter.ELEM_PARTITION_PARAM is Workset
                        worksetParam.Set(Ws.Id.IntegerValue)
                        # Sets link Name and Mark to make it the same as the link (this helps identify  the original link if it's duplicated)
                        worksetName = \
                            link.Parameter[DB.BuiltInParameter.RVT_LINK_INSTANCE_NAME]
                        worksetName.Set(instance_name)
                        # "Mark" links parameter was added in Revit 2023
                        if ver >= 2023:
                            worksetMark = \
                                link.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
                            worksetMark.Set(instance_name)
                        output.print_md( "> Link instance Workset was corrected.")
                    except Exception as e:
                        print('> Workset: {} could not be set to {} link\n> Error: {}'.format(workset_name, extension, e))
            if not existing_worksets[0].startswith(type_workset_name):
                with revit.Transaction('Set correct Workset for linked model Type'):
                    try:
                        worksetTypeParam = \
                            link_type.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
                        worksetTypeParam.Set(Ws.Id.IntegerValue)
                        output.print_md( "> Link Type Workset was corrected.")
                    except Exception as e:
                        print('> Workset: {} could not be set to {} link type\n> Error: {}'.format(workset_name, extension,e))
            used_workset_names.append(existing_worksets[0])
        else:
            # More than one workset with this new workset name beginning exists
            # First one will be used
            pass

    unused_workset_names = []
    if not enable_worksharing:
        output.print_md('---')
        output.print_md( "###Z-LINKED WORKSETS ELEMENTS CHECK" )

        for ws_name, ws in workset_dict.items():
            # Check if the workset name starts with "Z-Linked" but not "Z-Linked CAD" or "Z-Linked-CAD"
            if ws_name.startswith("Z-Linked") and not (ws_name.startswith("Z-Linked CAD") or ws_name.startswith("Z-Linked-CAD")):
                output.print_md("   **Workset: {}**".format(ws_name))
                # Use a filter to collect elements on the current workset
                workset_filter = DB.ElementWorksetFilter(ws.Id)
                workset_elements = DB.FilteredElementCollector(doc).WherePasses(workset_filter).WhereElementIsNotElementType().ToElements()

                # Count elements by category
                category_counts = {}
                for elem in workset_elements:
                    category = elem.Category
                    if category is not None:
                        category_name = category.Name
                        category_counts[category_name] = category_counts.get(category_name, 0) + 1

                # Output category counts
                for category_name, count in category_counts.items():
                    output.print_md("> - {}: {}".format(category_name, count))

                # Check for links in the workset
                has_link = False
                for elem in workset_elements:
                    if isinstance(elem, DB.RevitLinkInstance):
                        has_link = True
                        output.print_md("   Z-Linked Workset {} contains a Link!<br><br>".format(ws_name))
                        break

                if not has_link:
                    output.print_md("**Z-Linked Workset {} does not contain any Link! UNUSED AND CAN BE DELETED.**<br><br>".format(ws_name))
                    unused_workset_names.append(ws_name)
    
    if unused_workset_names:
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
                                    print('> Workset: {} could not be deleted\nError: {}'.format(ws.Name,e))
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
