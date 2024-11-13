# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-06-19
# Version: 1.0.4
# Description: Removes identical instances from the model omitting these in Groups and these which have any dependent elements connected to one of the duplicated element. It also deletes duplicated elements on the Design Options leaving an element in the Main Model only unless elements in the Design Option are in Groups or have dependent elements. At the end script shows a summary with number and IDs of deleted elements and a detailed summary of all elements showing IDs, which were deleted, which have what dependent elements and which are in the Design Options.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

from sys import exit

# Import pyRevit modules
from pyrevit import revit, DB, script, forms, framework

class DeleteBetweenOptionSets(forms.TemplateUserInputWindow):
    xaml_source = script.get_bundle_file('options.xaml')

    def _setup(self, **kwargs):
        message = kwargs.get('message', 'Pick a command option:')
        self.message_label.Content = message

        for option in self._context:
            my_button = framework.Controls.Button()
            my_button.Content = option
            my_button.Click += self.process_option
            self.button_list.Children.Add(my_button)
        self._setup_response()

    def _setup_response(self, response=None):
        self.response = response

    def _get_active_button(self):
        buttons = []
        for button in self.button_list.Children:
            if button.Visibility == framework.Windows.Visibility.Visible:
                buttons.append(button)
        if len(buttons) == 1:
            return buttons[0]
        else:
            for x in buttons:
                if x.IsFocused:
                    return x

    def handle_click(self, sender, args):
        """Handle mouse click."""
        self.Close()

    def handle_input_key(self, sender, args):
        """Handle keyboard inputs."""
        if args.Key == framework.Windows.Input.Key.Escape:
            self.Close()
        elif args.Key == framework.Windows.Input.Key.Enter:
            self.process_option(self._get_active_button(), None)

    def process_option(self, sender, args):
        """Handle click on command option button."""
        self.Close()
        if sender:
            self._setup_response(response=sender.Content)

opts = ['Yes', 'No',  'Cancel & Exit']
selected_option = DeleteBetweenOptionSets.show(
    opts,
    response = 'Yes',
    message='Delete identical instances between Design Option Sets?\n\n(Sets are ordered A to Z. Elements in the latest Set will be retained)\n',
    title = 'Delete between Design Option Sets',#not doing anything
)

if selected_option == 'Cancel & Exit':
    exit()

# Get the current document
doc = revit.doc
warnings = doc.GetWarnings()

# Filter out warnings related to identical instances
identical_instance_warnings = [
    warning for warning in warnings 
    if warning.GetFailureDefinitionId() == DB.BuiltInFailures.PerformanceFailures.DuplicateInstances or 
       warning.GetFailureDefinitionId() == DB.BuiltInFailures.OverlapFailures.DuplicateInstances
]


# DB.BuiltInFailures.PerformanceFailures.DuplicateInstances is a GUID representing this warning:
# There are duplicate instances of the same family in the same place. Delete duplicate instances to improve performance.

# DB.BuiltInFailures.OverlapFailures.DuplicateInstances is a GUID representing this warning:
# There are identical instances in the same place. This will result in double counting in schedules.
info = ""

# Create an element filter that accepts all elements
element_type_filter = DB.ElementIsElementTypeFilter(True)
element_instance_filter = DB.ElementIsElementTypeFilter(False)
element_filter = DB.LogicalOrFilter(element_type_filter, element_instance_filter)

checked_element_ids = set()
elements_ids_to_delete = set()
grouped_warnings_ids = []
InvalidElementId = DB.ElementId.InvalidElementId

# this creates lists of element ids which are duplicated
for warning in identical_instance_warnings:
    element_ids = warning.GetFailingElements()
    #info += warning.GetDescriptionText()+"\n" # not necessary
    # there are always two elements
    # if there are more then 2 then one warning will be created for each possible pair of elements in the same place
    if element_ids[0] in checked_element_ids or element_ids[1] in checked_element_ids:
        for list in grouped_warnings_ids:
            # added to avoid infinite loop
            add_element0 = False
            add_element1 = False
            for id in list:
                if element_ids[0] == id:
                    # the other id represents another duplicated element
                    add_element1 = True
                if element_ids[1] == id:
                    add_element0 = True
            if add_element0:
                # check necessary because other element could be in a different warning
                if element_ids[0] not in list:
                    list.append(element_ids[0])
                    checked_element_ids.add(element_ids[0])
            if add_element1:
                # check necessary because other element could be in a different warning
                if element_ids[1] not in list:
                    list.append(element_ids[1])
                    checked_element_ids.add(element_ids[1])
    else:
        checked_element_ids.update(element_ids)
        warning_ids = [element_ids[0],element_ids[1]]
        grouped_warnings_ids.append(warning_ids)

def check_host_and_groupId( element ):
    # check if element has a host parameter
    try:
        if element.Host:
            host = True
        else:
            host = None
    except:
        host = None
    # check if element has a parameter SuperComponent what means it's nested.
    try:
        if element.SuperComponent:
            super_component = True
        else:
            super_component = None
    except:
        super_component = None

    if host or super_component is not None:
        global info  # declare 'info' as global
        design_option_text = ""
        design_option = element.DesignOption
        if design_option:
            designOptionSetId = design_option.get_Parameter(DB.BuiltInParameter.OPTION_SET_ID).AsElementId()
            designOptionSet = doc.GetElement(designOptionSetId)
            DSName = designOptionSet.Name
            design_option_text = "(" + element.DesignOption.Name+" in " + DSName + ")"
        info += ("* {}Element ID: {} - Category: {} - Name: {} is nested.\n".format(design_option_text, element.Id.IntegerValue,element.Category.Name,element.Name ))
        return True

    if element.GroupId != InvalidElementId:
        global info  # declare 'info' as global
        design_option_text = ""
        design_option = element.DesignOption
        if design_option:
            designOptionSetId = design_option.get_Parameter(DB.BuiltInParameter.OPTION_SET_ID).AsElementId()
            designOptionSet = doc.GetElement(designOptionSetId)
            DSName = designOptionSet.Name
            design_option_text = "(" + element.DesignOption.Name+" in " + DSName + ") "
        info += ("* {}Element ID: {} - Category: {} - Name: {} is in a group ID: {}\n".format(design_option_text, element.Id.IntegerValue,element.Category.Name,element.Name, element.GroupId.IntegerValue ))
        return True
    else:
        return False

def check_dependencies( element ):
    global info  # declare 'info' as global
    dependent_element_ids = element.GetDependentElements(element_filter)
    dependencies_info = "   Dependent elements list:\n"
    i = 0
    for dep_element_id in dependent_element_ids:
        dep_element = doc.GetElement(dep_element_id)
        # Elements with no category can be omitted
        # Analytical columns and constants are omitted
        # Sketch lines creating an element can be omitted
        # Automatic sketch dimensions are omitted
        # The same element as the element is also always listed as a dependent element and needs to be omitted.
        if dep_element_id != element.Id and dep_element.Category and dep_element.Category.Name != "Constraints" and dep_element.Category.Name != "Analytical Columns" and dep_element.Category.Name != "<Sketch>" and dep_element.Category.Name != "Automatic Sketch Dimensions":
            i += 1
            dep_element = doc.GetElement(dep_element_id)
            dependencies_info += ("    "+str(i)+". ID: {} - Category: {} - Name: {}\n".format(dep_element_id.IntegerValue, dep_element.Category.Name,  dep_element.Name))
    if i != 0:
        design_option_text = ""
        design_option = element.DesignOption
        if design_option:
            design_option_text = "("+element.DesignOption.Name+") "
        info += ("* {}Element ID: {} - Category: {} - Name: {}\n".format(design_option_text, element.Id.IntegerValue, element.Category.Name,  element.Name))
        # element has dependent elements
        info += dependencies_info
        return True
    else:
        return False

warning_number = 1
# after sorting all warning ids it's easier to check which should be deleted
# There are only unique elements
for list in grouped_warnings_ids:
    if warning_number > 1:
        info += "\n"
    if len(list) == 2:
        info += ("Warning with {} duplicated elements ({}):\n".format(str(len(list)), warning_number))
    else:
        info += ("Warnings with {} duplicated elements ({}):\n".format(str(len(list)), warning_number))
    warning_number += 1
    check_to_delete = []
    check_to_delete_in_design_options = []
    retained = []
    retained_in_design_options = []
    for element_id in list:
        element = doc.GetElement(element_id)
        if not check_host_and_groupId( element ) and not check_dependencies( element ):
            # this element can be deleted
            design_option_text = ""
            design_option = element.DesignOption
            if design_option:      
                designOptionSetId = design_option.get_Parameter(DB.BuiltInParameter.OPTION_SET_ID).AsElementId()
                designOptionSet = doc.GetElement(designOptionSetId)
                DSName = designOptionSet.Name
                design_option_text = "(" + element.DesignOption.Name+" in " + DSName + ") "
                check_to_delete_in_design_options.append(element_id)
            else:
                check_to_delete.append(element_id)
            info += ("* {}Element ID: {} - Category: {} - Name: {}\n".format(design_option_text, element.Id.IntegerValue, element.Category.Name,  element.Name))
        else:
            design_option = element.DesignOption
            if design_option:
                retained_in_design_options.append(element_id)
            else:
                retained.append(element_id)
    # if there are any elements which can be safely deleted we need to retain one if there are no in retained or delete all.
    if retained:
        # it doesn't matter if some elements are on design options if there is a copy in the main model and all other can be deleted.
        check_to_delete.extend(check_to_delete_in_design_options)
        elements_ids_to_delete.update(check_to_delete)
    else:
        # if there are some elements on design option which can be deleted then only the element in the main model should stay. We need to check if there is at least one element in the main model.
        if check_to_delete:
            check_to_delete.extend(check_to_delete_in_design_options)
            check_to_delete.pop(0)
            elements_ids_to_delete.update(check_to_delete)
        # if warnings are only between elements in the same design option one should be left
        # if warnings are between elements different design option sets one should be left only if user selected 'Yes'
        else:
            if check_to_delete_in_design_options:
                # Check if elements are in the same design option or in a different design option Sets
                # The same design option will always have the same Id number while name may be the same in different DO sets and can't be used.
                DO_ids = set()
                #this is a list of elements which do not repeat on the same DO but may repeat on different DO sets
                DOSets_elements = []
                # we will add design option ids to this set and check if a next element's DO id is on this list already. If it is this next element can be deleted because it is duplicated. If it doesn't exist it must be on a different design option set (most likely) but it could be on a different DO in the same set and in a different set. In this scenario the set ds_ids will contain at least 3 elements. With 2 elements it will be sure that the warnings occurred in two different DO sets. We can first perform removing elements in the same DO and only if more then 2 (>1) check which elements should be deleted from other DO sets.
                for element_id in check_to_delete_in_design_options:
                    element = doc.GetElement(element_id)
                    DO_id = element.DesignOption.Id.IntegerValue
                    if DO_id in DO_ids:
                        #delete
                        elements_ids_to_delete.add(element_id)
                    else:
                        #add
                        DOSets_elements.append(element_id)
                        DO_ids.add(DO_id)
                if len(DO_ids) > 1:
                    if selected_option == 'Yes':
                        dos_elements = []
                        for element_id in DOSets_elements:
                            element = doc.GetElement(element_id)
                            design_option = element.DesignOption
                            designOptionSetId = design_option.get_Parameter(DB.BuiltInParameter.OPTION_SET_ID).AsElementId()
                            designOptionSet = doc.GetElement(designOptionSetId)
                            DSName = designOptionSet.Name

                                # Add the tuple (DSName, element_id) to the list
                            dos_elements.append((DSName, element_id))

                        # Sort the list of tuples by the first element of each tuple (the Design Option Set name)
                        sorted_dos = sorted(dos_elements)

                        # The elements in the last Design Option Set are the last elements in the sorted list
                        # whose first element (the Design Option Set name) is the same as the first element of the last tuple
                        last_dos_name = sorted_dos[-1][0]
                        # this is not needed
                        #last_dos_elements = [element_id for dos_name, element_id in sorted_dos if dos_name == last_dos_name]

                        # All elements not in the last Design Option Set should be deleted
                        elements_ids__in_DO_to_delete = [element_id for dos_name, element_id in sorted_dos if dos_name != last_dos_name]

                        elements_ids_to_delete.update(elements_ids__in_DO_to_delete)
if elements_ids_to_delete:
    DELETED = []
    with revit.Transaction('Delete identical instances'):
        for id in elements_ids_to_delete:
            try:
                #print("Parameter {} was deleted from the model.".format(pp.Name))
                doc.Delete(id)
                DELETED.append(str(id.IntegerValue))
            except Exception as del_err:
                element = doc.GetElement(id)
                if element:
                    name = element.Name
                else:
                    name = "No element"
                logger = script.get_logger()
                logger.error('Error purging duplicate objects: {} | {} | {}'.format(id, name, del_err))
    if len(DELETED) > 1:
        forms.alert(str(len(DELETED))+" elements: {} were deleted from the model.".format(', '.join(DELETED)))
    elif len(DELETED) == 1:
        forms.alert('Element "{}" was deleted from the model.'.format(DELETED[0]))
    # After the transaction, update the info string to reflect the deleted elements
    for id in DELETED:
        # Find the line in the info string that contains the deleted element's ID
        start_index = info.find("Element ID: " + id)
        if start_index != -1:
            # Find the end of the line
            end_index = info.find(id, start_index) 
            # Replace the line with the new message
            info = info[:start_index] + "** ELEMENT WAS DELETED! ID: " + info[end_index:]
else:
    info += "\nNothing was deleted."
logger = script.get_logger()
logger.warning(info)
#forms.alert(info, "Identical Instances Warnings Summary")