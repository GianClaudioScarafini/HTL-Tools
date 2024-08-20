# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-07-31
# Version: 1.0.0
# Description: PyRevit python script which will collect all elements in the models and from these elements will filter out all elements which do not have a parameter COBie check. For all other elements which have COBie parameter checked or set to True it will acquire values of instance parameter Width, Height and Length and if instance parameter will not return any value acquire these parameter values from the element Type. If also Type parameters will not return any value it will get a Bounding Box for each element and calculate the element Width, Height and Length from Bonding Box min and max value. At the and acquired values for Width, Height and Length will be saved to element Types corresponding parameters called: COBie.Type.NominalWidth, COBie.Type.NominalLength, and COBie.Type.NominalHeight
#
# IMPORTANT!
# COBie.Type.NominalWidth Length and Height needs to be created manually as Text parameters for all element Categories which are selected for COBie export.
# Tested with: Revit 2022+
# Requirements: pyRevit add-in

import clr
import sys
from collections import defaultdict

clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitNodes')

from pyrevit import revit, UI
from Autodesk.Revit.DB import *

# Get the current Revit document
doc = revit.doc

none_type = []

def newAlert(msg, title='pyRevit',
          cancel=False, yes=False, no=False):
    buttons = UI.TaskDialogCommonButtons.None   # noqa

    if cancel:
        buttons |= UI.TaskDialogCommonButtons.Cancel
    if yes:
        buttons |= UI.TaskDialogCommonButtons.Yes
    if no:
        buttons |= UI.TaskDialogCommonButtons.No

    res = UI.TaskDialog.Show(title, msg, buttons)

    if res == UI.TaskDialogResult.Yes:
        return True
    elif res == UI.TaskDialogResult.No:
        return False
    else:
        sys.exit()

# Input from a user: Should elements in Groups be included?

not_grouped = newAlert(
    'Should elements in Groups be excluded?',
    title="Exclude elements in Groups?",
    cancel=True,
    yes=True,
    no=True
)

# Conversion utility function from Revit internal units (feet) to millimetres
def convert_to_mm(value_in_feet):
    value_in_mm = UnitUtils.ConvertFromInternalUnits(value_in_feet, UnitTypeId.Millimeters)
    rounded_value = round(value_in_mm)  # Round to nearest integer
    return rounded_value

# Function to get parameter value by name


def get_parameter_value(element, param_name):
    param = element.LookupParameter(param_name)
    if param and param.HasValue:
        return convert_to_mm(param.AsDouble())
    return None

# Function to calculate bounding box dimensions in mm


def get_bounding_box_dimensions(element):
    bbox = element.get_BoundingBox(None)
    if bbox:
        min_pt = bbox.Min
        max_pt = bbox.Max
        width = convert_to_mm(max_pt.X - min_pt.X)
        height = convert_to_mm(max_pt.Y - min_pt.Y)
        length = convert_to_mm(max_pt.Z - min_pt.Z)
        return width, height, length
    return None, None, None

# Function to check if an element is in a group
def is_in_group(element):
    group_id = element.GroupId
    return group_id != ElementId.InvalidElementId

# Collect all elements in the model
collector = FilteredElementCollector(
    doc).WhereElementIsNotElementType().ToElements()

# Filter elements with COBie parameter checked
elements_with_cobie = [
    elem for elem in collector
    if elem.LookupParameter('COBie') and elem.LookupParameter('COBie').AsInteger() == 1
]

# Filter out elements based on group membership if 'not_grouped' is True
if not_grouped:
    elements_with_cobie = [
        elem for elem in elements_with_cobie if not is_in_group(elem)]

# Dictionary to store unique values for type parameters
type_params_values = defaultdict(
    lambda: {'width': set(), 'height': set(), 'length': set()})

# Set to track types to be excluded because at least one instance is in a group
types_to_exclude = set()

for elem in elements_with_cobie:
    # Do noting for Rooms
    if elem.Category.Name == "Rooms":# or elem.Category.Name != "Walls":
        continue

    # Check if the element is in a group, and mark its type for exclusion if so
    if not_grouped and is_in_group(elem):
        types_to_exclude.add(elem.GetTypeId())
        continue

    # Get instance parameter values
    width = get_parameter_value(elem, 'Width')
    height = get_parameter_value(elem, 'Height')
    length = get_parameter_value(elem, 'Length')

    # If no Width or Length use Depth
    if not width or not length:
        depth = get_parameter_value(elem, 'Depth')
        if depth:
            if not width:
                width = depth
            else:
                length = depth

    # If instance parameters are not available, get type parameter values
    if not width or not height or not length:
        if elem.GetTypeId() != ElementId.InvalidElementId:
            elem_type = doc.GetElement(elem.GetTypeId())
            width = width or get_parameter_value(elem_type, 'Width')
            height = height or get_parameter_value(elem_type, 'Height')
            length = length or get_parameter_value(elem_type, 'Length')
            
            # If no Width or Length use Depth
            if not width or not length:
                depth = get_parameter_value(elem, 'Depth')
                if depth:
                    if not width:
                        width = depth
                    else:
                        length = depth
        else:
            try:
                name = elem.Name
            except Exception as e:
                name = "No Name"
            print('Element with no Type: {} | {} | {}'.format(elem.Id, name, e))
            none_type.append( (elem.Id, elem.Category.Name, name) )

    # print('Element: {} | {} | {}'.format(width, height, length))

    # If type parameters are also not available, calculate from bounding box
    if not width or not height or not length:
        bbox_width, bbox_height, bbox_length = get_bounding_box_dimensions(elem)
        # Only supplement missing values
        width = width or bbox_width
        height = height or bbox_height
        length = length or bbox_length

    # Store values for type parameters
    type_id = elem.GetTypeId()
    if width:
        type_params_values[type_id]['width'].add(int(width))
    if height:
        type_params_values[type_id]['height'].add(int(height))
    if length:
        type_params_values[type_id]['length'].add(int(length))

# Remove types marked for exclusion
for type_id in types_to_exclude:
    if type_id in type_params_values:
        del type_params_values[type_id]

# Start a transaction to set values to type parameters
with revit.Transaction('Add COBie.NominalSize values from elements'):

    for type_id, dimensions in type_params_values.items():
        elem_type = doc.GetElement(type_id)
        if elem_type:
            # Convert sets to sorted comma-separated strings
            width_str = ', '.join(map(str, sorted(dimensions['width'])))
            height_str = ', '.join(map(str, sorted(dimensions['height'])))
            length_str = ', '.join(map(str, sorted(dimensions['length'])))

            # Uncheck comment to make a list of values.
            print('Element ID: {} | {} | {} | {}'.format(type_id, width_str, height_str, length_str))
            # Set type parameters
            width_param = elem_type.LookupParameter('COBie.Type.NominalWidth')
            height_param = elem_type.LookupParameter('COBie.Type.NominalHeight')
            length_param = elem_type.LookupParameter('COBie.Type.NominalLength')
            
            def PrintErr(elem, parameter_name):
                try:
                    name = elem.Name
                except Exception as e:
                    name = "No Name"
                #print('Element Type parameter "{}" not found: {} | {} | Add this parameter to this Category: {}'.format(parameter_name, elem.Id, name, elem.Category.Name))

            if width_param:
                width_param.Set(width_str)
                # print('Parameter: {} '.format(width_param.Name))
            else:
                PrintErr(elem_type, 'COBie.Type.NominalWidth')
            if height_param:
                height_param.Set(height_str)
            else:
                PrintErr(elem_type, 'COBie.Type.NominalHeight')
            if length_param:
                length_param.Set(length_str)
            else:
                PrintErr(elem_type, 'COBie.Type.NominalLength')
