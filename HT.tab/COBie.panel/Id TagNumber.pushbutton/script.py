# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-08-07
# Version: 1.0.0
# Description: Copies element Id to COBie.Component.TagNumber
#
# Tested with: Revit 2022+
# Requirements: pyRevit add-in

import clr
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')

from pyrevit import revit, DB

# Get the current document
doc = revit.doc

# Function to copy ElementId to COBie.Component.TagNumber
def copy_element_id_to_tag_number(element):
    # Get the ElementId as a string
    element_id_str = str(element.Id.IntegerValue)
    
    # Get the parameter "COBie.Component.TagNumber"
    tag_number_param = element.LookupParameter("COBie.Component.TagNumber")
    
    if tag_number_param and tag_number_param.IsReadOnly == False:
        # Set the parameter value to the ElementId
        tag_number_param.Set(element_id_str)
        return True
    return False

# Start a transaction
with revit.Transaction('Copy Element Id to COBie.Component.TagNumber'):

    # Collect all elements in the document
    elements = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    elements_with_cobie = [
    elem for elem in elements
    if elem.LookupParameter('COBie') and elem.LookupParameter('COBie').AsInteger() == 1
]
    if elements_with_cobie:
        count = 0
        print("Copied element Id:\n")
        # Process each element
        for element in elements_with_cobie:
            # Filter out elements without "COBie" set to Yes
            if copy_element_id_to_tag_number(element):
                print("{}\n".format(element.Id.IntegerValue))
                count +=1

        # Notify completion
        print("Operation completed for {} elements.".format(count))
    else:
        print("No elements with COBIe parameter set to Yes.")