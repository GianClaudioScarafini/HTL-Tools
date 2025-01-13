# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-06-13
# Version: 1.0.0
# Description: Removes identical instances from the model. The tool checks if there are any dependent elements connected to one of duplicated element and only deletes one which doesn't have any. It notifies about others.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

import clr
import System
from System.Collections.Generic import List
# Import pyRevit modules
from pyrevit import revit, DB, script, forms, framework

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

# Print the warnings
elements_ids_to_delete = {}
# Checked elements - some elements may exist in more then two warnings if elements are duplicated more then 2 times
# the retained list contains elements which are in groups or have dependent elements and must NOT be deleted.
checked_element_ids = {}
retained_element_ids = {}

grouped_warnings_ids = []

# this creates lists of element id  which are duplicated
for warning in identical_instance_warnings:
    element_ids = warning.GetFailingElements()
    # message = warning.GetDescriptionText() - not necessary
    # there are always two elements
    # if there are more then 2 then one warning will be created for each possible pair of elements in the same place
    # checks if there is a sense of doing it at all to speed the process.
    if element_ids[0] in checked_element_ids or element_ids[0] in checked_element_ids:
        for list in grouped_warnings_ids:
            for id in list:
                if element_ids[0] == id:
                    # the other id represents another duplicated element
                    grouped_warnings_ids.append(element_ids[1])
                if element_ids[1] == id:
                    grouped_warnings_ids.append(element_ids[0])
    else:
        checked_element_ids.add(element_ids[0])
        checked_element_ids.add(element_ids[1])
        warning_ids = [element_ids[0],element_ids[1]]
        grouped_warnings_ids.append(warning_ids)

def check_groupId( element ):
    if element.GroupId != InvalidElementId:
        retained_element_ids.Add(element.Id)
        info += (" Element (ID: {}) is in a group (ID: {})\n\n".format(element.Id.IntegerValue, element.GroupId.IntegerValue, ))
        return True
    else:
        return False

def check_dependencies( element ):
    dependent_element_ids = element.GetDependentElements(element_filter)
    dependencies_info = "Dependent elements list:\n"
    i = 0
    for dep_element_id in dependent_element_ids:
        dep_element = doc.GetElement(dep_element_id)
        if dep_element_id != element.Id and dep_element.Category and dep_element.Category.Name != "Constraints":
            i += 1
            dep_element = doc.GetElement(dep_element_id)
            dependencies_info += (str(i)+". ID number: {} - Category: {} - Name: {}\n\n".format(dep_element_id.IntegerValue, dep_element.Category.Name,  dep_element.Name))
    if i != 0:
        # element has dependent elements
        info += dependencies_info
        return True
    else:
        return False

warning_number = 1
# after sorting all warning ids it's easier to check which should be deleted
# There are only unique elements
for list in grouped_warnings_ids:
    info += ("Warning {}:\n".format(warning_number))
    check_to_delete = []
    retained = []
    for element_id in list:
        element = doc.GetElement(element_id)
        if not check_groupId( element ) and check_dependencies( element ):
            # this element can be deleted
            check_to_delete.append(element)
        else:
            retained.append(element)
    # if there are any elements which can be safely deleted we need to retain one if there are no in retained or delete all.
    if retained:
        elements_ids_to_delete.update(check_to_delete)
    else:
        retained.append(check_to_delete.pop(0))
        elements_ids_to_delete.update(check_to_delete)

    InvalidElementId = DB.ElementId.InvalidElementId
    if element_ids[0] != InvalidElementId and element_ids[1] != InvalidElementId:
        info += ("Both elements {} and {} are in groups {} and {} - open them, check manually and delete one of them.\n\n".format(element_ids[0].IntegerValue, element_ids[1].IntegerValue, element0.GroupId.IntegerValue, element1.GroupId.IntegerValue ))
    elif element_ids[0] == InvalidElementId and element_ids[1] != InvalidElementId:
        # if one of the elements is not in a group we need to check if it was in other warnings before and is saved for deleting or retaining and have any dependent elements if not for deleting.
    elif element_ids[1] == InvalidElementId and element_ids[0] != InvalidElementId:
        # if one of the elements is not in a group we need to check if it was in other warnings before and is saved for deleting or retaining and have any dependent elements if not for deleting.
    else:
        # if both elements are not in groups we need to check if they appeared in previous warnings. 
        # It is possible that:
        # 1. both elements are retained - check which one must stay
        # 2. both are to be deleted - do nothing
        # 3. One is retained and one is to be deleted - do nothing
        # 4. only one is on one of these lists and the other is not on any
        #  A. One is to be deleted - do nothing - it will appear in other warning
        #  B. One is to be retained - delete one of them
        # 5. both are not
        # We need to check if they were retained or deleted.
        # if one was retained it should be still retained and the other element can be deleted if it doesn't have any dependencies.

        if check_retained( element_ids[0] ) and not check_retained( element_ids[1] ):
            if not check_groupId( element1 ) and not check_dependencies( element1 ):
                # Element can be deleted
                elements_ids_to_delete.add( element_ids[1] )
        if check_retained( element_ids[1] ) and not check_retained( element_ids[0] ):
            if not check_groupId( element0 ) and not check_dependencies( element0 ):
                # Element can be deleted
                elements_ids_to_delete.add( element_ids[0] )

    warning_number += 1


    firstIsInGroup = False
    firstElementHasDependencies = False
    secondElementHasDependencies = False
    firstElementId = None
    firstGroupId = None
    for i in range(len(element_ids)):
        element_id = element_ids[i]
        element = doc.GetElement(element_id)
        # there are always two elements
        # if there are more then 2 then one warning will be created for each possible pair of elements in the same place
        if element.GroupId == DB.ElementId.InvalidElementId:
            info += "  Element is not in a group\n"
            info += ("  Element ID: {} - Category: {} - Name: {}\n\n".format(element_id.IntegerValue, element.Category.Name, element.Name))
            dependent_element_ids = element.GetDependentElements(element_filter)
            for dep_element_id in dependent_element_ids:
                dep_element = doc.GetElement(dep_element_id)
                if dep_element_id != element_id and dep_element.Category and dep_element.Category.Name != "Constraints":
                    dep_element = doc.GetElement(dep_element_id)
                    if i == 0:
                        firstElementHasDependencies = True
                    else:
                        secondElementHasDependencies = True
                    info += ("  Dependent Element ID: {} - Category: {} - Name: {}\n\n".format(dep_element_id.IntegerValue, dep_element.Category.Name,  dep_element.Name))
        else:
            if i == 0:
                firstIsInGroup = True
                firstElementId = element_id.IntegerValue
                firstGroupId = element.GroupId.IntegerValue
                info += "  Element is in a group\n"
            else:
                if firstIsInGroup == True:
                    info += ("Both elements {} and {} are in groups {} and {} - open them and check manually.\n\n".format(firstElementId, element_id.IntegerValue, firstGroupId, element.GroupId.IntegerValue ))
            info += ("  Element ID: {} - Category: {} - Name: {}\n\n".format(element_id.IntegerValue, element.Category.Name, element.Name))
    if firstElementHasDependencies == False and secondElementHasDependencies == False:
        # Delete element with a bigger ID number:

        info += ("Element ID: {} Element is not in a group, does not have dependent elements and can be deleted\n".format(element_id.IntegerValue))    

forms.alert(info)
