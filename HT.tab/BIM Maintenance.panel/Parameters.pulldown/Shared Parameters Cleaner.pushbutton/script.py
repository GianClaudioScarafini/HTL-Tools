# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-05-14
# Version: 1.1.0
# Description: This script allow to inspect Shared Parameters usage in the project by checking how many elements use each selected parameter and allows to delete any parameter from the project.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Import pyRevit modules
from pyrevit import revit, DB, script, forms

app = revit.doc.Application
ver = int(app.VersionNumber)
doc = revit.doc

if doc.IsFamilyDocument:
    forms.alert('This is a family document. Please open a project document.')
else:
    # SP - Shared Parameter
    sp_list = []

    def TypeElementsByCategory(catId):
        collector = DB.FilteredElementCollector(doc).OfCategoryId(catId).WhereElementIsElementType()
        return collector.ToElements()

    def InstanceElementByCategory(catId):
        collector = DB.FilteredElementCollector(doc).OfCategoryId(catId).WhereElementIsNotElementType()
        return collector.ToElements()

    class SP(forms.TemplateListItem):
        def __init__(self, name, sp_id, category_set, guid, is_inst):
            self.Name = name
            self.sp_id = sp_id
            self.category_set = category_set
            self.guid = guid
            self.is_inst_value = is_inst
            self.inUse = False

    class ViewFilterToPurge(forms.TemplateListItem):
        @property
        def name(self):
            return self.item.Name

    def checkIfInUse(elements, sp):
        # If there are no elements a parameter can be deleted.
        # None will be returned in this case and this is fine.
        count = 0
        if elements:
            for element in elements:
                par = element.get_Parameter(sp.guid)
                try:
                    if par.HasValue:
                        value = None
                        try:
                            if par.StorageType == DB.StorageType.String:
                                value = par.AsString()
                            elif par.StorageType == DB.StorageType.Integer:
                                if ver >= 2023: # ParameterType() got obsolete in Revit 2023 and above.
                                    if par.Definition.GetDataType().Equals(DB.SpecTypeId.Boolean.YesNo):
                                        if par.HasValue:
                                            count += 1
                                    else:
                                        value = par.AsInteger()
                                else:
                                    if DB.ParameterType.YesNo == par.Definition.ParameterType:
                                        if par.HasValue:
                                            count += 1
                                    else:
                                        value = par.AsInteger()
                            elif par.StorageType == DB.StorageType.Double:
                                value = par.AsDouble()
                            elif par.StorageType == DB.StorageType.ElementId:
                                value = par.AsElementId()
                            # If parameter has values of empty sting = "" it should be deleted. 
                            # par.HasValue for empty string would return True - has a value. We do not want this except YesNo parameters.
                            if value or value == 0:
                                count += 1
                        except Exception as del_err:
                            logger.error('Error checking parameter value: {} | {}'
                                    .format(sp.Name, del_err))
                            count += 1 # For safety it is better to not delete a parameter that created an error and assume it has a value and has been used.
                except Exception as del_err:
                    logger.error('Error checking parameter HasValue: {} | {}'
                                    .format(sp.Name, del_err))
                    forms.alert('Error checking parameter HasValue: {} | {} | {}'
                                    .format(sp.Name, del_err, str(element.Id) ))
                    count += 1 # For safety it is better to not delete a parameter that created an error and assume it has a value and has been used.
                    continue
        return count            

    logger = script.get_logger()

    parametersToDelete = []

    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        sp = doc.GetElement(iterator.Key.Id)
        if sp.GetType().ToString() == 'Autodesk.Revit.DB.SharedParameterElement':
            binding_map = doc.ParameterBindings
            binding = binding_map.Item[sp.GetDefinition()]
            category_set = []
            if binding != None:
                category_set = binding.Categories
            is_inst_value = iterator.Current.GetType(
            ).ToString() == 'Autodesk.Revit.DB.InstanceBinding'
            # Creates an object to store the information of each parameter
            sp_obj = SP(iterator.Key.Name, sp.Id,
                        category_set, sp.GuidValue, is_inst_value)
            sp_list.append(sp_obj)
            # Sorts a list of parameters alphabetically by name.
            sp_list.sort(key=lambda sp_obj: sp_obj.Name)

    if not sp_list:
        forms.alert('No Project Parameters in the model.')
    else:
        # Ask user to select parameters to checks
        return_options = \
            forms.SelectFromList.show(
                [ViewFilterToPurge(x) for x in sp_list],
                title='Select Shared Parameters to check if they are in use.',
                width=500,
                button_name='Select these parameters',
                multiselect=True
            )

        if return_options:
            parameters_with_counts = []
            AllTypeElements = {}
            AllInstanceElements = {}
            for sp in return_options:
                #if sp.category_set:
                allElementsOfAllCategories = []
                for cat in sp.category_set:
                    if not sp.is_inst_value:
                        if cat.Name not in AllTypeElements:
                            AllTypeElements[cat.Name] = TypeElementsByCategory(cat.Id)
                            allElementsOfAllCategories.extend(AllTypeElements[cat.Name])
                        else:
                            allElementsOfAllCategories.extend(AllTypeElements[cat.Name])
                    else:
                        if cat.Name not in AllInstanceElements:
                            AllInstanceElements[cat.Name] = InstanceElementByCategory(cat.Id)
                            allElementsOfAllCategories.extend(AllInstanceElements[cat.Name])
                        else:
                            allElementsOfAllCategories.extend(AllInstanceElements[cat.Name])
                sp.inUse = checkIfInUse(allElementsOfAllCategories, sp)
                parameters_with_counts.append((sp, sp.inUse))
            
            # Sort the list based on count
            parameters_with_counts.sort(key=lambda x: x[1])
            
            # Create a custom class to create a list item with the parameter name and the count
            class ParameterWithCount(forms.TemplateListItem):
                @property
                def name(self):
                    return "{} - Used in {} elements".format(self.item[0].Name, self.item[1])

            if parameters_with_counts:
                # Ask user to select parameters to delete.
                delete_options = \
                    forms.SelectFromList.show(
                        [ParameterWithCount(x) for x in parameters_with_counts],
                        title='Parameters sorted by usage. Select which to delete.',
                        width=500,
                        button_name='Delete parameters!',
                        multiselect=True
                    )
                if delete_options:
                    DELETED = []
                    with revit.Transaction('Purge Unused Project Parameters'):
                        for sp_tuple in delete_options:
                            sp = sp_tuple[0]
                            try:
                                #print("Parameter {} was deleted from the model.".format(sp.Name))
                                doc.Delete(sp.sp_id)
                                DELETED.append(sp.Name)
                            except Exception as del_err:
                                logger.error('Error purging parameter: {} | {}'
                                            .format(sp.Name, del_err))
                    if len(DELETED) > 1:
                        forms.alert("Parameters: {} were deleted from the model.".format(', '.join(DELETED)))
                    else:
                        forms.alert('Parameter "{}" was deleted from the model.'.format(DELETED[0]))
                else:
                    forms.alert('Nothing selected.')
        else:
            forms.alert('Nothing selected.')