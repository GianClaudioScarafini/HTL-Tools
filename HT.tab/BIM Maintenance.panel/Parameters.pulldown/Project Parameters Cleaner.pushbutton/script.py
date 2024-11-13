# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-05-14
# Version: 1.1.0
# Description: This script allow to inspect Project Parameters usage in the project by checking how many elements use each selected parameter and allows to delete any parameter from the project.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Import pyRevit modules
from pyrevit import revit, DB, script, forms

app = revit.doc.Application
ver = int(app.VersionNumber)
if ver <= 2022:
    parameter_type = DB.ParameterType
doc = revit.doc

if doc.IsFamilyDocument:
    forms.alert('This is a family document. Please open a project document.')
else:
    # PP - Project Parameter
    pp_list = []

    def TypeElementsByCategory(catId):
        collector = DB.FilteredElementCollector(doc).OfCategoryId(catId).WhereElementIsElementType()
        return collector.ToElements()

    def InstanceElementByCategory(catId):
        collector = DB.FilteredElementCollector(doc).OfCategoryId(catId).WhereElementIsNotElementType()
        return collector.ToElements()

    class PP(forms.TemplateListItem):
        def __init__(self, name, category_set, pp_id, is_inst):
            self.Name = name
            self.category_set = category_set
            self.pp_id = pp_id
            self.is_inst_value = is_inst
            self.inUse = 1 

    class ViewFilterToPurge(forms.TemplateListItem):
        @property
        def name(self):
            return self.item.Name

    def checkIfInUse(elements, pp):
        # If there are no elements a parameter can be deleted.
        # None will be returned in this case and this is fine.
        count = 0
        if elements:
            for element in elements:
                parameters = element.GetParameters(pp.Name)
                for par in parameters:
                    # Checks if correct parameter was acquired
                    if par.Id == pp.pp_id:
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
                                                break
                                        else:
                                            value = par.AsInteger()
                                    else:
                                        if DB.ParameterType.YesNo == par.Definition.ParameterType:
                                            if par.HasValue:
                                                count += 1
                                                break
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
                                    break
                            except Exception as del_err:
                                logger.error('Error checking parameter value: {} | {}'
                                        .format(pp.Name, del_err))
                                count += 1 # For safety it is better to not delete a parameter that created an error and assume it has a value and has been used.
                                break
                            
        return count
    
    logger = script.get_logger()

    parametersToDelete = []

    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        pp = doc.GetElement(iterator.Key.Id)
        if pp.GetType().ToString() == 'Autodesk.Revit.DB.ParameterElement':
            binding_map = doc.ParameterBindings
            binding = binding_map.Item[pp.GetDefinition()]
            category_set = []
            if binding != None:
                category_set = binding.Categories
            is_inst_value = iterator.Current.GetType(
            ).ToString() == 'Autodesk.Revit.DB.InstanceBinding'
            # Creates an object to store the information of each parameter
            pp_obj = PP(iterator.Key.Name, category_set, pp.Id, is_inst_value)
            pp_list.append(pp_obj)
            # Sorts a list of parameters alphabetically by name.
            pp_list.sort(key=lambda pp_obj: pp_obj.Name)

    if not pp_list:
        forms.alert('No Project Parameters in the model.')
    else:
        # Ask user to select parameters to checks
        return_options = \
            forms.SelectFromList.show(
                [ViewFilterToPurge(x) for x in pp_list],
                title='Select Project Parameters to check if they are in use.',
                width=500,
                button_name='Select these parameters',
                multiselect=True
            )

        if return_options:
            parameters_with_counts = []
            AllTypeElements = {}
            AllInstanceElements = {}
            for pp in return_options:
                allElementsOfAllCategories = []
                for cat in pp.category_set:
                    if pp.is_inst_value:
                        if cat.Name not in AllInstanceElements:
                            AllInstanceElements[cat.Name] = InstanceElementByCategory(cat.Id)
                            allElementsOfAllCategories.extend(AllInstanceElements[cat.Name])
                        else:
                            allElementsOfAllCategories.extend(AllInstanceElements[cat.Name])
                    else:
                        if cat.Name not in AllTypeElements:
                            AllTypeElements[cat.Name] = TypeElementsByCategory(cat.Id)
                            allElementsOfAllCategories.extend(AllTypeElements[cat.Name])
                        else:
                            allElementsOfAllCategories.extend(AllTypeElements[cat.Name])
                pp.inUse = checkIfInUse(allElementsOfAllCategories, pp)
                parameters_with_counts.append((pp, pp.inUse))
            
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
                        for pp_tuple in delete_options:
                            pp = pp_tuple[0]
                            try:
                                #print("Parameter {} was deleted from the model.".format(pp.Name))
                                doc.Delete(pp.pp_id)
                                DELETED.append(pp.Name)
                            except Exception as del_err:
                                logger.error('Error purging parameter: {} | {}'
                                            .format(pp.Name, del_err))
                    if len(DELETED) > 1:
                        forms.alert("Parameters: {} were deleted from the model.".format(', '.join(DELETED)))
                    else:
                        forms.alert('Parameter "{}" was deleted from the model.'.format(DELETED[0]))
                else:
                    forms.alert('Nothing selected.')
        else:
            forms.alert('Nothing selected.')