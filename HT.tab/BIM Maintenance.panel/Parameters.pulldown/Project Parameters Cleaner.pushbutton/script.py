# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-05-14
# Version: 1.0.0
# Description: This script will delete all not used Project Parameters from the project.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import StorageType, SpecTypeId

# Import pyRevit modules
from pyrevit import revit, DB, script, forms

app = __revit__.Application
ver = int(app.VersionNumber)
if ver <=2022:
    from Autodesk.Revit.DB import ParameterType
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
            self.inUse = False

    class ViewFilterToPurge(forms.TemplateListItem):
        @property
        def name(self):
            return self.item.Name

    def checkIfInUse(elements, pp):
        # If there are no elements a parameter can be deleted.
        # None will be returned in this case and this is fine.
        if elements:
            for element in elements:
                parameters = element.GetParameters(pp.Name)
                for par in parameters:
                    # Checks if correct parameter was acquired
                    if pp.pp_id.ToString() == par.Id.ToString():
                        if par.HasValue:
                            value = None
                            try:
                                if par.StorageType == StorageType.String:
                                    value = par.AsString()
                                elif par.StorageType == StorageType.Integer:
                                    if ver >= 2023: # ParameterType() got obsolete in Revit 2023 and above.
                                        if par.Definition.GetDataType().Equals(SpecTypeId.Boolean.YesNo):
                                            if par.HasValue:
                                                return True
                                        else:
                                            value = par.AsInteger()
                                    else:
                                        if ParameterType.YesNo == par.Definition.ParameterType:
                                            if par.HasValue:
                                                return True
                                        else:
                                            value = par.AsInteger()
                                elif par.StorageType == StorageType.Double:
                                    value = par.AsDouble()
                                elif par.StorageType == StorageType.ElementId:
                                    value = par.AsElementId()
                            except Exception as del_err:
                                logger.error('Error checking parameter value: {} | {}'
                                        .format(pp.Name, del_err))
                                value = 'For safety it is better to not delete a parameter that created an error and assume it has a value and has been used.'
                            # If parameter has values of empty sting = "" it should be deleted. 
                            # par.HasValue for empty string would return True - has a value. We do not want this except YesNo parameters.
                            if value or value == 0:
                                return True

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
                title='Select project parameters to check if they are in use',
                width=500,
                button_name='Check these parameters',
                multiselect=True
            )

        if return_options:
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
                if not pp.inUse: # If not True - it is the same as "if sp.inUse = False or sp.inUse = None".
                    parametersToDelete.append(pp)

            # Ask user to select parameters to delete.
            delete_options = \
                forms.SelectFromList.show(
                    [ViewFilterToPurge(x) for x in parametersToDelete],
                    title='Select not used project parameters to delete',
                    width=500,
                    button_name='Delete parameters!',
                    multiselect=True
                )
            if delete_options:
                DELETED = []
                with revit.Transaction('Purge Unused Project Parameters'):
                    for pp in delete_options:
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