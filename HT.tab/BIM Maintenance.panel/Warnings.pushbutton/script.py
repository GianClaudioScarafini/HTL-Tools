# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-07-23
# Version: 1.0.4
# Name: Warnings
# Description: This script is an alternative way to show Revit warnings. It allows to show warnings of a specific type only on a list which allows to select one of the objects in the model. It also shows the number of warnings of each category.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Import pyRevit modules
from pyrevit import revit, DB, script, forms
from collections import defaultdict
import re

# Get the current document
doc = revit.doc
warnings = doc.GetWarnings()
InvalidElementId = DB.ElementId.InvalidElementId

if warnings:
    warningTypes = []
    all_warnings = defaultdict(int)  # Sort Warnings by Type
    for w in warnings:
        description = w.GetDescriptionText()
        if w.HasResolutions():
            des_res = (description, w.GetDefaultResolutionCaption())
        else:
            des_res = (description, '')
        all_warnings[des_res] += 1

    class warningType(forms.TemplateListItem):
        def __init__(self, name, description, resolution, count):
            self.Name = name
            self.Description = description
            self.Resolution = resolution
            self.Count = count
        

    for warning in all_warnings:
        count = all_warnings[warning]
        list_name  = '(' + str(count) + ') ' + warning[0]
        warningTypes.append(warningType( list_name, warning[0], warning[1], count ))

    warningTypes.sort(key=lambda x: x.Description )

    class AllWarningTypes(forms.TemplateListItem):
        @property
        def name(self):
            return self.item.Name
        def __bool__(self):
            return self.Name is not None

    return_warnings = \
        forms.SelectFromList.show(
            [AllWarningTypes(x) for x in warningTypes],
            title='Select a warning type',
            width=870,
            button_name='Show warnings',
            multiselect=True
        )

    if return_warnings is not None:
        output = script.get_output()
        selected = len(return_warnings)
        current = 0
        total = 0
        for return_warning in return_warnings:
            selected_warnings = [
                warning for warning in warnings 
                if warning.GetDescriptionText() == return_warning.Description
            ]
            total += return_warning.Count
            title = str(total) + " WARNING"
            if return_warning.Count > 1:
                title += 'S'
            output.set_title(title)
            output.print_md( "**Warning type:**"  )
            output.print_md( '> ###*'+ return_warning.Description + '*'  )
            output.print_md('---')
            if return_warning.Resolution:
                output.print_md( "**Warning default resolution:**"  )
                output.print_md( '> ###*'+ return_warning.Resolution + '*'  )
                output.print_md('---')
            count = 1
            parameter_search = ''
            if '"Mark"' in return_warning.Description:
                parameter_search = 'Mark'
            if 'Type Mark' in return_warning.Description:
                parameter_search = 'Type Mark'
            if 'Number' in return_warning.Description:
                parameter_search = 'Number'
            # Initialize a global variable for warnings with Area Boundary lines
            areas = None

            def printWarningInfo(element_id, no, parameter_search):
                element = doc.GetElement(element_id)
                element_id = element.Id

                # Get Design Option Set and Design Option
                design_option_info = ''
                design_option = element.DesignOption
                if design_option:      
                    designOptionSetId = design_option.get_Parameter(DB.BuiltInParameter.OPTION_SET_ID).AsElementId()
                    designOptionSet = doc.GetElement(designOptionSetId)
                    element_DO_SetName = designOptionSet.Name
                    element_design_option = design_option.Name
                    design_option_info = ("{} : {} | "
                                .format(element_DO_SetName, 
                                        element_design_option))
                # Get element Category
                category = element.Category.Name
                if 'Area Boundary' in category:
                    AreaSchemeName =''
                    global areas

                    # Now you can use or modify the global variable
                    if areas is None:
                        areas = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Areas).ToElements()
                    # Find the area that contains the boundary line
                    for area in areas:
                        # Check if the area's boundary lines contain the given boundary line
                        regions = area.GetBoundarySegments(DB.SpatialElementBoundaryOptions())
                        for region in regions:
                            segment_ids = []
                            for segment in region:
                                if segment.ElementId != InvalidElementId:
                                    segment_ids.append(segment.ElementId)
                            if element_id in segment_ids:
                                # If so, print the area's scheme
                                AreaSchemeName = area.AreaScheme.Name
                    # Some Area  lines may not form part of a boundary. THey can be found by searching the Area Plan views elements.
                    if not AreaSchemeName:
                        # Get the level of the boundary line
                        boundary_line_level_id = element.LevelId

                        # Collect all ViewPlan elements
                        area_plan_views = DB.FilteredElementCollector(doc)\
                            .OfClass(DB.ViewPlan)\
                            .ToElements()
                        # Filter the ViewPlan elements to only include non-template Area Plans
                        area_plan_views = [v for v in area_plan_views if not v.IsTemplate and v.ViewType == DB.ViewType.AreaPlan and v.GenLevel.Id == boundary_line_level_id]

                        found_view = None
                        for view in area_plan_views:
                            # Collect all Area Boundary Lines in this view
                            view_boundary_lines_collector = DB.FilteredElementCollector(doc, view.Id)\
                                .WhereElementIsNotElementType()\
                                .OfCategory(DB.BuiltInCategory.OST_AreaSchemeLines)\
                                .ToElementIds()

                            # Check if our specific boundary line exists in this view
                            if element.Id in view_boundary_lines_collector:
                                found_view = view
                                AreaSchemeName = found_view.AreaScheme.Name
                                category = category +' ('+area.AreaScheme.Name + ')'
                                break
                        if not found_view:
                            category = category +' (Scheme not found)'
                    else:
                        # If so, print the area's scheme
                        category = category +' ('+area.AreaScheme.Name + ')'
                if 'Areas' in category:
                    category = category +' ('+element.AreaScheme.Name + ')'

                if category is not None:
                    category =  category.replace('<', '&lt;').replace('>', '&gt;')
                # else:
                #     bic = DB.BuiltInCategory.get_Item(element.Id.IntegerValue)
                #     category =  bic.ToString().replace('<', '&lt;').replace('>', '&gt;')
                
                # Get Level
                level = '_'
                def GetLevel(element):
                    try:
                        return 'Level : ' + element.LookupParameter('Level').AsValueString() + ' | ' 
                    except:
                        return ''
                def GetLevelConstraint(element):
                    try:
                        return 'Level : ' + element.LookupParameter('Base Constraint').AsValueString() + ' | ' 
                    except:
                        return ''
                def GetBaseLevel(element):
                    try:
                        return 'Level : ' + element.LookupParameter('Base Level').AsValueString() + ' | ' 
                    except:
                        return ''
                def GetLevelId(element):
                    try:
                        levelId = element.LevelId
                        element_level = doc.GetElement(levelId)
                        levelName = element_level.Name
                        return 'Level : ' + levelName + ' | ' 
                    except:
                        return ''
                try:
                    if category == 'Walls':
                        try:
                            level = GetLevelConstraint(element)
                            if not level:
                                level = GetLevel(element)
                            if not level:
                                level = GetLevelId(element)
                        except:
                            pass
                    else:
                        level = GetLevel(element)
                        if not level:
                            level = GetLevelId(element)
                        if not level:
                            level = GetBaseLevel(element)
                        if not level:
                            level = ' No Level | '
                except:
                    pass
                # Get the family and type name
                try:
                    if category == 'Rooms' or 'Areas' in category:
                        room_name = element.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsValueString()
                        room_number = element.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsValueString()
                        separator = ''
                        if room_name and room_number:
                            separator = ' '
                        family_and_type = room_name + separator + room_number
                        if not family_and_type:
                            family_and_type = 'No Name and Number'
                    else:
                        family_and_type = \
                            element.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
                    if not family_and_type and element.GetTypeId() == InvalidElementId:
                        family_and_type = \
                            element.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsValueString()
                    if not family_and_type:
                        family_and_type = \
                            element.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
                    if not family_and_type:
                        family_and_type = \
                            element.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                    if not family_and_type:
                        family_and_type = "Family not found!"
                except Exception as del_err:
                    family_and_type = ""
                    # logger = script.get_logger()
                    # logger.error('Error: {} | {} '.format(element.Id, del_err))
                
                # Get the element name
                try:
                    element_name = element.Name
                except:
                    element_name = ''

                if not family_and_type and element_name:
                    family_and_type = element_name
                # Check if the element name is at the end of the family and type string
                elif not family_and_type.endswith(': ' + element_name) and element_name:
                    # If it's not, append it to the end
                    family_and_type += ': ' + element_name

                # Get Workset
                workset = ''
                element_workset = revit.query.get_element_workset(element).Name
                if element_workset:
                    workset = ("**Workset : ** {} | "
                                    .format(element_workset))

                # Get element creator
                element_instance_creator = \
                    DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc, element.Id).Creator.replace("-", "&minus;")
                # Without replacing minus sign Markdown was failing to convert text

                if no == 'single':
                    order = "> "
                else:
                    order = str(no) +". " # **FIRST / SECOND ELEMENT | **"
            
                # Search for "Mark" or "Type Mark" parameter
                parameter_value = ''
                if parameter_search:
                    def escape_markdown_chars(text):
                        # Define the special characters that need to be escaped in markdown
                        special_chars = r"([#*_{}\\[\\()+\\.!`-])"

                        # Use a regular expression to replace each special character with its escaped version
                        escaped_text = re.sub(special_chars, r"\\\1", text)
                        # Replace < and > with their HTML entities
                        escaped_text = escaped_text.replace("<", "&lt;").replace(">", "&gt;")
                        return escaped_text
                
                if parameter_search == 'Mark':
                    mark_value = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK).AsValueString()
                    if mark_value:
                        parameter_value = ' - Mark : ' + escape_markdown_chars(mark_value)

                if parameter_search == 'Type Mark':
                    # Get the Type Mark parameter value
                    try:
                        type_mark_value = element.LookupParameter('Type Mark').AsValueString()
                    except:
                        type_mark_value = ''
                    
                    if type_mark_value:
                        parameter_value = ' - Type Mark : ' + escape_markdown_chars(type_mark_value)
                
                if parameter_search == 'Number':
                    number_value = element.LookupParameter('Number').AsValueString()
                    if number_value:
                        parameter_value = ' - Number : ' + escape_markdown_chars(number_value)

                text = order + workset + design_option_info + level + category + " : " + family_and_type + \
                        parameter_value + \
                        "<br>**Id:** " + output.linkify(element_id) + ' Created by: ' + element_instance_creator
                
                # Check if the element is a group.
                if element.GroupId != DB.ElementId.InvalidElementId:
                    group_id = element.GroupId
                    group = doc.GetElement(group_id)
                    group_name = group.Name
                    text += ' | **Group Id:** ' + output.linkify(group_id) + ' Name: ' + group_name
                #forms.alert( text )
                return  text

            for warning in selected_warnings:
                element_ids = warning.GetFailingElements()
                
                output.print_md('### Warning '+str(count))
                if len(element_ids) == 1:
                    output.print_md(printWarningInfo(element_ids[0], 'single', parameter_search ))
                else:
                    markdown_text = ''
                    for i in range(len(element_ids)):
                        markdown_text += printWarningInfo( element_ids[i], i+1, parameter_search )+'\n\n'
                    output.print_md(markdown_text)
                count += 1
            current +=1
            if current < selected:
                output.print_md('---')
                output.print_md('---')
else:
    forms.alert('There are no warnings in the model.', title='No warnings')