# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-12-05
# Version: 1.0.0
# Description: Delete Views safely from the Project Browser by checking if they do not have any dependent view or they are not placed on any sheet.
# Script shows at the end a list of views which are placed on sheets and these which have dependent views showing what views they are. 
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Import pyRevit modules
from pyrevit import revit, DB, script, forms

# app = revit.doc.Application
# ver = int(app.VersionNumber)
# if ver <= 2022:
#     parameter_type = DB.ParameterType
# doc = revit.doc

# Initialize logger for user feedback
logger = script.get_logger()

# Helper function: Check if a view is on a sheet
def get_view_sheet(view):
    sheet_number_param = view.LookupParameter("Sheet Number")
    if sheet_number_param and sheet_number_param.AsString() != "---":
        sheet_number = sheet_number_param.AsString()
        # Find the sheet using the sheet number
        collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSheet)
        for sheet in collector:
            if sheet.SheetNumber == sheet_number:
                return sheet
    return None 

# Helper function: Check if a view has dependent views
def get_all_dependent_views(view):
    # Collect dependent views from both sources
    dependent_views = view.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Views))
    # Method GetDependentViewIds() returns only views create and visible in the browser tree as dependent views but not dependent detail callouts.
    # Method GetDependentElements() returns all dependent views including detail callouts.
    dependent_ids = [v for v in dependent_views if v != view.Id]
    # Remove duplicates and invalid elements
    return list(set(dep_id for dep_id in dependent_ids if revit.doc.GetElement(dep_id)))

# Main execution
def main():
    # Collect the selected elements
    selected_elements = revit.get_selection()
    if not selected_elements:
        forms.alert("No views selected. Please select views to proceed.", exitscript=True)
    
    # Filter selected elements to ensure only views are considered
    selected_views = []
    for elem in selected_elements:
        if isinstance(elem, DB.View) and not isinstance(elem, (DB.ViewSchedule, DB.ViewSheet)):
            # Check if the view's type matches the specified criteria for deletion
            if elem.ViewType in [
                DB.ViewType.ThreeD,            # 3D view
                DB.ViewType.DraftingView,      # Drafting view
                DB.ViewType.FloorPlan,         # Floor plan
                DB.ViewType.AreaPlan,          # Area plan
                DB.ViewType.CeilingPlan,       # Ceiling plan
                DB.ViewType.Section,           # Section view
                DB.ViewType.Detail,            # Detail view
                DB.ViewType.Elevation,         # Elevation view
                DB.ViewType.Walkthrough        # Walkthrough view
            ]:
                selected_views.append(elem)
        else:
            logger.error("Non-view element selected: {}. Ignored.".format(elem.Name))

    # Exit if no valid views were selected
    if not selected_views:
        forms.alert("No valid views selected. Ensure views are selected and retry.", exitscript=True)
    
    views_to_delete = []
    views_on_sheets_or_with_dependents = []

    # Process each selected view
    for view in selected_views:
        sheet = get_view_sheet(view)
        all_dependents = get_all_dependent_views(view)
        if sheet or all_dependents:
            dependents_with_info = [{
                    "name": revit.doc.GetElement(dep_id).Name,
                    "sheet": get_view_sheet(revit.doc.GetElement(dep_id))
                }
                for dep_id in all_dependents
            ]
            views_on_sheets_or_with_dependents.append({
                "view_name": view.Name,
                "sheet": sheet,
                "dependents": dependents_with_info,
                "view": view,
                "all_dependents": all_dependents
            })
        else:
            views_to_delete.append(view)

    # Check for parent views whose dependents are all marked for deletion
    additional_deletions = set()
    new_views_on_sheets_or_with_dependents = []
    for view in views_on_sheets_or_with_dependents:
        parent_view = view["view"]
        dependent_views = view["all_dependents"]
        if dependent_views and all(dep_id in views_to_delete for dep_id in dependent_views):
            additional_deletions.add(parent_view)
        else:
            new_views_on_sheets_or_with_dependents.append(view)
    views_to_delete.extend(additional_deletions)

    # Collect templates from views marked for deletion
    templates_to_check = set()
    for view in views_to_delete:
        if view.ViewTemplateId != DB.ElementId.InvalidElementId:
            template = revit.doc.GetElement(view.ViewTemplateId)
            templates_to_check.add(template)

    # Check if templates are used by other views
    templates_unused = []
    for template in templates_to_check:
        views_using_template = [
            v for v in DB.FilteredElementCollector(revit.doc)
            .OfClass(DB.View)
            if v.ViewTemplateId == template.Id
        ]
        # Checks whether all views that are using a specific template are included in the views_to_delete
        if all(view in views_to_delete for view in views_using_template):
            templates_unused.append(template)

    # Delete views not on sheets and with no dependents
    with revit.Transaction("Delete unused views"):
        deleted_views_counter = 0
        for view in views_to_delete:
            try:
                revit.doc.Delete(view.Id)
                deleted_views_counter += 1
            except Exception as e:
                logger.error("Error deleting view {}: {}".format(view.Name, str(e)))
            
        if deleted_views_counter > 1:
            forms.alert("{} views deleted.".format(deleted_views_counter))
        elif deleted_views_counter == 1:
            forms.alert("1 view deleted.")
        else:
            forms.alert("No views deleted.")

    # Ask user about deleting unused templates
    if templates_unused:
        template_names = [t.Name for t in templates_unused]
        delete_templates = forms.alert(
            "The following view templates are not used any more:\n{}\n\nDo you want to delete them?".format(
                ", ".join(template_names)
            ),
            options=["Yes", "No"]
        )
        if delete_templates == "Yes":
            deleted_vt_counter = 0
            with revit.Transaction("Delete unused View Templates"):
                for template in templates_unused:
                    revit.doc.Delete(template.Id)
                if deleted_vt_counter > 1:
                    forms.alert("{} View Templates deleted.".format(deleted_vt_counter))
                elif deleted_vt_counter == 1:
                    forms.alert("1 View Template deleted.")

    # Display views on sheets or with dependents
    if new_views_on_sheets_or_with_dependents:
        output = script.get_output()
        output.print_md("### Number of Views on Sheets or with dependent Views -> {}".format(len(new_views_on_sheets_or_with_dependents)))
    for view_info in new_views_on_sheets_or_with_dependents:
        sheet_info = view_info["sheet"]
        dependents_info = view_info["dependents"]
        output.print_md("**View:** {}".format(view_info['view_name']))
        if sheet_info:
            output.print_md("  * On Sheet: {} - {}".format(sheet_info.SheetNumber, sheet_info.Name))
        if dependents_info:
            for dep in dependents_info:
                dep_sheet = dep["sheet"]
                if dep_sheet:
                    output.print_md("    - Dependent View: {} (Sheet: {} - {})".format(
                        dep['name'], dep_sheet.SheetNumber, dep_sheet.Name))
                else:
                    output.print_md("    - Dependent View: {}".format(dep['name']))

# if __name__ == "__main__": construct in Python is a standard way to ensure that a block of code is executed only when the script is run directly, and not when it is imported as a module into another script. 
if __name__ == "__main__":
    main()