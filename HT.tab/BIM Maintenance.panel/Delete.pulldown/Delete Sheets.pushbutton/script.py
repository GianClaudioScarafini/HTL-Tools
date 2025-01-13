# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-12-06
# Version: 1.0.0
# Description: Safely delete sheets from the Project Browser with all Legends, Schedules an Views by checking if they do not have any dependent views.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Import pyRevit modules
from pyrevit import revit, DB, forms, script
from collections import OrderedDict

# Initialize logger and output
logger = script.get_logger()
output = script.get_output()

# Check Revit version
app = revit.doc.Application
ver = int(app.VersionNumber)

# Helper function: Get all views on a sheet
def get_views_on_sheet(sheet):
    """
    Returns all views and schedules placed on a given sheet.
    Optimized to avoid repeated filtering.
    """
    # Get regular views placed on the sheet
    views = [revit.doc.GetElement(view_id) for view_id in sheet.GetAllPlacedViews()]
    
    # Get all dependent elements of the sheet and filter for ScheduleSheetInstance
    dependent_elements = sheet.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_ScheduleGraphics))
    schedules = [
        revit.doc.GetElement(elem_id)
        for elem_id in dependent_elements
        if isinstance(revit.doc.GetElement(elem_id), DB.ScheduleSheetInstance) and not revit.doc.GetElement(elem_id).IsTitleblockRevisionSchedule
    ]

    # Combine views and schedules
    return views + schedules

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

# Helper function: Check for dependent views
def get_all_dependent_views(view):
    """Collect all dependent views for a given view."""
    dependent_views = view.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Views))
    # Method GetDependentViewIds() returns only views create and visible in the browser tree as dependent views but not dependent detail callouts.
    # Method GetDependentElements() returns all dependent views including detail callouts.
    # Remove itself. (View is als listing itself in dependent elements)
    dependent_ids = [v for v in dependent_views if v != view.Id]
    return list(set(dep_id for dep_id in dependent_ids if revit.doc.GetElement(dep_id)))

# Legacy method to find schedule instances (Revit 2022 and earlier)
def find_other_sheets_with_schedule_legacy(schedule, selected_sheets):
    dependent_ids = schedule.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_ScheduleGraphics))
    other_sheets_with_schedule = []
    for instance_id in dependent_ids:
        instance = revit.doc.GetElement(instance_id)
        if instance:
            parent_sheet_id = instance.OwnerViewId
            if parent_sheet_id != DB.ElementId.InvalidElementId:
                selected_sheets_ids = [sheet.Id for sheet in selected_sheets]
                if isinstance(sheet, DB.ViewSheet) and parent_sheet_id not in selected_sheets_ids:
                    sheet = revit.doc.GetElement(parent_sheet_id)
                    other_sheets_with_schedule.append(sheet)
    return other_sheets_with_schedule

# Unified function for schedule instances
def find_other_sheets_with_schedule(schedule, selected_sheets):
    if ver >= 2023:
        schedule_instances = schedule.GetScheduleInstances()
        other_sheets_with_schedule = []
        for instance_id in schedule_instances:
            instance = revit.doc.GetElement(instance_id)
            if instance:
                parent_sheet_id = instance.OwnerViewId
                if parent_sheet_id != DB.ElementId.InvalidElementId:
                    selected_sheets_ids = [sheet.Id for sheet in selected_sheets]
                    if isinstance(sheet, DB.ViewSheet) and parent_sheet_id not in selected_sheets_ids:
                        sheet = revit.doc.GetElement(parent_sheet_id)
                        other_sheets_with_schedule.append(sheet)
        return other_sheets_with_schedule
    else:
        return find_other_sheets_with_schedule_legacy(schedule,selected_sheets)

# Helper function: Check if a legend is used elsewhere
def get_legend_usage_count(legend, selected_sheets):
    count = 0
    for sheet in DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSheet):
        if sheet not in selected_sheets and legend.Id in sheet.GetAllPlacedViews():
        # GetAllPlacedViews() does not return Schedules
            count += 1
    return count

def get_legend_viewport(sheet, legend):
    viewports = sheet.GetAllViewports()
    for viewport_id in viewports:
        viewport = revit.doc.GetElement(viewport_id)
        if legend.Id == viewport.ViewId:
            return viewport_id

def main():
    # Collect selected sheets
    selected_elements = revit.get_selection()
    selected_sheets = [elem for elem in selected_elements if isinstance(elem, DB.ViewSheet)]

    if not selected_sheets:
        forms.alert("No sheets selected. Please select sheets without any views placed on them.", exitscript=True)

    # Sort selected sheets by sheet number
    selected_sheets.sort(key=lambda s: s.SheetNumber)

    # Prepare sets to avoid duplicates
    views_to_delete = set()
    schedules_to_delete = set()
    legends_to_delete = set()
    all_views_with_dependents = []
    all_schedules_with_other_sheets = []
    all_schedules_instances_on_sheets = []
    all_legends_on_sheets = []

    # Process each sheet
    for sheet in selected_sheets:
        views_with_dependents = []
        schedules_with_other_sheets = []
        legends_with_other_instances = []
        sheet_name = "{} - {}".format(sheet.SheetNumber, sheet.Name)
        output.print_md("### Processing Sheet: **{}**".format(sheet_name))
        
        views = get_views_on_sheet(sheet)

        # Process views
        for view in views:
            if isinstance(view, DB.View) and not isinstance(view, (DB.ViewSchedule, DB.ScheduleSheetInstance)) and view.ViewType != DB.ViewType.Legend:
                all_dependents = get_all_dependent_views(view)
                if not all_dependents:
                    views_to_delete.add(view)  # Add to set to prevent duplicates
                else:
                    dependents_info = [
                        {
                            "name": revit.doc.GetElement(dep_id).Name,
                            "sheet": get_view_sheet(revit.doc.GetElement(dep_id))
                        }
                        for dep_id in all_dependents
                    ]
                    views_with_dependents.append({
                        "view_name": view.Name,
                        "all_dependents": dependents_info,
                        "view": view
                    })
                    all_views_with_dependents.append(views_with_dependents)

        # Process schedules
        for view in views:
            if isinstance(view, DB.ScheduleSheetInstance):
                schedule_instance = view
                all_schedules_instances_on_sheets.append(schedule_instance)
                schedule = revit.doc.GetElement(schedule_instance.ScheduleId)
                # Check if the schedule is used elsewhere
                other_sheets = find_other_sheets_with_schedule(schedule, selected_sheets)
                # Sort other sheets by sheet number
                other_sheets.sort(key=lambda s: s.SheetNumber)
                other_sheets = [s for s in other_sheets if s.Id != sheet.Id]
                if not other_sheets:
                    schedules_to_delete.add(schedule)
                else:
                    schedules_with_other_sheets.append({
                        "schedule_name": schedule.Name,
                        "other_sheets": other_sheets
                    })
                    all_schedules_with_other_sheets.append(schedules_with_other_sheets)

        # Process legends
        for view in views:
            if not isinstance(view, DB.ScheduleSheetInstance) and view.ViewType == DB.ViewType.Legend:
                legend = view
                legend_viewport_id = get_legend_viewport(sheet, legend)
                all_legends_on_sheets.append(legend_viewport_id)
                # Check if the legend is used on other sheets then selected
                usage_count = get_legend_usage_count(view, selected_sheets)
                if usage_count == 0:
                    legends_to_delete.add(view)
                else:
                    legends_with_other_instances.append({
                        "legend_name": view.Name,
                        "usage_count": usage_count
                    })
        
        # Print dependencies
        for view_info in views_with_dependents:
            output.print_md(">**View with Dependents:** {}".format(view_info['view_name']))
            for dep in view_info['all_dependents']:
                output.print_md(">    - Dependent View: {}".format(dep["name"]))

        for schedule_info in schedules_with_other_sheets:
            output.print_md(">**Schedule:** {} is also placed on other sheets:".format(schedule_info['schedule_name']))
            for other_sheet in schedule_info['other_sheets']:
                output.print_md(">    - {} - {}".format(other_sheet.SheetNumber, other_sheet.Name))

        for legend_info in legends_with_other_instances:
            output.print_md(">**Legend:** {} is used on {} other sheet(s).".format(
                legend_info['legend_name'], legend_info['usage_count']
            ))

        if not views_with_dependents and not schedules_with_other_sheets and not legends_with_other_instances:
            output.print_md("No dependencies found.")

    # Check for parent views whose dependents are all marked for deletion
    additional_deletions = set()
    new_views_with_dependents = []
    for view in views_with_dependents:
        parent_view = view["view"]
        dependent_views = view["all_dependents"]
        if dependent_views and all(dep_id in views_to_delete for dep_id in dependent_views):
            additional_deletions.add(parent_view)
        else:
            new_views_with_dependents.append(view)
    views_to_delete.update(additional_deletions)

    # Prepare dialogue for deletion process
    # User can decide what to delete
    # Sheets with all: views (except views with dependents), schedules and legends (if not placed on other sheets)
    plural_sheets = "s" if len(selected_sheets) > 1 else ""
    DeleteSheets = "Delete {} sheet{} (deletes all below)".format(len(selected_sheets), plural_sheets)

    # Confirmation dialog switches - on/off (using OrderedDict)
    switches = OrderedDict()

    # Add switches in the desired order
    switches[DeleteSheets] = False

    DeleteViews = ""
    if len(views_to_delete) > 0:
        if new_views_with_dependents:
            view_text = "except {} not deleted due to dependencies".format(len(new_views_with_dependents))
        else:
            view_text = "all views without dependent views"
        plural_views = "s" if len(views_to_delete) > 1 else ""
        DeleteViews = "Delete {} view{} ({})".format(len(views_to_delete), plural_views, view_text)
        switches[DeleteViews] = False

    DeleteSchedules = ""
    if len(schedules_to_delete) > 0:
        if all_schedules_with_other_sheets:
            schedule_text = "except {} used on other sheets".format(len(all_schedules_with_other_sheets))
        else:
            schedule_text = "all schedules not placed on other sheets"
        plural_schedules = "s" if len(schedules_to_delete) > 1 else ""
        DeleteSchedules = "Delete {} schedule{} ({})".format(len(schedules_to_delete), plural_schedules, schedule_text)
        switches[DeleteSchedules] = False

    DeleteLegends = ""
    if len(legends_to_delete) > 0:
        plural_legends = "s" if len(legends_to_delete) > 1 else ""
        DeleteLegends = "Delete {} Legend{} (not used elsewhere)".format(len(legends_to_delete), plural_legends)
        switches[DeleteLegends] = False

    # Remove Schedule Instances from sheets only
    RemoveScheduleInstances = ""
    if len(all_schedules_instances_on_sheets) > 0:
        plural_schedule_instances = "s" if len(all_schedules_instances_on_sheets) > 1 else ""
        RemoveScheduleInstances = "Remove {} Schedule Instance{} from selected Sheet{}".format(
            len(all_schedules_instances_on_sheets), plural_schedule_instances, plural_sheets
        )
        switches[RemoveScheduleInstances] = False

    # Remove Legend Viewports from sheets only
    RemoveLegendViewports = ""
    if len(all_legends_on_sheets) > 0:
        plural_legends_vp = "s" if len(all_legends_on_sheets) > 1 else ""
        RemoveLegendViewports = "Remove {} legend Viewport{} from selected Sheet{}".format(
            len(all_legends_on_sheets), plural_legends_vp, plural_sheets
        )
        switches[RemoveLegendViewports] = False

    # Emphasise Delete Sheets because it deletes everything
    cfgs = {DeleteSheets: { 'background': '0xFFFF55'}}
    _, toggled_switches = forms.CommandSwitchWindow.show(
    [],
    switches=switches,
    message="Select which elements to delete (toggle switches):",
    config=cfgs
    )

    ConfirmationMessage = "Deletion completed successfully for:"
    if any(toggled_switches.values()):
        with revit.Transaction("Delete Selected Sheets and Associated Elements"):
            # get(xxx, False) returns False if the key is not found
            if toggled_switches.get(DeleteSheets, False):  # If 'Delete Sheets' is ON
                for sheet in selected_sheets:
                    revit.doc.Delete(sheet.Id)
                toggled_switches.update({DeleteViews: True})
                toggled_switches.update({DeleteSchedules: True})
                toggled_switches.update({DeleteLegends: True})
                plural = "s" if len(selected_sheets) > 1 else ""
                ConfirmationMessage += "\n- {} Sheet{}".format(len(selected_sheets), plural)

            if toggled_switches.get(DeleteViews, False):  # If 'Delete Views' is ON
                count = 0
                not_deleted_views = ""
                for view in views_to_delete:
                    if revit.doc.GetElement(view.Id):  # Ensure view still exists
                        revit.doc.Delete(view.Id)
                        count += 1
                    else:
                        not_deleted_views += "{}".format(view.Name)
                plural = "s" if count > 1 else ""
                ConfirmationMessage += "\n- {} View{}".format(count, plural)
                if not_deleted_views:
                    ConfirmationMessage += "\n  not deleted Views: "+not_deleted_views
                

            if toggled_switches.get(DeleteSchedules, False):  # If 'Delete Schedules' is ON
                count = 0
                for schedule in schedules_to_delete:
                    if revit.doc.GetElement(schedule.Id):  # Ensure schedule still exists
                        revit.doc.Delete(schedule.Id)
                        count += 1
                plural = "s" if count > 1 else ""
                ConfirmationMessage += "\n- {} Schedule{}".format(count, plural)

            if toggled_switches.get(DeleteLegends, False):  # If 'Delete Legends' is ON
                count = 0
                for legend in legends_to_delete:
                    if revit.doc.GetElement(legend.Id):  # Ensure legend still exists
                        revit.doc.Delete(legend.Id)
                        count += 1
                plural = "s" if count > 1 else ""
                ConfirmationMessage += "\n- {} Legend{}".format(count, plural)

            # When all sheets are deleted schedule instances are always removed from sheets
            if not toggled_switches.get(DeleteSheets, False) and toggled_switches.get(RemoveScheduleInstances, False):  # If 'Remove Schedule Instances' is ON
                count = 0
                for instance in all_schedules_instances_on_sheets:
                    if revit.doc.GetElement(instance.Id):
                        revit.doc.Delete(instance.Id)
                        count += 1
                plural = "s" if count > 1 else ""
                ConfirmationMessage += "\n- {} Schedule Instance{} from selected Sheet{}".format(count, plural, plural_sheets)

            # When all sheets are deleted legend viewports are always removed from sheets
            if not toggled_switches.get(DeleteSheets, False) and toggled_switches.get(RemoveLegendViewports, False):  # If 'Remove Legend Viewports' is ON
                count = 0
                for legend_vp_id in all_legends_on_sheets:
                    if revit.doc.GetElement(legend_vp_id):
                        revit.doc.Delete(legend_vp_id)
                        count += 1
                plural = "s" if count > 1 else ""
                ConfirmationMessage += "\n- {} Legend Viewport{} from selected Sheet{}".format(count, plural, plural_sheets)
    else:
        ConfirmationMessage = "No switch toggled on. Nothing deleted."

    forms.alert(ConfirmationMessage, title="Deletion Confirmation")

if __name__ == "__main__":
    main()
