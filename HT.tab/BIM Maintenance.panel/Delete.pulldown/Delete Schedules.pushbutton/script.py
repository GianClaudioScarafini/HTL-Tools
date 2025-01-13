# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-12-10
# Version: 1.0.0
# Description: Delete Schedules safely checking first in the Output window how often they are used and on which sheets.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Import pyRevit modules
from pyrevit import revit, DB, forms, script

# Initialize logger and output
logger = script.get_logger()
output = script.get_output()

# Check Revit version
app = revit.doc.Application
ver = int(app.VersionNumber)

# Helper class for displaying schedules in a selection form
class ScheduleViewItem(forms.TemplateListItem):
    @property
    def name(self):
        if self.item['sheets']:
            sheet_details = ", ".join(
                "{} - {}".format(sheet.SheetNumber, sheet.Name) for sheet in self.item['sheets']
            )
            return "{} | {} Sheets: {}".format(self.item['schedule_name'], len(self.item['sheets']), sheet_details)
        else:
            return "{} | > Not on any Sheet > CAN BE DELETED!".format(self.item['schedule_name'])

    def __bool__(self):
        return True

def find_schedule_instances_legacy(schedule):
    """Find schedule instances for Revit 2022 and earlier."""
    dependent_ids = schedule.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_ScheduleGraphics))
    sheets_with_schedule = []
    for instance_id in dependent_ids:
        instance = revit.doc.GetElement(instance_id)
        if instance:
            parent_sheet_id = instance.OwnerViewId
            if parent_sheet_id != DB.ElementId.InvalidElementId:
                sheet = revit.doc.GetElement(parent_sheet_id)
                if isinstance(sheet, DB.ViewSheet):
                    sheets_with_schedule.append(sheet)
    return sheets_with_schedule

def find_schedule_instances(schedule):
    """Find schedule instances based on Revit version."""
    if ver >= 2023:
        # Use GetScheduleInstances() for Revit 2023 and newer
        schedule_instances = schedule.GetScheduleInstances()
        sheets_with_schedule = []
        for instance_id in schedule_instances:
            instance = revit.doc.GetElement(instance_id)
            if instance:
                parent_sheet_id = instance.OwnerViewId
                if parent_sheet_id != DB.ElementId.InvalidElementId:
                    sheet = revit.doc.GetElement(parent_sheet_id)
                    if isinstance(sheet, DB.ViewSheet):
                        sheets_with_schedule.append(sheet)
        return sheets_with_schedule
    else:
        # Use legacy method for Revit 2022 and earlier
        return find_schedule_instances_legacy(schedule)

def collect_schedule_views(schedules):
    schedules_data = []

    for schedule in schedules:
        # Exclude title block revision schedules
        if schedule.IsTitleblockRevisionSchedule:
            continue

        schedule_name = schedule.Name
        sheets_with_schedule = find_schedule_instances(schedule)

        # Store data about the schedule
        schedules_data.append({
            "schedule": schedule,
            "schedule_name": schedule_name,
            "sheets": sheets_with_schedule
        })

    # Sort by number of sheets, then alphabetically by name
    schedules_data.sort(key=lambda x: (len(x['sheets']), x['schedule_name']))
    return schedules_data

def main():
    # Collect the selected elements
    selected_elements = revit.get_selection()
    if not selected_elements:
        select_all = forms.alert(
            "No schedules selected. Select all Schedules in the project?",
            yes = True,
            no = True,
            exitscript=True
        )
        if select_all:
            """Collects all schedules in the model"""
            schedules = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSchedule).ToElements()
            schedules_data = collect_schedule_views(schedules)
    else:
        schedules = []
        for elem in selected_elements:
            if isinstance(elem, DB.ViewSchedule):
                schedules.append(elem)
        schedules_data = collect_schedule_views(schedules)

    # Print schedules and their sheet placements
    output.print_md("### Schedule Views and Sheet Placement")
    for schedule_info in schedules_data:
        schedule_name = schedule_info['schedule_name']
        sheets = schedule_info['sheets']
        if sheets:
            plural = "s" if len(sheets) > 1 else ""
            output.print_md("**Schedule:** {} | Placed on {} sheet{}".format(schedule_name, len(sheets), plural))
            for sheet in sheets:
                output.print_md("  * {} - {}".format(sheet.SheetNumber, sheet.Name))
        else:
            output.print_md("**Schedule:** {} | > NOT ON ANY SHEET!".format(schedule_name))

    # Show selection form
    selected_schedules = forms.SelectFromList.show(
        [ScheduleViewItem(schedule_info) for schedule_info in schedules_data],
        title='Select Schedules to Delete',
        width=1070,
        button_name='Delete Selected Schedules',
        multiselect=True
    )

    # Confirm deletion
    if selected_schedules:
        confirmation = forms.alert(
            "You are about to delete {} schedule(s). Do you want to continue?".format(len(selected_schedules)),
            options=["Yes", "No"]
        )
        if confirmation == "Yes":
            with revit.Transaction("Delete Selected Schedules"):
                for schedule_info in selected_schedules:
                    revit.doc.Delete(schedule_info['schedule'].Id)
            forms.alert("{} schedule(s) deleted successfully.".format(len(selected_schedules)))
        else:
            forms.alert("Schedule deletion canceled.")

if __name__ == "__main__":
    main()
