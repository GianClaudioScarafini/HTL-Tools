# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-12-06
# Version: 1.0.0
# Description: Delete Legends safely checking first in the Output window how often legends are used and on which sheets.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Import pyRevit modules
from pyrevit import revit, DB, script, forms

# app = revit.doc.Application
# ver = int(app.VersionNumber)
# if ver <= 2032:
#     parameter_type = DB.ParameterType
# doc = revit.doc

# Initialize logger and output
logger = script.get_logger()
output = script.get_output()

# Helper class for displaying legends in a selection form
class LegendViewItem(forms.TemplateListItem):
    @property
    def name(self):
        # Construct display name for the legend view
        sheet_info = self.item['sheets']
        if sheet_info:
            sheet_details = ", ".join(
                sheet.SheetNumber for sheet in sheet_info
            )
            if len(sheet_info) == 1:
                return "{} | on {} sheet: {}".format(self.item['legend_name'], len(sheet_info), sheet_details)
            else:
                return "{} | on {} sheets: {}".format(self.item['legend_name'], len(sheet_info), sheet_details)
        else:
            return "{} | > Not on any Sheet > CAN BE DELETED!".format(self.item['legend_name'])

    def __bool__(self):
        return True

def collect_legend_views(views):
    """Collects all legend views in the model and checks their sheet placement."""
    legends_data = []
    legend_views = []
    for elem in views:
        if isinstance(elem, DB.View) and not isinstance(elem, (DB.ViewSchedule, DB.ViewSheet)):
            # Check if the view's type matches the specified criteria for deletion
            if elem.ViewType == DB.ViewType.Legend:
                legend_views.append(elem)

    # Exit if no valid views were selected
    if not legend_views:
        forms.alert("No Legends in selection or in the model. Ensure valid Legends are selected and retry.", exitscript=True)

    # Process each legend view
    for legend in legend_views:
        legend_name = legend.Name
        # Find all sheets where the legend is placed
        sheets_with_legend = []
        sheets = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSheet).ToElements()
        for sheet in sheets:
            if legend.Id in sheet.GetAllPlacedViews():
                sheets_with_legend.append(sheet)
        
        # Store data about the legend
        legends_data.append({
            "legend": legend,
            "legend_name": legend_name,
            "sheets": sheets_with_legend
        })
    
    # Sort by number of sheets, then alphabetically by name
    legends_data.sort(key=lambda x: (len(x['sheets']), x['legend_name']))
    return legends_data

def main():
    legends_data = []
    # Collect the selected elements
    selected_elements = revit.get_selection()
    if not selected_elements:
        select_all = forms.alert(
            "No Legends selected. Select all Legends in the project?",
            yes = True,
            no = True,
            exitscript=True
        )
        if select_all:
            views = DB.FilteredElementCollector(revit.doc).OfClass(DB.View).ToElements()
            legends_data = collect_legend_views(views)
    else:
        legends_data = collect_legend_views(selected_elements)

    # Print legends and their sheet placements
    output.print_md("### Legend Views and Sheet Placement")
    for legend_info in legends_data:
        legend_name = legend_info['legend_name']
        sheets = legend_info['sheets']
        if sheets:
            output.print_md("**Legend:** {} | Placed on {} sheet(s)".format(legend_name, len(sheets)))
            for sheet in sheets:
                HTL_Sheet_Type_param = sheet.LookupParameter("HTL Sheet Type").AsValueString() or "???"
                HTL_Sheet_Sub_Type_param = sheet.LookupParameter("HTL Sheet Sub Type").AsValueString() or "???"
                output.print_md("  * {} - {} | Folder: {} / {}".format(sheet.SheetNumber, sheet.Name, HTL_Sheet_Type_param, HTL_Sheet_Sub_Type_param))
        else:
            output.print_md("**Legend:** {} | NOT ON ANY SHEET!".format(legend_name))

    # Show selection form
    selected_legends = forms.SelectFromList.show(
        [LegendViewItem(legend_info) for legend_info in legends_data],
        title='Select Legends to Delete',
        width=1070,
        button_name='Delete Selected Legends',
        multiselect=True
    )

    # Confirm deletion
    if selected_legends:
        plural = "s" if len(selected_legends) > 1 else ""
        confirmation = forms.alert(
            "You are about to delete {} legend{}. Do you want to continue?".format(len(selected_legends), plural),
            options=["Yes", "No"]
        )
        if confirmation == "Yes":
            with revit.Transaction("Delete Selected Legends"):
                for legend_info in selected_legends:
                    revit.doc.Delete(legend_info.item['legend'].Id)
            forms.alert("{} legend{} deleted successfully.".format(len(selected_legends), plural))

if __name__ == "__main__":
    main()
