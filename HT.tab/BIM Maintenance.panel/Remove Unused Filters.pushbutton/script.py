# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-10-29
# Version: 1.0.0
# Removes unused filters from View Templates and Views. 
# These are filters added to V/G Filters tab but not doing any changes.
# Filters enable/disable property is ignored
# Tested with: Revit +2022
# Requirements: pyRevit add-in
# Import pyRevit modules

from pyrevit import revit, DB, script, forms

# Document and transaction setup
doc = revit.doc

# Define helper functions to gather and check filters
class AllViewTemplates(forms.TemplateListItem):
    @property
    def name(self):
        viewTemplate = ""
        if self.item.IsTemplate:
            viewTemplate = " View Template"
        SheetNumber = self.item.LookupParameter("Sheet Number")
        sheetNumberText = ""
        if SheetNumber and SheetNumber.AsString() != "---":
            sheetNumberText = " - Sheet Number: "+SheetNumber.AsString()
        return self.item.Name + " ( " + self.item.ViewType.ToString() + viewTemplate + " )" + sheetNumberText

def get_all_views_and_templates(doc):
    """
    Get all view templates and views with no view templates applied.
    """
    # Collect all views in the document (including templates)
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View)

    # Prepare two lists: one for view templates, one for views without templates
    view_templates = []
    views_no_templates = []
    
    # Helper function to check if a view has a template applied
    def get_view_template(view):
        if not view: 
            return None
        elif hasattr(view, "ViewTemplateId"):
            # ViewTemplateId of -1 means no template is applied
            if view.ViewTemplateId.IntegerValue == -1: 
                return None
            else: 
                return view.Document.GetElement(view.ViewTemplateId)
        return None

    views = list(filter(lambda x: (
        x.ViewType != DB.ViewType.Legend and
        x.ViewType != DB.ViewType.Schedule and
        x.ViewType != DB.ViewType.DrawingSheet and
        x.ViewType != DB.ViewType.ProjectBrowser and
        x.ViewType != DB.ViewType.Report and
        x.ViewType != DB.ViewType.SystemBrowser and
        x.ViewType != DB.ViewType.EngineeringPlan and
        x.ViewType != DB.ViewType.CostReport and
        x.ViewType != DB.ViewType.LoadsReport and
        x.ViewType != DB.ViewType.Walkthrough and
        x.ViewType != DB.ViewType.Rendering and
        x.ViewType != DB.ViewType.Internal
    ), collector))

    # Loop through all collected views
    for view in views:
        # Check if the view is a template
        if view.IsTemplate:
            # Check if view is a View Template but we will remove unused filters from it regardless if Filter overrides are enabled or not.
            view_templates.append(view)
        else:
            # Check if the view has a template applied
            if get_view_template(view) is None:
                views_no_templates.append(view)
            else:
                # Check if the view has a template applied but the template has a filter overrides not enabled.
                viewTemplate = doc.GetElement(view.ViewTemplateId)
                view_template_not_controlled_settings = viewTemplate.GetNonControlledTemplateParameterIds()
                if -1006964 in view_template_not_controlled_settings:
                    views_no_templates.append(view)
    # Built-in parameter: VIS_GRAPHICS_FILTERS -1006964 V/G Overrides Filters

    # Combine the two lists: templates and views without templates
    view_templates.sort(key=lambda view: view.Name)
    views_no_templates.sort(key=lambda view: view.Name)
    combined_views = view_templates + views_no_templates
    count = len(combined_views)

    returned_views = forms.SelectFromList.show(
            [ AllViewTemplates(x) for x in combined_views ],
            title='List of '+str(count)+' View Templates and Views with no View Template - Select which to clean from unused filters.',
            width=800,
            button_name='Clean these views',
            multiselect=True
        )
    return returned_views

def get_unused_filters_in_view(view):
    """Check filters in the view and return unused filter IDs."""
    try:
        filter_ids = view.GetFilters()
        unused_filter_ids = []
        if filter_ids:
            for filter_id in filter_ids:
                try:
                    visibility = view.GetFilterVisibility(filter_id)
                except Exception as e:
                    pass # this exception happens when a view doesn't support filters
                if visibility: # must be visible with no overrides checked later. Not visible means it is used.
                    unused_filter_ids.append(filter_id)
        return unused_filter_ids
    except Exception as e:
        return []

def check_filter_overrides(view, filter_id):
    """Check if a filter has no overrides applied."""
    # Color and Haleftone must be False then it's not set up, Halftone not enabled
 	# Id, Lineweight or Patter must be -1 then not set up
 	# Background or Foreground Visibility must be True then it's not turn off
 	# Transparency must be 0 then it's not transparent
    overrides = view.GetFilterOverrides(filter_id)
    if (
        overrides.CutBackgroundPatternColor.IsValid == False and
        overrides.CutForegroundPatternColor.IsValid == False and
        overrides.CutLineColor.IsValid == False and
        overrides.ProjectionLineColor.IsValid == False and
        overrides.SurfaceBackgroundPatternColor.IsValid == False and
        overrides.SurfaceForegroundPatternColor.IsValid == False and
        overrides.Halftone == False and
        overrides.CutForegroundPatternId.ToString() == "-1" and
        overrides.CutBackgroundPatternId.ToString() == "-1" and
        overrides.CutLinePatternId.ToString() == "-1" and
        overrides.ProjectionLinePatternId.ToString() == "-1" and
        overrides.CutLineWeight == -1 and
        overrides.ProjectionLineWeight == -1 and
        overrides.SurfaceBackgroundPatternId.ToString() == "-1" and
        overrides.SurfaceForegroundPatternId.ToString() == "-1" and
        overrides.IsCutBackgroundPatternVisible == True and
        overrides.IsCutForegroundPatternVisible == True and
        overrides.IsSurfaceBackgroundPatternVisible == True and
        overrides.IsSurfaceForegroundPatternVisible == True and
        overrides.Transparency == 0
    ):
        return True
    return False

def remove_unused_filters(doc, views):
    """Remove unused filters from the provided views."""
    deleted_filters_total_counter = 0
    views_counter = 0
    error_report = []

    # Start the transaction for filter removal
    with revit.Transaction("Remove unused filters"):
        global output_text
        for view in views:
            deleted_view_filters_counter = 0
            deleted_view_filter_names = []
            try:
                unused_filters = get_unused_filters_in_view(view)
                for filter_id in unused_filters:
                    filter_name = doc.GetElement(filter_id).Name
                    if check_filter_overrides(view, filter_id):
                        view.RemoveFilter(filter_id)
                        deleted_filters_total_counter += 1
                        deleted_view_filters_counter += 1
                        deleted_view_filter_names.append(filter_name)
                if deleted_view_filter_names:
                    views_counter += 1
            except Exception as e:
                error_report.append(str(e))
            if deleted_view_filters_counter > 0:
                viewTemplate = ""
                if view.IsTemplate:
                    viewTemplate = " ( View Template )"
                filter_text = " filter" if deleted_filters_total_counter == 1 else " filters"
                output_text += "\n\n**" + view.Name + viewTemplate + "**<br>" + str(deleted_view_filters_counter) + " unused " + filter_text+ " removed: " + ', '.join(deleted_view_filter_names)

    return deleted_filters_total_counter, views_counter, error_report

output_text = "Summary:"
# Main execution
all_views = get_all_views_and_templates(doc)
if all_views:
    deleted_filters_total_counter, views_counter, errors = remove_unused_filters(doc, all_views)

    # Report the result to the user
    output = script.get_output()
    filter_text = " filter" if deleted_filters_total_counter == 1 else " filters"
    view_text = " view." if views_counter == 1 else " views."
    output.print_md("**{}**{} removed from **{}**{}".format(deleted_filters_total_counter, filter_text,views_counter, view_text))
    if deleted_filters_total_counter > 0:
        output.print_md(output_text)

    if errors:
        output.print_md("Errors occurred during execution:")
        for error in errors:
            output.print_md(" - {}".format(error))
