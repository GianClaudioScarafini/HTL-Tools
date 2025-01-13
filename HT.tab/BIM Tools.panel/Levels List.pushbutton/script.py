# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-11-27
# Version: 1.0.0
# Description: Creates a list of levels with number of dependent views and elements and a sub-list of dependent view names.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

from pyrevit import revit, DB, script
from collections import defaultdict

# Initialize output
output = script.get_output()
doc = revit.doc

# Collect all levels in the document
levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()

# Sort levels by name and then elevation
sorted_levels = sorted(levels, key=lambda l: (l.Name, l.Elevation))

# Function to collect dependent views and elements grouped by category
def get_dependents(level):
    dependent_views = []
    excluded_elements = 0
    category_counts = defaultdict(int)  # Dictionary to store category names and counts
    
    # Get the ID of the level
    level_id = level.Id
    
    # Find dependent elements
    for dependent_id in level.GetDependentElements(None):
        dependent = doc.GetElement(dependent_id)
        
        # Sort into views and other elements
        if isinstance(dependent, DB.View):
            dependent_views.append(dependent)
        else:
            # Exclude elements that are the level itself
            if dependent.Id == level_id:
                continue
            # Exclude elements with null Category
            if dependent.Category is None:
                continue
            
            # Add to category count
            category_name = dependent.Category.Name
            excluded_categories = ["Sun Path",
                          "Work Plane Grid",
                          "Viewports",
                          "Automatic Sketch Dimensions",
                          ""]
            if category_name in excluded_categories:
                excluded_elements += 1
                continue
            category_counts[category_name] += 1
    
    return dependent_views, category_counts, excluded_elements

output.print_md("# Level Dependency Report")
output.print_md("**Analysing levels and their dependencies...**")
output.print_md("*Excluded categories: Sun Path, Work Plane Grid, Viewports, Automatic Sketch Dimensions*")

# Process each level
for level in sorted_levels:
    output.print_md("---")  # Separator

    level_name = level.Name
    dependent_views, category_counts, excluded_elements = get_dependents(level)
    
    # Count dependent views and elements
    num_views = len(dependent_views)
    num_elements = sum(category_counts.values())
    
    # Header for level
    if num_views == 0 and num_elements == 0:
        # Highlight levels with no dependencies
        level_header = "Level: **{} - CAN BE DELETED**".format(level_name)
    else:
        level_header = "Level: **{}**".format(level_name)
    output.print_md(level_header)
    
    # Output dependency summary
    if excluded_elements > 0:
        excluded_text = " (Excluded: {})".format(excluded_elements)
    else:
        excluded_text = ""
    output.print_md("> Dependent Views: {}".format(num_views))
    output.print_md("> Dependent Elements: {}".format(num_elements) + excluded_text)
    
    if num_views > 0:
        # List dependent views
        output.print_md("> Views:")
        for view in dependent_views:
            output.print_md("> - {}".format(view.Name))
    
    if category_counts:
        # List dependent elements grouped by category
        output.print_md("> Elements by Category:")
        for category_name, count in category_counts.items():
            output.print_md("> - {}: {}".format(category_name, count))