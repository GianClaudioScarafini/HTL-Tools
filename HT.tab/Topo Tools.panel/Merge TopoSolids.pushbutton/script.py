# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-11-04
# Version: 1.0.0
# Description: Joins many TopoSolids into one. Before merging, copy any Sub-divisions and paste them in the same place after the conversion.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Explicit imports for pyRevit scripting and Revit API
from pyrevit import script
from pyrevit import revit, DB, UI, forms
from System.Collections.Generic import List  # Import for IList conversion

# Initialize logger
logger = script.get_logger()

# Get Revit application and active document
uidoc = revit.uidoc
doc = revit.doc

# Custom selection filter class for selecting only Toposolids
class ToposolidSelectionFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.Toposolid)

    def AllowReference(self, reference, position):
        return False

# Function to pick multiple Toposolids using a custom filter
def pick_toposolids():
    try:
        # Prompt user to pick multiple Toposolids using the custom filter
        selection = uidoc.Selection
        refs = selection.PickObjects(UI.Selection.ObjectType.Element, ToposolidSelectionFilter(), "Select Toposolids")
        elements = [doc.GetElement(ref) for ref in refs]
        return elements
    except Exception as e:
        logger.error("Error selecting Toposolids: {}".format(str(e)))
        script.exit()

# Function to extract unique points with the biggest Z value from Toposolids
def extract_unique_points_from_toposolids(toposolids):
    # Dictionary to store points with (X, Y) as keys and max Z value as values
    points_dict = {}

    for toposolid in toposolids:
        options = DB.Options()
        geometry = toposolid.get_Geometry(options)

        for geom_obj in geometry:
            if isinstance(geom_obj, DB.Solid):
                for face in geom_obj.Faces:
                    mesh = face.Triangulate()
                    for vertex in mesh.Vertices:
                        # Use a tuple of (rounded X, rounded Y) to ensure unique XY locations
                        xy_tuple = (round(vertex.X, 2), round(vertex.Y, 2))
                        
                        # Update the dictionary to store the point with the highest Z value
                        if xy_tuple not in points_dict or vertex.Z > points_dict[xy_tuple].Z:
                            points_dict[xy_tuple] = vertex

    # Return only the points with the maximum Z value for each unique (X, Y)
    return list(points_dict.values())

# Main function to create a Toposolid from multiple Toposolids
def create_toposolid_from_toposolids():
    # Start Transaction
    with revit.Transaction('Convert Toposolids to Toposolid'):

        try:
            # Pick multiple Toposolids
            toposolids = pick_toposolids()

            if not toposolids:
                raise ValueError("No Toposolids selected.")

            # Get topoTypeId and levelId from the first Toposolid
            first_toposolid = toposolids[0]
            topo_type_id = first_toposolid.GetTypeId()
            level_id = first_toposolid.LevelId

            # Extract and adjust the boundary points
            points = extract_unique_points_from_toposolids(toposolids)

            if len(points) < 3:
                raise ValueError("Not enough unique points to create a Toposolid. A minimum of 3 points is required.")

            # Convert points to an IList[DB.XYZ]
            points_list = List[DB.XYZ](points)

            # Create the Toposolid from the adjusted points
            toposolid = DB.Toposolid.Create(doc, points_list, topo_type_id, level_id)

            if toposolid:
                logger.info("Successfully created Toposolid from multiple Toposolids.")

                # Ask user if they want to delete the original Toposolids
                if forms.alert("Do you want to delete the original Toposolids?", options=["Yes", "No"]) == "Yes":
                    for toposolid in toposolids:
                        doc.Delete(toposolid.Id)
                    logger.info("Deleted original Toposolids.")

            else:
                logger.error("Failed to create Toposolid.")

        except Exception as e:
            logger.error("Failed to convert Toposolids to Toposolid: {}".format(str(e)))

# Run the main function
if __name__ == "__main__":
    create_toposolid_from_toposolids()
