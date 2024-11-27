# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-11-04
# Version: 1.0.0
# Description: Converts a TopoSolid to a TopoSurface.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# Explicit imports for pyRevit scripting and Revit API
from pyrevit import script
from pyrevit import revit, DB, UI

# Initialize logger
logger = script.get_logger()

# Get Revit application and active document
uidoc = revit.uidoc
doc = revit.doc

# Custom selection filter class for selecting only TopoSolids
class ToposolidSelectionFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.Toposolid)

    def AllowReference(self, reference, position):
        return False

# Function to pick a Toposolid using a custom filter
def pick_toposolid():
    try:
        # Prompt user to pick a Toposolid using the custom filter
        selection = uidoc.Selection
        ref = selection.PickObject(UI.Selection.ObjectType.Element, ToposolidSelectionFilter(), "Select a Toposolid")
        element = doc.GetElement(ref)
        return element
    except Exception as e:
        logger.error("Error selecting Toposolid: {}".format(str(e)))
        script.exit()

# Function to extract unique points with the biggest Z value from Toposolid
def extract_unique_points_from_toposolid(toposolid):
    options = DB.Options()
    geometry = toposolid.get_Geometry(options)

    # Dictionary to store points with (X, Y) as keys and max Z value as values
    points_dict = {}

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

    # Yield only the points with the maximum Z value for each unique (X, Y)
    for point in points_dict.values():
        yield point

# Main function to create a TopographySurface from a Toposolid
def create_topography_surface_from_toposolid():
    # Start Transaction
    with revit.Transaction('Convert Toposolid to TopographySurface'):

        try:
            toposolid = pick_toposolid()

            # Extract and adjust the boundary points
            points = list(extract_unique_points_from_toposolid(toposolid))

            if len(points) < 3:
                raise ValueError("Not enough unique points to create a TopographySurface. A minimum of 3 points is required.")

            # Create the TopographySurface from the adjusted points
            topo_surface = DB.Architecture.TopographySurface.Create(doc, points)

            if topo_surface:
                logger.info("Successfully created TopographySurface from Toposolid.")
            else:
                logger.error("Failed to create TopographySurface.")

        except Exception as e:
            logger.error("Failed to convert Toposolid to TopographySurface: {}".format(str(e)))

# Run the main function
if __name__ == "__main__":
    create_topography_surface_from_toposolid()
