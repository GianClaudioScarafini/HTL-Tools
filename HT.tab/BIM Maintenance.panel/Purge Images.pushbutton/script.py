# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-08-05
# Version: 1.0.1
# Description: Removes all images from the model
# Tested with: Revit +2022
# Requirements: pyRevit add-in

from pyrevit import forms, revit, DB

# Get the current Revit document
doc = revit.doc

# Input from a user
purge = forms.alert(
    'Purge all images?',
    title="Purge all images?",
    cancel=True,
    ok=False,
    yes=True,
    no=True,
    exitscript=True
)

if purge:
    # Initialize a list to store the names of deleted images
    deleted_images = []
    # Start a transaction
    with revit.Transaction('Purge all images'):

        # Collect all ImageType elements in the document
        image_types = DB.FilteredElementCollector(doc).OfClass(DB.ImageType).ToElements()

        # Iterate over each ImageType and remove it
        for image_type in image_types:
            
            # Get the name of the image file from the Element class
            image_name = image_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            # Store the name of the image file
            deleted_images.append(image_name)

            doc.Delete(image_type.Id)

    # Print the list of deleted images
    print("Following files were deleted:\n")
    for image_name in deleted_images:
        print(image_name + "\n")
