# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-08-15
# Version: 1.0.0
# Description: Shows Central Model GUID number if model is workshared and is based in the cloud or on the Revit Server. It also shows path and size of the Central Model
# Tested with: Revit +2022
# Requirements: pyRevit add-in

import clr
import os
# Import pyRevit modules
from pyrevit import revit, forms, DB

# Get the current document
doc = revit.doc

def get_file_size(file_path):
    try:
        size_bytes = os.path.getsize(file_path)
        size_kb = size_bytes / 1024.0
        size_mb = size_kb / 1024.0
        return "{:.2f} MB".format(size_mb)
    except OSError as error :
        out = "Error: " + str(error)
        return out
output = ""
try:
    guid = doc.WorksharingCentralGUID
    
    if guid:
        output += "This Central Model GUID number is:\n\n"+ str(guid)
    
except Exception as del_err:
    # logger = script.get_logger()
    # logger.error('Error: {}'
    #         .format( del_err))
    output += "This is not a Central Model stored on Revit Server or in the Cloud and it does not have a GUID number."

central_model_path = doc.GetWorksharingCentralModelPath()
if central_model_path:
    central_model_file_path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(central_model_path)
    output += "\n\nThis Central Model file path is:\n\n"+ str(central_model_file_path) + '\n\nModel size: '+ get_file_size(central_model_file_path)

forms.alert(output, title="Central Model Information")
