from datetime import datetime
import os
import os.path
from pyrevit import revit, DB 
from Autodesk.Revit.DB import BuiltInCategory, View, UnitUtils ,BuiltInParameter


sheet_collector = DB.FilteredElementCollector(revit.doc) \
                        .OfClass(ViewSheet) \
                        .ToElements()
                        
revisions_on_sheet_collection = [i.GetAllRevisionIds() for i in sheet_collector]
# get revision number of drawings 
# get revision                     


print(revisions_on_sheet_collection)