from datetime import datetime
import os
import os.path
from pyrevit.revit.db import query
from pyrevit.revit import  selection
from pyrevit import  forms
import re 
from Autodesk.Revit.DB import BuiltInCategory, View, UnitUtils ,BuiltInParameter, FilteredElementCollector, ViewSection, ViewPlan, View3D, ViewSheet, WorksharingTooltipInfo

from pyrevit import revit, DB 



RevitLinkType = FilteredElementCollector(revit.doc).OfClass(DB.RevitLinkInstance).ToElements()

def find_duplicates(lst):
    seen = set()
    duplicates = set()
    for item in lst:
        if item in seen:
            duplicates.add(item)
        else:
            seen.add(item)
    return list(duplicates)


name = []

for i in RevitLinkType:
    londname = i.Name
    rvtname = londname.split(".rvt")[0]
    name.append(rvtname)
    


if len(RevitLinkType) == 0:
    forms.alert (msg="No Revit Links in the model")
elif len(name) != len(set(name)):
    forms.alert (msg="There are Duplicated Revit Links in the model\n This are the models duplicated {}".format(find_duplicates(name)))
else:
    forms.alert (msg="There are no Duplicated Revit Links in the model you are good to go!")
    