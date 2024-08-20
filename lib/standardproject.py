from Autodesk.Revit.DB import BuiltInCategory, View, UnitUtils ,BuiltInParameter
from pyrevit import revit, DB 
from pyrevit.revit import query
import os
from os import path
import csv
from datetime import datetime
import time
import itertools
import json
# create a dictionary to use to diplay and record to mongo db

def containers_info():
    file_path = query.get_central_path(revit.doc)
    model_name = file_path.split("\\")[-1]
    container_code = model_name.split(".")[0]
    return {"file_path":file_path,"model_name":model_name,"container_code":container_code}

def container_name():
    file_path = query.get_central_path(revit.doc)
    model_name = file_path.split("\\")[-1]
    container_code = model_name.split(".")[0]
    return model_name

def number_of_OSTelement(myarg):
    element_collector = DB.FilteredElementCollector(revit.doc) \
                        .OfCategory(myarg) \
                        .WhereElementIsNotElementType().ToElements()
    return len(element_collector)

def view_contain_subStr (listOfView, strContian):
    viewId = []
    viewName = []
    for i in listOfView:
        if strContian in i.Name:
            viewId.append(i.Id)
            viewName.append(i.Name)
    return [viewId,viewName]

# def element_creator_name(elementid):
#     result = []
#     if isinstance(elementid, list):
#         for i in elementid:
#             result.append(DB.WorksharingUtils\
#                             .GetWorksharingTooltipInfo(revit.doc,i)\
#                             .Creator)
#         return result
#     else:
#         return DB\
#                 .WorksharingUtils\
#                 .GetWorksharingTooltipInfo(revit.doc,elementid)\
#                 .Creator
                
                
RE_DATE_REVISION = '(0[1-9]|1[1,2])(\/|-)(0[1-9]|[12][0-9]|3[01])(\/|-|\.)(19|20)\d{2}'


def get_Creator(elementId):
    result = []
    if isinstance(elementId, list):
        for i in elementId:
            result.append(DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc, i).Creator)
    else:
        return DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc, elementId).Creator
    return result
