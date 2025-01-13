# pyRevit script amended for HTL
# Author: 'Frederic Beaupere' (pyRevit) 
# Amendments author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-05-14
# Version: 1.0.3
# Description: This script will create a 3D view for each workset and will update existing Workset views. It allows to specify a View Template which can direct  new views to specific Project Browser folders.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

# -*- coding: UTF-8 -*-
"""Lists all linked and imported DWG instances with worksets, creator, and views."""
import clr
from collections import defaultdict

from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms

output = script.get_output()


def listdwgs(current_view_only=False):
    dwgs = DB.FilteredElementCollector(revit.doc)\
             .OfClass(DB.ImportInstance)\
             .WhereElementIsNotElementType()\
             .ToElements()

    dwgInst = defaultdict(list)

    output.print_md("## LINKED AND IMPORTED DWG FILES:")
    output.print_md('By: [{}]({})'.format('Frederic Beaupere',
                                          'https://github.com/frederic-beaupere'))

    for dwg in dwgs:
        if dwg.IsLinked:
            dwgInst["LINKED DWGs:"].append(dwg)
        else:
            dwgInst["IMPORTED DWGs:"].append(dwg)

    for link_mode in dwgInst:
        output.print_md("####{}".format(link_mode))
        for dwg in dwgInst[link_mode]:
            dwg_id = dwg.Id
            dwg_name = dwg.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()
            dwg_workset = revit.query.get_element_workset(dwg).Name
            dwg_instance_creator = DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc, dwg.Id).Creator
            owner_view_id = dwg.OwnerViewId

            if current_view_only and revit.active_view.Id != owner_view_id:
                continue

            # Get the view name where the DWG is placed
            view_name = "3D DWG - not views specific\n\n"
            if dwg.ViewSpecific == True and owner_view_id != DB.ElementId.InvalidElementId:
                owner_view = revit.doc.GetElement(owner_view_id)
                if isinstance(owner_view, DB.View):
                    view_name = "DWG placed in view: ***{}***\n\n".format(owner_view.Name)

            output.print_md("**DWG name:** {}\n\n"
                            "- DWG created by: {}\n\n"
                            "- DWG id: {}\n\n"
                            "- DWG workset: {}\n\n"
                            "- {}"
                            .format(dwg_name,
                                    dwg_instance_creator,
                                    output.linkify(dwg_id),
                                    dwg_workset,
                                    view_name))

selected_option = forms.CommandSwitchWindow.show(
    ['In Current View',
     'In Model'],
    message='Select search option:'
)

if selected_option:
    listdwgs(current_view_only=selected_option == 'In Current View')
