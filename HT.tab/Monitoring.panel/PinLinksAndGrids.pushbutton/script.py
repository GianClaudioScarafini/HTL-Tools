# -*- coding: UTF-8 -*-
# Author: Gian Claudio Scarafini
# Date: 2024-09-10
# Version: 1.0.0
# Description: this script pins all the grids and levels in the project
# Tested with: Revit 2022/2023/2024
# Requirements: pyRevit add-in

from pyrevit import revit, DB
from collections import Counter 


## Collector Grids
gridCollection = DB.FilteredElementCollector(revit.doc).OfClass(DB.Grid)

## Collectore of Levels
levelCollection = DB.FilteredElementCollector(revit.doc).OfClass(DB.Level)

t = DB.Transaction(revit.doc, 'Pin Links and Grids')

t.Start()
## Pin the links and grids

def pinElements(elements):
    for i in elements:
        i.Pinned = True
        
        

pinElements(gridCollection)
pinElements(levelCollection)

#finish the transaction
t.Commit()