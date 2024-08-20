from datetime import datetime
import os
import os.path
from pyrevit import revit, DB
from Autodesk.Revit.DB import Revision
import re
from standardproject import RE_DATE_REVISION, get_Creator

revisions_collector = DB.FilteredElementCollector(revit.doc) \
                        .OfClass(Revision)
                        
                        
RevisionDate = [i.RevisionDate for i in revisions_collector]
RevisionDateId = [i.Id for i in revisions_collector]

def message():
    print ("The date format is not correct and it follows the format DD/MM/YY or DD/MM/YYYY") 

for i in revisions_collector:
    pattern = r'^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[012])/((19|20)?\d\d)$'
    if not re.match(pattern, i.RevisionDate):
        print("NO {}: revision format is NOT correct{}".format(get_Creator(i.Id),i.RevisionDate))
    else:
        print("Correct  {}: revision format is correct: {}".format(get_Creator(i.Id),i.RevisionDate))

