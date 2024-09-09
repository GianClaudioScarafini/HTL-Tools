# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-08-19
# Version: 1.1.0
# Description: 
# This pyRevit startup script (hook) records synchronization events into separate CSV files for each unique Revit document. 
# More information here: 
# https://pyrevitlabs.notion.site/Extension-Bundles-10df16fe826040bc9dbd9e83bb4248e6
# Tested with: Revit 2022+
# Requirements: pyRevit add-in

import clr
import csv
import os
import time
from datetime import datetime
from pyrevit import DB, framework, HOST_APP


clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox

# Define the directory path for the CSV logs
base_csv_directory = r'Y:\9000 Office Admin\08-BIM\__HTL Revit Resources\Sync Times'

# Ensure the directory exists
if not os.path.exists(base_csv_directory):
    os.makedirs(base_csv_directory)

def get_size_MB(file_path):
        size_bytes = os.path.getsize(file_path)
        size_kb = size_bytes / 1024.0 # With no .0 integer is returned
        size_mb = size_kb / 1024.0
        # MessageBox.Show("{:.2f} MB".format(size_mb), "pyRevit Hook", 0)
        return "{:.2f}".format(size_mb) # Size in MB

def get_file_size(doc, file_path):
    try:
        return get_size_MB(file_path)
    except:
        try:
            # Get the current Windows username
            import getpass 
            username = getpass.getuser()
            # MessageBox.Show(username, "pyRevit Hook", 0)
            
            # Get the Revit version
            app = HOST_APP.app
            revit_version = app.VersionNumber
            # MessageBox.Show(revit_version, "pyRevit Hook", 0)
            
            # Get the user account ID
            user_account_id = app.LoginUserId # folder id

            # Get the project GUID and model GUID
            project_guid = doc.GetCloudModelPath().GetProjectGUID()
            model_guid = doc.WorksharingCentralGUID
            # MessageBox.Show(revit.doc.PathName, "pyRevit Hook", 0)

            # Construct the local file path
            file_path = "C:\\Users\\{}\\AppData\\Local\\Autodesk\\Revit\\Autodesk Revit {}\\CollaborationCache\\{}\\{}\\{}.rvt".format(username, revit_version, user_account_id, project_guid, model_guid)

            # Currently there is no simpler way to get a cloud based model local file path 
            # https://forums.autodesk.com/t5/revit-api-forum/local-copy-file-path-of-cloud-models-on-document-opening-event/td-p/12925825

            return get_size_MB(file_path)
        except:
            return 0
    
def central_model_path(doc):
    path = doc.GetWorksharingCentralModelPath()
    
    # Convert Model Path to a user-visible file path
    if path:
        file_path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(path)
        # MessageBox.Show(file_path, "pyRevit Hook", 0)
        return file_path
    else:
        # MessageBox.Show('Return 2', "pyRevit Hook", 0)
        return "No central model path found."

def get_document_file_name(doc):
    """Generates the file name for the CSV log based on the document."""
    central_model_file_path = central_model_path(doc)
    if central_model_file_path != "No central model path found.":
        # MessageBox.Show('OPT 1', "pyRevit Hook", 0)
        # Extract the file name from the path
        central_model_file_name = os.path.basename(central_model_file_path)
        root_ext = os.path.splitext(central_model_file_name)
        file_name = root_ext[0]
    try:
        # MessageBox.Show('Try 1', "pyRevit Hook", 0)
        central_guid = doc.WorksharingCentralGUID
        file_name = file_name + "_" + central_guid + ".csv"
    except:
        file_name += ".csv"
        # MessageBox.Show('Try 2', "pyRevit Hook", 0)
    
    return os.path.join(base_csv_directory, file_name)

def sync_start_event_handler(sender, args):
    """Handles the start of the synchronization event."""
    # MessageBox.Show("App Synchronization event started.", "pyRevit Hook", 0)
    global sync_start_time
    sync_start_time = datetime.now()
    # MessageBox.Show(sync_start_time.strftime("%Y/%m/%d %H:%M:%S"), "pyRevit Hook", 0)

def sync_end_event_handler(sender, args):
    """Handles the end of the synchronization event."""
    global sync_start_time
    # MessageBox.Show(sync_start_time.strftime("%Y/%m/%d %H:%M:%S"),"pyRevit Hook", 0)
    if sync_start_time:
        sync_end_time = datetime.now()
        # MessageBox.Show(sync_end_time.strftime("%Y/%m/%d %H:%M:%S"),"pyRevit Hook", 0)
        log_sync_event(args.Document, sync_start_time, sync_end_time)
    # MessageBox.Show("App Synchronization event ended.", "pyRevit Hook", 0)

def log_sync_event(doc, start_time, end_time):
    """Logs a synchronization event to the corresponding CSV file."""
    pc_name = os.environ.get("COMPUTERNAME", "Unknown PC")
    # MessageBox.Show(pc_name, "pyRevit Hook", 0)
    user_name = os.environ.get("USERNAME", "Unknown User")
    # MessageBox.Show(user_name, "pyRevit Hook", 0)
    duration = end_time - start_time
    duration = duration.total_seconds()
    duration = "{:.2f}".format(duration)
    # MessageBox.Show( str(duration), "pyRevit Hook", 0)
    csv_file_path = get_document_file_name(doc)
    # MessageBox.Show(csv_file_path, "pyRevit Hook", 0)
    file_size = get_file_size(doc, central_model_path(doc))
    # MessageBox.Show(file_size, "pyRevit Hook", 0)
    file_exists = os.path.isfile(csv_file_path)

    retries=3 # number of attempts to retry to save the file
    delay=2   # in seconds
    attempt = 0
    while attempt < retries:
        try:
            with open(csv_file_path, mode='ab') as csvfile: # Adding newline='' was creating an error
                fieldnames = ['PC Name', 'User Name', 'File Size', 'Sync Start Time', 'Sync End Time', 'Duration']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

                if not file_exists:
                    writer.writeheader()
                
                writer.writerow({
                    'PC Name': pc_name,
                    'User Name': user_name,
                    'File Size': file_size,
                    'Sync Start Time': start_time.strftime("%Y/%m/%d %H:%M:%S"),
                    'Sync End Time': end_time.strftime("%Y/%m/%d %H:%M:%S"),
                    'Duration': duration
                })
                
                # MessageBox.Show("File written successfully.", "pyRevit Hook", 0)
            break  # Exit the loop if the operation is successful
        except Exception as e:
            attempt += 1
            if attempt < retries:
                time.sleep(delay)  # Wait for the specified delay before retrying
            else:
                MessageBox.Show("Error writing to CSV Synchronization file: {}. Inform BIM Manager.".format(e), "pyRevit Hook", 0)

# Registers event handlers for Revit synchronization events
app = HOST_APP.app
app.DocumentSynchronizingWithCentral += framework.EventHandler[DB.Events.DocumentSynchronizingWithCentralEventArgs](sync_start_event_handler)
app.DocumentSynchronizedWithCentral += framework.EventHandler[DB.Events.DocumentSynchronizedWithCentralEventArgs](sync_end_event_handler)

# Unregistering Revit synchronization events is unnecessary as the events should work for the whole duration of the Revit session.