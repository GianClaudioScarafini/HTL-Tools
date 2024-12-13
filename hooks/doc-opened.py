# #! python3

# import os
# import subprocess
# import re
# import datetime
# import csv

# PATH_TO_SAVE = r"Y:\9000 Office Admin\08-BIM\__HTL Revit Resources\Computer Resources"

# timeestamp = datetime.datetime.now().strftime("%Y-%m-%d")


# def extract_number(text):
#     # Use regular expression to extract the numeric part
#     match = re.search(r'(\d+(\.\d+)?)', text)
#     if match:
#         return float(match.group(0))  # Convert the matched string to a float
#     else:
#         return None  # Return None if no number is found

# def get_pc_name():
#     command = ["hostname"]
#     result = subprocess.run(command, capture_output=True, text=True, check=True)
#     return str(result.stdout).strip()

# def filename_csv():
#     timeestamp = datetime.datetime.now().strftime("%Y-%m-%d")
#     pc_name = get_pc_name()
#     return "{0}_{1}.csv".format(pc_name, timeestamp)

# def get_revit_version_installed():
#     #DONE
#     command = [ "pyrevit",  "env"]
#     result = subprocess.run(command, capture_output=True, text=True, check=True)
#     pattern = r"Version: [0-9]{2}\.[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2} "
#     matches = re.findall(pattern, result.stdout)
#     output = []
#     for match in matches:
#         output.append(["Revit 20{}".format(match[9:11]), match.replace("Version: ", "")])
#     # return output
    
#     unique_uotput = list(set(tuple(item) for item in output))
#     unique_versions_list = [list(item) for item in unique_uotput]
#     return unique_versions_list


# def get_disk_free_space(drive):
#     #DONE
#     command = ["fsutil", "volume", "diskfree", drive]
#     result = subprocess.run(command, capture_output=True, text=True, check=True)
#     value = result
#     pattern = r"\(([^()]+)\)"
#     matches = list(re.findall(pattern,value.stdout))
#     return [["Total Space free" , extract_number(matches[0])],\
#             [ "Total space" , extract_number(matches[1])]]


# def get_nvidia_drive():
#     #done
#     command = ["nvidia-smi", "--query-gpu=driver_version", "--format=csv,noheader"]
#     result = subprocess.run(command, capture_output=True, text=True, check=True)
#     return [["Driver Version",result.stdout.strip()]]

# def get_nvidia_gpu_name():
#     #done 
#     # nvidia-smi --query-gpu=name --format=csv,noheader

#     command = ["nvidia-smi", "--query-gpu=name", "--format=csv,noheader"]
#     result = subprocess.run(command, capture_output=True, text=True, check=True)
#     return [["GPU Name",result.stdout.strip()]]
 

# data =  [["date",timeestamp]] +\
#         [["PC Name", get_pc_name()]] +\
#         get_disk_free_space("C:\\") +\
#         get_revit_version_installed() +\
#         get_nvidia_gpu_name() +\
#         get_nvidia_drive()


# if filename_csv() in os.listdir(PATH_TO_SAVE):
#     pass
# else:
#     with open(os.path.join(PATH_TO_SAVE, filename_csv()), "w", newline="") as file:
#         writer = csv.writer(file)
#         writer.writerows(data)