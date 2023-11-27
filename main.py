from file import create_directory, list_tab_files, unzip_packages, copy_unpackaged
from tableau import process_workbooks

# Create directories
assert create_directory("Unpackaged"), "Create directory failed"
assert create_directory("Fields"), "Create directory failed"

# Get lists of Tableau files
files = list_tab_files("Packaged")
assert files, "No files returned"

# Unzip packaged Workbooks
assert unzip_packages(files["twbx"], "Unpackaged"), "Unzipping failed"

# Copy unpackaged Worbooks
assert copy_unpackaged(files["twb"], "Unpackaged"), "Copying failed"

# Get new list of Unpackaged Workbooks
files = list_tab_files("Unpackaged")
assert files, "No files returned"

# Main Loop
for wb in files["twb"]:
    process_workbooks(wb)
