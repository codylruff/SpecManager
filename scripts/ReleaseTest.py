import shutil, os, argparse, json
from distutils.dir_util import copy_tree
    
# Takes github release tag as command line arg "vX.Y.Z"

# ****************************************************************************************
# 	1. Create new release folder C:/Users/cruff/source/SM - Final/bin/spec-manager-vX.Y.Z
# 	2. Copy Spec Manager vX.Y.Z into the release folder
# 	3. Copy the config, libs, logs, and scripts folders from the main repo directory into 
#      the release folder
# 	4. Copy the Spec-Manager.ico file in the release directory
# 	5. Update the json files in the config folder of the release directory
# 		a. Local_version.json :
# 			{
# 			  "thisworkbook_version": "1.0",
# 			  "updater_version": "1.0",
# 			  "app_version": "X.Y.Z"
# 			}
# 		b. User.json :
# 			{
# 			  "name": "",
# 			  "default_printer": "",
# 			  "app_version": "X.Y.Z",
# 			  "default_log_level": "",
# 			  "privledge_level": "",
# 			  "product_line": "User",
# 			  "repo_path": ""
# 			}
# 	6. Compress all contents of the release folder into spec-manager-vX.Y.Z.zip
# 	7. Create the new release on github
# 		a. Create the release and give it a tag of vX.Y.Z
# 		b. Fill in the release name and description
# 		c. Add the spec-manager-vX.Y.Z.zip to the release files.
# Finalize the release
# ****************************************************************************************
def get_arguments():
    """Parse the commandline arguments from the user"""

    parser = argparse.ArgumentParser(description='Create a new release of spec-manager.')
    parser.add_argument('-version', help='the version number that makes the tag value used for github release url')

    return parser.parse_args()

def create_release_folder(dirName):
    """Create the folder a new release"""
    
    print("Creating new release directory . . . ")

    # Create target Directory if don't exist
    if not os.path.exists(dirName):
        os.mkdir(dirName)
        print("Directory " , dirName ,  " Created ")
    else:    
        print("Directory " , dirName ,  " already exists")

def copy_files_to_release_folder(repo_dir, release_dir):
    """Copies all the necessary files from the repo to the release directory"""

    print("Copying files to release folder . . . ")

    # copy the icon file for the launcher shortcut
    icon_file = repo_dir + "/Spec-Manager.ico"
    shutil.copy(icon_file, release_dir)

    # repo directories
    repo_config_dir = repo_dir + "/config"
    repo_libs_dir = repo_dir + "/libs"
    repo_logs_dir = repo_dir + "/logs"
    repo_scripts_dir = repo_dir + "/scripts"

    # release directorys
    release_config_dir = release_dir + "/config"
    release_libs_dir = release_dir + "/libs"
    release_logs_dir = release_dir + "/logs"
    release_scripts_dir = release_dir + "/scripts"

    # copy from repo to release
    copy_tree(repo_config_dir, release_config_dir)
    copy_tree(repo_libs_dir, release_libs_dir) 
    copy_tree(repo_logs_dir, release_logs_dir) 
    copy_tree(repo_scripts_dir, release_scripts_dir)

def update_json_files(release_dir, release_ver):
    """Updates the application version in the local_version.json && user.json files"""

    # path to json files
    local_version_json_file = release_dir + "/config/local_version.json"
    user_json_file = release_dir + "/config/user.json"

    # modify the local_version file
    with open(local_version_json_file) as json_file:  
        local_version_json = json.load(json_file)
    local_version_json['app_version'] = release_ver
    with open(local_version_json_file, 'w') as out_file:
        json.dump(local_version_json, out_file)
    
    # modify the user file
    with open(user_json_file) as json_file:
        user_json = json.load(json_file)
    user_json['app_version'] = release_ver
    with open(user_json_file, 'w') as out_file:
        json.dump(user_json, out_file)

def main():

    # Get arguments
    ver = get_arguments().version

    # Create new release directory
    repo_dir = r"C:/Users/cruff/source/SM - Final/release-script-test"
    release_dir = repo_dir + "/bin/spec-manager-v" + ver
    create_release_folder(release_dir)
    
     # copy the new excel file into the release folder
    excel_file = repo_dir + "/Spec Manager v" + ver + ".xlsm"
    shutil.copy(excel_file, release_dir)

    # Copy all other files && folders to the release directory
    copy_files_to_release_folder(repo_dir, release_dir)

    # Update the json files with the new version
    print('Updating json configuration files . . .')
    update_json_files(release_dir, ver)

    # Zip the directory for release
    print('Creating spec manager archive . . . ')
    shutil.make_archive(release_dir, 'zip', release_dir)

if __name__ == "__main__":
    main()