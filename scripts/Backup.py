#!\usr\bin\env python
"""
This script creates a timestamped database backup,
and cleans backups older than a set number of dates

"""    

from __future__ import print_function
from __future__ import unicode_literals
from pathlib import Path

import argparse
import sqlite3
import shutil
import time
import os

#sqlite_file = 'S:\Data Manager\Database\SAATI_Spec_Manager.db3'

DESCRIPTION = """
              Create a timestamped SQLite database backup, and
              clean backups older than a defined number of days
              """
NO_OF_DAYS = 15

def sqlite3_backup(sqlite_file, backup_dir):
    """Create timestamped database copy"""
    
    if not os.path.isdir(backup_dir):
        raise Exception("Backup directory does not exist: {}".format(backup_dir))
    db_file = Path(sqlite_file)
    
    backup_file = Path(backup_dir + "\\" + db_file.stem + time.strftime("-%Y%m%d-%H%M%S") + db_file.suffix)
    
    connection = sqlite3.connect(sqlite_file)
    cursor = connection.cursor()
    
    # Lock database before making a backup
    cursor.execute('begin immediate')
    # Make new backup file
    shutil.copyfile(sqlite_file, backup_file)
    print ("\nCreating {}...".format(backup_file))
    # Unlock database
    connection.rollback()

def clean_data(backup_dir):
	"""Delete files older than NO_OF_DAYS days"""

	print ("\n------------------------------")
	print ("Cleaning up old backups")

	for filename in os.listdir(backup_dir):
	    backup_file = os.path.join(backup_dir, filename)
	    if os.stat(backup_file).st_ctime < (time.time() - NO_OF_DAYS * 86400):
	        if os.path.isfile(backup_file):
	            os.remove(backup_file)
	            print ("Deleting {}...".format(backup_file))

def get_arguments():
    """Parse the commandline arguments from the user"""

    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument('db_file',
                        help='the database file that needs backed up')
    parser.add_argument('backup_dir',
                         help='the directory where the backup'
                              'file should be saved')
    return parser.parse_args()

if __name__ == "__main__":
    #args = get_arguments()
    #sqlite3_backup(args.db_file, args.backup_dir)
    sqlite3_backup(r'c:\users\cruff\source\SM - Final\Database\SAATI_Spec_Manager.db3',
                   r'c:\users\cruff\desktop\Database_Backups')
    #clean_data(r'c:\users\cruff\desktop\Database_Backups')
    print ("\nBackup update has been successful.")