# VBS-Logs-Cleanup-Script

## This script was made for deleting folders with thousands of logs generated by servers.

## **How it works:**
### Basiscally the script gets information from a config file in xml format, from that file the script gets the followinf information:
#### 1. Path to the folder where the files to be deleted are stored. 
#### 2. If subfolders also required cleanup (True or False)
#### 3. Age of file, this the ammount of days old a file needs to be to qualify for deletion. (If value is 5 then files that are 5 days or older get deleted)
#### 4. File extension of the files that need to be deleted (make sure is uppser case)
