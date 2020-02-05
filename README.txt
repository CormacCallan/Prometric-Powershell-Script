<#
Surpass Log File Downloader SLDv1.1

PLEASE OPEN USING POWERSHELL ISE  

1.1 version allows;
Use of the confirmation number instead of keycode to search
Each file now saves as confiramtion number.csv rather than keycode.csv


#Logic
connect to ptc_sr_db get keycode for passed confirmation
use returned keycode to search secure_assess for log
use sql dump to save file to folder on desktop 
repeat for each confirmation number passed in


#Notes
If there is no log the exam may not have happened this will create empty file
Look to the invoke sql cmd to change servers for EU / Japan Logs

DATE: 02/07/2019
AUTH: Cormac Callan

#>