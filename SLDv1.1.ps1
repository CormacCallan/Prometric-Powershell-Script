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

#==================Surpass Log File Downloader SLDv1.1========================


#custom note for each file eg. client name
#0000000191053315_examLog.csv
#0000000191053315_ASEcandidateLog.csv
$customNote = 'examLog'


#Enter your confirmation numbers here;
$confirmationArray = (

'8660000001444838'


)



#==================Surpass Log File Downloader SLDv1.1========================




Clear-Host
write-host "`n============= Surpass Log Downloader SLDv1.1 ==================="



#checking for sql module and installing if not present
if(Get-Module -ListAvailable -Name sqlserver){
write-host "`n ☑ Sql Server Module already installed `n `t Continuing with application.."
}
else{
write-host "`n ☐ Installing SQL Server Module.."
install-module sqlserver -Scope CurrentUser
write-host "`n ☑ Continuing with application"
}


function getKeycode ($confNumber) {

    $QueryConf= "
    select distinct top 100
    a.confirmationnumber as 'confirmationNumber' ,        
    aa.deliveryKeycode
    from appointment a with (nolock)
    left join appointmentattribute aa with (nolock) on aa.appointmentid = a.appointmentid
    where
    a.confirmationnumber in 
    ('"+ $confNumber + "')"

    return (Invoke-Sqlcmd -ServerInstance bal-msq-4g -Database ptc_sr_db  -Query $QueryConf).Item('deliveryKeycode')

}



$folderPath = $env:userprofile + '\Desktop\'
$folderName = 'ExamLogs_' + $timeSignature
$timeSignature = Get-Date -Format "HHmm_ss"
New-Item -path $folderPath -Name $folderName -ItemType "directory"  | out-null



function getLogFile ($keycode, $confNumber, $count) {

    $AttachmentPath = $folderPath + $folderName +'\'+ $confNumber + '_' + $customNote  + '.csv'

     $QueryFmt= "
    Use Surpass_SecureAssess
    SELECT [TimeStamp],[ItemID],[ExamSessionCandidateInteractionEventsLookupTable].Description,[Data]
    FROM [dbo].[WAREHOUSE_ExamSessionCandidateInteractionLogsTable]
    INNER JOIN [dbo].ExamSessionCandidateInteractionEventsLookupTable
    ON [dbo].ExamSessionCandidateInteractionEventsLookupTable.EventType = [dbo].[WAREHOUSE_ExamSessionCandidateInteractionLogsTable].EventType
    INNER JOIN WAREHOUSE_ExamSessionTable
    ON WAREHOUSE_ExamSessionTable.ID = [dbo].[WAREHOUSE_ExamSessionCandidateInteractionLogsTable].ExamSessionID
    WHERE KeyCode =
    '"+ $keycode + "'
    order by TimeStamp"


    Write-Host $count - "fetching " $confNumber " -> " $keycode

    Invoke-Sqlcmd -ServerInstance bal-msq-72a -Database Surpass_SecureAssess  -Query $QueryFmt | Export-CSV $AttachmentPath

}




write-host "`n`n`n`n==================/= Fetching Records =\========================`n`n`n`n"

[System.Collections.ArrayList]$logFileObjects = @()


$count = 0

foreach($confirmation in $confirmationArray){
    
    $count++

    $object = [PSCustomObject] @{
    confirmationNumber = $confirmation
    deliveryKeycode = getKeycode -confNumber $confirmation

    }

    $logFileObjects.Add($object) | Out-Null


    #$object.confirmationNumber + " " +  $object.deliveryKeycode

    getLogFile -keycode $object.deliveryKeycode -confNumber $object.confirmationNumber -count $count

}



write-host "`n`n`n`nFiles dumped to " $folderPath$folderName 



write-host "`n=================\= Fetching Completed =/================ C.C ==`n`n"




