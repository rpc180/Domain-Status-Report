$startdate = get-date -format F
$shortdate = get-date -format MM-dd-yy
$useracct = get-content env:username
$systemname = hostname
$runpath = get-location
$runlog = $runpath.path+"\Domain_Status_Report_"+$shortdate+".txt"
$htmllog = $runpath.path+ "\Domain_Status_Report_"+$shortdate+".html"
#$ofs = ","


#Discover Servers
$servers = get-adcomputer -filter { OperatingSystem -like "Windows Server*" }


#Announce and Headers
write-output "Script running on $systemname at $startdate" | tee-object -filepath $runlog
write-output "in $runpath by $useracct" | tee-object -filepath $runlog -append
write-output "Runlog File: $runlog" | tee-object -filepath $runlog -append
add-content $runlog "`r"
write-output "Status,AD Name,Reported Name,HW Platform,Reported OS,Reported HDs,Reported IP,State,Services" | out-file -filepath $runlog -append


foreach ( $system in $servers) {
     $looptimestamp = get-date -format G
     $reportedCS = $NULL
     $reportedOS = $NULL
     $diskdata = $NULL
     $diskprops = $NULL
     $reportedHDs = $NULL    
     $reportedIP = $NULL
     $svcqry = $NULL
     try
     {
         get-wmiobject win32_operatingsystem -computername $system.name -erroraction stop | out-null
         $reportedCS = get-wmiobject win32_computersystem -computername $system.name
         $reportedOS = get-wmiobject win32_operatingsystem -computername $system.name
         $svcqry = get-wmiobject win32_service -computername $system.name -filter "DisplayName LIKE 'Remote Desktop Services'"
         $diskdata = get-wmiobject win32_logicaldisk -computername $system.name -filter "Mediatype=12"
         foreach ( $disk in $diskdata ) {
              $diskprops = [pscustomobject]@{
                   volume = $disk.name
                   remaining = [math]::round(($disk.freespace/1gb),2)
                    }
              $reportedHDs += "$($diskprops.volume) $($diskprops.remaining) "
             }        
          $reportedIP = get-wmiobject win32_networkadapterconfiguration -filter "ipenabled = 'True'" -computername $system.name | where-object {$_.ipaddress} | select-object -expand IPaddress | where-object {$_ -notlike "*:*"}
         write-output "$looptimestamp Processing Expected Host: $($system.name), $($reportedCS.name) Responded"
         "Online,"+$system.name+","+$reportedCS.name+","+$reportedCS.model+","+$reportedOS.caption+","+$reportedHDs+","+$reportedIP+","+$svcqry.state +","+$svcqry.displayname | out-file -filepath $runlog -append
     }
catch [System.Runtime.InteropServices.COMException]
     {
     write-output "$looptimestamp Processing Expected Host: $($system.name), No Response"
     "Offline,"+$system.name | out-file -filepath $runlog -append
     }
}


#Convert log to html
$newestfile = dir | sort-object -property creationtime -descending | select-object -first 1
$data = get-content $newestfile | where-object{![string]::IsNullOrWhiteSpace($_)}
$tabledata = $data | select-object -skip 3
$style = @"
<style>
    body {font-family: "Helvetica"}
    table{border: 1px solid black; width: 100%; border-collapse: collapse; font-family: "Helvetica"}
    th{border: 3px solid white; background: #B4D7EC; padding: 5px}
    td{border: 3px solid white; padding: 5px}
    tr:nth-child(odd) {background-color: #f2f2f2;}
</style>
"@
$tabledata | convertfrom-csv | convertto-html -Head $style | set-content $htmllog
$logdata = $data | select-object -first 3
$logdata = $logdata -join "<br />" 
add-content $htmllog "<hr><p>"
convertto-html -body $logdata | add-content $htmllog
remove-item $runlog


#Sending Report
$recipientlist = ""  #FILL IN as "Alias <email@domain.com>,Alias2 <email2@domain.com>"
$smtpserver = ""  #FILL IN
$emailsender = ""  #FILL IN
$emailrecipient = ""  #FILL IN or use $recipientlist
$emailsubject = "Domain Status Report"
$newestfile = get-childitem | sort-object -property creationtime -descending | select-object -first 1
$emailbody = $newestfile

send-mailmessage -to $emailrecipient -from $emailsender -subject $emailsubject -body $emailbody -BodyAsHtml -smtpserver $smtpserver