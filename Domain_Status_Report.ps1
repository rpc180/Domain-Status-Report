# Header
$startdate = get-date -format F
$shortdate = get-date -format MM-dd-yy
$useracct = get-content env:username
$systemname = hostname
$runpath = get-location
$runlog = $runpath.path+"\System_Status_"+$shortdate+".txt"
$htmllog = $runpath.path+"\System_Status_"+$shortdate+".html"

# Cleanup Old Logs
Get-ChildItem $runpath -File | Where-Object { $_.CreationTime -lt (Get-Date).Adddays(-90) -and $_.extension -like "*html" } `
     | Remove-Item -Force

# Discover Servers
$servers = get-adcomputer -filter { OperatingSystem -like "Windows Server*" } -properties * | Sort-Object -property Name

# Announce and Headers
write-output "Script running on $systemname at $startdate" | tee-object -filepath $runlog
write-output "in $runpath by $useracct" | tee-object -filepath $runlog -append
write-output "Runlog File: $runlog" `r | tee-object -filepath $runlog -append
write-output "Status,Name,IP Address,HD Free Space (GB),State, Service" | out-file -filepath $runlog -append

# Process Systems
foreach ( $system in $servers) {
     $looptimestamp = get-date -format G
     $reportedIP = $NULL
     $diskdata = $NULL
     $diskprops = $NULL
     $reportedHDs = $NULL    
     $monitoredsvc = $NULL
     try
     {
          get-wmiobject win32_operatingsystem -computername $system.dnshostname -erroraction stop | out-null
          $diskdata = get-wmiobject win32_logicaldisk -computername $system.dnshostname -filter "Mediatype=12"
          $monitoredsvc = get-wmiobject win32_service -computername $system.dnshostname -filter `
               { DisplayName LIKE 'VMware Tools%' and startmode = 'auto' }
          foreach ( $disk in $diskdata ) {
               $diskprops = [pscustomobject]@{
                    volume = $disk.name
                    remaining = [math]::round(($disk.freespace/1gb),2)}
               $reportedHDs += "$($diskprops.volume) $($diskprops.remaining) "
               }        
          $reportedIP = get-wmiobject win32_networkadapterconfiguration -filter "ipenabled = 'True'" `
               -computername $system.dnshostname | where-object {$_.ipaddress} | select-object -expand IPaddress `
               | where-object {$_ -notlike "*:*"}
          write-output "$looptimestamp Processing Host: $($system.description)"
          "Online,"+$system.description+","+$reportedIP+","+$reportedHDs+","+$monitoredsvc.state+","+$monitoredsvc.displayname `
               | out-file -filepath $runlog -append
     }
catch [System.Runtime.InteropServices.COMException]
     {
     write-output "$looptimestamp Processing Expected Host: $($system.description), No Response"
     "Offline,"+$system.description | out-file -filepath $runlog -append
     }
}

# Convert CSV outputs to HTML table for email report
$newestfile = Get-ChildItem | Sort-Object -property creationtime -descending | where-object { $_.extension -match ".txt" } `
     | select-object -first 1
$data = get-content $newestfile | where-object{![string]::IsNullOrWhiteSpace($_)}
$tabledata = $data | select-object -skip 3
$style = @"
<style>
    .redtext { color:red; }
    body {font-family: "Helvetica"}
    table{border: 1px solid black; width: 100%; border-collapse: collapse; font-family: "Helvetica"}
    th{border: 1px solid grey; background: #B4D7EC; padding: 5px}
    td{border: 1px solid grey; padding: 5px}
    tr:nth-child(odd) {background-color: #f2f2f2;}
</style>
"@
$tabledata | convertfrom-csv | convertto-html -head $style | set-content $htmllog
$logdata = $data | select-object -first 3
$logdata = $logdata -join "<br />" 
add-content $htmllog "<p>"
convertto-html -body $logdata | add-content $htmllog
add-content $htmllog '<p><hr width="50%"><p>'
add-content $htmllog "<p>"
$colorizetmp = "colorize.temp"
get-content $htmllog | foreach-object { `
     $_.replace("Stopped",'<span class="redtext"> Stopped </span>').replace("Offline",'<span class="redtext"> Offline </span>') `
     } | set-content $colorizetmp
copy-item $colorizetmp $htmllog -force
remove-item $runlog
remove-item $colorizetmp

# Sending Report
[System.Net.ServicePointManager]::SecurityProtocol = 'Tls,TLS11,TLS12'
$cred = import-clixml .\opt\utilsmtpcreds.xml
$newestfile = Get-ChildItem | Sort-Object -property creationtime -descending | where-object { $_.name -match ".html$" } | select-object -first 1
$emailsender = "Service Utility <emailaddress@somewhere.com>"
$emailrecipient = "SysAdmin <emailaddress@somewhere.com>"
$emailsubject = "Systems Status"
$emailbody = get-content $newestfile -raw
$sendmailargs = @{
     SmtpServer = "smtp.gmail.com"
     Credential = $cred
     UseSsl = $true
     Port = 587
     To = "$emailrecipient"
     From = "$emailsender"
     Subject = "$emailsubject"
     Body = $emailbody
}
Send-MailMessage @sendmailargs -BodyAsHTML
