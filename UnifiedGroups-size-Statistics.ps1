#---------------------------------------------------------------------------------------------
#Author: Sunil Chauhan
#Connect with me ## sunilkms@gmail.com
#this report fetches the unified group mailbox size report
#change Below.
# Preload the Connect365.ps1
#---------------------------------------------------------------------------------------------

$ReportPath="E:\Groups"
Start-Transcript -Path "$ReportPath\log\UnifiedGroupReport-$(get-date -Format MM-dd-yyyy).log"
$reportName="UnifiedGroupReport-$(get-date -Format MM-dd-yyyy).csv"
Import-Module "$ReportPath\Connect365.ps1"

$from="sunil@lab365.in"
$to="sunil@lab365.in"
$smtp="smtpserver"

$Report=$ReportPath +"\" + $reportName
#$mailboxReport="$ReportPath\MailboxReport.csv"

ConnectTo365

#Unified Groups
$UnifiedGroup=Get-UnifiedGroup -ResultSize unlimited

$UnifiedMailboxSizeReport=@()
#unified Group Mailbox Statistics

$c=0
foreach ($ug in $UnifiedGroup){
$c++
write-host "Fetching Mailbox stats#[$c]# $($ug.PrimarySmtpAddress)"
$ugs=Get-exoMailboxStatistics -Identity $ug.PrimarySmtpAddress -Properties LastInteractionTime,LastLogonTime

$ugs | select @{N="WhenCreated";E={$ug.WhenCreated}},
@{N="Group";E={$ug.PrimarySmtpAddress}},
@{N="HiddenFromAddressListsEnabled";E={$ug.HiddenFromAddressListsEnabled}},
LastInteractionTime,LastLogonTime,@{N="TotalItemSizeInGB";E={$_.TotalItemSize.value.toGB()}},ItemCount,
@{N="GroupMemberCount";E={$ug.GroupMemberCount}},
@{N="GroupExternalMemberCount";E={$ug.GroupExternalMemberCount}},
@{N="ResourceProvisioningOptions";E={$ug.ResourceProvisioningOptions}} | Export-Csv $Report -Append -NoTypeInformation
}

$data = ipcsv $Report
#send email report
$subject="Unified Group Mailbox Size Report:" + $(get-date -Format s)

$body=@"

<h4> Unified Group Mailbox Size Report </h4>
<ul>
<li>Total Unified Groups              #$($($data).count) </li>
<li>Unified Groups reached 50 GB limit#$($($data | ? {$_.TotalItemSizeInGB -eq 50}).count) </li>
<li>Unified Groups reached 45 GB limit#$($($data | ? {$_.TotalItemSizeInGB -ge 45}).count)</li>
<li>Unified Groups not used in last 90 days#$($($data | ? {[datetime]$_.LastLogonTime -lt (get-date).adddays(-90)}).count)</li>
</ul>

<p>Thanks,</br>
Sunil Chauhan</p>

"@
Send-MailMessage -From $from -To $to -Subject $subject -Attachments $Report -SmtpServer $smtp -Body $body -BodyAsHtml
Stop-Transcript
