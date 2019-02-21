#get users created in Last 24 hours UTC 22 Hours - Finland 00 hours

#notification settings
$to="sunil@sunilchauhan.info"
$from=$cred.UserName
$smtp ="smtp.outlook.com"
$subject="Daily Report:New Users created in Last 24 hours"

## Date Settings
$date = (get-date).AddDays(-2)
$day=$date.Day
$year=$date.Year
$month=$date.Month

#Report File Name
$reportName="Users_Created_in_Last_24hours" + "_"+$year+"_"+$month+"_"+$day+".csv"
$attchment=$reportName

#featching report
Write-Host "Fatching users created in last 24 hours"
$QuaryDate= "'"+"$year"+"-"+"$month"+"-"+"$day"+"T22:00"+"'"
$newusers=Get-user -Filter "WhenCreated -gt $QuaryDate" | select WhenCreatedUTC,DisplayName,WindowsEmailAddress,Title,RecipientTypeDetails,IsDirSynced
$newusers | Export-Csv $reportName -NoTypeInformation

$TotalMailboxCreated=(($newusers | ? {$_.RecipientTypeDetails -eq "UserMailbox"}).count)
$TotalMailuserCreated=(($newusers | ? {$_.RecipientTypeDetails -eq "MailUser"}).count)
$TotalGuestUserCreated=(($newusers | ? {$_.RecipientTypeDetails -eq "GuestMailUser"}).count)
$TotalTypesharedCreated=(($newusers | ? {$_.RecipientTypeDetails -eq "SharedMailbox"}).count)
$TotalTypeUserCreated=(($newusers | ? {$_.RecipientTypeDetails -eq "User"}).count)

$TotalUsersCreated=($newusers.count)
$CloudOnly = $newusers | ? {$_.IsDirSynced -eq $false}
$st = $CloudOnly | ? {$_.RecipientTypeDetails -ne 'GuestMailUser'} | select When*,*name,RecipientTypeDetails,IsDirSynced

$Conly = if ($st) {$($st | ConvertTo-HTML -Fragment -PreContent "<h3 style='background-color:#000000;color:#F8F8FF;'>Cloud Only users</h3>" | Out-String) }

$body = @"
<style>
table {
  font-family: arial, sans-serif;
  border-collapse: collapse;
  width: 100%;
}
td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 6px;
}
tr:nth-child(even) {
  background-color: #dddddd;
}
</style>

<body>
<h3 style='background-color:#000000;color:#F8F8FF;'>Report Fetched Since:$QuaryDate UTC Time</h3>
<table>
  <tr>
    <td>Type UserMailbox Created</td>
    <td>$TotalMailboxCreated</td>
  </tr>
  <tr>
    <td>Type Mailuser Created</td>
    <td>$TotalMailuserCreated</td>
  </tr>
  <tr>
    <td>Type GuestUser Created - (External users added to Teams or O365 Group)</td>
    <td>$TotalMailboxCreated</td>
  </tr>
  <tr>
    <td>Type User Created</td>
    <td>$TotalTypeUserCreated</td>
  </tr>
  <tr>
    <td>Type Shared Created</td>
    <td>$TotalTypesharedCreated</td>
  </tr>
  <tr>
    <td>Total Number of Users Created</td>
    <td>$TotalUsersCreated</td>
  </tr>
</table>

$Conly

</br>
<p style='background-color:#DAF7A6;color:#191970;padding:20px;text-align:left;width:80%;'>
<i>Report prepared By: Sunil Chauhan</i>
</p>

"@

#send Email
Write-Host "Sending Email..."
Send-MailMessage -To $to -From $from -SmtpServer $smtp -Subject $subject -Body $body -BodyAsHtml `
-Port 587 -UseSsl -Credential $credapp -Attachments $attchment
"Done"
