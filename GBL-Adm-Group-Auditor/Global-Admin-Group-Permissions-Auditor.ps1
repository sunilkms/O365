#------------------------------------------------------------------------
#Author: Sunil Chauhan 
#www.sunilchauhan.info
#About: This script will monitor Global Administrator Group Membership.
#------------------------------------------------------------------------

param(
[switch]$firstRun=$False,
$to="Sunil.chauhan@xyz.com",
$from="Sunil.chauhan@xyz.com",
$smtpServer="smtp.office365.com"
)

$existingList="C:\Users\sunil\Desktop\ExistingCAMembers.txt"
$EL=gc $existingList

$CA=Get-MsolRole -RoleName  "Company Administrator"
$members=Get-MsolRoleMember -RoleObjectId  $ca.ObjectId

if ($firstRun) {
"Script is running for the first time, Saving the existing member in the group"
Add-Content -Value $members.Emailaddress -Path $existingList
} else {

if ($members.count -gt $el.count) {

"A new entry has been added"
"Getting the newly added entry"

$addition = Compare-Object -DifferenceObject $members.EmailAddress -ReferenceObject $el
Write-host "New User:" $($addition.InputObject -join "; ")-f Yellow
$subject= "Global Administrator Group Membership Auditor::Change Detected-Member Added."
$Body="

A new addition to the group has been detected

User: $($addition.InputObject -join "; ")

"
Send-MailMessage -Body $body -To $to -From $from -Subject $subject -SmtpServer $smtpServe
Clear-Content -Path $existingList
Add-Content -Value $members.Emailaddress -Path $existingList
} 

elseif ($members.count -lt $el.count) {
"A member has been removed from the group."

"getting the member details"

$diff = Compare-Object -DifferenceObject $members.EmailAddress -ReferenceObject $el
Write-host "User Removed:" $($Diff.InputObject -join "; ")-f Yellow
$subject= "Global Administrator Group Membership Auditor::Change Detected-Member Removed"
$Body="

A Member has been removed from the group.

User:$($diff.InputObject -join "; ")

"
Send-MailMessage -Body $body -To $to -From $from -Subject $subject -SmtpServer $smtpServer
Clear-Content -Path $existingList
Add-Content -Value $members.Emailaddress -Path $existingList
} else {
"No change were detected in the group"
 }
}
