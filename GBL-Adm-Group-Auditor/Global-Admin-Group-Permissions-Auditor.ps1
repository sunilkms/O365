#------------------------------------------------------------------------
#Author: Sunil Chauhan 
#www.sunilchauhan.info
#About: This script will monitor Global Administrator Group Membership.
#------------------------------------------------------------------------

param(
[switch]$firstRun=$False,
$to="admins@xyz.com",
$from=Auditor@xyz.com,
$smtpServer="smtp.xyz.com"
)

#edit the existing memberlist url.
$existingList="C:\Users\su347719\Desktop\ExistingCAMembers.txt"
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
Write-host "New User:" $addition.InputObject -f Yellow
$subject= "Global Administrator Auditor: A new addition to the group has been detected"
$Body="

A new addition to the group has been detected

User: $($addition.InputObject)

"
Send-MailMessage -Body $body -To $to -From $from -Subject $subject -SmtpServer $smtpServer
Add-Content -Value $members.Emailaddress -Path $existingList
} 

elseif ($members.count -lt $el.count) {
"a member has been removed from the group"
"getting the member details"

$diff = Compare-Object -DifferenceObject $members.EmailAddress -ReferenceObject $el
$subject= "Global Administrator Auditor: Change detected Member has been removed from the group."
$Body="

A Member has been removed from the group.
User:$($diff.InputObject)

"
Send-MailMessage -Body $body -To $to -From $from -Subject $subject -SmtpServer $smtpServer

} else {

"No change were detected in the group"

}
}
