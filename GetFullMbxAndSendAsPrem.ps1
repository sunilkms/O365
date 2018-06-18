#
# This script will get users details who has full mailbox and send as permissions 
# I use this as a function for daily usage to fetch the details quickly.
#

#function GetmbxPrem-Cloud {

param ($user)
""
Write-Host "Full Mailbox Access Permissions" -ForegroundColor Yellow
""

Get-MailboxPermission $user | ? {$_.IsInherited -eq $false -and $_.User -notlike "*S-1-5-21*" -and $_.USER -notlike "*ITY\SELF"}| ft -AutoSize

#foreach ($e in $perm) {$e | select @{N="SharedMailbox";E={$e.Identity.NAme}},@{N="USER";e={(get-aduser $E.user.SecurityIdentifier.Value).UserPrincipalName}},AccessRights }

""
Write-Host "Send As Permission" -ForegroundColor Yellow
""
$adprem = Get-RecipientPermission $user | ? {$_.Trustee -notlike "*S-1-5-21*" -and $_.Trustee -notlike "*ITY\SELF"}

$adprem | ft -AutoSize

#$adprem | ? {($_.IsInherited) -like "False"}| fT -AutoSize

#}
