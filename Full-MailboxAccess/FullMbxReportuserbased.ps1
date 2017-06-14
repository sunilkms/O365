"Getting All Shared Mailboxes..."
#For all the users in org
$Allmailbox = Get-Mailbox -ResultSize unlimited

#for only Shared Mailbox
#$Allmailbox = Get-Mailbox -ResultSize unlimited #-RecipientTypeDetails sharedMailbox 

$MailboxPermissions=@()
foreach($Mailbox in $AllMailbox) {
$prm=Get-MailboxPermission -Identity $Mailbox.alias
$MailboxPermissions+=$prm
}

$FilteredPermissions = $MailboxPermissions | ? {$_.User -Notlike "NT AUTHORITY*" -and `
$_.User -Notlike "S-1-5-21*" -and $_.IsInherited -ne "True"}

#select the required properties in Report
$postfilter = $FilteredPermissions | select Identity,User,AccessRights

"Grouping by User.."
$Grouped = $postfilter | group User
$Report=@()

foreach ($entry in $Grouped){
$Report+=$entry.Group
}

$Report | select User,Identity, {$_.AccessRights} | Export-Csv "FullMailboxReport-Multi.csv" -NoTypeIn
