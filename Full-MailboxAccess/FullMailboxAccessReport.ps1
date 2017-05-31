########################################
# Full Mailbox Permission backup
# Author : Sunil Chauhan
# Email: sunilkms@hotmail.com
# This script takes backup of Full Mailbox permission.
#
########################################

#----------------------------------------
# Edit Mail Notification Settings Here
#-----------------------------------------

$mail=@{

  To = "sunilkms@LetsExchange.com"
  From = "sunilkms@LetsExchange.com"
  SMTPServer = "smtp.office365.com"
  
  }

#-------------------------------------------------------------
# you can customize the current msg as per your requirement.
#-------------------------------------------------------------

$body =

"
Hi,


Please Find the Attached Report for $(get-date -Format MM/dd/yyyy).


Thanks

"

#---------------------------------------------------------
# Log
#---------------------------------------------------------

$LogFile = "FullMailbox-SendAs-Report-Logs.log"

Write-host "Getting All Mailboxes in the org.." -f Yellow
Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Starting.."
Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Getting All Mailboxes in the org.."

$AllMailbox = Get-Mailbox -ResultSize unlimited

write-host "Getting FullAccess Permission for All Mailboxes"

Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Checking FullMailbox Permission..."

$data = foreach ($user in $allMailbox)

   {

   Get-MailboxPermission $user.alias | ? {$_.isinherited -ne "True" -and $_.User -notlike `
             "S-1-5*" -and $_.user -notlike "*Self"} | select Identity,@{Label="SMTPAddress";Expression= `
             {$user.PrimarySmtpAddress}},user,AccessRights,@{Label="RecipientType";Expression= `
             {$User.RecipientTypeDetails}}

   }

$FulMbxAcsesPrmAtcmnt = "FullMailboxAccessPermission-" + $(get-date -Format MM-dd-yyyy) + ".csv"

$data | Export-csv $FulMbxAcsesPrmAtcmnt -Notype

Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :FullMailbox Permission reported has been prepared."

"Sending Email"
$Subject = "Full Access Permission Backup for $(get-date -Format MM/dd/yyyy) "
Send-MailMessage -from $mail.From -to $mail.to -subject $Subject -Body $body -SmtpServer $mail.SMTPServer -Attachment $FulMbxAcsesPrmAtcmnt
Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Full Mailbox Report Was Sent."
