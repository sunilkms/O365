##################################################################
# Send As Permission backup report
# Author : Sunil Chauhan
# Email: sunilkms@hotmail.com
# This script takes backup of Send as permission.
##################################################################

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

$LogFile = "SendAs-Report-Logs.log"

Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Starting.."

# Send As Permission Report Code Start from Here

Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Now Getting Send As Permission.."
Write-host "Getting Send As Permission Report.." -f Yellow

# Getting Receipent Permission of all the mailboxes.
$RecipientPermissions = Get-RecipientPermission -ResultSize unlimited

#Formating data
$SendAsRights = $RecipientPermissions | ? {$_.trustee -ne "True" -and $_.Trustee `

-notlike "S-1-5*" -and $_.Trustee -notlike "*Self"} | select Identity,Trustee, Accessrights

# Attachment file Name
$sendAsAttachment = "SendAs-Permission-Backup" + $(get-date -Format MM-dd-yyyy) + ".csv"

#Exporting report Data.
$SendAsRights | Export-csv "$sendAsAttachment" -NoType

Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) : Send As Report is Ready to be sent."

"Sending Email"
$Subject = "Send As Permission Backup for $(get-date -Format MM/dd/yyyy)"
Send-MailMessage -from $mail.From -to $mail.to -subject $Subject -Body $body -SmtpServer $mail.SMTPServer -Attachment $sendAsAttachment
Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Send As Report Was Sent Successfully"
Add-content -path $logFile -value "$(get-date -Format MM/dd/yyyy-hh:mm:ss) :Reporting End."
