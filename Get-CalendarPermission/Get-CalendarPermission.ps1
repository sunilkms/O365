#this script gets users Calendar Permissions in a CSV file from the file "Users.txt"

$users = gc "users.txt"
$RFname="CalenderPermission.csv"
$logfile="UserwithoutMailbox.txt"
$report=@()
foreach ($user in $users) 
      {
        write-host "Checking for"$user
        $mbxStatus = Get-Mailbox $user
        if ($mbxStatus -ne $null) 
            {
             $prem=Get-MailboxFolderPermission $user":\Calendar" | ? {$_.User -notlike "Default" -and $_.user `
             -notlike "anonymous" -and $_.USer -notlike "*S-1-5-21*"}
             $Report+=$Prem
             }
        Else { 
             add-content $Logfile "$user" 
             }
      }
$report | export-csv $rfname -notype
