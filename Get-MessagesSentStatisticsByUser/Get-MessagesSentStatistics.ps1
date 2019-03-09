# Author : Sunil Chauhan
# Blog: https://www.sunilchauhan.info/

# This script gets all the messages sent by a user on a specific date can also include stats for sent to domains.
# The script is fully tested for Exchange Online to the script publising date.

#--------------------------------------------------------------------------------
# How to run Example:
# Get-MessagesSentStatistics -SenderEmailAddress user@domain.com -DaysofActivity 5

# Switch "-GetSentToDomainStats" includes stats for email sent to domains
# Get-MessagesSentStatistics -SenderEmailAddress user@domain.com -DaysofActivity 5 -GetSentToDomainStats
#
# Switch "-ShowProgress" will show progress while script is running.
# Get-MessagesSentStatistics -SenderEmailAddress user@domain.com -DaysofActivity 5 -GetSentToDomainStats -ShowProgress
#-------------------------------------------------------------------------------------------------------------

Function Get-MessagesSentStatistics {
param (
$SenderEmailAddress = "user@domain.com",
$DaysofActivity=7,
[switch]$GetSentToDomainStats=$false,
[switch]$ShowProgress=$false
)

$startDate=(Get-Date).AddDays(- $daysofActivity).ToShortDateString()
$endDate=(Get-Date).ToShortDateString()

#Conver the date to Exchange aware date format dd-MM-yyyy
$startDatetostring=$startDate.Split("-")
$EndDatetostring=$endDate.Split("-")
$startDateforMsgTrace = $startDatetostring[1] + "-" + $startDatetostring[0] + "-" + $startDatetostring[2]
$EndDateforMsgTrace = $EndDatetostring[1] + "-" + $endDatetostring[0] + "-" + $endDatetostring[2]

$AllMesage=@()
$pageNumber=1   

If ($ShowProgress) {Write-Host "Getting sent messages logs for dates between ($startDate) and ($endDate)"}
$Msg=Get-MessageTrace -SenderAddress $SenderEmailAddress -StartDate $startDateforMsgTrace -EndDate $EndDateforMsgTrace -PageSize 5000 -Page $pageNumber

if ($msg.count -gt 4999) {

$AllMesage+=$Msg
$pageNumber++

    Do {
    if ($ShowProgress){
        if ($pageNumber -gt 1) { Write-Host "More messages were sent on $($msg[$msg.count-1].received.ToShortDateString()) " -ForegroundColor Yellow -NoNewline}
        Write-Host "Getting Message Logs from Page Number $pageNumber" 
        }
    $Msg=Get-MessageTrace -SenderAddress $SenderEmailAddress -StartDate $startDateforMsgTrace -EndDate $EndDateforMsgTrace -PageSize 5000 -Page $pageNumber
    $AllMesage+=$msg
    $pageNumber++
    }

    until($($msg[$msg.count-1].received.ToShortDateString()) -le $startDate)    
    $AllMesage | % {$_.Received.ToShortDateString()} | group | select Name, Count
    if ( $GetsenttoDomainStats){"";$AllMesage.RecipientAddress | % {$_.split("@")[1]} | group | select Name, Count }

    } else {
    $msg  | % { $_.Received.ToShortDateString() } | group | select Name, Count
    ""
    if ($GetsenttoDomainStats){$msg.RecipientAddress | % {$_.split("@")[1]} | group | select Name, Count }
    }   
}
