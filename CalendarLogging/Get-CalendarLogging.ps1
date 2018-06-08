# This script will fetch the logs and the run the alalysis on that and then export and ReExport.
# Usage:
# Get-CalendarLogging.ps1 -User "Sunil.Chauhan@domain.com" -ExportReportToPath -MeetingSubject

param(
$User,
$ExportReportToPath = "D:\Sunil\CalLogs\",
$MeetingSubject = "Test Meeting One"
)

#Below CMD will get the Diagnostics Logs and then Run Alalysis on them
$logs = Get-CalendarDiagnosticLog -Identity $USER -Subject $subject; 
# Report will be exported on the path specificed.
$ReportPath = $ExportReportToPath +  $MeetingSubject + ".CSV"
Get-CalendarDiagnosticAnalysis -CalendarLogs $logs -DetailLevel Advanced > $ReportPath

# Export and Reimport to fix the excel formatting issue.

$importedLogs = IPCSV $ReportPath

#Export and Rewrite the exported File.
$importedLogs | Export-CSV $ReportPath -NoType



