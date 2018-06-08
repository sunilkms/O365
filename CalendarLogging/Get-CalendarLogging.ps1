param(
$USER,
$path = "D:\Sunil\CalLogs\$user",
$subject
)

$logs = Get-CalendarDiagnosticLog -Identity $USER -Subject $subject; 
Get-CalendarDiagnosticAnalysis -CalendarLogs $logs -DetailLevel Advanced > $subject
