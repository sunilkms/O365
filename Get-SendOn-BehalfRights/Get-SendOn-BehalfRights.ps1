$AllMailbox = Get-mailbox -resultsize unlimited

$Report=@()

write-host "Formating sob permission in required report format..."
foreach ($mbx in $allMailbox)

 {
  if (!$mbx.GrantSendOnBehalfto)
   {
    write-host "No One has Send On Behalf Rights on:" $mbx.alias
    $Reportdata = New-Object –TypeName PSObject    
    $reportData | Add-Member –MemberType NoteProperty –Name IDENTITY –Value $mbx.Alias;
    $reportData | Add-Member –MemberType NoteProperty –Name User –Value "Null";
    $reportData
    
    $Report += $reportdata
   }
  Else
   { 
    
    Write-host "SOB permission enery found" -f Green
    foreach ($entry in $mbx.Grantsendonbehalfto)
    { 
     $Reportdata = New-Object –TypeName PSObject
     $reportData | Add-Member –MemberType NoteProperty –Name IDENTITY –Value $mbx.Alias;
     $reportData | Add-Member –MemberType NoteProperty –Name User -Value $entry;
     $reportData
     $Report += $reportdata
    }
   }
 }
 $ReportName = "SendonBehalfto-Backup-" + $(get-date -Format MM-dd-yyyy) + ".Csv"
 $report | Export-csv $ReportName -notype
