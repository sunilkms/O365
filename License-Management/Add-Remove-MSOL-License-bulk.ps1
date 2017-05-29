#######################################################
# Script = Remove Lync Only License from Office365 
# Author Sunil Chauhan
# sunilkms@gmail.com
# Blog : Sunil-chauhan.blogspot.com
#######################################################
cls

#$cred = Get-credential
#Connect-MSolService -Credential $cred

#$txtfile = Read-host "Type File Name containing User SAMAccount"
$file = gc "b1.txt"
$logfile = "d:\sunil\b2Logs.txt"
Clear-content -path $logfile
$TenantDomain = "xyz.com"

# Select Service Plan
#=====================================
#YAMMER_ENTERPRISE
#RMS_S_ENTERPRISE
#OFFICESUBSCRIPTION
#MCOSTANDARD
#SHAREPOINTWAC
#SHAREPOINTENTERPRISE
#EXCHANGE_S_ENTERPRISE

$skuEnt = Get-MsolAccountSku | where {$_.SkuPartNumber -eq "ENTERPRISEPACK"}
$skuSTD = Get-MsolAccountSku | where {$_.SkuPartNumber -eq "STANDARDPACK"}

$EntSrv = New-MsolLicenseOptions -AccountSkuId $skuEnt.accountskuid -DisabledPlans RMS_S_ENTERPRISE, SHAREPOINTWAC, SHAREPOINTENTERPRISE, MCOSTANDARD, EXCHANGE_S_ENTERPRISE
$StdSrv = New-MsolLicenseOptions -AccountSkuId $SkuStd.AccountSkuId -DisabledPlans MCOSTANDARD, SHAREPOINTSTANDARD,EXCHANGE_S_STANDARD

#$file = Get-content $txtfile

$e=0
$s=0
$n=0
$r=0

$std ="motorolasolutions:STANDARDPACK"
$ent ="motorolasolutions:ENTERPRISEPACK"

foreach ($user in $file)

	{
		$r++
		$upn = $user + "@" + $Tenantdomain
		Write-host "Processing`tRecordNo:$r`tUser:"$upn
		
		#check License Type Assigned to User
		$cl = (Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid
			
		if ($cl -eq $std) {$st = $true} else {$st = $false}
		if ($cl -eq $ent) {$en = $true} else {$en = $false}
		
		$sku=(Get-MsolUser -UserPrincipalName $upn).licenses.accountsku.skupartnumber
		
		if ($en -eq $true) {
		
					$e++
					Write-host "
  Current Licence assigned:" $sku -f Yellow		
					#Write-Host "Running ent block.."
					Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $EntSrv 
					Write-Host (Get-Date) "`tSuccess $User@$TenantDomain`tLicense Modification Successful" -f Green
					$value = (Get-Date).tostring() + "`tSuccess :$User `tType :$sku"
					Add-content -path $logfile -Value $value				
		}
		
		if ($st -eq $true)
		
		{
					$s++
					Write-host "
Current Licence assigned:" $sku -f Yellow		
					#Write-Host "Running standard block"
					Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $stdSrv
					#Remove all licence
					#Set-MsolUserLicense -UserPrincipalName $upn  -RemoveLicenses $skuSTD.accountskuid
					Write-Host (Get-Date) "`tSuccess $User@$TenantDomain`tLicense Modification Successful" -f Green
					$value = (Get-Date).tostring() + "`tSuccess :$User `tType :$sku"
					Add-content -path $logfile -Value $value
				}
				
				if ($st -eq $false -and $en -eq $false) {
				$n++
				$value = (Get-Date).tostring() + "`tSuccess :$User `tType :Null"
					Add-content -path $logfile -Value $value				
				   }				
				}
	
	$wv = "Total Enterprice Users Lync Disabled:$e"
	Write-host $wv -f Green
	Add-content -path $logfile -Value $wv
	$wv = "Totel Standard Users Lync Disabled:$S"
	Write-host $wv -f Green
	Add-content -path $logfile -Value $wv
	$wv = "Total Users did not have License:$n"
	Write-host $wv -f Green
	Add-content -path $logfile -Value $wv
