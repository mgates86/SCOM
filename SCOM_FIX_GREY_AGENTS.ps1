#Variables
$SCOMSERVER = ##ENTERSCOMSERVER##
$SMTPSERVER = ##ENTER-SMTP-SERVER##


#Get the current module path
$p = [Environment]::GetEnvironmentVariable("PSModulePath")

#Add to the existing path the additional path to our module
$p += ";C:\Program Files\Microsoft System Center 2012 R2\Operations Manager\Powershell\"

#Set the new path to the PSModulePath variable
[Environment]::SetEnvironmentVariable("PSModulePath",$p)

Import-Module OperationsManager
New-SCOMManagementGroupConnection $SCOMSERVER



#Find grey agents:
$agent = Get-SCClass -name "Microsoft.SystemCenter.Agent"
$objects = Get-SCOMMonitoringObject -class:$agent | where {$_.IsAvailable -eq $false}

ForEach ($object in $objects){

$computername = $object.DisplayName
$computerdrive = "\\" + $computername
$SCOMPATHOLD = "\C$\Program Files\Microsoft Monitoring Agent\Agent\HealthOLD"
$SCOMPATHNEW = "\c$\Program Files\Microsoft Monitoring Agent\Agent\Health Service State"
$SCOMPATH = "\c$\Program Files\Microsoft Monitoring Agent\Agent\Health Service State"
$SCOMFULLOLD = $computerdrive + $SCOMPATHOLD
$SCOMFULLNEW = $computerdrive + $SCOMPATHNEW
$SCOMFULLPATH = $computerdrive + $SCOMPATH


#Restart grey agents:

#Send-MailMessage -To SCOM.Repository@example.com -Body ($objects | out-string) -From SCOM_repository@example.com -Subject "Repairing: SCOM Agents That are Unmanaged" -SmtpServer $SMTPSERVER

Get-Service -Name healthservice -ComputerName $computername | Stop-Service -Force
Start-Sleep -s 60
Remove-Item  $SCOMFULLPATH -force
#Rename-Item   $SCOMFULLNEW "HealthOLD"
Get-Service -Name healthservice -ComputerName $computername | Start-Service

}

#Wait for SCOM Console to refresh
Start-Sleep -s 240

$objects = Get-SCOMMonitoringObject -class:$agent | where {$_.IsAvailable -eq $false}
Send-MailMessage -To infra@example.com -Body ($objects | out-string) -From SCOM_repository@example.com -Subject "SCOM Agents to Repair: Rename C:\Program Files\Microsoft Monitoring Agent\Agent\Health Service State" -SmtpServer $SMTPSERVER