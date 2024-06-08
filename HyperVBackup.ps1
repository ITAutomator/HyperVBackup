###
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
###
### Main 
Param ## provide a comma separated list of switches
	(
	[switch] $quiet
	)

### Main function header - Put RethinkitFunctions.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$scriptVer      = "v"+(Get-Item $scriptFullname).LastWriteTime.ToString("yyyy-MM-dd")
if ((Test-Path("$scriptDir\ITAutomator.psm1"))) {Import-Module "$scriptDir\ITAutomator.psm1" -Force} else {write-host "Err: Couldn't find RethinkitFunctions.psm1";return}
# Get-Command -module ITAutomator  ##Shows a list of available functions
######################

#######################
## Main Procedure Start
#######################

########### Load From XML
$Globals=@{}
$Globals=GlobalsLoad $Globals $scriptXML $false
$GlobalsChng=$false
if (-not $Globals.BackupsFolder)   {$GlobalsChng=$true;$Globals.Add("BackupsFolder","D:\HyperVBackups")}
if (-not $Globals.BackupsToKeep)     {$GlobalsChng=$true;$Globals.Add("BackupsToKeep",4)}
if (-not $Globals.VM_Excludes)       {$GlobalsChng=$true;$Globals.Add("VM_Excludes",("VM to Exclude 1","VM to Exclude 2"))}
####
if ($GlobalsChng) {GlobalsSave $Globals $scriptXML;"A new settings file was created (edit as needed and run again): $(Split-Path $scriptXML -Leaf)";Start-Sleep 3;Exit}
########### Load From XML

$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
##
$date = Get-Date -format "yyyy-MM-dd_HH-mm"
New-Item -ItemType Directory -Force -Path $Globals.BackupsFolder -ErrorAction Ignore | out-null
if (-not (Test-Path $Globals.BackupsFolder)) {
    $ErrOut=102; Write-Host "Err $ErrOut : Couldn't create backups folder (edit xml to fix): $($Globals.BackupsFolder)";Start-Sleep -Seconds 3; Exit($ErrOut)
}

$old_dirs = @(Get-ChildItem "$($Globals.BackupsFolder)\Backup-*" -Directory | Sort-Object CreationTime -Descending | Select-Object -Skip ($Globals.BackupsToKeep-1))
Write-Host "-----------------------------------------------------------------------------"
Write-Host "$($scriptName) $($scriptVer)       Computer:$($env:computername) User:$($env:username) PSver:$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
if ($quiet) {Write-Host ("<<-quiet>>")}
Write-Host ""
Write-Host "Creates up to $($Globals.BackupsToKeep) live backups of the HyperV VMs on this machine (Removing the oldest)."
Write-Host ""
Write-Host "Get-VM | Export-VM"
Write-Host "VM Excludes: " -NoNewline
Write-Host ($Globals.VM_Excludes -join ", ")
Write-Host ""
Write-Host "Will create:"
Write-Host "Backup-$($date)"
Write-Host ""
Write-Host "Will delete:"
ForEach ($old_dir in $old_dirs)
{
    $sizemb    = (Get-ChildItem $old_dir -Recurse -File | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB
    $sizembtxt = BytestoString  ($sizemb * 1024)
    Write-Host "$($old_dir.name) ($($sizembtxt))"
}
Write-Host ""
Write-Host "Will backup:"
##
$VMs = Get-VM
##
$count=0
[long] $bytes_total = 0
ForEach ($VM in $VMs)
{
    $count+=1
    #$bytes = $VM.SizeOfSystemFiles*[math]::Pow(1024,2)
    #$bytes = $VM.HardDrives|Get-VHD | Select-Object -ExpandProperty Size | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $bytes = $VM.HardDrives|Get-VHD | Select-Object -ExpandProperty Filesize | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    if ($VM.Name -in $Globals.VM_Excludes)
    {$excl = "     [EXCLUDED]"}
    else
    {$excl = "";$bytes_total += $bytes}
    Write-Host "$($count): $($VM.Name) [$(BytestoString $bytes)]$($excl)"
}
Write-Host "Total Size: $($(BytestoString $bytes_total))"
Write-Host "-----------------------------------------------------------------------------"
If (-not(IsAdmin))
    {
    $ErrOut=101; Write-Host "Err $ErrOut : This script requires Administrator priviledges, re-run with elevation (right-click and Run as Admin)";Start-Sleep -Seconds 3; Exit($ErrOut)
    }
if (Test-Path "$($Globals.BackupsFolder)\Backup-$($date)")
    {
    $ErrOut=102; Write-Host "Err $ErrOut : The folder already exists.";Start-Sleep -Seconds 3; Exit($ErrOut)
    }

if ($quiet) {PauseTimed -quiet} else {PauseTimed}

Write-Host "Removing Old Dirs..."
ForEach ($old_dir in $old_dirs)
{
    $sizemb    = (Get-ChildItem $old_dir -Recurse -File | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB
    $sizembtxt = BytestoString  ($sizemb * 1024)
    $old_dir | Remove-Item -Recurse -Force
    Write-Host "$($old_dir.name) ($($sizembtxt))"
}
Write-Host ""
Write-Host "Creating Backup..."
Write-Host "Backup-$($date)"
#### Backup
#Get-VM | Export-VM -Path "$($Globals.BackupsFolder)\Backup-$($date)"
$count=0
[long] $bytes_total = 0
ForEach ($VM in $VMs)
{
    $count+=1
    #$bytes = $VM.SizeOfSystemFiles*[math]::Pow(1024,2)
    #$bytes = $VM.HardDrives|Get-VHD | Select-Object -ExpandProperty Size | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $bytes = $VM.HardDrives|Get-VHD | Select-Object -ExpandProperty Filesize | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    if ($VM.Name -in $Globals.VM_Excludes)
    {$excl = "     [EXCLUDED]"}
    else
    {$excl = "";$bytes_total += $bytes}
    Write-Host "$($count): $($VM.Name) [$(BytestoString $bytes)]$($excl)"
    if ($excl -eq "")
    {
        $VM | Export-VM -Path "$($Globals.BackupsFolder)\Backup-$($date)"
    }
}
Write-Host "Total Size: $($(BytestoString $bytes_total))"

#### Debug (just create an empty folder
# New-Item -ItemType Directory -Path "$($Globals.BackupsFolder)\Backup-$($date)" | Out-Null
####

### Maybe send to recycle bin one day
#Add-Type -AssemblyName Microsoft.VisualBasic
#[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory('d:\foo.txt','OnlyErrorDialogs','SendToRecycleBin')

## Main Procedure End
#######################
Write-Host "-----------------------------------------------------------------------------"
Write-Host "Done"
#################### Transcript Save
Stop-Transcript | Out-Null
$date = get-date -format "yyyy-MM-dd_HH-mm"
$TranscriptTarget = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+"_log.txt"
If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
Move-Item $Transcript $TranscriptTarget -Force
#################### Transcript Save
if ($quiet) {PauseTimed -quiet} else {PauseTimed}
Return