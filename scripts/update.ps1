# This script is called from within a vba code module.
# Upon user prompt the current spec-manager version will be removed
# and the newest release will be downloaded from github

$Shell = New-Object -ComObject ("WScript.Shell")

$ErrorActionPreference = 'Stop'
$tls12 = [Net.ServicePointManager]::SecurityProtocol =  [Enum]::ToObject([Net.SecurityProtocolType], 3072)

# ----------------------------------------------------------------------------------------------------
# UPDATE :
# ----------------------------------------------------------------------------------------------------
# Download latest dotnet/codeformatter release from github
$repo = "codylruff/SpecManager"
$releases = "https://api.github.com/repos/$repo/releases"

[Net.ServicePointManager]::SecurityProtocol = $tls12
$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name

# Initialize variables
$Version = $tag
$SpecManagerDir = "$env:APPDATA\Spec-Manager-$Version"
#$LibsDir = "$SpecManagerDir\libs"
$ConfigDir = "$SpecManagerDir\config"
#$LogsDir = "$SpecManagerDir\logs"
$ZipFile = "$SpecManagerDir\spec-manager.zip"

function SpecManagerShortcut() {
    $ShortCut = $Shell.CreateShortcut("$env:USERPROFILE\Desktop\Spec-Manager.lnk")
    $ShortCut.TargetPath="$SpecManagerDir\scripts\start.bat"
    $ShortCut.Description = "Spec-Manager Launcher";
    $ShortCut.IconLocation="$SpecManagerDir\Spec-Manager.ico"
    $ShortCut.WindowStyle = 7
    $ShortCut.Save()
}

$ReleaseUri = "https://github.com/codylruff/SpecManager/releases/download/$Version/spec-manager-$Version.zip";

if (!(Test-Path $SpecManagerDir)) {
New-Item $SpecManagerDir -ItemType Directory | Out-Null
}
  
[Net.ServicePointManager]::SecurityProtocol = $tls12
Invoke-WebRequest $ReleaseUri -Out $ZipFile

Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
Remove-Item $ZipFile

SpecManagerShortcut

# -----------------------------------------------------------------------------------------------------------
# RESTART : This powershell code will start the application with the updated version.
# -----------------------------------------------------------------------------------------------------------
#$Excel = New-Object -comobject Excel.Application
#$FilePath = "C:\Users\cruff\AppData\Roaming\Spec-Manager-$tag\Spec Manager $tag.xlsm"
#$Excel.Workbooks.Open($FilePath)
#$Excel.visible = $true