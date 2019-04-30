# This script is called from within a vba code module.
# Upon user prompt the current spec-manager version will be removed
# and the newest release will be downloaded from github

$Shell = New-Object -ComObject ("WScript.Shell")

$ErrorActionPreference = 'Stop'

# Kill the spec-manager workbook so that the $SpecManagerDir can be overwritten.
$excel = Get-Process excel -ea 0 | Where-Object { $_.MainWindowTitle -like '*Spec Manager*' }; 
Stop-Process $excel

# Download latest dotnet/codeformatter release from github
$repo = "codylruff/SpecManager"
$releases = "https://api.github.com/repos/$repo/releases"

Write-Host Determining latest release
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name
Write-Output "Current release is : $tag"

# Initialize variables
$Version = $tag
$SpecManagerDir = "$env:APPDATA\Spec-Manager-$Version"
#$LibsDir = "$SpecManagerDir\libs"
$ConfigDir = "$SpecManagerDir\config"
#$LogsDir = "$SpecManagerDir\logs"
$ZipFile = "$SpecManagerDir\spec-manager.zip"

function SpecManagerShortcut() {
    $ShortCut = $Shell.CreateShortcut("$env:USERPROFILE\Desktop\Spec-Manager.lnk")
    $ShortCut.TargetPath="$SpecManagerDir\Spec Manager $Version.xlsm"
    $ShortCut.Description = "Spec-Manager Shortcut";
    $shortcut.IconLocation="$SpecManagerDir\Spec-Manager.ico"
    $ShortCut.Save()
}

$ReleaseUri = "https://github.com/codylruff/SpecManager/releases/download/$Version/spec-manager-$Version.zip";

if (!(Test-Path $SpecManagerDir)) {
New-Item $SpecManagerDir -ItemType Directory | Out-Null
}
  
Write-Output ("Downloading spec-manager-$Version. . .")
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest $ReleaseUri -Out $ZipFile

Write-Output ("Extracting spec-manager-$Version. . .")
Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
Remove-Item $ZipFile

if (Test-Path ($ConfigDir + "\user.json")) {
	Remove-Item ($ConfigDir + "\user.json")
}

Write-Output "Creating Shortcut"
SpecManagerShortcut

function Enable-VBOM ($App) {
    Try {
      $CurVer = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer -ErrorAction Stop
      $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"
  
      Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\$Version\$App\Security -Name AccessVBOM -Value 1 -ErrorAction Stop
    } Catch {
      Write-Output "Failed to enable access to VBA project object model for $App."
    }
  }
  
  Write-Output "Enabling access to VBA project object model..."
  Enable-VBOM "Excel"