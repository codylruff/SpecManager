$Shell = New-Object -ComObject ("WScript.Shell")

$ErrorActionPreference = 'Stop'

$SpecManagerDir = "$env:APPDATA\Spec-Manager"
$LibsDir = "$SpecManagerDir\libs"
$ConfigDir = "$SpecManagerDir\config"
$ZipFile = "$SpecManagerDir\spec-manager.zip"

$ShortCut = $Shell.CreateShortcut($env:USERPROFILE + "\Desktop\Spec-Manager.lnk")
$ShortCut.TargetPath=".\Spec Manager (Test).xlsm"
$ShortCut.WorkingDirectory = $SpecManagerDir;
$ShortCut.Description = "Your Custom Shortcut Description";
$shortcut.IconLocation=".\Spec-Manager.ico"
$ShortCut.Save()

if ($args.Length -gt 0) {
    $Version = $args.Get(0)
  }

$ReleaseUri = if (!$Version) {
    "https://github.com/codylruff/DataManager/releases/download/$Version/dm_v0.0.1.zip";
}else {
    "https://github.com/codylruff/DataManager/releases/download/v0.0.1/dm_v0.0.1.zip"
}

if (!(Test-Path $SpecManagerDir)) {
New-Item $SpecManagerDir -ItemType Directory | Out-Null
}
  
Write-Output "Downloading spec-manager..."
Write-Output "($ReleaseUri)"
Invoke-WebRequest $ReleaseUri -Out $ZipFile

Write-Output "Extracting spec-manager..."
Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
Remove-Item $ZipFile

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