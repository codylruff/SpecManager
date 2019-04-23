$Shell = New-Object -ComObject ("WScript.Shell")

$ErrorActionPreference = 'Stop'

$SpecManagerDir = "$env:APPDATA\Spec-Manager"
$LibsDir = "$SpecManagerDir\libs"
$ConfigDir = "$SpecManagerDir\config"
$LogsDir = "$SpecManagerDir\logs"
$ZipFile = "$SpecManagerDir\spec-manager.zip"

if ($args.Length -gt 0) {
  $Version = $args.Get(0)
} else {
  $Version = "v0.0.3"
}

function SpecManagerShortcut() {
    $ShortCut = $Shell.CreateShortcut($env:USERPROFILE + "\Desktop\Spec-Manager.lnk")
    $ShortCut.TargetPath="$SpecManagerDir\Spec Manager" + $Version + ".xlsm"
    $ShortCut.Description = "Spec-Manager Shortcut";
    $shortcut.IconLocation="$SpecManagerDir\Spec-Manager.ico"
    $ShortCut.Save()
}

$ReleaseUri = if (!$Version) {
    "https://github.com/codylruff/SpecManager/releases/download/$Version/spec-manager-v" + $Version + ".zip";
}else {
    "https://github.com/codylruff/SpecManager/releases/download/v0.0.3/spec-manager-v0.0.3.zip"
}

if (!(Test-Path $SpecManagerDir)) {
New-Item $SpecManagerDir -ItemType Directory | Out-Null
}
  
Write-Output "Downloading spec-manager..."
Write-Output "($ReleaseUri)"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest $ReleaseUri -Out $ZipFile

Write-Output "Extracting spec-manager..."
Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
Remove-Item $ZipFile

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