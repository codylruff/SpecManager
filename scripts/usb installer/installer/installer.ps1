# This script is called from within a vba code module.
# Upon user prompt the current spec-manager version will be removed
# and the newest release will be downloaded from github
# ----------------------------------------------------------------------------
# GUI Script w/ progress bar for user to see
# ----------------------------------------------------------------------------
# Load Assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Init Form Object
$Form = New-Object System.Windows.Forms.Form
$Form.width = 450
$Form.height = 150
$Form.Text = "Spec-Manager Launcher"
$Form.StartPosition = "CenterScreen"

# Init Status Text label
$StatusText                      = New-Object system.Windows.Forms.Label
$StatusText.text                 = "Initializing . . ."
$StatusText.AutoSize             = $true
$StatusText.width                = 25
$StatusText.height               = 10
$StatusText.location             = New-Object System.Drawing.Point(75,30)
$StatusText.Font                 = 'Microsoft Sans Serif,15'
$Form.Controls.Add($StatusText)

# Show Form
$Form.Show()

$Shell = New-Object -ComObject ("WScript.Shell")

$ErrorActionPreference = 'Stop'
$tls12 = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
# ----------------------------------------------------------------------------------------------------
# UPDATE :
# ----------------------------------------------------------------------------------------------------
# Download latest dotnet/codeformatter release from github
$repo = "codylruff/SpecManager"
$releases = "https://api.github.com/repos/$repo/releases"

$StatusText.Text = "Determining latest release . . ."
# Create a web client object
$webClient = New-Object System.Net.WebClient
$json = $webclient.DownloadString($releases)
[Net.ServicePointManager]::SecurityProtocol = $tls12
#$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name
$tag = ($json | ConvertFrom-Json)[0].tag_name

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
    $ShortCut.WindowStyle = 7
    $ShortCut.Save()
}

$ReleaseUri = "https://github.com/codylruff/SpecManager/releases/download/$Version/spec-manager-$Version.zip";

if (!(Test-Path $SpecManagerDir)) {
New-Item $SpecManagerDir -ItemType Directory | Out-Null
}
  
$StatusText.Text = "Downloading Spec-Manager. . ."
[Net.ServicePointManager]::SecurityProtocol = $tls12
$client = New-Object System.Net.WebClient
$client.DownloadFile($ReleaseUri, $ZipFile)
#Invoke-WebRequest $ReleaseUri -Out $ZipFile

$StatusText.Text = "Installing Spec-Manager. . ."
Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
Remove-Item $ZipFile

$StatusText.Text = "Creating Shortcut . . ."
SpecManagerShortcut

function Enable-VBOM ($App) {
    Try {
      $CurVer = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer -ErrorAction Stop
      $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"
  
      Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\$Version\$App\Security -Name AccessVBOM -Value 1 -ErrorAction Stop
    } Catch {
      $StatusText.Text = "Failed to enable access to VBA project object model for $App."
    }
  }
  
  Enable-VBOM "Excel"
# -----------------------------------------------------------------------------------------------------------
# STARTUP : This powershell code will start the application for the first time.
# -----------------------------------------------------------------------------------------------------------
$StatusText.Text = "Loading Spec-Manager . . ."
$Excel = New-Object -comobject Excel.Application
$FilePath = "C:\Users\cruff\AppData\Roaming\Spec-Manager-$tag\Spec Manager $tag.xlsm"
$Excel.Workbooks.Open($FilePath)
$Excel.visible = $true