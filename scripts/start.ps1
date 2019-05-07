# This script is called from the spec-manager shortcut on the users desktop
# ----------------------------------------------------------------------------
# GUI Script w/ progress bar for user to see
# ----------------------------------------------------------------------------
# Load Assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Init Form Object
$Form = New-Object System.Windows.Forms.Form
$Form.width = 425
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

# Initialize variables
$ErrorActionPreference = 'Stop'
$tls12 = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$SpecManagerDir = (Get-Item .\).Parent.FullName
$ConfigDir = "$SpecManagerDir\config"
$releases = "https://api.github.com/repos/codylruff/SpecManager/releases"

# ----------------------------------------------------------------------------------------------------
# CHECK FOR UPDATE :
# ----------------------------------------------------------------------------------------------------
$StatusText.Text = "Checking for updates . . ."
# Create a web client object
$webClient = New-Object System.Net.WebClient
$json = $webclient.DownloadString($releases)
[Net.ServicePointManager]::SecurityProtocol = $tls12
#$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name
$tag = ($json | ConvertFrom-Json)[0].tag_name

$json_file = "$ConfigDir\local_version.json"
$JSON = Get-Content $json_file | Out-String | ConvertFrom-Json
$local_version = $JSON.app_version

if ($tag -ne $local_version) {
    $StatusText.Text = "Updating Spec-Manager . . ."   
    # Initialize variables
    $NewSpecManagerDir = "$env:APPDATA\Spec-Manager-$tag"
    $ZipFile = "$NewSpecManagerDir\spec-manager.zip"
    
    $ReleaseUri = "https://github.com/codylruff/SpecManager/releases/download/$tag/spec-manager-$tag.zip";
    
    if (!(Test-Path $NewSpecManagerDir)) {
    New-Item $NewSpecManagerDir -ItemType Directory | Out-Null
    }
      
    [Net.ServicePointManager]::SecurityProtocol = $tls12
    Invoke-WebRequest $ReleaseUri -Out $ZipFile
    
    Expand-Archive $ZipFile -Destination $NewSpecManagerDir -Force
    #Remove-Item $ZipFile
    $Shell = New-Object -ComObject ("WScript.Shell")
    $ShortCut = $Shell.CreateShortcut("$env:USERPROFILE\Desktop\Spec-Manager.lnk")
    $ShortCut.TargetPath="$NewSpecManagerDir\scripts\start.bat"
    $ShortCut.Description = "Spec-Manager Launcher";
    $ShortCut.IconLocation="$NewSpecManagerDir\Spec-Manager.ico"
    $ShortCut.WindowStyle = 7
    $ShortCut.Save()

}

# -------------------------------------------------------------------------------------------------
# START : This powershell code will start the application.
# -------------------------------------------------------------------------------------------------
$StatusText.Text = "Loading Spec-Manager . . ."
$Excel = New-Object -comobject Excel.Application
$FilePath = "C:\Users\cruff\AppData\Roaming\Spec-Manager-$tag\Spec Manager $tag.xlsm"
$Excel.Workbooks.Open($FilePath)
$Excel.visible = $true

$Form.Hide


