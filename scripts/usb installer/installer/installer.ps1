###############################################
#  GUI Script w/ progress bar for user to see #
###############################################

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

Start-Sleep -Seconds 2

$Shell = New-Object -ComObject ("WScript.Shell")
$ErrorActionPreference = 'Stop'
###################
#    FUNCTIONS    #
###################
function ConvertFrom-Json20([object] $item){ 
    add-type -assembly system.web.extensions 
    $ps_js=new-object system.web.script.serialization.javascriptSerializer 
    return $ps_js.DeserializeObject($item) 
}

function Expand-ZipFile($file, $destination)
{
    $shell_ = new-object -com shell.application
    $zip = $shell_.NameSpace($file)

    foreach($item in $zip.items())
    {
        $shell_.Namespace($destination).copyhere($item)
    }
}

function SpecManagerShortcut() {
    $ShortCut = $Shell.CreateShortcut("$env:USERPROFILE\Desktop\Spec-Manager.lnk")
    $ShortCut.TargetPath="$SpecManagerDir\scripts\start.bat"
    $ShortCut.Description = "Spec-Manager Shortcut";
    $shortcut.IconLocation="$SpecManagerDir\Spec-Manager.ico"
    $ShortCut.WindowStyle = 7
    $ShortCut.Save()
}

function Enable-VBOM ($App) {
    Try {
      $CurVer = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer -ErrorAction Stop
      $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"
  
      Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\$Version\$App\Security -Name AccessVBOM -Value 1 -ErrorAction Stop
    } Catch {
      $StatusText.Text = "Failed to enable access to VBA project object model for $App."
      $StatusText.Refresh()
      Start-Sleep -Seconds 2
      Exit
    }
  }

# ----------------------------------------------------------------------------------------------------
# Download latest archive from github for installation :
# ----------------------------------------------------------------------------------------------------
$tls12 = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$repo = "codylruff/SpecManager"
$releases = "https://api.github.com/repos/$repo/releases"

$StatusText.Text = "Determining latest release . . ."
$StatusText.Refresh()
Start-Sleep -Seconds 2

# Create a web client object
[Net.ServicePointManager]::SecurityProtocol = $tls12
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Only a test!")
$json = $webclient.DownloadString($releases)
$tag = (ConvertFrom-Json20($json))[0].tag_name

# Initialize variables
$Version = $tag
$SpecManagerDir = "$env:APPDATA\Spec-Manager"
$ZipFile = "$SpecManagerDir\spec-manager.zip"
$ReleaseUri = "https://github.com/codylruff/SpecManager/releases/download/$Version/spec-manager-$Version.zip";

if (!(Test-Path $SpecManagerDir)) {
  New-Item $SpecManagerDir -ItemType Directory | Out-Null
} else {
  # Remove old installation
  Remove-Item -LiteralPath $SpecManagerDir -Force -Recurse
  New-Item $SpecManagerDir -ItemType Directory | Out-Null
}
  
$StatusText.Text = "Downloading Spec-Manager. . ."
$StatusText.Refresh()
Start-Sleep -Seconds 2

# Check version to speed up program if PSVersion 3.0 or higher.
if($PSVersionTable.PSVersion.Major -gt 4){
    [Net.ServicePointManager]::SecurityProtocol = $tls12
    Invoke-WebRequest $ReleaseUri -Out $ZipFile
    $StatusText.Text = "Installing Spec-Manager. . ."
    $StatusText.Refresh()
    Start-Sleep -Seconds 2
    Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
    Remove-Item $ZipFile
}else {
    [Net.ServicePointManager]::SecurityProtocol = $tls12
    $client = New-Object System.Net.WebClient
    $client.Headers.Add("user-agent", "Only a test!")
    $client.DownloadFile($ReleaseUri, $ZipFile)
    $StatusText.Text = "Installing Spec-Manager. . ."
    $StatusText.Refresh()
    Start-Sleep -Seconds 2
    Expand-ZipFile $ZipFile -destination $SpecManagerDir 
    Remove-Item $ZipFile
}

$StatusText.Text = "Creating Shortcut . . ."
$StatusText.Refresh()
Start-Sleep -Seconds 2
SpecManagerShortcut
  
Enable-VBOM "Excel"
# -----------------------------------------------------------------------------------------------------------
# STARTUP : This powershell code will start the application for the first time.
# -----------------------------------------------------------------------------------------------------------
$StatusText.Text = "Loading Spec-Manager . . ."
$StatusText.Refresh()
Start-Sleep -Seconds 2
$Excel = New-Object -comobject Excel.Application
$FilePath = "$SpecManagerDir\Spec Manager $tag.xlsm"
$Excel.Workbooks.Open($FilePath)
$Excel.visible = $true
$Form.Hide