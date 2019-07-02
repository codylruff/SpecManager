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
    $Shell = New-Object -ComObject ("WScript.Shell")
    $ShortCut = $Shell.CreateShortcut("$env:USERPROFILE\Desktop\Spec-Manager.lnk")
    $ShortCut.TargetPath="$SpecManagerDir\scripts\start.bat"
    $ShortCut.Description = "Spec-Manager Shortcut";
    $shortcut.IconLocation="$SpecManagerDir\Spec-Manager.ico"
    $ShortCut.WindowStyle = 7
    $ShortCut.Save()
}

# Initialize variables
$ErrorActionPreference = 'Stop'
$tls12 = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$SpecManagerDir = (Get-Item .\).Parent.FullName + "\AppData\Roaming\Spec-Manager"
$ConfigDir = "$SpecManagerDir\config"
$releases = "https://api.github.com/repos/codylruff/SpecManager/releases"

# ----------------------------------------------------------------------------------------------------
# CHECK FOR UPDATE :
# ----------------------------------------------------------------------------------------------------
$StatusText.Text = "Checking for updates . . ."
$StatusText.Refresh()
Start-Sleep -Seconds 1

# Create a web client object and check for new releases
[Net.ServicePointManager]::SecurityProtocol = $tls12
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Only a test!")
$release_json = $webclient.DownloadString($releases)
$tag = (ConvertFrom-Json20($release_json))[0].tag_name
$Version = $tag

# Get current version number for local_version.json
$json_file = "$ConfigDir\local_version.json"
$version_json = Get-Content $json_file | Out-String
$version_string = ConvertFrom-Json20($version_json)
$local_version = "v" + $version_string.app_version

if ($tag -ne $local_version) {
    $StatusText.Text = "Removing Previous Version . . ."
    $StatusText.Refresh()
    Start-Sleep -Seconds 1
    # Remove old installation
    #Remove-Item -LiteralPath $SpecManagerDir -Force -Recurse
    #New-Item $SpecManagerDir -ItemType Directory | Out-Null
    
    # Initialize variables
    $ZipFile = "$SpecManagerDir\spec-manager.zip"
    $ReleaseUri = "https://github.com/codylruff/SpecManager/releases/download/$tag/spec-manager-$tag.zip";
    
    # Check version to speed up program if PSVersion 5.0 or higher.
    if($PSVersionTable.PSVersion.Major -gt 4){
        $StatusText.Text = "Downloading Latest Version . . ."
        $StatusText.Refresh()
        [Net.ServicePointManager]::SecurityProtocol = $tls12
        Invoke-WebRequest $ReleaseUri -Out $ZipFile
        $StatusText.Text = "Installing Latest Version . . ."
        $StatusText.Refresh()  
        Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
    }else {
        $StatusText.Text = "Downloading Latest Version (w/ PSv2) . . ."
        $StatusText.Refresh()
        [Net.ServicePointManager]::SecurityProtocol = $tls12
        $client = New-Object System.Net.WebClient
        $client.Headers.Add("user-agent", "Only a test!")
        $client.DownloadFile($ReleaseUri, $ZipFile)
        StatusText.Text = "Installing Latest Version . . ."
        $StatusText.Refresh()
        Start-Sleep -Seconds 1
        Expand-ZipFile $ZipFile -destination $SpecManagerDir 
    }
    
    $StatusText.Text = "Cleaning Up . . ."
    $StatusText.Refresh()
    Start-Sleep -Seconds 1
    Remove-Item $ZipFile  
    SpecManagerShortcut

}

# -------------------------------------------------------------------------------------------------
# START : This powershell code will start the application.
# -------------------------------------------------------------------------------------------------
$StatusText.Text = "Launching Spec-Manager . . ."
$StatusText.Refresh()
Start-Sleep -Seconds 1
$Excel = New-Object -comobject Excel.Application
$FilePath = "$SpecManagerDir\Spec Manager $tag.xlsm"
$Excel.WindowState = -4140
$Excel.visible = $true
$Excel.Workbooks.Open($FilePath)
# Close the form
$Form.Hide


