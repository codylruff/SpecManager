# Include
."..\include.ps1"

###############################################
#  GUI Script w/ progress bar for user to see #
###############################################
# Load Assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Init Form Object
$Launcher = New-Object System.Windows.Forms.Form
$Launcher.Width = 425
$Launcher.Height = 150
$Launcher.Text = "Spec-Manager Launcher"
$Launcher.StartPosition = "CenterScreen"

# Init Status Text label
$Status                      = New-Object system.Windows.Forms.Label
$Status.Text                 = "Initializing . . ."
$Status.AutoSize             = $true
$Status.Width                = 25
$Status.Height               = 10
$Status.Location             = New-Object System.Drawing.Point(75,30)
$Status.Font                 = 'Microsoft Sans Serif,15'
$Launcher.Controls.Add($Status)

# Show Form
$Launcher.Show()
Start-Sleep -Seconds 2


$ErrorActionPreference = 'Stop'
###################
#    FUNCTIONS    #
###################

# ----------------------------------------------------------------------------------------------------
# CHECK FOR UPDATE :
# ----------------------------------------------------------------------------------------------------
$tls12 = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$SpecManagerDir = "$env:APPDATA\Spec-Manager"
$ConfigDir = "$SpecManagerDir\config"
$repo = "codylruff/SpecManager"
$releases = "https://api.github.com/repos/$repo/releases"

$Status.Text = "Checking for updates . . ."
$Status.Refresh()
Start-Sleep -Seconds 2

# Create a web client object and check for new releases
[Net.ServicePointManager]::SecurityProtocol = $tls12
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Only a test!")
$release_json = $webclient.DownloadString($releases)
$tag = (ConvertFrom-Json20($release_json))[0].tag_name

# Get current version number for local_version.json
$json_file = "$ConfigDir\local_version.json"
$version_json = Get-Content $json_file | Out-String
$version_string = ConvertFrom-Json20($version_json)
$local_version = "v" + $version_string.app_version

if ($tag -ne $local_version) {
    $Status.Text = "Removing Previous Version . . ."
    $Status.Refresh()
    Start-Sleep -Seconds 1
    # Remove old installation
    Remove-Item -LiteralPath $SpecManagerDir -Force -Recurse
    New-Item $SpecManagerDir -ItemType Directory | Out-Null
    
    # Initialize variables
    $ZipFile = "$SpecManagerDir\spec-manager.zip"
    $ReleaseUri = "https://github.com/codylruff/SpecManager/releases/download/$tag/spec-manager-$tag.zip";
    
    # Check version to speed up program if PSVersion 5.0 or higher.
    if($PSVersionTable.PSVersion.Major -gt 4){
        $Status.Text = "Downloading Latest Version . . ."
        $Status.Refresh()
        [Net.ServicePointManager]::SecurityProtocol = $tls12
        Invoke-WebRequest $ReleaseUri -Out $ZipFile
        $Status.Text = "Installing Latest Version . . ."
        $Status.Refresh()  
        Expand-Archive $ZipFile -Destination $SpecManagerDir -Force
    }else {
        $Status.Text = "Downloading Latest Version (w/ PSv2) . . ."
        $Status.Refresh()
        [Net.ServicePointManager]::SecurityProtocol = $tls12
        $client = New-Object System.Net.WebClient
        $client.Headers.Add("user-agent", "Only a test!")
        $client.DownloadFile($ReleaseUri, $ZipFile)
        $Status.Text = "Installing Latest Version . . ."
        $Status.Refresh()
        Start-Sleep -Seconds 1
        Expand-ZipFile $ZipFile -destination $SpecManagerDir 
    }
    
    $Status.Text = "Cleaning Up . . ."
    $Status.Refresh()
    Start-Sleep -Seconds 1
    Remove-Item $ZipFile  
    SpecManagerShortcut($SpecManagerDir)

}

# -------------------------------------------------------------------------------------------------
# START : This powershell code will start the application.
# -------------------------------------------------------------------------------------------------
$Status.Text = "Launching Spec-Manager . . ."
$Status.Refresh()
Start-Sleep -Seconds 1
$Excel = New-Object -comobject Excel.Application
$FilePath = "$SpecManagerDir\Spec Manager $tag.xlsm"
#$Excel.WindowState = -4140
$Excel.visible = $true
$Excel.Workbooks.Open($FilePath)
# Close the form
$Launcher.Hide


