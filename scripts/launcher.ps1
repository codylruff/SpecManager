# ----------------------------------------------------------------------------
# GUI Script w/ progress bar for user to see
# ----------------------------------------------------------------------------
# Load Assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

function IncrementProgressBar($increment, $status){
    $current = $pbr.Value + 20
    $pbr.Value = $current
    $StatusText.text = $status
}

$Form = New-Object System.Windows.Forms.Form
$Form.width = 400
$Form.height = 600
$Form.Text = "Add Resource"

# Init ProgressBar
$pbr                            = New-Object System.Windows.Forms.ProgressBar
$pbr.Maximum                    = 100
$pbr.Minimum                    = 0
$pbr.Location                   = new-object System.Drawing.Size(10,10)
$pbr.size                       = new-object System.Drawing.Size(100,50)
$pbr.Value                      = 0
$Form.Controls.Add($pbr)

# Init Status Text label
$StatusText                      = New-Object system.Windows.Forms.Label
$StatusText.text                 = "Initializing . . ."
$StatusText.AutoSize             = $true
$StatusText.width                = 25
$StatusText.height               = 10
$StatusText.location             = New-Object System.Drawing.Point(38,144)
$StatusText.Font                 = 'Microsoft Sans Serif,10'
$Form.Controls.Add($StatusText)

# Show Form
$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()

# ----------------------------------------------------------------------------
# Startup Script
# ----------------------------------------------------------------------------
# This script is called from the spec-manager shortcut on the users desktop
$Shell = New-Object -ComObject ("WScript.Shell")
# Initialize variables
$ErrorActionPreference = 'Stop'
$tls12 = [Net.ServicePointManager]::SecurityProtocol =  [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$SpecManagerDir = (Get-Item .\).Parent.FullName
$ConfigDir = "$SpecManagerDir\config"

# CHECK FOR UPDATE :
IncrementProgressBar(20,"Checking for updates . . .")
$releases = "https://api.github.com/repos/codylruff/SpecManager/releases"
[Net.ServicePointManager]::SecurityProtocol = $tls12
$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name

# Get local version for comparison
$json_file = "$ConfigDir\local_version.json"
$JSON = Get-Content $json_file | Out-String | ConvertFrom-Json
$local_version = $JSON.app_version

if ($tag -ne $local_version) {

    IncrementProgressBar 35, "Checking for updates . . ."
    Start-Process .\update.bat
    IncrementProgressBar 15, "Verifying update . . ."
    $UpdatedFilePath = "C:\Users\cruff\AppData\Roaming\Spec-Manager-$tag\Spec Manager $tag.xlsm"
    if (!(Test-Path $UpdatedFilePath)) {
        # Notify the user of update failure and close the launcher.
        IncrementProgressBar 30, "Update failed. Contact Administrator."
        Exit
    }else {
        # Notify the user of update success and open the application
        IncrementProgressBar 20, "Update Successful!"
    }
    
} else {
    IncrementProgressBar 60, "Starting Spec-Manager . . ."
}

# START : This powershell code will start the application if there are no updates.
$Excel = New-Object -comobject Excel.Application
$FilePath = "C:\Users\cruff\AppData\Roaming\Spec-Manager-$tag\Spec Manager $tag.xlsm"
$Excel.Workbooks.Open($FilePath)
$Excel.visible = $true

# Finished Loading Spec-Manager
$complete = 100 - $pbr.Value
IncrementProgressBar $complete, "Finished."
Exit