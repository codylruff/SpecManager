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
$Launcher.width = 425
$Launcher.height = 150
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
Start-Sleep -Seconds 1

$ErrorActionPreference = 'Stop'
# ----------------------------------------------------------------------------------------------------
# Download latest archive from github for installation :
# ----------------------------------------------------------------------------------------------------
$Status.Text = "Determining latest release . . ."
$Status.Refresh()

# Initialize variables
$SpecManagerDir = "$env:APPDATA\Spec-Manager"
$ZipFile = "$SpecManagerDir\spec-manager.zip"

if (!(Test-Path $SpecManagerDir)) {
  New-Item $SpecManagerDir -ItemType Directory | Out-Null
} else {
  # Remove old installation
  Remove-Item -LiteralPath $SpecManagerDir -Force -Recurse
  New-Item $SpecManagerDir -ItemType Directory | Out-Null
}

# Download
$Status.Text = "Downloading Spec-Manager. . ."
$Status.Refresh()
DownloadZipLegacy($ZipFile)

# Install
$Status.Text = "Installing Spec-Manager . . ."
$Status.Refresh()
ExtractZipLegacy -Zip $ZipFile -OutDir $SpecManagerDir
Remove-Item $ZipFile
SpecManagerShortcut($SpecManagerDir)
Enable-VBOM "Excel"

# -----------------------------------------------------------------------------------------------------------
# STARTUP : This powershell code will start the application for the first time.
# -----------------------------------------------------------------------------------------------------------
$Status.Text = "Launching Spec-Manager . . ."
$Status.Refresh()
$Excel = New-Object -comobject Excel.Application
$tag = GetLatestVersion("Only a test!")
$FilePath = "$SpecManagerDir\Spec Manager $tag.xlsm"
#$Excel.WindowState = -4140
$Excel.visible = $true
$Excel.Workbooks.Open($FilePath)
Start-Sleep -Seconds 2
$Launcher.Hide