# This script is called from the spec-manager shortcut on the users desktop

# Initialize variables
$Shell = New-Object -ComObject ("WScript.Shell")
$ErrorActionPreference = 'Stop'
$tls12 = [Net.ServicePointManager]::SecurityProtocol =  [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$SpecManagerDir = (Get-Item .\).Parent.FullName
$ConfigDir = "$SpecManagerDir\config"
$repo = "codylruff/SpecManager"
$releases = "https://api.github.com/repos/$repo/releases"
# ----------------------------------------------------------------------------------------------------
# CHECK FOR UPDATE :
# ----------------------------------------------------------------------------------------------------
[Net.ServicePointManager]::SecurityProtocol = tls12
$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name

$json_file = "$ConfigDir\local_version.json"
$JSON = Get-Content $json_file | Out-String | ConvertFrom-Json
$local_version = $JSON.app_version

if ($tag -ne $local_version) {
	Start-Process .\update.bat
}else {
    # -------------------------------------------------------------------------------------------------
    # START : This powershell code will start the application if there are no updates.
    # -------------------------------------------------------------------------------------------------
    $Excel = New-Object -comobject Excel.Application
    $FilePath = "C:\Users\cruff\AppData\Roaming\Spec-Manager-$tag\Spec Manager $tag.xlsm"
    $Excel.Workbooks.Open($FilePath)
    $Excel.visible = $true
}

