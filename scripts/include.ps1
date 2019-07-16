###################
#    FUNCTIONS    #
###################
$tls12 = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$repo = "codylruff/SpecManager"
$releases = "https://api.github.com/repos/$repo/releases"

function GetLatestVersion($text){
    # Create a web client object
    [Net.ServicePointManager]::SecurityProtocol = $tls12
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("user-agent", $text)
    $json = $webclient.DownloadString($releases)
    return (ConvertFrom-Json20($json))[0].tag_name
}
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

function SpecManagerShortcut($MainDir) {
    $Shell = New-Object -ComObject ("WScript.Shell")
    $ShortCut = $Shell.CreateShortcut("$env:USERPROFILE\Desktop\Spec-Manager.lnk")
    $ShortCut.TargetPath="$MainDir\scripts\start.bat"
    $ShortCut.Description = "Spec-Manager Shortcut";
    $shortcut.IconLocation="$MainDir\Spec-Manager.ico"
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

function DownloadZipLegacy($Zip){
    # Check version to speed up program if PSVersion 5.0 or higher.
    $Version = GetLatestVersion("Only a test!")
    $Uri = "https://github.com/codylruff/SpecManager/releases/download/$Version/spec-manager-$Version.zip";
    if($PSVersionTable.PSVersion.Major -gt 4){
        [Net.ServicePointManager]::SecurityProtocol = $tls12
        Invoke-WebRequest $Uri -Out $Zip
    }else {
        [Net.ServicePointManager]::SecurityProtocol = $tls12
        $client = New-Object System.Net.WebClient
        $client.Headers.Add("user-agent", "Only a test!")
        $client.DownloadFile($Uri, $Zip)
    }
}

function ExtractZipLegacy($Zip, $OutDir){
    # Check version to speed up program if PSVersion 5.0 or higher.
    if($PSVersionTable.PSVersion.Major -gt 4){
        Expand-Archive $Zip -DestinationPath $OutDir -Force
    }else {
        Expand-ZipFile $Zip -destination $OutDir
    }
}