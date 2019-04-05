$Shell = New-Object -ComObject ("WScript.Shell")
$ShortCut = $Shell.CreateShortcut($env:USERPROFILE + "\Desktop\Spec-Manager.lnk")
$ShortCut.TargetPath=".\Spec Manager (Test).xlsm"
$ShortCut.WorkingDirectory = $env:USERPROFILE + "\Desktop\Spec Manager";
$ShortCut.Description = "Your Custom Shortcut Description";
$shortcut.IconLocation=".\Spec-Manager.ico"
$ShortCut.Save()