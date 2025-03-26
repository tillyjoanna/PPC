$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Notepad++.lnk")
$Shortcut.TargetPath="C:\Program Files\Notepad++"
$Shortcut.Arguments="/l"
$Shortcut.IconLocation="C:\Program Files\Notepad++\notepad++.exe"
$Shortcut.Save()
