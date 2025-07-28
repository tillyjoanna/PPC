# Define the URL and the shortcut path
$url = "https://nottslibraries.getnetloan.co.uk/netloan"
$shortcutPath = "C:\Users\Public\Desktop\Book A PC.lnk"

# Create a COM object for WScript.Shell
$wshShell = New-Object -ComObject WScript.Shell

# Create the shortcut
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $url
$shortcut.IconLocation = "shell32.dll,15"
$shortcut.Save()

# Define the file path and the shortcut path
$url = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$shortcutPath = "C:\Users\Public\Desktop\Microsoft Word.lnk"

# Create a COM object for WScript.Shell
$wshShell = New-Object -ComObject WScript.Shell

# Create the shortcut
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $url
$shortcut.IconLocation = "%ProgramFiles%\Microsoft Office\Root\VFS\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\wordicon.exe"
$shortcut.Save()

# Define the file path and the shortcut path
$url = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$shortcutPath = "C:\Users\Public\Desktop\Microsoft Excel.lnk"

# Create a COM object for WScript.Shell
$wshShell = New-Object -ComObject WScript.Shell

# Create the shortcut
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $url
$shortcut.IconLocation = "%ProgramFiles%\Microsoft Office\Root\VFS\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\xlicons.exe"
$shortcut.Save()

# Define the file path and the shortcut path
$url = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
$shortcutPath = "C:\Users\Public\Desktop\Microsoft Powerpoint.lnk"

# Create a COM object for WScript.Shell
$wshShell = New-Object -ComObject WScript.Shell

# Create the shortcut
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $url
$shortcut.IconLocation = "%ProgramFiles%\Microsoft Office\Root\VFS\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\pptico.exe"
$shortcut.Save()

New-Item -Path 'C:\Program Files\uvnc bvba\UltraVNC' -ItemType Directory 
Invoke-WebRequest 'https://raw.githubusercontent.com/tillyjoanna/PPC/refs/heads/main/background.bmp' -OutFile 'C:\Program Files\uvnc bvba\UltraVNC\Background.bmp'
