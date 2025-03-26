# Define the URL and the shortcut path
$url = "https://nottslibraries.getnetloan.co.uk/netloan"
$shortcutPath = "C:\Users\Public\Desktop\Book A PC.lnk"

# Create a COM object for WScript.Shell
$wshShell = New-Object -ComObject WScript.Shell

# Create the shortcut
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $url
$shortcut.Save()
