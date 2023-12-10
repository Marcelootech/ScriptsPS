$wshell = New-Object -com WScript.Shell

$wshell | Get-Member

$wshell.Popup("Esse Curso é muito legal")

$wshell.run("Notepad")
$wshell.AppActivate("Notepad")
Start-Sleep 2
$wshell.SendKeys("Esse Curso é muito legal")