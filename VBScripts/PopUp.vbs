Dim oShell
Set oShell = CreateObject("WScript.Shell")
oShell.AppActivate "Register"
oShell.SendKeys "{ENTER}"
WScript.Sleep 2000
oShell.AppActivate "Authentication Required"
WScript.Sleep 2000
oShell.SendKeys "amoeba"
oShell.SendKeys "{TAB}"
WScript.Sleep 2000
oShell.SendKeys "amo2012"
WScript.Sleep 2000
oShell.SendKeys "{ENTER}"
