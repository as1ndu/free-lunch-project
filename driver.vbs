Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c C:\path\driver.bat"
oShell.Run strArgs, 0, false