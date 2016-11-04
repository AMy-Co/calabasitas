'<job>
'<script language="VBScript">
Option Explicit
On Error Resume Next
Dim WshShell
set WshShell = CreateObject("WScript.Shell")

'Initilizate variables
Dim strSafeDate, strSafeTime, strDateTime, strLogFilePath, strLogFileName, strUserPath, strProcIP

'Prompt forIP
Dim message, title, defaultValue
Dim strProcIPInput 
' Set prompt.
message = "Enter the ProcIP" 
' Set title.
title = "ProcIP Input"
defaultValue = "192.168.250.1"   ' Set default value.
' Display message, title, and default value.
strProcIP = InputBox(message, title, defaultValue)
' If user has clicked Cancel, set strProcIP to defaultValue 



strUserPath = WshShell.ExpandEnvironmentStrings("%userprofile%")
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
'Set strDateTime equal to a string representation of the current date and time, for use as part of a valid Windows filename
strDateTime = strSafeDate & "-" & strSafeTime
'Assemble the path and filename

strLogFileName = chr(34) & strUserPath & "\Desktop\" & strDateTime & "-" &strProcIP & ".txt" & chr(34) 



msgBox "Your entire session will be saved here: " & strLogFileName, vbOKOnly+vbInformation+vbDefaultButton1, "MonsterPickle!"

WshShell.run "cmd.exe"
WScript.Sleep 1000
'Send commands to the window as needed - IP and commands need to be customized
'Step 1 - Telnet to remote IP'

WshShell.SendKeys "telnet -f " & strLogFileName & " " & strProcIP & " 23"
WshShell.SendKeys ("{Enter}")
WScript.Sleep 1000
'Step 2 - Issue Commands with pauses'
WshShell.SendKeys "USERNAME"


WshShell.SendKeys ("{Enter}")
WScript.Sleep 1000
WshShell.SendKeys "PASSWORD"
WshShell.SendKeys ("{Enter}")
WScript.Sleep 1000
WshShell.SendKeys "PRINTCONNECTEDDEVICES"

WshShell.SendKeys ("{Enter}")
WScript.Sleep 1000
'Step 3 - Exit Command Window

'WshShell.SendKeys "exit"
'WshShell.SendKeys ("{Enter}")
'WScript.Quit 

'</script>
'</job>






