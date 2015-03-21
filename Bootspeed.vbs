Option Explicit
On Error Resume Next
Dim Wsh, Time1, Time2, Result, PathFile, MsgResult, MsgA, AppName, KeyA, KeyB, TimeDiff
MsgA = "Please close all running applications and click on OK."
KeyA = "HKEY_CURRENT_USER\Software\RestartTime\"
KeyB = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\RestartTime"
AppName = "ReBoot-Time"
Set Wsh = CreateObject("WScript.Shell")
PathFile = """" & WScript.ScriptFullName & """"
Result = wsh.RegRead(KeyA & "Times")
if Result = "" then
MsgResult = Msgbox (MsgA, vbOKCancel, AppName)
If MsgResult = vbcancel then WScript.Quit
Wsh.RegWrite KeyA & "Times", left(Time,8), "REG_SZ"
Wsh.RegWrite KeyB, PathFile, "REG_SZ"
Wsh.Run "cmd /c Shutdown -r -t 00", false, 0 
else
Wsh.RegDelete KeyA & "Times"
Wsh.RegDelete KeyA
Wsh.RegDelete KeyB
TimeDiff = DateDiff("s",Result,left(Time,8))
MsgBox "Your computer reboots in " & TimeDiff & " seconds", VbInformation, AppName
end if
wscript.Quit
