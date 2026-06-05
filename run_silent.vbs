' run_silent.vbs - Run PT Report silently (no console window)
' Called by Windows Task Scheduler
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' This script's directory = exe location
strDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Try exe first (distributed version), fall back to py (dev version)
strExe = strDir & "\ptreport.exe"
strPy  = strDir & "\run.py"

If objFSO.FileExists(strExe) Then
    ' exe version - pass --auto to run headless (no GUI window)
    objShell.Run """" & strExe & """ --auto", 0, False
ElseIf objFSO.FileExists(strPy) Then
    ' py version - use pythonw (no console)
    objShell.Run "pythonw.exe """ & strPy & """", 0, False
End If
