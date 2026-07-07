On Error Resume Next
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = scriptDir
scriptPath = fso.BuildPath(scriptDir, "start_mini_window.py")

Err.Clear
shell.Run "pythonw.exe """ & scriptPath & """", 0, False
If Err.Number <> 0 Then
    Err.Clear
    shell.Run "pyw -3 """ & scriptPath & """", 0, False
End If
If Err.Number <> 0 Then
    Err.Clear
    shell.Run "python """ & scriptPath & """", 0, False
End If
