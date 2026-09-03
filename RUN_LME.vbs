Option Explicit
Dim fso, sh, folder, rc
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("Wscript.Shell")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = folder

If fso.FileExists(folder & "\.venv\Scripts\python.exe") Then
  rc = sh.Run(".venv\Scripts\python.exe generate_lme_daily.py", 1, True)
ElseIf fso.FileExists(folder & "\venv\Scripts\python.exe") Then
  rc = sh.Run("venv\Scripts\python.exe generate_lme_daily.py", 1, True)
Else
  rc = sh.Run("py -3 generate_lme_daily.py", 1, True)
End If

If rc <> 0 Then
  MsgBox "FAILED, exit code " & rc & ". Read the black console window.", 16, "LME Daily"
End If
WScript.Quit rc
