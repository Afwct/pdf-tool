' Drag image files onto THIS .vbs file (more reliable than .bat on Windows)
Option Explicit
Dim sh, fso, root, cmd, i, arg
If WScript.Arguments.Count = 0 Then
  MsgBox "Please drag image files onto this VBS file icon." & vbCrLf & vbCrLf & "Do NOT drop on folder blank area.", vbInformation, "image2pdf"
  WScript.Quit 1
End If
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
cmd = "cmd /c cd /d """ & root & """ && """
cmd = cmd & root & "\common\python.bat"" """
cmd = cmd & root & "\core\pdf_tool.py"" image2pdf"
For i = 0 To WScript.Arguments.Count - 1
  arg = WScript.Arguments(i)
  cmd = cmd & " """ & Replace(arg, """", """""") & """"
Next
cmd = cmd & " && pause"
sh.Run cmd, 1, True
