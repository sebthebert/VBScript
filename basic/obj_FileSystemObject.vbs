'
' VBS FileSystemObject demo
'
Dim dir
dir = "D:\Temp\"
Dim file
file = "D:\Temp\demo.txt"
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

' CreateFolder & FolderExists
If (fso.FolderExists(dir)) Then
	WScript.Echo("Folder " & dir & " already exists.")
Else
	Set folder = fso.CreateFolder(dir)
End If

' CreateTextFile & FileExists 
If (fso.FileExists(file)) Then
	WScript.Echo("File " & file & " already exists.")
Else
	Set file = fso.CreateTextFile(file)
End If