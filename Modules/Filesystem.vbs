'
' # Filesystem.vbs
'

Const OPEN_MODE_READ = 1
Const OPEN_MODE_WRITE = 2
Const OPEN_MODE_APPEND = 8

Const OPEN_FORMAT_SYSDEFAULT = -2
Const OPEN_FORMAT_UNICODE = -1
Const OPEN_FORMAT_ASCII = 0

' ## Functions

'
' ### Current_Directory()
'
' Returns Current Directory
'
Function Current_Directory
	
	Dim script_fullname
	
	script_fullname = WScript.ScriptFullName
	Current_Directory = Left(script_fullname, InStrRev(script_fullname, "\"))
	
End Function

'
' ### File_Read(filename)
'
' Returns data from file 'filename'
'
Function File_Read(filename)

	Dim data, fso
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(filename) = True then
		Set file = fso.OpenTextFile(filename, OPEN_MODE_READ, False)
		if Err.number <> 0 then
			Set fso = Nothing
			data = -1
			Exit Function
		end if
		data = file.ReadAll()
		if Err.number <> 0 then
			data = -1
		end if
		file.Close()
	end if
	
	Set file = Nothing
	Set fso = Nothing
	
	File_Read = data
	
End Function

'
' ### File_Write(filename, data)
'
' Writes data to file 'filename'
'
Function File_Write(filename, data)

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.CreateTextFile(filename, True)
	if Err.number <> 0 then
		Set fso = Nothing
		File_Write = -4
		Exit Function
	end if
	file.Write(data)
	if Err.number <> 0 then
		File_Write = -4
	else
		File_Write = 0
	end if
	file.Close()
	Set file = Nothing
	Set fso = Nothing
	
End Function