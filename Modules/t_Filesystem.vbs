'
' Test suite for Filesystem Module
'

Dim data, rc
Dim script_fullname

script_fullname = WScript.ScriptFullName

data = File_Read(script_fullname)
WScript.Echo(data)

Dim file_bak
file_bak = script_fullname & ".bak"

rc = File_Write(file_bak, data)