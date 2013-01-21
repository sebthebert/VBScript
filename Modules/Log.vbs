'
' # Log.vbs
'

Const OPEN_FOR_APPEND = 8

' ## Procedures

'
' ### Log(level, msg)
'
' Logs message 'msg' with loglevel 'level'
' 
Sub Log(level, msg)
	
	Dim datetime, log, logfile, program
	datetime = DatetimeISO_Now()
	logfile = WScript.ScriptFullName & ".log"
	program = WScript.ScriptName
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.OpenTextFile(logfile, OPEN_FOR_APPEND, True)
	log = datetime & " " & program & ": [" & UCase(level) & "] " & msg & vbCRLF
	file.Write log
	file.Close
	Set file = Nothing
	Set fso = Nothing

End Sub
