'
' VBS WScript object demo
'

' WScript.Arguments
Set args = WScript.Arguments
For i = 0 to args.count - 1
	WScript.Echo("Argument " & i+1 & " => " & args(i))
Next

' WScript.FullName
WScript.Echo("WScript.FullName => " & WScript.FullName)

' WScript.Name
WScript.Echo("WScript.Name => " & WScript.Name)

' WScript.Path
WScript.Echo("WScript.Path => " & WScript.Path)

' WScript.ScriptFullName
WScript.Echo("WScript.ScriptFullName => " & WScript.ScriptFullName)

' WScript.ScriptName
WScript.Echo("WScript.ScriptName => " & WScript.ScriptName)