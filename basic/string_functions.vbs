'
' VBS String functions demo
'

Dim str 
str = "ABCdef"

' Case functions

' LCase(str) for LowerCase(str)
WScript.Echo("LCase(" & str & ") => " & LCase(str))

' UCase(str) for UpperCase(str)
WScript.Echo("UCase(" & str & ") => " & UCase(str))

' Extract functions

' Left(str, n) returns the first n chars from str
WScript.Echo("Left(" & str & ", 2) => " & Left(str, 2))

' Right(str, n) returns the last n chars from str
WScript.Echo("Right(" & str & ", 2) => " & Right(str, 2))

' Other functions

' Len(str) returns str length
WScript.Echo("String " & str & " has " & Len(str) & " chars. (Len(" & str & ") => " & Len(str) & ")")

' Search functions

' InStr
WScript.Echo()