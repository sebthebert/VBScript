'
' VBS String functions demo
'

Option Explicit

Dim str 
str = "ABCdef"

'
' Case functions
'

' LCase(str) for LowerCase(str)
WScript.Echo("LCase(" & str & ") => " & LCase(str))

' UCase(str) for UpperCase(str)
WScript.Echo("UCase(" & str & ") => " & UCase(str))

'
' Extract functions
'

' Left(str, n) returns the first n chars from str
WScript.Echo("Left(" & str & ", 2) => " & Left(str, 2))

' Right(str, n) returns the last n chars from str
WScript.Echo("Right(" & str & ", 2) => " & Right(str, 2))

' Mid(str, start, n) returns n characters starting at start from str
WScript.Echo("Mid(" & str & ", 2, 1) => " & Mid(str, 2, 1))
WScript.Echo("Mid(" & str & ", 3, 2) => " & Mid(str, 3, 2))
WScript.Echo("Mid(" & str & ", 4) => " & Mid(str, 4))

'
' Other functions
'

' Chr(n)

Dim Special_Chars
'Special_Chars(0) = 34
'Special_Chars(1) = 65
Special_Chars = Array(34,65)

Dim i
For i = 0 To UBound(Special_Chars)
	WScript.Echo("Chr(" & Special_Chars(i) & ") => " & Chr(Special_Chars(i)))
Next

' Join(list[, delimiter])
Dim array(2)
Dim str_join
array(0) = "One"
array(1) = "Two"
array(2) = "Three"
WScript.Echo("Join(array) => " & Join(array))
WScript.Echo("Join(array, ,) => " & Join(array, ","))

' Len(str) returns str length
WScript.Echo("String " & str & " has " & Len(str) & " chars. (Len(" & str & ") => " & Len(str) & ")")

' String(number, char) returns a string of n chars
WScript.Echo("String(3, A) => " & String(3, "A"))
WScript.Echo("String(8, *) => " & String(8, "*"))

' Search functions
Dim str2
str2 = "OneTwoThreeTwoOne"

' InStr
WScript.Echo("InStr(" & str2 & ", Two) => " & InStr(str2, "Two"))
WScript.Echo("InStr(6, " & str2 & ", Two) => " & InStr(6, str2, "Two"))

' InStrRev
WScript.Echo("InStrRev(" & str2 & ", Two) => " & InStrRev(str2, "Two"))
WScript.Echo("InStrRev(" & str2 & ", Two, 6) => " & InStrRev(str2, "Two", 6))

' Trim functions
Dim str_with_spaces 
str_with_spaces = "  2 spaces before, 2 spaces after  "

' LTrim(str_with_spaces)
WScript.Echo("LTrim(" & str_with_spaces & ") => |" & LTrim(str_with_spaces) & "|")

' RTrim(str_with_spaces)
WScript.Echo("RTrim(" & str_with_spaces & ") => |" & RTrim(str_with_spaces) & "|")

' Trim(str_with_spaces)
WScript.Echo("Trim(" & str_with_spaces & ") => |" & Trim(str_with_spaces) & "|")
