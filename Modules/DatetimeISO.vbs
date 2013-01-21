'
' # DatetimeISO.vbs
'

' ## Functions

'
' ### DatetimeISO_Now()
'
' Returns current Datetime in ISO format (yyyy-mm-ddThh:mm:ss)
'
Function DatetimeISO_Now

	Dim isodate, isomonth, isoday, isohour, isominute, isosecond

	isodate = Now
	isomonth = Month(isodate)
	isoday = Day(isodate)
	isohour = Hour(isodate)
	isominute = Minute(isodate)
	isosecond = Second(isodate)

	if (isomonth < 10) Then
		isomonth = "0" & isomonth
	End If
	if (isoday < 10) Then
		isoday = "0" & isoday
	End If
	if (isohour < 10) Then
		isohour = "0" & isohour
	End If
	if (isominute < 10) Then
		isominute = "0" & isominute
	End If
	if (isosecond < 10) Then
		isosecond = "0" & isosecond
	End If
	
	DatetimeISO_Now = Year(isodate) & "-" & isomonth & "-" & isoday & "T" & isohour & ":" & isominute & ":" & isosecond
	
End Function