
<%
Function DatMon(Consump,MonthString)
Dim sYear,sMonth,sLast,sLastday,Final,sFirstday
If Trim(Consump) = "F" Then
    
    sYear = Left(MonthString, 4)
    sMonth = Right(MonthString, 2)
    sFirstDay = "01/"&sMonth&"/"&sYear
    sLast = LastDay(sFirstDay,"D")
    sLastday = "15/"&sMonth&"/"&sYear
    
     Final = sFirstDay&"|"&sLastday
    
End If

If Trim(Consump) = "D" Then
    sYear = Left(MonthString, 4)
    sMonth = Right(MonthString, 2)
    
    sFirstDay = "15/"&sMonth&"/"&sYear
    sLast = LastDay(sFirstDay,"D")
    sLastday = sLast&"/"&sMonth&"/"&sYear
    
     Final = sFirstDay&"|"&sLastday
End If
If Trim(Consump) = "M" Then
    sYear = Left(MonthString, 4)
    sMonth = Right(MonthString, 2)
    
    sFirstDay = "01/"&sMonth&"/"&sYear
    sLast = LastDay(sFirstDay,"D")
    
    sLastday = sLast&"/"&sMonth&"/"&sYear
    
     Final = sFirstDay&"|"&sLastday
End If
If Trim(Consump) = "Q" Then
    sYear = Left(MonthString, 4)
    sMonth = Right(MonthString, 2)
    
    If sMonth = "06" Then
        sFirstDay = "01/04/"&sYear
        sLastday = "30/06/"&sYear
    ElseIf sMonth = "09" Then
        sFirstDay = "01/07/"&sYear
        sLastday = "30/09/"&sYear
    ElseIf sMonth = "12" Then
        sFirstDay = "01/10/"&sYear
        sLastday = "31/12/"&sYear
    ElseIf sMonth = "03" Then
        sFirstDay = "01/01/"&sYear
        sLastday = "31/03/"&sYear
    End If
    Final = sFirstDay&"|"&sLastday
End If
If Trim(Consump) = "H" Then
    sYear = Left(MonthString, 4)
    sMonth = Right(MonthString, 2)
    If sMonth = "09" Then
        sFirstDay = "01/04/"&sYear
        sLastday = "30/09/"&sYear
   ElseIf sMonth = "03" Then
        sFirstDay = "01/10/"&sYear - 1
        sLastday = "31/03/"&sYear
    End If
    Final = sFirstDay&"|"&sLastday
End If
If Trim(Consump) = "Y" Then
    sYear = Left(MonthString, 4)
    sMonth = Right(MonthString, 2)
    sFirstDay = "01/04/"&sYear - 1
    sLastday = "31/03/"&sYear
    Final = sFirstDay&"|"&sLastday
End If
DatMon = Final
End Function
%>
<%
Function Lastday(sDate,sReturntype)
Dim stDay,sLeap,sMonth,sYear,sLastday
sMonth = Mid(sDate,4,2)
sYear = Right(sDate,4)

if(cInt(sMonth) <> 2) Then
	if Cint(sMonth) = 2 or Cint(sMonth) = 4 or Cint(sMonth) = 6 or Cint(sMonth) = 9 or Cint(sMonth) = 11 Then
		sLastday = 30
	else
		sLastday = 31
	end if
else
	If sYear Mod 4 = 0 Then
		If sYear Mod 100 = 0 Then
			If sYear Mod 400 = 0 Then
				'sLastday = 29
			Else
				'LeapYear = False
				sLastday = 28
			End If
		Else
			'LeapYear = True
			sLastday = 28
		End if
	Else
		'LeapYear = False
		sLastday = 28
	End If

end if
sTDay = sLastday&"/"&sMonth&"/"&sYear
If Trim(sReturntype) = "F" then
	LastDay = sTDay
else
	LastDay = sLastday
End if
End Function
%>
