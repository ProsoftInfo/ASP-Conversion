Function CheckPartyAvail(sTemp)
	Dim objhttp,sRetVal
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	Objhttp.Open "POST","PartyCheck.asp?sCallType="&sTemp, false
	Objhttp.send
	sRetVal = objhttp.responsetext
	IF Cstr(sRetVal) = "0" Then
		IF Cstr(sTemp) = "K" Then
			Msgbox "Please Enter Commission Agents "
			CheckPartyAvail = 1
		End IF
		IF Cstr(sTemp) = "U" Then
			Msgbox "Please Enter Agents "
			CheckPartyAvail = 1
		End IF
		IF Cstr(sTemp) = "P" Then
			Msgbox "Please Enter Depot Agents "
			CheckPartyAvail = 1
		End IF
		IF Cstr(sTemp) = "D" Then
			Msgbox "Please Enter Party "
			CheckPartyAvail = 1
		End IF
	Else
		CheckPartyAvail = 0
	End IF
End Function