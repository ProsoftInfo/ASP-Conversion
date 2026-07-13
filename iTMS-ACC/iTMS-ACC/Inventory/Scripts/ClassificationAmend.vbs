Function AmendClass(gName,gPath,gKey,sVal)
Dim gArr, gcode, sTemp
	gArr = Split(gKey, ":")

	If UBound(gArr) < 3 Then 
		alert("Select Classification Created")
		Exit function
	end if
	sTemp = "C" & gArr(0) & ":" & gArr(UBound(gArr))

	document.formname.pGroup.value = sTemp
	document.formname.pName.value = gName
	document.formname.hPara.value = sVal
	document.formname.action = "MasClassificationNameAmendEntry.asp"	 
	document.formname.submit()
End Function

Function DeleteItem(gKey,sOrgID)
	if sOrgID = "select" then 
		alert("Select Organization")
		exit Function
	elseif gKey = "GRP" then 
		alert("Select an Category or an Classification")
		exit Function
	end if
	document.formname.pGroup.value = gKey
	document.formname.action = "ItmDeletionDetailsEntry.asp"
	document.formname.submit()
End function

Function DeleteClass(gKey)
	gArr = Split(gKey, ":")

	If UBound(gArr) < 3 Then 
		alert("Select Classification Created")
		Exit function
	end if
	document.formname.pGroup.value = gKey
	document.formname.action = "ClassDeletionDetailsEntry.asp"
	document.formname.submit()
End function
