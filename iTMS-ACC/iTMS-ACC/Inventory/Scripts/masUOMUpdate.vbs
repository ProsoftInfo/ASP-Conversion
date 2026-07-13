Function checkSubmit()
	if (document.formname.selUoM.selectedIndex = "0") then
		alert("Select UoM")
		document.formname.selUoM.focus()
		exit function
	elseif (trim(document.formname.txtUOMName.value) = "") then
		alert("Enter UoM description")
		document.formname.txtUOMName.select()
		exit function
	else
		document.formname.action = "masUOMUpdate.asp"
		document.formname.submit()
	end if
End Function

Function GetDetails(obj)
	if (document.formname.selUoM.selectedIndex = "0") then 
		document.formname.txtUOMName.value = ""
		exit function
	end if
	document.formname.txtUOMName.value = ""
	arrTemp = split(obj.value,"|")
	document.formname.txtUOMName.value = trim(document.formname.selUoM(document.formname.selUoM.selectedIndex).text)
	if arrTemp(2) = "Y" then
		document.formname.radDecimal(0).checked = true	
	elseif arrTemp(2) = "N" then
		document.formname.radDecimal(1).checked = true	
	end if
End Function

Function Delete()
	if (document.formname.selUoM.selectedIndex = "0") then
		alert("Select UoM")
		document.formname.selUoM.focus()
		exit function
	end if
	document.formname.hUoMName.value = trim(document.formname.selUoM(document.formname.selUoM.selectedIndex).text)
	document.formname.action = "UoMDeletionUpdate.asp"
	document.formname.submit()
End function
