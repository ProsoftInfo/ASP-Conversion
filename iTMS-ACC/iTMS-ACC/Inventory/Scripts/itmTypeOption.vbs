dim i,j
j = 0
dim objTemp

Function window_onload()
	'set objTemp = window.dialogArguments
end Function

Function CheckName(iItm)
	dim Root
	set Root = OutDataValue.documentElement
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			sTemp = HeaderNode.Attributes.Item(0).nodeValue 
			if not(iItm = sTemp) then
				CheckName = true
			else
				CheckName = false
				exit for
			end if
		next
	else
		CheckName = true
	end if
end Function

Function CheckEntry()
	
	if (document.formname.txtValue.value = "") then
		MsgBox "Enter Option Name",0,"Option Name"
		document.formname.txtValue.select()
		Exit Function
	elseif(not CheckName(Trim(document.formname.txtValue.value))) then
		MsgBox "Option Name already entered",0,"Option Name"
		document.formname.txtValue.focus()
		Exit Function
	else
		j = j + 1
		set oRow = document.all.tblData.insertRow(j)

		set headerCell=oRow.insertCell()									
		headerCell.innerHTML = j
		headerCell.className="ExcelSerial"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.innerHTML= trim(document.formname.txtValue.value)
		headerCell.className="ExcelDisplayCell"
		headerCell.align="left"

		Set Root = OutDataValue.documentElement
		
		Set newElem = OutDataValue.createElement("OptionEntry")
		newElem.setAttribute "ONAME", trim(document.formname.txtValue.value)
		Root.appendChild newElem

		document.formname.txtValue.value = ""
	end if
		
end Function

Function CheckSubmit()
	if j = 0 then
		Msgbox "No Option Name entered",0,"Option Name"
		document.formname.txtValue.select()
		exit function
	else
		set window.returnValue = OutDataValue.documentElement
		window.close
	end if
end Function

Function window_onunload() 
	set window.returnValue = OutDataValue.documentElement
	window.close
end Function
