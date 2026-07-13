dim maxLen,iCtr
Function getCode(obj)
	document.formname.selCode.options.length = 1
	dim Root,HeaderNode
	sIType=trim(obj.value)
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLSelectCode.asp?sIType=" & sIType, false
	objhttp.send 
	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				document.formname.selCode.length = document.formname.selCode.length+1
				document.formname.selCode.options(document.formname.selCode.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selCode.options(document.formname.selCode.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue & "|" & HeaderNode.Attributes.Item(2).nodeValue
			next
		end if
	else
		MsgBox "No Code Type for the Item Type Selected",0,"No Code"
		document.formname.selItmType.focus()
		Exit Function
	end if
End Function

Function setMaxLen(obj)
	 if not document.formname.selCode.selectedIndex = "0" then
		document.formname.txtCode.value = ""
		arrTemp = split(obj.value,"|")
		maxLen = cint(arrTemp(1))
		document.formname.txtCode.maxLength = maxLen
		if document.formname.selCode.selectedIndex = "5" then
			document.formname.btnAdd.disabled = false
		else
			document.formname.btnAdd.disabled = true
		end if
	else
			document.formname.btnAdd.disabled = false
	end if
	document.formname.hdMaxLen.value = maxLen
end Function

Function CheckSubmit()
	if (document.formname.selItmType(document.formname.selItmType.selectedIndex).value = "select") then
		Msgbox "Select Item Type",0,"Item Type"
		document.formname.selItmType.focus()
		exit function
	elseif (document.formname.selCode(document.formname.selCode.selectedIndex).value = "select") then
		Msgbox "Select Code Type",0,"Code Type"
		document.formname.selCode.focus()
		exit function
	elseif (trim(document.formname.txtCode.value) = "") then
		Msgbox "Enter Code",0,"Code"
		document.formname.txtCode.select()
		exit function
	elseif (len(document.formname.txtCode.value) <> Cint(Trim(document.formname.hdMaxLen.value))) then
		Msgbox "Code should be " & document.formname.hdMaxLen.value & " Characters",0,"Code"
		document.formname.txtCode.select()
		exit function
	elseif (trim(document.formname.txtCodeName.value) = "") then
		Msgbox "Enter Code Name",0,"Code Name"
		document.formname.txtCodeName.select()
		exit function
	elseif (trim(document.formname.txtDesc.value) = "") then
		Msgbox "Enter Code Description",0,"Code Description"
		document.formname.txtDesc.select()
		exit function
	else
		Set Root = OutData.documentElement
		if not Root.hasChildNodes() then
			Set newElem = OutData.createElement("Entry")
			newElem.setAttribute "CodeType", document.formname.selCode(document.formname.selCode.selectedIndex).value
			newElem.setAttribute "CodeC", trim(document.formname.txtCode.value)
			newElem.setAttribute "CodeName", trim(document.formname.txtCodeName.value)
			newElem.setAttribute "CodeDesc", trim(document.formname.txtDesc.value)
			newElem.setAttribute "ItemType", trim(document.formname.selItmType.value)
			Root.appendChild newElem
		end if

		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","MascodeAttributeInsert.asp", false
		objhttp.send OutData.XMLDocument
		
		if objhttp.responseText = "Y<BR>" then
			MsgBox("Code Type " + trim(document.formname.txtCodeName.value) +" has been defined sucessfully")
			document.formname.reset()
			if (confirm("Do You want to create another Code Type")) then
				window.location.href = "MascodeAttributeEntry.asp"
			else
				window.location.href = "../welcome_Inventory.asp"
			end if
		elseif objhttp.responseText = "N<BR>" then
			MsgBox("Code Type " + trim(document.formname.txtCodeName.value) + " has already been defined")
			window.location.href = "MascodeAttributeEntry.asp"
		else
			MsgBox objhttp.responseText
		end if	
	end if
end Function

Function winOpen()
	if (document.formname.selItmType(document.formname.selItmType.selectedIndex).value = "select") then
		Msgbox "Select Item Type",0,"Item Type"
		document.formname.selItmType.focus()
		exit function
	elseif (document.formname.selCode(document.formname.selCode.selectedIndex).value = "select") then
		Msgbox "Select Code Type",0,"Code Type"
		document.formname.selCode.focus()
		exit function
	elseif (trim(document.formname.txtCode.value) = "") then
		Msgbox "Enter Code",0,"Code"
		document.formname.txtCode.select()
		exit function
	elseif (trim(document.formname.txtCodeName.value) = "") then
		Msgbox "Enter Code Name",0,"Code Name"
		document.formname.txtCodeName.select()
		exit function
	elseif (trim(document.formname.txtDesc.value) = "") then
		Msgbox "Enter Code Description",0,"Code Description"
		document.formname.txtDesc.select()
		exit function
	else
		Set Root = OutData.documentElement
		if not Root.hasChildNodes() then
			Set newElem = OutData.createElement("Entry")
			newElem.setAttribute "CodeType", document.formname.selCode.value
			newElem.setAttribute "CodeC", trim(document.formname.txtCode.value)
			newElem.setAttribute "CodeName", trim(document.formname.txtCodeName.value)
			newElem.setAttribute "CodeDesc", trim(document.formname.txtDesc.value)
			newElem.setAttribute "ItemType", trim(document.formname.selItmType.value)
			Root.appendChild newElem
		end if
	
		For Each HeaderNode In Root.childNodes
			if HeaderNode.nodeName <> "Entry" then
				set a = Root.removeChild(HeaderNode)
			end if
		Next
		window.showModalDialog "masAttrAddlPopEntry.asp?Code="&trim(document.formname.txtCode.value),OutData,"dialogHeight:380px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No"
	end if
End Function

Function DisplayDet(sText)
	FontFace="Verdana,8,,bold" 
	TopicText = sText
	document.formname.Det.TextPopup TopicText, FontFace, 10,10,0,0
end Function
