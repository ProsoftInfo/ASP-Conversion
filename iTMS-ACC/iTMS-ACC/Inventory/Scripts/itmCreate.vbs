dim OutValue
Function DisplayItemCode()
	if document.formname.selIType.value = "select" then
		alert("Select Item Type")
		document.formname.selIType.focus
		exit function
	end if

	sTempValues = document.formname.selIType.value&":"&document.formname.selIType(document.formname.selIType.selectedIndex).text
	
	showModalDialog "ExistingItemCodePop.asp?sTemp="&sTempValues,"","dialogHeight:330px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No"
end Function

Function ChangeLabel(obj)
	if obj.value = "FAB" or obj.value = "GAR" then
		'idCat.innerHTML = "Dia"
		'idDrw.innerHTML = "GSM"
	else
		'idCat.innerHTML = "Catalogue No."
		'idDrw.innerHTML = "Draw. Ver"
	end if
end Function

Function CreateItemCode(obj)
	dim sTemp
	
	if left(obj,1) = "C" then
		sTemp = document.formname.selIType.value&"&ItemCode="&document.formname.txtitmCode.value&"&ItemDesc="&document.formname.txtItmDesc.value
		if document.formname.selIType.value = "FAB" then
			OutValue = showModalDialog("itmCodeCreate.asp?sTemp="&sTemp,"","dialogHeight:490px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		elseif document.formname.selIType.value = "GAR" then
			OutValue = showModalDialog("itmGarCodeCreate.asp?sTemp="&sTemp,"","dialogHeight:490px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		else
			OutValue = showModalDialog("itmCodeCreate.asp?sTemp="&sTemp,"","dialogHeight:285px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
		end if
		
		arrTemp = split(OutValue,"``")
		document.formname.txtItmDesc.value =  arrTemp(0)

		if document.formname.selIType.value = "GAR" then
			arrValue = split(arrTemp(1),",")
			for i = 0 to ubound(arrValue)
				arrTemp = split(arrValue(i),"|")
				iLevel = iLevel & "," & arrTemp(0)
				iGroup = iGroup & "," & arrTemp(1)
				iCode = iCode & arrTemp(1)
			next
			iLevel = mid(iLevel,2)
			iGroup = mid(iGroup,2)
			document.formname.txtitmCode.value =  iCode
			document.formname.hGroup.value =  iGroup
			document.formname.hLevel.value =  iLevel
		else
			document.formname.txtitmCode.value =  arrTemp(1)
		end if
		document.formname.txtitmCode.readOnly = true
	elseif left(obj,1) = "M" then
		OutValue = showModalDialog("MachineCenterCreate.asp",document.formname.hWMCCode.value,"dialogHeight:270px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
		document.formname.hWMCCode.value = OutValue
	end if	
end Function

Function itemCodeCheck(iValue,sDesc,str)
	dim Root,HeaderNode,objhttp

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","itmCodeXMLSelect.asp", false
	
	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement

		if str ="T" then
			if Root.HaschildNodes() then
				For Each HeaderNode In Root.childNodes
					if not left(HeaderNode.Attributes.Item(0).nodeName,1) = "T" then
						if not (trim(lcase(iValue)) = lcase(HeaderNode.Attributes.Item(0).nodeValue)) then
							itemCodeCheck = false
						else
							itemCodeCheck = true
							exit for
						end if
					'	if not (trim(lcase(sDesc)) = lcase(HeaderNode.Attributes.Item(2).nodeValue)) then
					'		itemCodeCheck = false
					'	else
					'		itemCodeCheck = true
					'		exit for
					'	end if
					end if
				next
			else
				itemCodeCheck = false
			end if
		else
			if Root.HaschildNodes() then
				For Each HeaderNode In Root.childNodes
					if not (trim(lcase(iValue)) = lcase(HeaderNode.Attributes.Item(0).nodeValue)) then
						itemCodeCheck = false
					else
						itemCodeCheck = true
						exit for
					end if
				'	if not (trim(lcase(sDesc)) = lcase(HeaderNode.Attributes.Item(2).nodeValue)) then
				'		itemCodeCheck = false
				'	else
				'		itemCodeCheck = true
				'		exit for
				'	end if
				next
			else
				itemCodeCheck = false
			end if
		end if
	end if
end Function

Function DisableEnable(iValue)
	dim Root,HeaderNode,objhttp

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","XMLCodeMaster.asp?sItmType="&iValue, false
	
	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement

		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				if HeaderNode.Attributes.Item(0).nodeValue = "Y" then
					DisableEnable = true
				else
					DisableEnable = false
				end if
			next
		end if
	end if
end Function


Function CheckName()
	dim Root,HeaderNode,objhttp

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","XMLCheckItemName.asp?ItemName="&document.formname.txtItmDesc.value, false
	
	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement

		if Root.Attributes.getNamedItem("STATUS").Value = "Y" then
			if confirm("Item already exists with this name."&vbcrlf&"Do you want to create the Item.") then
				CheckName = True
			else
				CheckName = False
			end if
		else
			CheckName = True
		end if
	end if
	
end Function
