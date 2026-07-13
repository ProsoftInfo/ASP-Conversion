dim OutValue,Root,tempItmCode,sStatus
dim sTempCode,sTempDesc,sTempShDesc,sTempAddDesc,iTempICode
dim sRet
Set Root = OutData.documentElement

sRet = "-1"

Function CreateItemCode(obj)
	'OutValue = showModalDialog("itmCodeCreate.asp?sTemp="&document.formname.selItmType.value,"","dialogHeight:270px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	OutValue = showModalDialog("itmCodeCreate.asp","","dialogHeight:270px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,"``")
	document.formname.txtItmDesc.value =  arrTemp(0)
	document.formname.itmCode.value =  arrTemp(1)
	document.formname.itmCode.readOnly = true
end Function

Function itemCodeCheck(iValue)
	sOrgCode = trim(document.formname.hOrgCode.value)
	sStatus = ""
	sTempCode = ""
	sTempDesc = ""
	sTempShDesc = ""
	sTempAddDesc = ""
	iTempICode = ""
	
	dim Root,HeaderNode,objhttp

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","itmCodeXMLSelect.asp?sOrgCode="&sOrgCode, false
	
	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement

		if Root.HaschildNodes() then
			
			bClaExists = false
			bItmExists = false
			
			For Each HeaderNode In Root.childNodes
				if not (trim(lcase(iValue)) = lcase(HeaderNode.Attributes.Item(0).nodeValue)) then
					itemCodeCheck = false
				else
					if left(HeaderNode.Attributes.Item(0).nodeName,1) = "T" then
						
						sTempCode = HeaderNode.Attributes.Item(0).nodeValue
						sTempDesc = HeaderNode.Attributes.Item(2).nodeValue
						sTempShDesc = HeaderNode.Attributes.Item(1).nodeValue
						sTempAddDesc = HeaderNode.Attributes.Item(3).nodeValue
						iTempICode = HeaderNode.Attributes.Item(4).nodeValue
						if HeaderNode.Attributes.Item(5).nodeValue = "N" then
							sStatus = "Temporary Item code already Exists"
						elseif HeaderNode.Attributes.Item(5).nodeValue = "Y" then
							sStatus = " Temporary Item code already been created / mapped with Permanent Item " & vbcrlf & " [" & HeaderNode.Attributes.Item(6).nodeValue & "]. So select this Item from the catalogue itself."
						end if
					elseif left(HeaderNode.Attributes.Item(0).nodeName,1) = "C" then
						bClaExists = true
					elseif left(HeaderNode.Attributes.Item(0).nodeName,1) = "I" then
						bItmExists = true
					end if
					
					if bItmExists then
						sStatus = "Item code already Exists but not defined for the Unit"
					elseif bClaExists then
						sStatus = "Item code already Exists"
					end if
					
					itemCodeCheck = true
					exit for
				end if
			next
		else
			itemCodeCheck = false
		end if

	end if
end Function

Function CheckSubmit()
'	iType = document.formname.selItmType.value
'	if (document.formname.selItmType.value = "select") then
'		alert("Select Item Type")
'		document.formname.selItmType.focus()
'		exit function
'	elseif (iType = "YRN" and trim(document.formname.itmCode.value) = "") then
'		CreateItemCode(iType)
'	else
	if (not iType = "YRN" and trim(document.formname.itmCode.value) = "") then
		alert("Enter Item Code")
		document.formname.itmCode.select()
		exit function
	elseif (not iType = "YRN" and itemCodeCheck(trim(document.formname.itmCode.value))) then
		if left(sStatus,1) = "T" then
			if confirm(sStatus & ". Do you want to have the existing Temporary Item Code?") then
				document.formname.itmCode.value = sTempCode
				document.formname.txtItmDesc.value = sTempDesc
				document.formname.txtItmShDesc.value = sTempShDesc
				document.formname.txtItmAddDesc.value = sTempAddDesc
		
				InsertDetails trim(document.formname.itmCode.value),trim(document.formname.txtItmDesc.value),trim(document.formname.txtItmShDesc.value),trim(document.formname.txtItmAddDesc.value)
				tempItmCode = iTempICode
				sRet = "YES"
				window.close
				exit function
			else
				document.formname.itmCode.select()
				exit function
			end if
		elseif left(sStatus,2) = " T" then
			alert(sStatus)
			document.formname.itmCode.select()
			exit function
		else
			alert(sStatus)
			document.formname.itmCode.select()
			exit function
		end if
	elseif (iType = "YRN" and itemCodeCheck(trim(document.formname.itmCode.value))) then
		if left(sStatus,1) = "T" then
			if confirm(sStatus & ". Do you want to have the existing Temporary Item Code?") then
				document.formname.itmCode.value = sTempCode
				document.formname.txtItmDesc.value = sTempDesc
				document.formname.txtItmShDesc.value = sTempShDesc
				document.formname.txtItmAddDesc.value = sTempAddDesc
		
				InsertDetails trim(document.formname.itmCode.value),trim(document.formname.txtItmDesc.value),trim(document.formname.txtItmShDesc.value),trim(document.formname.txtItmAddDesc.value)
				tempItmCode = iTempICode
				sRet = "YES"
				window.close
				exit function
			else
				document.formname.itmCode.select()
				exit function
			end if
		else
			alert(sStatus)
			document.formname.itmCode.select()
			exit function
		end if
		alert(sStatus)
		CreateItemCode(iType)
	elseif (trim(document.formname.txtItmDesc.value) = "") then
		alert("Enter Item Description")
		document.formname.txtItmDesc.select()
		exit function
	elseif (trim(document.formname.txtItmShDesc.value) = "") then
		alert("Enter Item Short Description")
		document.formname.txtItmShDesc.select()
		exit function
	'elseif (trim(document.formname.txtItmAddDesc.value) = "") then
	'	alert("Enter Item Additional Description")
	'	document.formname.txtItmAddDesc.select()
	'	exit function
	else 
		InsertDetails trim(document.formname.itmCode.value),trim(document.formname.txtItmDesc.value),trim(document.formname.txtItmShDesc.value),trim(document.formname.txtItmAddDesc.value)

		Set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","tempitmInsert.asp", false
		objhttp.send OutData.XMLDocument

		if objhttp.responseText = "" then
			sRet = "YES"
			window.close
			exit function
		else
			MsgBox objhttp.responseText
		end if
	end if
end Function

Function window_onunload() 
	'set window.returnValue = OutData.documentElement
	if sRet = "YES" then
		window.returnValue = trim(document.formname.itmCode.value)&"``"&trim(document.formname.txtItmShDesc.value)
	end if
	window.close()
	
end Function


Function InsertDetails(iItemCodeP,sItemDescP,sItemShDescP,sItemAddDescP)
	Root.setAttribute "ITMTYPE",""' trim(document.formname.selItmType.value)
	Root.setAttribute "APPCODE", trim(document.formname.hAppCode.value)
	Root.setAttribute "MODCODE", trim(document.formname.hModCode.value)
	Root.setAttribute "CRESTAGE", trim(document.formname.hCreStage.value)

	Set newElem = OutData.createElement("TDETAILS")
	newElem.setAttribute "ITMCODE", iItemCodeP
	newElem.setAttribute "ITMDESC", sItemDescP
	newElem.setAttribute "ITMSHDESC", sItemShDescP
	newElem.setAttribute "ITMADDDESC", sItemAddDescP
	Root.appendChild newElem
	
end Function

