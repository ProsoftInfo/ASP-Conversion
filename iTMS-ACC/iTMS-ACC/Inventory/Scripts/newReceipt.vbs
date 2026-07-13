Function Search()
'	if document.formname.selItmType.value = "" then
'		alert("Select Item Type")
'		document.formname.selItmType.focus
'		exit function
'	else
    if document.formname.seldepart.value ="select" Then
		alert("select Usage")
		document.formname.seldepart.focus
		exit function	
	end if
		
	sUnit = document.formname.hOrgID.value
	'sIType = document.formname.selItmType.value
	
	Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
    
    sTempValWindowSize = GetWindowSizeForPopup("1")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)
	

	'set ResData=showModalDialog("../../Common/ItemSelectCommon.asp?orgID=" & sUnit & "&sIType=" & sIType & "&hSelectMode=R&Flag="+cstr(nFlag),Data,"dialogHeight:650px;dialogWidth:730px;center:Yes;help:No;resizable:No;status:No")
	set ResData = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&hSelectMode=R&Flag="+cstr(nFlag),Data,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	'alert(ResData.xml)
	sAct = UCase(trim(ResData.getAttribute("Action")))
	sQuery = trim(ResData.getAttribute("PassQuery"))
'	if ucase(trim(sAct)) <> "CLOSE" then
'		do while sAct <> "DONE"
'			set OutValue=showModalDialog("../../Common/ItemSelectCommon.asp?" & sQuery,Data,"dialogHeight:650px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
'			sAct = UCase(trim(ResData.getAttribute("Action")))
'			if ucase(Trim(sAct)) = "CLOSE" then exit do
'			sQuery = trim(ResData.getAttribute("PassQuery"))
'		loop
'	end if 'if ucase(trim(sAct)) <> "CLOSE" then
	Set Root = Data.documentElement
			
	If not Root.hasChildNodes Then 	exit function	

	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			'sTemp = HeaderNode.Attributes.Item(3).nodeValue & ":" & HeaderNode.Attributes.Item(2).nodeValue
			sTemp = HeaderNode.getAttribute("ItemCode")				
			sAttributeList = HeaderNode.getAttribute("AttributeList")
			'alert(HeaderNode.xml)
			'if not CheckExists(sTemp) then
				document.formname.hItmCode.value = sTemp
				
			'	if triM(sAttributeList)<>"" then
			'	    sArrList = split(sAttributeList,":")
			'	    sAttList = split(sArrList(0),"#")
			'	    sAttID = sAttList(0)    
			'	else
			'	    sAttID = ""
			'	end if
			
			
			if Trim(sAttributeList) <> "" then
				Temp = split(sAttributeList,",")
				sAttList = ""
				For i = 0 to UBOUND(Temp) 
					sValTemp = Split(Temp(i),":")
					sArrTemp = split(sValTemp(0),"#")
					if sArrTemp(1)<>"" then
					    sAttList = sAttList &","&sArrTemp(0) 
					end if
				Next
			end if
			
			if trim(sAttList)<>"" then
			    sAttList = Mid(sAttList,2)
			    sAttrID = sAttList
			end if
			
				document.formname.hAttributeList.value = sAttList
				document.formname.submit()
			'end if
			'alert sTemp
		next
	else
		alert("No Items found")
		exit function
	end if

End Function

Function CheckType(obj) 
	document.formname.selFrombox.options.length = 0
	document.formname.selTobox.options.length = 0
	if (document.formname.selUnit.selectedIndex = "0") then
		alert("Select Unit")
		document.formname.selUnit.focus()
		obj.selectedIndex = 0
		exit function
	elseif (document.formname.selDepart.selectedIndex = "0") then
		alert("Select Usage")
		document.formname.selDepart.focus()
		obj.selectedIndex = 0
		exit function
	elseif (document.formname.selDepart.value = "PRD" and document.formname.selAddType.selectedIndex = "0") then
		alert("Select Type")
		document.formname.selAddType.focus()
		exit function
	elseif (document.formname.selItmType.selectedIndex = "0") then
		alert("Select Item Type")
		document.formname.selItmType.focus()
		obj.selectedIndex = 0
		exit function
	else
		OutDataValue = AddClass()
		if OutDataValue = "" then exit function
		
		document.formname.selFrombox.multiple = False
		
		'sTemp = document.formname.selUnit.value
		'OutDataValue = showModalDialog("receiptNewClassPoP.asp?orgID="&sTemp,OutData,"dialogHeight:250px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
		
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","itmRecXMLSelect.asp?Check=N&sTemp=" & OutDataValue &"&orgID=" & document.formname.selUnit.value , false
		objhttp.send
		
		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			popItmDisplay "N"
		else
			alert("No Item found for Classification selected.")
			clearXML
			'document.formname.selSrc.selectedIndex = "0"
		end if
	end if
end Function

Function popItmDisplay(strr)
	set recParent = OutData.recordset
	recParent.moveFirst()
	Set Root = OutData.documentElement
	if strr = "M" then
		do while not recParent.EOF
			document.formname.selFrombox.options.length = 0
			document.formname.selTobox.options.length = 0
			For Each HeaderNode In Root.childNodes
				document.formname.selFrombox.length = document.formname.selFrombox.length+1
				document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.Attributes.Item(4).nodeValue & " / " & HeaderNode.Attributes.Item(6).nodeValue 'HeaderNode.Attributes.Item(1).nodeValue & " -- " & HeaderNode.Attributes.Item(3).nodeValue
				document.formname.selFrombox.options(document.formname.selFrombox.length-1).Value = HeaderNode.Attributes.Item(5).nodeValue & ":" & HeaderNode.Attributes.Item(4).nodeValue & ":" & HeaderNode.Attributes.Item(0).nodeValue & ":" & HeaderNode.Attributes.Item(2).nodeValue
			next
			recParent.moveNext()
		loop
	else
		do while not recParent.EOF
			document.formname.selFrombox.options.length = 0
			document.formname.selTobox.options.length = 0
			For Each HeaderNode In Root.childNodes
				document.formname.selFrombox.length = document.formname.selFrombox.length+1
				document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.Attributes.Item(6).nodeValue  'HeaderNode.Attributes.Item(1).nodeValue & " -- " & HeaderNode.Attributes.Item(3).nodeValue
				document.formname.selFrombox.options(document.formname.selFrombox.length-1).Value = HeaderNode.Attributes.Item(5).nodeValue & ":" & HeaderNode.Attributes.Item(4).nodeValue & ":" & HeaderNode.Attributes.Item(0).nodeValue & ":" & HeaderNode.Attributes.Item(2).nodeValue
			next
			recParent.moveNext()
		loop
	end if
end Function

Function DisplayMRSItem()
	set recParent = OutData.recordset
	recParent.moveFirst()
	Set Root = OutData.documentElement
	do while not recParent.EOF
		document.formname.selFrombox.options.length = 0
		document.formname.selTobox.options.length = 0
		For Each HeaderNode In Root.childNodes
			document.formname.selFrombox.length = document.formname.selFrombox.length+1
			if HeaderNode.Attributes.Item(4).nodeValue <> 0 then
				if HeaderNode.Attributes.Item(9).nodeValue <> "0" then
					document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.Attributes.Item(9).nodeValue & " / " & HeaderNode.Attributes.Item(6).nodeValue
				else
					document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.Attributes.Item(4).nodeValue & " / " & HeaderNode.Attributes.Item(6).nodeValue
				end if
				document.formname.selFrombox.options(document.formname.selFrombox.length-1).Value = "M:" & HeaderNode.Attributes.Item(5).nodeValue & ":" & HeaderNode.Attributes.Item(4).nodeValue & ":" & HeaderNode.Attributes.Item(0).nodeValue & ":" & HeaderNode.Attributes.Item(2).nodeValue
			else
				if HeaderNode.Attributes.Item(8).nodeValue <> "0" then
					document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.Attributes.Item(8).nodeValue & " / " & HeaderNode.Attributes.Item(6).nodeValue
				else
					document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.Attributes.Item(7).nodeValue & " / " & HeaderNode.Attributes.Item(6).nodeValue
				end if
				document.formname.selFrombox.options(document.formname.selFrombox.length-1).Value = "D:" & HeaderNode.Attributes.Item(5).nodeValue & ":" & HeaderNode.Attributes.Item(7).nodeValue & ":" & HeaderNode.Attributes.Item(0).nodeValue & ":" & HeaderNode.Attributes.Item(2).nodeValue
			end if
		next
		recParent.moveNext()
	loop
end Function


Function clearXML()
	Set Root = OutData.documentElement
	For Each HeaderNode In Root.childNodes
		set a=Root.removeChild(HeaderNode)
	next
	document.formname.selFrombox.options.length = 0

	document.formname.selTobox.options.length = 0
	
	document.formname.reset()
end Function

Function resetAll(strr)
	if strr = "U" then 
		document.formname.selDepart.selectedIndex = "0"
		'document.formname.selItmType.selectedIndex = "0"
	end if

	if strr = "R" then 
		'document.formname.selItmType.selectedIndex = "0"
	end if
	
'	document.formname.selFrombox.options.length = 0
'	document.formname.selTobox.options.length = 0
	'document.formname.selSrc.selectedIndex = "0"
	
end Function

'Code to get the added classes
Function AddClass()
	dim sOrgID, sIType, sITypeName, arrTemp, sClass

	if (document.formname.selItmType.selectedIndex = "0") then
		alert("Select Item Type")
		document.formname.selItmType.focus()
		'document.formname.selSrc.selectedIndex = 0
		exit function
	end if
	
	sOrgID = document.formname.selUnit.value
	sIType = document.formname.selItmType.value
	sITypeName = document.formname.selItmType.options(document.formname.selItmType.selectedIndex).text

	OutValue = showModalDialog("../../include/ClassificationSelectPop.asp?sIType=" & sIType & "&sOrgID=" & sOrgID & "&sITypename="& sITypeName,"Classification","dialogHeight:460px;dialogWidth:625px;center:Yes;help:No;resizable:No;status:No")

	arrTemp = split(OutValue,"*****")

	if arrTemp(0) = "-1" then exit function

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","../Master/XMLSelectItemClass.asp?sOrgID=" & sOrgID &"&sText="&arrTemp(0), false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				sClass = sClass & ":" & HeaderNode.Attributes.getNamedItem("CLASSCODE").Value
			next
		end if
	end if

	AddClass = mid(sClass,2)
End Function

