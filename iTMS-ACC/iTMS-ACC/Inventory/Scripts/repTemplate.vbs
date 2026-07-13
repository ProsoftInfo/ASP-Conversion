dim TemplateRoot

set TemplateRoot = Template.documentElement

Function SaveTemplate(sFromDate,sToDate)
	dim HeaderNode,Node,SubNode,objHttp 
	dim i, l, sName

	set TemplateRoot = Template.documentElement

	'alert TemplateRoot.xml
'	if document.formname.selUnit.value = "" then
'		alert("Select Unit")
'		document.formname.selUnit.focus
'		exit function
'	else
	if document.formname.hFileName.value <> "MaterialConsumption" then
		if document.formname.selIType.value = "" then
			alert("Select Item Type")
			document.formname.selIType.focus
			exit function
		end if
'	elseif document.formname.selClass.length > 0 and document.formname.selClass.selectedindex = -1 then 
'		alert("Select Classification")
'		document.formname.selClass.focus
'		exit function
	end if

	i = 1
	if trim(document.formname.hEditNo.value) <> "" then
		for each HeaderNode in TemplateRoot.ChildNodes
			if trim(HeaderNode.attributes.getNamedItem("NO").value) = trim(document.formname.hEditNo.value) then
				sName = HeaderNode.attributes.getNamedItem("NAME").value
				TemplateRoot.removeChild HeaderNode
				i = trim(document.formname.hEditNo.value)
				exit for
			end if
		next
	else
		for each HeaderNode in TemplateRoot.ChildNodes
			i = i + 1
		next
	end if

	if i > 5 then exit function

	'sName = InputBox("Enter Template Name","Template Creation",sName)
	sName = GetTemplateName(sName)

	'alert "sName = " & sName
	if trim(sName) = "" then exit function
	
	set HeaderNode = Template.CreateElement("TEMPLATE")
	TemplateRoot.appendChild HeaderNode
	HeaderNode.setAttribute "NO",i
	HeaderNode.setAttribute "NAME",sName

	if sFromDate <> "" and sToDate <> "" then
		set Node = Template.CreateElement("DATE")
		Node.setAttribute "FROM",document.formname.ctlFromDate.getDate()
		Node.setAttribute "TO",document.formname.ctlToDate.getDate()
		HeaderNode.appendChild Node
	end if
	
	set Node = Template.CreateElement("UNIT")
'	l = document.formname.selUnit.length - 1
'	for i = 0 to l
'		if document.formname.selUnit.options(i).selected then
			set SubNode = Template.CreateElement("DETAILS")
			SubNode.setAttribute "CODE",document.formname.hOrgID.value
			SubNode.setAttribute "NAME",document.formname.hOrgName.value
			Node.appendChild SubNode
'		end if
'	next
	HeaderNode.appendChild Node

	if document.formname.hFileName.value <> "MaterialConsumption" then
		set Node = Template.CreateElement("ITEMTYPE")
		l = document.formname.selIType.length - 1
		for i = 0 to l
			if document.formname.selIType.options(i).selected then
				set SubNode = Template.CreateElement("DETAILS")
				SubNode.setAttribute "CODE",document.formname.selIType.options(i).value
				SubNode.setAttribute "NAME",document.formname.selIType.options(i).Text
				Node.appendChild SubNode
			end if
		next
		HeaderNode.appendChild Node
	end if
	
'	set Node = Template.CreateElement("CLASSIFICATION")
'	l = document.formname.selClass.length - 1
'	for i = 0 to l
'		if document.formname.selClass.options(i).selected then
'			set SubNode = Template.CreateElement("DETAILS")
'			SubNode.setAttribute "CODE",document.formname.selClass.options(i).value
'			SubNode.setAttribute "NAME",document.formname.selClass.options(i).Text
'			Node.appendChild SubNode
'		end if
'	next
'	HeaderNode.appendChild Node

'	if document.formname.hFileName.value <> "ItemtobeDefined" then
'		set Node = Template.CreateElement("ITEM")
'		l = document.formname.selItem.length - 1
'		for i = 0 to l
'			if document.formname.selItem.options(i).selected then
'				set SubNode = Template.CreateElement("DETAILS")
'				SubNode.setAttribute "CODE",document.formname.selItem.options(i).value
'				SubNode.setAttribute "NAME",document.formname.selItem.options(i).Text
'				Node.appendChild SubNode
'			end if
'		next
'		HeaderNode.appendChild Node
'	end if
	
	'Creating additional Elements based on Report
	if document.formname.hFileName.value = "StoreLedger" then
		set Node = Template.CreateElement("LOCATION")
		l = document.formname.selLocName.length - 1
		
		if document.formname.selLocName.value <> "" then
			for i = 0 to l
				if document.formname.selLocName.options(i).selected then
					set SubNode = Template.CreateElement("DETAILS")
					SubNode.setAttribute "CODE",document.formname.selLocName.options(i).value
					SubNode.setAttribute "NAME",document.formname.selLocName.options(i).Text
					Node.appendChild SubNode
				end if
			next
		else
			for i = 0 to l
				set SubNode = Template.CreateElement("DETAILS")
				SubNode.setAttribute "CODE",document.formname.selLocName.options(i).value
				SubNode.setAttribute "NAME",document.formname.selLocName.options(i).Text
				Node.appendChild SubNode
			next
		end if
		
		HeaderNode.appendChild Node
		
	elseif document.formname.hFileName.value = "StageWiseStock" then
		set Node = Template.CreateElement("STAGE")
		l = document.formname.selStage.length - 1
		
		if document.formname.selStage.value <> "" then
			for i = 0 to l
				if document.formname.selStage.options(i).selected then
					set SubNode = Template.CreateElement("DETAILS")
					SubNode.setAttribute "CODE",document.formname.selStage.options(i).value
					SubNode.setAttribute "NAME",document.formname.selStage.options(i).Text
					Node.appendChild SubNode
				end if
			next
		else
			for i = 0 to l
				set SubNode = Template.CreateElement("DETAILS")
				SubNode.setAttribute "CODE",document.formname.selStage.options(i).value
				SubNode.setAttribute "NAME",document.formname.selStage.options(i).Text
				Node.appendChild SubNode
			next
		end if
		
		HeaderNode.appendChild Node
		
	elseif document.formname.hFileName.value = "PartyWiseReceiptItem" then
		set Node = Template.CreateElement("PARTY")
		l = document.formname.selParty.length - 1
		
		if document.formname.selParty.value <> "" then
			for i = 0 to l
				if document.formname.selParty.options(i).selected then
					set SubNode = Template.CreateElement("DETAILS")
					SubNode.setAttribute "CODE",document.formname.selParty.options(i).value
					SubNode.setAttribute "NAME",document.formname.selParty.options(i).Text
					Node.appendChild SubNode
				end if
			next
		else
			for i = 0 to l
				set SubNode = Template.CreateElement("DETAILS")
				SubNode.setAttribute "CODE",document.formname.selParty.options(i).value
				SubNode.setAttribute "NAME",document.formname.selParty.options(i).Text
				Node.appendChild SubNode
			next
		end if
		
		HeaderNode.appendChild Node
	end if

	'alert TemplateRoot.xml

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objHttp.open "POST","..\Reports\XMLSave.asp?UserId=True&Value="&document.formname.hFileName.value&"&Folder=..\XmlData\Reports",false
	objhttp.send Template.XMLDocument

	if objHttp.responseText = "" then
		alert "Template created successfully"
	else
		alert objHttp.responseText
	end if
End Function

Function OpenTemplate()
	dim No

	set TemplateRoot = Template.documentElement

	document.formname.hEditNo.value = ""
	TemplateRoot.attributes.getNamedItem("NO").value = ""

	set EditDataValue = showModalDialog("..\Reports\TemplatePopup.asp?FileName="&document.formname.hFileName.value,Template,"dialogHeight:225px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
	set TemplateRoot = Template.DocumentElement

	No = TemplateRoot.attributes.getNamedItem("NO").value

	if trim(No) <> "" then 
		document.formname.hEditNo.value = No
		EditTemplate No 
	end if

End Function

Function EditTemplate(No)

	For Each HeaderNode in TemplateRoot.childNodes
		if trim(HeaderNode.attributes.getNamedItem("NO").value) = trim(No) then
			for each Node in HeaderNode.childNodes
				if Node.nodeName = "DATE" then
					document.formname.ctlFromDate.SetDate = Node.attributes.getNamedItem("FROM").value
					document.formname.ctlToDate.SetDate = Node.attributes.getNamedItem("TO").value

				elseif Node.nodeName = "UNIT" then
					ClearSelected document.formname.selUnit
					for each SubNode in Node.childNodes
						setSelected document.formname.selUnit, SubNode.attributes.getNamedItem("CODE").value
					Next

				elseif Node.nodeName = "ITEMTYPE" then
					ClearSelected document.formname.selIType 
					for each SubNode in Node.childNodes
						setSelected document.formname.selIType, SubNode.attributes.getNamedItem("CODE").value
					Next

			'	elseif Node.nodeName = "CLASSIFICATION" then
			'		document.formname.selClass.length = 0
			'		for each SubNode in Node.childNodes
			'			document.formname.selClass.length = document.formname.selClass.length + 1
			'			document.formname.selClass.options(document.formname.selClass.length-1).value = SubNode.attributes.getNamedItem("CODE").value
			'			document.formname.selClass.options(document.formname.selClass.length-1).text = SubNode.attributes.getNamedItem("NAME").value
			'			document.formname.selClass.options(document.formname.selClass.length-1).selected = True
			'		Next
'
			'	elseif Node.nodeName = "ITEM" then
			'		document.formname.selItem.length = 0
			'		for each SubNode in Node.childNodes
			'			document.formname.selItem.length = document.formname.selItem.length + 1
			'			document.formname.selItem.options(document.formname.selItem.length-1).value = SubNode.attributes.getNamedItem("CODE").value
			'			document.formname.selItem.options(document.formname.selItem.length-1).text = SubNode.attributes.getNamedItem("NAME").value
			'			document.formname.selItem.options(document.formname.selItem.length-1).selected = True
			'		Next
				elseif Node.nodeName = "LOCATION" then
					document.formname.selLocName.length = 0
					for each SubNode in Node.childNodes
						document.formname.selLocName.length = document.formname.selLocName.length + 1
						document.formname.selLocName.options(document.formname.selLocName.length-1).value = SubNode.attributes.getNamedItem("CODE").value
						document.formname.selLocName.options(document.formname.selLocName.length-1).text = SubNode.attributes.getNamedItem("NAME").value
						document.formname.selLocName.options(document.formname.selLocName.length-1).selected = True
					Next
				elseif Node.nodeName = "STAGE" then
					document.formname.selStage.length = 0
					for each SubNode in Node.childNodes
						document.formname.selStage.length = document.formname.selStage.length + 1
						document.formname.selStage.options(document.formname.selStage.length-1).value = SubNode.attributes.getNamedItem("CODE").value
						document.formname.selStage.options(document.formname.selStage.length-1).text = SubNode.attributes.getNamedItem("NAME").value
						document.formname.selStage.options(document.formname.selStage.length-1).selected = True
					Next
				elseif Node.nodeName = "PARTY" then
					document.formname.selParty.length = 0
					for each SubNode in Node.childNodes
						document.formname.selParty.length = document.formname.selParty.length + 1
						document.formname.selParty.options(document.formname.selParty.length-1).value = SubNode.attributes.getNamedItem("CODE").value
						document.formname.selParty.options(document.formname.selParty.length-1).text = SubNode.attributes.getNamedItem("NAME").value
						document.formname.selParty.options(document.formname.selParty.length-1).selected = True
					Next
					
				end if
			next
		end if
	next
End Function

Function setSelected(obj,sTemp)
	dim i
	for i = 0 to obj.length - 1
		if trim(sTemp) = trim(obj.options(i).value) then
			obj.item(i).selected = true
			exit function
		end if
	next
End Function

Function ClearSelected(obj)
	dim i
	for i = 0 to obj.length - 1
		obj.item(i).selected = False
	next
End Function

Function GetTemplateName(sValue)

	sValue = InputBox("Enter Template Name","Template Creation",sValue)

	set TemplateRoot = Template.documentElement

	for each HeaderNode in TemplateRoot.ChildNodes
		if Ucase(trim(HeaderNode.attributes.getNamedItem("NAME").value)) = Ucase(trim(sValue)) then
			alert "Template already exists with this Name"
			GetTemplateName sValue
		end if
	next
	
	GetTemplateName = sValue
End Function

