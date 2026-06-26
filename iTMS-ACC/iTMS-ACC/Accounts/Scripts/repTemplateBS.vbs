dim TemplateRoot

set TemplateRoot = Template.documentElement

Function SaveTemplate()
	dim HeaderNode,Node,SubNode,objHttp
	dim i, l, sName,sMonth
	
	set TemplateRoot = Template.documentElement
	
	IF Trim(document.formname.selOrg.options(document.formname.selOrg.selectedIndex).text) = "Select" Then
		alert("Select Unit")
		document.formname.selOrg.focus
		exit function	
	ElseIF document.formname.SelLedgType.selectedIndex = 0 Then
		alert "Select Report For"
		document.formname.SelLedgType.focus
		exit Function
	End IF
	
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

	sName = GetTemplateName(sName)
	
	if trim(sName) = "" then exit function
	
	set HeaderNode = Template.CreateElement("TEMPLATE")
	TemplateRoot.appendChild HeaderNode
	HeaderNode.setAttribute "NO",i
	HeaderNode.setAttribute "NAME",sName

	set Node = Template.CreateElement("UNIT")
	l = document.formname.selOrg.length - 1
	for i = 0 to l
		if document.formname.selOrg.options(i).selected then
			set SubNode = Template.CreateElement("DETAILS")
			SubNode.setAttribute "CODE",document.formname.selOrg.options(i).value
			SubNode.setAttribute "NAME",document.formname.selOrg.options(i).Text
			Node.appendChild SubNode
		end if
	next
	HeaderNode.appendChild Node

	set Node = Template.CreateElement("LEDGERTYPE")
	l = document.formname.SelLedgType.length - 1
	for i = 0 to l
		if document.formname.SelLedgType.options(i).selected then
			set SubNode = Template.CreateElement("DETAILS")
			SubNode.setAttribute "CODE",document.formname.SelLedgType.options(i).value
			SubNode.setAttribute "NAME",document.formname.SelLedgType.options(i).Text
			Node.appendChild SubNode
		end if
	next
	HeaderNode.appendChild Node
	
	set Node = Template.CreateElement("TILLDATE")
	if document.formname.optPeriod(0).checked then
		l = document.formname.SelMonth.length - 1
		for i = 0 to l	

			if document.formname.SelMonth.options(i).selected then
				set SubNode = Template.CreateElement("DETAILS")
				SubNode.setAttribute "CODE",document.formname.SelMonth.options(i).value
				SubNode.setAttribute "NAME",document.formname.SelMonth.options(i).Text
				Node.appendChild SubNode
			end if
		Next
	else
		set SubNode = Template.CreateElement("DETAILS")
		SubNode.setAttribute "CODE",0
		Node.appendChild SubNode	
	end if
	HeaderNode.appendChild Node
	
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
	
	set EditDataValue = showModalDialog("..\Reports\TemplatePopupBS.asp?FileName="&document.formname.hFileName.value,Template,"dialogHeight:225px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
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
				
				if Node.nodeName = "UNIT" then
					ClearSelected document.formname.selOrg
					for each SubNode in Node.childNodes
						setSelected document.formname.selOrg, SubNode.attributes.getNamedItem("CODE").value
					Next

				elseif Node.nodeName = "LEDGERTYPE" then
						ClearSelected document.formname.SelLedgType
						for each SubNode in Node.childNodes
							setSelected document.formname.SelLedgType, SubNode.attributes.getNamedItem("CODE").value
						Next
				elseif Node.nodeName = "TILLDATE" then
						
						ClearSelected document.formname.Selmonth						
						for each SubNode in Node.childNodes
							IF SubNode.attributes.getNamedItem("CODE").value = 0 Then
								document.formname.optPeriod(1).checked = True

							Else
								document.formname.optPeriod(0).checked = True
								setSelected document.formname.Selmonth, SubNode.attributes.getNamedItem("CODE").value

							End IF	
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