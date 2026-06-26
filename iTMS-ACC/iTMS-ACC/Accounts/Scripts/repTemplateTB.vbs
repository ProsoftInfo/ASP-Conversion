dim TemplateRoot

set TemplateRoot = Template.documentElement

Function SaveTemplate()
	
	dim HeaderNode,Node,SubNode,objHttp
	dim i, l, sName
	
	set TemplateRoot = Template.documentElement
	
	'IF Trim(document.formname.selOrg.options(document.formname.selOrg.selectedIndex).text) = "Select" Then
	'	alert("Select Unit")
	'	document.formname.selOrg.focus
	'	exit function
	'End IF
	
	IF document.formname.hFileName.value = "GRPTRIALBALANCE" Then
		IF document.formname.OptUnitSel(1).checked Then
			IF document.formname.selunit.selectedindex = 0 Then
				alert ("Select unit")
				document.formname.selunit.focus
				exit function
			End IF	
		End IF	

		IF document.formname.SelLedgType.value = "S" Then
			alert ("Select Ledger Type")
			document.formname.SelLedgType.focus
			exit Function
		End IF		
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
	
	IF document.formname.hFileName.value <> "GRPTRIALBALANCE" Then
	set Node = Template.CreateElement("DATE")	
		set SubNode = Template.CreateElement("DETAILS")
		SubNode.setAttribute "FROMMONTH",document.formname.selFromMonth.options(document.formname.selFromMonth.selectedIndex).value
		SubNode.setAttribute "TOMONTH",document.formname.selToMonth.options(document.formname.selToMonth.selectedIndex).value
		Node.appendChild SubNode
	HeaderNode.appendChild Node	
	Else	
	set Node = Template.CreateElement("DATE")	
		set SubNode = Template.CreateElement("DETAILS")
		SubNode.setAttribute "TILLMONTH",document.formname.selFromMonth.options(document.formname.selFromMonth.selectedIndex).value		
		Node.appendChild SubNode
	HeaderNode.appendChild Node	
	End IF
	
	set Node = Template.CreateElement("UNIT")
	'l = document.formname.selOrg.length - 1
	'for i = 0 to l
	'	if document.formname.selOrg.options(i).selected then
	'		set SubNode = Template.CreateElement("DETAILS")
	'		SubNode.setAttribute "CODE",document.formname.selOrg.options(i).value
	'		SubNode.setAttribute "NAME",document.formname.selOrg.options(i).Text
	'		Node.appendChild SubNode
	'	end if
	'next
	set SubNode = Template.CreateElement("DETAILS")
	SubNode.setAttribute "CODE",document.formname.hUnitID.value
	SubNode.setAttribute "NAME",document.formname.hUnitName.value
	Node.appendChild SubNode
	
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
	
	IF document.formname.hFileName.value = "TRIALBALANCE" Then
		set Node = Template.CreateElement("PARTYHEAD")
			l = document.formname.SelPartyHead.length - 1
			for i = 0 to l
				if document.formname.SelPartyHead.options(i).selected then
					set SubNode = Template.CreateElement("DETAILS")
					SubNode.setAttribute "CODE",document.formname.SelPartyHead.options(i).value
					SubNode.setAttribute "NAME",document.formname.SelPartyHead.options(i).Text
					SubNode.setAttribute "SNAME",Window.spPartyHead.innerHTML					
					SubNode.setAttribute "FROMMONTH",document.formname.SelFromMonth.options(document.formname.SelFromMonth.selectedindex).value
					SubNode.setAttribute "TOMONTH",document.formname.SelToMonth.options(document.formname.SelToMonth.selectedindex).value
					SubNode.setAttribute "PARTYSUBTYPE",document.formname.hdPartySubType.value
					Node.appendChild SubNode					
					end if
			next
		HeaderNode.appendChild Node
		
	ElseIF document.formname.hFileName.value = "LEDSelection" Then
		set Node = Template.CreateElement("PARTYHEAD")
			l = document.formname.SelPartyHead.length - 1
			for i = 0 to l
				if document.formname.SelPartyHead.options(i).selected then
					set SubNode = Template.CreateElement("DETAILS")
					SubNode.setAttribute "CODE",document.formname.SelPartyHead.options(i).value
					SubNode.setAttribute "NAME",document.formname.SelPartyHead.options(i).Text
					SubNode.setAttribute "SNAME",Window.spPartyHead.innerHTML					
					SubNode.setAttribute "FROMMONTH",document.formname.SelFromMonth.options(document.formname.SelFromMonth.selectedindex).value
					SubNode.setAttribute "TOMONTH",document.formname.SelToMonth.options(document.formname.SelToMonth.selectedindex).value
					SubNode.setAttribute "PARTYSUBTYPE",document.formname.hdPartySubType.value
					SubNode.setAttribute "PARTYCODE",document.formname.hdPartyCode.value
					SubNode.setAttribute "PARTYNAME",document.formname.hdPartyName.value
					IF document.formname.hdAcc.value <> "" Then
						SubNode.setAttribute "ACC",document.formname.hdAcc.value
					Else
						SubNode.setAttribute "ACC","0:0"
					End IF
				Node.appendChild SubNode	
				end if				
			next
		HeaderNode.appendChild Node
	Else
	set Node = Template.CreateElement("DISPLAYTYPE")			
			
				if document.formname.OptUnitSel(1).checked then
					set SubNode = Template.CreateElement("DETAILS")
					SubNode.setAttribute "TYPE",document.formname.OptUnitSel(1).value
					SubNode.setAttribute "CODE","S:"&document.formname.SelUnit.value
					SubNode.setAttribute "NAME",document.formname.SelUnit.options(document.formname.SelUnit.selectedindex).Text
					
					
				else
					set SubNode = Template.CreateElement("DETAILS")
					SubNode.setAttribute "TYPE",document.formname.OptUnitSel(0).value
					SubNode.setAttribute "CODE", "A:0"
					
				end if
					SubNode.setAttribute "FROMMONTH",document.formname.SelFromMonth.options(document.formname.SelFromMonth.selectedindex).value
					Node.appendChild SubNode
		HeaderNode.appendChild Node
	End IF
	
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
	
	IF document.formname.hFileName.value = "TRIALBALANCE" Then
		set EditDataValue = showModalDialog ("..\Reports\TemplatePopupTB.asp?FileName="&document.formname.hFileName.value,Template,"dialogHeight:225px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	Else
		set EditDataValue = showModalDialog ("..\Reports\TemplatePopupTB.asp?FileName="&document.formname.hFileName.value,Template,"dialogHeight:225px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
	End IF
	set TemplateRoot = Template.DocumentElement

	No = TemplateRoot.attributes.getNamedItem("NO").value

	if trim(No) <> "" then 
		document.formname.hEditNo.value = No
		EditTemplate No 
	end if

End Function

Function EditTemplate(No)
	
	Dim sTemp
	For Each HeaderNode in TemplateRoot.childNodes
		if trim(HeaderNode.attributes.getNamedItem("NO").value) = trim(No) then
			for each Node in HeaderNode.childNodes
				
				
				IF document.formname.hFileName.value <> "GRPTRIALBALANCE" Then
					IF Node.nodeName = "DATE" Then
						ClearSelected document.formname.SelFromMonth
						ClearSelected document.formname.SelToMonth
						for each SubNode in Node.childNodes
							setSelected document.formname.SelFromMonth, SubNode.attributes.getNamedItem("FROMMONTH").value
							setSelected document.formname.SelToMonth, SubNode.attributes.getNamedItem("TOMONTH").value
						Next				
					End IF
				Else
					IF Node.nodeName = "DATE" Then
						ClearSelected document.formname.SelFromMonth
						for each SubNode in Node.childNodes
							setSelected document.formname.SelFromMonth, SubNode.attributes.getNamedItem("TILLMONTH").value							
						Next				
					End IF
				End IF
				if Node.nodeName = "UNIT" then
					'ClearSelected document.formname.selOrg
					'for each SubNode in Node.childNodes
					'	setSelected document.formname.selOrg, SubNode.attributes.getNamedItem("CODE").value
					'Next
					
				elseif Node.nodeName = "LEDGERTYPE" then
						ClearSelected document.formname.SelLedgType
						for each SubNode in Node.childNodes
							setSelected document.formname.SelLedgType, SubNode.attributes.getNamedItem("CODE").value
						Next					
					IF document.formname.hFileName.value = "LEDSelection" Then
						GetLedgType1()						
					End IF
				end if
				
				
				IF document.formname.hFileName.value = "TRIALBALANCE" Then
					if Node.nodeName = "PARTYHEAD" then
						ClearSelected document.formname.SelPartyHead
						for each SubNode in Node.childNodes
							setSelected document.formname.SelPartyHead, SubNode.attributes.getNamedItem("CODE").value
							Window.spPartyHead.innerHTML = SubNode.attributes.getNamedItem("SNAME").value
							document.formname.hdTemp.value = SubNode.attributes.getNamedItem("PARTYSUBTYPE").value
						Next
					end if
				elseif document.formname.hFileName.value = "LEDSelection" Then
						if Node.nodeName = "PARTYHEAD" then
							ClearSelected document.formname.SelPartyHead
							for each SubNode in Node.childNodes
								setSelected document.formname.SelPartyHead, SubNode.attributes.getNamedItem("CODE").value
								Window.spPartyHead.innerHTML = SubNode.attributes.getNamedItem("SNAME").value
								
								Window.spParty.innerHTML = SubNode.attributes.getNamedItem("PARTYNAME").value
								document.formname.hdTemp.value = SubNode.attributes.getNamedItem("PARTYSUBTYPE").value
								IF SubNode.attributes.getNamedItem("PARTYCODE").value <> "" Then
									document.formname.hdTemp1.value = SubNode.attributes.getNamedItem("PARTYCODE").value
									
								Else
									document.formname.hdTemp1.value = 0
								End IF	
							Next
						end if
					
				
				elseif Node.nodeName = "DISPLAYTYPE" then
						
						for each SubNode in Node.childNodes
							if SubNode.attributes.getNamedItem("TYPE").value = "A" Then							
								document.formname.optUnitSel(0).checked = True
								document.formname.selunit.disabled = false
								document.formname.selunit.selectedindex = 0
								document.formname.selunit.disabled = True
							else							
								document.formname.optUnitSel(1).checked = True
								document.formname.selunit.disabled = false
								sTemp = split(SubNode.attributes.getNamedItem("CODE").value,":")
								document.formname.selunit.value = sTemp(1)
							end if
						Next
				
				End IF
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

Function GetLedgType1()
	sLedgType=document.formname.SelLedgType.value
	If sLedgType="CR" or sLedgType="DR" Then
		document.formname.SelPartyHead.length =0
		document.formname.SelPartyHead.length =document.formname.SelPartyHead.length+1
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).text="All Party SubTypes"
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).Value="A"
		document.formname.SelPartyHead.length =document.formname.SelPartyHead.length+1
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).text="Select Party SubType"
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).Value="S"
		'sPartySubtype=0
		'sSubTypeName=""
		'window.spPartyHead.innerHTML =""
		'window.spParty.innerHTML =""
		'sPartyCode =0
		'sPartyName =""
	Elseif sLedgType="GL" Then
		document.formname.SelPartyHead.length =0
		document.formname.SelPartyHead.length =document.formname.SelPartyHead.length+1
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).text="All Account Heads"
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).Value="A"
		document.formname.SelPartyHead.length =document.formname.SelPartyHead.length+1
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).text="Select Account Head"
		document.formname.SelPartyHead.options(document.formname.SelPartyHead.length-1).Value="S"
		document.formname.SelParty.disabled =true
		'window.spPartyHead.innerHTML =""
		'window.spParty.innerHTML =""
	End if
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