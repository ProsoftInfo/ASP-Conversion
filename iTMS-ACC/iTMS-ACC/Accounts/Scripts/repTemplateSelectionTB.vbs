dim objTemp
Dim sLedgCode,sLedgName,sType,sPartySubType
set objTemp = window.dialogArguments
set Root = objTemp.DocumentElement

Function Init()
	
	dim Node,i

	for i = 1 to document.all.tblDetails.rows.length - 1
		document.all.tblDetails.deleteRow(1)
	next
	
	i = 0
	if Root.hasChildNodes then
		For Each Node in Root.childNodes
			i=i+1
			set oRow = document.all.tblDetails.insertRow(i)

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=i
			headerCell.className="ExcelSerial"
			headerCell.align="center"

			set headerCell=oRow.insertCell()									
			set oLink = document.createElement("<a href=""#"" onClick=""ShowDetails("&i&")"">")
			oLink.innerText=Node.attributes.getNamedItem("NAME").value
			oLink.className="ExcelDisplayLink"
			headerCell.appendChild(oLink)
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"
			
			IF document.formname.hFileName.value = "TRIALBALANCE" Then
			
			set headerCell=oRow.insertCell()
			set oButton = document.createElement("<input type=""button"" name =BtnView:" & i & " value=""View Closing"" style=""WIDTH:80px"" onClick=""ViewTemplate(" & i & ")"" class=""AddButton"">")
			headerCell.appendChild oButton
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"
			
			set headerCell=oRow.insertCell()
			set oButton = document.createElement("<input type=""button"" name =BtnView:" & i & " value=""View Complete"" style=""WIDTH:90px"" onClick=""ViewTemplate1(" & i & ")"" class=""AddButton"">")
			headerCell.appendChild oButton
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"
			
			Else
			
			set headerCell=oRow.insertCell()
			set oButton = document.createElement("<input type=""button"" name =BtnView:" & i & " value=""View"" style=""WIDTH:50px"" onClick=""ViewTemplate(" & i & ")"" class=""AddButton"">")
			headerCell.appendChild oButton
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"
			
			End IF
			
			set headerCell=oRow.insertCell()
			set oButton = document.createElement("<input type=""button"" name =BtnEdit:" & i & " value=""Edit"" style=""WIDTH:50px"" onClick=""EditTemplate(" & i & ")"" class=""AddButton"">")
			headerCell.appendChild oButton
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oButton = document.createElement("<input type=""button"" name =BtnDelete:" & i & " value=""Delete"" style=""WIDTH:50px"" onClick=""DeleteTemplate(" & i & ")"" class=""AddButton"">")
			headerCell.appendChild oButton
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"
		next
	else
		set oRow = document.all.tblDetails.insertRow(1)

		set headerCell=oRow.insertCell()
		headerCell.innerHTML = "No Templates Defined"
		headerCell.colspan="5"
		headerCell.className="ExcelDisplayCell"
		headerCell.align="center"

	end if
End Function

Function ViewTemplate(No)
	
	dim HeaderNode,Node,SubNode,sValue
	dim i,sUnit,sUnitName,sPassStr,sSelUnit,sFromMonth,sToMonth,sTemp,sPartyCode,sPartyName,sAcc
	
	Root.setAttribute "NO", ""
	For Each HeaderNode in Root.childNodes
		if trim(HeaderNode.attributes.getNamedItem("NO").value) = trim(No) then
			for each Node in HeaderNode.childNodes
				
				if Node.nodeName = "UNIT" then
					for each SubNode in Node.childNodes
						sUnit = SubNode.attributes.getNamedItem("CODE").value
						sUnitName = SubNode.attributes.getNamedItem("NAME").value
					Next					
				elseif Node.nodeName = "LEDGERTYPE" then
					for each SubNode in Node.childNodes
						sLedgCode = SubNode.attributes.getNamedItem("CODE").value
						sLedgName = SubNode.attributes.getNamedItem("NAME").value
					Next
				end if	
				IF document.formname.hFileName.value = "GRPTRIALBALANCE" Then
					if Node.nodeName = "DISPLAYTYPE" then
						for each SubNode in Node.childNodes
							sSelUnit = SubNode.attributes.getNamedItem("CODE").value
							sFromMonth = SubNode.attributes.getNamedItem("FROMMONTH").value
						Next	
					end if
				ElseIF document.formname.hFileName.value="LEDSelection" Then
					if Node.nodeName = "PARTYHEAD" then						
						for each SubNode in Node.childNodes										
							sFromMonth = SubNode.attributes.getNamedItem("FROMMONTH").value
							sToMonth = SubNode.attributes.getNamedItem("TOMONTH").value
							sType = SubNode.attributes.getNamedItem("NAME").value
							sPartySubType = SubNode.attributes.getNamedItem("PARTYSUBTYPE").value
							IF SubNode.attributes.getNamedItem("PARTYCODE").value <> "" Then
								sPartyCode = SubNode.attributes.getNamedItem("PARTYCODE").value
							Else
								sPartyCode = 0
							End IF
							sPartyName = SubNode.attributes.getNamedItem("PARTYNAME").value
							sAcc = SubNode.attributes.getNamedItem("ACC").value
						next
					end if
				ElseIF document.formname.hFileName.value="TRIALBALANCE" Then
					if Node.nodeName = "PARTYHEAD" then						
						for each SubNode in Node.childNodes										
							sFromMonth = SubNode.attributes.getNamedItem("FROMMONTH").value
							sToMonth = SubNode.attributes.getNamedItem("TOMONTH").value
							sType = SubNode.attributes.getNamedItem("NAME").value
							sPartySubType = SubNode.attributes.getNamedItem("PARTYSUBTYPE").value							
						next
					end if
				End IF
			next
			
			'Adding Query string values based on Report
			
			IF document.formname.hFileName.value = "LEDSelection" Then
				IF sLedgCode = "GL" Then
					sTemp = split(sAcc,":")
					sPassStr1 = "LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth&"&AccHead="&sTemp(0)&"&AccName="&sTemp(1)
				ElseIF Trim(sType) = "All Party SubTypes" Then
					sPartySubType = 0
					sPassStr1 = "PartySubtype="&sPartySubtype&"&LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth				
				Else
					sPassStr1="PartySubtype="&sPartySubtype&"&LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth&"&SubTypeName="&sSubTypeName&"&PartyCode="&sPartycode&"&PartyName="&sPartyName
				End IF
				
			End IF
			
			Select Case document.formname.hFileName.value			
				Case "GRPTRIALBALANCE"   	: sPassStr="LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&DispTy="&sSelUnit
				Case "TRIALBALANCE"	   	: sPassStr="LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth&"&CallTy=C"
				Case "LEDSelection"		: sPassStr=sPassStr1
				Case Else			: alert "No Reports to view" : exit function
			End Select
			
			exit for
		end if
	next

	ShowReport(sPassStr)
End Function

Function ShowReport(sPassStr)
	Select Case document.formname.hFileName.value
		Case "GRPTRIALBALANCE" 		:
		IF sLedgCode = "GL" Then
			open "CompTBGLView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
		Else
			open "CompTBPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"	
		End IF
		
		Case "TRIALBALANCE" 		:
		IF sLedgCode = "GL" Then
			open "TBGLView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
		Else
			
			if Trim(sType) = "All Party Type" then
				sPassStr = "PartySubtype=A&"&sPassStr						
				open "TBPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			else					
				sPassStr = "PartySubtype="&sPartySubType&"&"&sPassStr				
				open "TBPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			end if
		End IF
		
		Case "LEDSelection" 		:
		IF sLedgCode = "GL" Then
			open "LEDGLView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"

		Else
			
			if Trim(sType) = "All Party SubTypes" then

				open "LEDPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			else				
				open "LEDPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			end if
		End IF
		
	End Select
End Function

Function ViewTemplate1(No)
	
	dim HeaderNode,Node,SubNode,sValue
	dim i,sUnit,sUnitName,sPassStr,sSelUnit,sFromMonth,sToMonth,sTemp,sPartyCode,sPartyName,sAcc
	
	Root.setAttribute "NO", ""
	For Each HeaderNode in Root.childNodes
		if trim(HeaderNode.attributes.getNamedItem("NO").value) = trim(No) then
			for each Node in HeaderNode.childNodes
				
				if Node.nodeName = "UNIT" then
					for each SubNode in Node.childNodes
						sUnit = SubNode.attributes.getNamedItem("CODE").value
						sUnitName = SubNode.attributes.getNamedItem("NAME").value
					Next					
				elseif Node.nodeName = "LEDGERTYPE" then
					for each SubNode in Node.childNodes
						sLedgCode = SubNode.attributes.getNamedItem("CODE").value
						sLedgName = SubNode.attributes.getNamedItem("NAME").value
					Next
				end if	
				IF document.formname.hFileName.value = "GRPTRIALBALANCE" Then
					if Node.nodeName = "DISPLAYTYPE" then
						for each SubNode in Node.childNodes
							sSelUnit = SubNode.attributes.getNamedItem("CODE").value
							sFromMonth = SubNode.attributes.getNamedItem("FROMMONTH").value
						Next	
					end if
				ElseIF document.formname.hFileName.value="LEDSelection" Then
					if Node.nodeName = "PARTYHEAD" then						
						for each SubNode in Node.childNodes										
							sFromMonth = SubNode.attributes.getNamedItem("FROMMONTH").value
							sToMonth = SubNode.attributes.getNamedItem("TOMONTH").value
							sType = SubNode.attributes.getNamedItem("NAME").value
							sPartySubType = SubNode.attributes.getNamedItem("PARTYSUBTYPE").value
							sPartyCode = SubNode.attributes.getNamedItem("PARTYCODE").value
							sPartyName = SubNode.attributes.getNamedItem("PARTYNAME").value
							sAcc = SubNode.attributes.getNamedItem("ACC").value
						next
					end if
				ElseIF document.formname.hFileName.value="TRIALBALANCE" Then
					if Node.nodeName = "PARTYHEAD" then						
						for each SubNode in Node.childNodes										
							sFromMonth = SubNode.attributes.getNamedItem("FROMMONTH").value
							sToMonth = SubNode.attributes.getNamedItem("TOMONTH").value
							sType = SubNode.attributes.getNamedItem("NAME").value
							sPartySubType = SubNode.attributes.getNamedItem("PARTYSUBTYPE").value							
						next
					end if
				End IF
			next
			
			'Adding Query string values based on Report
			
			IF document.formname.hFileName.value = "LEDSelection" Then
				IF sLedgCode = "GL" Then
					sTemp = split(sAcc,":")
					sPassStr1 = "LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth&"&AccHead="&sTemp(0)&"&AccName="&sTemp(1)
				ElseIF Trim(sType) = "All Party SubTypes" Then
					sPartySubType = 0
					sPassStr1 = "PartySubtype="&sPartySubtype&"&LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth				
				Else
					sPassStr1="PartySubtype="&sPartySubtype&"&LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth&"&SubTypeName="&sSubTypeName&"&PartyCode="&sPartycode&"&PartyName="&sPartyName
				End IF
				
			End IF
			
			Select Case document.formname.hFileName.value			
				Case "GRPTRIALBALANCE"   	: sPassStr="LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&DispTy="&sSelUnit
				Case "TRIALBALANCE"	   	: sPassStr="LedgType="&sLedgCode&"&OrgId="&sUnit&"&OrgName="&sUnitName&"&FromMonth="&sFromMonth&"&ToMonth="&sToMonth
				Case "LEDSelection"		: sPassStr=sPassStr1
				Case Else			: alert "No Reports to view" : exit function
			End Select
			
			exit for
		end if
	next

	ShowReport1(sPassStr)
End Function

Function ShowReport1(sPassStr)
	Select Case document.formname.hFileName.value
		Case "GRPTRIALBALANCE" 		:
		IF sLedgCode = "GL" Then
			open "CompTBGLView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
		Else
			open "CompTBPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"	
		End IF
		
		Case "TRIALBALANCE" 		:
		IF sLedgCode = "GL" Then
			open "TBGLView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
		Else
			
			if Trim(sType) = "All Party Type" then
				sPassStr = "PartySubtype=A&"&sPassStr						
				open "TBPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			else					
				sPassStr = "PartySubtype="&sPartySubType&"&"&sPassStr				
				open "TBPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			end if
		End IF
		
		Case "LEDSelection" 		:
		IF sLedgCode = "GL" Then
			open "LEDGLView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"

		Else
			
			if Trim(sType) = "All Party SubTypes" then

				open "LEDPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			else				
				open "LEDPartyView.asp?"&sPassStr,"A","height=590,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No"
			end if
		End IF
		
	End Select
End Function


Function EditTemplate(No)
	Root.setAttribute "NO", No
	window.close 
End Function

Function window_onunload(No)
	set window.returnValue = objTemp.DocumentElement
End Function

Function DeleteTemplate(No)
	dim Node,i

	For Each Node in Root.childNodes
		if trim(Node.attributes.getNamedItem("NO").value) = trim(No) then
			Root.removeChild Node
			exit for
		end if
	next

	For Each Node in Root.childNodes
		i = i + 1
		Node.attributes.getNamedItem("NO").value = i
	next

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objHttp.open "POST","..\Reports\XMLSave.asp?UserId=True&Value="&document.formname.hFileName.value&"&Folder=..\XmlData\Reports",false
	objhttp.send objTemp.XMLDocument

	Init()
End Function

Function ShowDetails(No)	
	showModalDialog "..\Reports\TemplateDisplayTB.asp?FILENAME="&document.formname.hFileName.value&"&NO="&No,"","dialogHeight:270px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No"
End Function