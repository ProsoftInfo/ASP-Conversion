dim objTemp,Root,RootO,j,iTotQty,iwcCtr
dim iClass,iItem,sUsage,sOrgID
j = 0
iTotQty = 0

Set Root = Data.documentElement

Function fnInit(sOrg,sClass,sItm,sUsag)
	set objTemp = window.dialogArguments
	Set RootO = objTemp.documentElement

	iClass = sClass
	iItem = sItm
	sUsage = sUsag
	sOrgID = sOrg
	'alert RootO.Xml
	sExp11 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter"
	Set WCNode1 = RootO.Selectnodes(sExp11)
	For iwcCtr = 0 To WCNode1.Length - 1
		sExp12 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter[@WCODE="""&WCNode1.Item(iwcCtr).Attributes.getNamedItem("WCODE").Value&"""]/MachineCenter"
		Set MCNode1 = RootO.Selectnodes(sExp12)

		if MCNode1.Length = 0 and WCNode1.Length > 0 then
			'setIndex document.formname.selWC,WCNode1.Item(0).Attributes.getNamedItem("WCODE").Value
		elseif MCNode1.Length > 0 and WCNode1.Length > 0 then
			for z = 0 to MCNode1.Length - 1
				j = j + 1
				set oRow = document.all.tblData.insertRow(j)

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML=j
				headerCell.className="ExcelSerial"
				headerCell.align="center"

				'Add the check box to delete an entry
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox""  name=""chkDelete"& WCNode1.Item(iwcCtr).Attributes.getNamedItem("WCODE").Value &"Z"& trim(MCNode1.Item(z).Attributes.getNamedItem("MCODE").Value) &""" value=""Y""  class=""FormElem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML= trim(MCNode1.Item(z).Attributes.getNamedItem("NAME").Value)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML= trim(MCNode1.Item(z).Attributes.getNamedItem("QTY").Value)
				headerCell.className="ExcelDisplayCell"
				headerCell.width="10"

				iTotQty = cdbl(iTotQty) + cdbl(MCNode1.Item(z).Attributes.getNamedItem("QTY").Value)

			next		
		end if
	Next
end Function

Function GetWC(obj)
	if (sUsage = "MAT") or (sUsage = "CON")  then
			document.formname.selWC.options.length = 1

			if obj.value ="select" then exit function

			dim Root,HeaderNode

			sWG = document.formname.selWG.value
			set objhttp = CreateObject("MSXML2.XMLHTTP")

			objhttp.Open "GET","XMLWorkCenter.asp?sOrgID="& sOrgID &"&WG="&sWG, false
	
			objhttp.send 

			if objhttp.responseXML.xml <> "" then
				WGData.loadXML objhttp.responseXML.xml
				Set Root = WGData.documentElement
				if Root.HaschildNodes() then
					For Each HeaderNode In Root.childNodes
						document.formname.selWC.length = document.formname.selWC.length+1
						document.formname.selWC.options(document.formname.selWC.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
						document.formname.selWC.options(document.formname.selWC.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
					next
				end if
			else
				alert("No Work Center defined for the Work Group Selected")
				document.formname.selWG.focus()
				Exit Function
			end if
	else
		 exit function
	end if  'if (sUsage = "MAT") or (sUsage = "CON")  then
end Function

Function GetMC(obj)
	if (sUsage = "MAT") or (sUsage = "CON")  then
			document.formname.selMC.options.length = 1

			if obj.value ="select" then exit function

			dim Root,HeaderNode

			sWC = document.formname.selWC.value
			set objhttp = CreateObject("MSXML2.XMLHTTP")

			objhttp.Open "GET","XMLMachineCenter.asp?sOrgID="& sOrgID &"&WC="&sWC, false
	
			objhttp.send 
			'alert objhttp.responseXML.xml
			if objhttp.responseXML.xml <> "" then
				MCData.loadXML objhttp.responseXML.xml
				Set Root = MCData.documentElement
				if Root.HaschildNodes() then
					For Each HeaderNode In Root.childNodes
						document.formname.selMC.length = document.formname.selMC.length+1
						document.formname.selMC.options(document.formname.selMC.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
						document.formname.selMC.options(document.formname.selMC.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
					next
				end if
			else
				alert("No machine Center defined for the Work Center Selected")
				document.formname.selWC.focus()
				Exit Function
			end if
	else
		exit function
	end if 'if (sUsage = "MAT") or (sUsage = "CON")  then
end Function

Function GetMCDetails(obj)

	if obj.value ="select" then exit function

	dim Root,HeaderNode

	sWC = document.formname.selWC.value
	sMC = document.formname.selMC.value
	Set Root = MCData.documentElement
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
		If HeaderNode.Attributes.Item(0).nodeValue = sMC Then
			MCModel.innerHTML = HeaderNode.Attributes.Item(3).nodeValue &"&nbsp;"
			MCSNo.InnerHTML = HeaderNode.Attributes.Item(4).nodeValue &"&nbsp;"
		End If			
		next
	end if

end Function

Function CheckEntry()
	iTotQty = 0

	if document.formname.selWG.selectedIndex = "0" then
		alert("Select Work Group")
		document.formname.selWG.focus()
		exit Function
	elseif document.formname.selWC.selectedIndex = "0" then
		alert("Select Work Center")
		document.formname.selWC.focus()
		exit Function
	elseif document.formname.selMC.selectedIndex = "0" and document.formname.selMC.length > 1 and sUsage<>"CON" then 
		alert("Select Machine Center")
		document.formname.selMC.focus()
		exit Function
	elseif trim(document.formname.txtQty.value) = "" then
		alert("Enter Quantity")
		document.formname.txtQty.select()
		exit Function
	elseif cdbl(document.formname.txtQty.value) <= 0 then
		alert("Quantity cannot be ZERO")
		document.formname.txtQty.select()
		exit Function
	else
		iTotQty = 0
		sWCCode = document.formname.selWC.value
'alert(RootO.xml)
		sExp4 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter/MachineCenter"
		Set MC1Node = RootO.Selectnodes(sExp4)
		if MC1Node.Length > 0 then
			For iCtr = 0 to MC1Node.Length - 1
				iTotQty = cdbl(iTotQty) + cdbl(MC1Node.Item(iCtr).Attributes.getNamedItem("QTY").Value)
			next
			iTotQty = cdbl(iTotQty) + cdbl(document.formname.txtQty.value)
		else
			iTotQty = cdbl(document.formname.txtQty.value)
		end if

		if (cdbl(iTotQty) > cdbl(idQty.innerText)) then
			alert("Quantity breakup should be equal to Quantity Issue")
			exit function
		end if

		sEx ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet"
		Set ADNode = RootO.Selectnodes(sEx)
		
		sExp ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter"
		Set WCNode = RootO.Selectnodes(sExp)

		sExp1 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]"
		Set HeaderNode = RootO.Selectnodes(sExp1)
		' alert HeaderNode.Length
		if WCNode.Length > 0 then
			sExp2 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter [ @WCODE = '"&sWCCode&"']"
			Set WC1Node = RootO.Selectnodes(sExp2)
			
			if WC1Node.Length > 0 then
				WC1Node.Item(0).Attributes.getNamedItem("WCODE").Value = sWCCode
			else
				Set newElem2 = objTemp.createElement("WorkCenter")
				newElem2.setAttribute "WCODE", document.formname.selWC.value

				ADNode.Item(0).appendChild newElem2
				
			end if

			sExp2 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter [ @WCODE = '"&sWCCode&"']"
			Set WC1Node = RootO.Selectnodes(sExp2)

			sExp3 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter [ @WCODE = '"&sWCCode&"']/MachineCenter [ @MCODE = '"&document.formname.selMC.value&"']"
			Set MCNode = RootO.Selectnodes(sExp3)
			if MCNode.Length > 0 then
				ClearTable
				MCNode.Item(0).Attributes.getNamedItem("MCODE").Value = document.formname.selMC.value
				MCNode.Item(0).Attributes.getNamedItem("QTY").Value = trim(document.formname.txtQty.value)
			
				if document.formname.selMC.selectedIndex > 0 then
					MCNode.Item(0).Attributes.getNamedItem("NAME").Value = document.formname.selWC(document.formname.selWC.selectedIndex).text & " / " & document.formname.selMC(document.formname.selMC.selectedIndex).text & " (" & MCModel.InnerHTML & "-" & MCSNo.InnerHTML & ")"
				else
					MCNode.Item(0).Attributes.getNamedItem("NAME").Value = document.formname.selWC(document.formname.selWC.selectedIndex).text
				end if

				for z = 0 to MC1Node.Length - 1
					j = j + 1
					set oRow = document.all.tblData.insertRow(j)

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML=j
					headerCell.className="ExcelSerial"
					headerCell.align="center"
					
					'Add the check box to delete an entry
					set headerCell=oRow.insertCell()
					set oText = document.createElement("<input type=""checkbox""  name=""chkDelete"& document.formname.selWC(document.formname.selWC.selectedIndex).value &"Z"& document.formname.selMC(document.formname.selMC.selectedIndex).value &""" value=""Y""  class=""FormElem"">")
					headerCell.appendChild(oText)
					headerCell.width = "10"
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
			
					set headerCell=oRow.insertCell()									
					headerCell.innerHTML= trim(MC1Node.Item(z).Attributes.getNamedItem("NAME").Value)
					headerCell.className="ExcelDisplayCell"
					headerCell.align="left"

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML= trim(MC1Node.Item(z).Attributes.getNamedItem("QTY").Value)
					headerCell.className="ExcelDisplayCell"
					headerCell.width="10"

					iTotQty = cdbl(iTotQty) + cdbl(MC1Node.Item(z).Attributes.getNamedItem("QTY").Value)

				next		

			else
				Set newElem3 = objTemp.createElement("MachineCenter")
				newElem3.setAttribute "MCODE", document.formname.selMC.value
				newElem3.setAttribute "QTY", trim(document.formname.txtQty.value)
				if document.formname.selMC.selectedIndex > 0 then
					newElem3.setAttribute "NAME", document.formname.selWC(document.formname.selWC.selectedIndex).text & " / " & document.formname.selMC(document.formname.selMC.selectedIndex).text & " (" & MCModel.InnerHTML & "-" & MCSNo.InnerHTML & ")"
				else
					newElem3.setAttribute "NAME", document.formname.selWC(document.formname.selWC.selectedIndex).text
				end if
				
				WC1Node.Item(0).appendChild newElem3

				j = j + 1
				set oRow = document.all.tblData.insertRow(j)

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML=j
				headerCell.className="ExcelSerial"
				headerCell.align="center"

				'Add the check box to delete an entry
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox""  name=""chkDelete"& document.formname.selWC(document.formname.selWC.selectedIndex).value &"Z"& document.formname.selMC(document.formname.selMC.selectedIndex).value &""" value=""Y""  class=""FormElem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"
			
				set headerCell=oRow.insertCell()									
				if document.formname.selMC.selectedIndex > 0 then
					headerCell.innerHTML= document.formname.selWC(document.formname.selWC.selectedIndex).text & " / " & document.formname.selMC(document.formname.selMC.selectedIndex).text & " (" & MCModel.InnerHTML & "-" & MCSNo.InnerHTML & ")"
				else
					headerCell.innerHTML= document.formname.selWC(document.formname.selWC.selectedIndex).text
				end if
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML= trim(document.formname.txtQty.value)
				headerCell.className="ExcelDisplayCell"
				headerCell.width="10"

			end if

		else
		
			Set newElem1 = objTemp.createElement("AddDet")
			Set newElem2 = objTemp.createElement("WorkCenter")
			newElem2.setAttribute "WCODE", document.formname.selWC.value

			Set newElem3 = objTemp.createElement("MachineCenter")
			newElem3.setAttribute "MCODE", document.formname.selMC.value
			newElem3.setAttribute "QTY", trim(document.formname.txtQty.value)
		
			if document.formname.selMC.selectedIndex > 0 then
				newElem3.setAttribute "NAME", document.formname.selWC(document.formname.selWC.selectedIndex).text & " / " & document.formname.selMC(document.formname.selMC.selectedIndex).text & " (" & MCModel.InnerHTML & "-" & MCSNo.InnerHTML & ")"
			else
				newElem3.setAttribute "NAME", document.formname.selWC(document.formname.selWC.selectedIndex).text
			end if
			newElem2.appendChild newElem3
			
			newElem1.appendChild newElem2
			HeaderNode.Item(0).appendChild newElem1

			j = j + 1
			set oRow = document.all.tblData.insertRow(j)

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML=j
			headerCell.className="ExcelSerial"
			headerCell.align="center"
			
			'Add the check box to delete an entry
			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""checkbox""  name=""chkDelete"& document.formname.selWC(document.formname.selWC.selectedIndex).value &"Z"& document.formname.selMC(document.formname.selMC.selectedIndex).value &""" value=""Y""  class=""FormElem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()									
			if document.formname.selMC.selectedIndex > 0 then
				headerCell.innerHTML= document.formname.selWC(document.formname.selWC.selectedIndex).text & " / " & document.formname.selMC(document.formname.selMC.selectedIndex).text & " (" & MCModel.InnerHTML & "-" & MCSNo.InnerHTML & ")"
			else
				headerCell.innerHTML= document.formname.selWC(document.formname.selWC.selectedIndex).text
			end if
			headerCell.className="ExcelDisplayCell"
			headerCell.align="left"

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML= trim(document.formname.txtQty.value)
			headerCell.className="ExcelDisplayCell"
			headerCell.width="10"

		end if

		document.formname.selWC.selectedIndex = "0"
		document.formname.selMC.selectedIndex = "0"
		document.formname.txtQty.value = ""

	end if
end Function

Function CheckSubmit()
	if document.formname.selWC.selectedIndex = "0" and sUsage = "PRD" then
		alert("Select Work Center")
		document.formname.selWC.focus()
		exit Function
	end if

	Set Root = Data.documentElement

	sExp ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter"
	Set WCNode = RootO.Selectnodes(sExp)

	sExp1 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]"
	Set HeaderNode = RootO.Selectnodes(sExp1)


	if sUsage = "MAT" or sUsage = "CON" then 
		'if (cdbl(iTotQty) <> cdbl(idQty.innerText)) then
		if (cdbl(iTotQty) > cdbl(idQty.innerText)) then
			alert("Enter Full Breakup Quantity")
			exit function
		end if
	else
		if WCNode.Length > 0 then
			WCNode.Item(0).Attributes.getNamedItem("WCODE").Value = document.formname.selWC.value
		else
			Set newElem1 = objTemp.createElement("AddDet")
			Set newElem2 = objTemp.createElement("WorkCenter")
			newElem2.setAttribute "WCODE", document.formname.selWC.value

			newElem1.appendChild newElem2
			HeaderNode.Item(0).appendChild newElem1
		end if
	end if
	window.close
end Function

Function setIndex(objSch,sTemp)
	for ic = 0 to objSch.length - 1
		if sTemp = objSch.options(ic).value then
			objSch.selectedIndex = ic
			exit function
		end if
	next
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

Function ClearTable()
	dim i
	for i=1 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(1) 
	next
	j = 0
	iTotQty = 0
end Function

Function DeleteEntry()

	dim sExp,AddDetNode,WCNode,MCNode,itr,iWctr,sWCDel,sMCDel,objSel,objSer,isSelected,ItemNode
	isSelected = False
	'alert(RootO.XML)

	sExp ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter"
	Set WCNode = RootO.Selectnodes(sExp)
	
	If WCNode.Length > 0 Then
		For iWctr = 0 To WCNode.Length - 1
			sWCDel = WCNode.item(iWctr).Attributes.getNamedItem("WCODE").value
			sExp ="//ITEMDETAILS[ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet/WorkCenter[ @WCODE="""&sWCDel&"""]/MachineCenter"
		
			Set MCNode = RootO.Selectnodes(sExp)
			if MCNode.Length > 0 then
				for itr = 0 to MCNode.Length - 1
					sMCDel = MCNode.item(itr).Attributes.getNamedItem("MCODE").value
					set objSel = eval("document.formname.chkDelete"&CStr(sWCDel)&"Z"&CStr(sMCDel))
					'set objSer = eval("document.formname.txtSerA"&CStr(iClassDel)&"A"&CStr(iItemDel))
					'alert(objSel.value)
					if objSel.checked then 'DeleteItem objSel,objSer.value + 2
						'alert("wc="&sWCDel)
						'alert("MC="&sMCDel)
						'Remove the node
						Set ItemNode = WCNode.Item(iWctr).removeChild(MCNode.item(itr))
						isSelected = True
						If Not WCNode.Item(iWctr).HasChildNodes then
							'Remove Work Center node
							sExp = "//ITEMDETAILS[ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet"
							Set AddDetNode = RootO.selectNodes(sExp)
							Set ChkNode = AddDetNode.Item(0).removeChild(WCNode.Item(iWctr))
							If Not AddDetNode.Item(0).HasChildNodes Then
								'Remove AddDet node
								sExp = "//ITEMDETAILS[ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]"
								Set iTemNode = RootO.selectNodes(sExp)
								Set ChkNode = iTemNode.Item(0).removeChild(AddDetNode.Item(0))
							End If
						End If
					End If
				next
			end if
		Next
	End If
	If Not isSelected Then
		alert("Please select an entry to delete")
		Exit Function
	Else
		ClearTable
		fnInit sOrgID,iClass,iItem,sUsage
	End If
End Function
