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
'================================================
Function CheckEntry()
Dim AttNode,sTempNode
	if (document.formname.txtValue.value = "") then
		MsgBox "Enter Option Name",0,"Option Name"
		document.formname.txtValue.select()
		Exit Function
	elseif(not CheckName(Trim(document.formname.txtValue.value))) then
		MsgBox "Option Name already entered",0,"Option Name"
		document.formname.txtValue.focus()
		Exit Function
	else
		iSno = iSno + 1

		Set Root = OutDataValue.documentElement
		For each AttNode in Root.childNodes
			if strcomp(AttNode.nodeName,"Attribute")=0 then
				set sTempNode= AttNode 
			end if
		next
		
		Set newElem = OutDataValue.createElement("Option")
		newElem.setAttribute "Type",""
		newElem.setAttribute "Name", trim(document.formname.txtValue.value)
		sTempNode.appendChild newElem

		document.formname.txtValue.value = ""
	end if
ClearTable()
DisplayTable()
End Function
'========================================================
Function CheckSubmit()
		window.close
end Function
'===========================================================
Function window_onunload() 
	set window.returnValue = OutDataValue.documentElement
end Function
'===========================================================
Function Init()
	Dim Root,AttNode,sAttributeNode
	Dim sItmType,sHeader,sAttName,iSno
	set Root = objTemp.documentElement

	sItmType =  document.formname.hItemType.value
	sHeader  =	document.formname.hHeader.value
	sAttName =  document.formname.hAttribute.value
	iSno= 0
	for each AttNode in Root.childNodes
		if strcomp(AttNode.getAttribute("Name"),sAttName)=0 and strcomp(AttNode.getAttribute("HName"),sHeader)=0 then
			set sAttributeNode = AttNode
			RootNode.appendChild(sAttributeNode)
			for each OptNode in AttNode.childNodes
					iSno = iSno + 1
				set oRow = document.all.tblData.insertRow(document.all.tblData.Rows.Length)
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = iSno 
				CurrCell.ClassName="ExcelHeaderCell"
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=checkbox name='chkOptZ"&iSno&"' value='"& OptNode.getAttribute("Name") &"'>")
				CurrCell.appendChild(oText)
				CurrCell.ClassName="ExcelDisplayCell"
				
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = OptNode.getAttribute("Name")
				CurrCell.ClassName="ExcelDisplayCell"
			next
		end if
	next
	document.formname.hRow.value = iSno
End Function 
'================================================
Function DelItem()
Dim iRow,AttNode,OptNode

for iRow = 1 to cint(document.formname.hRow.value)
	if eval("document.formname.chkOptZ"&iRow).checked = true then
		for each AttNode in RootNode.childNodes
			for each OptNode in AttNode.childNodes
				if strcomp(OptNode.nodeName,"Option")=0 then
					if strcomp(OptNode.getAttribute("Name"),eval("document.formname.chkOptZ"&iRow).value)=0 then
						AttNode.removechild(OptNode)
					end if
				end if
			next
		next
	end if
next
ClearTable 
DisplayTable	
End Function
'====================================================================
Function ClearTable()
	Dim iRow
	for iRow =1 to cint(document.formname.hRow.value)
		document.all.tblData.deleteRow(1)
	next
End Function
'=================================================================
Function DisplayTable()
Dim	AttNode,iSno
	iSno =0
	
	for each AttNode in RootNode.childNodes
		for each OptNode in AttNode.childNodes
				iSno = iSno + 1
				set oRow = document.all.tblData.insertRow(iSno)

						set headerCell=oRow.insertCell()									
						headerCell.innerHTML = iSno
						headerCell.className="ExcelSerial"
						headerCell.align="center"
						
						set CurrCell = oRow.insertCell()
						set oText = document.createElement("<input type=checkbox name='chkOptZ"&iSno&"' value='"& OptNode.getAttribute("Name") &"'>")
						CurrCell.appendChild(oText)
						CurrCell.ClassName="ExcelDisplayCell"

						set headerCell=oRow.insertCell()
						headerCell.innerHTML= OptNode.getAttribute("Name")
						headerCell.className="ExcelDisplayCell"
						headerCell.align="left"
		next
	next
	document.formname.hRow.value = iSno
End Function
'=====================================================================

