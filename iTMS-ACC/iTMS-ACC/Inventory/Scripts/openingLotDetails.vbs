
dim RowNo, objTemp, Root, NewElem, sType, iItem, iClass, sOrgID, iTotQty, sStoName
dim sStoresUom, iQty, sCheck, iNo, sAltUom, sAltCheck, iAltGross,iAltNett
dim iTotGross, iTotNett

iTotGross = 0
iTotNett = 0

Function Init(sTemp,sTempCheck,sTempAltCheck)
	dim HeaderNode, Node, arrTemp

	arrTemp = split(sTemp ,":")
	sType	= arrTemp(0)
	iItem	= arrTemp(1)
	iClass	= arrTemp(2)
	sOrgID	= arrTemp(3)
	iTotQty = arrTemp(6)
	sStoName = arrTemp(7)
	sStoresUom = arrTemp(8)
	iQty = arrTemp(10)
	iNo = arrTemp(11)
	sAltUom = arrTemp(12)
	iAltGross = arrTemp(13)
	iAltNett = arrTemp(14)
	
	if iAltGross = "" then iAltGross = "0"
	if iAltNett = "" then iAltNett = 0

	if trim(sAltUom) = "select" then sAltUom = "-"	
	sCheck = sTempCheck
	sAltCheck = sTempAltCheck
	
	Set objTemp= Window.DialogArguments
	Set Root = objTemp.documentElement
	
	if Root.hasChildNodes() then
		for Each HeaderNode in Root.childNodes
			
			if HeaderNode.attributes.getNamedItem("NO").value = iNo then
				if trim(HeaderNode.Attributes.getNamedItem("CHECK").value) = "Y" then
					document.formname.chkAltUom.Checked = True
					document.formname.txtGross.disabled = False
					document.formname.txtNett.disabled = False
					document.formname.txtGross.value = HeaderNode.Attributes.getNamedItem("ALTGROSS").value
					document.formname.txtNett.value = HeaderNode.Attributes.getNamedItem("ALTNETT").value
				end if
				
				for each Node in HeaderNode.childNodes
					set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
					RowNo = RowNo +1

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML=RowNo
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"

'					set headerCell=oRow.insertCell()									
'					headerCell.innerHTML=Node.attributes.getNamedItem("QUANTITY").value
'					headerCell.className="ExcelDisplayCell"
'					headerCell.align="right"
'
					set headerCell=oRow.insertCell()									
					set oText = document.createElement("<input type=""text"" name=""txtDetails"&CStr(RowNo)&""" size=""12"" value="""&Node.attributes.getNamedItem("QUANTITY").value&""" class=""Formelem"" onkeypress=""DoKeyPress('"&sCheck&"',7,1)"" onBlur=""CalculateQty()"" style=""text-align=right"">" )
					headerCell.appendChild(oText)
					headerCell.width = "8"
					headerCell.className="ExcelInputCell"

					set headerCell=oRow.insertCell()									
					set oText = document.createElement("<input type=""text"" name=""txtAltGross"&CStr(RowNo)&""" size=""12"" value="""&Node.attributes.getNamedItem("ALTGROSS").value&""" class=""Formelem"" onkeypress=""DoKeyPress('"&sAltCheck&"',7,1)"" onBlur=""CheckGrossQty(this)"" style=""text-align=right"">" )
					headerCell.appendChild(oText)
					headerCell.width = "8"
					headerCell.className="ExcelInputCell"

					set headerCell=oRow.insertCell()									
					set oText = document.createElement("<input type=""text"" name=""txtAltNett"&CStr(RowNo)&""" size=""12"" value="""&Node.attributes.getNamedItem("ALTNETT").value&""" class=""Formelem"" onkeypress=""DoKeyPress('"&sAltCheck&"',7,1)"" onBlur=""CheckNettQty(this)"" style=""text-align=right"">" )
					headerCell.appendChild(oText)
					headerCell.width = "8"
					headerCell.className="ExcelInputCell"

					idQtyEntered.innerText = cdbl(idQtyEntered.innerText) + cdbl(Node.attributes.getNamedItem("QUANTITY").value)
					
				next
			end if
		next
		EnableDone()
	end if
'	alert Root.xml	
End Function

Function AddRow()
	
	if trim(document.formname.txtQty.value) = "" then
		alert "Enter Quantity"
		document.formname.txtQty.focus
		exit function
	elseif document.formname.txtQty.value = 0 then
		alert "Quantity should be greater than zero"
		document.formname.txtQty.select
		exit function
	end if
	
	if CheckQty() then
		alert "Entered Quantity should be less than or equal to Quantity " & cdbl(idQty.innerText)
		document.formname.txtQty.select
		exit function
	end if
	
	set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
	RowNo = RowNo +1

	set headerCell=oRow.insertCell()									
	headerCell.innerHTML=RowNo
	headerCell.className="ExcelDisplayCell"
	headerCell.align="center"

'	set headerCell=oRow.insertCell()									
'	headerCell.innerHTML=document.formname.txtQty.value
'	headerCell.className="ExcelDisplayCell"
'	headerCell.align="right"

	set headerCell=oRow.insertCell()									
	set oText = document.createElement("<input type=""text"" name=""txtDetails"&CStr(RowNo)&""" size=""12"" value="""&document.formname.txtQty.value&""" class=""Formelem"" onkeypress=""DoKeyPress('"&sCheck&"',7,1)"" onBlur=""CalculateQty()"" style=""text-align=right"">" )
	headerCell.appendChild(oText)
	headerCell.width = "8"
	headerCell.className="ExcelInputCell"

	set headerCell=oRow.insertCell()									
	set oText = document.createElement("<input type=""text"" name=""txtAltGross"&CStr(RowNo)&""" size=""12""  class=""Formelem"" onkeypress=""DoKeyPress('"&sAltCheck&"',7,1)"" onBlur=""CheckGrossQty(this)"" style=""text-align=right"">" )
	headerCell.appendChild(oText)
	headerCell.width = "8"
	headerCell.className="ExcelInputCell"

	set headerCell=oRow.insertCell()									
	set oText = document.createElement("<input type=""text"" name=""txtAltNett"&CStr(RowNo)&""" size=""12""  class=""Formelem"" onkeypress=""DoKeyPress('"&sAltCheck&"',7,1)"" onBlur=""CheckNettQty(this)"" style=""text-align=right"">" )
	headerCell.appendChild(oText)
	headerCell.width = "8"
	headerCell.className="ExcelInputCell"

	idQtyEntered.innerText = cdbl(idQtyEntered.innerText) + cdbl(document.formname.txtQty.value)

	document.formname.txtQty.value = ""
	document.formname.txtQty.focus

	EnableDone()	
End Function

Function CheckQty()
	if cdbl(idQty.innerText) < cdbl(idQtyEntered.innerText) + cdbl(document.formname.txtQty.value) then
		CheckQty = True
	else
		CheckQty = False
	end if
End Function

Function window_onunload() 
	
	if cdbl(idQty.innerText) <> cdbl(idQtyEntered.innerText) then
		for Each HeaderNode in Root.childNodes
			if HeaderNode.attributes.getNamedItem("NO").value = iNo then
				for each Node in HeaderNode.childNodes
					HeaderNode.removechild Node
				next
			end if
		next
	else
		for Each HeaderNode in Root.childNodes
			if HeaderNode.attributes.getNamedItem("NO").value = iNo then
				HeaderNode.setAttribute "QUANTITY", cdbl(idQtyEntered.innerText)
			end if
		next
	end if

'	alert objTemp.xml	
	set window.returnValue = objTemp.documentElement
	window.close()
End Function

Function CheckSubmit() 
	dim i,objDetails,objGross,objNett, iTempGross, iTempNett

	if document.formname.chkAltUom.Checked then				
		if trim(document.formname.txtGross.value) = "" then
			alert "Enter Gross in " & sAltUom 
			document.formname.txtGross.focus
			exit function
		elseif trim(document.formname.txtNett.value) = "" then
			alert "Enter Nett in " & sAltUom 
			document.formname.txtNett.focus
			exit function
		end if
	elseif cdbl(iAltGross) > 0 then
		if CheckGrossEntry() then
			alert "Enter Gross in " & sAltUom 
			exit function
		elseif CheckNettEntry() then
			alert "Enter Nett in " & sAltUom 
			exit function
		end if
		
	end if
	
	if cdbl(idQty.innerText) = cdbl(idQtyEntered.innerText) then
		'Removing Previous Node
		for Each HeaderNode in Root.childNodes
			if HeaderNode.attributes.getNamedItem("NO").value = iNo then
				for each Node in HeaderNode.childNodes
					HeaderNode.removechild Node
				next
			end if
		next
		
		iTempGross = 0
		iTempNett = 0
		
		'Adding to XML
		for Each HeaderNode in Root.childNodes
			if HeaderNode.attributes.getNamedItem("NO").value = iNo then
				
				for i = 1 to RowNo
					set objDetails = eval("document.formname.txtDetails"&cstr(i))
					set objGross = eval("document.formname.txtAltGross"&cstr(i))
					set objNett = eval("document.formname.txtAltNett"&cstr(i))
		
					set NewElem = objTemp.createElement("DETAILS")
					NewElem.setAttribute "PIECENO", i
					NewElem.setAttribute "QUANTITY", objDetails.value
					NewElem.setAttribute "ALTGROSS", objGross.value
					NewElem.setAttribute "ALTNETT", objNett.value
					if trim(objGross.value) <> "" then iTempGross = cdbl(iTempGross) + cdbl(objGross.value)
					if trim(objNett.value) <> "" then iTempNett = cdbl(iTempNett) + cdbl(objNett.value)
					HeaderNode.appendChild NewElem
				next

				if document.formname.chkAltUom.Checked then				
					HeaderNode.setAttribute "CHECK","Y"
					HeaderNode.setAttribute "ALTGROSS",document.formname.txtGross.value
					HeaderNode.setAttribute "ALTNETT",document.formname.txtNett.value
					HeaderNode.setAttribute "TOTGROSS",document.formname.txtGross.value
					HeaderNode.setAttribute "TOTNETT",document.formname.txtNett.value
				else
					HeaderNode.setAttribute "CHECK","N"
					HeaderNode.setAttribute "ALTGROSS",""
					HeaderNode.setAttribute "ALTNETT",""
					HeaderNode.setAttribute "TOTGROSS",iTempGross
					HeaderNode.setAttribute "TOTNETT",iTempNett
				end if

			end if
		next

		'alert Root.xml

		window.close()
	end if

End Function

Function CalculateQty()
	dim i,objDetails, iQty
	
	for i = 1 to RowNo
		set objDetails = eval("document.formname.txtDetails"&cstr(i))
		iQty = cdbl(iQty) + cdbl(objDetails.value)
	next
	
	idQtyEntered.innerText = iQty
	EnableDone()
	
End Function

Function EnableDone()
	if cdbl(idQty.innerText) = cdbl(idQtyEntered.innerText) then
		document.formname.BtnDone.disabled = False
	else
		document.formname.BtnDone.disabled = True
	end if
End Function

Function CheckEnable()
	if document.formname.chkAltUom.Checked then
		document.formname.txtGross.disabled = False
		document.formname.txtNett.disabled = False
	else
		document.formname.txtGross.disabled = True
		document.formname.txtNett.disabled = True
	end if
End Function

Function CheckGrossQty(obj)
	if trim(obj.value) = "" then exit function
	
	FindQty()
	
	if cdbl(iTotGross) > cdbl(iAltGross) then
		alert("Gross should be less than or equal to (" & (cdbl(iAltGross) - (cdbl(iTotGross) - cdbl(obj.value))) & ")")
		obj.select
		exit function
	end if
End Function

Function CheckNettQty(obj)
	if trim(obj.value) = "" then exit function
	
	FindQty()
	
	if cdbl(iTotNett) > cdbl(iAltNett) then
		alert("Nett should be less than or equal to (" & (cdbl(iAltNett) - (cdbl(iTotNett) - cdbl(obj.value))) & ")")
		obj.select
		exit function
	end if
End Function

Function FindQty()
	dim i,objGross,objNett
	
	iTotGross = 0
	iTotNett = 0
	
	if Root.hasChildNodes() then
		for Each HeaderNode in Root.childNodes
			if HeaderNode.attributes.getNamedItem("NO").value <> iNo then

				if trim(HeaderNode.attributes.getNamedItem("TOTGROSS").value) <> "" then
					iTotGross = cdbl(iTotGross) + cdbl(HeaderNode.attributes.getNamedItem("TOTGROSS").value)
				end if

				if trim(HeaderNode.attributes.getNamedItem("TOTNETT").value) <> "" then
					iTotNett = cdbl(iTotNett) + cdbl(HeaderNode.attributes.getNamedItem("TOTNETT").value)
				end if
			end if
		next
	end if
	
	for i = 1 to RowNo
		set objGross = eval("document.formname.txtAltGross"&cstr(i))
		if trim(objGross.value) <> "" then iTotGross = cdbl(iTotGross) + cdbl(objGross.value)
		
		set objNett = eval("document.formname.txtAltNett"&cstr(i))
		if trim(objNett.value) <> "" then iTotNett = cdbl(iTotNett) + cdbl(objNett.value)
	next
	
	if not document.formname.txtGross.disabled then
		if trim(document.formname.txtGross.value) <> "" then iTotGross = cdbl(iTotGross) + cdbl(document.formname.txtGross.value)
	end if
	if not document.formname.txtNett.disabled then
		if trim(document.formname.txtNett.value) <> "" then iTotNett = cdbl(iTotNett) + cdbl(document.formname.txtNett.value)
	end if
	
End Function

Function CheckGrossEntry()
	dim objGross
	
	for i = 1 to RowNo
		set objGross = eval("document.formname.txtAltGross"&cstr(i))
		if trim(objGross.value) <> "" then 
			CheckGrossEntry = False
			exit function
		end if
	next

	CheckGrossEntry = True
End Function

Function CheckNettEntry()
	for i = 1 to RowNo
		set objNett = eval("document.formname.txtAltNett"&cstr(i))

		if trim(objNett.value) <> "" then 
			CheckNettEntry = False
			exit function
		end if

	next

	CheckNettEntry = True
End Function