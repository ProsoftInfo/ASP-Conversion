dim objTemp,Root,newElem,SerialFlag
dim iClass,iItem,sOrgID,sType,j,iStore,iBin
dim iQtyTotGross,iQtyTotTare, iFixedNo
dim sLotNo, iCounter, iCount

Function LoadDetails(obj)
	dim sTempLotNo

	arrTemp = split(obj,":")
	sType = arrTemp(0)
	iItem = arrTemp(1)
	iClass = arrTemp(2)
	sOrgID = arrTemp(3)
	iStore = arrTemp(4)
	iBin = arrTemp(5)
	iTareValue = arrTemp(7)
	
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
'	alert Root.xml
	iQtyTotGross = 0
	iQtyTotTare = 0

	'Code to get the ItemName and Qty
	sExp ="//ITEM [ @ICODE = "&iItem&" and @CCODE = "&iClass&"]"
	Set ItemNode = Root.Selectnodes(sExp)

	For iCounter = 0 to ItemNode.Length - 1
		sItemName = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").Value
		idItemName.innerHTML = sItemName
		iQtyRec = ItemNode.Item(iCounter).Attributes.getNamedItem("QTY").Value
	next

	'Code to get the Entered Qty
	sExp ="//ITEM [ @ICODE = "&iItem&" and @CCODE = "&iClass&"]/STORAGE"
	Set StorageNode = Root.Selectnodes(sExp)
	For iCounter = 0 to StorageNode.Length - 1
		iTotStQty = cdbl(iTotStQty) + cdbl(StorageNode.Item(iCounter).Attributes.getNamedItem("QTY").Value)
	next
	idQty.innerHTML = cdbl(iQtyRec) - cdbl(iTotStQty)


	sExp ="//STOREDLOTDETAILS//SERIAL" '[ @ICODE = "&iItem&" and @CCODE = "&iClass&"]/STORAGE [ @STORE = "&iStore&" and @BIN = "&iBin&"]/LotSerial/LotSerialDetails"
	Set StorageNode = Root.Selectnodes(sExp)

	j=1
	For iCounter = 0 to StorageNode.Length - 1
		
		sLotNo=StorageNode.Item(iCounter).Attributes.getNamedItem("LOT").Value
		sLotserialNo=StorageNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value
		sQtyRec=StorageNode.Item(iCounter).Attributes.getNamedItem("QTYREC").Value
		sTar=StorageNode.Item(iCounter).Attributes.getNamedItem("TAREREC").Value

		sSell=StorageNode.Item(iCounter).Attributes.getNamedItem("SELLINGTYPE").Value
		iWeight=StorageNode.Item(iCounter).Attributes.getNamedItem("WEIGHTSTYPE").Value
		sPack=StorageNode.Item(iCounter).Attributes.getNamedItem("PACKINGTYPE").Value
		sSellingForm=StorageNode.Item(iCounter).Attributes.getNamedItem("SELLINGFORM").Value
		iPackNo=StorageNode.Item(iCounter).Attributes.getNamedItem("PACKNUMBER").Value
		sTempGross=StorageNode.Item(iCounter).Attributes.getNamedItem("QTYGRO").Value
		Flag=StorageNode.Item(iCounter).Attributes.getNamedItem("FLAG").Value

		'Checking whether the serial is already selected		
		sExp ="//ITEM [ @ICODE = "&iItem&" and @CCODE = "&iClass&"]/STORAGE [ @STORE = "&iStore&" and @BIN = "&iBin&"]/LotSerial/LotSerialDetails"
		set Node = Root.SelectNodes(sExp)
		for i = 0 to Node.length - 1
			if Node.item(i).attributes.getNamedItem("LOTSERIAL").value = sLotserialNo and Node.Item(i).Attributes.getNamedItem("QTYREC").Value = sQtyRec and Node.Item(i).Attributes.getNamedItem("PACKNUMBER").Value = iPackNo then
				Flag="E"
				exit for
			end if
		next
		
		if sTempLotNo <> sLotNo then
			sTempLotNo = sLotNo 
			iFixedNo = iFixedNo + 1
			set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
			set headerCell=oRow.insertCell()									
			headerCell.innerHTML=sLotNo
			headerCell.className="ExcelDisplayCell"
			headerCell.align="left"
			headerCell.colspan="10"
		end if 

		set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)

		set headerCell=oRow.insertCell()									
		headerCell.innerHTML=j
		headerCell.className="ExcelDisplayCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()		
		if Flag = "Y" then 'Not Selected
			set oText = document.createElement("<input type=""Checkbox"" name=""ChkItem"&CStr(j)&""" class=""Formelem"" >" )
		elseif Flag = "N" then 'Selected
			set oText = document.createElement("<input type=""Checkbox"" name=""ChkItem"&CStr(j)&""" class=""Formelem"" DISABLED>" )
		elseif Flag = "E" then 'Already Selected
			set oText = document.createElement("<input type=""Checkbox"" name=""ChkItem"&CStr(j)&""" class=""Formelem"" CHECKED>" )
		end if
		headerCell.className="ExcelFieldCell"
		headerCell.appendChild(oText)
		headerCell.width = "8"
		headerCell.align="center"


		'set headerCell=oRow.insertCell()		
		'set oText = document.createElement("<input type=""text"" name=""txtSerial"&CStr(j)&""" size=""8"" value="""&sLotserialNo&""" maxlength=10 READONLY class=""FormelemRead"" style=""text-align=center"">" )
		'headerCell.className="ExcelDisplayCell"
		'headerCell.appendChild(oText)
		'headerCell.width = "8"
		'headerCell.align="center"
		
		set oText = document.createElement("<input type=""hidden"" name=""txtSerial"&CStr(j)&""" size=""8"" value="""&sLotserialNo&""" >" )
		document.formname.appendchild(oText)
		
		set oText = document.createElement("<input type=""hidden"" name=""txtLot"&CStr(j)&""" size=""8"" value="""&sLotNo&""" >" )
		document.formname.appendchild(oText)
		
		set headerCell=oRow.insertCell()									
		set oText = document.createElement("<input type=""text"" name=""txtGross"&CStr(j)&""" size=""12"" value="""&sQtyRec&""" maxlength=10 READONLY class=""FormelemRead"" style=""text-align=right"">" )
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelDisplayCell"

		set headerCell=oRow.insertCell()									
		set oText = document.createElement("<input type=""text"" name=""txtTare"&CStr(j)&""" size=""12"" value="""&sTar&""" maxlength=10 READONLY class=""FormelemRead"" style=""text-align=right"">")
		headerCell.className="ExcelDisplayCell"
		headerCell.width = "10"
		headerCell.appendChild(oText)

		set headerCell=oRow.insertCell()
		CheckFlag = False

		Set RootO = SellingData.documentElement
		if RootO.hasChildNodes() then
			For Each HeaderONode In RootO.childNodes
				if trim(HeaderONode.Attributes.Item(0).nodeValue) = sSell then
					set oText = document.createElement("<input type=""text"" name=""selSell"&CStr(j)&""" size=""12"" maxlength=10 value="""&trim(HeaderONode.Attributes.Item(1).nodeValue)&""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
					CheckFlag = True
					exit for
				end if
			next
			if not CheckFlag then set oText = document.createElement("<input type=""text"" name=""selSell"&CStr(j)&""" size=""12"" maxlength=10 value="""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
		else
			set oText = document.createElement("<input type=""text"" name=""selSell"&CStr(j)&""" size=""12"" maxlength=10 value="""" READONLY class=""FormelemRead"" style=""text-align=left"">" )

		end if
		
		headerCell.appendChild(oText)
		headerCell.className="ExcelDisplayCell"
		headerCell.align="left"

		set headerCell=oRow.insertCell()									
		set oText = document.createElement("<input type=""text"" name=""txtSellWeight"&CStr(j)&""" size=""12"" maxlength=10 value="""&iWeight&""" READONLY class=""FormelemRead"" style=""text-align=right"">" )
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelDisplayCell"


		set oText = document.createElement("<input type=""hidden"" name=""txtTempGross"&CStr(j)&""" size=""8"" value="""&sTempGross&""" >" )
		document.formname.appendchild(oText)

		set headerCell=oRow.insertCell()
		CheckFlag = False

		Set RootO = PackingData.documentElement
		if RootO.hasChildNodes() then
			For Each HeaderONode In RootO.childNodes
				if trim(HeaderONode.Attributes.Item(0).nodeValue) = sPack then
					set oText = document.createElement("<input type=""text"" name=""selPack"&CStr(j)&""" size=""12"" maxlength=10 value="""&trim(HeaderONode.Attributes.Item(1).nodeValue)&""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
					CheckFlag = True
					exit for
				end if
			next
			if not CheckFlag then set oText = document.createElement("<input type=""text"" name=""selPack"&CStr(j)&""" size=""12"" maxlength=10 value="""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
		else
			set oText = document.createElement("<input type=""text"" name=""selPack"&CStr(j)&""" size=""12"" maxlength=10 value="""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
		end if

		headerCell.appendChild(oText)
		headerCell.className="ExcelDisplayCell"
		headerCell.align="left"

		set headerCell=oRow.insertCell()
		CheckFlag = False

		Set RootO = SellingFormData.documentElement
		if RootO.hasChildNodes() then
			For Each HeaderONode In RootO.childNodes
				if trim(HeaderONode.Attributes.Item(0).nodeValue) = sSellingForm then
					set oText = document.createElement("<input type=""text"" name=""selForm"&CStr(j)&""" size=""12"" maxlength=10 value="""&trim(HeaderONode.Attributes.Item(1).nodeValue)&""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
					CheckFlag = True
					exit for
				end if
			next
			if not CheckFlag then set oText = document.createElement("<input type=""text"" name=""selForm"&CStr(j)&""" size=""12"" maxlength=10 value="""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
		else
			set oText = document.createElement("<input type=""text"" name=""selForm"&CStr(j)&""" size=""12"" maxlength=10 value="""" READONLY class=""FormelemRead"" style=""text-align=left"">" )
		end if

		headerCell.appendChild(oText)
		headerCell.className="ExcelDisplayCell"
		headerCell.align="left"


		set headerCell=oRow.insertCell()									
		set oText = document.createElement("<input type=""text"" name=""txtPackNo"&CStr(j)&""" size=""17"" value="""&iPackNo&""" maxlength=30 READONLY class=""FormelemRead"" style=""text-align=left"">" )
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelDisplayCell"
		
		k = k+1
		j = j+1
	next
	
end Function

Function AddChecked()
	dim iSellingForm, iPackCode, iPackingForm, sTempLot

	'Adding Lot Details
	Set Root = objTemp.documentElement
	'alert Root.xml
	For Each HeaderNode In Root.childNodes
		if HeaderNode.nodeName = "ITEM" then
			if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass and HeaderNode.Attributes.Item(2).nodeValue = sOrgID then
				For Each childNod In HeaderNode.childNodes
					if childNod.Attributes.Item(0).nodeValue = iStore and childNod.Attributes.Item(1).nodeValue = iBin then
						for each Node in ChildNod.ChildNodes
							ChildNod.removeChild Node
						next

						childNod.setAttribute "QTY", CalculateQty

						for i = 1 to j - 1
							set objCheck = eval("document.formname.chkItem"&CStr(i))
							set objLotNo = eval("document.formname.txtLot"&CStr(i))

							if objCheck.Checked and (sTempLot = "" or sTempLot <> objLotNo.value) then

								Set newElem = objTemp.createElement("LotSerial")

								newElem.setAttribute "QTYIN", ""
								newElem.setAttribute "TARE", ""
								newElem.setAttribute "LOT", objLotNo.value
								newElem.setAttribute "SERIALFROM", ""
								newElem.setAttribute "SERIALTO", ""
								newElem.setAttribute "TAREWEIGHT", ""
								newElem.setAttribute "QTY", CalculateLotQty(objLotNo.value)
								iCount = cdbl(iCount) + 1
								newElem.setAttribute "COUNTER", cdbl(iCount)
								childNod.appendChild newElem							

								sTempLot = objLotNo.value
								
								if sType="S" then exit for

							end if
						next
					end if
				next
			end if
		end if
	next

	'Adding Serial Details	
	for i = 1 to j - 1
		set objCheck = eval("document.formname.chkItem"&CStr(i))

		set objLotNo = eval("document.formname.txtLot"&CStr(i))
		set objSerial = eval("document.formname.txtSerial"&CStr(i))
		set objGross = eval("document.formname.txtGross"&CStr(i))
		set objTare = eval("document.formname.txtTare"&CStr(i))
		set objSel = eval("document.formname.selSell"&CStr(i))
		set objWei = eval("document.formname.txtSellWeight"&CStr(i))
		set objPac = eval("document.formname.selPack"&CStr(i))
		set objFor = eval("document.formname.selForm"&CStr(i))	
		set objPackNo = eval("document.formname.txtPackNo"&CStr(i))
		set objTempGross = eval("document.formname.txtTempGross"&CStr(i))
		
		iSellingForm = GetSellingForm(objSel.value)
		iPackCode = GetPackingCode(objPac.value)
		iPackingForm = GetPackingForm(objFor.value) 
		
		if objCheck.checked then
			
			Set newElem1 = objTemp.createElement("LotSerialDetails")
			newElem1.setAttribute "LOTSERIAL", trim(objSerial.value)
			newElem1.setAttribute "QTYREC", trim(objGross.value)
			newElem1.setAttribute "TAREREC", trim(objTare.value)
			newElem1.setAttribute "SELLINGTYPE", iSellingForm 
			newElem1.setAttribute "WEIGHTSTYPE", trim(objWei.value)
			newElem1.setAttribute "PACKINGTYPE", iPackCode 
			newElem1.setAttribute "LOT", trim(objLotNo.value)
			newElem1.setAttribute "SELLINGFORM", iPackingForm 
			newElem1.setAttribute "PACKNUMBER", trim(objPackNo.value)
			newElem1.setAttribute "QTYGRO", trim(objTempGross.value)
			
			For Each HeaderNode In Root.childNodes
				if HeaderNode.NodeName = "ITEM" then
					if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass and HeaderNode.Attributes.Item(2).nodeValue = sOrgID then
						For Each childNod In HeaderNode.childNodes
							if childNod.Attributes.Item(0).nodeValue = iStore and childNod.Attributes.Item(1).nodeValue = iBin then
								for each Node in ChildNod.ChildNodes
									if Node.attributes.getNamedItem("LOT").value = objLotNo.value then
										Node.appendChild newElem1
										exit for
									end if
								next
							end if
						next
					end if
				end if
			next
			
			
			sExp ="//STOREDLOTDETAILS//SERIAL [@LOTSERIAL = '" & trim(objSerial.value) & "' and @QTYREC = '" & trim(objGross.value) & "' and @TAREREC = '" & trim(objTare.value) & "' and @SELLINGTYPE = '" & trim(iSellingForm) & "' and @WEIGHTSTYPE = '" & trim(objWei.value) & "' and @PACKINGTYPE = '" & trim(iPackCode) & "' and @SELLINGFORM = '" & trim(iPackingForm) & "' and @PACKNUMBER = '" & trim(objPackNo.value) & "']"
			'[ @ICODE = "&iItem&" and @CCODE = "&iClass&"]/STORAGE [ @STORE = "&iStore&" and @BIN = "&iBin&"]/LotSerial/LotSerialDetails"
			set Node = root.selectNodes(sExp)
			'alert sExp & ">>" & Node.length
			if Node.length > 0 then Node.item(0).attributes.getNamedItem("FLAG").value = "N"

		elseif not objCheck.Disabled then
			sExp ="//STOREDLOTDETAILS//SERIAL [@LOTSERIAL = '" & trim(objSerial.value) & "' and @QTYREC = '" & trim(objGross.value) & "' and @TAREREC = '" & trim(objTare.value) & "' and @SELLINGTYPE = '" & trim(iSellingForm) & "' and @WEIGHTSTYPE = '" & trim(objWei.value) & "' and @PACKINGTYPE = '" & trim(iPackCode) & "' and @SELLINGFORM = '" & trim(iPackingForm) & "' and @PACKNUMBER = '" & trim(objPackNo.value) & "']"
			'[ @ICODE = "&iItem&" and @CCODE = "&iClass&"]/STORAGE [ @STORE = "&iStore&" and @BIN = "&iBin&"]/LotSerial/LotSerialDetails"
			set Node = root.selectNodes(sExp)
			if Node.length > 0 then Node.item(0).attributes.getNamedItem("FLAG").value = "Y"
		end if
	next

'	alert Root.xml
	window.close

End Function

Function CalculateQty()
	dim i, objCheck, objQty, objTare, iQty, iTare
	
	for i = 1 to j-1
		set objCheck = eval("document.formname.ChkItem"&CStr(i))
		if objCheck.checked then
			set objQty = eval("document.formname.txtGross"&CStr(i))
			set objTare = eval("document.formname.txtTare"&CStr(i))
			iQty = cdbl(iQty) + cdbl(objQty.value)
			'iTare = cdbl(iTare) + cdbl(objTare.value)
			iTare = 0
		end if
	next
	
	CalculateQty = cdbl(iQty) - cdbl(iTare) 
End Function

Function CalculateLotQty(sLotNo)
	dim i, objCheck, objQty, objTare, iQty, iTare, objLotNo
	
	for i = 1 to j-1
		set objCheck = eval("document.formname.ChkItem"&CStr(i))
		set objLotNo = eval("document.formname.txtLot"&CStr(i))
		
		if objCheck.checked and objLotNo.value=sLotNo then
			set objQty = eval("document.formname.txtGross"&CStr(i))
			set objTare = eval("document.formname.txtTare"&CStr(i))
			iQty = cdbl(iQty) + cdbl(objQty.value)
			iTare = cdbl(iTare) + cdbl(objTare.value)
		end if
	next
	
	CalculateLotQty = cdbl(iQty) - cdbl(iTare) 
End Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

Function SelectAll()
	dim i, objCheck, Flag
	
	if document.formname.ChkAll.checked then
		Flag = True
	else
		Flag = False
	end if
	
	for i = 1 to j-1
		set objCheck = eval("document.formname.ChkItem"&CStr(i))
		if not objCheck.Disabled then
			objCheck.Checked = Flag
		end if
	next
End Function

Function GetSellingForm(iSellingForm)
	
	Set RootO = SellingData.documentElement
	if RootO.hasChildNodes() then
		For Each HeaderONode In RootO.childNodes
			if trim(HeaderONode.Attributes.Item(1).nodeValue) = iSellingForm then
				GetSellingForm = trim(HeaderONode.Attributes.Item(0).nodeValue)
				exit function
			end if
		next
	end if
	GetSellingForm = 0
	
End Function

Function GetPackingCode(iPackCode)

	Set RootO = PackingData.documentElement
	if RootO.hasChildNodes() then
		For Each HeaderONode In RootO.childNodes
			if trim(HeaderONode.Attributes.Item(1).nodeValue) = iPackCode then
				GetPackingCode = trim(HeaderONode.Attributes.Item(0).nodeValue)
				exit function
			end if
		next
	end if
	GetPackingCode = 0
	
End Function

Function GetPackingForm(iPackingForm)

	Set RootO = SellingFormData.documentElement
	if RootO.hasChildNodes() then
		For Each HeaderONode In RootO.childNodes
			if trim(HeaderONode.Attributes.Item(1).nodeValue) = iPackingForm then
				GetPackingForm = trim(HeaderONode.Attributes.Item(0).nodeValue)
				exit function
			end if
		next
	end if
	GetPackingForm = 0
	
End Function