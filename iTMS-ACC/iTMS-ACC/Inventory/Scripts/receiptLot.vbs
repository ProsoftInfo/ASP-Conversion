dim objTemp,Root,newElem,SerialFlag
dim iClass,iItem,sOrgID,sType,j,iStore,iBin
dim iQtyTotGross,iQtyTotTare, iFixedNo
dim sLotNo, iCounter, iCount

Function LoadDetails(obj)
	dim sTempLotNo
	'alert obj
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
	'alert Root.XML
	iQtyTotGross = 0
	iQtyTotTare = 0

	'Code to get the ItemName and Qty
	sExp ="//ITEM [ @ITEM = "&iItem&" and @CLASS = "&iClass&"]"
	Set ItemNode = Root.Selectnodes(sExp)

	For iCounter = 0 to ItemNode.Length - 1
		sItemName = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").Value
		idItemName.innerHTML = sItemName
		iQtyRec = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEMQTY").Value
	next

	'Code to get the Entered Qty
	sExp ="//ITEM [ @ITEM = "&iItem&" and @CLASS = "&iClass&"]/STORAGE"
	Set StorageNode = Root.Selectnodes(sExp)
	For iCounter = 0 to StorageNode.Length - 1
		iTotStQty = cdbl(iTotStQty) + cdbl(StorageNode.Item(iCounter).Attributes.getNamedItem("STOREQTY").Value)
	next
	idQty.innerHTML = cdbl(iQtyRec) - cdbl(iTotStQty)

	'alert iTotStQty
	sExp ="//ITEM[@ITEM = "&iItem&" and @CLASS = "&iClass&"]/STORAGE [ @STORE = "&iStore&" and @BIN = "&iBin&"]/LOT"
	Set StorageNode = Root.Selectnodes(sExp)

	j=1
	For iCounter = 0 to StorageNode.Length - 1
		sLotNo=StorageNode.Item(iCounter).Attributes.getNamedItem("LOT").Value
		sLotserialNo=StorageNode.Item(iCounter).Attributes.getNamedItem("SERIALNO").Value
		sTempGross=StorageNode.Item(iCounter).Attributes.getNamedItem("GROSSQTY").Value
		sQtyRec=StorageNode.Item(iCounter).Attributes.getNamedItem("QTY").Value
		sTar=cdbl(sTempGross)-cdbl(sQtyRec)
		sSell=StorageNode.Item(iCounter).Attributes.getNamedItem("SELLINGNUMBER").Value
		iWeight=StorageNode.Item(iCounter).Attributes.getNamedItem("WEIGHTPERSELLINGFORM").Value
		sPack=StorageNode.Item(iCounter).Attributes.getNamedItem("PACKINGCODE").Value
		sSellingForm=StorageNode.Item(iCounter).Attributes.getNamedItem("SELLINGFORM").Value
		iPackNo=StorageNode.Item(iCounter).Attributes.getNamedItem("PACKINGNUMBER").Value
		sAttList=StorageNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTE").Value
		iStockQuality = StorageNode.Item(iCounter).Attributes.getNamedItem("SQ").Value
		iRate = StorageNode.Item(iCounter).Attributes.getNamedItem("RATE").Value
		
		Flag=StorageNode.Item(iCounter).Attributes.getNamedItem("FLAG").Value

		'Checking whether the serial is already selected		
		sExp ="//ITEM [ @ITEM = "&iItem&" and @CLASS = "&iClass&"]/STORAGE [ @STORE = "&iStore&" and @BIN = "&iBin&"]/LOT"
		set Node = Root.SelectNodes(sExp)
		for i = 0 to Node.length - 1
			if Node.item(i).attributes.getNamedItem("SERIALNO").value = sLotserialNo and Node.Item(i).Attributes.getNamedItem("QTY").Value = sQtyRec and Node.Item(i).Attributes.getNamedItem("PACKINGNUMBER").Value = iPackNo then
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
		set oText = document.createElement("<input type=""text"" name=""txtG"&CStr(j)&""" size=""12"" value="""&sTempGross&""" maxlength=10 READONLY class=""FormelemRead"" style=""text-align=right"">" )
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelDisplayCell"

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
		set oText = document.createElement("<input type=""hidden"" name=""hAttList"&Cstr(j)&""" value="""&sAttList&""" >")
		headerCell.appendChild(oText)
		set oText = document.createElement("<input type=""hidden"" name=""hSQ"&Cstr(j)&""" value="""&iStockQuality&""" >")
		headerCell.appendChild(oText)
		set oText = document.createElement("<input type=""hidden"" name=""hRate"&Cstr(j)&""" value="""&iRate&""">")
		headerCell.appendChild(oText)
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
	'alert "Adding Lot Details>>>"
	'Adding Lot Details
	Set Root = objTemp.documentElement
	For Each HeaderNode In Root.childNodes
		if HeaderNode.nodeName = "ITEM" then
			if HeaderNode.getAttribute("ITEM") = iItem and HeaderNode.getAttribute("CLASS") = iClass then
				For Each childNod In HeaderNode.childNodes
					if childNod.getAttribute("ITEM")=iItem and childNod.getAttribute("CLASS")=iClass then
						iStoreEntryNo =childNod.getAttribute("STOENTRYNO")
						for each Node in ChildNod.ChildNodes
							ChildNod.removeChild Node
						next
						childNod.setAttribute "STOREQTY","0"
					end if
				next
				exit for
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
		set objAttList = eval("document.formname.hAttList"&Cstr(i))
		set objSQ = eval("document.formname.hSQ"&cstr(i))
		set objRate = eval("document.formname.hRate"&cstr(i))
		
		iSellingForm = GetSellingForm(objSel.value)
		iPackCode = GetPackingCode(objPac.value)
		iPackingForm = GetPackingForm(objFor.value) 
		
		if objCheck.checked then
	
			Set newElem = objTemp.createElement("LOT")	
				newElem.setAttribute "LOTENTRYNO",iStoreEntryNo
				newElem.setAttribute "ITEM",iItem
				newElem.setAttribute "CLASS",iClass
				newElem.setAttribute "STORE",iStore
				newElem.setAttribute "BIN",iBin
				newElem.setAttribute "LOT",trim(objLotNo.value)
				newElem.setAttribute "QTY",trim(objGross.value)
				newElem.setAttribute "RATE",objRate.value
				newElem.setAttribute "GROSSQTY",trim(objTempGross.value)
				newElem.setAttribute "PACKINGNUMBER",trim(objPackNo.value)
				newElem.setAttribute "PACKINGCODE",iPackCode 
				newElem.setAttribute "SELLINGNUMBER",iSellingForm
				newElem.setAttribute "WEIGHTPERSELLINGFORM",trim(objWei.value)
				newElem.setAttribute "SELLINGFORM",iPackingForm
				newElem.setAttribute "STAGE","0"
				newElem.setAttribute "ATTRIBUTE",objAttList.value
				newElem.setAttribute "SERIALNO",trim(objSerial.value)
				newElem.setAttribute "SQ",objSQ.value
				newElem.setAttribute "FLAG","Y"
		
			
			For Each HeaderNode In Root.childNodes
				if HeaderNode.NodeName = "ITEM" then
					if HeaderNode.getAttribute("ITEM") = iItem and HeaderNode.getAttribute("CLASS") = iClass then
						For Each childNod In HeaderNode.childNodes
							if childNod.getAttribute("ITEM")=iItem and childNod.getAttribute("CLASS")=iClass and childNod.getAttribute("STORE") = iStore and childNod.getAttribute("BIN") = iBin then
								childNod.appendChild newElem
								childNod.setAttribute "STOREQTY", CalculateQty
								exit for
							end if
						next
					end if
				end if
			next
			
			
			sExp ="//ITEM//STORE//LOT[@SERIALNO = '" & trim(objSerial.value) & "' and @QTY = '" & trim(objGross.value) & "' and @SELLINGNUMBER = '" & trim(iSellingForm) & "' and @WEIGHTPERSELLINGFORM = '" & trim(objWei.value) & "' and @PACKINGCODE = '" & trim(iPackCode) & "' and @SELLINGFORM = '" & trim(iPackingForm) & "' and @PACKINGNUMBER = '" & trim(objPackNo.value) & "' and @ITEM="& iItem &" and @CLASS="& iClass &" and @STORE="& iStore &" and @BIN="& iBin &"]"
			set Node = root.selectNodes(sExp)
			if Node.length > 0 then Node.item(0).attributes.getNamedItem("FLAG").value = "N"

		elseif not objCheck.Disabled then
			sExp ="//ITEM//STORE//LOT[@SERIALNO = '" & trim(objSerial.value) & "' and @QTY = '" & trim(objGross.value) & "' and @SELLINGNUMBER = '" & trim(iSellingForm) & "' and @WEIGHTPERSELLINGFORM = '" & trim(objWei.value) & "' and @PACKINGCODE = '" & trim(iPackCode) & "' and @SELLINGFORM = '" & trim(iPackingForm) & "' and @PACKINGNUMBER = '" & trim(objPackNo.value) & "' and @ITEM="& iItem &" and @CLASS="& iClass &" and @STORE="& iStore &" and @BIN="& iBin &"]"
			set Node = root.selectNodes(sExp)
			if Node.length > 0 then Node.item(0).attributes.getNamedItem("FLAG").value = "Y"
		end if
	next
	
	For Each HeaderNode In Root.childNodes
		if HeaderNode.nodeName = "ITEM" then
			if HeaderNode.getAttribute("ITEM") = iItem and HeaderNode.getAttribute("CLASS") = iClass then
				For Each childNod In HeaderNode.childNodes
					if childNod.getAttribute("ITEM")=iItem and childNod.getAttribute("CLASS")=iClass then
						iStoQty = childNod.getAttribute("STOREQTY")
						if cdbl(iStoQty)<=0 then
							HeaderNode.removeChild(childNod)
						end if
					end if
				next
				exit for
			end if
		end if
	next
	
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