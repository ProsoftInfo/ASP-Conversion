dim objTemp,Root,newElem
dim iClass,iItem,sOrgID,sType,j,iRecNo,iMRSNo,iIssNo,iLot
dim bFlag
bFlag = true

Function checkNumbers(val)
	dim valid,temp,i
	valid = "0123456789."
	for i=1 to len(val)
		temp = mid(val,i,1)
		if Instr(1,valid,temp) > 0 then
			checkNumbers = true
		else
			checkNumbers = false
			exit for
		end if
	next
end Function

Function fnInit(obj)
	arrTemp = split(obj,":")
	iRecNo	= arrTemp(1)
	iClass	= arrTemp(2)
	iItem	= arrTemp(3)
	sOrgID	= arrTemp(4)
	iMRSNo  = arrTemp(5)
	iIssNo  = arrTemp(6)
	iLot    = arrTemp(7)

	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement

	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass and HeaderNode.Attributes.Item(2).nodeValue = sOrgID and HeaderNode.Attributes.Item(4).nodeValue = iMRSNo and HeaderNode.Attributes.Item(5).nodeValue = iLot and HeaderNode.Attributes.Item(6).nodeValue = iIssNo then
				idQty.innerHTML = HeaderNode.Attributes.Item(7).nodeValue
			end if
		next
	end if

end Function

Function DisableTxt(obj)
	document.formname.txtTare.value = ""
	if obj.value = "I" then
		document.formname.txtTare.disabled = true
	else
		document.formname.txtTare.disabled = false
	end if
end Function

Function CheckAdd()
	dim i,k
	
	if not document.formname.chkNew.checked then exit function

	if not (document.formname.radQtyIn(0).checked or document.formname.radQtyIn(1).checked) then
		alert("Select Quantity in")
		document.formname.radQtyIn(0).focus
		exit function
	elseif not (document.formname.radTare(0).checked or document.formname.radTare(1).checked) then
		alert("Select Tare Weight")
		document.formname.radTare(0).focus
		exit function
	elseif (document.formname.radTare(0).checked and trim(document.formname.txtTare.value) = "") then
		alert("Enter Tare Weight")
		document.formname.txtTare.select
		exit function
	elseif (document.formname.radTare(0).checked and trim(document.formname.txtTare.value) <> "" and not checkNumbers(trim(document.formname.txtTare.value))) then
		alert("Enter only Numerals")
		document.formname.txtTare.select
		exit function
	elseif (trim(document.formname.txtLotNumber.value) = "") then
		alert("Enter Lot Number")
		document.formname.txtLotNumber.select
		exit function
	elseif (not checkNumbers(trim(document.formname.txtLotNumber.value))) then
		alert("Enter only Numerals")
		document.formname.txtLotNumber.select
		exit function
	elseif (trim(document.formname.txtSerialFrom.value) = "") then
		alert("Enter Lot Serial Number From")
		document.formname.txtSerialFrom.select
		exit function
	elseif (not checkNumbers(trim(document.formname.txtSerialFrom.value))) then
		alert("Enter only Numerals")
		document.formname.txtSerialFrom.select
		exit function
	elseif (trim(document.formname.txtSerialTo.value) = "") then
		alert("Enter Lot Serial Number To")
		document.formname.txtSerialTo.select
		exit function
	elseif (not checkNumbers(trim(document.formname.txtSerialTo.value))) then
		alert("Enter only Numerals")
		document.formname.txtSerialTo.select
		exit function
	elseif (trim(cdbl(document.formname.txtSerialFrom.value)) > trim(cdbl(document.formname.txtSerialTo.value))) then
		alert("Lot Serial From Number should be less than Lot Serial To Number")
		document.formname.txtSerialTo.select
		exit function
	else
		bFlag = false
		iLotNo = cdbl(document.formname.txtSerialTo.value) - cdbl(document.formname.txtSerialFrom.value)

		k = 0
		j = 0

		if iLotNo = 0 then
			j = 2
			for i=2 to document.all.tblLot.rows.length - 1
				document.all.tblLot.deleteRow(2) 
			next

			set oRow = document.all.tblLot.insertRow(2)

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML=1
			headerCell.className="ExcelSerial"
			headerCell.align="center"
														
			set headerCell=oRow.insertCell()									
			headerCell.innerHTML=trim(document.formname.txtLotNumber.value) & " - " & cdbl(document.formname.txtSerialFrom.value)
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()									
			set oText = document.createElement("<input type=""text"" name=""txtGross1"" size=""12"" maxlength=10 class=""Formelem"">" )
			headerCell.appendChild(oText)
			headerCell.className="ExcelInputCell"
		
			set headerCell=oRow.insertCell()									
			if document.formname.radTare(1).checked then
				set oText = document.createElement("<input type=""text"" name=""txtTare1"" size=""12"" maxlength=10 class=""Formelem"">")
			headerCell.className="ExcelInputCell"
			else
				set oText = document.createElement("<input type=""text"" name=""txtTare1"" size=""12"" maxlength=10 value="""&trim(document.formname.txtTare.value)&""" READONLY class=""FormelemRead"">")
			headerCell.className="ExcelDisplayCell"
			end if
			headerCell.appendChild(oText)

			exit function
		end if
		
		ClearTable

		for j = 1 to iLotNo + 1
			set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML=j
			headerCell.className="ExcelSerial"
			headerCell.align="center"
														
			set headerCell=oRow.insertCell()									
			headerCell.innerHTML=trim(document.formname.txtLotNumber.value) & " - " & cdbl(document.formname.txtSerialFrom.value) + k
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()									
			set oText = document.createElement("<input type=""text"" name=""txtGross"&CStr(j)&""" size=""12"" maxlength=10 class=""Formelem"">" )
			headerCell.appendChild(oText)
			headerCell.className="ExcelInputCell"
		
			set headerCell=oRow.insertCell()									
			if document.formname.radTare(1).checked then
				set oText = document.createElement("<input type=""text"" name=""txtTare"&CStr(j)&""" size=""12"" maxlength=10 class=""Formelem"">")
				headerCell.className="ExcelInputCell"
			else
				set oText = document.createElement("<input type=""text"" name=""txtTare"&CStr(j)&""" size=""12"" maxlength=10 value="""&trim(document.formname.txtTare.value)&""" READONLY class=""FormelemRead"">")
				headerCell.className="ExcelDisplayCell"
			end if
			headerCell.appendChild(oText)
		
			k = k+1
		next
		
	end if
end Function

Function ClearTable()
	dim i
	for i=2 to document.all.tblLot.rows.length - 1
		document.all.tblLot.deleteRow(2) 
	next
end Function

Function clearAll()
	ClearTable
	document.formname.reset()
	j = 0
	i = 0
	k = 0
	bFlag = true
end Function

Function CheckSubmit()
	
	if document.formname.chkNew.checked then 
		if bFlag then 
			CheckAdd()
			exit function
		end if
	else
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass and HeaderNode.Attributes.Item(2).nodeValue = sOrgID and HeaderNode.Attributes.Item(4).nodeValue = iMRSNo and HeaderNode.Attributes.Item(5).nodeValue = iLot and HeaderNode.Attributes.Item(6).nodeValue = iIssNo then
					For Each childNod In HeaderNode.childNodes
						set oRem = HeaderNode.removeChild(childNod)
					next
				end if
			next
		end if
		'alert RootO.XML
		'Exit Function
		Set RootO = OutData.documentElement
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass and HeaderNode.Attributes.Item(2).nodeValue = sOrgID and HeaderNode.Attributes.Item(4).nodeValue = iMRSNo and HeaderNode.Attributes.Item(5).nodeValue = iLot and HeaderNode.Attributes.Item(6).nodeValue = iIssNo then
				For Each HeaderONode In RootO.childNodes
					HeaderNode.AppendChild HeaderONode
				next
				exit for
			end if		
		next
		
		window.close
		exit function
	end if

	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass and HeaderNode.Attributes.Item(2).nodeValue = sOrgID and HeaderNode.Attributes.Item(4).nodeValue = iMRSNo and HeaderNode.Attributes.Item(5).nodeValue = iLot and HeaderNode.Attributes.Item(6).nodeValue = iIssNo then
				For Each childNod In HeaderNode.childNodes
					set oRem = HeaderNode.removeChild(childNod)
				next
			end if
		next
	end if

	dim i,iQtyTotGross,iQtyTotTare
	dim objQ,objD,z
	
	'if j = 0 then
	'	alert("No Lot Details Enterted")
	'	exit function
	'end if
	
	for i = 1 to j - 1
		set objD = eval("document.formname.txtGross"&CStr(i))
		set objQ = eval("document.formname.txtTare"&CStr(i))

		if trim(objD.value) = "" then
			alert("Enter Gross / Nett")
			objD.select()
			exit function
		elseif not checkNumbers(objD.value) then
			alert("Enter Numerals Only")
			objD.select()
			exit function
		elseif trim(objQ.value) = "" then
			alert("Enter Tare")
			objQ.select()
			exit function
		elseif not checkNumbers(objQ.value) then
			alert("Enter Numerals Only")
			objQ.select()
			exit function
		else
			iQtyTotGross = cdbl(iQtyTotGross) + cdbl(objD.value)
			iQtyTotTare = cdbl(iQtyTotTare) + cdbl(objQ.value)
		end if
	next

	if (cdbl(iQtyTotGross+iQtyTotTare) < cdbl(idQty.innerHTML)) or (cdbl(iQtyTotGross+iQtyTotTare) > cdbl(idQty.innerHTML)) then
		alert("Total Received Quantity should be equal to Quantity (" &idQty.innerHTML& ")")
		exit function
	end if
	
	i = 1
	dim a

	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass and HeaderNode.Attributes.Item(2).nodeValue = sOrgID and HeaderNode.Attributes.Item(4).nodeValue = iMRSNo and HeaderNode.Attributes.Item(5).nodeValue = iLot and HeaderNode.Attributes.Item(6).nodeValue = iIssNo then
				Set newElem = objTemp.createElement("LotSerial")

				if document.formname.radQtyIn(0).checked then
					newElem.setAttribute "QTYIN", "G"
				else
					newElem.setAttribute "QTYIN", "N"
				end if

				newElem.setAttribute "TARE", trim(document.formname.txtTare.value)
				newElem.setAttribute "LOT", trim(document.formname.txtLotNumber.value)
				newElem.setAttribute "SERIALFROM", trim(document.formname.txtSerialFrom.value)
				newElem.setAttribute "SERIALTO", trim(document.formname.txtSerialTo.value)

				if document.formname.radTare(0).checked then
					newElem.setAttribute "TAREWEIGHT", "U"
				else
					newElem.setAttribute "TAREWEIGHT", "I"
				end if
				
				HeaderNode.appendChild newElem

				for i = 1 to j - 1
					set objD = eval("document.formname.txtGross"&CStr(i))
					set objQ = eval("document.formname.txtTare"&CStr(i))

					Set newElem1 = objTemp.createElement("LotSerialDetails")
					newElem1.setAttribute "LOTSERIAL", trim(document.formname.txtLotNumber.value) & " - " & cdbl(document.formname.txtSerialFrom.value) + z
					newElem1.setAttribute "QTYREC", trim(objD.value)
					newElem1.setAttribute "TAREREC", trim(objQ.value)
					newElem.appendChild newElem1
					z = z+1
				next
			window.close
			exit function
			end if
		next
	end if
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

