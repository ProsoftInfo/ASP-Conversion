dim objTemp,Root,newElem
dim iLineNo,iPickNo,iQty,ii,sLot,sBin,sStore,sOrgID,iItem,iClass

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
'--------------------------------------------------------------------
Function fnInit(obj)
	Dim ndRootTo
	dim sTemp
	Dim iTotQty,iNoofPack
	sTemp = document.formname.hSUBC.value
	set objTemp = window.dialogArguments

	Set Root = objTemp.documentElement
	arrTemp = split(obj,":")
	iLineNo = arrTemp(1)
	sLot = arrTemp(2)
	sStore = arrTemp(3)
	sBin = arrTemp(4)
	iClass = arrTemp(6)
	iItem = arrTemp(7)
	sOrgID = arrTemp(8)
	iEntNo = arrTemp(10)
	sOptName= arrTemp(11)
	ii = 0

	if trim(Root.getAttribute("NoofPack"))<>"" then
		document.formname.hRowSelect.value = Root.getAttribute("NoofPack")
	else
		document.formname.hRowSelect.value = 0
	end if ' if trim(Root.getAttribute("NoofPack"))<>"" then
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPack in HeaderNode.childNodes
					if ndPack.nodeName="Selection" and ndPack.getAttribute("YesNo")="Y" then
						document.formname.selTobox.length=document.formname.selTobox.length+1
						document.formname.selTobox(document.formname.selTobox.length-1).value = ndPack.getAttribute("SerialNo")
						document.formname.selTobox(document.formname.selTobox.length-1).text = ndPack.getAttribute("SerialNo")
						iTotQty = cdbl(iTotQty)+cdbl(ndPack.getAttribute("Qty"))
						iNoofPack=iNoofPack + 1
					end if'if ndPack.nodeName="Selection" and ndPack.getAttribute("YesNo")="Y" then
				next ' if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			end if
		next
	end if 'if ndRoot.hasChildNodes() then
	
	spaNoofPackSelected.innerText = FormatNumber(iTotQty,3)&"["&iNoofPack&"]"
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPack in HeaderNode.childNodes
					if ndPack.nodeName="Selection" and ndPack.getAttribute("YesNo")="N" then
						HeaderNode.removeChild(ndPack)
					end if'if ndPack.nodeName="Selection" and ndPack.getAttribute("YesNo")="Y" then
				next
			end if 'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
		next
	end if 'if ndRoot.hasChildNodes() then
	
end Function
'--------------------------------------------------------------------
function CheckSer()
	Dim iCnt, iCtr, sFlag,chkCount
	iCnt = document.formname.hiCtr.value
	'chkCount =document.formname.hRowSelect.value
	For iCtr = 1 to iCnt
		if eval("document.formname.chkSer"&iCtr).checked then
			sFlag = True
			Set iQty = 	eval("document.formname.txtQty"&iCtr)
			Set iStQty = eval("document.formname.txtStQty"&iCtr)
			iQty.value = iStQty.value
			chkCount = chkCount + 1
		else
			sFlag = False
			Set iQty = 	eval("document.formname.txtQty"&iCtr)
			iQty.value = 0
		'	chkCount = chkCount - 1
		end if
		eval("document.formname.chkSer"&iCtr).checked = sFlag	
	next
	document.formname.hRowSelect.value= chkCount
	'NoofPack.innerText = document.formname.hRowSelect.value
end function
'**********************************************************************
Function btnAddToList_Click()
	Dim iCnt, iCtr, sFlag,chkCount,iTotalQuantity,iSerNo,iPackSerNo
	Dim ndPack,ndPick
	
	iCnt = document.formname.hiCtr.value
	
	if trim(iCnt)="" or trim(iCnt)="0"  then
		alert("Select the Pack Numbers")
		exit function
	end if

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			set ndPick = HeaderNode
		end if
	Next
	
	'chkCount =document.formname.hRowSelect.value
	For iCtr = 1 to iCnt
		if eval("document.formname.chkSer"&iCtr).checked then
			sFlag = True
			Set iQty = 	eval("document.formname.txtQty"&iCtr)
			Set iStQty = eval("document.formname.txtStQty"&iCtr)
			iQty.value = iStQty.value
			
			iSerNo = eval("document.formname.hSerial"&iCtr).value
			
			if cdbl(document.formname.selTobox.length)>0 then
				for iPackSerNo = 0 to cdbl(document.formname.selTobox.length)-1
					if trim(document.formname.selTobox(iPackSerNo).value) = trim(iSerNo) then
						alert("Pack Number is Already Selected")
						exit function
					end if
				next
				
				document.formname.selTobox.length=document.formname.selTobox.length+1
				document.formname.selTobox(document.formname.selTobox.length-1).value=eval("document.formname.hSerial"&iCtr).value
				document.formname.selTobox(document.formname.selTobox.length-1).text=eval("document.formname.hSerial"&iCtr).value
				
				set ndPack = PackToData.createElement("Selection")
					ndPack.setAttribute "SerialNo",iSerNo
					ndPack.setAttribute "Qty",iStQty.Value
					ndPack.setAttribute "YesNo","Y"
				ndPick.appendChild ndPack
				
			else
				document.formname.selTobox.length=document.formname.selTobox.length+1
				document.formname.selTobox(document.formname.selTobox.length-1).value=eval("document.formname.hSerial"&iCtr).value
				document.formname.selTobox(document.formname.selTobox.length-1).text=eval("document.formname.hSerial"&iCtr).value
				
				set ndPack = PackToData.createElement("Selection")
					ndPack.setAttribute "SerialNo",iSerNo
					ndPack.setAttribute "Qty",iStQty.Value
					ndPack.setAttribute "YesNo","Y"
				ndPick.appendChild ndPack
				
			end if
		else
			sFlag = False
			Set iQty = 	eval("document.formname.txtQty"&iCtr)
			iQty.value = 0
		'	chkCount = chkCount - 1
		end if
		eval("document.formname.chkSer"&iCtr).checked = sFlag	
	next
	
	
	iTotalQuantity=0
	chkCount=0
	
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPack in HeaderNode.childNodes
					if trim(ndPack.nodeName)="Selection" then
						if ndPack.getAttribute("YesNo")="Y" then
							iTotalQuantity = cdbl(iTotalQuantity) + cdbl(ndPack.getAttribute("Qty"))
							chkCount = chkCount + 1
						end if
					end if'if trim(ndPack.nodeName)="Selection" then
				next
			end if
		Next
	end if
	
	if iTotalQuantity>0 then
		spaNoofPackSelected.innerText = FormatNumber(iTotalQuantity,3)&"["&chkCount&"]"
	else
		alert("Select the Pack Numbers")
		exit function
	end if
	
	For iCtr = 1 to iCnt
		eval("document.formname.chkSer"&iCtr).checked=false
		eval("document.formname.txtQty"&iCtr).value = 0
	next

End Function
'----------------------------------------------------------------------
Function btnRemove_Click()
	Dim ndRootFrom,ndRootTo,ndPackFrom,ndPackTo,ndTempNode
	Dim iCnt
	for iCnt = 0 to eval(document.formname.selFrombox.length)-1
		if Root.hasChildNodes() then
			For Each HeaderNode In Root.childNodes
				if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
					for each ndPack in HeaderNode.childNodes
						if trim(ndPack.nodeName)="Selection" then
							if ndPack.getAttribute("SerialNo")=trim(document.formname.selFrombox(iCnt).value) then
								ndPack.setAttribute "YesNo","N"	
								exit for
							end if
						end if'if trim(ndPack.nodeName)="Selection" then
					next
				end if
			Next
		end if
	next
	
	iTotalQuantity=0 
	chkCount=0
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPack in HeaderNode.childNodes
					if trim(ndPack.nodeName)="Selection" then
						if ndPack.getAttribute("YesNo")="Y" then
							iTotalQuantity = cdbl(iTotalQuantity) + cdbl(ndPack.getAttribute("Qty"))
							chkCount = chkCount + 1
						end if
					end if'if trim(ndPack.nodeName)="Selection" then
				next
			end if
		Next
	end if
	
	if iTotalQuantity>0 then
		spaNoofPackSelected.innerText = FormatNumber(iTotalQuantity,3)&"["&chkCount&"]"
	end if
	
End Function
'-------------------------------------------------------------------------
Function btnAdd_Click()
	Dim ndRootFrom,ndRootTo,ndPackFrom,ndPackTo,ndTempNode
	Dim iCnt
	for iCnt = 0 to eval(document.formname.selTobox.length)-1
		if Root.hasChildNodes() then
			For Each HeaderNode In Root.childNodes
				if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
					for each ndPack in HeaderNode.childNodes
						if trim(ndPack.nodeName)="Selection" then
							if ndPack.getAttribute("SerialNo")=trim(eval("document.formname.selTobox("&iCnt&")").value) then
								ndPack.setAttribute "YesNo","Y"	
								exit for
							end if
						end if'if trim(ndPack.nodeName)="Selection" then
					next
				end if
			Next
		end if
	next
	
	iTotalQuantity=0 
	chkCount=0
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPack in HeaderNode.childNodes
					if trim(ndPack.nodeName)="Selection" then
						if ndPack.getAttribute("YesNo")="Y" then
							iTotalQuantity = cdbl(iTotalQuantity) + cdbl(ndPack.getAttribute("Qty"))
							chkCount = chkCount + 1
						end if
					end if'if trim(ndPack.nodeName)="Selection" then
				next
			end if
		Next
	end if
	
	if iTotalQuantity>0 then
		spaNoofPackSelected.innerText = FormatNumber(iTotalQuantity,3)&"["&chkCount&"]"
	end if
	
	
End Function
'***************************************************************
Function CheckSubmit()
	dim ictr,objQ,iQtyTot,objSTQ,objSerial,i,iNoofPack
	dim sTemp,nSerialNumber
	Dim ndRootTo,ndPack
	iNoofPack = 0
	
	ictr = document.formname.selTobox.length-1

	sTemp = document.formname.hSUBC.value
	if lcase(sTemp) = "no" then
		if Root.hasChildNodes() then
			for each HeaderNode	in Root.childNodes
				if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
					for each ndPack in HeaderNode.childNodes
						if trim(ndPack.nodeName)="Selection" and ndPack.getAttribute("YesNo")="Y" then
							iQtyTot = cdbl(iQtyTot) + cdbl(ndPack.getAttribute("Qty"))
							iNoofPack = iNoofPack + 1
						end if 'if trim(ndPack.nodeName)="Selection" and ndPack.getAttribute("YesNo")="Y" then
					next
				end if'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			next
		end if 'if ndRootTo.hasChildNodes() then
	end if
	
	Root.setAttribute "NoofPack", iNoofPack
	
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			if HeaderNode.HaschildNodes() then
				For Each HNode In HeaderNode.childNodes
					if StrComp(Trim(HNode.NodeName),"SERIALHEADER") = 0 then
						set a = HeaderNode.removeChild(HNode)
					end if
				next
			end if
		end if
	Next

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			Set newElem = objTemp.createElement("SERIALHEADER")
				for each ndHeader in Root.childNodes
					if ndHeader.Attributes.getNamedItem("LOTNO").Value = sLot and ndHeader.Attributes.getNamedItem("LOC").Value = sStore and ndHeader.Attributes.getNamedItem("BIN").Value = sBin  then
						for each ndPack in ndHeader.childNodes
							if trim(ndPack.nodeName)="Selection" then
								if trim(ndPack.getAttribute("YesNo"))="Y" then
									Set newElem1 = objTemp.createElement("SERIALDETAILS")
									nSerialNumber = ndPack.getAttribute("SerialNo")
							
									if trim(nSerialNumber)="" or isNull(nSerialNumber) then nSerialNumber="NULL"
							
									newElem1.setAttribute "SERIALNO", nSerialNumber
									newElem1.setAttribute "QTY", ndPack.getAttribute("Qty")
									newElem.appendChild newElem1
								end if 'if trim(ndPack.getAttribute("YesNo"))="Y" then
							end if 'if trim(ndPack.nodeName)="Selection" then
						next
					end if 'if ndHeader.Attributes.getNamedItem("LOTNO").Value = sLot and ndHeader.Attributes.getNamedItem("LOC").Value = sStore and ndHeader.Attributes.getNamedItem("BIN").Value = sBin  then
				next
			HeaderNode.appendChild newElem
		end if
	Next

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			HeaderNode.Attributes.getNamedItem("QTYISS").value = cdbl(iQtyTot)
		end if
	Next
	
	
	
	sRet="DONE```YES"
	
	sSearchBy = document.formname.selSearchBy(document.formname.selSearchBy.selectedIndex).value 
	sSearchFor = document.formname.txtSearchFor.value 
	sSearchType = document.formname.selSearchType.value 
	sTemp = document.formname.hTemp.value
	iCurr=document.formname.hCurrentPage.value
	sRowCount=document.formname.hRowCount.value
	
	sValue = sTemp&"|SearchBy="&sSearchBy&"|SearchFor="&sSearchFor&"|SearchType="&sSearchType&"|hCurrentPage="&iCurr&"|hWho="&sRet&"|hSubmit="&sPage&"|hRowCount="&sRowCount
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPagination in HeaderNode.childNodes 
					if ndPagination.nodeName="Pagination" then
						HeaderNode.removeChild(ndPagination)
					end if
				next
			end if 'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
		next
	end if
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				set ndPagination = objTemp.createElement("Pagination")
				ndPagination.setAttribute "Details",sValue
				HeaderNode.appendChild ndPagination
			end if 'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
		next
	end if
	
	window.close
	exit function
end Function
'****************************************************************
Function btnClose_Click()
	dim ictr,objQ,iQtyTot,objSTQ,objSerial,i,iNoofPack
	dim sTemp,nSerialNumber
	Dim ndRootTo,ndPack
	iNoofPack = 0
	
	sRet="DONE```NO"
	
	sSearchBy = document.formname.selSearchBy(document.formname.selSearchBy.selectedIndex).value 
	sSearchFor = document.formname.txtSearchFor.value 
	sSearchType = document.formname.selSearchType.value 
	sTemp = document.formname.hTemp.value
	iCurr=document.formname.hCurrentPage.value
	sRowCount=document.formname.hRowCount.value
	
	sValue = sTemp&"|SearchBy="&sSearchBy&"|SearchFor="&sSearchFor&"|SearchType="&sSearchType&"|hCurrentPage="&iCurr&"|hWho="&sRet&"|hSubmit="&sPage&"|hRowCount="&sRowCount
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPagination in HeaderNode.childNodes 
					if ndPagination.nodeName="Pagination" then
						HeaderNode.removeChild(ndPagination)
					end if
				next
			end if 'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
		next
	end if
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				set ndPagination = objTemp.createElement("Pagination")
				ndPagination.setAttribute "Details",sValue
				HeaderNode.appendChild ndPagination
			end if 'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
		next
	end if
	
	window.close
	exit function
end Function
'***************************************************************
Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

'**********************************************************************
Function NextSelection(sPage)
	Dim sSearchFor,sSearchBy,sSearchType,sTemp,sValue,sRet
	Dim objhttp,ndPick,ndPagination
	Dim iCtr,sAttributeName,sSerialNo,sQty,oRow,oCell,sInvNo
	Dim iCurr,sQuery
	
	sRet="NEXT"
	
	sSearchBy = document.formname.selSearchBy(document.formname.selSearchBy.selectedIndex).value 
	sSearchFor = document.formname.txtSearchFor.value 
	sSearchType = document.formname.selSearchType.value 
	sTemp = document.formname.hTemp.value
	iCurr=document.formname.hCurrentPage.value
	sRowCount=document.formname.hRowCount.value
	
	sValue = sTemp&"|SearchBy="&sSearchBy&"|SearchFor="&sSearchFor&"|SearchType="&sSearchType&"|hCurrentPage="&iCurr&"|hWho="&sRet&"|hSubmit="&sPage&"|hRowCount="&sRowCount
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				for each ndPagination in HeaderNode.childNodes 
					if ndPagination.nodeName="Pagination" then
						HeaderNode.removeChild(ndPagination)
					end if
				next
			end if 'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
		next
	end if
	
	if Root.hasChildNodes() then
		for each HeaderNode in Root.childNodes
			if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
				set ndPagination = objTemp.createElement("Pagination")
				ndPagination.setAttribute "Details",sValue
				HeaderNode.appendChild ndPagination
			end if 'if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
		next
	end if
	window.close
End Function
'*********************************************************
