dim objTemp,Root,newElem
dim iClass,iItem,iQty,ii

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

Function fnInit(sItem,sClass,sQty,sEntNo)
	iClass = sClass
	iItem = sItem
	iQty = sQty
	iEntNo = sEntNo
	ictr = document.formname.hiCtr.value

	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement

	if document.formname.hiCtr.value = 0 then exit function

	For Each HeaderNode In Root.childNodes
		'if HeaderNode.Attributes.Item(0).nodeValue = iClass and HeaderNode.Attributes.Item(1).nodeValue = iItem then
		if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = iItem and HeaderNode.Attributes.Item(2).nodeValue = iClass then
			if HeaderNode.HaschildNodes() then
				For Each CNode In HeaderNode.childNodes 
					if StrComp(Trim(CNode.NodeName),"STDETAILS") = 0 then
						ii = ii + 1
						set Q = eval("document.formname.txtST"&ii)
						Q.value = CNode.Attributes.Item(1).nodeValue
					end if
				next
			else
				for ii = 1 to document.formname.hiCtr.value
					set Q = eval("document.formname.txtST"&ii)
					Q.value = "0"
				next
			end if
		end if
	Next
end Function

Function CheckSubmit()
	dim ictr,objQ,iQtyTot
	
	ictr = document.formname.hiCtr.value
	'alert(ictr)
	if ictr = "" then exit function 
	
	for i=1 to ictr
		set objQ = eval("document.formname.txtST"&i)
		set objh = eval("document.formname.hST"&i)
 
		if trim(objQ.value) = "" then
			msgbox "Enter Quantity",0,"Quantity"
			objQ.select()
			exit function
		elseif not checkNumbers(objQ.value) then
			msgbox "Enter Numerals Only",0,"Numerals"
			objQ.select()
			exit function
		else
			sQty = objQ.value
			if (cdbl(sQty) > cdbl(objh.value)) then
				msgbox "Transfer Quantity should be equal to or less than Stock Quantity "&objh.value,0,"Quantity"
				objQ.select()
				exit function
			end if

			iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
		end if

	next

	Set Root = objTemp.documentElement
	For Each HeaderNode In Root.childNodes
		'if HeaderNode.Attributes.Item(0).nodeValue = iClass and HeaderNode.Attributes.Item(1).nodeValue = iItem then
		if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = iItem and HeaderNode.Attributes.Item(2).nodeValue = iClass then
			if HeaderNode.HaschildNodes() then
				For Each HNode In HeaderNode.childNodes
					if StrComp(Trim(HNode.NodeName),"STDETAILS") = 0 then
						set a = HeaderNode.removeChild(HNode)
					end if
				next
			end if
		end if
	Next

	if not iQtyTot = 0 then
		Set Root = objTemp.documentElement
			 
		For Each HeaderNode In Root.childNodes
			'if HeaderNode.Attributes.Item(2).nodeValue = iClass and HeaderNode.Attributes.Item(1).nodeValue = iItem then
			if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = iItem and HeaderNode.Attributes.Item(2).nodeValue = iClass then
				for i=1 to ictr
					Set newElem = objTemp.createElement("STDETAILS")
					newElem.setAttribute "ORGID", trim(eval("document.formname.hOrgID"&i&".value"))
					newElem.setAttribute "QTY", trim(eval("document.formname.txtST"&i&".value"))
					HeaderNode.appendChild newElem
				next
			end if
		Next
	end if
	If trim(document.formname.hMrsNo.value) <> "" then 
		if (cdbl(iQtyTot) > cdbl(iQty)) then
			msgbox "Total Transfer Quantity should be equal to or less than Quantity Pending " &iQty ,0,"Quantity"
			exit function
		end if
	End If
	Set Root = objTemp.documentElement
	For Each HeaderNode In Root.childNodes
		'if HeaderNode.Attributes.Item(0).nodeValue = iClass and HeaderNode.Attributes.Item(1).nodeValue = iItem then
		if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = iItem and HeaderNode.Attributes.Item(2).nodeValue = iClass then
			HeaderNode.setAttribute "TRAQTY", cdbl(iQtyTot)
			window.close
			exit function
		end if
	Next
	alert(objTemp.xml)
	 window_onunload() 
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function
