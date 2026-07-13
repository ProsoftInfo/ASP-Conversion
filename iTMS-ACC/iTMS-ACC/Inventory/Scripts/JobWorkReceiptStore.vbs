dim objTemp,Root,newElem
dim iClass,iItem,iIssNo,iMRSNo,sOrgID,iStore,iBin,iTareValue

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
	dim CheckQtyFlag
'alert "AAA"
	arrTemp = split(obj,":")
	iItem = arrTemp(0)
	iClass = arrTemp(1)
	sOrgID = arrTemp(2)
	iTareValue = arrTemp(4)
	
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
'alert document.formname.hiCtr.value
	if document.formname.hiCtr.value = 0 then 
		exit function
	end if
	
	For Each HeaderNode In Root.childNodes
		if HeaderNode.nodeName = "ITEM" then
			if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass then
				For Each HNode In HeaderNode.childNodes
					ii = ii + 1
					set Q = eval("document.formname.txtQty"&ii)
					Q.value = HNode.Attributes.Item(2).nodeValue
					if HNode.Attributes.Item(2).nodeValue > 0 then 
						CheckQtyFlag = True
					else
						CheckQtyFlag = False
					end if
				next
			end if
		end if
	Next

	iRecNo = Root.attributes.getNamedItem("RECNO").value

	if CheckQtyFlag then
		GetLotDetails iRecNo,iItem,iClass,sOrgID,"N"
	else
		GetLotDetails iRecNo,iItem,iClass,sOrgID,"Y"
	end if
	
end Function

Function GetLot(obj,str,iCtr,sStoName)
	dim sItem,sClass,sOrgID,aobj
	arrTemp = split(obj.name,":")
	
	sClass = arrTemp(1)
	sItem = arrTemp(2)
	sOrgID = arrTemp(3)
	iStore = arrTemp(4)
	iBin = arrTemp(5)
	
	sTempValues = str&":"&sItem&":"&sClass&":"&sOrgID&":"&iStore&":"&iBin&":"&sStoName&":"&iTareValue & ":True"

'	alert objTemp.xml
	
	Set OutDataValue = showModalDialog("JobWorkReceiptLotSerPop.asp?sTemp="&sTempValues,objTemp,"dialogHeight:370px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
	
'	alert objTemp.xml
	
	For Each HeaderNode In Root.childNodes
		if HeaderNode.nodeName = "ITEM" then
			if HeaderNode.Attributes.getNamedItem("ICODE").Value = sItem and HeaderNode.Attributes.getNamedItem("CCODE").Value = sClass then
				if HeaderNode.HaschildNodes() then
					For Each HNode In HeaderNode.childNodes
						if HNode.Attributes.Item(0).nodeValue = iStore and HNode.Attributes.Item(1).nodeValue = iBin then
							set Q = eval("document.formname.txtQty"&iCtr)
							Q.value = HNode.Attributes.Item(2).nodeValue
						end if
					next
				end if
			end if 
		end if
	Next
end Function

Function GetLotDetails(iRecNo,iItem,iClass,sOrgID,sFlag)
	dim objhttp
	
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","XMLGetJobWorkLotDetails.asp?iRecNo="&iRecNo&"&iItem="&iItem&"&iClass="&iClass&"&sOrgID="&sOrgID&"&sFlag="&sFlag&"&sRecType="&document.formname.hRec.value, false

	objhttp.send 
	
	'alert objhttp.responsetext
	'alert root.xml

	if objhttp.responseXML.xml <> "" then
		XmlData.loadXML objhttp.responseXML.xml
		Set RootData = XmlData.documentElement
		Root.appendChild RootData
	end if

'	alert root.xml
'	exit function
End Function

Function CheckSubmit()
	dim ictr,objQ,iQtyTot,objSTQ,objSerial

	if document.formname.hiCtr.value = 0 then 
		window.close
		exit function
	end if
	
	ictr = document.formname.hiCtr.value

	for i=1 to ictr
		set objQ = eval("document.formname.txtQty"&i)

		if trim(objQ.value) = "" then
			msgbox "Enter Quantity",0,"Quantity"
			objQ.select()
			exit function
		elseif not checkNumbers(objQ.value) then
			alert("Enter Numerals Only")
			objQ.select()
			exit function
		else
			iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
		end if

	next

	if (cdbl(iQtyTot) <> cdbl(idQty.innerText)) then
		alert("Total Quantity should be equal to Quantity to Account (" &trim(idQty.innerText)& ")")
		exit function
	end if

	ictr = 0

	Set RootO = objTemp.documentElement
	
	For Each HeaderNode In RootO.childNodes
		if HeaderNode.nodeName="ITEM" then
			if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass then
				For Each HNode In HeaderNode.childNodes
					ii = ii + 1
					set Q = eval("document.formname.txtQty"&ii)
					HNode.setAttribute "QTY", cdbl(Q.value)	
					HNode.setAttribute "RECTYPE", document.formname.hRec.value	
					HNode.setAttribute "LOTNUMBER", idLotNumber.innerText
				next
			end if
		end if
	Next
	
'	alert RootO.xml
	window.close

end Function

Function window_onunload() 
	dim ictr,objQ,iQtyTot,objSTQ,objSerial

	Set RootO = objTemp.documentElement
	
	if document.formname.hiCtr.value = 0 then 
		set window.returnValue = objTemp.documentElement
		window.close()
	end if

	For Each HeaderNode In RootO.childNodes
		if HeaderNode.nodeName = "STOREDLOTDETAILS" then
			RootO.removechild(HeaderNode)
		end if
	next

	ictr = document.formname.hiCtr.value
	for i=1 to ictr
		set objQ = eval("document.formname.txtQty"&i)
		iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
	next

	'Removing if Entered Qty is not equal to Qty
	if (cdbl(iQtyTot) <> cdbl(idQty.innerText)) then
		Set RootO = objTemp.documentElement
'		alert RootO.xml
		For Each HeaderNode In RootO.childNodes
			if HeaderNode.nodeName="ITEM" then
				if HeaderNode.Attributes.Item(0).nodeValue = iItem and HeaderNode.Attributes.Item(1).nodeValue = iClass then
					For Each StorageNode In HeaderNode.childNodes
						For Each LotNode in StorageNode.childNodes
							StorageNode.attributes.getNamedItem("QTY").value = 0
							StorageNode.removeChild(LotNode)
						next
					next
				end if
			end if
		next
	end if
	
'	alert RootO.xml
	
	set window.returnValue = objTemp.documentElement
	window.close()
end Function
