dim OutDataValue

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

Function DisableTxt(obj)
	arrTemp = split(obj.name,"n")
	
	set Q = eval("document.formname.txtTare"&arrTemp(1))
	
	Q.value = ""
	if obj.value = "I" then
		Q.disabled = true
	else
		Q.disabled = false
	end if
end Function

Function GetLot(obj)

	dim sItem,sClass,sOrgID,aobj
	arrTemp = split(obj.name,":")
	
	sClass = arrTemp(1)
	sItem = arrTemp(2)
	sOrgID = arrTemp(3)
	iQty = arrTemp(4)
	sVar = "btn:"&sClass&":"&sItem&":"&sOrgID
	
	dim i
	dim arrTemp
	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "button") then
			if (Instr(document.formname.elements(i).name,sVar) > 0) then
				if Trim(document.formname.elements(i-4).value) <> "N" then
					iValue = "0"
				end if
			end if
		end if
	next
	if iValue = "" then iValue = "N"
	sTempValues = sItem&":"&sClass&":"&sOrgID&":"&iQty&":"&iValue

	Set OutDataValue = showModalDialog("receiptStorePop.asp?sTemp="&sTempValues,Data,"dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	
end Function

Function DisplayItem(obj)
	sTempValues = obj

	showModalDialog "itmDetailsPop.asp?sTemp="&sTempValues,document.formname.hOrgName.value,"dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
end Function

Function CheckSubmit(todaysdate,sReturnTo)
	dim i,iTotal, iItemStr,sItmType
	iItemStr = ""	
	iTotal = 0
	dim arrTemp
	Set Root = Data.documentElement
	'sTempDate = document.formname.hGDate.value
	sTempDate = document.formname.ctlDDate.GetDate
	if(datediff("d",sTempDate,document.formname.ctlDDate.GetDate)) < 0 then
		alert("Date should be greater than or equal to GRN Date")
		exit function
	end if
	if(datediff("d",todaysdate,document.formname.ctlDDate.GetDate)) > 0 then
		alert("Date should be less than or equal to "&todaysdate)
		exit function
	end if
	If datediff("d",document.formname.ctlDDate.GetDate,document.formname.hFinTo.Value)< 0 or datediff("d",document.formname.ctlDDate.GetDate,document.formname.hFinFrom.Value) > 0 Then
		alert("Date should be Selected Between " & document.formname.hFinFrom.Value & " and " & document.formname.hFinTo.Value)		
		Exit Function
	End If
	
	
	sExp ="//ITEM"
	Set ItemNode = Root.Selectnodes(sExp)

	For iCounter = 0 to ItemNode.Length - 1
		iTotStQty = 0
		iItem = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEM").Value
		iClass = ItemNode.Item(iCounter).Attributes.getNamedItem("CLASS").Value
		iQtyRec = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEMQTY").Value
		iItemValue = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEMVALUE").Value
		
		if iItemValue<>0 then
			iItemRate = cdbl(iItemValue)/cdbl(iQtyRec)
		end if

		sItemName = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").Value
		
		sExp1 ="//ITEM [ @ITEM = "&iItem&" and @CLASS = "&iClass&"]/STORAGE"
		Set StorageNode = Root.Selectnodes(sExp1)
		
		For iCtr = 0 to StorageNode.Length - 1
			iTotStQty = cdbl(iTotStQty) + cdbl(StorageNode.Item(iCtr).Attributes.getNamedItem("STOREQTY").Value)
		next
		
		if cdbl(iQtyRec) <> cdbl(iTotStQty) then
			alert("Total Quantity received should be accounted fully for the Item "& sItemName)
			exit function
		end if
		If cDbl(iItemRate) = 0 Then
			If iItemStr = "" Then
				iItemStr = sItemName
			Else
				iItemStr = iItemStr & "," & sItemName
			End If
		End If			
		iTotal = cdbl(iTotal) + cdbl(iTotStQty)
	next

	if cdbl(iTotal) <= 0 then
		alert("Cannot account ZERO Quantity")
		exit function
	end if

'	If iItemStr <> "" Then		
'		alert ("Item Value Can't Be Zero while Receipt accounting. Please Amend the Item values in Actual Receipt Stage And Come Back for accounting...!")
'		Exit Function	
'	End If 

	
	Root.setAttribute "RECEIVEDON", document.formname.ctlDDate.GetDate
	
'	document.formname.B1.disabled = True
	
	Set objhttp = CreateObject("Microsoft.XMLHTTP")

	'storing data - for testing
	'objhttp.Open "POST","XMLSaveInv.asp?Name=InvReceipt&Mod=INV", False
	'objhttp.send Data.XML
	
	'alert Data.xml
	'exit function

	objhttp.Open "POST","receiptInsert.asp", false
	objhttp.send Data.XMLDocument

'	msgbox objhttp.responseText
'	exit function

	if trim(objhttp.responseText)="" then
		alert("Receipt has been done Successfully")
		window.location.href = "MATERIALRECEIPTS.ASP?RCPT=A"
	else
		alert(objhttp.responseText)
		document.formname.B1.disabled = False
	end if
	
end Function

Function Init()
	Dim i,j, gDate
	document.formname.ctlDDate.setMindate = document.formname.hFinFrom.value
	document.formname.ctlDDate.SetDate = document.formname.hFormattedGDate.value
	j = 0
	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "button") then
			if (Instr(document.formname.elements(i).name,"btn") > 0) then
				'alert document.formname.elements(i).name
				j = j + 1
				arrTemp = split(document.formname.elements(i).name,":")
				'alert sClass
				sClass = arrTemp(1)
				sItem = arrTemp(2)
				sOrgID = arrTemp(3)
				document.formname.ctlDDate.setDate = arrTemp(5)	
				'To get whether the Quantity is Gross/Nett
				Dim objHttp
				Set objhttp = CreateObject("Microsoft.XMLHTTP")

				objhttp.Open "GET","XMLReceiptQuantity.asp?Item="&sItem&"&Class="&sClass&"&OrgID="&sOrgID&"&ReceiptNo="&document.formname.hiRecNo.value, false
				objhttp.send 
				if Trim(objHttp.responseXml.xml)<>"" then
					TempData.loadxml objHttp.responseXml.Xml
					set Root = TempData.documentElement
				else
					alert("File is Not Found or It is returning Invalid Data - XMLReceiptQuantity.asp")
				end if
				
'				If document.formname.hGDate.value <> "" Then
'					document.formname.ctlDDate.setDate = document.formname.hGDate.value
'				End If

				if Trim(Root.attributes.getNamedItem("QUANTITYIN").value) = "G" then
				'	eval("document.formname.selQtyIn"&j).selectedIndex = 1
				'	eval("document.formname.selQtyIn"&j).disabled = True
				elseif Trim(Root.attributes.getNamedItem("QUANTITYIN").value) = "N" then
				'	eval("document.formname.selQtyIn"&j).selectedIndex = 2
				'	eval("document.formname.selQtyIn"&j).disabled = True
				end if
				
			end if
		end if
	next
	
	set Root = Data.documentElement
	dGDate = Root.getAttribute("RECEIVEDON")
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.nodeName = "ITEM" then
			sTempObjVal= ""
				iItem = HeaderNode.getAttribute("ICODE")
				iClass = HeaderNode.getAttribute("CCODE")
				iQtyAcc = HeaderNode.getAttribute("QTYACC")
				sOCode = HeaderNode.getAttribute("OCODE")
				sTempObjVal = "txtZ"&iClass&":"&iItem&":"&sOCode&":"&iQtyAcc&":"&dGDate
					For Each HNode In HeaderNode.childNodes
						for i=0 to document.formname.elements.length - 1
							if (document.formname.elements(i).type = "text") then
								if (Instr(document.formname.elements(i).name,cstr(trim(sTempObjVal))) > 0) then
									document.formname.elements(i).value = cdbl(FormatNumber(HNode.getAttribute("QTY"),3,,,0))
								end if
							end if
						next
					next
			end if
		Next
	end if'if Root.hasChildNodes() then

end Function

