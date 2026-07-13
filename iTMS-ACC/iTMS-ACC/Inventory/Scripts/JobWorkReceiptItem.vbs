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
	sRec = arrTemp(5)
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
	sTempValues = sItem&":"&sClass&":"&sOrgID&":"&iQty&":"&iValue&":"&sRec

	Set OutDataValue = showModalDialog("JobWorkReceiptStorePop.asp?sTemp="&sTempValues,Data,"dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	'window.open "JobWorkReceiptStorePop.asp?sTemp="&sTempValues,Data,""
	
end Function

Function DisplayItem(obj)
	sTempValues = obj

	showModalDialog "itmDetailsPop.asp?sTemp="&sTempValues,document.formname.hOrgName.value,"dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
end Function

Function CheckSubmit(todaysdate)
	dim i,iTotal
	iTotal = 0
	dim arrTemp
	
	sTempDate = document.formname.hGDate.value
	if(datediff("d",sTempDate,document.formname.ctlDDate.GetDate)) < 0 then
		alert("Date should be greater than or equal to GRN Date")
		exit function
	end if
	if(datediff("d",todaysdate,document.formname.ctlDDate.GetDate)) > 0 then
		alert("Date should be less than or equal to Today's Date")
		exit function
	end if

	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "button") then
			if (Instr(document.formname.elements(i).name,"btn") > 0) then
				arrTemp = split(document.formname.elements(i).name,":")

				sClass = arrTemp(1)
				sItem = arrTemp(2)
				sOrgID = arrTemp(3)

				Set Root = Data.documentElement
				For Each HeaderNode In Root.childNodes
					if HeaderNode.Attributes.Item(1).nodeValue = sClass and HeaderNode.Attributes.Item(0).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sOrgID then
						HeaderNode.Attributes.getNamedItem("PRUNFROM").value = trim(document.formname.elements(i-2).value)
						HeaderNode.Attributes.getNamedItem("PRUNTO").value = trim(document.formname.elements(i+2).value)
						HeaderNode.Attributes.getNamedItem("QTYIN").value = trim(document.formname.elements(i-1).value)
'						HeaderNode.Attributes.getNamedItem("TARE").value = trim(document.formname.elements(i-1).value)
						HeaderNode.Attributes.getNamedItem("PRESS").value = trim(document.formname.elements(i-3).value)
						HeaderNode.Attributes.getNamedItem("CROP").value = trim(document.formname.elements(i+1).value)
'						HeaderNode.Attributes.getNamedItem(BALES").value = trim(document.formname.elements(i+3).value)
'						HeaderNode.Attributes.getNamedItem("TAREWEIGHT").value = trim(document.formname.elements(i+4).value)
					end if
				Next
			end if
		end if
		
	next

	sExp ="//ITEM"
	Set ItemNode = Root.Selectnodes(sExp)

	For iCounter = 0 to ItemNode.Length - 1
		iTotStQty = 0
		iItem = ItemNode.Item(iCounter).Attributes.getNamedItem("ICODE").Value
		iClass = ItemNode.Item(iCounter).Attributes.getNamedItem("CCODE").Value
		iQtyRec = ItemNode.Item(iCounter).Attributes.getNamedItem("QTY").Value

		sItemName = ItemNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").Value

		sExp1 ="//ITEM [ @ICODE = "&iItem&" and @CCODE = "&iClass&"]/STORAGE"
		Set StorageNode = Root.Selectnodes(sExp1)
		
		For iCtr = 0 to StorageNode.Length - 1
			iTotStQty = cdbl(iTotStQty) + cdbl(StorageNode.Item(iCtr).Attributes.getNamedItem("QTY").Value)
		next
		
		if cdbl(iQtyRec) <> cdbl(iTotStQty) then
			alert("Total Quantity received should be accounted fully for the Item "& sItemName)
			exit function
		end if
		iTotal = cdbl(iTotal) + cdbl(iTotStQty)
	next

	if cdbl(iTotal) <= 0 then
		alert("Cannot account ZERO Quantity")
		exit function
	end if

	Set Root = Data.documentElement
	Root.setAttribute "RECEIVEDON", document.formname.ctlDDate.GetDate
	
	dim objHttp, Root, Node
	set objHttp = CreateObject("MSXML2.XMLHTTP")

	objHttp.open "POST","XMLSave.asp?SessionFlag=False&Value=JobWorkReceipt&Folder=Transaction",false
	objHttp.send Data.xml

	document.formname.B1.disabled = True

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","JobWorkReceiptInsert.asp", false
	objhttp.send Data.XMLDocument

	if objhttp.responseText = "Y<BR>" then
		alert("Receipt has been done Successfully")
		if confirm("Do you want to create another Receipt") then
			window.location.href = "MATERIALRECEIPTS.ASP"
		else
			window.location.href = "../welcome_Inventory.asp"
		end if
	else
		alert(objhttp.responseText)
		document.formname.B1.disabled = False
	end if

end Function

Function Init()
	Dim i,j
	
	j = 0
	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "button") then
			if (Instr(document.formname.elements(i).name,"btn") > 0) then
				j = j + 1
				arrTemp = split(document.formname.elements(i).name,":")

				sClass = arrTemp(1)
				sItem = arrTemp(2)
				sOrgID = arrTemp(3)

				'To get whether the Quantity is Gross/Nett
				Dim objHttp
				Set objhttp = CreateObject("Microsoft.XMLHTTP")

				objhttp.Open "GET","XMLJobWorkReceiptQuantity.asp?Item="&sItem&"&Class="&sClass&"&OrgID="&sOrgID&"&ReceiptNo="&document.formname.hiRecNo.value, false
				objhttp.send 
				
				TempData.loadxml objHttp.responseXml.Xml
				set Root = TempData.documentElement
				
				if Trim(Root.attributes.getNamedItem("QUANTITYIN").value) = "G" then
					eval("document.formname.selQtyIn"&j).selectedIndex = 1
					eval("document.formname.selQtyIn"&j).disabled = True
				elseif Trim(Root.attributes.getNamedItem("QUANTITYIN").value) = "N" then
					eval("document.formname.selQtyIn"&j).selectedIndex = 2
					eval("document.formname.selQtyIn"&j).disabled = True
				end if
				
				'Setting Values
				eval("document.formname.txtPressMark"&j).value = Trim(Root.attributes.getNamedItem("PRESSMARKNO").value)
				eval("document.formname.txtPressRunFro"&j).value = Trim(Root.attributes.getNamedItem("PRNOFROM").value)
				eval("document.formname.txtCropYear"&j).value = Trim(Root.attributes.getNamedItem("CROPYEAR").value)
				eval("document.formname.txtPressRunTo"&j).value = Trim(Root.attributes.getNamedItem("PRNOTO").value)
				
			end if
		end if
	next

end Function