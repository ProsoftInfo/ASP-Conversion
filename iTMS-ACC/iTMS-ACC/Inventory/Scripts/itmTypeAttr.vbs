dim OutDataValue,bFlag

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

Function checkSelect()
	Set Root = OutData.documentElement
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			set a = Root.removeChild(HeaderNode)
		Next
	end if		

	document.formname.txtDataLen.value = ""
	document.formname.txtDecimals.value = ""

	if (document.formname.selDataType(document.formname.selDataType.selectedIndex).value = "String") then
		document.formname.txtDataLen.disabled = false
		document.formname.txtDecimals.disabled = true
	elseif (document.formname.selDataType(document.formname.selDataType.selectedIndex).value = "Numeric") then
		document.formname.txtDataLen.disabled = false
		document.formname.txtDecimals.disabled = false
	elseif (document.formname.selDataType(document.formname.selDataType.selectedIndex).value = "Options") then
		document.formname.txtDataLen.disabled = true
		document.formname.txtDecimals.disabled = true

	'	if(document.formname.selType(document.formname.selType.selectedIndex).value = "select") then
	'		MsgBox "Select Type",0,"Type"
	'		document.formname.selType.focus()
	'		document.formname.selDataType.selectedIndex = 0
	'		Exit Function
	'	else
	    if(document.formname.selHeader(document.formname.selHeader.selectedIndex).value = "select") then
			MsgBox "Select Header",0,"Header"
			document.formname.selHeader.focus()
			document.formname.selDataType.selectedIndex = 0
			Exit Function
		elseif (Trim(document.formname.txtAttribute.value) = "") then
			MsgBox "Enter Attribute Name",0,"Attribute Name"
			document.formname.txtAttribute.select()
			document.formname.selDataType.selectedIndex = 0
			Exit Function
		end if

		sTempValues = document.formname.hItemType.value&":"&document.formname.selHeader(document.formname.selHeader.selectedIndex).text&":"&document.formname.txtAttribute.value
		set OutDataValue = showModalDialog("ItmTypOptionPoPEntry.asp?sTemp="&sTempValues,"","dialogHeight:270px;dialogWidth:435px;center:Yes;help:No;resizable:No;status:No")

		set Root1 = OutDataValue
		if not Root1.HaschildNodes() then
			document.formname.selDataType.selectedIndex = 0
		else
			bFlag = true
		end if		
	else
		document.formname.txtDataLen.disabled = false
		document.formname.txtDecimals.disabled = false
	end if
end Function
'=======================================
Function CheckSubmit()
'	if(document.formname.selType(document.formname.selType.selectedIndex).value = "select") then
'		MsgBox "Select Type",0,"Type"
'		document.formname.selType.focus()
'		Exit Function
'	else
    if(document.formname.selHeader(document.formname.selHeader.selectedIndex).value = "select") then
		MsgBox "Select Header",0,"Header"
		document.formname.selHeader.focus()
		Exit Function
	elseif (Trim(document.formname.txtAttribute.value) = "") then
		MsgBox "Enter Attribute Name",0,"Attribute Name"
		document.formname.txtAttribute.select()
		Exit Function
	elseif(document.formname.selDataType(document.formname.selDataType.selectedIndex).value = "select") then
		MsgBox "Select Data Type",0,"Data Type"
		document.formname.selDataType.focus()
		Exit Function
	elseif((document.formname.selDataType.selectedIndex = "1" or document.formname.selDataType.selectedIndex = "2") and (Trim(document.formname.txtDataLen.value) = "")) then
		MsgBox "Enter Data Length",0,"Data Length"
		document.formname.txtDataLen.select()
		Exit Function
	elseif((document.formname.selDataType.selectedIndex = "1" or document.formname.selDataType.selectedIndex = "2") and (Trim(document.formname.txtDataLen.value) <> "" and not checkNumbers(document.formname.txtDataLen.value))) then
		msgbox "Enter Numerals Only",0,"Numerals"
		document.formname.txtDataLen.select()
		Exit Function
	elseif((document.formname.selDataType.selectedIndex = "2") and (Trim(document.formname.txtDecimals.value) = "")) then
		MsgBox "Enter Decimal Length",0,"Decimal Length"
		document.formname.txtDecimals.select()
		Exit Function
	elseif((document.formname.selDataType.selectedIndex = "2") and (Trim(document.formname.txtDecimals.value) <> "" and not checkNumbers(document.formname.txtDecimals.value))) then
		msgbox "Enter Numerals Only",0,"Numerals"
		document.formname.txtDecimals.select()
		Exit Function
	else
		set Root = OutData.documentElement
		
		Root.setAttribute "ITYPE", trim(document.formname.hItemType.value)
		Root.setAttribute "HEADER", trim(document.formname.selHeader(document.formname.selHeader.selectedIndex).value)
		Root.setAttribute "ATTRNAME", trim(document.formname.txtAttribute.value)
		Root.setAttribute "DATATYPE", trim(document.formname.selDataType(document.formname.selDataType.selectedIndex).value)
		Root.setAttribute "DATALENGTH", trim(document.formname.txtDataLen.value)
		Root.setAttribute "DECIMAL", trim(document.formname.txtDecimals.value)
		Root.setAttribute "CLASSCODE", trim(document.formname.hClassCode.value)

		if bFlag then 
			set Root1 = OutDataValue
			if Root1.HaschildNodes() then
				For Each HeaderNode In Root1.childNodes
					Root.appendChild HeaderNode
				next
			end if		
		end if

		Set objhttp = CreateObject("Microsoft.XMLHTTP")
		
		objhttp.Open "POST","ItmTypeAttributeInsert.asp", false
		objhttp.send OutData.XMLDocument
		
		if objhttp.responseText = "Y<BR>" then
			Msgbox ("Item Type Attribute " + trim(document.formname.txtAttribute.value) + " has been defined Successfully")
			document.formname.submit
		elseif objhttp.responseText = "N<BR>" then
			Msgbox ("Item Type Attribute " + trim(document.formname.txtAttribute.value) + " has been already defined")
			document.formname.submit
		else
			MsgBox objhttp.responseText
		end if
		
	end if	
end Function
'==========================================================
Function Init()
	if cint(document.formname.hValue.value)>0 then	
		populateHeader(document.formname.hItemType.value)
		for iCnt = 0 to cint(document.formname.selHeader.length) -1
		    if strcomp(document.formname.hHeader.value,document.formname.selHeader(iCnt).value)=0 then
		        document.formname.selHeader.selectedIndex = iCnt
		    end if
	    next
	    for iCnt = 0 to cint(document.formname.selDataType.length) -1
            if StrComp(document.formname.hType.value,document.formname.selDataType(iCnt).value)=0 then
                document.formname.selDataType.selectedIndex = iCnt
            end if
        next
        document.formname.txtAttribute.value = document.formname.hAttribute.value
	    document.formname.txtDataLen.value = document.formname.hLength.value
	    document.formname.txtDecimals.value = document.formname.hDecimal.value
	else
	    populateHeader(document.formname.hItemType.value)
	end if
End Function
'========================================================
Function ClearTable()
	Dim nRow 
	nRow = cint(document.formname.hRow.value)
	for i=1 to nRow
		document.all.tblDisplay.deleteRow(1)
	next
	
	for i=1 to nRow
		dofument.all.tblAttSelect.deleteRow(1)
	next
End Function
'======================================================
Function populateUpdate()
Dim oRow,Root,AttNode,iSno
iSno = 0
ClearTable()

'populateHeader(document.formname.hItemType.value)
set objhttp = CreateObject("Microsoft.XMLHTTP")
objhttp.open "GET","ItmTypeAttributeDisplayEdit.asp?ItemType="& document.formname.hItemType.value&"&ClassCode="& document.formname.hClassCode.value ,false
objhttp.send
'alert(objhttp.responseXML.xml)
	if objhttp.responseXML.xml<>"" then
		EditData.loadXML objhttp.responseXML.xml
	else
		alert(objhttp.responseText)
	end if
	
	set Root = EditData.documentElement
	if Root.hasChildNodes() then
		for each AttNode in Root.childNodes
		iSno = iSno + 1
			set oRow = document.all.tblDisplay.insertRow(document.all.tblDisplay.Rows.length)
			
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = iSno
				CurrCell.className="ExcelHeaderCell"
				CurrCell.width = "10"
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=button name='btn"&iSno&"' value ='Edit' class='ActionButtonX' onClick='EditAttribute("""&iSno&""")'>")
				CurrCell.appendChild(oText)
				CurrCell.className="ExcelDisplayCell"
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=checkbox name='Chk"&iSno&"' value ='"& iSno &"' class='FormElem'>")
				CurrCell.appendChild(oText) 
				CurrCell.align = "Center"
				CurrCell.className="ExcelDisplayCell"
				
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = AttNode.getAttribute("HName")
				CurrCell.className="ExcelDisplayCell"
				
				
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = AttNode.getAttribute("Name")
				CurrCell.className="ExcelDisplayCell"
				
				
			'	set CurrCell = oRow.insertCell()
			'	CurrCell.innerHtml = AttNode.getAttribute("Type")
			'	CurrCell.className="ExcelDisplayCell"
			'	
			'	
			'	set CurrCell = oRow.insertCell()
			'	CurrCell.innerHtml = AttNode.getAttribute("Length")
			'	CurrCell.className="ExcelDisplayCell"
			'	
			'	
			'	set CurrCell = oRow.insertCell()
			'	CurrCell.innerHtml = AttNode.getAttribute("Decimal")
			'	CurrCell.className="ExcelDisplayCell"
			'	
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hItemTypeZ"&iSno&"' value='"& AttNode.getAttribute("ItemType") &"'>")
				currCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hHeaderZ"&iSno&"' value='"& AttNode.getAttribute("Header") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hAttNameZ"&iSno&"' value='"& AttNode.getAttribute("Name") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hTypeZ"&iSno&"' value='"& trim(AttNode.getAttribute("Type")) &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hLengthZ"&iSno&"' value='"& AttNode.getAttribute("Length") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hDecimalZ"&iSno&"' value='"& AttNode.getAttribute("Decimal") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hAttIDZ"&iSno&"' value='"& AttNode.getAttribute("ID") &"'>")
				CurrCell.appendChild(oText)
				
		next 'for each AttNode in Root.childNodes
		document.formname.hRowCtr.value = iSno
	Else
		iSno = iSno + 1
		set oRow = document.all.tblDisplay.insertRow(document.all.tblDisplay.Rows.length)
		set CurrCell = oRow.insertCell()
		CurrCell.innerHtml = "No Records Found"
		CurrCell.ClassName= "ExcelDisplayCell"
		CurrCell.colspan= "7"
		CurrCell.align="Center"
	end if 'if Root.hasChildNodes() then
	document.formname.hRow.value = iSno
	
	iSno=0
	set Root = EditData.documentElement
	if Root.hasChildNodes() then
		for each AttNode in Root.childNodes
		iSno = iSno + 1
			set oRow = document.all.tblAttSelect.insertRow(document.all.tblAttSelect.Rows.length)
			
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = iSno
				CurrCell.className="ExcelHeaderCell"
				CurrCell.width = "10"
				CurrCell.align = "Center"
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=checkbox name='ChkSelect"&iSno&"' value ='"& iSno &"' class='FormElem'>")
				CurrCell.appendChild(oText) 
				CurrCell.align = "Center"
				CurrCell.className="ExcelDisplayCell"
				
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = AttNode.getAttribute("HName")
				CurrCell.className="ExcelDisplayCell"
				
				set CurrCell = oRow.insertCell()
				CurrCell.innerHtml = AttNode.getAttribute("Name")
				CurrCell.className="ExcelDisplayCell"
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hItemTypeSZ"&iSno&"' value='"& AttNode.getAttribute("ItemType") &"'>")
				currCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hHeaderSZ"&iSno&"' value='"& AttNode.getAttribute("Header") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hAttNameSZ"&iSno&"' value='"& AttNode.getAttribute("Name") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hTypeSZ"&iSno&"' value='"& trim(AttNode.getAttribute("Type")) &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hLengthSZ"&iSno&"' value='"& AttNode.getAttribute("Length") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hDecimalSZ"&iSno&"' value='"& AttNode.getAttribute("Decimal") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hAttIDSZ"&iSno&"' value='"& AttNode.getAttribute("ID") &"'>")
				CurrCell.appendChild(oText)
				
				set CurrCell = oRow.insertCell()
				set oText = document.createElement("<input type=hidden name='hAttValSZ"&iSno&"' value=''>")
				CurrCell.appendChild(oText)
				
		next 'for each AttNode in Root.childNodes
		document.formname.hRowCtr.value = iSno
	Else
		iSno = iSno + 1
		set oRow = document.all.tblAttSelect.insertRow(document.all.tblAttSelect.Rows.length)
		set CurrCell = oRow.insertCell()
		CurrCell.innerHtml = "No Records Found"
		CurrCell.ClassName= "ExcelDisplayCell"
		CurrCell.colspan= "6"
		CurrCell.align="Center"
	end if 'if Root.hasChildNodes() then
	
	
End Function
'===================================================
Function EditAttribute(sVal)
    sAttrID = eval("document.formname.hAttIDZ"&sVal).value
	document.formname.action = "ItmTypeAttributeEntry.asp?sValue="&sAttrID&"&Mod=U&ClassCode="&document.formname.hClassCode.value
	document.formname.submit 
End Function
'=========================================================
Function DeleteAttribute()
    dim iSelCnt,sVal
    'alert(sVal)
	'document.formname.selType.value = eval("document.formname.hItemTypeZ"&sVal).value
	For iCnt = 1 to cint(document.formname.hRowCtr.value)
	    if eval("document.formname.chk"&iCnt).checked = true then
	        iSelCnt = iSelCnt + 1
	        sVal = eval("document.formname.chk"&iCnt).value
	    end if
	Next
	
	if iSelCnt > 1 then
	    alert("Select any one attribute to Delete")
	    exit function
	end if
	sAttrID = eval("document.formname.hAttIDZ"&sVal).value
	document.formname.action = "ItmTypeAttributeDelete.asp?sValue="&sAttrID&"&ItemType="& eval("document.formname.hItemTypeZ"&sVal).value&"&ClassCode="&document.formname.hClassCode.value
	document.formname.submit 
End Function
'=========================================================
Function populateHeader(sItemType)
	Dim ndRoot,ndTypeHd,sTempIType,sTypeID
	set ndRoot = ItemTypeHeader.documentElement
	sItemType = document.formname.hItemType.value
	document.formname.selHeader.length = 1
	if ndRoot.hasChildNodes() then
		for each ndTypeHd in ndRoot.childNodes
			sTempIType = ndTypeHd.getAttribute("IType")
			sTypeID = ndTypeHd.getAttribute("ID")
			if trim(sItemType)="select" then
				document.formname.selHeader.length = document.formname.selHeader.length +1
				document.formname.selHeader(document.formname.selHeader.length-1).value = sTypeID
				document.formname.selHeader(document.formname.selHeader.length-1).text =  ndTypeHd.getAttribute("Name")
			else
				if trim(sTypeID)="2" then
					if trim(sItemType)="3" then
						document.formname.selHeader.length = document.formname.selHeader.length +1
						document.formname.selHeader(document.formname.selHeader.length-1).value = sTypeID
						document.formname.selHeader(document.formname.selHeader.length-1).text =  ndTypeHd.getAttribute("Name")	
					end if
				else
					document.formname.selHeader.length = document.formname.selHeader.length +1
					document.formname.selHeader(document.formname.selHeader.length-1).value = sTypeID
					document.formname.selHeader(document.formname.selHeader.length-1).text =  ndTypeHd.getAttribute("Name")	
				end if
			end if
		next
	end if
End Function
'============================================================