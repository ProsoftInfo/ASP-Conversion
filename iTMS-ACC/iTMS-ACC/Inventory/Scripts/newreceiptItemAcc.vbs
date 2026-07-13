dim OutDataValue, Root
Function GetLot(obj)
	sTempValues = obj.name

	OutDataValue = showModalDialog("newreceiptLotSerPop.asp?sTemp="&sTempValues,Data,"dialogHeight:320px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	
end Function

Function GetRate(obj)
Dim i
	sTempValues = obj.name
	'alert sTempValues
	
	i = 1
	Set OutDataValue = showModalDialog("newReceiptPackingDetailsPop.asp?sTemp="&sTempValues,Data,"dialogHeight:420px;dialogWidth:380px;center:Yes;help:No;resizable:No;status:No")
	Set Root = Data.documentElement
	
	'**************Following is added by Kumar K A to Display the itemrate*********************
	For Each HeaderNode In Root.childNodes
		For Each PageNode In HeaderNode.childNodes
			if StrComp(PageNode.nodeName,"STAGE") = 0 then
				eval("document.formname.txtRate"&i).Value = cDbl(trim(PageNode.Attributes.getNamedItem("IVALUE").Value)) / cDbl(trim(PageNode.Attributes.getNamedItem("IQTY").Value))								
				i = i + 1
			end if
		next
	next	
	'*******************************************************************************************
end Function

Function DisplayItem(obj)
	sTempValues = obj

	showModalDialog "itmDetailsPop.asp?sTemp="&sTempValues,document.formname.hOrgName.value,"dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
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

Function GetLott(obj,iQty)
	dim i
	dim arrTemp
	
	arrTemp = split(obj.name,":")
	
	iRec = arrTemp(1)
	sClass = arrTemp(2)
	sItem = arrTemp(3)
	sOrgID = arrTemp(4)

	sVar = "btn:"&iRec&":"&sClass&":"&sItem&":"&sOrgID
	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "button") then
			if (Instr(document.formname.elements(i).name,sVar) > 0) then
				'if document.formname.elements(i-4).value <> "N" then
					
					if document.formname.elements(i-2).selectedIndex = "0" then
						alert("Select Quantity in")
						document.formname.elements(i-2).focus
						exit function
					elseif (document.formname.elements(i-2).selectedIndex = "1" and document.formname.elements(i-1).selectedIndex = "0") then
						alert("Select Tare Weight")
						document.formname.elements(i-1).focus
						exit function
					elseif (document.formname.elements(i-1).selectedIndex = "1" and trim(document.formname.elements(i+2).value) = "") then
						alert("Enter Tare Weight")
						document.formname.elements(i+2).select
						exit function
					elseif (document.formname.elements(i-1).selectedIndex = "1" and trim(document.formname.elements(i+2).value) <> "" and not checkNumbers(trim(document.formname.elements(i+2).value))) then
						alert("Enter only Numerals")
						document.formname.elements(i+2).select
						exit function
					else
						iValue = document.formname.elements(i+2).value
					end if
				'end if
			end if
		end if
	next
	
	if iValue = "" then iValue = "N"
	sTempValues = obj.name&":"&iValue&":"&iQty

	Set OutDataValue = showModalDialog("newreceiptStorePop.asp?sTemp="&sTempValues,Data,"dialogHeight:320px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	
end Function

Function Init()
	spanDate.innerText = document.formname.hDate.value
end Function

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

Function CheckSubmit(todaysdate)
	dim i,iCtr,iCtrr
	dim arrTemp,sSrc
	iCtrr = document.formname.hiCtr.value
		
	if iCtrr = "" then exit function
	iCtr = 0

	sTempDate = document.formname.hDate.value
'	if(datediff("d",sTempDate,document.formname.ctlDDate.GetDate)) < 0 then
'		alert("Date should be greater than or equal to Receipt Date")
'		exit function
'	end if
'	if(datediff("d",todaysdate,document.formname.ctlDDate.GetDate)) > 0 then
'		alert("Date should be less than or equal to Today's Date")
'		exit function
'	end if
'	If(datediff("d",document.formname.hFinTo.value, document.formname.ctlDDate.GetDate)) > 0 Then
'		alert("Date Should be less than or equal to " & document.formname.hFinTo.value)
'		Exit Function
'	End If
'	
'	Set Root = Data.documentElement
'	Root.setAttribute "RECEIVEDON", sTempDate
'
'	For Each HeaderNode In Root.childNodes
'		iCtr = iCtr + 1
'		iMRSNo = HeaderNode.Attributes.Item(4).nodeValue 
'		sSrc = HeaderNode.Attributes.Item(11).nodeValue 
'		
'		set Q = eval("document.formname.txtRate"&iCtr)
'		'if (iMRSNo = "0" and trim(Q.value) = "") then
'		'	alert("Enter Item Rate")
'		'	Q.select
'		'	exit function
'		'end if
'
'		if cint(HeaderNode.Attributes.Item(10).nodeValue) = cint(iCtr) then
'			set Q = eval("document.formname.selStore"&iCtr)
'			HeaderNode.setAttribute "STORE", trim(Q.value)
'			
'			set Q = eval("document.formname.txtRate"&iCtr)
'			HeaderNode.setAttribute "IRATE", trim(Q.value)
'		end if
'	Next
'	
'	For Each HeaderNode In Root.childNodes
'		For Each PageNode In HeaderNode.childNodes
'			if StrComp(PageNode.nodeName,"STAGE") = 0 then
'				iQty = trim(PageNode.Attributes.getNamedItem("IVALUE").Value)				
'				
'				' This Line is added by Kumar K A to Get ItemRate
'				HeaderNode.setAttribute "IRATE", cDbl(trim(PageNode.Attributes.getNamedItem("IVALUE").Value)) / cDbl(trim(PageNode.Attributes.getNamedItem("IQTY").Value))
'				if iQty = "" then iQty = "0"
'				iValue = cdbl(iValue) + cdbl(iQty)
'			end if
'		next
'	next		
'	
'	if (iValue = 0) and sSrc ="F" then
'		alert("Enter Packing Details")
'		exit function
'	end if
	
'	msgbox Data.xml
'	exit function
'	Set objhttp = CreateObject("Microsoft.XMLHTTP")

'	objhttp.Open "POST","newreceiptInsert.asp", false
'	objhttp.send Data.XMLDocument
	
'	alert objhttp.responseText
	
'	exit function

'	if objhttp.responseText = "Y<BR>" then
'		alert ("Receipt has been done Successfully")
'		window.location.href = "MATERIALRECEIPTS.ASP"
'	elseif objhttp.responseText = "N<BR>" then
'		alert ("Receipt Number has been already defined")
'	else
'		alert objhttp.responseText
'	end if

	document.formname.action = "newInternalReceiptAcc.asp?RecNo="&document.formname.hRecNo.value
	document.formname.submit

end Function
