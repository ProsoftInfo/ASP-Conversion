Function Init(FDate,TDate)
	document.formname.ctlFromDate.SetDate=FDate
	document.formname.ctlToDate.SetDate=date()
	document.formname.ctlFromDate.SetMinDate=FDate
	document.formname.ctlToDate.SetMaxDate=TDate
end Function

'************Newly added by Maheswari on March 18th 2008 to check login financial year*****************
Function MinDate()
	Dim sMinDate,sFinPeriod,sSelDate,sMaxDate,sQtnDate	
	'alert("date check")
	'sFinPeriod = document.formname.hFinPeriod.value
	sMinDate = document.formname.hFinFromDate.value
	sMaxDate = document.formname.hFinToDate.value
	
	sFromDate = document.formname.ctlFromDate.getdate
	sToDate = document.formname.ctlToDate.getdate
	 
	If dateDiff("d",sFromDate,sMinDate) > 0 or  dateDiff("d",sFromDate,sMaxDate) < 0 then
		alert("Date Should be within the Financial Year  "& sMinDate&" to " & sMaxDate )
		document.formname.ctlToDate.SetDate =document.formname.hFinToDate.value
		exit Function
	end if
	If dateDiff("d",sToDate,sMinDate) > 0 or datediff("d",sToDate,sMaxDate) < 0  then
		alert("Date Should be within the Financial Year  "& sMinDate &" to " & sMaxDate )
		document.formname.ctlFromDate.SetDate = document.formname.hFinFromDate.value
		exit Function
	end if
End Function
'****************************************************************************************

Function clearIT()
	'for i=0 to document.formname.selIType.length - 1
	'	document.formname.selIType.options(i).selected = false
	
	'next
	
	
	document.formname.selStorage.options.length = 0	
	sOrgID = document.formname.hOrgID.value	
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	 
	objhttp.Open "GET","../transaction/XMLSelectStorage.asp?sOrgID="& sOrgID, false

	objhttp.send 
	'alert(objhttp.responseText)
	'alert(objhttp.responseXML.XML)
	if objhttp.responseXML.xml <> "" then
		storageData.loadXML objhttp.responseXML.xml

		Set Root = storageData.documentElement
		document.formname.selStorage.options.length = 0	
		For Each StoreNode In Root.childNodes
			document.formname.selStorage.length = document.formname.selStorage.length+1
			document.formname.selStorage.options(document.formname.selStorage.length-1).text = StoreNode.Attributes.Item(1).nodeValue
			document.formname.selStorage.options(document.formname.selStorage.length-1).Value = StoreNode.Attributes.Item(0).nodeValue
		next
		
	end if
	'ClearClass()
end Function

'Code to clear the classes
Function ClearClass()
	'document.formname.selClass.options.length = 0
	
	for i=0 to document.formname.selClass.length - 1
		document.formname.selClass.options(i).selected = false
	next
	document.formname.selItem.options.length = 0
End Function

'Code to get the added classes
Function AddClass()
	dim i, j, k, sOrgID, sIType, sITypeName, arrTemp, arrTempClass, sClass
	dim arrTempName, sClassName, arrDisplayValue, arrDisplayName

'	if document.formname.selUnit.selectedindex = -1 then
'		alert "Select Unit"
'		document.formname.selUnit.focus
'		exit function
'	else
	if document.formname.selIType.selectedindex = -1 then
		alert "Select Item Type"
		document.formname.selIType.focus
		exit function
	end if
	
'	for i=0 to document.formname.selUnit.length - 1
'		if document.formname.selUnit.options(i).selected then
'			sOrgID = sOrgID&":"&document.formname.selUnit.options(i).value
'		end if
'	next
'	sOrgID = trim(mid(sOrgID,2))
	sOrgID = document.formname.hOrgID.value

	for i=0 to document.formname.selIType.length - 1
		if document.formname.selIType.options(i).selected then
			sIType = sIType&":"&document.formname.selIType.options(i).value
			sITypeName = sITypeName&":"&document.formname.selIType.options(i).text
		end if
	next
	sIType = trim(mid(sIType,2))
	sITypeName = trim(mid(sITypeName,2))

	OutValue = showModalDialog("../../include/ClassificationSelectPop.asp?sIType=" & sIType & "&sOrgID=" & sOrgID & "&sITypename="& sITypeName,"Classification","dialogHeight:460px;dialogWidth:625px;center:Yes;help:No;resizable:No;status:No")

	arrTemp = split(OutValue,"*****")

	if arrTemp(0) = "-1" then exit function

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","../Master/XMLSelectItemClass.asp?sOrgID=" & sOrgID &"&sText="&arrTemp(0), false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				if not CheckExists(HeaderNode.Attributes.getNamedItem("CLASSCODE").Value) then
					document.formname.selClass.length = document.formname.selClass.length+1
					document.formname.selClass.options(document.formname.selClass.length-1).text = HeaderNode.Attributes.getNamedItem("CLASSNAME").Value
					document.formname.selClass.options(document.formname.selClass.length-1).Value = HeaderNode.Attributes.getNamedItem("CLASSCODE").Value
				end if
			next
		end if
	end if

End Function

'Code to check whether the class already exists or not
Function CheckExists(iClass)
	Dim i, l
	l = document.formname.selClass.length - 1
	For i = 0 to l
		if document.formname.selClass.options(i).value = iClass then
			CheckExists = True
			exit function
		end if
	next

	CheckExits = False
End Function

Function popClass(obj,sWho)

'	if document.formname.selUnit.value = "" then 
'		alert("Select Unit")
'		document.formname.selUnit.focus
'		exit function
'	end if

	'document.formname.selClass.options.length = 0
	document.formname.selItem.options.length = 0

	dim Root,HeaderNode

'	for i=0 to document.formname.selUnit.length - 1
'		if document.formname.selUnit.options(i).selected then
'			sOrgID = sOrgID&":"&"'"&document.formname.selUnit.options(i).value&"'"
'		end if
'	next
'	sOrgID = trim(mid(sOrgID,2))
	sOrgID = document.formname.hOrgID.value

	for i=0 to document.formname.selIType.length - 1
		if document.formname.selIType.options(i).selected then
			sIType = sIType&":"&"'"&document.formname.selIType.options(i).value&"'"
		end if
	next
	sIType = trim(mid(sIType,2))

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","../transaction/XMLSelectClass.asp?Type=ST&sIType=" & sIType & "&sOrgID="& sOrgID, false

	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				document.formname.selClass.length = document.formname.selClass.length+1
				document.formname.selClass.options(document.formname.selClass.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selClass.options(document.formname.selClass.length-1).Value = cstr(HeaderNode.Attributes.Item(0).nodeValue)
			next
		end if
	end if
end Function

Function popItem(obj)

	if document.formname.selClass.value = "" then 
		alert("Select Classification")
'		document.formname.selClass.focus
		exit function
	end if

	hFromDate = document.formname.ctlFromDate.GetDate
	hToDate = document.formname.ctlToDate.GetDate

	document.formname.selItem.options.length = 0

	dim Root,HeaderNode

'	for i=0 to document.formname.selUnit.length - 1
'		if document.formname.selUnit.options(i).selected then
'			sOrgID = sOrgID&":"&"'"&document.formname.selUnit.options(i).value&"'"
'		end if
'	next
'	sOrgID = trim(mid(sOrgID,2))
	sOrgID = document.formname.hOrgID.value

	i = 0
	if document.formname.selClass.value <> "" then
		for i=0 to document.formname.selClass.length - 1
			if document.formname.selClass.options(i).selected then
				selClass = selClass&"?"&document.formname.selClass.options(i).value	
			end if
		next
	else
		for i=0 to document.formname.selClass.length - 1
			selClass = selClass&"?"&document.formname.selClass.options(i).value	
		next
	end if
	selClass = trim(mid(selClass,2))

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","../transaction/XMLSelectClass.asp?Type=STI&sClass=" & selClass & "&sOrgID="& sOrgID & "&sFDt="& hFromDate & "&sFTt="& hToDate, false

	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		ItemData.loadXML objhttp.responseXML.xml
		Set Root = ItemData.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				document.formname.selItem.length = document.formname.selItem.length+1
				document.formname.selItem.options(document.formname.selItem.length-1).text = HeaderNode.Attributes.Item(3).nodeValue
				document.formname.selItem.options(document.formname.selItem.length-1).Value = cstr(HeaderNode.Attributes.Item(2).nodeValue)
			next
		end if
	end if
end Function

Function GetData()

	dim Root,HeaderNode

	Set Root = Data.documentElement

	sIType = document.formname.selIType.value
	
	
	sExp ="//CODE [ @ITEMTYPEID = '"&sIType&"']"
	Set CodeNode = Root.Selectnodes(sExp)
	
	i= 0
	k = 0
	sExp ="//ATTRIBUTES [ @ITEMTYPEID = '"&sIType&"']"
	Set CodeNode = Root.Selectnodes(sExp)
	
	AddAttrib(CodeNode.Length)
	'alert(CodeNode.Length)
	for i = 1 to CodeNode.Length 
		iCode = CodeNode.item(i-1).attributes.getNamedItem("ATTRID").value		 
		set selObj = eval("document.formname.selAttrZ"&i)		
		selObj.length = 0
		 
	next
	i=0
	for i = 1 to CodeNode.Length 
	
		iCode = CodeNode.item(i-1).attributes.getNamedItem("ATTRID").value
		
		set selObj = eval("document.formname.selAttrZ"&i)
		selObj.length = selObj.length+1
		selObj.options(selObj.length-1).text = CodeNode.item(i-1).attributes.getNamedItem("ATTRNAME").value
		selObj.options(selObj.length-1).Value = CodeNode.item(i-1).attributes.getNamedItem("ATTRID").value
		
		sExp1 ="//ATTRIBUTES [ @ITEMTYPEID = '"&sIType&"']/GROUP[ @ATTRID = "&iCode&"]"
		Set GNode = Root.Selectnodes(sExp1)
		for k = 0 to GNode.Length - 1
			selObj.length = selObj.length+1
			selObj.options(selObj.length-1).text = GNode.item(k).attributes.getNamedItem("OPTIONNAME").value
			selObj.options(selObj.length-1).Value = GNode.item(k).attributes.getNamedItem("OPTIONVALUE").value
		next
	next
end Function

Function DisplayBy()
	IF trim(document.formname.selView.value) = "Rec" or trim(document.formname.selView.value) = "Iss" then
		
		idDispBy.innerhtml = "<input type=""Radio""  Class=""ExcelDisplayCell"" name=""radOpt"" value=""I"" Checked>Itemwise"
		idDispBy.innerhtml = idDispBy.innerhtml + "&nbsp;<input type=""Radio""  Class=""ExcelDisplayCell"" name=""radOpt"" value=""D"" >Datewise"
	End IF
	IF trim(document.formname.selView.value) = "Iss" then	
		idDispBy.innerhtml = idDispBy.innerhtml + "&nbsp;<input type=""Radio"" Class=""ExcelDisplayCell"" name=""radOpt"" value=""C"" >CostCenterwise"		
		document.formname.txtIssueNo.readonly = False
	Else
		document.formname.txtIssueNo.readonly = True
	End IF
	IF trim(document.formname.selView.value) <> "Rec" And trim(document.formname.selView.value) <> "Iss" then   
		idDispBy.innerhtml = ""
	End IF
	
End Function

Function CheckSubmit(todaysdate)

	if DateDiff("d",document.formname.ctlFromDate.GetDate,todaysdate) < 0 then
		alert("From Date should be less than or equal to Today's Date")
		Exit Function
	elseif DateDiff("d",document.formname.ctlToDate.GetDate,todaysdate) < 0 then
		alert("To Date should be less than or equal to Today's Date")
		Exit Function
	elseif DateDiff("d",document.formname.ctlFromDate.GetDate,document.formname.ctlToDate.GetDate) < 0 then
		alert("To Date should be greater than or equal to From Date")
		Exit Function
'	elseif document.formname.selUnit.value = "" then
'		alert("Select Unit")
'		document.formname.selUnit.focus
'		exit function
	elseif document.formname.selIType.value = "" then
		alert("Select Item Type")
		document.formname.selIType.focus
		exit function
	else
		
		Set Root = Data.documentElement
		
		hFromDate = document.formname.ctlFromDate.GetDate
		hToDate = document.formname.ctlToDate.GetDate
		hFinFromDate = document.formname.hFinFromDate.value
		
	'	for i=0 to document.formname.selUnit.length - 1
	'		if document.formname.selUnit.options(i).selected then
	'			selUnit = selUnit&"?"&document.formname.selUnit.options(i).value
	'			hUnitName = hUnitName&"?"&document.formname.selUnit.options(i).text
	'		end if
	'	next
	'	selUnit = trim(mid(selUnit,2))
	'	hUnitName = trim(mid(hUnitName,2))
		selUnit = trim(document.formname.hOrgID.value)
		hUnitName = trim(document.formname.hOrgName.value)

		selIType = document.formname.selIType.value
		hIType = document.formname.selIType(document.formname.selIType.selectedIndex).text

		i=0
		sExp ="//ATTRIBUTES [ @ITEMTYPEID = '"&selIType&"']"
		Set CodeNode = Root.Selectnodes(sExp)

		for i = 1 to CodeNode.Length 
			iCode = CodeNode.item(i-1).attributes.getNamedItem("ATTRID").value
			set selObj = eval("document.formname.selAttrZ"&i)
			if selObj.selectedIndex > 0 then
				selAttr = selAttr&"?"&selObj.value
			end if
		next
		selAttr = trim(mid(selAttr,2))
		'alert(selAttr)
		sCatVal = ""
		sItemCode = "0"
		IF trim(document.formname.selCategory.value) <> "select" then
			sCatCode = document.formname.selCategory.value					
		End IF
		'alert(document.formname.hItemCode.value)
		IF(document.formname.hItemCode.value) <> "" then
			sItemCode = document.formname.hItemCode.value
		End IF
		'	 alert(sCatCode & "  " & sItemCode)
		'alert(document.formname.selStorage.value)
		if document.formname.selStorage.value <> "" then
		
			for i=0 to document.formname.selStorage.length - 1
				if document.formname.selStorage.options(i).selected then
					sStorVal = sStorVal&"?"&document.formname.selStorage.options(i).value	
				end if
			next
			sStorVal = trim(mid(sStorVal,2))
		else
			for i=0 to document.formname.selStorage.length - 1
				sStorVal = sStorVal&"?"&document.formname.selStorage.options(i).value	
			next
			sStorVal = trim(mid(sStorVal,2))
		end if
		'alert(sStorVal)
		If Trim(document.formname.txtIssueNo.Value) <> "" Then
			sIssueString = Trim(document.formname.txtIssueNo.Value)
		End If
		
		IF trim(document.formname.selView.value) = "Sto" then
			sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selClass&":"&selAttr
		ElseIF trim(document.formname.selView.value) = "Led" then
			sTempValues = selUnit&":"&hUnitName&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selClass&":"&selAttr
		ElseIF trim(document.formname.selView.value) = "Rec" then
		
			IF trim(document.formname.radOpt(0).checked) = "True" then
				sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&selIType&":"&hIType&":"&selClass&":I:"&sItemCode
			ElseIF trim(document.formname.radOpt(1).checked) = "True" then
				sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&selIType&":"&hIType&":"&selClass&":D:"&sItemCode
			End IF	
			
		ElseIF trim(document.formname.selView.value) = "Iss" then	
			IF trim(document.formname.radOpt(0).checked) = "True" then
				sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&selIType&":"&hIType&":"&selClass&":I:"&sItemCode
				
			ElseIF trim(document.formname.radOpt(1).checked) = "True" then
				sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&selIType&":"&hIType&":"&selClass&":D:"&sItemCode
			ElseIF trim(document.formname.radOpt(2).checked) = "True" then
				sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&selIType&":"&hIType&":"&selClass&":C:"&sItemCode
			End IF	
		ElseIF trim(document.formname.selView.value) = "Mat" then
			sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&selClass&":"&sItemCode&":"&selIType
		Elseif trim(document.formname.selView.value) = "StoMIS" then
			sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selClass
		Elseif trim(document.formname.selView.value)="RcptLot" then
			sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selAttr
		Elseif trim(document.formname.selView.value)="IssueLot" then
			sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selAttr
		Elseif trim(document.formname.selView.value)="IssCon" then
			sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selAttr
		Elseif trim(document.formname.selView.value)="IssDet" then
			sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selAttr
	    Elseif trim(document.formname.selView.value)="SRS" then
	        sTempValues = selUnit&":"&cstr(hFromDate)&":"&cstr(hToDate)&":"&(hFinFromDate)&":"&selIType&":"&hIType&":"&selAttr
	    Elseif Trim(document.formname.selView.value)="AWSR" then
	        sTempValues = selUnit &":"& CStr(hFromDate)&":"&cstr(hToDate)&":"&selIType&":"&hIType&":"& selAttr
		End IF
		
		if selAttr <> "" then
			sTempValues = sTempValues & ":" & "SELECTED"
		else
			sTempValues = sTempValues & ":" & "ALL"
		end if
			
		sTempValues = sTempValues &":"& sCatCode &":"& sItemCode &":"& sStorVal
		IF trim(document.formname.selView.value) = "Iss" then
			If sIssueString <> "" Then
				sTempValues = sTempValues & ":" & "ISSUESTRING:" & sIssueString
			Else
				sTempValues = sTempValues & ":" & "ALL:" & sIssueString
			End IF
		End IF		
		'alert(sTempValues)
		
		'Exit function
		
		IF trim(document.formname.selView.value) = "Sto" then
			win = open ( "ItemAttrStockDetailsEntry.asp?sTemp="&sTempValues, "StockStatement", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		ElseIF trim(document.formname.selView.value) = "Led" then
			win = open ( "StoreLedgerDetailsEntry.asp?sTemp="&sTempValues, "StockStatement", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		ElseIF trim(document.formname.selView.value) = "Rec" then			
			win = open ( "ReceiptItemDetailsEntry.asp?sTemp="&sTempValues&"&Attribute="&selAttr, "ReceiptItem", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )		
		ElseIF trim(document.formname.selView.value) = "Iss" then				
			win = open ( "IssueItemDetailsEntry.asp?sTemp="&sTempValues&"&Attribute="&selAttr, "IssuedItems", "height=540,width=795, toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		ElseIF trim(document.formname.selView.value) = "Mat" then
			win = open ("MaterialConsumptionDetailsEntry.asp?sTemp="&sTempValues,"", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		ElseIF trim(document.formname.selView.value) = "StoMIS" then
			win = open ( "StockStmtDetailsEntry_MIS.asp?sTemp="&sTempValues,"","height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		ElseIF trim(document.formname.selView.value) = "RcptLot" then
			win = open ( "ReceiptLotDetails.asp?sTemp="&sTempValues,"","height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		Elseif trim(document.formname.selView.value)="IssueLot" then
			win = open ( "IssueLotDetails.asp?sTemp="&sTempValues,"","height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		Elseif trim(document.formname.selView.value)="IssCon" then
			win = open ( "StockLotConsolidated.asp?sTemp="&sTempValues,"","height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		Elseif trim(document.formname.selView.value)="IssDet" then
			win = open ( "StockLotDetailed.asp?sTemp="&sTempValues,"","height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
	    Elseif trim(document.formname.selView.value)="SRS" then
	        win = open ( "StockReplishmentStmt.asp?sTemp="&sTempValues,"","height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
	        'win = open ( "StockReplishmentStmt.asp?sTemp="&sTempValues, "StockReplishment", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )			
	    Elseif Trim(document.formname.selView.value)="AWSR" then
	        win = open ("InvAgeingReport.asp?sTemp="&sTempValues,"","height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0" )
		End If
	end if
end Function
Function popCC(obj)
	
	document.formname.selCC.options.length = 0

	dim Root,HeaderNode
	sOrgID = document.formname.hOrgID.value
	
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	sIType=trim(sIType)
	objhttp.Open "GET","../transaction/XMLSelectCostCenter.asp?sOrgID="& sOrgID, false
	
	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				document.formname.selCC.length = document.formname.selCC.length+1
				document.formname.selCC.options(document.formname.selCC.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selCC.options(document.formname.selCC.length-1).Value = cstr(HeaderNode.Attributes.Item(0).nodeValue)
			next
		end if
	else
		alert("No Cost Center for the Unit Selected")
		'document.formname.selUnit.focus()
		Exit Function
	end if
End Function

Function ClearAll()
	'document.formname.selClass.options.length = 0	
	'document.formname.selItem.options.length = 0
	document.formname.reset()
end Function

'Newly Added On 8th Oct 2007 by Maheshwari 
Function Search()

'	if document.formname.selUnit.value = "" then
'		alert("Select Unit")
'		document.formname.selUnit.focus
'		exit function		
'	else
	if document.formname.selIType.value = "" then
		alert("Select Item Type")
		document.formname.selIType.focus
		exit function
	end if
		
	sUnit = document.formname.hOrgID.value
	sIType = document.formname.selIType.value
	iStock = "N"
	set Root = ItemData.documentelement 
	'alert(sIType)
	'set ResData=showModalDialog("../transaction/ItemSelect.asp?orgID=" & sUnit & "&sIType=" & sIType & "&hSelectMode=M&Flag="+cstr(nFlag),ItemData,"dialogHeight:650px;dialogWidth:730px;center:Yes;help:No;resizable:No;status:No")
	set ResData=showModalDialog("../../Common/ItemSelectCommon.asp?orgID=" & sUnit & "&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag),ItemData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
	'alert(ResData.xml)
	sAct = UCase(trim(ResData.getAttribute("Action")))
	sQuery = trim(ResData.getAttribute("PassQuery"))
	document.formname.hItemType.value = trim(ResData.getAttribute("ItemType"))
 	if ucase(trim(sAct)) <> "CLOSE" then
		do while sAct <> "DONE"		
			'set OutValue=showModalDialog("../transaction/ItemSelect.asp?" & sQuery,ItemData,"dialogHeight:650px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
			set OutValue=showModalDialog("../../Common/ItemSelectCommon.asp?" & sQuery,ItemData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
			
			sAct = UCase(trim(ResData.getAttribute("Action")))
			if ucase(Trim(sAct)) = "CLOSE" then exit do
			sQuery = trim(ResData.getAttribute("PassQuery"))
		loop
	end if 'if ucase(trim(sAct)) <> "CLOSE" then
				
	If not Root.hasChildNodes Then 	exit function	
	sTemp = ""
	'alert(Root.xml)
	if Root.hasChildNodes() then
		
		For Each HeaderNode In Root.childNodes
		
			'sTemp = HeaderNode.Attributes.Item(4).nodeValue & ":" & HeaderNode.Attributes.Item(3).nodeValue
			sTemp = sTemp &"|"& HeaderNode.Attributes.Item(2).nodeValue
			IF Left(sTemp,1) = "|" then
				sTemp = Mid(sTemp,2)
			End IF	
		
			'alert(sTemp)
			sVal = split(HeaderNode.Attributes.Item(4).nodeValue,"--")
			sText = sText &","& sVal(0)
			idItemName.innerText = Mid(sText,2)
								
		next
		document.formname.hItemCode.value = sTemp
	else
		alert("No Items found")
		exit function
	end if

end Function
