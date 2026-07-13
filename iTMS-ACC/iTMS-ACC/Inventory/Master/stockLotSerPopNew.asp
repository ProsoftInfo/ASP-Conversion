<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	stockLotSerPopNew.asp
	'Module Name				:	Inventory (Opening Stock Lot and Serial Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 28, 2003
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	May 24,2010 'Update the Auto Generation Entry Selection as a default
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/NoSeries.asp" -->
<%
	Dim oDom,Root,HeaderNode,newElem

	'Declaration of Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	' Declaration of variables
	Dim dcrs,rsTemp
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set rsTemp = Server.CreateObject("ADODB.RecordSet")

	dim iItem,iClass,sOrgID,sType,arrTemp,sStoName,iQty,sItmType,sFormCode
	dim bPackNoSerFlag,sComCode,sStoresUom,sTemp
	dim arrUoM,sUoMDesc,sUoMCode,sCheck,sPackNoFlag,bPackFlag
	Dim ct, StDate,sItemName,sAttributeID,sQuery
	ct = 0
	bPackNoSerFlag = false
	bPackFlag = false

	sTemp = trim(Request.QueryString("sTemp"))
	 'Response.write "sTemp="&sTemp
sTemp = replace(sTemp,"'","~")
sTemp = Replace(sTemp,Chr(34),"~~")
	arrTemp = split(sTemp ,"``")
	sType	= arrTemp(0)
	iItem	= arrTemp(1)
	iClass	= arrTemp(2)
	sOrgID	= arrTemp(3)
	iQty = arrTemp(6)
	sStoName = arrTemp(7)
	sStoresUom = arrTemp(8)
	sItmType =  arrTemp(11)
	sAttributeID=arrTemp(12)
'	Response.Write "sAttributeID = "& sAttributeID
	'Response.Write iItem
	if iItem = "0" then
		sItemName = replace(arrTemp(10),"~~","'")
	else
		sItemName = replace(ItemDisplay(iItem,"NULL"),"~~","'")
	end if
	If UBound(arrTemp) = 10 Then
		StDate = arrTemp(10)
	Else
		StDate  = arrTemp(11)
	End If
	IF  sItmType = "" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMTYPEID FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & iClass & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			sItmType = trim(dcrs(0))
		end if
		dcrs.Close
	End IF
	
	''blocked by ragav on Feb 27, 2012
	''begin
'	with dcrs
'		.CursorLocation = 3
'		.CursorType = 3
'		'.Source = "SELECT COMPANYITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ""
'		.Source = "SELECT ITEMTYPEID FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ""
'		.ActiveConnection = con
'		.Open
'	end with
'	set dcrs.ActiveConnection = nothing
'	if not dcrs.EOF then
'		'sComCode = trim(dcrs(0))
'		sItmType = dcrs(0)
'		'sFormCode = mid(sComCode,8,1)
''	end if
'	dcrs.Close
''end

	'IF sFormCode <> "" then
	'	if sItmType <> "YRN" then
	'		sFormCode = "NULL"
	'	else
	'		sFormCode = Pack(sFormCode)
	'	end if
	'End IF
	if sItmType <> "YRN" then
		sFormCode = "NULL"
	else
		sFormCode = Pack(sFormCode)
	end if

	if sItmType = "YRN" then
		' Packing Number - Number Series
		' OrganisationCode,SeriesNumber,SeriesCode,Date
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'PN' AND ITEMTYPE = " & Pack(sItmType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
			.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'PN' AND ITEMTYPE = 'YRN' AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			bPackNoSerFlag = true
		end if
		dcrs.close
	end if

	Set Root = oDOM.createElement("Packing")
	oDOM.appendChild Root
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT PACKINGCODE,PACKINGSHORTNAME,PACKINGNAME FROM APP_M_PACKINGTYPE WHERE PACKINGCODE IN (SELECT PACKINGCODE FROM SAL_R_ITEMTYPEPACK WHERE ITEMTYPEID = " & Pack(sItmType) & ")"
		.Source = "SELECT PACKINGCODE,PACKINGSHORTNAME,PACKINGNAME FROM APP_M_PACKINGTYPE" 'WHERE PACKINGCODE IN (SELECT PACKINGCODE FROM SAL_R_ITEMTYPEPACK WHERE ITEMTYPEID = " & Pack(sItmType) & ")"
		.ActiveConnection = con
		.Open
	end with
	'Response.Write dcrs.source
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		Do While Not dcrs.EOF
			Set newElem = oDOM.createElement("PACK")
			newElem.setAttribute "PACKCODE",trim(dcrs(0))
			newElem.setAttribute "PACKSHNAME",trim(dcrs(1))
			newElem.setAttribute "PACKNAME",trim(dcrs(2))
			Root.appendChild newElem
		dcrs.MoveNext
		Loop
	else
		bPackFlag = true
	end if
	dcrs.Close

	oDOM.Save server.MapPath("../temp/master/PACKING"&Session.SessionID&".xml")

	set oDOM = nothing

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	Set Root = oDOM.createElement("Selling")
	oDOM.appendChild Root


	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS WHERE CODE = " & sFormCode & " AND ITEMTYPEID = " & Pack(sItmType) & ""
		.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		Do While Not dcrs.EOF
			Set newElem = oDOM.createElement("SELL")
			newElem.setAttribute "UNITID",trim(dcrs(0))
			newElem.setAttribute "UNITNAME",trim(dcrs(1))
			Root.appendChild newElem
		dcrs.MoveNext
		Loop
	else
		bPackFlag = true
	end if
	dcrs.Close

	oDOM.Save server.MapPath("../temp/master/SELLING"&Session.SessionID&".xml")

	set oDOM = nothing

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	Set Root = oDOM.createElement("SellingForm")
	oDOM.appendChild Root

	GetSellingForm

	oDOM.Save server.MapPath("../temp/master/SELLINGFORM"&Session.SessionID&".xml")

	arrUoM = split(DisplayUoM(iItem),":")
	sUoMCode = arrUoM(0)
	sUoMDesc = arrUoM(1)
	sCheck = UoMDecimal(sUoMCode)

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT NUMBERINGTYPE FROM PRD_M_PACKINGNUMBERSERIESCHECK WHERE ORGANISATIONCODE = " & Pack(sOrgID)
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sPackNoFlag  = trim(dcrs(0))
	else
		sPackNoFlag = ""
	end if
	dcrs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock - Lot / Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="Data" data-itms-xml-island>
<root/>
</script>
<script type="application/xml" id="SerialData" data-itms-xml-island>
<ROOT/>
</script>
<script type="application/xml" id="FabricData" data-itms-xml-island>
<ROOT/>
</script>

<script type="application/xml" id="PackingData" data-itms-xml-island data-src="<%="../temp/master/PACKING"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" id="SellingData" data-itms-xml-island data-src="<%="../temp/master/SELLING"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" id="SellingFormData" data-itms-xml-island data-src="<%="../temp/master/SELLINGFORM"&Session.SessionID&".xml"%>"></script>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<!--SCRIPT LANGUAGE=vbscript SRC="../scripts/openingLot.vbs"></SCRIPT-->
<SCRIPT LANGUAGE=vbscript>
dim objTemp,Root,newElem,SerialFlag,Cols
dim iClass, iItem, sOrgID, sType, j, iStore, iBin, iValue, iTotValue
dim iQtyTotGross, iQtyTotTare, No, iFixedNo
dim iQtyTotGrossSerial, iQtyTotTareSerial
dim sLotNo, iCounter, objTempFabric, iCnt, avail,sOldLotNo
dim LotSerialNode, LotSerialDetNode, STDate, sPubLot
dim AutoLot, AutoSerial, NewVal
AutoLot = 0
AutoSerial = 0
NewVal = 0
iCnt = 0

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

Function SerialCheck(iFromValue,iToValue)
	dim Root,HeaderNode,objhttp

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","../../Inventory/transaction/XMLSerialCheck.asp?iFrom="&iFromValue&"&iTo="&iToValue, false

	objhttp.send

	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement
		if Root.Attributes.Item(0).nodeValue = "Y" then
			SerialCheck = false
		else
			SerialCheck = true
		end if
	end if

end Function

Function LotCheck(iValue)
	if iValue = "N/A" then
		LotCheck = false
		exit function
	end if
	dim Root,HeaderNode,objhttp

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","LotXMLSelect.asp?orgID="&sOrgID&"&iLot="&iValue, false

	objhttp.send

	'alert(objhttp.responseText)

	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement
		if Root.Attributes.Item(0).nodeValue = "N" then
			LotCheck = false
		else
			LotCheck = true
		end if
	end if
end Function

Function LotSerialLocalCheck(iValue,iFromValue,iToValue)
	Set Root = objTemp.documentElement
	sExp ="//STORAGE/LotSerial [ @LOT = '"&iValue&"']/LotSerialDetails"
	Set LotNode = Root.Selectnodes(sExp)

	if LotNode.Length = 0 or iValue= "N/A" then
		LotSerialLocalCheck = true
		exit function
	else
		LotSerialLocalCheck = false
		exit function
	end if

	For iCounter = 0 to LotNode.Length - 1
		arrTemp = split(LotNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value," - ")
		sLotserialNo = arrTemp(0)
		for iTemp = cint(iFromValue) to cint(iToValue)
			if cint(sLotserialNo) = cint(iTemp) then
				LotSerialLocalCheck = false
				exit function
			else
				LotSerialLocalCheck = true
			end if
		next
	next
end Function
'****************************************
Function init()
	Dim iCnt
	if strcomp(trim(document.formname.hRec.value),"S")<>0 then
		for iCnt = 0 to cint(document.formname.sellot.length)
			if strcomp(trim(document.formname.sellot(iCnt).value),"N")=0 then
				document.formname.sellot.selectedIndex = iCnt
				exit for
			end if
		next
	end if
	DispLot()
End Function
'****************************************
Function LoadDetails(obj)
	dim sTempLotNo, iLotCtr
	dim cntr,NoofPacks, iSubVal,iTemp,Qty
	arrTemp = split(obj,"``")
	sType = arrTemp(0)
	iItem = arrTemp(1)
	iClass = arrTemp(2)
	sOrgID = arrTemp(3)
	iStore = arrTemp(4)
	iBin = arrTemp(5)
	iValue = arrTemp(9)
	'alert sType
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
	'alert Root.XML
	iQtyTotGross = 0
	iQtyTotTare = 0
	iTotValue = 0
	iCnt = 0
	iFixedNo = 0
'	if ucase(document.formname.hItemType.value)<>"FAB" then
		Cols="9"
		idQtyEntered.innerText = FormatNumber(DisplayQty(),3,,,0)
'	else
'		Cols="10"
'		idQtyEntered.innerText = FormatNumber(DisplayFabricQty(),3,,,0)
'	end if

	EnableDone()

	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial"
	Set StorageNode = Root.Selectnodes(sExp)
	'alert(Root.xml)
	For iCounter = 0 to StorageNode.Length - 1
		setIndex document.formname.selAltUom,StorageNode.Item(iCounter).Attributes.getNamedItem("ALTUOM").Value
		document.formname.txtAltGross.value= StorageNode.Item(iCounter).Attributes.getNamedItem("ALTGROSS").Value
		document.formname.txtAltNett.value= StorageNode.Item(iCounter).Attributes.getNamedItem("ALTNETT").Value
		iTotValue = StorageNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value
		iSubRate = cDbl(StorageNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value) /  cDbl(StorageNode.Item(iCounter).Attributes.getNamedItem("QTY").Value)
		Qty = cDbl(StorageNode.Item(iCounter).Attributes.getNamedItem("QTY").Value)
		idAltUom.InnerText = document.formname.selAltUom.options(document.formname.selAltUom.selectedIndex).text
		if lcase(idAltUom.InnerText) = "select" then idAltUom.InnerText = ""
	next
	j=1


	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial"
	Set LotSerialNode = Root.Selectnodes(sExp)
	For iLotCtr = 0 to LotSerialNode.Length - 1
		 cntr = LotSerialNode.Item(iLotCtr).Attributes.getNamedItem("COUNTER").Value
		 sTempLotNo = ""
	'alert cntr
	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial [ @COUNTER ='"&cntr&"']/LotSerialDetails"
	Set StorageNode = Root.Selectnodes(sExp)


'	alert StorageNode.Length
	For iCounter = 0 to StorageNode.Length - 1

		sLotNo=StorageNode.Item(iCounter).Attributes.getNamedItem("LOT").Value
		sLotserialNo=StorageNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value
		sQtyRec=StorageNode.Item(iCounter).Attributes.getNamedItem("QTYREC").Value
		sTar=StorageNode.Item(iCounter).Attributes.getNamedItem("TAREREC").Value
		'alert sLotNo
		sSell=StorageNode.Item(iCounter).Attributes.getNamedItem("SELLINGTYPE").Value
		iWeight=StorageNode.Item(iCounter).Attributes.getNamedItem("WEIGHTSTYPE").Value
		sPack=StorageNode.Item(iCounter).Attributes.getNamedItem("PACKINGTYPE").Value
		sSellingForm=StorageNode.Item(iCounter).Attributes.getNamedItem("SELLINGFORM").Value
		iPackNo=StorageNode.Item(iCounter).Attributes.getNamedItem("PACKNUMBER").Value

		iSubVal = FormatNumber((CDbl(sQtyRec) - CDbl(sTar))* CDbl(iSubRate),2,,,0)

		if document.formname.hrec.value="S" then iTempValue=StorageNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value
		if document.formname.hrec.value="LS" then iTempValue=StorageNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value
		if document.formname.hrec.value="L" then iTempValue=StorageNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value


		sExp1 ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial[@LOT = '"&sLotNo&"' and  @COUNTER ='"&cntr&"']"
		Set LotNode = Root.Selectnodes(sExp1)
		iTemp = LotNode.Item(0).Attributes.getNamedItem("IVALUE").Value
		'alert LotNode.Item(0).Attributes.getNamedItem("IVALUE").Value

		If (sType = "S" or sType="LS") and iCnt = 0 Then
			'alert iTemp
				sTempLotNo = sLotNo
				iFixedNo = iFixedNo + 1

				set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
				set headerCell=oRow.insertCell()
				If sType <> "S" Then
				headerCell.innerHTML=sLotNo & " / "
				headerCell.className="ExcelDisplayCell"
				End If
				'if document.formname.hrec.value <> "S" then
				'ALERT("iTotValue")
					set oText = document.createElement("<input type=""text"" name=""txtIValueZ"&CStr(iFixedNo)&""" size=""12"" value="""&iTemp&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" >")
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
					headerCell.width = "8"
					headerCell.appendChild(oText)

					set oText = document.createElement("<input type=""text"" name=""txtValZ"&CStr(iFixedNo)&""" size=""12"" value="""&" ["&cDbl(iTotValue)\cDbl(Qty)&"]"&""" onBlur=""RecalculateTotal()"" READONLY class=""Formelem"" style=""text-align=Left"" >")
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
					headerCell.width = "8"
					headerCell.appendChild(oText)
				'end if
				headerCell.align="left"
				headerCell.colspan=Cols
		End If


		if sTempLotNo <> sLotNo and sLotNo <> "" Then
			sTempLotNo = sLotNo
			iFixedNo = iFixedNo + 1
			set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
			set headerCell=oRow.insertCell()
			If sType <> "S" Then
				headerCell.innerHTML=sLotNo & " / "
				headerCell.className="ExcelDisplayCell"
			End If
			'if document.formname.hrec.value <> "S" then
				set oText = document.createElement("<input type=""text"" name=""txtIValueZ"&CStr(iFixedNo)&""" size=""12"" value="""&iTemp&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"
				headerCell.width = "8"
				headerCell.appendChild(oText)

				set oText = document.createElement("<input type=""text"" name=""txtValZ"&CStr(iFixedNo)&""" size=""10"" value="""&" ["&cDbl(iTotValue)\cDbl(Qty)&"]"&""" onBlur=""RecalculateTotal()"" READONLY class=""Formelem"" style=""text-align=Left"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"
				headerCell.width = "8"
				headerCell.appendChild(oText)

			'end if
			headerCell.align="left"
			headerCell.colspan=Cols
		end if

		set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)

		set headerCell=oRow.insertCell()
		headerCell.innerHTML=j
		headerCell.className="ExcelDisplayCell"
		headerCell.align="center"
		iCnt = iCnt + 1

'		set headerCell=oRow.insertCell()
'		set oText = document.createElement("<input type=""text"" name=""txtSerialZ"&CStr(j)&""" size=""8"" value="""&sLotserialNo&""" maxlength=10 READONLY class=""FormelemRead"" style=""text-align=center"">" )
'		headerCell.className="ExcelDisplayCell"
'		headerCell.appendChild(oText)
'		headerCell.width = "8"
'		headerCell.align="center"
		set oText = document.createElement("<input type=""hidden"" name=""txtSerialZ"&CStr(j)&""" value="""&sLotserialNo&""" >" )
		document.formname.appendchild(oText)

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text"" name=""txtGrossZ"&CStr(j)&""" onBlur=""RecalculateQty()"" size=""12"" value="""&sQtyRec&""" maxlength=10 class=""Formelem"" style=""text-align=right"">" )
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelInputCell"

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text"" name=""txtTareZ"&CStr(j)&""" onBlur=""RecalculateQty()"" size=""12"" value="""&sTar&""" maxlength=10 class=""Formelem"" style=""text-align=right"">")
		headerCell.className="ExcelInputCell"
		headerCell.width = "10"
		headerCell.appendChild(oText)

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<SELECT name=""selPackZ"&CStr(j)&""" class=""FormElem"" >" )

		Set RootO = PackingData.documentElement
		if RootO.hasChildNodes() then
			set oText1 = document.createElement("<Option>" )
			oText1.Text = "Select"
			oText1.Value = "select"
			oText.Options.Add(oText1)
			For Each HeaderONode In RootO.childNodes
				set oText1 = document.createElement("<Option>" )
				oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
				oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
				oText.Options.Add(oText1)
			next
		else
			set oText1 = document.createElement("<Option>" )
			oText1.Text = "Select"
			oText1.Value = "select"
			oText.Options.Add(oText1)
		end if
		headerCell.appendChild(oText)
		headerCell.className="ExcelFieldCell"
		headerCell.align="center"

		setIndex eval("document.formname.selPackZ"&CStr(j)),sPack

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text"" name=""txtPackNoZ"&CStr(j)&""" size=""17"" value="""&iPackNo&""" maxlength=30 class=""Formelem"" style=""text-align=right"">" )
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelInputCell"

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text"" name=""txtValueZ"&CStr(j)&""" size=""12"" value="""&iSubVal&"""  tabindex="&TabInd+5&" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
		headerCell.className="ExcelInputCell"
		headerCell.width = "10"
		headerCell.appendChild(oText)

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<SELECT name=""selSellZ"&CStr(j)&""" class=""FormElem"" >" )

		Set RootO = SellingData.documentElement
		if RootO.hasChildNodes() then
			set oText1 = document.createElement("<Option>" )
			oText1.Text = "Select"
			oText1.Value = "select"
			oText.Options.Add(oText1)
			For Each HeaderONode In RootO.childNodes
				set oText1 = document.createElement("<Option>" )
				oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
				oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
				oText.Options.Add(oText1)
			next
		else
			set oText1 = document.createElement("<Option>" )
			oText1.Text = "-N/A-"
			oText1.Value = "0"
			oText.Options.Add(oText1)
		end if
		headerCell.appendChild(oText)
		headerCell.className="ExcelFieldCell"
		headerCell.align="center"

		setIndex eval("document.formname.selSellZ"&CStr(j)),sSell

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text"" name=""txtSellWeightZ"&CStr(j)&""" size=""12"" maxlength=10 value="""&iWeight&""" class=""Formelem"" style=""text-align=right"">" )
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelInputCell"

'		if ucase(document.formname.hItemType.value)="FAB" then
'			set headerCell=oRow.insertCell()
'			set oText = document.createElement("<input type=""button"" name=""BtnEnterZ"&CStr(j)&""" value=""Yes"" onClick=""ShowPopup(this)"" class=""AddButtonX"">" )
'			headerCell.appendChild(oText)
'			headerCell.align="center"
'			headerCell.className="ExcelFieldCell"
'
'			'Code to load the FabricDetails
'			set FabricRoot = FabricData.documentElement
'			for each Node in StorageNode.item(iCounter).ChildNodes
'				set newElem = FabricData.createElement("FABRICDETAILS")
'				newelem.setAttribute "NO",j
'				newelem.setAttribute "QUANTITY",Node.attributes.getNamedItem("QUANTITY").value
'				newelem.setAttribute "QUANTITYIN",Node.attributes.getNamedItem("QUANTITYIN").value
'				newelem.setAttribute "CHECK",trim(Node.Attributes.getNamedItem("CHECK").value)
'				newelem.setAttribute "ALTGROSS",trim(Node.Attributes.getNamedItem("ALTGROSS").value)
'				newelem.setAttribute "ALTNETT",trim(Node.Attributes.getNamedItem("ALTNETT").value)
'				newelem.setAttribute "TOTGROSS",trim(Node.Attributes.getNamedItem("TOTGROSS").value)
'				newelem.setAttribute "TOTNETT",trim(Node.Attributes.getNamedItem("TOTNETT").value)
'
'				for each SubNode in Node.ChildNodes
'					set NewElem1 = FabricData.createElement("DETAILS")
'					NewElem1.setAttribute "PIECENO", SubNode.attributes.getNamedItem("PIECENO").value
'					NewElem1.setAttribute "QUANTITY", SubNode.attributes.getNamedItem("QUANTITY").value
'					NewElem1.setAttribute "ALTGROSS",SubNode.attributes.getNamedItem("ALTGROSS").value
'					NewElem1.setAttribute "ALTNETT",SubNode.attributes.getNamedItem("ALTNETT").value
'					newElem.appendChild NewElem1
'				next
'				FabricRoot.appendChild newElem
'			next
'		end if

	set headerCell=oRow.insertCell()
		set oText = document.createElement("<SELECT name=""selFormZ"&CStr(j)&""" class=""FormElem"" >" )

		Set RootO = SellingFormData.documentElement
		if RootO.hasChildNodes() then
			set oText1 = document.createElement("<Option>" )
			oText1.Text = "Select"
			oText1.Value = "select"
			oText.Options.Add(oText1)
			For Each HeaderONode In RootO.childNodes
				set oText1 = document.createElement("<Option>" )
				oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
				oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
				oText.Options.Add(oText1)
			next
		else
			set oText1 = document.createElement("<Option>" )
			oText1.Text = "-N/A-"
			oText1.Value = "0"
			oText.Options.Add(oText1)
		end if

		headerCell.appendChild(oText)
		headerCell.className="ExcelFieldCell"
		'headerCell.width = "10"
		headerCell.align="center"
		k = k+1
		j = j+1
	next
'	alert FabricRoot.xml
	next
	dim TempRoot,Node
	'Code for loading the PackingType Details
	Set TempRoot=PackingData.documentElement
	if TempRoot.hasChildNodes() then
		for each Node in TempRoot.childNodes
			document.formname.selPackType.length = document.formname.selPackType.length+1
			document.formname.selPackType.options(document.formname.selPackType.length-1).text = Node.attributes.getNamedItem("PACKSHNAME").value
			document.formname.selPackType.options(document.formname.selPackType.length-1).Value = Node.attributes.getNamedItem("PACKCODE").value
		next
	end if

	'Code for loading the SellingType Details
	Set TempRoot=SellingData.documentElement
	if TempRoot.hasChildNodes() then
		for each Node in TempRoot.childNodes
			document.formname.selSellType.length = document.formname.selSellType.length+1
			document.formname.selSellType.options(document.formname.selSellType.length-1).text = Node.attributes.getNamedItem("UNITNAME").value
			document.formname.selSellType.options(document.formname.selSellType.length-1).Value = Node.attributes.getNamedItem("UNITID").value
		next
	end if
	If document.formname.selSellType.length > 1 Then
		document.formname.txtWeight.disabled = false
	Else
		document.formname.txtWeight.disabled = True
	End If
	'Code for loading the SellingFormType Details
	Set TempRoot=SellingFormData.documentElement
	if TempRoot.hasChildNodes() then
		for each Node in TempRoot.childNodes
			document.formname.selForm.length = document.formname.selForm.length+1
			document.formname.selForm.options(document.formname.selForm.length-1).text = Node.attributes.getNamedItem("UNITNAME").value
			document.formname.selForm.options(document.formname.selForm.length-1).Value = Node.attributes.getNamedItem("UNITID").value
		next
	end if

	'Code for Disable the SellingType Details
	if document.formname.selSellType.length =1 then
		'document.formname.radSellType(1).disabled = true
		'document.formname.radSellType(0).disabled = true
		document.formname.selSellType.disabled = true
		'document.formname.radWeight(1).disabled = true
		'document.formname.radWeight(0).disabled = true
		'document.formname.txtWeight.disabled = true
	else
		'document.formname.radSellType(1).disabled = false
		'document.formname.radSellType(0).disabled = false
		document.formname.selSellType.disabled = false
		'document.formname.radWeight(1).disabled = false
		'document.formname.radWeight(0).disabled = false
		document.formname.txtWeight.disabled = false
	end if

	'Code for Disable the PackingType Details
	if document.formname.selPackType.length =1 then
		'document.formname.radPackType(1).disabled = true
		'document.formname.radPackType(0).disabled = true
		document.formname.selPackType.disabled = true
	else
		'document.formname.radPackType(1).disabled = false
		'document.formname.radPackType(0).disabled = false
		document.formname.selPackType.disabled = false
	end if

	'Code for Disable the SellingFormType Details
	if document.formname.selSellType.length =1 then
		'document.formname.radForm(1).disabled = true
		'document.formname.radForm(0).disabled = true
		'document.formname.selForm.disabled = true
	else
		'document.formname.radForm(1).disabled = false
		'document.formname.radForm(0).disabled = false
		'document.formname.selForm.disabled = false
	end if

	'Code for Disable the Stage Details
	if document.formname.selStage.length =1 then
		document.formname.selStage.disabled = true
	else
		document.formname.selStage.disabled = false
	end if

	UpdateLot

end Function

Function DisableTxt(obj)
	document.formname.txtTare.value = ""
	if obj.value = "I" then
		document.formname.txtTare.disabled = true
	else
		document.formname.txtTare.disabled = false
	end if
end Function


Function DisableSelSellType(obj)
	document.formname.selSellType.SelectedIndex=0
	if obj.value = "I" then
		document.formname.selSellType.disabled = true
	else
		document.formname.selSellType.disabled = false
	end if
end Function

Function DisableTxtWt(obj)
	document.formname.txtWeight.value = ""
	if obj.value = "I" then
		document.formname.txtWeight.disabled = true
	else
		document.formname.txtWeight.disabled = false
	end if
end Function

Function DisableselPack(obj)
	document.formname.selPackType.SelectedIndex=0
	if obj.value = "I" then
		document.formname.selPackType.disabled = true
	else
		document.formname.selPackType.disabled = false
	end if
end Function

Function DisableselForm(obj)
	'document.formname.selForm.SelectedIndex=0
	'if obj.value = "I" then
	'	document.formname.selForm.disabled = true
	'else
	'	document.formname.selForm.disabled = false
	'end if
end Function


Function DisableTxtQty(obj)
	document.formname.txtQuantity.value = ""
	if obj.value = "I" then
		document.formname.txtQuantity.disabled = true
	else
		document.formname.txtQuantity.disabled = false
	end if
end Function

Function DisableTxtValue(obj)
	document.formname.txtValue.value = ""
	if obj.value = "I" then
		document.formname.txtValue.disabled = true
	else
		document.formname.txtValue.disabled = false
	end if
end Function

Function CheckValue()
	CheckValue = False

	if Trim(sType) = "S" then
		if not (document.formname.radValue(1).checked or document.formname.radValue(0).checked) then
			alert("Select Value - Uniform or Individual")
			document.formname.radValue(1).focus
			CheckValue = True
		elseif (document.formname.radValue(1).checked and trim(document.formname.txtValue.value) = "") then
			alert("Enter Serial Value")
			document.formname.txtValue.focus
			CheckValue = True
		end if
	else
		if Trim(document.formname.txtValue.value) = "" then
			alert("Enter Lot Value")
			document.formname.txtValue.focus
			CheckValue = True
		end if
	end if
end Function

Function AddDetails(sCheck)
	Dim FType, objhttp, sLotNo, sSrlNo
	dim i,k,sQty, iTempVal, iLotNo
	Dim TabInd, autogen, SValue
	TabInd = 0
	STDate = document.formname.hDate.value
	'alert STDate
	If trim(document.formname.txtQtyIn.value) = "" Then
		alert "Select Gross Quantity"
		document.formname.txtQtyIn.focus
		Exit Function
	Elseif Not checkNumbers(Trim(document.formname.txtQtyIn.value)) Then
		alert "Enter Numerals Only"
		document.formname.txtQtyIn.focus
		Exit Function
	Elseif trim(document.formname.txtTare.value)="" then
		alert("Enter Tare Weight")
		document.formname.txtTare.focus
		Exit Function
	Elseif Not checkNumbers(trim(document.formname.txtTare.value)) then
		alert "Enter Numerals Only"
		document.formname.txtTare.focus
		Exit Function
	ElseIf trim(document.formname.txtValue.value) = "" Then
		alert "Enter Rate Per Unit"
		document.formname.txtValue.focus
		Exit Function
	ElseIf Not checkNumbers(trim(document.formname.txtValue.value)) Then
		alert "Enter Rate Per Unit"
		document.formname.txtValue.focus
		Exit Function
	Elseif trim(document.formname.txtWeight.value)="" then
		alert("Enter Weight Per Form")
		document.formname.txtWeight.focus
		Exit Function
	Elseif Not checkNumbers(trim(document.formname.txtWeight.value)) then
		alert "Enter Numerals Only"
		document.formname.txtWeight.focus
		Exit Function
	ElseIf (trim(document.formname.txtNoOfPacks.value)) = "" Then
		alert ("Enter No. of Packs")
		document.formname.txtNoOfPacks.focus
		Exit Function
	Elseif Not checkNumbers(Trim(document.formname.txtNoOfPacks.value)) Then
		alert "Enter Numerals Only"
		document.formname.txtNoOfPacks.focus
		Exit Function
	ElseIf (trim(document.formname.txtRcptNumbering.value) = "LOT\SERIAL" or trim(document.formname.txtRcptNumbering.value) = "LOT") and (Trim(document.formname.txtLotNumber.value) = "" and Trim(document.formname.chkNew.checked=false)) Then
		alert "Select\Enter Lot Number"
		document.formname.txtRcptNumbering.focus
		Exit Function
	Elseif document.formname.selPackType.selectedIndex = 0 then
		alert("Select Pack Type")
		document.formname.selPackType.focus
		Exit function
	Elseif trim(document.formname.txtWeightPerPack.value)="" then
		 alert("Enter Weight Per Pack")
		 document.formname.txtWeightPerPack.focus
		 Exit Function
	ElseIf Not checkNumbers(trim(document.formname.txtWeightPerPack.value)) Then
		alert "Enter Numerals Only"
		document.formname.txtWeightPerPack.focus
		Exit Function
	Elseif trim(document.formname.txtPackTare.value)="" then
		 alert("Enter Tare Pack Weight")
		 document.formname.txtPackTare.focus
		 Exit Function
	ElseIf Not checkNumbers(trim(document.formname.txtPackTare.value)) Then
		alert "Enter Numerals Only"
		document.formname.txtPackTare.focus
		Exit Function
	End If

	If document.formname.chkNew.checked  Then
		autogen = True
	Else
		autogen = False
	End IF
		iLotNo = document.formname.txtNoOfPacks.value

		No = document.all.tblLot.rows.length - 2 - iFixedNo

		CheckAvail

		If avail = False Then Exit Function

		if iLotNo = 0 then

			j = No

			'alert sLotNo
			if sLotNo = "" then
				sLotNo = document.formname.txtLotNumber.value

				iFixedNo = iFixedNo + 1

				set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
				set headerCell=oRow.insertCell()
				headerCell.innerHTML=sLotNo & " / "
				headerCell.className="ExcelDisplayCell"
				if Trim(sType) <> "S" then
					set oText = document.createElement("<input type=""text"" name=""txtIValueZ"&CStr(iFixedNo)&""" size=""12"" value="""&document.formname.txtValue.value&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
					headerCell.width = "8"
					headerCell.appendChild(oText)
				end if
				headerCell.align="left"
				headerCell.colspan=Cols
			elseif sLotNo <> document.formname.txtLotNumber.value then

				sLotNo = document.formname.txtLotNumber.value
				iFixedNo = iFixedNo + 1
				set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
				set headerCell=oRow.insertCell()
				headerCell.innerHTML=sLotNo & " / "
				headerCell.className="ExcelDisplayCell"
				if Trim(sType) <> "S" then
					set oText = document.createElement("<input type=""text"" name=""txtIValueZ"&CStr(iFixedNo)&""" size=""12"" value="""&document.formname.txtValue.value&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
					headerCell.width = "8"
					headerCell.appendChild(oText)
				end if
				headerCell.colspan=Cols
				headerCell.align="left"
			end if
			sPubLot = sLotNo
			set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=j
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"
			exit function

'
		'	set oText = document.createElement("<input type=""hidden"" name=""txtSerial"&CStr(j)&""" value="""&trim(document.formname.txtSerialFrom.value)&""" >" )
			set oText = document.createElement("<input type=""hidden"" name=""txtSerialZ"&CStr(j)&""" value="" >" )
			document.formname.appendchild(oText)

			set headerCell=oRow.insertCell()

			'set oText = document.createElement("<input type=""text"" name=""txtGross1"" size=""12"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">" )
			if document.formname.radQty(1).checked then
				set oText = document.createElement("<input type=""text"" name=""txtGrossZ"&CStr(j)&""" size=""12"" onBlur=""RecalculateQty()"" value="""&trim(document.formname.txtQuantity.value)&""" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
			elseif document.formname.radQty(0).checked then
				set oText = document.createElement("<input type=""text"" name=""txtGrossZ"&CStr(j)&""" size=""12"" onBlur=""RecalculateQty()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
			end if

			headerCell.width = "10"
			headerCell.appendChild(oText)
			headerCell.className="ExcelInputCell"

			set headerCell=oRow.insertCell()

			set oText = document.createElement("<input type=""text"" name=""txtTareZ"&CStr(j)&""" size=""12"" value=""0"" READONLY class=""FormelemRead"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
			headerCell.className="ExcelDisplayCell"

			headerCell.width = "10"
			headerCell.appendChild(oText)

			set headerCell=oRow.insertCell()

			set oText = document.createElement("<SELECT name=""selSellZ"&CStr(j)&""" class=""FormElem"" >" )

			Set RootO = SellingData.documentElement
			'alert(RootO.xml)
			if RootO.hasChildNodes() then
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
				For Each HeaderONode In RootO.childNodes
					set oText1 = document.createElement("<Option>" )
					oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
					oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
					oText.Options.Add(oText1)
				next
			else
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
			end if

			headerCell.className="ExcelFieldCell"
			headerCell.align="center"
			headerCell.appendChild(oText)

			set headerCell=oRow.insertCell()
			if (document.formname.radWeight(1).Checked=true  and document.formname.radWeight(1).disabled=false) then
				set oText = document.createElement("<input type=""text"" name=""txtSellWeightZ"&CStr(j)&""" size=""12"" value="""&trim(document.formname.txtWeight.value)&""" class=""Formelem"" style=""text-align=right"">")
			else
				set oText = document.createElement("<input type=""text"" name=""txtSellWeightZ"&CStr(j)&""" size=""12"" maxlength=10 class=""Formelem"" style=""text-align=right"">" )
			end if
			headerCell.className="ExcelInputCell"
			headerCell.appendChild(oText)
			headerCell.width = "10"

			set headerCell=oRow.insertCell()

			set oText = document.createElement("<SELECT name=""selPackZ"&CStr(j)&""" class=""FormElem"" >" )

			Set RootO = PackingData.documentElement
			if RootO.hasChildNodes() then
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
				For Each HeaderONode In RootO.childNodes
					set oText1 = document.createElement("<Option>" )
					oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
					oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
					oText.Options.Add(oText1)
				next
			else
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
			end if
			headerCell.className="ExcelFieldCell"
			'headerCell.width = "10"
			headerCell.align="center"
			headerCell.appendChild(oText)


			set headerCell=oRow.insertCell()

			set oText = document.createElement("<SELECT name=""selFormZ"&CStr(j)&""" class=""FormElem"" >" )

			Set RootO = SellingFormData.documentElement
			if RootO.hasChildNodes() then
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
				For Each HeaderONode In RootO.childNodes
					set oText1 = document.createElement("<Option>" )
					oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
					oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
					oText.Options.Add(oText1)
				next
			else
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
			end if
			headerCell.className="ExcelFieldCell"
			headerCell.align="center"
			headerCell.appendChild(oText)

			set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtPackNoZ"&CStr(j)&""" value="""&trim(document.formname.txtLotNumber.value)&""" size=""17"" maxlength=30 class=""Formelem"" onBlur=""PackNoLocalCheck1(this,"&CStr(j)&")"" style=""text-align=left"">" )
			headerCell.className="ExcelInputCell"
			headerCell.appendChild(oText)
			headerCell.width = "10"


'			if ucase(document.formname.hItemType.value)="FAB" then
'				set headerCell=oRow.insertCell()
'				set oText = document.createElement("<input type=""button"" name=""BtnEnterZ"&CStr(j)&""" onClick=""ShowPopup(this)""  value=""Yes"" class=""AddButtonX"">" )
'				headerCell.appendChild(oText)
'				headerCell.align="center"
'				headerCell.className="ExcelFieldCell"
'
'				sQty = "N:" & document.formname.txtQtyIn.value
'
'				set FabricRoot=FabricData.documentElement
'				set newelem = FabricData.createElement("FABRICDETAILS")
'				newelem.setAttribute "NO",j
'				newelem.setAttribute "QUANTITYIN",sQty
'
'				newelem.setAttribute "CHECK",""
'				newelem.setAttribute "ALTGROSS",""
'				newelem.setAttribute "ALTNETT",""
'				newelem.setAttribute "TOTGROSS",""
'				newelem.setAttribute "TOTNETT",""
'
'				FabricRoot.appendChild newelem
'			end if



			eval("document.formname.txtGrossZ"&CStr(j)).focus

			j = j + 1
			exit function
		end if

		'ClearTable

		k = cdbl(document.formname.txtNoOfPacks.value)

			if sLotNo = "" then

				sLotNo = document.formname.txtLotNumber.value

				iFixedNo = iFixedNo + 1

				If autogen = True and (trim(document.formname.txtRcptNumbering.value) = "LOT" or trim(document.formname.txtRcptNumbering.value)="LOT\SERIAL") Then
'					FType = "LO"  '"SL"
'					set objhttp = CreateObject("MSXML2.XMLHTTP")
'					objhttp.Open "GET","../../Inventory/transaction/LotSerialGenrate.asp?orgID="&sOrgID&"&iItem="&iItem&"&FTYPE="&FType&"&STDATE="&STDate,False
'					objhttp.send
'					If Trim(objhttp.responsetext) <> "" Then
'						sLotNo = Trim(objhttp.responsetext)
'						'alert sLotNo
'						'sSrlNo = Trim(objhttp.responsetext)
'						'set oText = document.createElement("<input type=""text"" name=""txtIValue"&CStr(iFixedNo)&""" size=""12"" value="""&sLotNo&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
'					End If
					AutoLot = AutoLot + 1
					sLotNo = AutoLot
				End If
				'If Trim(sLotNo) <> Trim(sOldLotNo) Then
			'	alert " sLotNo ="& sLotNo
				If Not LotExist(sLotNo) Then

					set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
					set headerCell=oRow.insertCell()

					if Trim(sType) <> "S" then
					'	alert cDbl(document.formname.txtQtyIn.value)
					'	alert cDbl(document.formname.txtTare.value)
					'	sLotNo = sLotNo & " " & FormatNumber((CDbl(document.formname.txtQtyIn.value)-cDbl(document.formname.txtTare.value))*cdBL(document.formname.txtValue.value),2,,,0)
						NewVal = FormatNumber((CDbl(document.formname.txtQtyIn.value)-cDbl(document.formname.txtTare.value))*cdBL(document.formname.txtValue.value),2,,,0)
						headerCell.innerHTML=sLotNo &  " / "
						headerCell.className="ExcelDisplayCell"
					end if
					NewVal =  FormatNumber((CDbl(document.formname.txtQtyIn.value)-cDbl(document.formname.txtTare.value))*cdBL(document.formname.txtValue.value),2,,,0)
					'alert("NewVal="&NewVal)
'					set oText = document.createElement("<input type=""text"" name=""txtIValue"&CStr(iFixedNo)&""" size=""12"" value="""&cDbl(document.formname.txtValue.value)*cDbl(document.formname.txtNoOfPacks.value) &""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					set oText = document.createElement("<input type=""text"" name=""txtIValueZ"&CStr(iFixedNo)&""" size=""12"" value="""&NewVal&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
					headerCell.width = "100"
					headerCell.appendChild(oText)

					set oText = document.createElement("<input type=""text"" name=""txtValZ"&CStr(iFixedNo)&""" size=""10"" value="""&" ["&cDbl(document.formname.txtValue.value)&"]"&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=Left"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
					headerCell.width = "100"
					headerCell.appendChild(oText)

					headerCell.align="left"
					headerCell.colspan=Cols

				End If
				'End If
			elseif sLotNo <> document.formname.txtLotNumber.value then

				sLotNo = document.formname.txtLotNumber.value
				If Not LotExist(sLotNo) Then
					iFixedNo = iFixedNo + 1
					set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)
					set headerCell=oRow.insertCell()

					if Trim(sType) <> "S" then
					'	sLotNo = sLotNo & " " & FormatNumber((CDbl(document.formname.txtQtyIn.value)-cDbl(document.formname.txtTare.value))*cdBL(document.formname.txtValue.value),2,,,0)
						NewVal = FormatNumber((CDbl(document.formname.txtQtyIn.value)-cDbl(document.formname.txtTare.value))*cdBL(document.formname.txtValue.value),2,,,0)
						headerCell.innerHTML=sLotNo &  " / "
						headerCell.className="ExcelDisplayCell"
					end if

					NewVal = FormatNumber((CDbl(document.formname.txtQtyIn.value)-cDbl(document.formname.txtTare.value))*cdBL(document.formname.txtValue.value),2,,,0)
					'set oText = document.createElement("<input type=""text"" name=""txtIValue"&CStr(iFixedNo)&""" size=""12"" value="""&document.formname.txtValue.value&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					set oText = document.createElement("<input type=""text"" name=""txtIValueZ"&CStr(iFixedNo)&""" size=""12"" value="""&NewVal&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					headerCell.className="ExcelDisplayCell"

					headerCell.align="center"
					headerCell.width = "8"
					headerCell.appendChild(oText)

					set oText = document.createElement("<input type=""text"" name=""txtValZ"&CStr(iFixedNo)&""" size=""12"" value="""&" ["&document.formname.txtValue.value&"]"&""" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					headerCell.className="ExcelDisplayCell"

					headerCell.align="center"
					headerCell.width = "8"
					headerCell.appendChild(oText)


					headerCell.align="left"
					headerCell.colspan=Cols
				End If
			end if
			sPubLot = sLotNo

		'for j = No  to No + iLotNo

		for j = No  to No + iLotNo -1

			set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length)

			iCnt = iCnt + 1

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=j
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtGrossZ"&CStr(j)&""" size=""12"" tabindex="&TabInd+1&" value="""&trim(document.formname.txtWeightPerPack.value)&""" class=""Formelem"" onBlur=""RecalculateTotal()"" style=""text-align=right"">")
			headerCell.className="ExcelInputCell"
			headerCell.width = "10"
			headerCell.appendChild(oText)

			set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtTareZ"&CStr(j)&""" size=""12"" tabindex="&TabInd+2&"  value="""&trim(document.formname.txtPackTare.value)&""" class=""Formelem"" style=""text-align=right""  onBlur=""RecalculateTotal()"" >")
				headerCell.className="ExcelInputCell"
			headerCell.width = "10"
			headerCell.appendChild(oText)

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<SELECT name=""selPackZ"&CStr(j)&""" class=""FormElem"" >" )

			Set RootO = PackingData.documentElement
			if RootO.hasChildNodes() then
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
				For Each HeaderONode In RootO.childNodes
					set oText1 = document.createElement("<Option>" )
					oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
					oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
					oText.Options.Add(oText1)
				next
			else
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "N/A"
				oText1.Value = "N/A"
				oText.Options.Add(oText1)
			end if

			headerCell.className="ExcelFieldCell"
			'headerCell.width = "10"
			headerCell.align="center"
			headerCell.appendChild(oText)

			setIndex eval("document.formname.selPackZ"&CStr(j)),document.formname.selPackType.options(document.formname.selPackType.selectedIndex).Value

			If autogen = True and (trim(document.formname.txtRcptNumbering.value) = "LOT" or trim(document.formname.txtRcptNumbering.value)="LOT\SERIAL"  or trim(document.formname.txtRcptNumbering.value)="SERIAL") Then
				AutoSerial = AutoSerial + 1
				sSrlNo = AutoSerial
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtPackNoZ"&CStr(j)&""" maxlength=30 class=""Formelem"" Value="&sSrlNo&" tabindex="&TabInd+4&" onBlur=""PackNoLocalCheck1(this,"&CStr(j)&")"" style=""text-align=right"">" )
				headerCell.className="ExcelInputCell"
				headerCell.appendChild(oText)
				headerCell.width = "10"
'				FType = "SL"
'				set objhttp = CreateObject("MSXML2.XMLHTTP")
'				objhttp.Open "GET","../../Inventory/transaction/LotSerialGenrate.asp?orgID="&sOrgID&"&iItem="&iItem&"&FTYPE="&FType&"&STDATE="&STDate,False
'				objhttp.send
'				If Trim(objhttp.responsetext) <> "" Then
'					sSrlNo = Trim(objhttp.responsetext)
'						set headerCell=oRow.insertCell()
'						'set oText = document.createElement("<input type=""text"" name=""txtPackNo"&CStr(j)&""" value="&CalculatePackNo(k)&" size=""17"" maxlength=30 class=""Formelem"" onBlur=""PackNoLocalCheck1(this,"&CStr(j)&")"" style=""text-align=left"">" )
'						set oText = document.createElement("<input type=""text"" name=""txtPackNoZ"&CStr(j)&""" maxlength=30 class=""Formelem"" Value="&sSrlNo&" tabindex="&TabInd+4&" onBlur=""PackNoLocalCheck1(this,"&CStr(j)&")"" style=""text-align=right"">" )
'						headerCell.className="ExcelInputCell"
'						headerCell.appendChild(oText)
'						headerCell.width = "10"
'				End If
			Else
				set headerCell=oRow.insertCell()
				'set oText = document.createElement("<input type=""text"" name=""txtPackNo"&CStr(j)&""" value="&CalculatePackNo(k)&" size=""17"" maxlength=30 class=""Formelem"" onBlur=""PackNoLocalCheck1(this,"&CStr(j)&")"" style=""text-align=left"">" )
				set oText = document.createElement("<input type=""text"" name=""txtPackNoZ"&CStr(j)&""" maxlength=30 class=""Formelem""  tabindex="&TabInd+4&" onBlur=""PackNoLocalCheck1(this,"&CStr(j)&")"" style=""text-align=right"">" )
				headerCell.className="ExcelInputCell"
				headerCell.appendChild(oText)
				headerCell.width = "10"
			End If

	'		If Trim(sType) = "S" then
				set headerCell=oRow.insertCell()
				'Individual Value
'				if document.formname.radValue(0).checked then
'					set oText = document.createElement("<input type=""text"" name=""txtIValue"&CStr(j)&""" size=""12"" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
'					headerCell.className="ExcelInputCell"
'				'Uniform Value
'				elseif document.formname.radValue(1).checked then
					SValue = FormatNumber((cDbl(document.formname.txtWeightPerPack.value)-cDbl(document.formname.txtPackTare.value)) * CDbl(document.formname.txtValue.value),2,,0)
					'set oText = document.createElement("<input type=""text"" name=""txtValue"&CStr(j)&""" size=""12"" value="""&iTempVal&"""  tabindex="&TabInd+5&" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
					set oText = document.createElement("<input type=""text"" name=""txtValueZ"&CStr(j)&""" size=""12"" value="""&SValue&"""  tabindex="&TabInd+5&" onBlur=""RecalculateTotal()"" class=""Formelem"" style=""text-align=right"" >")
					headerCell.className="ExcelInputCell"
'				end if

				headerCell.width = "10"
				headerCell.appendChild(oText)


		'	if ucase(document.formname.hItemType.value)="FAB" then
		'		set headerCell=oRow.insertCell()
		'		set oText = document.createElement("<input type=""button"" name=""BtnEnterZ"&CStr(j)&""" onClick=""ShowPopup(this)"" value=""Yes"" class=""AddButtonX"">" )
		'		headerCell.appendChild(oText)
		'		headerCell.align="center"
		'		headerCell.className="ExcelFieldCell"

		'			sQty = "N:" & document.formname.txtQtyIn.value
		'
		'		set FabricRoot=FabricData.documentElement
		'		set newelem = FabricData.createElement("FABRICDETAILS")
		'		newelem.setAttribute "NO",j
		'		newelem.setAttribute "QUANTITYIN",sQty

		'		newelem.setAttribute "CHECK",""
		'		newelem.setAttribute "ALTGROSS",""
		'		newelem.setAttribute "ALTNETT",""
		'		newelem.setAttribute "TOTGROSS",""
		'		newelem.setAttribute "TOTNETT",""
'
'				FabricRoot.appendChild newelem
'
'			end if

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<SELECT name=""selSellZ"&CStr(j)&""" class=""FormElem"" >" )
			Set RootO = SellingData.documentElement
			if RootO.hasChildNodes() then
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
				For Each HeaderONode In RootO.childNodes
					set oText1 = document.createElement("<Option>" )
					oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
					oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
					oText.Options.Add(oText1)
				next
			else
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "-N/A-"
				oText1.Value = "0"
				oText.Options.Add(oText1)
			end if
			headerCell.className="ExcelFieldCell"
			headerCell.align="center"
			headerCell.appendChild(oText)

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtSellWeightZ"&CStr(j)&""" size=""12""  tabindex="&TabInd+3&" value="""&trim(document.formname.txtWeight.value)&"""  class=""Formelem"" style=""text-align=right"">")
			headerCell.className="ExcelInputCell"
			headerCell.appendChild(oText)
			headerCell.width = "10"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<SELECT name=""selFormZ"&CStr(j)&""" class=""FormElem"" >" )

			Set RootO = SellingFormData.documentElement
			if RootO.hasChildNodes() then
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)
				For Each HeaderONode In RootO.childNodes
					set oText1 = document.createElement("<Option>" )
					oText1.Text = trim(HeaderONode.Attributes.Item(1).nodeValue)
					oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)
					oText.Options.Add(oText1)
				next
			else
				set oText1 = document.createElement("<Option>" )
				oText1.Text = "-N/A-"
				oText1.Value = "0"
				oText.Options.Add(oText1)
			end if
			headerCell.className="ExcelFieldCell"
			headerCell.align="center"
			headerCell.appendChild(oText)

			k = k+1
			sOldLotNo = sLotNo
		next
		eval("document.formname.txtGrossZ"&CStr(No)).focus

		AddLot
		UpdateLot

		ct = ct +1
end Function

Function ClearTable()
	dim i
	for i=2 to document.all.tblLot.rows.length - 1
		document.all.tblLot.deleteRow(2)
	next
end Function

Function clearAll()
	'document.formname.txtQuantity.value = ""
	document.formname.txtTare.value = "0"
	document.formname.txtLotNumber.value = ""
	document.formname.txtValue.value = "0"
	document.formname.txtWeight.value = "0"
	'document.formname.selForm.selectedIndex = 0
	document.formname.selPackType.selectedIndex = 0
	document.formname.selSellType.selectedIndex = 0
	document.formname.selStage.selectedIndex = 0
end Function

Function CheckAvail()
	avail = True
	dim i,iTempTotGross,iTempTotTare, iTempTotValue
	dim objQ,objD,objSerial,TempRoot,Node, FabricRoot, FabricNode, objValue
	if (document.all.tblLot.rows.length = 2 ) then
		alert("No Details Entered")
		exit function
	end if
	document.formname.txtLotNumber.ReadOnly = false
	iQtyTotGross = 0
	iQtyTotTare = 0
	iTempTotGross=0
	iTempTotTare=0
	iTempTotValue=0
	for i = 1 to iCnt
		set objD = eval("document.formname.txtGrossZ"&CStr(i))
		set objQ = eval("document.formname.txtTareZ"&CStr(i))
		set objSel = eval("document.formname.selSellZ"&CStr(i))
		set objWei = eval("document.formname.txtSellWeightZ"&CStr(i))
		set objPac = eval("document.formname.selPackZ"&CStr(i))
		set objPackNo = eval("document.formname.txtPackNoZ"&CStr(i))
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
		'elseif (objSel.length > 1 and objSel.selectedIndex = 0) then
		'	alert("Select Selling Type")
		'	objSel.focus()
		'	exit function
		elseif (objSel.selectedIndex > 0 and trim(objWei.value) = "") then
			alert("Enter Weight Of Selling Type")
			objWei.select()
			exit function
		elseif (objSel.selectedIndex > 0 and trim(objWei.value) <> "") and not checkNumbers(objWei.value) then
			alert("Enter Numerals Only")
			objWei.select()
			exit function
		elseif PackNoCheck(objPackNo,objPac) then
			alert("Packing Number already Exists")
			objPackNo.select()
			exit function
		else
			iTempTotGross = cdbl(iTempTotGross) + cdbl(objD.value)
			iTempTotTare = cdbl(iTempTotTare) + cdbl(objQ.value)
		end if
	next
	'alert cdbl(iQtyTotGross)

	'alert cdbl(iTempTotGross)
	'alert CDbl(document.formname.txtqtyin.value)
	'alert(iQtyTotGross &"+"& iTempTotGross &"+"& document.formname.txtWeightPerPack.value &"-"& document.formname.txtPackTare.value)
	iQtyTotGross = cdbl(iQtyTotGross) + cdbl(iTempTotGross) + CDbl(document.formname.txtWeightPerPack.value)- CDbl(document.formname.txtPackTare.value)
	iQtyTotTare = cdbl(iQtyTotTare) + cdbl(iTempTotTare)
	'alert(iQtyTotGross &" - "& iQtyTotTare &"  > " &idQty.innerText)
	''Blocked By Ragav Feb 27,2012
	''begin
	'if (cdbl(iQtyTotGross-iQtyTotTare) > cdbl(idQty.innerText)) then
	'	alert("Total Received Quantity should be Less than or equal to Quantity (" &idQty.innerText& ")")
	'	avail = False
	'	exit function
	'end if
	''end

End Function


Function AddLot()
	dim i,iTempTotGross,iTempTotTare, iTempTotValue
	dim objQ,objD,objSerial,TempRoot,Node, FabricRoot, FabricNode, objValue
	dim autogen, k
	dim objhttp, Root1, sSrlNo,sAttributeID
'	Set objhttp = CreateObject("Microsoft.XMLHTTP")
'		objhttp.Open "POST","DirectIssueInsert.asp", false
'		objhttp.send OutData.XMLDocument


	'alert Root1.XML


	Set Root = SerialData.documentElement
	'alert Root.XML
	if Root.hasChildNodes  then
		CheckExists()
		ClearAll()
		exit function
	end if
	'alert Root.XML
	if (document.all.tblLot.rows.length = 2 ) then
		alert("No Details Entered")
		exit function
	end if
	If document.formname.chkNew.checked  Then
		autogen = True
	Else
		autogen = False
	End IF


	document.formname.txtLotNumber.ReadOnly = false
	iTempTotGross=0
	iTempTotTare=0
	iTempTotValue=0
'	alert(No)
'	alert(j)

	for i = No to j - 1
		set objD = eval("document.formname.txtGrossZ"&CStr(i))
		set objQ = eval("document.formname.txtTareZ"&CStr(i))

		set objSel = eval("document.formname.selSellZ"&CStr(i))
		set objWei = eval("document.formname.txtSellWeightZ"&CStr(i))
		set objPac = eval("document.formname.selPackZ"&CStr(i))
		'set objFor = eval("document.formname.selFormZ"&CStr(i))
		set objPackNo = eval("document.formname.txtPackNoZ"&CStr(i))

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
		'elseif (objSel.length > 1 and objSel.selectedIndex = 0) then
		'	alert("Select Selling Type")
		'	objSel.focus()
		'	exit function
		elseif (objSel.selectedIndex > 0 and trim(objWei.value) = "") then
			alert("Enter Weight Of Selling Type")
			objWei.select()
			exit function
		elseif (objSel.selectedIndex > 0 and trim(objWei.value) <> "") and not checkNumbers(objWei.value) then
			alert("Enter Numerals Only")
			objWei.select()
			exit function

		elseif PackNoCheck(objPackNo,objPac) then
			alert("Packing Number already Exists")
			objPackNo.select()
			exit function
		else
			iTempTotGross = cdbl(iTempTotGross) + cdbl(objD.value)
			iTempTotTare = cdbl(iTempTotTare) + cdbl(objQ.value)
		end if
	next


	iQtyTotGross = cdbl(iQtyTotGross) + cdbl(iTempTotGross)
	iQtyTotTare = cdbl(iQtyTotTare) + cdbl(iTempTotTare)

	RecalculateQty()
	'alert("Gross Qty : " & iQtyTotGross)
	'alert("Tare Qty : " & iQtyTotTare)
	'alert("Diff of Gross and Tare Qty : " & cdbl(iQtyTotGross-iQtyTotTare))
	'alert("Difference Qty : " & cdbl(iQtyTotGross-iQtyTotTare) - cdbl(idQty.innerText))
	'alert cdbl(iQtyTotGross-iQtyTotTare)
'	alert idQty.innerText
'alert(cdbl(iQtyTotGross-iQtyTotTare))
'alert(cdbl(idQty.innerText))
''Blocked By Ragav Feb 27,2012
''begin
'	if (cdbl(iQtyTotGross-iQtyTotTare) > cdbl(idQty.innerText)) then
'		ClearTable()
'		LoadDetails(document.formname.sTemp.value)
'		alert("Total Received Quantity should be Less than or equal to Quantity (" &idQty.innerText& ")")
''		iQtyTotGross = cdbl(iQtyTotGross) - cdbl(iTempTotGross)
''		iQtyTotTare = cdbl(iQtyTotTare) - cdbl(iTempTotTare)
'		exit function
'	end if
''end
	idQtyEntered.innerText = FormatNumber(cdbl(iQtyTotGross) - cdbl(iQtyTotTare),3,,,0)
	i = 1

'	if ucase(document.formname.hItemType.value)="FAB" then
'		if CheckFabricData then exit function
'	end if

	Set Root = objTemp.documentElement

	For Each childNod In Root.childNodes

		if childNod.Attributes.Item(0).nodeValue = iStore and childNod.Attributes.Item(1).nodeValue = iBin then
			childNod.setAttribute "QTY", cdbl(iTempTotGross+iTempTotTare)
				If LotExist(sPubLot) Then   'Lot Exist
					For Each GrandChildNod In ChildNod.childNodes  'GrandChild
						if cstr(GrandChildNod.attributes.getNamedItem("LOT").value) = cstr(sPubLot) then    'LotValCheck

						For k = cDbl(iCnt)-cDbl(document.formname.txtNoOfPacks.value)+ 1 To cDbl(iCnt)
						i = k
						set objD = eval("document.formname.txtGrossZ"&CStr(i))
						set objQ = eval("document.formname.txtTareZ"&CStr(i))

						set objSel = eval("document.formname.selSellZ"&CStr(i))
						set objWei = eval("document.formname.txtSellWeightZ"&CStr(i))
						set objPac = eval("document.formname.selPackZ"&CStr(i))
						set objFor = eval("document.formname.selFormZ"&CStr(i))
						set objPackNo = eval("document.formname.txtPackNoZ"&CStr(i))

						if document.formname.hrec.value="S"  then set objValue = eval("document.formname.txtValueZ"&CStr(i))
						if document.formname.hrec.value="LS"  then set objValue = eval("document.formname.txtValueZ"&CStr(i))
						if document.formname.hrec.value="L"  then set objValue = eval("document.formname.txtValueZ"&CStr(i))

						sAttributeID = document.formname.hAttributeList.value
						
						if sAttributeID<>"NULL" and sAttributeID<>"0" and not IsNull(sAttributeID) and sAttributeID<>"" then
						    if document.formname.selAttribute.selectedIndex > 0 then
							    sAttributeID = document.formname.selAttribute(document.formname.selAttribute.selectedIndex).value
						    else
						        sAttributeID = ""
						    end if
						end if


						Set newElem1 = objTemp.createElement("LotSerialDetails")
						newElem1.setAttribute "LOTSERIAL", i
						newElem1.setAttribute "QTYREC", trim(objD.value)
						newElem1.setAttribute "TAREREC", trim(objQ.value)
						newElem1.setAttribute "SELLINGTYPE", trim(objSel.value)
						newElem1.setAttribute "WEIGHTSTYPE", trim(objWei.value)
						newElem1.setAttribute "PACKINGTYPE", trim(objPac.value)
						newElem1.setAttribute "LOT",sPubLot  '   trim(document.formname.txtLotNumber.value)
						newElem1.setAttribute "SELLINGFORM", trim(objFor.value)
						newElem1.setAttribute "PACKNUMBER", trim(objPackNo.value)
						if document.formname.hrec.value="S" then newElem1.setAttribute "IVALUE", trim(objValue.value)
						if document.formname.hrec.value="LS" then newElem1.setAttribute "IVALUE", trim(objValue.value)
						if document.formname.hrec.value="L" then newElem1.setAttribute "IVALUE", trim(objValue.value)
						newElem.setAttribute "ATTRIBUTELIST",trim(sAttributeID)


						iQtyTotGross = cdbl(iQtyTotGross) + cdbl(objD.value)
						iQtyTotTare = cdbl(iQtyTotTare) + cdbl(objQ.value)

					'	if ucase(document.formname.hItemType.value)="FAB" then
					'	set FabricRoot = FabricData.documentElement
					'	for each FabricHeaderNode in FabricRoot.childNodes
					'		if cstr(FabricHeaderNode.attributes.getNamedItem("NO").value) = cstr(i) then
					'			set newElem2 = objTemp.createElement("FABRICDETAILS")
					'			newelem2.setAttribute "NO",i
					'			newelem2.setAttribute "QUANTITY",FabricHeaderNode.attributes.getNamedItem("QUANTITY").value
					'			newelem2.setAttribute "QUANTITYIN",FabricHeaderNode.attributes.getNamedItem("QUANTITYIN").value
					'			newelem2.setAttribute "CHECK",trim(FabricHeaderNode.Attributes.getNamedItem("CHECK").value)
					'			newelem2.setAttribute "ALTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("ALTGROSS").value)
					'			newelem2.setAttribute "ALTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("ALTNETT").value)
					'			newelem2.setAttribute "TOTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("TOTGROSS").value)
					'			newelem2.setAttribute "TOTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("TOTNETT").value)
'
'								for each FabricNode in FabricHeaderNode.childNodes
'									set NewElem3 = objTemp.createElement("DETAILS")
'									NewElem3.setAttribute "PIECENO", FabricNode.attributes.getNamedItem("PIECENO").value
'									NewElem3.setAttribute "QUANTITY", FabricNode.attributes.getNamedItem("QUANTITY").value
'									NewElem3.setAttribute "ALTGROSS",FabricNode.attributes.getNamedItem("ALTGROSS").value
'									NewElem3.setAttribute "ALTNETT",FabricNode.attributes.getNamedItem("ALTNETT").value
'
'									newElem2.appendChild NewElem3
'								next
'								newElem1.appendChild newElem2
'							end if
'						Next
'					End If '	if ucase(document.formname.hItemType.value)="FAB" then
						GrandChildNod.appendChild newElem1

					'	if ucase(document.formname.hItemType.value)="FAB" then
					'		GrandChildNod.setAttribute "QTY", cdbl(iQtyTotGross)
					'	else
							GrandChildNod.setAttribute "QTY", cdbl(iQtyTotGross-iQtyTotTare)
					'	end if

						iCounter = cdbl(iCounter) + 1
						GrandChildNod.setAttribute "COUNTER", cdbl(iCounter)
						GrandChildNod.setAttribute "STAGE", document.formname.selStage.value
						GrandChildNod.setAttribute "ALTGROSS", document.formname.txtAltGross.value
						GrandChildNod.setAttribute "ALTNETT", document.formname.txtAltNett.value
						GrandChildNod.setAttribute "ALTUOM", document.formname.selAltUom.value
						GrandChildNod.setAttribute "IVALUE", document.formname.txtValue.value
						If document.formname.chkNew.checked Then
							GrandChildNod.setAttribute "AUTOGEN", "AUTO"
						Else
							GrandChildNod.setAttribute "AUTOGEN", ""
						End If
						childNod.appendChild GrandChildNod
						Next

						End If 'LotValCheck
						UpdateValues
				Next   ' GrandChild

				ClearTable()
				LoadDetails(document.formname.sTemp.value)


				Exit Function

			End If	' Lot Exist





			set newElem = objTemp.createElement("LotSerial")
			newElem.setAttribute "QTYIN", "N"
			newElem.setAttribute "TARE", trim(document.formname.txtTare.value)
			newElem.setAttribute "LOT", sPubLot ' trim(document.formname.txtLotNumber.value)
			newElem.setAttribute "SERIALFROM", "" 'trim(document.formname.txtSerialFrom.value)
			newElem.setAttribute "SERIALTO","" ' trim(document.formname.txtSerialTo.value)
			newElem.setAttribute "TAREWEIGHT", "U"
			newElem.setAttribute "IVALUE", trim(document.formname.txtValue.value)
			iQtyTotGross = 0
			iQtyTotTare = 0
			set TempRoot=SerialData.documentElement
			for each Node in TempRoot.childNodes
				iQtyTotGross = cdbl(iQtyTotGross) + cdbl(Node.attributes.getNamedItem("QTYREC").value)
				iQtyTotTare = cdbl(iQtyTotTare) + cdbl(Node.attributes.getNamedItem("TAREREC").value)
				newElem.appendChild Node
			next

			for i = 1 to iCnt
				set objD = eval("document.formname.txtGrossZ"&CStr(i))
				set objQ = eval("document.formname.txtTareZ"&CStr(i))

				set objSel = eval("document.formname.selSellZ"&CStr(i))
				set objWei = eval("document.formname.txtSellWeightZ"&CStr(i))
				set objPac = eval("document.formname.selPackZ"&CStr(i))
				set objFor = eval("document.formname.selFormZ"&CStr(i))
				set objPackNo = eval("document.formname.txtPackNoZ"&CStr(i))
				if document.formname.hrec.value="S"  then set objValue = eval("document.formname.txtValueZ"&CStr(i))
				if document.formname.hrec.value="LS"  then set objValue = eval("document.formname.txtValueZ"&CStr(i))
				if document.formname.hrec.value="L"  then set objValue = eval("document.formname.txtValueZ"&CStr(i))
				sAttributeID = document.formname.hAttributeList.value

				if sAttributeID<>"NULL" and sAttributeID<>"0" and not IsNull(sAttributeID) and sAttributeID<>"" then
					if document.formname.selAttribute.selectedIndex > 0 then
					    sAttributeID = document.formname.selAttribute(document.formname.selAttribute.selectedIndex).value
				    else
				        sAttributeID = ""
				    end if
				end if

				Set newElem1 = objTemp.createElement("LotSerialDetails")
				newElem1.setAttribute "LOTSERIAL", i
				newElem1.setAttribute "QTYREC", trim(objD.value)
				newElem1.setAttribute "TAREREC", trim(objQ.value)
				newElem1.setAttribute "SELLINGTYPE", trim(objSel.value)
				newElem1.setAttribute "WEIGHTSTYPE", trim(objWei.value)
				newElem1.setAttribute "PACKINGTYPE", trim(objPac.value)
				newElem1.setAttribute "LOT",sPubLot  '   trim(document.formname.txtLotNumber.value)
				newElem1.setAttribute "SELLINGFORM", trim(objFor.value)
				newElem1.setAttribute "PACKNUMBER", trim(objPackNo.value)

				if document.formname.hrec.value="S" then newElem1.setAttribute "IVALUE", trim(objValue.value)
				if document.formname.hrec.value="LS" then newElem1.setAttribute "IVALUE", trim(objValue.value)
				if document.formname.hrec.value="L" then newElem1.setAttribute "IVALUE", trim(objValue.value)
				newElem1.setAttribute "ATTRIBUTELIST",trim(sAttributeID)
				newElem.appendChild newElem1

				iQtyTotGross = cdbl(iQtyTotGross) + cdbl(objD.value)
				iQtyTotTare = cdbl(iQtyTotTare) + cdbl(objQ.value)
				'iTempTotValue = cdbl(iTempTotValue) + cdbl(objValue.value)

			'	if ucase(document.formname.hItemType.value)="FAB" then
			'		set FabricRoot = FabricData.documentElement
			'		for each FabricHeaderNode in FabricRoot.childNodes
			''			if cstr(FabricHeaderNode.attributes.getNamedItem("NO").value) = cstr(i) then
			'				set newElem2 = objTemp.createElement("FABRICDETAILS")
			'				newelem2.setAttribute "NO",i
			'				newelem2.setAttribute "QUANTITY",FabricHeaderNode.attributes.getNamedItem("QUANTITY").value
			'				newelem2.setAttribute "QUANTITYIN",FabricHeaderNode.attributes.getNamedItem("QUANTITYIN").value
			'				newelem2.setAttribute "CHECK",trim(FabricHeaderNode.Attributes.getNamedItem("CHECK").value)
			''				newelem2.setAttribute "ALTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("ALTGROSS").value)
			'				newelem2.setAttribute "ALTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("ALTNETT").value)
			'				newelem2.setAttribute "TOTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("TOTGROSS").value)
			'				newelem2.setAttribute "TOTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("TOTNETT").value)
'
'							for each FabricNode in FabricHeaderNode.childNodes
'								set NewElem3 = objTemp.createElement("DETAILS")
'								NewElem3.setAttribute "PIECENO", FabricNode.attributes.getNamedItem("PIECENO").value
'								NewElem3.setAttribute "QUANTITY", FabricNode.attributes.getNamedItem("QUANTITY").value
'								NewElem3.setAttribute "ALTGROSS",FabricNode.attributes.getNamedItem("ALTGROSS").value
'								NewElem3.setAttribute "ALTNETT",FabricNode.attributes.getNamedItem("ALTNETT").value
'
'								newElem2.appendChild NewElem3
'							next
'							newElem1.appendChild newElem2
'						end if
'					next
'				end if

			next


	'		if ucase(document.formname.hItemType.value)="FAB" then
	'			newElem.setAttribute "QTY", cdbl(iQtyTotGross)
	'		else
				newElem.setAttribute "QTY", cdbl(iQtyTotGross-iQtyTotTare)
	'		end if

			iCounter = cdbl(iCounter) + 1
			newElem.setAttribute "COUNTER", cdbl(iCounter)
			newElem.setAttribute "STAGE", document.formname.selStage.value
			newElem.setAttribute "ALTGROSS", document.formname.txtAltGross.value
			newElem.setAttribute "ALTNETT", document.formname.txtAltNett.value
			newElem.setAttribute "ALTUOM", document.formname.selAltUom.value
			newElem.setAttribute "IVALUE", document.formname.txtValue.value
			If document.formname.chkNew.checked Then
				newElem.setAttribute "AUTOGEN", "AUTO"
			Else
				newElem.setAttribute "AUTOGEN", ""
			End If

			childNod.appendChild newElem
		ClearAll()
		EnableDone()
		exit function
		end if
	next
end Function

Function CheckSubmit()
	AddSerial()

	Set Root = SerialData.documentElement

	UpdateValues()
	''Blocked By Ragav Feb 27,2012
	''begin
	'if cdbl(idQtyEntered.innerText) <> cdbl(idQty.innerText) then
	'	alert("Total Received Quantity should be equal to Quantity (" &idQty.innerText& ")")
	'	exit function
	'end if
	''end

	if CheckFabricData() then exit function

	if CheckFabricQtyBreakup() then exit function

	UpdateXml()

	RecalculateQty()
	RecalculateTotal()


	If CheckTotalValue() then Exit Function
	If avail = False Then Exit Function

	iQtyTotGross=0
	iQtyTotTare=0
	Alert ("Lot\Serial Details added successfully")
	window.close
	exit function
End Function


Function window_onunload()
	dim iCalQty,iInnerQty

'	if ucase(document.formname.hItemType.value)<>"FAB" then
		iCalQty = FormatNumber(trim(idQtyEntered.innerText),3)
		'iInnerQty = FormatNumber(trim(idQty.innerText),3)
'	else
'		iCalQty = FormatNumber(trim(idQtyEntered.innerText),3)
'		iInnerQty = FormatNumber(trim(idQty.innerText),3)
'
'	end if
iInnerQty = 0

iTotalValue = cdbl(iCalQty)* cdbl(document.formname.txtValue.value)
'alert(iTotalValue)
	if (iCalQty <> iInnerQty) then
		Set Root = objTemp.documentElement
		'alert(Root.xml)
		for each HeaderNode in Root.childNodes
			if StrComp(Trim(HeaderNode.NodeName),"STORAGE") = 0 then
				if HeaderNode.attributes.getNamedItem("STORE").value = iStore and HeaderNode.attributes.getNamedItem("BIN").value = iBin then
				    HeaderNode.setAttribute "QTY",iCalQty
			        HeaderNode.setAttribute "STORAGEVALUE",iTotalValue
					for each SubNode in HeaderNode.childNodes
					'	HeaderNode.Removechild SubNode
					next
				end if
			end if
		next
	end if

	set window.returnValue = objTemp.documentElement
'	alert(objTemp.xml)
	window.close()
end Function

Function ClearTable()
	dim i
	for i=3 to document.all.tblLot.rows.length - 1
		document.all.tblLot.deleteRow(3)
	next

end Function


Function setIndex(objSch,sTemp)
	for ic = 0 to objSch.length - 1
		if sTemp = objSch.options(ic).value then
			objSch.selectedIndex = ic
			exit function
		end if
	next
end Function

Function LotExist(sTemp)
	for ic = 0 to document.formname.sellot.length - 1
		if Trim(sTemp) = Trim(document.formname.sellot.options(ic).Text) then
			LotExist = True
			exit function
		end if
		LotExist = False
	next
end Function

Function PackNoLocalCheck1(obj,iNo)
	exit function
end Function

'Function to check whether the Packing Number exists in the XML / TextBoxes
Function PackNoLocalCheck(obj,iNo)
	dim sCheckValue,Root,HeaderNode,Node,SubNode
	dim sTemp,n,i,k,sCheck

	If iNo = "" Then Exit Function
	sCheckValue = obj.value

	sCheck = document.formname.hPackNoFlag.value
'	alert(sCheck)
'	alert(iNo)
	set objPac = eval("document.formname.selPack"&CStr(iNo))
	sValue = objPac.value
'	alert(sValue)

	'n= document.all.tblLot.rows.length
	sTemp = obj.name
	'i = replace(sTemp,"txtPackNo","")
	i = sTemp
	Set Root = objTemp.documentElement
	For Each HeaderNode In Root.childNodes
		for Each Node In HeaderNode.childNodes
			for Each SubNode in Node.childNodes
				if SubNode.nodeName="LotSerialDetails" then
					if not sCheck = "K" then
						if SubNode.Attributes.getNamedItem("PACKNUMBER").Value = sCheckValue then
							PackNoLocalCheck = true
							alert("Packing Number already entered previously.")
							obj.select
							exit function
						else
							PackNoLocalCheck = false
						end if
					elseif sCheck = "K" then
						if SubNode.Attributes.getNamedItem("PACKNUMBER").Value = sCheckValue and SubNode.Attributes.getNamedItem("PACKINGTYPE").Value = sValue then
							PackNoLocalCheck = true
							alert("Packing Number already Exists")
							obj.select
							exit function
						else
							PackNoLocalCheck = false
						end if
					end if
				end if
			next
		next
	next

	Set Root = SerialData.documentElement
	For Each SubNode In Root.childNodes
		if SubNode.nodeName="LotSerialDetails" then
			if not sCheck = "K" then
				if SubNode.Attributes.getNamedItem("PACKNUMBER").Value = sCheckValue then
					PackNoLocalCheck = true
					alert("Packing Number already entered previously.")
					obj.select
					exit function
				else
					PackNoLocalCheck = false
				end if
			elseif sCheck = "K" then
				if SubNode.Attributes.getNamedItem("PACKNUMBER").Value = sCheckValue and SubNode.Attributes.getNamedItem("PACKINGTYPE").Value = sValue then
					PackNoLocalCheck = true
					alert("Packing Number already Exists")
					obj.select
					exit function
				else
					PackNoLocalCheck = false
				end if
			end if
		end if
	next

	for k=No to j-1
		if k <> cint(i) then
			set objPackg = eval("document.formname.txtPackNo"&k)
			set objPackForm = eval("document.formname.selPack"&k)
			if objPackg.value <> "" then
				if not sCheck = "K" then
					if objPackg.value = sCheckValue then
						PackNoLocalCheck  = true
						alert("Packing Number already Exists..")
						obj.select
						exit function
					else
						PackNoLocalCheck = false
					end if
				elseif sCheck = "K" then
					if objPackg.value = sCheckValue and objPackForm.value = sValue then
						PackNoLocalCheck  = true
						alert("Packing Number already Exists..")
						obj.select
						exit function
					else
						PackNoLocalCheck = false
					end if
				end if
			end if
		end if
	next
End Function

Function PackNoCheck(obj,objPac)
	dim Root,HeaderNode,objhttp

	if document.formname.hPackNoFlag.value="" then
		PackNoCheck = false
		exit function
	end if

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	'alert(objPac.value)
	objhttp.Open "GET","XMLPackingNoCheck.asp?sOrgId="&sOrgID&"&sItemCode="&iItem&"&sPackNoFlag="&document.formname.hPackNoFlag.value&"&sPackNo="&obj.value&"&sPackCode="&objPac.value, false

	objhttp.send

	if objhttp.responseXML.xml <> "" then
		Data.loadXML objhttp.responseXML.xml
		Set Root = Data.documentElement
		if Root.Attributes.Item(0).nodeValue = "N" then
			PackNoCheck = false
		else
			PackNoCheck = true
		end if
	end if
end Function


Function SerialLocalCheck(iValue,iFromValue,iToValue)
	dim Root,Node

	Set Root = SerialData.documentElement

	if not Root.hasChildNodes then
		SerialLocalCheck = true
		exit function
	end if

	For each Node in Root.childNodes
		sLotserialNo = Node.Attributes.getNamedItem("LOTSERIAL").Value
		for iTemp = cint(iFromValue) to cint(iToValue)
			if cint(sLotserialNo) = cint(iTemp) then
				SerialLocalCheck = false
				exit function
			else
				SerialLocalCheck = true
			end if
		next
	next
end Function

Function AddSerial()
	dim i,iTempTotGross,iTempTotTare
	dim objQ,objD,objSerial,Root,objValue,sAttributeID
	avail = True
	if document.all.tblLot.rows.length = 2 then
		alert("No Details Entered")
		exit function
	end if

	document.formname.txtLotNumber.ReadOnly = true
	for i = 1 to iCnt
		set objD = eval("document.formname.txtGrossZ"&CStr(i))
		set objQ = eval("document.formname.txtTareZ"&CStr(i))

		set objSel = eval("document.formname.selSellZ"&CStr(i))
		set objWei = eval("document.formname.txtSellWeightZ"&CStr(i))
		set objPac = eval("document.formname.selPackZ"&CStr(i))
		'set objFor = eval("document.formname.selFormZ"&CStr(i))
		set objPackNo = eval("document.formname.txtPackNoZ"&CStr(i))
		'alert document.formname.hrec.value

		if document.formname.hrec.value="S" or document.formname.hrec.value="LS" or document.formname.hrec.value="L" then
			set objValue = eval("document.formname.txtValueZ"&CStr(i))
			'alert  eval("document.formname.txtIValue"&CStr(i)).value
		End If
		if trim(objD.value) = "" then
			alert("Enter Gross / Nett")
			objD.select()
			avail = False
			exit function
		elseif not checkNumbers(objD.value) then
			alert("Enter Numerals Only")
			objD.select()
			avail = False
			exit function
		elseif trim(objQ.value) = "" then
			alert("Enter Tare")
			objQ.select()
			avail = False
			exit function
		elseif not checkNumbers(objQ.value) then
			alert("Enter Numerals Only")
			objQ.select()
			avail = False
			exit function
		'elseif (objSel.length > 1 and objSel.selectedIndex = 0) then
		'	alert("Select Selling Type")
		'	objSel.focus()
		'	avail = False
		'	exit function
		elseif (objSel.selectedIndex > 0 and trim(objWei.value) = "") then
			alert("Enter Weight Of Selling Type")
			objWei.select()
			avail = False
			exit function
		elseif (objSel.selectedIndex > 0 and trim(objWei.value) <> "") and not checkNumbers(objWei.value) then
			alert("Enter Numerals Only")
			objWei.select()
			avail = False
			exit function
		elseif (objPac.length > 1 and objPac.selectedIndex = 0) then
			alert("Select Packing Type")
			objPac.focus()
			avail = False
			exit function
'		elseif (objFor.length > 1 and objFor.selectedIndex = 0) then
'			alert("Select Selling Form")
'			objFor.focus()
'			exit function
		elseif trim(objPackNo.value) = "" then
			alert("Enter Packing Number")
			objPackNo.select()
			avail = False
			exit function
		elseif PackNoCheck(objPackNo,objPac) then
			alert("Packing Number already Exists")
			objPackNo.select()
			avail = False
			exit function
		'elseif PackNoLocalCheck(objPackNo,i) then
		'	exit function
		else
			if document.formname.hrec.value="S" then
				if Trim(objValue.value) = "" then
					alert("Enter Value")
					objValue.focus()
					avail = False
					exit function
				end if
			end if
'		else
'			iTempTotGross = cdbl(iTempTotGross) + cdbl(objD.value)
'			iTempTotTare = cdbl(iTempTotTare) + cdbl(objQ.value)
		end if
	next

'	iQtyTotGross = cdbl(iQtyTotGross) + cdbl(iTempTotGross)
'	iQtyTotTare = cdbl(iQtyTotTare) + cdbl(iTempTotTare)

	RecalculateQty()
	'alert(iQtyTotGross &"***"&iQtyTotTare)
	idQtyEntered.innerText = FormatNumber(cdbl(iQtyTotGross) - cdbl(iQtyTotTare),3,,,0)

	'alert(idQtyEntered.innerText & ">" & idQty.innerText)
	''Blocked By Ragav Feb 27,2012
	''begin
	'if (cdbl(idQtyEntered.innerText) > cdbl(idQty.innerText)) then
	'	alert("Total Received Quantity should be Less than or equal to Quantity (" &idQty.innerText& ")")
	'	avail = False
'	'	iQtyTotGross = cdbl(iQtyTotGross) - cdbl(iTempTotGross)
'	'	iQtyTotTare = cdbl(iQtyTotTare) - cdbl(iTempTotTare)
	'	exit function
	'end if

	RecalculateTotal()

	iTotValue = FormatNumber(cdbl(iTotValue),4,,,0)
	iValue = FormatNumber(cdbl(iValue),4,,,0)


'	if cdbl(iTotValue) > cdbl(iValue) then
'		alert "Total value should be equal to (" & iValue & "). You have entered (" & cdbl(iTotValue) & ")."
'		avail = False
'		exit function
'	end if

'	if ucase(document.formname.hItemType.value)="FAB" then
'		if CheckFabricData then exit function
'	end if


	set Root=SerialData.documentElement
	for i = 1 to iCnt
		set objD = eval("document.formname.txtGrossZ"&CStr(i))
		set objQ = eval("document.formname.txtTareZ"&CStr(i))
		set objSel = eval("document.formname.selSellZ"&CStr(i))
		set objWei = eval("document.formname.txtSellWeightZ"&CStr(i))
		set objPac = eval("document.formname.selPackZ"&CStr(i))
		set objFor = eval("document.formname.selFormZ"&CStr(i))
		set objPackNo = eval("document.formname.txtPackNoZ"&CStr(i))

		if document.formname.hrec.value="S" then set objValue = eval("document.formname.txtValueZ"&CStr(i))
		if document.formname.hrec.value="LS" then set objValue = eval("document.formname.txtValueZ"&CStr(i))
		if document.formname.hrec.value="L" then set objValue = eval("document.formname.txtValueZ"&CStr(i))
		sAttributeID = document.formname.hAttributeList.value

		if sAttributeID<>"NULL" and sAttributeID<>"0" and not IsNull(sAttributeID) and sAttributeID<>"" then
			if document.formname.selAttribute.selectedIndex > 0 then
		        sAttributeID = document.formname.selAttribute(document.formname.selAttribute.selectedIndex).value
		    else
		        sAttributeID = ""
		    end if
		end if


		Set newElem = SerialData.createElement("LotSerialDetails")
'		newElem.setAttribute "SERIALNO", i
'		newElem.setAttribute "LOTSERIAL", trim(objSerial.value)
		newElem.setAttribute "LOTSERIAL", i
		newElem.setAttribute "QTYREC", trim(objD.value)
		newElem.setAttribute "TAREREC", trim(objQ.value)
		newElem.setAttribute "SELLINGTYPE", trim(objSel.value)
		newElem.setAttribute "WEIGHTSTYPE", trim(objWei.value)
		newElem.setAttribute "PACKINGTYPE", trim(objPac.value)
		newElem.setAttribute "LOT", trim(document.formname.txtLotNumber.value)
		newElem.setAttribute "SELLINGFORM", trim(objFor.value)
		newElem.setAttribute "PACKNUMBER", trim(objPackNo.value)

		if document.formname.hrec.value="S" then newElem.setAttribute "IVALUE", trim(objValue.value)
		if document.formname.hrec.value="LS" then newElem.setAttribute "IVALUE", trim(objValue.value)
		if document.formname.hrec.value="L" then newElem.setAttribute "IVALUE", trim(objValue.value)
		newElem.setAttribute "ATTRIBUTELIST",trim(sAttributeID)

		Root.appendChild newElem

	'	if ucase(document.formname.hItemType.value)="FAB" then
	'		set FabricRoot = FabricData.documentElement
	'		for each FabricHeaderNode in FabricRoot.childNodes
	'			if cstr(FabricHeaderNode.attributes.getNamedItem("NO").value) = cstr(i) then
	'				set newElem1 = objTemp.createElement("FABRICDETAILS")
	'				newelem1.setAttribute "NO",i
	'				newelem1.setAttribute "QUANTITY",FabricHeaderNode.attributes.getNamedItem("QUANTITY").value
	'				NewElem1.setAttribute "QUANTITYIN", FabricHeaderNode.attributes.getNamedItem("QUANTITYIN").value
	'				NewElem1.setAttribute "CHECK",trim(FabricHeaderNode.Attributes.getNamedItem("CHECK").value)
	'				NewElem1.setAttribute "ALTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("ALTGROSS").value)
	'				NewElem1.setAttribute "ALTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("ALTNETT").value)
	'				NewElem1.setAttribute "TOTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("TOTGROSS").value)
	'				NewElem1.setAttribute "TOTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("TOTNETT").value)
'
'					for each FabricNode in FabricHeaderNode.childNodes
'						set NewElem2 = objTemp.createElement("DETAILS")
'						NewElem2.setAttribute "PIECENO", FabricNode.attributes.getNamedItem("PIECENO").value
'						NewElem2.setAttribute "QUANTITY", FabricNode.attributes.getNamedItem("QUANTITY").value
'						NewElem2.setAttribute "ALTGROSS",FabricNode.attributes.getNamedItem("ALTGROSS").value
'						NewElem2.setAttribute "ALTNETT",FabricNode.attributes.getNamedItem("ALTNETT").value
'
'						newElem1.appendChild NewElem2
'					next
'					newElem.appendChild newElem1
'				end if
'			next
'		end if

	next

	EnableDone()

End Function


Function CheckExists()
	dim FabricRoot,FabricHeaderNode,FabricNode, iTempTotValue

	iTempTotValue = 0
	If iTempTotGross = Empty Then iTempTotGross = 0
	If iTempTotTare = Empty Then iTempTotTare = 0
	if iQtyTotGross = Empty Then iQtyTotGross = 0
	If iQtyTotTare = Empty Then iQtyTotTare = 0
	Set Root = objTemp.documentElement
	For Each childNod In Root.childNodes
		if childNod.Attributes.Item(0).nodeValue = iStore and childNod.Attributes.Item(1).nodeValue = iBin then
			childNod.setAttribute "QTY", cdbl(iTempTotGross+iTempTotTare)
			Set newElem = objTemp.createElement("LotSerial")
			newElem.setAttribute "QTYIN", "N"
			newElem.setAttribute "TARE", trim(document.formname.txtTare.value)
			newElem.setAttribute "LOT", trim(document.formname.txtLotNumber.value)
			newElem.setAttribute "SERIALFROM","" ' trim(document.formname.txtSerialFrom.value)
			newElem.setAttribute "SERIALTO","" ' trim(document.formname.txtSerialTo.value)
			newElem.setAttribute "TAREWEIGHT", "U"

			RecalculateQty()

		'	if ucase(document.formname.hItemType.value)="FAB" then
		'   	newElem.setAttribute "QTY", cdbl(iQtyTotGross)
		'	else
				newElem.setAttribute "QTY", cdbl(iQtyTotGross-iQtyTotTare)
		'	end if

			iCounter = cdbl(iCounter) + 1
			newElem.setAttribute "COUNTER", cdbl(iCounter)
			newElem.setAttribute "STAGE", document.formname.selStage.value
			newElem.setAttribute "ALTGROSS", document.formname.txtAltGross.value
			newElem.setAttribute "ALTNETT", document.formname.txtAltNett.value
			newElem.setAttribute "ALTUOM", document.formname.selAltUom.value
			 newElem.setAttribute "IVALUE", cdbl(document.formname.txtValue.value)

			set TempRoot=SerialData.documentElement
			for each Node in TempRoot.childNodes
				newElem.appendChild Node
				if sType = "S" or sType="LS" then
					iTempTotValue = cdbl(iTempTotValue) + cdbl(Node.attributes.getNamedItem("IVALUE").value)
					newElem.setAttribute "IVALUE", iTempTotValue
				end if
			next

			childNod.appendChild newElem
		exit function
		end if
	next
End Function

Function DisplayQty()
	dim Root, Node, sExp, i, iQty, iTare

	Set Root = objTemp.documentElement
	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial/LotSerialDetails"
	Set Node = Root.Selectnodes(sExp)

	For i = 0 to Node.Length - 1

		iQty = cdbl(iQty) + cdbl(Node.Item(i).Attributes.getNamedItem("QTYREC").Value)
		iTare = cdbl(iTare) + cdbl(Node.Item(i).Attributes.getNamedItem("TAREREC").Value)
	next

	iQtyTotGross = iQty
	iQtyTotTare = iTare

	DisplayQty = cdbl(iQty) - cdbl(iTare)

End Function

Function DisplayFabricQty()
	dim Root, Node, sExp, i, iQty, iTare

	Set Root = objTemp.documentElement

	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial/LotSerialDetails"
	Set Node = Root.Selectnodes(sExp)

	For i = 0 to Node.Length - 1
		iQty = cdbl(iQty) + cdbl(Node.Item(i).Attributes.getNamedItem("QTYREC").Value)
	next

	iQtyTotGross = iQty
	iQtyTotTare = cdbl(0)

	DisplayFabricQty = cdbl(iQty)

End Function

Function RecalculateQty()
	dim i, iQty, iTare, objQty, objTare

'	if ucase(document.formname.hItemType.value)="FAB" then
'		CalculateFabricQty()
'		exit Function
'	end if

	for i = 1 to j - 1
		set objQty = eval("document.formname.txtGrossZ"&CStr(i))
		set objTare = eval("document.formname.txtTareZ"&CStr(i))

		if objQty.value = "" then
			iEQty = "0"
		else
			iEQty = objQty.value
		end if

		if objTare.value = "" then
			iETare = "0"
		else
			iETare = objTare.value
		end if

		iQty = cdbl(iQty) + cdbl(iEQty)
		if cdbl(iEQty) <> 0 then iTare = cdbl(iTare) + cdbl(iETare)
	next

	iQtyTotGross = FormatNumber(cdbl(iQty),3,,,0)

	iQtyTotTare = FormatNumber(cdbl(iTare),3,,,0)
	'alert("Recal="&iQtyTotGross &"and"&iQtyTotTare)

	idQtyEntered.innerText = FormatNumber(cdbl(iQty) - cdbl(iTare),3,,,0)

End Function

Function CalculateFabricQty()
	dim i, iQty, objQty, iEQty

	for i = 1 to j-1
		set objQty = eval("document.formname.txtGrossZ"&CStr(i))

		if objQty.value = "" then
			iEQty = "0"
		else
			iEQty = objQty.value
		end if

		iQty = cdbl(iQty) + cdbl(iEQty)
	next

	iQtyTotGross = iQty
	iQtyTotTare = cdbl(0)

	idQtyEntered.innerText = FormatNumber(cdbl(iQty),3,,,0)

End Function

Function CalculateLotQty()
	dim Root, Node, sExp, i, iQty, iTare

	Set Root = objTemp.documentElement

	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial/LotSerialDetails"
	Set Node = Root.Selectnodes(sExp)

	For i = 0 to Node.Length - 1
		iQty = cdbl(iQty) + cdbl(Node.Item(i).Attributes.getNamedItem("QTYREC").Value)
		iTare = cdbl(iTare) + cdbl(Node.Item(i).Attributes.getNamedItem("TAREREC").Value)
	next

	CalculateLotQty = cdbl(iQty) - cdbl(iTare)

End Function

Function UpdateValues()
	dim Root, i,Node,sExp
	dim TotGross, TotTare,iLotCtr,sAttributeID
	DIM StorageNode, LotSerialNode, cntr
	Set Root = objTemp.documentElement
	sAttributeID = document.formname.hAttributeList.value

	if sAttributeID<>"NULL" and sAttributeID<>"0" and not IsNull(sAttributeID) and sAttributeID<>"" then
        if document.formname.selAttribute.selectedIndex > 0 then
	        sAttributeID = document.formname.selAttribute(document.formname.selAttribute.selectedIndex).value
        else
            sAttributeID = ""
        end if
	end if

	'alert Root.XML
	sExp ="//STORAGE/LotSerial/LotSerialDetails"
	Set Node = Root.Selectnodes(sExp)
	For i = 1 to Node.Length
		Node.item(i-1).Attributes.getNamedItem("PACKNUMBER").value = eval("document.formname.txtPackNoZ"&CStr(i)).value
		Node.item(i-1).Attributes.getNamedItem("PACKINGTYPE").value = eval("document.formname.selpackZ"&CStr(i)).value
		Node.item(i-1).Attributes.getNamedItem("WEIGHTSTYPE").value = eval("document.formname.txtSellWeightZ"&CStr(i)).value
		Node.item(i-1).Attributes.getNamedItem("QTYREC").value =  eval("document.formname.txtGrossZ"&CStr(i)).value
		Node.item(i-1).Attributes.getNamedItem("TAREREC").value = eval("document.formname.txtTareZ"&CStr(i)).value
		Node.item(i-1).Attributes.getNamedItem("SELLINGFORM").value  = eval("document.formname.selFormZ"&CStr(i)).value
		Node.item(i-1).Attributes.getNamedItem("ATTRIBUTELIST").Value = sAttributeID
	Next
	'alert Root.XML

	sExp ="//STORAGE/LotSerial"
	Set Node = Root.Selectnodes(sExp)
	For i = 1 to Node.Length
			Node.item(i-1).Attributes.getNamedItem("IVALUE").value = eval("document.formname.txtIvalueZ"&CStr(i)).value
	Next

	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial"
	Set LotSerialNode = Root.Selectnodes(sExp)
	For iLotCtr = 0 to LotSerialNode.Length - 1
		cntr = LotSerialNode.Item(iLotCtr).Attributes.getNamedItem("COUNTER").Value

		sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial [ @COUNTER ='"&cntr&"']/LotSerialDetails"

		Set StorageNode = Root.Selectnodes(sExp)
		For iCounter = 0 to StorageNode.Length - 1
				'alert cDbl(StorageNode.item(iCounter).Attributes.getNamedItem("QTYREC").value)
				TotGross = TotGross + cDbl(StorageNode.item(iCounter).Attributes.getNamedItem("QTYREC").value)
				TotTare = TotTare + cDbl(StorageNode.item(iCounter).Attributes.getNamedItem("TAREREC").value)
		Next
		LotSerialNode.Item(iLotCtr).Attributes.getNamedItem("TARE").Value = TotTare
		LotSerialNode.Item(iLotCtr).Attributes.getNamedItem("QTY").Value = TotGross
'		LotSerialNode.Item(iLotCtr).Attributes.getNamedItem("ALTGROSS").Value = TotGross
'		LotSerialNode.Item(iLotCtr).Attributes.getNamedItem("ALTGROSS").Value = TotGross
		'alert "ok"
	next
'alert "ook"

End Function

Function UpdateLot()
	dim Root, i,Node,sExp, OldLot
	Set Root = objTemp.documentElement
	'alert Root.XML
	sExp ="//STORAGE/LotSerial"
	Set Node = Root.Selectnodes(sExp)
	document.formname.sellot.options.Length = 0

	document.formname.sellot.length = document.formname.sellot.length+1
	document.formname.sellot.options(document.formname.sellot.length-1).text = "select"
	document.formname.sellot.options(document.formname.sellot.length-1).Value = "select"

	document.formname.sellot.length = document.formname.sellot.length+1
	document.formname.sellot.options(document.formname.sellot.length-1).text = "New Lot No."
	document.formname.sellot.options(document.formname.sellot.length-1).Value = "N"

	For i = 1 to Node.Length
'		Node.item(i-1).Attributes.getNamedItem("PACKNUMBER").value = eval("document.formname.txtPackNoZ"&CStr(i)).value
		'alert Node.item(i-1).Attributes.getNamedItem("LOT").value
		If OldLot <>  Node.item(i-1).Attributes.getNamedItem("LOT").value Then
		document.formname.sellot.length = document.formname.sellot.length + 1
		document.formname.sellot.options(document.formname.sellot.length-1).text = Node.item(i-1).Attributes.getNamedItem("LOT").value
		document.formname.sellot.options(document.formname.sellot.length-1).value = document.formname.sellot.length
		End If
		OldLot =  Node.item(i-1).Attributes.getNamedItem("LOT").value
	Next
End Function

Function UpdateXml()
	dim i, iQty, iTare, objQty, objTare, objLotSerial, k, objValue
	dim sExp, Root, Node, sAttribute, HeaderNode, SubNode, ChildNode

	Set Root = objTemp.documentElement
	'alert Root.XML
	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial"
'	alert sExp
	Set Node = Root.Selectnodes(sExp)

	For i = 1 to Node.Length
		Node.item(i-1).Attributes.getNamedItem("ALTGROSS").value = document.formname.txtAltGross.value
		Node.item(i-1).Attributes.getNamedItem("ALTNETT").value = document.formname.txtAltNett.value
		Node.item(i-1).Attributes.getNamedItem("ALTUOM").value = document.formname.selAltUom.value
		if document.formname.hrec.value<>"S" then Node.item(i-1).Attributes.getNamedItem("IVALUE").value = eval("document.formname.txtIValueZ"&CStr(i)).value
	next

	sExp ="//STORAGE [ @STORE = "&iStore&" and @BIN = '"&iBin&"']/LotSerial/LotSerialDetails"
	Set Node = Root.Selectnodes(sExp)

	For i = 1 to Node.Length
		set objQty = eval("document.formname.txtGrossZ"&CStr(i))
		set objTare = eval("document.formname.txtTareZ"&CStr(i))
		'set objLotSerial = eval("document.formname.txtSerial"&CStr(i))
		if document.formname.hrec.value="S" then set objValue = eval("document.formname.txtValueZ"&CStr(i))
		if document.formname.hrec.value="LS" then set objValue = eval("document.formname.txtValueZ"&CStr(i))
		if document.formname.hrec.value="L" then set objValue = eval("document.formname.txtValueZ"&CStr(i))

		if cdbl(objQty.value) = 0 then
			k = 0
			if document.formname.hrec.value="L" then
				for each HeaderNode in Root.childNodes
					if HeaderNode.attributes.getNamedItem("STORE").value = iStore and HeaderNode.attributes.getNamedItem("BIN").value = iBin then
						for each SubNode in HeaderNode.childNodes
							k = k + 1
							if k = i then
								HeaderNode.Removechild SubNode
								exit For
							end if
						next
					end if
				next
			else
				for each HeaderNode in Root.childNodes
					if HeaderNode.attributes.getNamedItem("STORE").value = iStore and HeaderNode.attributes.getNamedItem("BIN").value = iBin then
						for each SubNode in HeaderNode.childNodes
							for each ChildNode in SubNode.childNodes
								k = k + 1
								if k = i then
									SubNode.Removechild ChildNode
									exit For
								end if
							next
						next
					end if
				next

			end if
		else
			Node.item(i-1).Attributes.getNamedItem("LOTSERIAL").value = i
			Node.item(i-1).Attributes.getNamedItem("QTYREC").value = cdbl(objQty.value)
			Node.item(i-1).Attributes.getNamedItem("TAREREC").value = cdbl(objTare.value)


		'	if ucase(document.formname.hItemType.value)="FAB" then

				'Removing Childs
		'		for each SubNode in Node.item(i-1).childNodes
		'			Node.item(i-1).removechild SubNode
		'		next
'
'				'Adding Childs
'				set FabricRoot = FabricData.documentElement
'				for each FabricHeaderNode in FabricRoot.childNodes
'					if cstr(FabricHeaderNode.attributes.getNamedItem("NO").value) = cstr(i) then
'						set newElem2 = objTemp.createElement("FABRICDETAILS")
'						newelem2.setAttribute "NO",i
'						newelem2.setAttribute "QUANTITY",FabricHeaderNode.attributes.getNamedItem("QUANTITY").value
'						newelem2.setAttribute "QUANTITYIN",FabricHeaderNode.attributes.getNamedItem("QUANTITYIN").value
'						newelem2.setAttribute "CHECK",trim(FabricHeaderNode.Attributes.getNamedItem("CHECK").value)
'						newelem2.setAttribute "ALTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("ALTGROSS").value)
'						newelem2.setAttribute "ALTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("ALTNETT").value)
'						NewElem2.setAttribute "TOTGROSS",trim(FabricHeaderNode.Attributes.getNamedItem("TOTGROSS").value)
'						NewElem2.setAttribute "TOTNETT",trim(FabricHeaderNode.Attributes.getNamedItem("TOTNETT").value)
'
'						for each FabricNode in FabricHeaderNode.childNodes
'							set NewElem3 = objTemp.createElement("DETAILS")
'							NewElem3.setAttribute "PIECENO", FabricNode.attributes.getNamedItem("PIECENO").value
'							NewElem3.setAttribute "QUANTITY", FabricNode.attributes.getNamedItem("QUANTITY").value
'							NewElem3.setAttribute "ALTGROSS",FabricNode.attributes.getNamedItem("ALTGROSS").value
'							NewElem3.setAttribute "ALTNETT",FabricNode.attributes.getNamedItem("ALTNETT").value
'
'							newElem2.appendChild NewElem3
'						next
'						Node.item(i-1).appendChild newElem2
'					end if
'				next
'			end if


		end if
	next
	'alert Root.XML

End Function

Function EnableDone()

'	if cdbl(idQtyEntered.innerText) = cdbl(idQty.innerText) then
'		document.formname.B1.disabled = false
'	else
'		'document.formname.B1.disabled = true
'	end if

End Function

Function CalculatePackNo(iNo)
	dim i,l,PackNo

	l = len(iNo) + 1
	for i = l to 5
		PackNo = PackNo & "0"
	next

	CalculatePackNo = document.formname.txtPrefix.value & PackNo & iNo & document.formname.txtSuffix.value


End Function

Function ShowPopup(Obj)
	Dim No, n, objQty,sTempValue,iCnt,sArrValue

	n = instr(1,Obj.name,"r")
	No = mid(Obj.name,n+1,len(Obj.name)-n+1)

	set ObjQty = eval("document.formname.txtGross"&CStr(No))

	if trim(ObjQty.value) = "" then
		alert "Enter Quantity"
		ObjQty.focus
		exit function
	elseif ObjQty.value = 0 then
		alert "Quantity should be greater than zero"
		ObjQty.select
		exit function
	end if

		sArrValue = split(document.formname.sTemp.value,"``")
		for iCnt =0 to 9
			sTempValue = sTempValue & sArrValue(iCnt)&":"
		next

	sTemp =  sTempValue & ObjQty.value & ":" & No &":"&document.formname.selAltUom.value&":"&document.formname.txtAltGross.value&":"&document.formname.txtAltNett.value
	'alert sTemp
	set OutValue= showModalDialog("stockLotSerFabPop.asp?sTemp=" & sTemp,FabricData,"dialogHeight:460px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")

End Function

Function CheckFabricData()
	dim i,objBtn,objGross

	i = 0
	set FabricRoot = FabricData.documentElement

	for each FabricHeaderNode in FabricRoot.childNodes
		i = i + 1
		set ObjBtn = eval("document.formname.BtnEnterZ"&CStr(i))
		set ObjGross = eval("document.formname.txtGrossZ"&CStr(i))

		if not FabricHeaderNode.hasChildNodes then
			alert "Enter Quantity Details"
			'objBtn.focus
			CheckFabricData = True
			ShowPopup(ObjBtn)
			exit function
		else
			if FabricHeaderNode.attributes.getNamedItem("QUANTITY").value <> objGross.value then
				alert "Quantity mismatch"
				'objBtn.focus
				CheckFabricData = True
				ShowPopup(ObjBtn)
				exit function
			end if
		end if
	next
	CheckFabricData = False
End Function

Function CheckFabricQtyBreakup()
	dim iTotGross, iTotNett,sAltUom

'	if ucase(document.formname.hItemType.value)<>"FAB" then
'		CheckFabricQtyBreakup = False
'		exit Function
'	end if

	iTotGross = 0
	iTotNett = 0
	sAltUom = document.formname.selAltUom.value

	set FabricRoot = FabricData.documentElement
	for each FabricHeaderNode in FabricRoot.childNodes
		iTotGross = cdbl(iTotGross) + cdbl(FabricHeaderNode.attributes.getNamedItem("TOTGROSS").value)
		iTotNett = cdbl(iTotNett) + cdbl(FabricHeaderNode.attributes.getNamedItem("TOTNETT").value)
	next
	if document.formname.txtAltGross.value = "" then document.formname.txtAltGross.value = 0
	if document.formname.txtAltNett.value = "" then document.formname.txtAltNett.value = 0
	if (cdbl(iTotGross) <> cdbl(document.formname.txtAltGross.value)) or (cdbl(iTotNett) <> cdbl(document.formname.txtAltNett.value)) then
		alert "Quantity mismatch in Alternate UoM " & vbcrlf & "Quantity in Gross : " & document.formname.txtAltGross.value & " " & sAltUom & vbcrlf & "Quantity in Nett : " & document.formname.txtAltNett.value & " " & sAltUom & vbcrlf & "Entered Gross : " & iTotGross & " " & sAltUom & vbcrlf & "Entered Nett : " & iTotNett & " " & sAltUom
		CheckFabricQtyBreakup = True
	else
		CheckFabricQtyBreakup = False
	end if

End Function

Function CheckTotalValue()
	Dim iNo, iVal,iCtr
	Dim Root,Node
	CheckTotalValue = False
	iTotValue = 0
	if Trim(sType) = "S" then
		iCtr = j - 1
	else
		iCtr = iFixedNo
	end if

'	for i = 1 to iCtr
'		set objValue = eval("document.formname.txtIValue"&CStr(i))
'		if Trim(objValue.value) <> "" then iTotValue = cdbl(iTotValue) + cdbl(objValue.value)
'	next
	set Root = objTemp.documentElement

	sExp ="//STORAGE/LotSerial"
	Set Node = Root.Selectnodes(sExp)
	For i = 1 to Node.Length
		iTotValue = iTotValue +	CDBL(Node.item(i-1).Attributes.getNamedItem("IVALUE").value)
	Next

'	alert iTotValue
'	alert iValue
	if cdbl(iTotValue) < cdbl(iValue) then
		alert "Total value should be equal to (" & iValue & "). You have Entered (" & cdbl(iTotValue) & ")."
		document.formname.txtValue.select
		avail = False
		CheckTotalValue = True
	end if
	iTotValue = FormatNumber(cdbl(iTotValue),4,,,0)

End Function

Function RecalculateTotal()
	dim i, objValue,iCtr
	dim Root, Node
	iTotValue = 0

	if Trim(sType) = "S" then
		iCtr = j - 1
	else
		iCtr = iFixedNo
	end if

	set Root = objTemp.documentElement

	sExp ="//STORAGE/LotSerial"
	Set Node = Root.Selectnodes(sExp)
	For i = 1 to Node.Length
		iTotValue = iTotValue +	CDBL(Node.item(i-1).Attributes.getNamedItem("IVALUE").value)
	Next


'	for i = 1 to iCtr
'		set objValue = eval("document.formname.txtIValue"&CStr(i))
'
'		if Trim(objValue.value) <> "" then iTotValue = cdbl(iTotValue) + cdbl(objValue.value)
'	next
End Function
Function DispLot()
	If document.formname.sellot.selectedIndex > 1 Then
		document.formname.txtLotNumber.value = document.formname.sellot.options(document.formname.sellot.selectedIndex).text
		'document.formname.txtLotNumber.disabled = True
		document.formname.txtLotNumber.readOnly = True
	Else
		document.formname.txtLotNumber.value = ""
		document.formname.txtLotNumber.readOnly = False
	End If
End Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="LoadDetails('<%=trim(sTemp)%>');init();">
<%	if sPackNoFlag = "" then %>
<SCRIPT LANGUAGE=vbscript>
	alert("Number Series for Packing Number has not been defined.")
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
	set window.returnValue = objTemp.documentElement
	window.close()
</SCRIPT>
<%	end if
	if bPackFlag then
%>
<!--SCRIPT LANGUAGE=vbscript>
	alert("Packing Type or Selling Form has not been defined.")
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
	set window.returnValue = objTemp.documentElement
	window.close()
</SCRIPT-->
<%	end if %>

<form method="POST" name="formname" action="">
<input type="hidden" name="hRec" value="<%=sType%>">
<input type="hidden" name="sTemp" value="<%=sTemp%>">
<input type="hidden" name="hItemType" value="<%=sItmType%>">
<input type="hidden" name="hPackNoFlag" value="<%=sPackNoFlag%>">
<input type="hidden" name="hDate" value="<%=StDate%>">
<input type="hidden" name="hAttributeList" value="<%=sAttributeID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot / Serial Details
		</td>
    </tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
											<td class="FieldCell"></td>
											<td class="FieldCellSub"></td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell" colspan="2"></td>
                                        </tr>

<!--Table 1-->
										<tr>
											<td align="left" class="FieldCell" width="60%">

												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class='GroupTitleLeft' width="10">&nbsp;</td>
																	<td class='GroupTitle' width="50"><p align="center">Details</td>
																	<td class='GroupTitleRight'><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																    <td class="FieldCell"></td>
																    <td class="FieldCell">Item</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idItemName"><%=sItemName%>&nbsp;</span>
																    </td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Store -- Bin</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idStoreName"><%=sStoName%>&nbsp;</span>
																    </td>
																</tr>
																<!--<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Nett Quantity&nbsp;</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idQty"><%=iQty%></span> -
																        <span class="DataOnly" ><%=sStoresUom%>&nbsp;</span>
																    &nbsp;Quantity Entered&nbsp;
																		<span class="DataOnly" id="idQtyEntered">0</span> -
																        <span class="DataOnly"><%=sStoresUom%>&nbsp;</span>
																    </td>
																</tr>-->
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity Entered&nbsp;</td>
																    <td class="FieldCellSub">
																    	<span class="DataOnly" id="idQtyEntered">0</span> -
																        <span class="DataOnly"><%=sStoresUom%>&nbsp;</span>
																    </td>
																</tr>
																
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>

															</table>
											            </td>
													</tr>
												</table>
											</td>

											<td align="left" class="FieldCell" width="40%">
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class='GroupTitleLeft' width="10">&nbsp;</td>
																	<td class='GroupTitle' width="90"><p align="center">Alternate UoM</td>
																	<td class='GroupTitleRight'><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																    <td class="FieldCell"></td>
																    <td class="FieldCell">Quantity in Gross</td>
																    <td class="FieldCellSub">
																        <input type="text" name="txtAltGross" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																		<select size="1" name="selAltUom" class="FormElem" onChange="idAltUom.innerText=document.formname.selAltUom.options(document.formname.selAltUom.selectedIndex).text">
																			<option value="select">Select</option>
																			<%
																				populateUoMList sStoresUom
																			%>
																		</select>
																    </td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity in Nett</td>
																    <td class="FieldCellSub">
																        <input type="text" name="txtAltNett" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																        <span class="DataOnly" id="idAltUom"></span>
																    </td>
																</tr>
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
															</table>
											            </td>
													</tr>
												</table>
											</td>
										</tr>
<!--Table 1-->

<!--Table 2-->
										<tr>
											<td align="center" class="FieldCell" width="100%" colspan="2">
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class='GroupTitleLeft' width="10">&nbsp;</td>
																	<td class='GroupTitle' width="100"><p align="center">Quantity Details</td>
																	<td class='GroupTitleRight'><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td align="center" colspan="5" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Receipt Numbering&nbsp;</td>
																    <td class="FieldCell">
																    <%	if sType = "S" then %>
																        <input type="text" name="txtRcptNumbering" READONLY size="15" value="SERIAL" class="FormElemRead">
																    <%	Elseif sType = "LS" then %>
																        <input type="text" name="txtRcptNumbering" READONLY size="15" value="LOT\SERIAL" class="FormElemRead">
																    <%	Elseif sType = "L" then %>
																        <input type="text" name="txtRcptNumbering" READONLY size="15" value="LOT" class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtRcptNumbering" size="15" class="FormElem">
																    <%	end if %>
																    </td>
																    <td class="FieldCellSub">Lot Number</td>
																    <td class="FieldCellSub">

																		<% If sType = "S" Then %>
																			<select size="1" name="sellot" class="FormElem" OnChange="" disabled >
																			<option value="select">select</option>
																			<option value="N">New Lot No.</option>
																		</select>
																		<input type="text" name="txtLotNumber" size="10" value="" class="FormElem" disabled>
																		<% Else %>
																    		<select size="1" name="sellot" class="FormElem" OnChange="DispLot()">
																				<option value="select">select</option>
																				<!--<option value="N">New Lot No.</option>-->
																				<option value="N">New Lot No.</option>
																			</select>
																			<input type="text" name="txtLotNumber" size="10" value="" class="FormElem">
																		<% End If %>

																	<!--<input type="checkbox" name="chkNew" value="1" >&nbsp;Auto Generate&nbsp;-->
																		<input type="checkbox" name="chkNew" value="1" checked>&nbsp;Auto Generate&nbsp;
																		<%	if sType = "S" then %>
																			Serial
																		<%	Elseif sType = "LS" then %>
																		    Lot & Serial
																		<%	Elseif sType = "L" then %>
																		    Lot
																		<%	else %>

																		<%	end if %>

																    </td>
																</tr>


																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity Gross&nbsp;&nbsp;&nbsp;</td>
																    <td class="FieldCell">
																        <!--input type="radio" value="G" name="radQtyIn" class="FormElem">  Gross&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																        <input type="radio" value="N" name="radQtyIn" class="FormElem">  Nett &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-->
																    <%	if ucase(sItmType) = "FAB" then %>
																        <input type="text" name="txtQtyIn" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																	<%	else %>
																		<input type="text" name="txtQtyIn" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																    <%	end if	%>

																    </td>
																    <td class="FieldCellSub">Tare weight&nbsp;</td>
																    <td class="FieldCellSub">
																		<input type="text" name="txtTare" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem"> <%=sStoresUom%>
																		&nbsp; Rate / <%=sStoresUom%> :&nbsp;&nbsp;
																		<input type="text" name="txtValue" size="10" value="" class="FormElem">
																    </td>
																</tr>

																<tr>
																	<td class="FieldCell"></td>
														            <td class="FieldCell">Selling Form</td>
														            <td class="FieldCell">
																		<!--input type="radio" value="I" name="radSellType" class="FormElem" onclick="DisableselSellType(this)">  Individual
																		<input type="radio" value="U" name="radSellType" class="FormElem" onclick="DisableselSellType(this)">  Uniform&nbsp;&nbsp;-->
																		<select size="1" name="selSellType" class="FormElem">
																			<option value="select">-N/A-</option>
																		</select>
																	</td>
																	<td class="FieldCellSub">W/t. Per Form</td>
																		<!--input type="radio" value="I" name="radWeight" class="FormElem" onclick="DisableTxtWt(this)">  Individual
																		<input type="radio" value="U" name="radWeight" class="FormElem" onclick="DisableTxtWt(this)">  Uniform&nbsp;&nbsp;-->
																	<td class="FieldCellSub">
																		<input type="text" name="txtWeight" size="11" class="FormElem" value="0">
																		&nbsp; No of Packs : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																		<input type="text" name="txtNoOfPacks" size="11" class="FormElem">
														            </td>
														        </tr>
																<tr>
																	<td class="FieldCell"></td>
														            <td class="FieldCell">Packing Type</td>
														            <td class="FieldCell">
																		<!--input type="radio" value="I" name="radPackType" class="FormElem" onclick="DisableselPack(this)">  Individual
																		<input type="radio" value="U" name="radPackType" class="FormElem" onclick="DisableselPack(this)">  Uniform&nbsp;&nbsp;-->
																		<select size="1" name="selPackType" class="FormElem">
																			<option value="select">Select</option>
																		</select>
																	</td>
														            <td class="FieldCellSub">W/t. Per Pack</td>
														            <td class="FieldCellSub">
																		<input type="text" name="txtWeightPerPack" size="11" class="FormElem">
																		&nbsp; Tare w/t per pack:&nbsp;
																		<input type="text" name="txtPackTare" size="11" class="FormElem">
																	</td>

																</tr>
																<tr>
																	<td class="FieldCell"></td>
																	<td class="FieldCell">Packing Form &nbsp	</td>
																	<td class="FieldCell">
																		<select size="1" name="selForm" class="FormElem">
																			<option value="select">-N/A-</option>
																		</select>
																	</td>
																	<td class="FieldCellSub">Packing Quality </td>
																	<td class="FieldCellSub">
																		<select size="1" name="selStage" class="FormElem">
																		<option value="select">-N/A-</option>
																		<%
																			PopulateStage sItmType
																		%>
																		</select>&nbsp;&nbsp;&nbsp;&nbsp;
																		<%if trim(sAttributeID)="" or IsNull(sAttributeID) or sAttributeID="NULL" then%>
																				<input type="button" value="Add Details" name="B3" class="AddButtonX" onClick="AddDetails('<%=sCheck%>')">
																		<%end if%>
																	</td>
																</tr>
																<%if trim(sAttributeID)<>"" and not IsNull(sAttributeID) and sAttributeID<>"NULL" then%>
																<tr>
																	<td class="FieldCell"></td>
																	<td class="FieldCell">Attributes &nbsp	</td>
																	<td class="FieldCell">
																		<select size="1" name="selAttribute" class="FormElem">
																			<%
																			'	sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID from  INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where "&_
																			'		" O.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemTypeID='"& sItmType &"' and A.ItemTypeAttributeID in ("& sAttributeID &") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID"
																				sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,'' from  INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where "&_
																				    " O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID in ("& sAttributeID &") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName"
																				rsTemp.Open sQuery,con
																				if not rsTemp.EOF then
																					Response.Write "<option value="& trim(rsTemp(0))&">"&Trim(rsTemp(1))&"</option>"
																				end if
																				rsTemp.Close

																			'	sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O "&_
																			'	" where O.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemTypeID = '"& sItmType &"' and A.ItemTypeAttributeID = "& sAttributeID
																				sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O "&_
																				    " where O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID = "& sAttributeID
																				rsTemp.Open sQuery,con
																				if not rsTemp.EOF then
																					do while not rsTemp.EOF
																						Response.Write "<option value="& trim(rsTemp(2))&">"&Trim(rsTemp(3))&"</option>"
																						rsTemp.MoveNext
																					loop
																				end if
																				rsTemp.Close
																			%>
																		</select>
																	</td>
																	<td class="FieldCellSub"><input type="button" value="Add Details" name="B3" class="AddButtonX" onClick="AddDetails('<%=sCheck%>')"></td>
																	<td class="FieldCellSub">
																	</td>
																</tr>
																<%end if%>
																<!--tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Lot Number</td>
																    <td class="FieldCellSub">
																    <%	if sType = "S" then %>
																        <input type="text" name="txtLotNumber" READONLY size="5" value="N/A" class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtLotNumber" size="15" class="FormElem">
																    <%	end if %>
																    &nbsp;
																    Serial No. Prefix&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtPrefix" READONLY value=0 size="12" maxlength=12 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtPrefix" size="12" maxlength=12 class="FormElem">
																    <%	end if %>
																    &nbsp;From&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtSerialFrom" READONLY value=0 size="5" maxlength=5 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtSerialFrom" size="5" maxlength=5 class="FormElem" onkeypress="DoKeyPress('<%=sCheck%>',5,1)">
																    <%	end if %>
																    &nbsp;To&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtSerialTo" READONLY value=0 size="5" maxlength=5 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtSerialTo" size="5" maxlength=5 class="FormElem" onkeypress="DoKeyPress('<%=sCheck%>',5,1)">
																    <%	end if %>
																    &nbsp;Suffix&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtSuffix" READONLY value=0 size="13" maxlength=13 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtSuffix" size="13" maxlength=13 class="FormElem">
																    <%	end if %>
																    </td>
																</tr-->

																<!--tr>
																	<td class="FieldCell"></td>
																	<%	if sType = "S" then %>
																    <td class="FieldCell">Serial Value</td>
																    <td class="FieldCell">
																		<input type="radio" value="I" name="radValue" class="FormElem" onclick="DisableTxtValue(this)">  Individual
																		<input type="radio" value="U" name="radValue" class="FormElem" onclick="DisableTxtValue(this)">  Uniform&nbsp;&nbsp;
																		<input type="text" name="txtValue" size="11" onkeypress="DoKeyPress('Y',7,3)" class="FormElem">
																    </td>
																	<%	else	%>
																    <td class="FieldCell">Lot Value</td>
																    <td class="FieldCellSub">
																		<input type="text" name="txtValue" size="15" onkeypress="DoKeyPress('Y',7,3)" class="FormElem">
																    </td>
																	<%	end if	%>
																</tr-->
																<!--tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity</td>
																    <td class="FieldCell">
																		<input type="radio" value="I" name="radQty" class="FormElem" onclick="DisableTxtQty(this)">  Individual
																		<input type="radio" value="U" name="radQty" class="FormElem" onclick="DisableTxtQty(this)">  Uniform&nbsp;&nbsp;
																		<input type="text" name="txtQuantity" size="11" onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem">
																    </td>
																</tr>
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr-->
															</table>
                                                        </td>
													</tr>
												</table>
                                            </td>
										</tr>
<!--Table 2-->

<!--Table 3-->

<!--Table 3-->

                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td>
									<div class="frmBody" id="frm2" style="width: 100%; height:170;">
										<table border="0" cellspacing="1" id="tblLot" class="ExcelTable" width="100%">
											<tr>
											<%	if ucase(sItmType) = "FAB" then %>
												<td class="ExcelHeaderCell" align="Left" colspan="10" >Lot No. / Value [Rate/<%=sStoresUom%>]</td>
											<%	else	%>
												<td class="ExcelHeaderCell" align="Left" colspan="9" >Lot No. / Value [Rate/<%=sStoresUom%>]</td>
											<%	end if	%>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
												<!--td class="ExcelHeaderCell" align="center" rowspan="2">Serial No.</td-->
												<td class="ExcelHeaderCell" align="center" colspan="2">Received</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Packing Type</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="100">Packing No</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="50">Value</td>
											<%	if ucase(sItmType) = "FAB" then %>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="50">Quantity Breakup</td>
											<%	end if
												'if ucase(sType) = "S" then %>

											<%	'end if%>
												<td class="ExcelHeaderCell" align="center" colspan="2">Selling</td>
												<!--td class="ExcelHeaderCell" align="center" rowspan="2">Selling Form</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">W/t. Per Form</td-->
												<td class="ExcelHeaderCell" align="center" rowspan="2">Packing Form</td>

											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center">Gross</td>
												<td class="ExcelHeaderCell" align="center">Tare</td>
												<td class="ExcelHeaderCell" align="center">Form</td>
												<td class="ExcelHeaderCell" align="center">W/t. Per</td>
											</tr>

										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<!--Begining of Code added by Tajudeen-->
							<tr>
							    <td class="FieldCell" colspan="2"><p align="center">
									<!--input type="button" value="Add Serial" name="B4" class="AddButtonX" onClick="AddSerial()">
									<input type="button" value="  Add Lot " name="B5" class="AddButtonX"  onClick="AddLot()"-->
								</td>
							</tr>
							<!--End of Code added by Tajudeen-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()" >
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton" onClick="clearAll()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

<%
	' Function to populate Stage
	Function PopulateStage(sItmType)
		' Declaration of variables
		Dim dcrs, StageCode,StageName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STAGEID, STAGENAME FROM INV_M_STAGE WHERE ITEMTYPEID = " & Pack(sItmType) & " AND (CATEGORYID IS NULL OR CATEGORYID = 'Q')"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set StageCode = dcrs(0)
		set StageName = dcrs(1)

		do while Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(StageCode)&""">"&trim(StageName)&"</OPTION>" &vbcrlf)
			dcrs.movenext
		loop
		dcrs.Close

	End Function
%>

<%
	' Function to Display UoM
	Function DisplayUoM(iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ")"
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to Get Selling Form
	Function GetSellingForm()
		' Declaration of variables
		Dim dcrs,dcrs1,iCodeLen,iCodeSize,sForm
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	'Response.Write "sItmType="&sItmType
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT CODELENGTH FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form' AND ITEMTYPEID = " & Pack(sItmType) & ""
			.Source = "SELECT CODELENGTH FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		If not dcrs.EOF Then
			iCodeLen = cint(dcrs(0))
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT ISNULL(SUM(CODELENGTH)+1,0) FROM APP_M_CODETYPES WHERE DISPLAYORDER < (SELECT DISPLAYORDER FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form') AND ITEMTYPEID = " & Pack(sItmType) & ""
				.Source = "SELECT ISNULL(SUM(CODELENGTH)+1,0) FROM APP_M_CODETYPES WHERE DISPLAYORDER < (SELECT DISPLAYORDER FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form') "
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			If not dcrs1.EOF Then
				iCodeSize = cint(dcrs1(0))
			end if
			dcrs1.Close
		end if
		dcrs.Close

		if iCodeSize > 0 and iCodeLen > 0 then
			sForm = trim(mid(sComCode,iCodeSize,iCodeLen))

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS WHERE CODE = " & Pack(sForm) & " AND ITEMTYPEID = " & Pack(sItmType) & ""
				.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS "
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			If not dcrs.EOF Then
				do while not dcrs.EOF
					'Response.Write "<option value="""&trim(dcrs(0))&""">"&trim(trim(dcrs(1)))&"</option>" & vbCrLf

					Set newElem = oDOM.createElement("SELLFORM")
					newElem.setAttribute "UNITID",trim(dcrs(0))
					newElem.setAttribute "UNITNAME",trim(dcrs(1))

					Root.appendChild newElem

				dcrs.MoveNext
				loop
			end if
			dcrs.Close
		end if
	End Function
%>

<%
	' Function to populate UoM List
	Function populateUoMList(sUoM)
		' Declaration of variables
		Dim dcrs,sUomDesc,sUomShDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE NOT IN (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ") AND UOMCODE <> " & Pack(sUoM) & " ORDER BY UOMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUoMCode = dcrs(0)
		set sUomDesc = dcrs(1)
		set sUomShDesc = dcrs(2)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUoMCode)&""">"&trim(sUomShDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
