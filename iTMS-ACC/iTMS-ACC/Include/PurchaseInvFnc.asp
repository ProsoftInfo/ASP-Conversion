

<%
'Function to fetch Order Confirmation Item Rate for given Receipt No / Item
function getOrderRateForReceipt(iRcptNo,sOrgID,iClassCode,iItemCode)

	Dim sSql,objRS,dOrderRate

	Set objRS = Server.CreateObject("ADODB.Recordset")

	'sSql  = "SELECT DISTINCT ISNULL(A.ITEMRATE ,0) FROM TPOTRPOCONFDET A, PUR_T_REFFERENCENUMBERDET B," &_
	'		" TPOMSITMUNITREL C WHERE A.CONFIRMATIONNO = B.OCNUMBER AND B.RECEIPTNUMBER =" & iRcptNo & " AND " &_
	'		" RIGHT(A.ORDDRWGSTORENO,12) = C.DRAWINGVERSIONNO AND C.ITEMCODE =" & trim(iItemCode)  & " AND " &_
	'		" C.CLASSIFICATIONCODE =" & trim(iClassCode) & " AND C.ORGANISATIONCODE = '" & trim(sOrgID) & "'"

	'with objRS
	'	.CursorLocation = 3
	'	.CursorType = 3
	'	.Source = sSql
	'	.ActiveConnection = con
	'	.Open
	'end with
	'set objRS.ActiveConnection = nothing

	dOrderRate = 0

	'If not objRS.eof then
	'	dOrderRate = objRs(0)
	'Else

	'	objRs.close

		'' To fetch Item rate from Actual receipts if Order Rate is not available

		sSql  = "SELECT DISTINCT ISNULL(ITEMRATE ,0) FROM RCV_T_ACTUALRCPTITEMDET WHERE ITEMCODE =" & trim(iItemCode) & " AND " &_
					"CLASSIFICATIONCODE =" & trim(iClassCode) & " AND ORGANISATIONCODE = '" & trim(sOrgID) & "' and RECEIPTNUMBER=" & trim(iRcptNo) & ""

		with objRS
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		set objRS.ActiveConnection = nothing

		If not objRS.eof then
			dOrderRate = objRs(0)
		end if

	'End if
	objRs.close

	getOrderRateforReceipt = dOrderRate

end function
%>


<%
'Function to Update QtyPurchased in Sal_T_ForProcessingDet table,
'After Receipt Inv. Acc. if PR is based on Sales Order

function setPurQtyForSalOrder(iRcptNo,sOrgID,iClassCode,iItemCode,dReceiptQty)
	setPurQtyForSalOrder = false
	exit function
	Dim sSql,objRS, iSaleProcessNo, blnUptFlag,sReceiptUoM,sSalesProcessUoM
	Dim sTemp,saTemp,sStoreUoM,sPurUoM,iPurToStoreRate,iPurToStoreOpr,sSalesUoM
	Dim iSaleToStoreRate,iSaleToStoreOpr,iStoreToSaleOpr,dStoreUoMPurQty,iBaseToOptOperator
	Dim dSalesUoMRcptQty,dPurQty,iOptToBaseRate,iOptToBaseOperator,dSalesPrcUoMRcptQty

	Set objRS = Server.CreateObject("ADODB.Recordset")

	sSql = "Select SalesProcessNo from Pur_T_RefferenceNumberDet where PRNumber = " &_
			" (Select Distinct B.IndentNo from Pur_T_RefferenceNumberDet A ,tpotrPOQuote B Where " &_
			" A.OCNumber = B.ConfirmationNo and A.ReceiptNumber =" & trim(iRcptNo) & ")"
	with objRS
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set objRS.ActiveConnection = nothing

	iSaleProcessNo = 0

	If not objRS.eof then
		iSaleProcessNo = objRs(0)
	End if
	objRs.close

	blnUptFlag = False

	'' To fetch Receipt UoM
	sSql = "Select Isnull(UoMCode,'') from RCV_T_ActualRcptItemDet Where ReceiptNumber=" & trim(iRcptNo) &_
			 " and OrganisationCode='" & trim(sOrgID) & "' and ItemCode=" & trim(iItemCode) & " and ClassificationCode=" & trim(iClassCode) & ""
	with objRS
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set objRS.ActiveConnection = nothing

	If not objRS.eof then
		sReceiptUoM = objRs(0)
	Else
		sReceiptUoM = ""
	End if
	objRs.close
	'------------------------------------------------------------------------

	'' To fetch SalesProcess UoM
	sSql = "Select Isnull(ProcessingUoM,'') from SAL_T_ForProcessingDet Where ProcessingNo=" & trim(iSaleProcessNo) &_
			 " and ItemCode=" & trim(iItemCode) & " and ClassificationCode=" & trim(iClassCode) & ""
	with objRS
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set objRS.ActiveConnection = nothing

	If not objRS.eof then
		sSalesProcessUoM  = objRs(0)
	Else
		sSalesProcessUoM = ""
	End if
	objRs.close
	'------------------------------------------------------------------------

	'' To convert Receipt Qty to Sales Process UoM before SalesProcess Qty updation

	'' To fetch Purchase, Stores, SalesUoM & corresponding conversion factors based on Flag

	'Store UoM ---------
	sTemp = getUoMConvFactor(sOrgID,iClassCode,iItemCode,"STORE")
	saTemp = Split(sTemp,":")
	sStoreUoM = saTemp(0)
	''------------------

	'Purchase UoM ------
	sTemp = getUoMConvFactor(sOrgID,iClassCode,iItemCode,"PUR")
	saTemp = Split(sTemp,":")
	sPurUoM = saTemp(0)
	iPurToStoreRate  = saTemp(1)
	iPurToStoreOpr  = saTemp(2)
	''------------------

	'Sales UoM
	sTemp = getUoMConvFactor(sOrgID,iClassCode,iItemCode,"SALES")
	saTemp = Split(sTemp,":")
	sSalesUoM = saTemp(0)
	iSaleToStoreRate = saTemp(1)
	iSaleToStoreOpr = saTemp(2)

	''------------------------------------------------------------------------------
	''ReceiptUoM - PurchaseUoM
	''PurchaseUoM - StoreUoM
	''StoreUoM - Sales UoM
	''SalesUoM - SalesProcessUoM

	' To check if SalesProcess UoM & SalesUoM differ & To convert Qty if so
	'Response.Write sSalesProcessUoM + " : " + sReceiptUoM
	If trim(sSalesProcessUoM) = trim(sReceiptUoM) Then
		dSalesPrcUoMRcptQty = dReceiptQty
	Else

		If trim(sReceiptUoM) = trim(sPurUoM) Then
			dPurQty = dReceiptQty

		ElseIf trim(sReceiptUoM) <> trim(sPurUoM) Then

			sTemp = getOptionBaseRateForSale(sOrgID,iClassCode,iItemCode,sReceiptUoM )

			saTemp = split(sTemp,"|")

			iOptToBaseRate = saTemp(0)
			iOptToBaseOperator = saTemp(1)

			dPurQty = getRatePerQtyUoM(iOptToBaseOperator,iOptToBaseRate,dReceiptQty)
		End If

		If trim(sPurUoM) = trim(sSalesProcessUoM) Then
			dSalesPrcUoMRcptQty = dPurQty
		Else
			If trim(sPurUoM) = trim(sStoreUoM) Then
				dStoreUoMPurQty = dPurQty
			ElseIf trim(sPurUoM) <> trim(sStoreUoM) Then
				if iPurToStoreOpr  = 0 then
					iPurToStoreOpr = 1
				else
					iPurToStoreOpr = 0
				end if
				dStoreUoMPurQty = getRatePerQtyUoM(iPurToStoreOpr,iPurToStoreRate,dPurQty)
			End If

			If trim(sStoreUoM) = trim(sSalesUoM) Then
				dSalesUoMRcptQty = dStoreUoMPurQty
			ElseIf trim(sSalesUoM) <> trim(sStoreUoM) Then
				'if iSaleToStoreOpr  = 0 then
				'	iStoreToSaleOpr = 1
				'else
				'	iStoreToSaleOpr = 0
				'end if
				dSalesUoMRcptQty = getRatePerQtyUoM(iSaleToStoreOpr,iSaleToStoreRate,dStoreUoMPurQty)
			End If

			If trim(sSalesUoM) = trim(sSalesProcessUoM) Then
				dSalesPrcUoMRcptQty = dSalesUoMRcptQty
			ElseIf trim(sSalesUoM) <> trim(sSalesProcessUoM) Then

				sTemp = getOptionBaseRateForSale(sOrgID,iClassCode,iItemCode,sSalesProcessUoM )

				saTemp = split(sTemp,"|")

				iOptToBaseRate = saTemp(0)
				iOptToBaseOperator = saTemp(1)

				If iOptToBaseOperator = 0 then
					iBaseToOptOperator = 1
				Elseif iOptToBaseOperator = 1 then
					iBaseToOptOperator = 0
				End if

				dSalesPrcUoMRcptQty = getRatePerQtyUoM(iBaseToOptOperator,iOptToBaseRate,dSalesUoMRcptQty)
			End If
		End if
	End if
	'Response.Write cstr(dPurUoMSalesQty) + " : " + cstr(dPRQty) + " : " + cstr(cdbl(dPurUoMSalesQty) - cdbl(dPRQty)) + "<br>"

	'-------------------------------------------------------------------------------

	If cint(iSaleProcessNo) <> 0 then

	  sSql = "Update Sal_T_ForProcessingDet set QtyPurchased = QtyPurchased + " & trim(dSalesPrcUoMRcptQty) & " Where ProcessingNo=" & trim(iSaleProcessNo) & " and " &_
		" ItemCode = " & trim(iItemCode) & " and ClassificationCode = " & trim(iClassCode)  & ""
	  'Response.write sSql

	  con.Execute sSql

	  blnUptFlag = True

	End if

	setPurQtyForSalOrder = blnUptFlag
end function
%>


<%
function getUoMConvFactor(sOrgID,iClassCode,iItemCode,blnFlag)

Dim sSql,rsItem, sUoM,iToStoreRate,iToStoreOpr,sTemp
Set rsItem = Server.createobject("adodb.recordset")

If blnFlag = "STORE" Then
	sSql = "Select StoresUoM,0,'0' from Inv_M_ItemOrgMaster Where Organisationcode='" & trim(sOrgID) & "' and ClassificationCode=" & trim(iClassCode) & " and ItemCode=" & trim(iItemCode) & ""

ElseIf blnFlag = "PUR" Then
	sSql = "Select PurchaseUoM, isnull(PurToStoreRate,1) , isnull(PurToStoreOperator,'0')  from Inv_M_ItemOrgMaster Where Organisationcode='" & trim(sOrgID) & "' and ClassificationCode=" & trim(iClassCode) & " and ItemCode=" & trim(iItemCode) & ""

ElseIf blnFlag = "SALES" Then
	sSql = "Select SalesUoM,isnull(SaleToStoreRate,1) ,isnull(SaleToStoreOperator,'0') from Inv_M_ItemOrgMaster Where Organisationcode='" & trim(sOrgID) & "' and ClassificationCode=" & trim(iClassCode) & " and ItemCode=" & trim(iItemCode) & ""
End if

With rsItem
	.CursorLocation = 3
	.CursorType = 3
	.Source =  sSql
	.ActiveConnection = con
	.Open
End With
Set rsItem.ActiveConnection = nothing

If Not rsItem.EOF and not isnull(rsItem(0)) then
	sUoM = rsItem(0)
	iToStoreRate = rsItem(1)
	iToStoreOpr = rsItem(2)
Else
	sUoM = ""
	iToStoreRate = 0
	iToStoreOpr = 0
End if
rsItem.Close

sTemp = cstr(sUoM) + ":" + cstr(iToStoreRate) + ":" + cstr(iToStoreOpr)

getUoMConvFactor = sTemp

end function
%>

<%
function getRatePerQtyUoM(iOptToBaseOperator,iOptToBaseRate,RATE)
Dim dRatePerQtyUoM

	'' if conversion operator is 0 : Multiply

	if cint(iOptToBaseOperator) = 0 then	' Multiply

		dRatePerQtyUoM = cdbl(RATE) * cdbl(iOptToBaseRate)

	elseif cint(iOptToBaseOperator) = 1 then ' Divide

		'' if conversion operator is 1 : divide
		dRatePerQtyUoM = cdbl(RATE) / cdbl(iOptToBaseRate)
	end if

	getRatePerQtyUoM = dRatePerQtyUoM	'RETURN RATE/QTY UOM

end function
%>

<%
'Function to fetch the Receipt Type based on Actual Receipt No.
function getReceiptType(iRcptNo)

	Dim sSql,objRS, sReceiptType

	Set objRS = Server.CreateObject("ADODB.Recordset")

	' Ref Aganist because of finding reference aganist
	sSql = "Select ReceiptAs from RCV_T_ActualReceiptHeader where ReceiptNumber=" & trim(iRcptNo) & ""
	with objRS
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set objRS.ActiveConnection = nothing

	If not objRS.EOF Then
		sReceiptType = objRS(0)
	Else
		sReceiptType = "0"
	End if
	objRS.close

	getReceiptType = sReceiptType

end function

%>

<%
'Function to fetch the Receipt Type based on Actual Receipt No.
function getReceiptRefType(iRcptNo)

	Dim sSql,objRS, sReceiptType

	Set objRS = Server.CreateObject("ADODB.Recordset")

	' Ref Aganist because of finding reference aganist
	sSql = "Select A.RefAgainst from Rcv_T_GateReceiptHeader A,Pur_T_RefferenceNumberDet B where " &_
			" A.GRNNumber = B.GRNNumber and B.ReceiptNumber=" & trim(iRcptNo) & ""
	with objRS
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set objRS.ActiveConnection = nothing

	If not objRS.EOF Then
		sReceiptType = objRS(0)
	Else
		sReceiptType = "0"
	End if
	objRS.close

	getReceiptRefType = sReceiptType

end function

%>

<%
Function ItemLocationStockUpdate (sOrgID,iClassCode,iItemCode,LocationType,LocationNo,sBinNo,passQty,PassValue,dtTransDate,sAction)
'Note : here sAction indicate  A (Add ) / S (Subtract )
Dim sTempMonYr,sMonYr,arrFin,sFinFrom,sFinTo,sSql
Dim dcrs

set dcrs = Server.CreateObject("ADODB.RecordSet")

sTempMonYr = mid(dtTransDate ,4,2)
sMonYr = sTempMonYr&Year(dtTransDate)

arrFin = split(GetFinancialYear(sMonYr),":")
sFinFrom = arrFin(0)
sFinTo = arrFin(1)

if sBinNo ="NULL" then sBinNo = "0"

If sAction ="A" or sAction = "S" Then

'' to insert/Update Location stock table only if location is selected
if trim(LocationNo) <> "0" and not isnull(locationNo) then

	'sSql = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & LocationNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND MONTHANDYEAR = " & Pack(sMonYr) & ""
	sSql = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & LocationNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103) and FinancialYearTo = Convert(datetime,'"& sFinTo &"',103)"
	Response.Write "<p>" &  sSql & vbCrLf & vbCrLf
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if dcrs.EOF then
		If sAction ="A" Then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,FinancialYearFrom,FinancialYearTo,ReceiptQuantity,ReceiptValue) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClassCode & "," & iItemCode & "," &_
				"" & LocationNo & "," & sBinNo & ",Convert(datetime," & Pack(sFinFrom) & ",103),Convert(datetime,"& Pack(sFinTo) &",103)," & passQty  & "," & PassValue  & ")"
		Else
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,FinancialYearFrom,FinancialYearTo,ISSUEQUANTITY,ISSUEVALUE) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClassCode & "," & iItemCode & "," &_
				"" & LocationNo & "," & sBinNo & ",Convert(datetime," & Pack(sFinFrom) & ",103),Convert(datetime,"& Pack(sFinTo) &",103)," & passQty  & "," & PassValue  & ")"
		End If 'If sAction ="A" Then
	'	Response.Write "<p>" &  sSql & vbCrLf & vbCrLf
		con.Execute sSql
	else
		If sAction ="A" Then
			'sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET ReceiptQuantity = (ReceiptQuantity + " & passQty & ")," &_
			'	"ReceiptValue = (ReceiptValue + " & PassValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
			'	"CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			'	"LOCATIONNUMBER = " & LocationNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
			'	"MONTHANDYEAR = " & Pack(sMonYr) & ""
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YearReceiptQuantity = (YearReceiptQuantity + " & passQty & ")," &_
				"YearReceiptValue = (YearReceiptValue + " & PassValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & LocationNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
				"FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103) and FinancialYearTo = Convert(datetime,'"& sFinTo &"',103)"
		Else
			'sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET ISSUEQUANTITY = (ISSUEQUANTITY + " & passQty & ")," &_
			'	"ISSUEVALUE = (ISSUEVALUE + " & PassValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
			'	"CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			'	"LOCATIONNUMBER = " & LocationNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
			'	"MONTHANDYEAR = " & Pack(sMonYr) & ""
			
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & passQty & ")," &_
				"YEARISSUEVALUE = (YEARISSUEVALUE + " & PassValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & LocationNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
				"FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103) and FinancialYearTo = Convert(datetime,'"& sFinTo &"',103)"
		End If ''If sAction ="A" Then
		Response.Write "<p>" & sSql & vbCrLf & vbCrLf
		con.Execute sSql
	end if
	dcrs.Close


	If sAction ="A" Then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
				"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClassCode & "," & iItemCode & "," &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & passQty & "," & PassValue & "," & passQty & "," & PassValue & ")"
		else
			sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YearReceiptQuantity = (YearReceiptQuantity + " & passQty & ")," &_
				"YearReceiptValue = (YearReceiptValue + " & PassValue & ")," &_
				"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & passQty & "), " &_
				"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & PassValue & ") WHERE " &_
				"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
				"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
				"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		end if
	'	Response.Write "<p>" & sSql & vbCrLf & vbCrLf
		con.Execute sSql
		dcrs.Close

	Else ''If sAction ="S" Then
		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & passQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & PassValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & passQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & PassValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
	'	Response.Write "<p>" & sSql & vbCrLf & vbCrLf
		con.Execute sSql

		'Response.Write "<p>" & sSql & vbCrLf & vbCrLf
		con.Execute sSql
	End If ''If sAction ="A" Then

	End if ' For Location NA case

End If  'If sAction ="A" or sAction = "S" Then
End function

'Test Data
'ItemLocationStockUpdate "010101",23,7,"PU",3,"NULL",100,1000,"23/12/2003","S"
%>

<%
' Function to check for Number series definition in Inventory table for
' Lot No. or Sample Label No. for the Item Type & organisation
Function checkInvNumSeriesEntry(sOrgID,sActivityType,sItemType)

Dim dRSet,sSql, blnFlag
'response.write sOrgID + " : " + sActivityType + " : " + sItemType
sSql = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE =" & Pack(sActivityType) & " AND ITEMTYPE = " & Pack(sItemType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
'response.write sSql
Set dRSet = Server.CreateObject("ADODB.RecordSet")
with dRSet
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
end with
set dRSet.ActiveConnection = nothing

If not dRSet.eof Then
	blnFlag = True
Else
	blnFlag = False
End if
dRSet.close

checkInvNumSeriesEntry = blnFlag

End Function
%>

<%
Function getOptionBaseRateForSale(sOrgID,iClassCode,iItemCode,sRateUoM)
Dim sSql,rsItem,iOptToBaseRate,iOptToBaseOperator, sTemp

Set rsItem = Server.createobject("adodb.recordset")

	sSql = "Select IM.OptionToBaseRate,IM.OptionToBaseOperator from INV_M_ITEMORGOPTIONALUOM IM " &_
	 " Where IM.itemcode="& trim(iItemCode) &" and IM.classificationcode="& trim(iClassCode) &" and IM.organisationcode='"&trim(sOrgID)&"' " &_
	 " and IM.UOMCode = '" & trim(sRateUoM) & "' And IM.OptionalUoMFor='S'"
	'response.write ssql + "<br>"
	With rsItem
		.CursorLocation = 3
		.CursorType = 3
		.Source =  sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsItem.ActiveConnection = nothing

	If Not rsItem.EOF then
		iOptToBaseRate = rsItem(0)
		iOptToBaseOperator = rsItem(1)
	Else
		iOptToBaseRate = 1
		iOptToBaseOperator = 0
	End if
	rsItem.Close

	sTemp = cstr(iOptToBaseRate) + "|" + cstr(iOptToBaseOperator)

	getOptionBaseRateForSale = sTemp

end function
%>