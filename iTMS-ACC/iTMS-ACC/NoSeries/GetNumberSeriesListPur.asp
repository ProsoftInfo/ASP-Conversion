<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetNumberSeriesListPur.asp
	'Module Name				:	Purchase (Master Creation)
	'Author Name				:	Malathi N
	'Created On					:	Sep 29,2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	PurNoSeriesInsert.asp
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<%
	Dim sUnit,sActitity,sTemp,sTempVal,sQuery,Objrs,iCtr,sSerCode
	Dim sMon,sYear,sMonYr,sFinYear,sFinFrom,sFinTo,saTemp,sCrValue,sDrValue
	Dim oDom,newElem,Root,sSeriesCode,sSeriesNo,sRetTy,objRs1,sItemTypeName
	Dim iSeriesCode,iSeriesNo,iAgentCode,iAddSerNo,iAddSerCode,sAgentName
	Dim iEntNo,sItemTy,sInvTy,sSalTy,sAgTy,iSerNo,iSerCode,sItemVal,sSalVal,sInvVal,sNumTy
	Dim sItemDesc,sInvDesc,sSalDesc,sTempVar,sNoUsed,Objrs3,sSelClassCodes,sSelCatCodes,sSelClassCodesName,sSelCatCodesName
	
	
	sTempVal = Request("sVal")
	sTemp = Split(sTempVal,":")
	
	iCtr = 1
	sUnit = sTemp(0)
	sActitity = sTemp(1)
	
	Set oDom = server.CreateObject("Microsoft.xmlDom")
	Set Root = oDom.createElement("Root")
	oDom.appendChild Root
	
	sMon = Month(Date)
	sYear = Year(Date)

	IF CInt(sMon) <=9 Then
		sMon = 0&sMon
	End IF
	sMonYr = sMon&sYear
	sFinYear = GetFinancialYear(sMonYr)
	saTemp = Split(sFinYear,":")
	sYear = Right(saTemp(0),4)
	sMon = Mid(saTemp(0),4,2)
	sFinFrom = sYear&sMon

	sYear = Right(saTemp(1),4)
	sMon = Mid(saTemp(1),4,2)
	sFinTo = sYear&sMon
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs3 = Server.CreateObject("ADODB.RecordSet")
	
	sQuery = "Select NoSeriesTransactionNo,OrganisationCode,NumberFor,isNull(SeriesNo,0),isNull(SeriesCode,0) "&_
			 "From PUR_M_Noseries Where NoSeriesStatus <> '1'  and OrganisationCode = '"&sUnit&"' and ActivityType = '"&sActitity&"' "
															 
	With Objrs
		.CursorType = 3
		.CursorLocation = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	'Response.Write sQuery
	Set Objrs.ActiveConnection = Nothing
	
	Do While Not Objrs.EOF
		iSeriesNo = Objrs(3)
		iSeriesCode = Objrs(4)
		sSelClassCodes=""
		sSelCatCodes =""
		sSelClassCodesName=""
		sSelCatCodesName =""
		sQuery = "Select isNull(ItemType,0),isNull(MainSeriesNo,0),isNull(MainSeriesCode,0),isnull(ItemValue,0) "&_
					 " From VwPurNoSeriesSel Where NoSeriesTransactionNo = "&Objrs(0)&" "
			With Objrs1
				.CursorType = 3
				.CursorLocation = 3
				.ActiveConnection = Con
				.Source = sQuery
				.Open
			End With
	'		Response.Write sQuery
			Set Objrs1.ActiveConnection = Nothing
			IF Not Objrs1.EOF Then
				iEntNo = "1"
				sItemTy = objRs1(0)
				sItemVal = objRs1(3)
			End IF
			objRs1.Close
			
			sNoUsed = CheckNoSerUsed(sUnit,sActitity,sItemTy,sItemVal)
		'	Response.Write sNoUsed
			
			IF CStr(sItemTy) = "0" Then
				sItemTy = "All"
				sItemDesc = ""
			Else
				sItemTy = "Specific"
				sTempVar = sItemVal
				sItemDesc = GetDesc(sTempVar)
			End IF
		
			
			IF CStr(Objrs(2)) = "B" Then
				sNumTy = "Both"
			Elseif CStr(Objrs(2)) = "D" Then
				sNumTy = "Domestic"
			Else
				sNumTy = "Import"
			End IF
			
			sQuery = "Select IsNull(ClassCode,0),isNull(CatCode,0) from PUR_M_NoSeriesClass where SeriesNo = "& iSeriesNo &" and SeriesCode = "& iSeriesCode 
			objRs1.Open sQuery,con
			if not objRs1.EOF then
				do while not objRs1.EOF 
					if Trim(objRs1(0))<>"0" then
						sSelClassCodes = sSelClassCodes &","& objRs1(0)
					else
						sSelClassCodes = sSelClassCodes &","& "NULL"
					end if
					if Trim(objRs1(1))<>"0" then
						sSelCatCodes = sSelCatCodes &","& objRs1(1)
					else
						sSelCatCodes = sSelCatCodes &","& "NULL"
					end if
					
					objRs1.MoveNext
				loop
			end if
			objRs1.Close 
			
			
			if Trim(sSelClassCodes)<>"" then sSelClassCodes = Mid(sSelClassCodes,2)
			if Trim(sSelCatCodes)<>"" then sSelCatCodes = Mid(sSelCatCodes,2)
			
			sQuery = "Select GroupName from INV_M_Classification where GroupCode in ("& sSelClassCodes  &")"
			objRs1.Open sQuery,con
			if not objRs1.EOF then
				do while not objRs1.EOF 
					sSelClassCodesName =sSelClassCodesName & ","&  Trim(objRs1(0))
					objRs1.MoveNext 
				loop
			end if
			objRs1.Close 
			if Trim(sSelClassCodesName)<>"" then sSelClassCodesName =  Mid(sSelClassCodesName,2)
			
			sQuery = "Select CategoryName from Inv_M_ClassificationCategory where CategoryCode in ("& sSelCatCodes  &")"
			objRs1.Open sQuery,con
			if not objRs1.EOF then
				do while not objRs1.EOF 
					sSelCatCodesName =sSelCatCodesName & ","&  Trim(objRs1(0))
					objRs1.MoveNext 
				loop
			end if
			objRs1.Close 
			if Trim(sSelCatCodesName)<>"" then sSelCatCodesName =  Mid(sSelCatCodesName,2)  
			
			 
			Set newElem = oDom.createElement("NumSeriesList")
			newElem.setAttribute "TransNo",Objrs(0)
			newElem.setAttribute "NumFor",sNumTy
			newElem.setAttribute "SeriesNo",iSeriesNo
			newElem.setAttribute "SeriesCode",iSeriesCode
			newElem.setAttribute "EntryNo",iEntNo
			newElem.setAttribute "ItemTy",sItemTy
			newElem.setAttribute "ItemValue",sItemVal
			newElem.setAttribute "ItemDesc",sItemDesc
			newElem.setAttribute "EditCheck","N"
			newElem.setAttribute "NoUsed",sNoUsed
			newElem.setAttribute "SelClass",sSelClassCodes
			newElem.setAttribute "SelCat",sSelCatCodes
			newElem.setAttribute "SelClassName",sSelClassCodesName
			newElem.setAttribute "SelCatName",sSelCatCodesName  
		Root.appendChild newElem
		iCtr = iCtr + 1
		Objrs.MoveNext
	Loop
	Objrs.Close
		
	
	Response.ContentType="text/xml"
	Response.Write oDom.xml											
	
	
	
%>
<%
Function CheckNoSerUsed(Unit,Activity,ItemType,ItemVal)
dim sSql,rsTemp,sDesc,sOrgID
sOrgID = getUnitNoOUDefID(Unit,"U")
'Response.Write "sOrgID"&sOrgID


Set rsTemp = server.CreateObject("ADODB.RecordSet")
'-----------------------------------------------------------------------------
if Activity = "1" then ''Purchase requisition
'	sSql = "select PREntryNumber from PUR_T_PRHeader where PRNumber is not null and RequestedByUnit='"&sOrgID&"' "&_
'		"	and Category='"&ItemVal&"' "
		
	sSql = "select PREntryNumber from PUR_T_PRHeader where PRNumber is not null and RequestedByUnit='"&sOrgID&"' "
elseif Activity = "2" then  ''Request for Quote
'	sSql = "Select RFQNumber from PUR_T_RFQHeader where RFQCode is not null and ForUnit='"&sOrgID&"' "&_
'		"	and ItemTypeID='"&ItemVal&"' "
	sSql = "Select RFQNumber from PUR_T_RFQHeader where RFQCode is not null and ForUnit='"&sOrgID&"' "

elseif Activity = "3" then ''Quotation
	sSql = "Select QuotationNo from PUR_T_QUoteSampleDetails where LabelNo is not null and QuotationNo in ( "&_
		" Select QuotationNo from PUR_T_QuoteHeader where ForUnit='"&sOrgID&"' and ItemTypeID='"&ItemVal&"') "	
	

elseif Activity = "4" then ''Purchase Order
'	sSql = "Select PurchaseOrderNo from PUR_T_POHeader where PurchaseOrderCode is not null and ForUnit='"&sOrgID&"' "&_
'		"	and Itemtype='"&ItemVal&"' "

	sSql = "Select PurchaseOrderNo from PUR_T_POHeader where PurchaseOrderCode is not null and ForUnit='"&sOrgID&"' "
	
elseif Activity = "5" then	''Schedule Order Release
	sSql = "Select ConfirmEntryNo from PUR_T_POConfirmHeader where ConfirmEntryCode is not null and ItemtypeID='"&ItemVal&"' "&_
		"	and ForUnit='"&sOrgID&"' "

elseif Activity = "6" then	''Payment
	sSql = "Select PaymentReferenceNo from tpotrpaydet where PaymentCode is not null "

elseif Activity = "7" then  ''gate Receipt
'	sSql = "select Grnnumber from RCV_T_GateReceiptHeader where GRNCode is not null and "&_
'		"	Grnnumber in (Select Grnnumber from RCV_T_GRNItemDetails where ItemTypeID='"&ItemVal&"' "&_
'		"	and OrganisationCode='"&Unit&"') "	
	sSql = "select Grnnumber from RCV_T_GateReceiptHeader where GRNCode is not null and "&_
		"	Grnnumber in (Select Grnnumber from RCV_T_GRNItemDetails where OrganisationCode='"&Unit&"') "	

elseif Activity = "8" then '' Sample Label
	
elseif Activity = "9" then  ''Actual receipt
'	sSql = "Select ReceiptNumber from RCV_T_ActualReceiptHeader where ReceiptCode is not null and ItemType='"&ItemVal&"' "&_
'		"	and ReceiptNumber in (Select ReceiptNumber from RCV_T_ActualRcptItemDet where OrganisationCode='"&sUnit&"') "
		
	sSql = "Select ReceiptNumber from RCV_T_ActualReceiptHeader where ReceiptCode is not null "&_
		"	and ReceiptNumber in (Select ReceiptNumber from RCV_T_ActualRcptItemDet where OrganisationCode='"&sUnit&"') "  
elseif Activity = "10" then  ''Receipt Invoice
	sSql = "Select InvoiceNumber from RCV_T_InvoiceHeader where InvoiceCode is not null and InvoiceNumber in "&_
		"	(select invoicenumber from RCV_T_InvoiceDetails where OrganisationCode='"&sUnit&"') "
elseif Activity = "11" then
    sSql = "Select InspectionNumber from RCV_T_PurchInspectionHeader where InspectionCode is not null and OrganisationCode in ("& pack(sUnit) &")"
end if
'--------------------------------------------------------------------------
'Response.Write sSql
with rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sSql
	.Open
end with
set rsTemp.ActiveConnection = nothing

if not rsTemp.EOF then
	CheckNoSerUsed = "Y"
else
	CheckNoSerUsed = "N"
end if
rsTemp.Close

	
End Function
%>


<%
	Function GetDesc(sVal)
		Dim sFullVal,arrtemp,Salrs
		Set Salrs = Server.CreateObject("ADODB.RecordSet")
		sQuery = "Select ItemTypeName from Inv_M_ItemType Where ItemTypeID = '"&sVal&"' "
		with Salrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sQuery
			.Open
		end with
		set Salrs.ActiveConnection = nothing
		if not Salrs.EOF then
			sFullVal = Salrs(0)
		end if
		Salrs.Close
		
		GetDesc = sFullVal
	End Function
	
%>
<%
	Function GetItemTypeName(ItemID)
		Dim sSql,rsTemp,sItemtypeName
		set rsTemp = server.CreateObject("ADODB.RecordSet")
		
		sSql = "Select ItemtypeName from Inv_M_ItemType where ItemTypeId = '"&ItemID&"' "
		with rsTemp
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sSql
			.Open
		End with
		set rsTemp.ActiveConnection = nothing
		If not rsTemp.eof then
			sItemtypeName = rsTemp(0)
		end if
		rsTemp.Close	
			
		GetItemTypeName = sItemtypeName			
		
	End function
%>
