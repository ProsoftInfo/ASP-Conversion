<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>

<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/PurchaseTermsConditions.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/PurChkItemSpecPack.asp"-->
<!--#include virtual="/include/CommonFunctions.asp"-->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="tex|/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<!---XML Data Island---->
<script type="application/xml" id="TempData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="OutData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="ITEMData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="TaxFormData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="PurTypeData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="ItemTaxData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="PartySubTypeData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="InvoiceDet" data-itms-xml-island="1" data-src="<%="../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml"%>"></script>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/calcAlternateUoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RefTypePop.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/amdInvPurInvoiceEntry.js"></SCRIPT>



<%
'declaring variables
Dim sOrgID,sOrgSName,sSelRcptNo,sConfNo,sPoNo,saSelRcptNo,sarrSelRcptno
Dim Curr1,Mod1,Mop,IssueBank,PayTerm,Bop,Transporter,sFlag
Dim rsTemp,rsItem,sSql,iActivityNo,sTempRcptCode
Dim iRcptNo,sRcptCode,sRcptDt,sActualReceiptNos,sRemarks
Dim nInvNo,nPurType,nCurrCode,nBenifitBank,nLoadPort,nDestPort
Dim nDisPer,nRate,nDisAmt,nNetBasicForItem,iPurTypeForItem,nRatePerQtyUOM,nItemValue

Dim oNodTaxRoot,ItemNode
Dim sTaxName,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxPer,dTaxValue,sAccHead
Dim iCrEligible,iRegCode,iToBeAccounted,iTRndoff,sTaxCatType,sOrghdno,sTaxDetailsAlreadyExist
Dim nTotalBasic,nTotalNet,nTotalTax,nTotalSubTotal

Set rsTemp = Server.CreateObject("ADODB.Recordset")
Set rsItem = Server.CreateObject("ADODB.Recordset")


dim rsSql,iCtr
dim sItemCode,sclassCode,iQtyRecd,arrdesc,sItemDesc,sClassDesc,iClassCode,iItemCode,iQtyInv
dim sRefNum,sPartyName,iEntryNo,cSuppCode
dim sSuppInvNo,sSuppInvDt,iPartyCode,sPartyType,sPartySubType
Dim sRateUoM,saTemp2,iOptToBaseRate,iOptToBaseOperator
dim sDRGNo,iTempItemCode,sOrdDrgStoreNo,dRatePerQtyUoM,dInvQTy,dItemValue,dBVValue,popPurOptionalUOM,dBalQty
dim sSql1,rsItem1,sUOM,dOrderRate,UomCode,sCategoryCode
Dim dcrs,oDOM,Root,newElem,newElem1
Dim SubNode1,SubNode,sStockType,sVAT
Dim ObjFs,objDOM2,sAccroot,sExp,sRoundoffHead,sTempnode,sBillType
Dim sFinFun,sFinPeriod,sFinTemp,sMaxDate,sMinDate



sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sFinTemp = Split(sFinPeriod,":")
	sMaxDate = "31/03/"&sFinTemp(1)
	sMinDate = "01/04/"&sFinTemp(0)
End IF


dOrderRate=0
dItemValue=0

set dcrs = Server.CreateObject("ADODB.Recordset")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

set objDOM2 = server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

'sFlag = trim(Request("Flag"))

'Response.Write "<p> Request.QueryString = " & Request.QueryString

nInvNo = Request.QueryString("InvNo")
sOrgID = Request.QueryString("ForUnit")
sOrgSName = getUnitName(sOrgID)
'
iActivityNo = "10"		'Activity Number for "RECEIPT INVOICE"
'Response.Write "nInvNo = "& nInvNo
'' To check for No. Series entry for recd. by unit.
'If not checkNumSeriesEntry(iActivityNo,sOrgID) then
'	Response.Redirect "CommonMessageNoSeries.asp?TranName=Receipt Invoice"

'End if

saSelRcptNo = Request.QueryString("hRcptNo")
'Response.Write "<p> saSelRcptNo = " &  saSelRcptNo

if instr(1,saSelRcptNo,",") > 0 then
	sFlag = "Multiple"
else
	sFlag = "Single"
end if

if sFlag = "Multiple" then
	sarrSelRcptno = split(saSelRcptNo,",")
	'sSelRcptno = sarrSelRcptno(0)
	sSelRcptno = saSelRcptno
else
	sSelRcptno = saSelRcptno
end if
if trim(sSelRcptNo) = "" then sSelRcptNo = "0"

'To find Purchase Order No  & Take currency,Mode of Desp. etc from it
sConfNo = ""
sPoNo = ""
sSql = "Select isNull(OCNumber,0) From PUR_T_RefferenceNumberDet Where ReceiptNumber in ( " & sSelRcptNo  &" ) "

'Response.Write sSql
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source =  sSql
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = nothing
If not rsTemp.EOF Then
	If trim(rsTemp(0)) <> ""  and  trim(rsTemp(0)) <> "0"  Then
		sConfNo =  trim(sConfNo) & rsTemp(0) & ","
	end If 'If trim(rsTemp(0)) <> "" Then
End If 'If not rsTemp.EOF Then
rsTemp.Close
if trim(sConfNo) <> ""  and  trim(sConfNo) <> "0" then
	sConfNo = mid(sConfNo,1,len(sConfNo) - 1 )
	sConfNo = replace(sConfNo ,",","','")
	sConfNo = "'" &  sConfNo  & "'"
end if
 'Response.Write "<p> sConfNo = " & sConfNo

'Response.Write("Conf:"&sSelRcptNo)
If trim(sConfNo) <> ""  and  trim(sConfNo) <> "0" Then
	sSql = "Select isNull(PurchaseOrderNo,'') From tpoTrPOQuote Where ConfirmationNo in (" & sConfNo &" )"
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source =  sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing
	If not rsTemp.EOF Then
		If trim(rsTemp(0)) <> "" Then
			sPoNo = sPoNo & rsTemp(0) & ","
		end If 'If trim(rsTemp(0)) <> "" Then
	End If 'If not rsTemp.EOF Then
	rsTemp.Close

	if trim(sPoNo) <> "" then
		sPoNo = mid(sPoNo,1,len(sPoNo) - 1 )
		sPoNo = replace(sPoNo ,",","','")
		sPoNo = "'" &  sPoNo  & "'"
	end if


	If trim(sPoNo)<> "" then
		sSql = "Select IsNull(CurrencyCode,0),isNull(ModeOfDespatch,0),isNull(ModeOfPayment,0),isNull(IssueBank,0),isNull(PaymentTermsNo,0),isNull(BasisOfPrice,0),isNull(TransporterCode,0) From tpoTrPOrderHdr Where PurchaseOrderNo in ( " & sPoNo  &" )"
		With rsTemp
			.CursorLocation = 3
			.CursorType = 3
			.Source =  sSql
			.ActiveConnection = con
			.Open
		End With
		Set rsTemp.ActiveConnection = nothing
		If not rsTemp.EOF Then
			Curr1		= rsTemp(0)
			Mod1		= rsTemp(1)
			Mop			= rsTemp(2)
			IssueBank	= rsTemp(3)
			PayTerm		= rsTemp(4)
			Bop			= rsTemp(5)
			Transporter	= rsTemp(6)
		End If 'If not rsTemp.EOF Then
		rsTemp.Close
	End If 'If trim(sPoNo)<> "" then
End If 'If trim(sConfNo) <> ""  Then


''to generate the receiptcode inorder to display the Rcpt code in a span
if Trim(saSelRcptNo)<>"" then
    sSql ="SELECT ReceiptNumber,ReceiptCode,convert(varchar,ReceiptDate,103) FROM RCV_T_ActualReceiptHeader where GRNNumber in " &_
	    " (Select GRNNumber from RCV_T_GateReceiptHeader where ReceivedForUnit='"& sOrgID& "') and ReceiptNumber in (" & saSelRcptNo & ")"
    'Response.Write "<p>" & sSql
    With rsTemp
	    .CursorLocation = 3
	    .CursorType = 3
	    .Source =  sSql
	    .ActiveConnection = con
	    .Open
    End With
    'Set rsTemp.ActiveConnection = nothing
    If 	not rsTemp.EOF then
	    iRcptNo	= rsTemp(0)
	    sRcptCode = rsTemp(1)
	    sRcptDt  = rsTemp(2)


	    sTempRcptCode = ""
	    Do while not rsTemp.eof
		    sActualReceiptNos = trim(sRcptCode)&"--"&trim(sRcptDt)
		    rsTemp.MoveNext
		    sTempRcptCode = sTempRcptCode + sActualReceiptNos + ","
	    Loop
    End if
    rsTemp.Close
end if 'if Trim(saSelRcptNo)<>"" then
'----------------------------------------------------------------------------

nPurType = 0
sSql = "Select isNull(PurchaseType,0),isNull(CurrencyCode,0),isNull(PaymentTerms,0),isNull(BasisOfPricing,0),isNull(DespatchMode,0),isNull(PaymentMode,0),isNull(TransporterCode,0),isNull(DestinationPort,0),isNull(LoadingPort,0),isNull(IssueBank,0),isNull(BenifBank,0),isNull(Remarks,''),isNull(SuppInvoiceNo,''),SuppInvoiceDate,ISNULL(InvCategoryCode,'0'),IsNull(BillType,'0') from RCV_T_InvoiceHeader Where InvoiceNumber =" & nInvNo  &" "
'Response.Write sSql
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source =  sSql
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = nothing
If not rsTemp.EOF Then
	nPurType		= rsTemp(0)
	nCurrCode		= rsTemp(1)
	PayTerm			= rsTemp(2)
	Bop				= rsTemp(3)
	Mod1			= rsTemp(4)
	Mop				= rsTemp(5)
	Transporter		= rsTemp(6)
	nDestPort		= rsTemp(7)
	nLoadPort		= rsTemp(8)
	IssueBank		= rsTemp(9)
	nBenifitBank	= rsTemp(10)
	sRemarks		= rsTemp(11)
	sSuppInvNo		= rsTemp(12)
	sSuppInvDt		= formatdate(rsTemp(13))
	sCategoryCode	= rsTemp(14)
	sBillType		= rsTemp(15)
End If 'If not rsTemp.EOF Then
rsTemp.Close
'Response.Write "Bill:"	&sBillType
Set Root = oDOM.createElement("Root")
oDOM.appendChild Root

If objfs.FileExists(Server.MapPath("../xmldata/General.xml")) then
	objDOM2.Load server.MapPath("../xmldata/General.xml")
	Set sAccroot = objdom2.documentElement
	sExp = "//ROUNDOFF"
	Set sTempnode = sAccroot.Selectnodes(sExp)
	If sTempnode.Length > 0 then
		sRoundoffHead = sTempnode.Item(0).Attributes.Item(0).nodevalue
	else
		sRoundoffHead = "0"
	end if
else
	sRoundoffHead = "0"
End if

Dim iParTypeID,sParTypeName,sParType,sTraUnit,sRecptAg
sRefNum = sSelRcptNo
''To fetch Party name
''To fetch Supp Inv No. & dt.
'sSql = "Select isNull(InvoiceNumber,'0'),InvoiceDate,PartyCode,PartyType,ReceivedFrom,TRANSFEREDFROM,RECEIPTAGAINST from RCV_T_GateReceiptHeader where ReceivedForUnit='"& sOrgID& "' and " &_
'		" GRNNumber = (Select GRNNumber from RCV_T_ActualReceiptHeader where ReceiptNumber in ( "& sRefNum & " ) )"
sSql = "Select isNull(InvoiceNumber,'0'),InvoiceDate,PartyCode,PartyType,ReceivedFrom,TRANSFEREDFROM,'' from RCV_T_GateReceiptHeader where ReceivedForUnit='"& sOrgID& "' and " &_
		" GRNNumber in (Select GRNNumber from RCV_T_ActualReceiptHeader where ReceiptNumber in ( "& sRefNum & " ) )"
'Response.Write "<p> sSql = "& sSql
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = nothing
if not rsTemp.EOF then
	'sSuppInvNo = rsTemp(0)
	'sSuppInvDt = formatdate(rsTemp(1))
	iPartyCode = rsTemp(2)
	sPartyType = rsTemp(3)
	sPartySubType = rsTemp(4)
	sTraUnit = rsTemp(5)
	sRecptAg = rsTemp(6)
end if
rsTemp.Close

if Trim(iPartyCode)="" then
    sSql = "Select PartyCode,PartyType,PartySubType from RCV_T_InvoiceHeader where InvoiceNumber = "& nInvNo
    rsTemp.Open sSql,con
    if not rsTemp.EOF then
        iPartyCode = rsTemp(0)
        sPartyType = rsTemp(1)
        sPartySubType = rsTemp(2)
    end if
    rsTemp.Close 
end if

if trim(iPartyCode) = "" then iPartyCode = 0
sSql = "Select PartyName from App_M_PartyMaster where PartyCode=" & trim(iPartyCode) & ""
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = nothing
if not rsTemp.EOF then
	sPartyName = rsTemp(0)
End if
rsTemp.Close

	Set SubNode= oDOM.createElement("InvoiceHeader")
	Root.appendChild SubNode
	Set SubNode1=oDOM.createElement("ItemDetails")
	Root.appendChild SubNode1
	Set NewElem1 = oDOM.createElement("Header")
	newElem1.setAttribute "OrgID", sOrgID
	newElem1.setAttribute "Party", sPartyName
	newElem1.setAttribute "PurchaseType", nPurType
	newElem1.setAttribute "Currency", nCurrCode
	newElem1.setAttribute "InvAgainst", "Receipt"
	newElem1.setAttribute "RefNum", sRefNum
	newElem1.setAttribute "PartyCode", iPartyCode
	newElem1.setAttribute "PartyType",sPartyType
	newElem1.setAttribute "PartySubType",sPartySubType
	newElem1.setAttribute "CurrencyNo", nCurrCode
	newElem1.setAttribute "DespatchMode",Mod1
	newElem1.setAttribute "PaymentMode",Mop
	newElem1.setAttribute "PayTerms",PayTerm
	newElem1.setAttribute "IssueBank",IssueBank
	newElem1.setAttribute "BenificiaryBank",nBenifitBank
	newElem1.setAttribute "PricingBasis",Bop
	newElem1.setAttribute "Transporter",Transporter
	newElem1.setAttribute "LoadingPort",nLoadPort
	newElem1.setAttribute "DestPort",nDestPort
	newElem1.setAttribute "Remarks",sRemarks
	newElem1.setAttribute "SuppInvNo",sSuppInvNo
	newElem1.setAttribute "SuppInvDt",sSuppInvDt
	newElem1.setAttribute "TransporterFlag",""
	newElem1.setAttribute "PoNo",sPoNo
	newElem1.setAttribute "ConfNum",sConfNo
	newelem1.setAttribute "InvoiceFlag",sFlag
	newelem1.setAttribute "InvValue",0
	newelem1.setAttribute "RoundOff",0
	newelem1.setAttribute "InvoiceNumber",nInvNo

	cSuppCode=""
	'added by kalai selvi on 20/12/2005
	'finding supplier code related to selected party
	' to pass supplier code in Item Detail Popup

	'ith rsTemp
	'.ActiveConnection = con
	'.CursorLocation = 3
	'.CursorType = 3
	'.Source  = "select isNull(SupplierCode,'') from MAP_Supplier where OUDefinitionID ='" & sOrgID & "' and  PartyType = '" & sPartyType  & "' and PartySubType= " & sPartySubType & " and PartyCode = " & iPartyCode
	'.Open
	'end with
	'Response.Write " <p> " & rsTemp.Source
	'if not rsTemp.EOF then
	'	cSuppCode = rsTemp(0)
	'end if 'if not rsTemp.EOF
	'rsTemp.Close
	'Response.Write "<p> cSuppCode = " & cSuppCode

	newelem1.setAttribute "SuppCode",cSuppCode
	SubNode.appendChild NewElem1

	'==============================================================================
	''Added by Maheshwari on Aug 30 2007 for With Material case 'To get ItemDesc from AdditionalDescr field if itemcode is 0

Dim sAttributeList,sOptName,iOptVal,sTemp,i,rsAtt
set rsAtt = Server.CreateObject("ADODB.Recordset")
	
		sSql =	" SELECT DISTINCT AR.ClassificationCode,AR.ItemCode,AR.InvoiceQuantity,IC.GroupName,AR.AdditionalDescr, "&_
				" AR.EntryNo,AR.InvItemUOM,isNull(AR.InvoiceRate,0),isNull(AR.ItemDiscountPercent,0),isNull(AR.ItemDiscountValue,0),"&_
				" isNull(AR.ItemNettBasicValue,0), isNull(AR.PurchaseType,0),isNull(AR.ItemRate,0),isNull(AR.ItemValue,0),isNull(AR.RateUOM,''),"&_
				" ISNULL(VATELIGIBILITY,'N'),ISNULL(ITEMATTRIBUTES,'') FROM RCV_T_InvoiceDetails AR, INV_M_CLASSIFICATION IC WHERE AR.OrganisationCode='" & sOrgID & "' AND "&_
				" AR.InvoiceNumber in ( "& nInvNo & " ) AND AR.ClassificationCode = IC.GroupCode group by AR.ClassificationCode,AR.ItemCode,AR.InvoiceQuantity,IC.GroupName,"&_
				" AR.AdditionalDescr,AR.EntryNo,AR.InvItemUOM,isNull(AR.InvoiceRate,0),isNull(AR.ItemDiscountPercent,0),isNull(AR.ItemDiscountValue,0),"&_
				" isNull(AR.ItemNettBasicValue,0), isNull(AR.PurchaseType,0),isNull(AR.ItemRate,0),isNull(AR.ItemValue,0),isNull(AR.RateUOM,''),ISNULL(VATELIGIBILITY,'N'),ISNULL(ITEMATTRIBUTES,'') "
	
		'Response.Write ssql
		 With rsItem
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		End With

		Set rsItem.ActiveConnection = nothing

		iEntryNo = 0

		Do While Not rsItem.EOF

			iEntryNo = iEntryNo + 1

			iClassCode		= rsItem(0)
			iItemCode		= rsItem(1)
			iQtyRecd		= rsItem(2)
			sClassDesc		= rsItem(3)
			sItemDesc		= rsItem(4)
			'iEntryNo		= rsItem(5)
			UomCode			= rsItem(6)
			nRate			= rsItem(7)
			nDisPer			= rsItem(8)
			nDisAmt			= rsItem(9)
			nNetBasicForItem= rsItem(10)
			iPurTypeForItem	= rsItem(11)
			nRatePerQtyUOM	= rsItem(12)
			nItemValue		= rsItem(13)
			sRateUOM		= rsItem(14)
			sVAT			= rsItem(15)
			sAttributeList	= rsItem(16)
			If trim(sAttributeList) <> "" then
				sOptName = ""
				sTemp = split(sAttributeList,",")
				For i = 0 to UBOUND(sTemp)
					iOptVal = sTemp(i)
					sSql = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iOptVal&" "
					rsAtt.Open sSql,con
					If not rsAtt.EOF then
						sOptName =sOptName &","& rsAtt(0)
					End If
					rsAtt.Close
				Next
			End If
			IF sOptName <> "" then
				sOptName = " [" & mid(sOptName,2) &"] "
			End IF

			sStockType = ""
			
			
		    sSql = "Select ItemDescription from VWITEM where ITemCode = "& iItemCode
		    dcrs.Open sSql,con
		    if not dcrs.EOF then
		        sItemDesc = dcrs(0)
		    end if
		    dcrs.Close 
			

			sSql = "Select isNull(ReceiptNumber,0) from PUR_T_RefferenceNumberDet where isNull(InvoiceNumber,0) = " & nInvNo & ""
			With dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			End With
			Set dcrs.ActiveConnection = nothing

			do while not dcrs.EOF


				sSql = "Select isNull(StockType,'') from RCV_T_ActualRcptItemDet where isNull(ReceiptNumber,0) = " & dcrs(0) & " " & _
						" AND ClassificationCode = " & iClassCode & " AND ITEMCODE = " & iItemCode  & " " &_
						" AND OrganisationCode='" & sOrgID & "'"
				'	Response.Write sSql
				With rsTemp
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				End With
				Set rsTemp.ActiveConnection = nothing

				if not rsTemp.EOF then
					sStockType = trim(sStockType) + "," + trim(rsTemp(0))
				end if
				rsTemp.Close

				dcrs.MoveNext
			loop
			dcrs.Close

			if trim(sStockType) <> "" then
				sStockType = mid(sStockType,2)
			end if

			dBalQty = round(cdbl(iQtyRecd) - cdbl(dInvQty),3)

			if trim(sOptName) <> "" then sItemDesc = sItemDesc & sOptName
			'if cdbl(dBalQty) > 0 then
			
			
			if trim(sItemDesc)="" or IsNull(sItemDesc) then sItemDesc=""



				Set newElem1 = oDOM.createElement("Item")
				newElem1.setAttribute "ItemCode", iItemCode
				newElem1.setAttribute "ClassificationCode", iClassCode
				newElem1.setAttribute "ItmDescription", sItemDesc
				newElem1.setAttribute "Uom", UomCode
				newElem1.setAttribute "Qty", dBalQty
				newElem1.setAttribute "Rate", nRate
				newElem1.setAttribute "DisPer", nDisPer
				newElem1.setAttribute "DisAmount", nDisAmt
				newElem1.setAttribute "NettBasic", nNetBasicForItem
				newElem1.setAttribute "UomDesc", UomCode
				newElem1.setAttribute "EntryNo", iEntryNo
				newElem1.setAttribute "RatePerQtyUoM", nRate
				newElem1.setAttribute "SourceEntryNo", iEntryNo
				newElem1.setAttribute "PurchaseType", iPurTypeForItem
				newElem1.setAttribute "Amount", "0"
				newElem1.setAttribute "ItemValue", nItemValue
				newElem1.setAttribute "ItemRate", nRatePerQtyUOM
				newElem1.setAttribute "RateUOM", sRateUOM
				newElem1.setAttribute "StockType", sStockType
				newElem1.setAttribute "VAT", sVAT
				newElem1.setAttribute "AttributeList",sAttributeList


				SubNode1.appendChild NewElem1
			'End if
			rsItem.MoveNext

		Loop

		rsItem.Close
		'Root.appendChild SubNode1

		'''''To get New Items''''''''''
		sSql = "SELECT A.TempItemCode,A.InvoiceQuantity,B.ItemDescription,'',A.InvItemUOM,isNull(A.InvoiceRate,0)," & _
			" isNull(A.ItemDiscountPercent,0),isNull(A.ItemDiscountValue,0),isNull(A.ItemNettBasicValue,0)," & _
			" isNull(A.PurchaseType,0),isNull(A.ItemRate,0),isNull(A.ItemValue,0),isNull(A.RateUOM,''),VATELIGIBILITY,Isnull(A.ItemAttributes,'') FROM RCV_T_InvoiceDetails A,MS_TemporaryItemMaster B WHERE " &_
			" A.OrganisationCode='" & trim(sOrgID) & "' AND A.InvoiceNumber = " & nInvNo & " and A.TempItemCode = B.TempItemCode" &_
			" and A.TempItemCode is not null and A.TempItemCode <> '0'"
		'Response.Write ssql
		With rsItem
			.CursorLocation = 3
			.CursorType = 3
			.Source =  sSql
			.ActiveConnection = con
			.Open
		End With

		Set rsItem.ActiveConnection = nothing

		iClassCode = "TEMP"

		Do While Not rsItem.EOF
			iItemCode		= rsItem(0)
		    iQtyRecd		= rsItem(1)
			sItemDesc		= rsItem(2)
			UomCode			= rsItem(4)
			nRate			= rsItem(5)
			nDisPer			= rsItem(6)
			nDisAmt			= rsItem(7)
			nNetBasicForItem= rsItem(8)
			iPurTypeForItem	= rsItem(9)
			nRatePerQtyUOM	= rsItem(10)
			nItemValue		= rsItem(11)
			sRateUOM		= rsItem(12)
			sVAT			= rsItem(13)
			sAttributeList  = rsItem(14)
			If trim(sAttributeList) <> "" then
				sOptName = ""
				sTemp = split(sAttributeList,",")
				For i = 0 to UBOUND(sTemp)
					iOptVal = sTemp(i)
					sSql = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iOptVal&" "
					rsAtt.Open sSql,con
					If not rsAtt.EOF then
						sOptName =sOptName &","& rsAtt(0)
					End If
					rsAtt.Close
				Next
			End If
			IF sOptName <> "" then
				sOptName = " [" & mid(sOptName,2) &"] "
			End IF
			sStockType = ""

			sSql = "Select isNull(ReceiptNumber,0) from PUR_T_RefferenceNumberDet where isNull(InvoiceNumber,0) = " & nInvNo & ""
			With dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			End With
			Set dcrs.ActiveConnection = nothing

			do while not dcrs.EOF


				sSql = "Select isNull(StockType,'') from RCV_T_ActualRcptItemDet where isNull(ReceiptNumber,0) = " & dcrs(0) & " " & _
						" AND TempItemCode = " & iItemCode  & " AND OrganisationCode='" & sOrgID & "'"
				With rsTemp
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				End With
				Set rsTemp.ActiveConnection = nothing

				if not rsTemp.EOF then
					sStockType = trim(sStockType) + "," + trim(rsTemp(0))
				end if
				rsTemp.Close

				dcrs.MoveNext
			loop
			dcrs.Close

			if trim(sStockType) <> "" then
				sStockType = mid(sStockType,2)
			end if

			'' to add Qty Validation for Receipt
			sSql = "Select Sum(B.InvoiceQuantity) from Rcv_T_InvoiceHeader A, Rcv_T_InvoiceDetails B " &_
					" Where isnull(TempItemCode,0)=" & trim(iItemCode) & "  and A.InvoiceAgainst = 1 and A.ReferenceNumber= '" & trim(sRefNum) & "'" &_
					" and B.InvoiceNumber=A.InvoiceNumber group by TempItemCode "

			With dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			End With
			Set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				dInvQty = dcrs(0)
			else
				dInvQty = 0
			end if
			dcrs.Close

			dBalQty = cdbl(iQtyRecd) - cdbl(dInvQty)

			iEntryNo = iEntryNo + 1

			if trim(sOptName) <> "" then sItemDesc = sItemDesc & sOptName

			Set newElem1 = oDOM.createElement("Item")
			newElem1.setAttribute "ItemCode", iItemCode
			newElem1.setAttribute "ClassificationCode", "TEMP"
			newElem1.setAttribute "ItmDescription", sItemDesc
			newElem1.setAttribute "Uom", UomCode
			newElem1.setAttribute "Qty", dBalQty
			newElem1.setAttribute "Rate", nRate
			newElem1.setAttribute "DisPer", nDisPer
			newElem1.setAttribute "DisAmount", nDisAmt
			newElem1.setAttribute "NettBasic", nNetBasicForItem
			newElem1.setAttribute "UomDesc", UomCode
			newElem1.setAttribute "EntryNo", iEntryNo
			newElem1.setAttribute "RatePerQtyUoM", nRate
			newElem1.setAttribute "SourceEntryNo", iEntryNo
			newElem1.setAttribute "PurchaseType", iPurTypeForItem
			newElem1.setAttribute "Amount", "0"
			newElem1.setAttribute "ItemValue", nItemValue
			newElem1.setAttribute "ItemRate", nRatePerQtyUOM
			newElem1.setAttribute "RateUOM", sRateUOM
			newElem1.setAttribute "StockType", sStockType
			newElem1.setAttribute "VAT", sVAT
			newElem1.setAttribute "AttributeList",sAttributeList
			SubNode1.appendChild newElem1
			rsItem.MoveNext
		Loop




		rsItem.Close

		if trim(nPurType) <> "0" then
			Set oNodTaxRoot = oDOM.createElement("TaxDetails")
			oNodTaxRoot.setAttribute "Basicvalue","0"
			oNodTaxRoot.setAttribute "NettValue","0"
			oNodTaxRoot.setAttribute "TotalTax","0"
			oNodTaxRoot.setAttribute "SubTotal","0"
			oNodTaxRoot.setAttribute "PurchaseType",nPurType
			Root.appendChild oNodTaxRoot



			sSql = "Select TaxShortName,TaxCategoryCode,TaxCode,ComputationMode,isnull(SumOfFields,''),isnull(FlatAmount,0),"&_
				" AccountHead,TaxCreditEligibility,isnull(RegisterCode,0),AccountTaxAccHead,isnull(Roundoff,'0'),TaxCategoryType, isnull(orgTaxAccHdNo,0) from VwPurchaseTaxDetails where OUDefinitionID='"&sOrgID&"' and PurchaseType="&nPurType& " order by TaxHierarchy"
			'Response.Write "<p>sSql ="+ sSql
			With rsTemp
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			End with
			Set rsTemp.ActiveConnection = nothing

			Set sTaxName=rsTemp(0)
			Set sCatCode=rsTemp(1)
			Set sTaxCode=rsTemp(2)
			Set sTaxMode=rsTemp(3)
			Set sFormula=rsTemp(4)
			Set dTaxValue=rsTemp(5)
			Set sAccHead=rsTemp(6)
			'Set iCrEligible=rsTemp(7)
			Set iRegCode=rsTemp(8)
			Set iToBeAccounted = rsTemp(9)
			Set iTRndoff = rsTemp(10)
			Set sTaxCatType  = rsTemp(11)
			Set sOrghdno = rsTemp(12)

			Do while not rsTemp.EOF


				'' to store tax details from invoice tax details table
				sSql = "Select isNull(TaxCreditEligibility,''),isNull(TaxPercentage,0),isNull(TaxAmount,0),isNull(PurchaseType,0) from RCV_T_InvoiceTaxDetails " &_
						" Where InvoiceNumber=" & trim(nInvNo) & "  and TaxCategoryCode =  " & sCatCode & " and TaxCode = " & sTaxCode  & ""

				With dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				End With
				Set dcrs.ActiveConnection = nothing

				if not dcrs.EOF then
					iCrEligible = dcrs(0)
					dTaxPer		= dcrs(1)
					dTaxValue	= dcrs(2)
				else
					iCrEligible = "0"
					dTaxPer = 0
					dTaxValue = 0
				end if
				dcrs.Close

				Set newElem = oDOM.createElement("Tax")
				newElem.setAttribute "CatCode",sCatCode
				newElem.setAttribute "TaxCode",sTaxCode
				newElem.setAttribute "TaxMode",sTaxMode
				newElem.setAttribute "TaxFormula",sFormula
				if trim(sTaxMode) ="F" then
					newElem.setAttribute "TaxValue",CStr(dTaxValue)
				else
					newElem.setAttribute "TaxValue",CStr(dTaxPer)
				end if
				newElem.setAttribute "TaxAmount",CStr(dTaxValue)
				newElem.setAttribute "AccHead",sAccHead
				newElem.setAttribute "CrEligible",iCrEligible
				newElem.setAttribute "RegisterCode",iRegCode
				newElem.setAttribute "ToBeAccounted",iToBeAccounted
				newElem.SetAttribute "Rndoff",iTRndoff
				newElem.SetAttribute "TaxCategoryType",sTaxCatType
				newElem.Text= sTaxName
				oNodTaxRoot.appendChild newElem
				rsTemp.MoveNext
			Loop
			rsTemp.Close
		else


			for each ItemNode in SubNode1.ChildNodes

				if trim(ItemNode.getAttribute("PurchaseType")) <> "0" then

					'check Tax Details for the purchase type is already exist,
					' if so add current item value along with that otherwise create a new Tax Detail Node
					sTaxDetailsAlreadyExist = "N"

					for each newElem in Root.ChildNodes
						if trim(NewElem.NodeName) = "TaxDetails" then
							if trim(NewElem.getAttribute("PurchaseType"))  = trim(ItemNode.getAttribute("PurchaseType")) then
								sTaxDetailsAlreadyExist = "Y"
								Set oNodTaxRoot = NewElem
								exit for
							end if
						end if
					next

					nTotalBasic		= 0
					nTotalNet		= 0
					nTotalTax		= 0
					nTotalSubTotal  = 0

					nTotalBasic		= ItemNode.getAttribute("Qty") * ItemNode.getAttribute("Rate")
					nTotalNet		= ItemNode.getAttribute("NettBasic")

					if sTaxDetailsAlreadyExist = "N" then
						Set oNodTaxRoot = oDOM.createElement("TaxDetails")
						Root.appendChild oNodTaxRoot
					end if

					if sTaxDetailsAlreadyExist = "N" then
						sSql = "Select TaxShortName,TaxCategoryCode,TaxCode,ComputationMode,isnull(SumOfFields,''),isnull(FlatAmount,0),"&_
							" AccountHead,TaxCreditEligibility,isnull(RegisterCode,0),AccountTaxAccHead,isnull(Roundoff,'0'),TaxCategoryType, isnull(orgTaxAccHdNo,0) from VwPurchaseTaxDetails where OUDefinitionID='"&sOrgID&"' and PurchaseType="& ItemNode.getAttribute("PurchaseType") & " order by TaxHierarchy"
						'Response.Write "<p>sSql ="+ sSql
						With rsTemp
							.CursorLocation = 3
							.CursorType = 3
							.Source = sSql
							.ActiveConnection = con
							.Open
						End with
						Set rsTemp.ActiveConnection = nothing

						Set sTaxName=rsTemp(0)
						Set sCatCode=rsTemp(1)
						Set sTaxCode=rsTemp(2)
						Set sTaxMode=rsTemp(3)
						Set sFormula=rsTemp(4)
						Set sAccHead=rsTemp(6)
						'Set iCrEligible=rsTemp(7)
						Set iRegCode=rsTemp(8)
						Set iToBeAccounted = rsTemp(9)
						Set iTRndoff = rsTemp(10)
						Set sTaxCatType  = rsTemp(11)
						Set sOrghdno = rsTemp(12)

						Do while not rsTemp.EOF

							iCrEligible = "0"
							dTaxPer = 0
							dTaxValue=rsTemp(5)

							'' to store tax details from invoice tax details table
							sSql = "Select isNull(TaxCreditEligibility,''),isNull(TaxPercentage,0),isNull(TaxAmount,0),isNull(PurchaseType,0) from RCV_T_InvoiceTaxDetails " &_
									" Where InvoiceNumber=" & trim(nInvNo) & "  and TaxCategoryCode =  " & sCatCode & " and TaxCode = " & sTaxCode  & "  and PurchaseType="& ItemNode.getAttribute("PurchaseType") & " "
							'Response.Write "<p> <p> sSql = "& sSql
							With dcrs
								.CursorLocation = 3
								.CursorType = 3
								.Source = sSql
								.ActiveConnection = con
								.Open
							End With
							Set dcrs.ActiveConnection = nothing
							'Response.Write "<p> <p> dcrs.EOF = "& dcrs.EOF
							if not dcrs.EOF then
								iCrEligible = dcrs(0)
								dTaxPer		= dcrs(1)
								dTaxValue	= dcrs(2)
							end if
							dcrs.Close

							Set newElem = oDOM.createElement("Tax")
							newElem.setAttribute "CatCode",sCatCode
							newElem.setAttribute "TaxCode",sTaxCode
							newElem.setAttribute "TaxMode",sTaxMode
							newElem.setAttribute "TaxFormula",sFormula
							if trim(sTaxMode) ="F" then
								newElem.setAttribute "TaxValue",CStr(dTaxValue)
							else
								newElem.setAttribute "TaxValue",CStr(dTaxPer)
							end if
							newElem.setAttribute "TaxAmount",CStr(dTaxValue)
							newElem.setAttribute "AccHead",sAccHead
							newElem.setAttribute "CrEligible",iCrEligible
							newElem.setAttribute "RegisterCode",iRegCode
							newElem.setAttribute "ToBeAccounted",iToBeAccounted
							newElem.SetAttribute "Rndoff",iTRndoff
							newElem.SetAttribute "TaxCategoryType",sTaxCatType
							newElem.Text= sTaxName
							oNodTaxRoot.appendChild newElem

							nTotalTax = cdbl(nTotalTax) + cdbl(dTaxValue)
							rsTemp.MoveNext
						Loop
						rsTemp.Close
					end if 'if sTaxDetailsAlreadyExist = "N" then

					nTotalSubTotal = cdbl(nTotalNet) +  cdbl(nTotalTax)

					'storing tax details node details
					if sTaxDetailsAlreadyExist = "N" then
						oNodTaxRoot.setAttribute "Basicvalue",nTotalBasic
						oNodTaxRoot.setAttribute "NettValue",nTotalNet
						oNodTaxRoot.setAttribute "TotalTax",nTotalTax
						oNodTaxRoot.setAttribute "SubTotal",nTotalSubTotal
						oNodTaxRoot.setAttribute "PurchaseType",ItemNode.getAttribute("PurchaseType")
					else
						nTotalBasic		= cdbl(nTotalBasic)		+ cdbl(oNodTaxRoot.getAttribute("Basicvalue"))
						nTotalNet		= cdbl(nTotalNet)		+ cdbl(oNodTaxRoot.getAttribute("NettValue"))
						nTotalSubTotal	= cdbl(nTotalSubTotal)	+ cdbl(oNodTaxRoot.getAttribute("SubTotal"))

						oNodTaxRoot.setAttribute "Basicvalue",nTotalBasic
						oNodTaxRoot.setAttribute "NettValue",nTotalNet
						oNodTaxRoot.setAttribute "SubTotal",nTotalSubTotal
					end if

				end if 'if trim(ItemNode.getAttribute("PurchaseType")) <> "0" then
			next 'for each ItemNode in SubNode1.ChildNodes


		end if 'end if 'if trim(nPurType) <> "0" then


		'*******************************************************************************************************
		oDOM.save server.MapPath("..\temp\transaction\AmdNewInvItemValue_PUR_"&Session.SessionID&".xml")

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="init('<%=sSelRcptNo%>','<%=sOrgID%>');SetDate()">

<form method="POST" name="formname" action>
	
	<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
	<Input type="hidden" name="hPoNo" value="<%=sPoNo%>" >
	<input type="hidden" name="hPartyCode" value="<%=iPartyCode%>">
	<input type="hidden" name="hCurrDate" value="<%=FormatDate(Date)%>">
	<input type="hidden" name="hInvDate" value="<%=sSuppInvDt%>">
	<input type="hidden" name="hMaxDate" value="<%=sMaxDate%>">
	<input type="hidden" name="hMinDate" value="<%=sMinDate%>">
	<input type="hidden" name="hPartyType" value="<%=sPartyType%>">
	<input type="hidden" name="hPSubType" value="<%=sPartySubType%>">
	<input type="hidden" name="txtSuppInvDt" value="">
	<Input type="hidden" name="hConfNum" value="<%=sConfNo%>" >
	<input type="hidden" name="hFlag" value="<%=sFlag%>">
	<input type="hidden" name="hRcptno" value="<%=saSelRcptNo%>">
	<input type="hidden" name="hRcptCode" value="<%=sTempRcptCode%>">
	<input type="hidden" name="txtPartyName" value="">
	<input type="hidden" name="hRoundOffAccHead" value="<%=sRoundoffHead%>">
	<input type="hidden" name="hRcptNum" value="<%=sSelRcptNo%>">
	<input type="hidden" name="hInvNo" value="<%=nInvNo%>">

	<input type="hidden" name="hSuppInvDate" value="<%=sSuppInvDt %>" >
	<input type=hidden name="hRecNo" value="">
	<input type=hidden name="hRefType" value="">
	<input type=hidden name="hRefDate" value="">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Purchase Invoice Entry Amendment
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0">
											<tr>
											    <td class="FieldCell" valign="top">
											    <% if Trim(saSelRcptNo)<>"" then
											        response.Write "Against Receipt"
											      else
											        response.Write "Party Name"
											      end if
											    %>
												</td>
												<td class="FieldCellSub" colspan="4">
												    <%if Trim(saSelRcptNo)<>"" then %>
													<span id=SpnRcptCode class="Dataonly"><%=saSelRcptNo%></span>&nbsp;-
													<span class="Dataonly" id="spnSuppName"><%=sPartyName%>&nbsp;</span>
													<%else %>
													<span id=Span1 class="Dataonly"></span>
													<span class="Dataonly" id="Span2"><%=sPartyName%>&nbsp;</span>
													<%end if %>
												</td>
											</tr>

											<tr>
											    <td class="FieldCell" valign="top">Reference Name
												</td>
												<td class="FieldCellSub">
													<Select name="selRefName" class="FormElem" onChange="RefType_Click()">
													<%
													    RefTypePop 10,2
													%>
													</Select>
													&nbsp;&nbsp;
														<img src="../../assets/images/iTMS%20icons/Entryicon.gif" onclick="RefType_Click()">
												</td>
												<td class="FieldCell">Supplier Type
												</td>
												<td class="FieldCellSub" colspan="4">
												<select size="1" class="Formelem" name="cmbPartyType">
														<option value="0" selected>Select</option>
														<%populatePartyType( )%>
													</select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Supplier Inv. No. / Dt.
												</td>
												<td class="FieldCellSub" colspan="4">
													<table>
													<tr><td>
													<input type="text" name="txtSuppInvNo" size="30" class="FormElem" align=top>
													</td><td>
														<input type="date" id="ctlDate" name="ctlDate" onblur="MinDate()" class="formelem itms-date-picker" style="width:89px">
 													<% ' Function Call to Insert Date Picker
													'Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
													<td>
														<select size="1" class="Formelem" name="cmbBillType">
															<option value="0" selected>Select Bill type</option>
															<option value="P" <%If sBillType = "P" Then Response.write("selected")%>>Credit bill</option>
															<option value="C" <%If sBillType = "C" Then Response.write("selected")%>>Cash bill</option>
														</select>
													</td>
													</tr>
													</table>
												</td>
											</tr>

											<tr>
												<td class="FieldCell" valign="top">Purchase Type
												</td>
												<td class="FieldCellSub" colspan="4">
												<select size="1" class="FormElem" name="cmbPurType" onchange="getTaxDet()">
														<option  value="" selected>Select</option>
														<option  value="0" <%if trim(nPurType) = "0" then Response.Write "Selected" %>>--------------ITEMWISE-------------</option>
														<%

															popSelPurTypeFull(nPurType)
														%>
													</select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Currency
												</td>
												<td class="FieldCellSub"><select size="1" name="D29" class="Formelem">
														<%
								  						'populateCurrency
								  						popSelCurrency(nCurrCode)
								  						%>
													</select>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Preferences&nbsp&nbsp
												<a>
													<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" width="11" height="11" onclick="ShowPrefDet()" alt="Preferences" style="cursor:hand">
													</a>
												</td>
												<td class="FieldCellSub" align="left">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Category
												</td>
												<td class="FieldCellSub" colspan="3">
												<select size="1" name="cmbInvCat" class="FormElem" >
													<option value="0">Select</option>
													<%

													sSql = "Select CategoryCode,CategoryName From APP_M_InvoiceCategory Where ApplicableFor = 'P' Order By CategoryName "
													With dcrs
															.CursorLocation = 3
															.CursorType = 3
															.Source = sSql
															.ActiveConnection = Con
															.Open
														End With
														Set dcrs.ActiveConnection = Nothing
														Do While Not dcrs.EOF
															if trim(sCategoryCode) = trim(dcrs(0)) then
														%>
														<Option value="<%=dcrs(0)%>" selected><%=dcrs(1)%></Option>
														<%
															else
														%>
														<Option value="<%=dcrs(0)%>"><%=dcrs(1)%></Option>
														<%
															end if
														dcrs.MoveNext
														loop
														dcrs.Close
													%>


													</select>
												</td>
											</tr>

										</table>
									</td>
									<td>
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<div class="frmBody" id="frm2" style="width: 100%; height:242;">
											<table border="0" cellspacing="1" class="ExcelTable" width="610" id="tblItemDet">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
														<a href="#"><img name="ImgDeleteIcon" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onClick="DeleteItems()" alt="Delete Selected Item"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description
														<%if trim(nPurType) <> "0" then%>
															<a><img name=imgPurchaseDet border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" disabled=true alt="Enter Purchase Details" onclick="ShowPurchaseDet('<%=sOrgID%>','<%=sRefNum%>')" width="15" height="15" style="cursor:hand"></a>
														<%else%>
															<a><img name=imgPurchaseDet border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" alt="Enter Purchase Details" onclick="ShowPurchaseDet('<%=sOrgID%>','<%=sRefNum%>')" width="15" height="15" style="cursor:hand"></a>
														<%end if %>

													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Invoice<br>
														UoM
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Invoice<br>
														Quantity
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Invoice<br>
														Rate
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Basic<br>
														Value
													</td>
													<td class="ExcelHeaderCell" align="center" colspan="2">Discount
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Nett<br>
														Basic
													</td>
													<%if sFlag = "True" then%>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Org Supp Inv
													<a onClick="showInvoiceEntryPop()" href="#">
													<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" width="11" height="11" alt="Enter Invoice  Details"></a></td>
													<%end if%>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Sub Total
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">VAT
													<input type="Checkbox" name="ChkAll" class="Formelem" onClick="SelectAll()">
													</td>
													<!--td class="ExcelHeaderCell" align="center" rowspan="2">Purchase Type <a href="InvoiceTaxDetails.html" target="_flank">
														<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" width="11" height="11">
														</a>
													</td-->
												</tr>

												<tr>
													<td class="ExcelHeaderCell" align="center">%
													</td>
													<td class="ExcelHeaderCell" align="center">Value
													</td>
												</tr>
											</table>
										</div>
									</td>
									<td>
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td width="65" class="FieldCell" valign="top">Remarks
												</td>
												<td><textarea rows="2" name="mTextAreaRemarks" cols="93" class="FormElem"><%=sRemarks%></textarea>
												</td>
											</tr>

										</table>
									</td>
									<td>
									</td>
								</tr>

								<tr>
									<td colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td>
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
											        <!--<input type="Button" value="View Item Values"	name="BtnViewItem" onClick="ViewItemValues()" class="ActionButton" tabindex="3" style="width:120">-->
													<input type="Button" value="Next"				name="BtnNext" onclick="javascript: return Next_Click('S')" class="ActionButton" tabindex="3">
 													<input type="reset"  value="Cancel"				name="BtnCancel" class="ActionButton" tabindex="4">
 													<input type="reset"  value="Reset"				name="BtnReset" class="ActionButton" tabindex="4">
												</td>
											</tr>

										</table>
									</td>
									<td>
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td colspan="3" class="BottomPack">
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
</body>
</html>
<%'ref :invPurItemValueCalcVw%>
