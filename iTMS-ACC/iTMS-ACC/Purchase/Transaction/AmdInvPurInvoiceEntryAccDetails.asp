<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AmdInvPurInvoiceEntryAccDetails.asp
	'Module Name				:	Purchase (Transaction - Invoice Entry)
	'Author Name				:	Kalaiselvi R
	'Created On					:	January 24, 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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



<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/PurchaseTermsConditions.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/PurChkItemSpecPack.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="tex|/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">

<script type="application/xml" id="InvoiceDet" data-itms-xml-island="1" data-src="<%="../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml"%>"></script>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/calcAlternateUoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/amdInvPurInvoiceAccDetails.js"></SCRIPT>

<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
function DisplayTotal(nPassCtr)

Dim sAccHeads

dim Arr1,nArrCtr,nTotal,nTotalTaxValue,nTotalValue

sAccHeads = ""

nTotalTaxValue	 = document.formname.hTotalTaxValue.value
'nAdjustValue	 = document.formname.mAdjustValue.value

'alert("nTotalTaxValue = " + trim(nTotalTaxValue)  )
nTotalValue		 = nTotalTaxValue

set RootNode=Invoicedet.documentElement
	
set ItemNd=RootNode.SelectNodes("//ItemDetails/Item")

for i=1 to ItemNd.Length-1
	sItemAccHead = trim(eval("document.formname.mCmbItemAccHeadZ1.value"))
	set obj = eval("document.formname.mCmbItemAccHeadZ"& trim(i+1)) 
	for iCnt = 0 to cint(obj.length)
		if StrComp(obj(iCnt).value,sItemAccHead)=0 then
			obj.selectedIndex = iCnt
			exit for
		end if
	next
next		

sAccHeads = ""
for i=0 to ItemNd.Length-1

	if i+1  = nPassCtr then
		sItemAccHead = trim(eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) &".value"))
		if trim(sItemAccHead) = "0" then
			alert("Select Account Head")
			set obj = eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) )
			obj.focus()
			exit function 
		end if 'if trim(sItemAccHead) = "0" then
	end if 	'if i+1  = nPassCtr then
next	


for i=0 to ItemNd.Length-1

	sItemAccHead = trim(eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) &".value"))
	if trim(sItemAccHead) <> "0" then
		sItemAccHead = mid(sItemAccHead,1,instr(1,sItemAccHead,"|")-1 )
		if instr(1,sAccHeads,sItemAccHead) <= 0 then
			sAccHeads = sAccHeads  &  trim(sItemAccHead) & ","
		end if 	
	end if 'if trim(sItemAccHead) = "0" then

next	

if right(sAccHeads,1) ="," then
	sAccHeads = mid(sAccHeads,1,len(sAccHeads) -1 )
end if 

if trim(sAccHeads) <> "" then

	
	for i = 0 to ItemNd.Length-1
		set obj = eval("document.formname.ItemValueZ"+ trim(i+1) )
		obj.value = ""
	next
	
	Arr1 = Split(sAccHeads,",")
	
	for nArrCtr = LBound(Arr1) to UBound(Arr1)
	
		nTotal = 0 
		
		for i = 0 to ItemNd.Length-1
			sItemAccHead = trim(eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) &".value"))
			if trim(sItemAccHead) <> "0" then
				sItemAccHead = mid(sItemAccHead,1,instr(1,sItemAccHead,"|")-1 )
				if trim(sItemAccHead) = trim(Arr1(nArrCtr)) then
					nTotal = cdbl(nTotal) + cdbl(ItemNd.Item(i).getAttribute("ItemValue"))
				end if 
			end if 'if trim(sItemAccHead) <> "0" then	
		next
		
		set obj = eval("document.formname.ItemValueZ"+ trim(nArrCtr+1) )
		obj.value = Formatnumber(nTotal,2,,,0)
		
		if cdbl(nTotal) > 0 then
			nTotalValue = cdbl(nTotalValue) + cdbl(nTotal )
		end if 
	next
end if 

document.formname.mTotalValue.value = Formatnumber(nTotalValue,2,,,0)
end function
'-------------------------------------------------------------------------------------------
function Goto_AccItemValuePage()

Dim sAccHeadExist,sItemAccHead

Dim nPrevTotalAmt

Dim RootNode,ItemNd,AccHeadElem,ItemAccHead,AccDet,AccItemDet


set RootNode=Invoicedet.documentElement
	
set ItemNd=RootNode.SelectNodes("//ItemDetails/Item")

for i=0 to ItemNd.Length-1

	sItemAccHead = trim(eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) &".value"))
	if trim(sItemAccHead) = "0" then
		alert("Select Account Head")
		set obj = eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) )
		obj.focus()
		exit function 
	end if 'if trim(sItemAccHead) = "0" then
next	

sPrevAccHead = ""
' check all account head are same or not 
for i=0 to ItemNd.Length-1

	sItemAccHead = trim(eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) &".value"))
	if Trim(sPrevAccHead) = "" then sPrevAccHead = sItemAccHead
	
	if Trim(sPrevAccHead) <> Trim(sItemAccHead) then
		alert("All Items Account Head should be same")
		exit function 
	end if 'if trim(sItemAccHead) = "0" then
next		

sAccHeadExist = "N"	
for each Node1 in RootNode.childNodes
	if Node1.nodename = "Voucher" then
		for each Node2 in Node1.ChildNodes
			if Node2.nodename = "Details" then
				set ndDetails = Node2
			end if 
		Next	
	end if 'if Node1.nodename = "Voucher" then
	
	if Node1.nodename = "AccountHead" then
		sAccHeadExist = "Y"
		set AccHeadElem = Node1
	end if 
next
'' updating detail node - start
For each ndTemp in ndDetails.childNodes

	iItemCnt = ndTemp.getAttribute("No")
	iItemAccHead = ndTemp.childnodes(0).getAttribute("No")
	iItemCode = ndTemp.getAttribute("ItemCode")
	iClassCode = ndTemp.getAttribute("ClassCode")
	sItemDesc = ndTemp.getAttribute("ItemDesc")
	dItemValue  = ndTemp.getAttribute("ItemValue")

	sSelAccHeadStr = eval("document.formname.mCmbItemAccHeadZ"+trim(iItemCnt) ).value

	'alert(sSelAccHeadStr)
	
	saTemp = split(sSelAccHeadStr,"|")

	iSelAccHead = saTemp(0)

	if iSelAccHead <> iItemAccHead then
		iCCExist = saTemp(1)
		iAHExist = saTemp(2)

		ndTemp.removechild ndTemp.childnodes(0)

		Set newElem2  = Invoicedet.createElement("AccHead")
		newElem2.setAttribute "No",iSelAccHead
		newElem2.setAttribute "CostCenter",iCCExist
		newElem2.setAttribute "Analytical",iAHExist
		newElem2.setAttribute "Name",sItemDesc
		newElem2.setAttribute "Type","G"
		newElem2.setAttribute "Group",""

		ndTemp.appendChild newElem2
	end if

Next 'For each ndTemp in ndDetails.childNodes
'' updating detail node - end 

if sAccHeadExist = "N" then
	Set AccHeadElem = Invoicedet.CreateElement("AccountHead")
	RootNode.appendChild AccHeadElem
end if

Set ItemAccHead = Invoicedet.CreateElement("ItemAccountHead")
AccHeadElem.appendChild ItemAccHead


for i=0 to ItemNd.Length-1

	sItemAccHead = trim(eval("document.formname.mCmbItemAccHeadZ"+ trim(i+1) &".value"))
	sItemAccHead = mid(sItemAccHead,1,instr(1,sItemAccHead,"|")-1 )
	ItemNd.Item(i).setAttribute "ItemAccHead",sItemAccHead
	
	sAccHeadExist = "N"
	nPrevTotalAmt = 0
	
	for each Node1 in ItemAccHead.ChildNodes
		if trim(Node1.getAttribute("No")) = sItemAccHead then
			set AccDet = Node1
			nPrevTotalAmt = Node1.getAttribute("TotalAmt")
			sAccHeadExist = "Y"
		end if 
	next
	
	if sAccHeadExist = "N" then
		Set AccDet = Invoicedet.CreateElement("Acc")
		AccDet.SetAttribute "No",sItemAccHead
		AccDet.SetAttribute "TotalAmt",ItemNd.Item(i).getAttribute("ItemValue")
		ItemAccHead.appendChild AccDet 
	else
		AccDet.SetAttribute "TotalAmt",cdbl(nPrevTotalAmt) +  cdbl(ItemNd.Item(i).getAttribute("ItemValue"))
	end if 'sAccHeadExist = "N"
	
	Set AccItemDet = Invoicedet.CreateElement("Item")
	AccItemDet.setAttribute "Desc",ItemNd.Item(i).getAttribute("ItmDescription")
	AccItemDet.setAttribute "Amount",ItemNd.Item(i).getAttribute("ItemValue")
	AccItemDet.setAttribute "ItemCode",ItemNd.Item(i).getAttribute("ItemCode")
	AccItemDet.setAttribute "ClassCode",ItemNd.Item(i).getAttribute("ClassificationCode")
	AccDet.appendChild AccItemDet
	
next	

'alert(InvoiceDet.XML)
'exit function
set objhttp1 = CreateObject("Microsoft.XMLHTTP")
objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=AmdNewInvItemValue", false
objhttp1.send InvoiceDet.XMLDocument	


document.formname.action = "AmdInvPurInvoiceInsert_New.asp"
document.formname.submit()

end function
</script>

<%
'declaring variables
Dim sSqlTemp,sSuppInvNo,sItemDesc,sOrgID,sAccHeadExist,sItemAccHead,sTaxExist
Dim sPartyName,sPurType,sPartyType,nInvNo,sInvDt,sRndOffAccHead,sDesc
Dim sPurTypeName,sParSubTypeName,sAccHeads

Dim iAccountHead,dSumItemAccHead,dSumTaxAccHead,dAmt,iTaxCode,iCatCode,dCrTaxAmt
Dim dAccTotal,dToAdjust,dToPay,iItemCode,iClassCode,dAccValue,nTotal
Dim nItemValue,iItemCtr,iItemAccHead,nPrevTotalAmt,nTempCtr,iTempPurType,nArrCtr
Dim dRndOffAmt,dCrItemValue,dDrItemValue,dInvValue,dRoundOff,nCreatedTransNo
Dim iPartyCode,iPartySubtype,iCostCenterExist,iAnalHeadExist,iFormNo,iExempFormNo
Dim dDisTotal,dNettTotal,dBasicTotal,dItemBV,dDisAmt,dItemNett,dTotalitemValue

Dim dtSuppInvDate

Dim blnFirstTime,blnRoundOff,blnDisplayAcc

Dim Arr1,ArrTotalDisplay

Dim oDom,Node1,Node2,Node3,Root,ItemNd,AccHeadElem,ItemAccHead,AccDet,AccItemDet
Dim oNodTemp,oNodEntry,RtItem,RtAcc,RtTax,RtItemAcc,RtTaxAcc,oNodTaxRoot
Dim oNodVoucher,NewElem,NewElem1,NewElem2,elemItem,elemTax,NdItemCR,NdItemDR
Dim rsTemp

											
set rsTemp = Server.CreateObject("ADODB.RecordSet")
set oDom   = Server.CreateObject("Microsoft.XMLDOM")


oDOM.load server.MapPath("../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml")

set Root = oDom.DocumentElement

'creating node for account posting - start

For each oNodTemp in Root.childNodes
	if oNodTemp.nodeName="InvoiceHeader" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Header" then
								
				sPartyName	= oNodEntry.Attributes.getNamedItem("Party").value
				
				sOrgID		= oNodEntry.Attributes.Item(0).nodeValue
				sPurType	= oNodEntry.Attributes.Item(2).nodeValue
				iPartyCode	= oNodEntry.Attributes.Item(6).nodeValue
				sPartyType	= oNodEntry.Attributes.Item(7).nodeValue
				iPartySubtype  = oNodEntry.Attributes.Item(8).nodeValue
				
				nInvNo		= oNodEntry.Attributes.getNamedItem("InvoiceNumber").value

				sSuppInvNo = oNodEntry.Attributes.Item(20).nodeValue
				dtSuppInvDate = oNodEntry.Attributes.Item(21).nodeValue

				sInvDt = oNodEntry.Attributes.getNamedItem("SuppInvDt").value
				
				dInvValue = oNodEntry.Attributes.getNamedItem("InvValue").value
				dRoundOff = oNodEntry.Attributes.getNamedItem("RoundOff").value
				
			end if 'if oNodEntry.nodeName="Header" then
		next
	end if 'if oNodTemp.nodeName="InvoiceHeader" then
Next 'For each oNodTemp in Root.childNodes

if Root.hasChildNodes then
	'removing exist AccountHead Node
	for each node1 in Root.childNodes
		if Node1.nodename = "Voucher" then
			Root.RemoveChild Node1
		end if 
	next
end if 'if Root.hasChildNodes then
		
Set oNodVoucher = oDOM.createElement("Voucher")
Root.appendChild oNodVoucher

Set newElem = oDOM.createElement("Header")
oNodVoucher.appendchild newelem

Set newElem1 = oDom.createElement("Organization")
newelem1.setAttribute "OrgId",sOrgID
newelem1.text = getUnitName(sOrgID)
newElem.appendChild newElem1

' Book Node appended on 25-12-03
Set newElem1 = oDom.createElement("Book")
newelem1.setAttribute "BookId",0
newelem1.setAttribute "BKAccHead",0
newelem1.setAttribute "BKOtherUnits",0
newElem.appendChild newElem1

sSqlTemp = "Select PurTypeShortName from APP_M_PurchaseTypes where PurchaseType=" &sPurType& ""
'Response.Write sSqlTemp
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSqlTemp
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = Nothing
if not rsTemp.EOF then
	sPurTypeName = rsTemp(0)
end if
rsTemp.Close

Set newElem1 = oDom.createElement("PurchaseType")
newelem1.setAttribute "PurTypeId",sPurType
newelem1.text = sPurTypeName
newElem.appendChild newElem1

Set newElem1 = oDom.createElement("PurInvoice")
newelem1.setAttribute "PurInvNo",sSuppInvNo
newelem1.setAttribute "PurInvDate",dtSuppInvDate
newelem1.setAttribute "InvoiceNo",nInvNo
newElem.appendChild newElem1

sSqlTemp = "Select SubTypeName from APP_M_PartyTypes where PartyType='" &sPartyType& "' and PartySubType=" & trim(iPartySubtype) & ""
'Response.Write "<p> sSqlTemp= " & sSqlTemp
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSqlTemp
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = Nothing
if not rsTemp.EOF then
	sParSubTypeName = rsTemp(0)
end if
rsTemp.Close

Set newElem1 = oDom.createElement("Party")
newelem1.setAttribute "ParType",sPartyType
newelem1.setAttribute "ParSubType",iPartySubtype
newelem1.setAttribute "ParSubTypeName",sParSubTypeName
newelem1.setAttribute "ParCode",iPartyCode
newelem1.text = sPartyName
newElem.appendChild newElem1

dDisTotal = 0
dNettTotal = 0
dBasicTotal = 0


	for each Node1 in Root.ChildNodes
		if Node1.NodeName = "ItemDetails" then
		
			'storing  item details
			Set newElem  = oDOM.createElement("Details")
			oNodVoucher.appendchild newelem

			For Each Node2 in Node1.childNodes
				Set newElem1  = oDOM.createElement("Entry")
				
				sItemDesc = replace(Node2.getAttribute("ItmDescription"),"|","--")
				iItemCode = Node2.getAttribute("ItemCode")
				iClassCode  = Node2.getAttribute("ClassificationCode")
				
				dItemBV		= Node2.getAttribute("Qty")  * Node2.getAttribute("Rate")
				dDisAmt		= Node2.getAttribute("DisAmount")
				dItemNett	= Node2.getAttribute("NettBasic")
				
				newElem1.setAttribute "No",Node2.getAttribute("EntryNo")
				newElem1.setAttribute "PayTo",sItemDesc
				newElem1.setAttribute "Amount", Node2.getAttribute("Amount")
				newElem1.setAttribute "Qty",Node2.getAttribute("Qty")
				newElem1.setAttribute "UOM",Node2.getAttribute("UomDesc")
				newElem1.setAttribute "UOMValue",Node2.getAttribute("Uom")
				newElem1.setAttribute "Rate",Node2.getAttribute("Rate")
				newElem1.setAttribute "ActValue", dItemBV
				newElem1.setAttribute "DisPer", Node2.getAttribute("DisPer")
				newElem1.setAttribute "DisAmount",dDisAmt
				newElem1.setAttribute "ItemCode",iItemCode
				newElem1.setAttribute "ClassCode",iClassCode
				newElem1.setAttribute "ItemDesc",sItemDesc
				newElem1.setAttribute "ItemValue",Node2.getAttribute("ItemValue")
				newElem1.setAttribute "PurchaseType",Node2.getAttribute("PurchaseType")
				newElem1.setAttribute "AttributeList",Node2.getAttribute("AttributeList")

				newelem.appendChild newElem1
				
				iAccountHead = ""
				iCostCenterExist = ""
				iAnalHeadExist  = ""
				'********Modified Qry by Maheshwari on Aug 17,2007***************
				'sSqlTemp = "Select Distinct A.AccountHead,B.CostCenterExists,B.AnalyticalHeadExists from Inv_M_ItemOrgAccountHead A, VwOrgGLHeads B Where " &_
				'		" A.ItemCode=" & iItemCode & " and A.ClassificationCode=" & iClassCode & " and A.OrganisationCode ='" & trim(sOrgID) & "' " &_
				'		" and A.AccountHeadFor='P' And A.AccountHead = B.AccountHead And A.OrganisationCode=B.OUDefinitionID "
				sSqlTemp = " Select Distinct A.AccountHead,B.CostCenterExists,B.AnalyticalHeadExists from Acc_M_GLAccountHead A, VwOrgGLHeads B Where "&_
						   " B.OUDefinitionID ='" & trim(sOrgID) & "' And A.AccountHead = B.AccountHead "									
				'Response.Write sSqlTemp 
				With rsTemp
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSqlTemp
					.ActiveConnection = con
					.Open
				End With
				Set rsTemp.ActiveConnection = Nothing
				if not rsTemp.EOF then
					iAccountHead = rsTemp(0)
					iCostCenterExist = rsTemp(1)
					iAnalHeadExist  = rsTemp(2)
				End if
				rsTemp.Close

				Set newElem2  = oDOM.createElement("AccHead")
				newElem2.setAttribute "No",iAccountHead
				newElem2.setAttribute "CostCenter",iCostCenterExist
				newElem2.setAttribute "Analytical",iAnalHeadExist
				newElem2.setAttribute "Name",sItemDesc
				newElem2.setAttribute "Type","G"
				newElem2.setAttribute "Group",""
				newelem1.appendChild newElem2
														

				dBasicTotal  = cdbl(dBasicTotal) + cdbl(dItemBV)
				dDisTotal = cdbl(dDisTotal) + cdbl(dDisAmt)
				dNettTotal = cdbl(dNettTotal) + cdbl(dItemNett)
			Next 'For Each Node2 in Node1.childNodes
			
			newElem.setAttribute "BasicValue", dBasicTotal
			newElem.setAttribute "Discount", dDisTotal
			newElem.setAttribute "ActualValue", dNettTotal
			newElem.setAttribute "VouDate", FormatDate(date())
		end if 'if Node1.NodeName = "ItemDetails" then

		if Node1.NodeName = "TaxDetails" then
			Set oNodTaxRoot = oDOM.createElement("TaxDetails")
			oNodTaxRoot.setAttribute "InvoiceValue",dInvValue
			oNodTaxRoot.setAttribute "Basicvalue",dBasicTotal
			oNodTaxRoot.setAttribute "NettValue",dNettTotal
			oNodTaxRoot.setAttribute "RoundOff",dRoundOff
			oNodTaxRoot.setAttribute "PurchaseType",Node1.getAttribute("PurchaseType")
			
			oNodVoucher.AppendChild oNodTaxRoot
			
			
			For Each Node2 in Node1.ChildNodes
			
				''To fetch Tax form No. if its available in Exemption forms
				sSqlTemp = "Select isnull(FormNumber,'') from Pur_T_ExemptionForms Where InvoiceNumber=" & trim(nInvNo) & "" &_
						" and OUDefinitionID = '" & trim(sOrgID) & "' and TaxCode=" & trim(Node2.getAttribute("TaxCode")) & " and TaxCategoryCode = " & Node2.getAttribute("CatCode") & ""
				with rsTemp
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSqlTemp
					.ActiveConnection = con
					.Open
				end with
				'Response.write sSql
				set rsTemp.ActiveConnection = nothing

				If not rsTemp.EOF then
					iFormNo = rsTemp(0)
				Else
					iFormNo = ""
				End if
				rsTemp.Close

				sSqlTemp = "Select A.FormNumber from VwPurchaseTaxDetails A, RCV_T_InvoiceTaxDetails B  " & _
							" where A.OUDefinitionID='"& sOrgID &"' and A.TaxCategoryCode = " & Node2.getAttribute("CatCode") & " and " & _
							" A.TaxCategoryCode = B.TaxCategoryCode and A.TaxCode = B.TaxCode and " & _
							" A.TaxCode = " & Node2.getAttribute("TaxCode") & " and A.PurchaseType = "& sPurType  &"  and B.InvoiceNumber=" & trim(nInvNo) & " order by A.TaxHierarchy"
				with rsTemp
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSqlTemp
					.ActiveConnection = con
					.Open
				end with

				set rsTemp.ActiveConnection = nothing

				If not rsTemp.EOF then
					iExempFormNo = rsTemp(0)
				Else
					iExempFormNo = ""
				End if
				rsTemp.Close
				
				Set newElem = oDOM.createElement("Tax")
				newElem.setAttribute "CatCode",Node2.getAttribute("CatCode")
				newElem.setAttribute "TaxCode",Node2.getAttribute("TaxCode")
				newElem.setAttribute "TaxMode",Node2.getAttribute("TaxMode")
				newElem.setAttribute "TaxFormula",Node2.getAttribute("TaxFormula")
				newElem.setAttribute "TaxValue",Node2.getAttribute("TaxValue")
				newElem.setAttribute "TaxAmount",Node2.getAttribute("TaxAmount")
				newElem.setAttribute "AccHead",Node2.getAttribute("AccHead")
				If trim(Node2.getAttribute("CatCode")) = "0" and trim(Node2.getAttribute("TaxCode")) ="0" then
					newElem.setAttribute "Formnumber","0"
					newElem.setAttribute "TransAmt","0"
				else
					newElem.setAttribute "ToBeAccounted",Node2.getAttribute("ToBeAccounted")
					newElem.setAttribute "FormNo",iFormNo
					newElem.setAttribute "RoundOff",Node2.getAttribute("Rndoff")
					'newElem.setAttribute "CrEligible",Node2.getAttribute("CrEligible")
					'newElem.setAttribute "RegisterCode",Node2.getAttribute("RegisterCode")
					newElem.setAttribute "ExempFormNo",iExempFormNo
				end if
				

				newElem.Text= Node2.text
				'storing formnumber information in TaxDetails/tax Node
				Node2.setAttribute "FormNo",iFormNo
				oNodTaxRoot.appendChild newElem
			Next
		
		end if 'if Node1.NodeName = "TaxDetails" then
		
	Next 'for each Node1 in Root.ChildNodes
' account posting - end

blnFirstTime = true
sAccHeads = ""

nCreatedTransNo = 0
sSqlTemp = "Select isNull(CreatedVoucherNo,0) from RCV_T_InvoiceHeader where InvoiceNumber= " & nInvNo & ""
'Response.Write "<p> sSqlTemp= " & sSqlTemp
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSqlTemp
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = Nothing
if not rsTemp.EOF then
	nCreatedTransNo = rsTemp(0)
end if
rsTemp.Close

'Response.Write "<p> nCreatedTransNo = " & trim(nCreatedTransNo) 
if Root.hasChildNodes() then

	'removing exist AccountHead Node
	for each node1 in Root.childNodes
		if Node1.nodename = "AccountHead" then
			Root.RemoveChild Node1
		end if 
	next

	for each Node1 in Root.ChildNodes

		if Node1.NodeName = "InvoiceHeader" then
			for each Node2 in Node1.ChildNodes
				if Node2.NodeName = "Header" then
					sSuppInvNo		= Node2.getAttribute("SuppInvNo")
					dtSuppInvDate	= Node2.getAttribute("SuppInvDt")
					sOrgID			= Node2.getAttribute("OrgID")
					exit for
				end if 'if Node2.NodeName = "Header" then
			next	
		end if 'if Node1.NodeName = "InvoiceHeader" then
		
		if Node1.NodeName = "ItemDetails" then
			for each Node2 in Node1.ChildNodes
				if Node2.NodeName = "Item" then

										
									
					
					
			
					iItemAccHead = 0
					
					iItemCode = Node2.getAttribute("ItemCode")
					iClassCode = Node2.getAttribute("ClassificationCode")
					iTempPurType = Node2.getAttribute("PurchaseType")
					
					sSqlTemp = "Select isNull(AccUnitAccountHead,0) from Acc_T_CreatedVoucherDetails where CreatedTransNo = " &  nCreatedTransNo & " and AccountingUnit = '" & sOrgID & "' and ItemCode=" & iItemCode &" and ClassificationCode = " & iClassCode & " and InvoiceType = " & iTempPurType & " "
					'Response.Write "<p> sSqlTemp = "& sSqlTemp
					with rsTemp
						.ActiveConnection = con
						.CursorLocation = 3
						.CursorType = 3
						.Source = sSqlTemp
						.Open 
					end with
					
					set rsTemp.ActiveConnection = nothing
					
					if not rsTemp.EOF then
						iItemAccHead = rsTemp(0)
					end if 
					rsTemp.Close 
					
					'Response.Write "<p> iItemAccHead = "& trim(iItemAccHead)
					
					Node2.setAttribute "ItemAccHead" , iItemAccHead
					if instr(1,sAccHeads,iItemAccHead) <= 0 then
						sAccHeads = sAccHeads  &  trim(iItemAccHead) & ","
					end if 	
				end if 	'if Node2.NodeName = "Item" then
			next	
		end if 'if Node1.NodeName = "ItemDetails" then		
		
		if Node1.NodeName = "TaxDetails" then
		
			if blnFirstTime then
				blnFirstTime = false
				
				Set AccHeadElem = oDom.CreateElement("AccountHead")
				Root.appendChild AccHeadElem
				
				'adding Tax Details
				Set ItemAccHead = oDom.CreateElement("TaxAccountHead")
				AccHeadElem.appendChild ItemAccHead
			end if 	

			For Each Node2 in Node1.ChildNodes

				if trim(Node2.getAttribute("ToBeAccounted"))  = "1" then
					if CDbl(Node2.getAttribute("TaxAmount")) > 0 then
						sAccHeadExist = "N"
						nPrevTotalAmt = 0
				
						for each Node3 in ItemAccHead.ChildNodes
							if trim(Node3.getAttribute("No")) = trim(Node2.getAttribute("AccHead") ) then
								set AccDet = Node3
								nPrevTotalAmt = Node3.getAttribute("TotalAmt")
								sAccHeadExist = "Y"
							end if 
						next
				
						if sAccHeadExist = "N" then
							Set AccDet = oDom.CreateElement("Acc")
							AccDet.SetAttribute "No",Node2.getAttribute("AccHead")
							AccDet.SetAttribute "TotalAmt",Node2.getAttribute("TaxAmount")
							ItemAccHead.appendChild AccDet 
						else
							AccDet.SetAttribute "TotalAmt",cdbl(nPrevTotalAmt) +  cdbl(Node2.getAttribute("TaxAmount"))
						end if 'sAccHeadExist = "N"
					
						sTaxExist = "N"
						nPrevTotalAmt = 0
						for each Node3 in AccDet.ChildNodes
							if trim(Node3.getAttribute("TaxCode")) = trim(Node2.getAttribute("TaxCode") ) and trim(Node3.getAttribute("CatCode")) = trim(Node2.getAttribute("CatCode") )  then
								set AccItemDet = Node3
								nPrevTotalAmt = Node3.getAttribute("Amount")
								sTaxExist = "Y"
							end if 
						next
				
						if sTaxExist = "N" then
							Set AccItemDet = oDom.CreateElement("Tax")
							AccItemDet.setAttribute "Name",Node2.text
							AccItemDet.setAttribute "Amount",Node2.getAttribute("TaxAmount")
							AccItemDet.setAttribute "TaxCode",Node2.getAttribute("TaxCode")
							AccItemDet.setAttribute "CatCode",Node2.getAttribute("CatCode")
							AccDet.appendChild AccItemDet
						else
							AccItemDet.setAttribute "Amount",cdbl(nPrevTotalAmt) +  cdbl(Node2.getAttribute("TaxAmount") )
						end if 'if sTaxExist = "N" then
					end if 'if CDbl(Node2.getAttribute("TaxAmount")) > 0 then
				end if 'if trim(Node2.getAttribute("ToBeAccounted"))  = "1" then
			next
			
		end if 'if Node1.NodeName = "TaxDetails" then	

	next
end if 'if Root.hasChildNodes() then	

if right(sAccHeads,1) ="," then
	sAccHeads = mid(sAccHeads,1,len(sAccHeads) -1 )
end if 

'Response.Write "<p> sAccHeads = "& trim(sAccHeads)
if trim(sAccHeads) <> "" then


	Arr1 = Split(sAccHeads,",")
	ArrTotalDisplay = Split(sAccHeads,",")
	
	for nArrCtr = LBound(Arr1) to UBound(Arr1)
	
		nTotal = 0 
		
		for each Node1 in Root.ChildNodes
			if Node1.NodeName = "ItemDetails" then
				for each Node2 in Node1.ChildNodes
					if Node2.NodeName = "Item" then
						iItemAccHead = Node2.getAttribute("ItemAccHead")
						if trim(iItemAccHead) <> "0" then
							if trim(iItemAccHead) = trim(Arr1(nArrCtr)) then
								nTotal = cdbl(nTotal) + cdbl(Node2.getAttribute("ItemValue"))
							end if 
						end if 'if trim(sItemAccHead) <> "0" then	
					end if 'if Node2.NodeName = "Item" then	
				next 'for each Node2 in Node1.ChildNodes
			end if 	'if Node1.NodeName = "ItemDetails" then
		next 'for each Node1 in Root.ChildNodes	
		
		ArrTotalDisplay(nArrCtr ) = nTotal
	next
end if  'if trim(sAccHeads) <> "" then
oDOM.save server.MapPath("../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml")



For each oNodTemp in Root.childNodes
	
	If oNodTemp.nodeName="ItemDetails" then
		Set RtItem = oNodTemp 
	End if
	
	If oNodTemp.nodeName="TaxDetails" then
		Set RtTax = oNodTemp 
		blnRoundOff = false
		
		For each oNodEntry in oNodTemp.childnodes	' for Round Off
			If oNodEntry.getAttribute("CatCode") = "0" and cstr(oNodEntry.getAttribute("AccHead")) <> "0" then
				blnRoundOff = True
				sRndOffAccHead = oNodEntry.getAttribute("AccHead")
				dRndOffAmt = oNodEntry.getAttribute("TaxAmount")
				Exit For
			End if 		
		Next 
		
		if not blnRoundOff then
			dRndOffAmt = 0.00
		end if 
		
	End if
	
	If oNodTemp.nodeName="AccountHead" then
		Set RtAcc = oNodTemp 
	End if
Next

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
<form method="POST" name="formname" >
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Purchase Invoice Entry
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
	)								<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0">
											<tr>
												<td class="FieldCell">Supplier Inv. No. / Dt.
												</td>
												<td class="FieldCellSub" colspan="4">
													<table>
													<tr>
														<td>
															<span class="DataOnly"><%=sSuppInvNo%></span>
															&nbsp;
															<span class="DataOnly"><%=dtSuppInvDate%></span>
														</td>
													</tr>
													</table>
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Supplier Name
												</td>
												<td class="FieldCellSub" colspan="4">
													
													<span class="DataOnly"><%=sPartyName%></span>
													
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
										<div class="frmBody" id="frm2" style="width: 585; height:300;">
											<table border="0" cellspacing="1" class="ExcelTable" width="550" id="tblItemDet">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
													<td class="ExcelHeaderCell" align="center">Item Description</td>
													<td class="ExcelHeaderCell" align="center">Account Head</td>
													<td class="ExcelHeaderCell" align="center">Nett<br>
														Basic
													</td>
													<td class="ExcelHeaderCell" align="center" width="85">Total
													</td>													
												</tr>
												<%
												nArrCtr = 0
												iItemCtr = 0
												dAccTotal = 0
												dTotalitemValue = 0
												
												if Root.hasChildNodes() then
													for each Node1 in Root.ChildNodes

														if Node1.NodeName = "ItemDetails" then
															for each Node2 in Node1.ChildNodes
																if Node2.NodeName = "Item" then
																
																	sItemDesc		= Node2.getAttribute("ItmDescription")
																	iItemAccHead	= Node2.getAttribute("ItemAccHead")
																	nItemValue		= Node2.getAttribute("ItemValue")
																	iItemCtr = iItemCtr + 1
																	
																	'dAccTotal = cdbl(dAccTotal) + cdbl(nItemValue)
																	dTotalitemValue = cdbl(dTotalitemValue) + cdbl(nItemValue)
																	'Response.Write "nItemValue="&nItemValue
																	%>
																	<tr>
																		<td class="ExcelSerial" align="center"><%=iItemCtr%></td>
																		<td class="ExcelDisplayCell" align="left" width="100"><%=sItemDesc%></td>
																		<td class="ExcelDisplayCell" align="left">
																		<Select name="mCmbItemAccHeadZ<%=iItemCtr%>" size = 1 class="FormElem" onChange="DisplayTotal(<%=iItemCtr%>)">
																			<option value="0">Select Account Head</option>
																			<%
																				'' To pupulate Acc Heads
																				if instr(1,Node2.getAttribute("StockType"),"C") > 0 then
																					popAccountHead sOrgID,iItemAccHead, "C"
																				else
																					popAccountHead sOrgID,iItemAccHead, "S"
																				end if 	
																			%>
																		</Select>
																		</td>
																		<td class="ExcelDisplayCell" align="right"><%=formatNumber(nItemValue,2,,,0)%></td>
																		<td class="ExcelDisplayCell">
																			<%
																			if nArrCtr >= ubound(ArrTotalDisplay) then%>
																				<input type="text" name="ItemValueZ<%=iItemCtr%>" value="<%=nItemValue%>" class="FormElemRead" readonly >
																			<%else%>
																				<input type="text" name="ItemValueZ<%=iItemCtr%>" value="" class="FormElemRead" readonly >
																			<%end if %>
																			
																		</td>
																	</tr>
																	<%
																	nArrCtr = nArrCtr + 1
																end if 'if Node2.NodeName = "Item" then
															next	
															exit for
														end if 'if Node1.NodeName = "InvoiceHeader" then
													next
												end if 'if Root.hasChildNodes() then	
												%>

												<%
								
								
								For Each oNodEntry in RtAcc.childNodes
									If oNodEntry.nodeName="ItemAccountHead" then
										Set RtItemAcc = oNodEntry
									End if
									If oNodEntry.nodeName="TaxAccountHead" then
										Set RtTaxAcc  = oNodEntry
									End if
								Next
								
								

								Dim arrAccHead, sArrAccHead, iCtr, iCheckCtr
								For Each oNodEntry in RtAcc.childNodes
									If oNodEntry.nodeName = "ItemAccountHead" then
										For each oNodTemp in RtItemAcc.childnodes
											iAccountHead = oNodTemp.getAttribute("No")
											sArrAccHead = Split(arrAccHead,",")
											If UBound(sArrAccHead) >= 0 Then
												iCheckCtr = 0
												For iCtr = 0 to UBound(sArrAccHead)
													If sArrAccHead(iCtr) = iAccountHead Then
														iCheckCtr = iCheckCtr + 1
													End If
												Next
												If iCheckCtr = 0 Then arrAccHead = arrAccHead&","&iAccountHead
											Else
												arrAccHead = iAccountHead
											End If
										Next
									End If
								Next

								dim tempNode, sExp, iLen
								sArrAccHead = Split(arrAccHead,",")
								For iCtr = 0 to UBound(sArrAccHead)
									sExp = "//AccountHead/ItemAccountHead/Acc[@No = '"&sArrAccHead(iCtr)&"']"
									Set tempNode = RtAcc.selectNodes(sExp)

									If tempNode.length > 0 Then
										iAccountHead = TempNode.item(0).attributes.getnameditem("No").value
										dSumItemAccHead =  TempNode.item(0).attributes.getnameditem("TotalAmt").value
										dAccTotal = cdbl(dAccTotal) + cdbl(dSumItemAccHead)
										iItemCtr = iItemCtr + 1
								%>
										<tr>
											<td class="ExcelSerial" align="center"><%= iItemCtr %></td>
											<td class="ExcelDisplayCell" colSpan="2"><b><%= getAccountHeadName(iAccountHead) %></b></td>
											<td class="ExcelDisplayCell"></td>
											<td class="ExcelDisplayCell" align="right"><%= FormatNumber(dSumItemAccHead,2,,,0) %></td>
										</tr>
								<%
										sExp = "//AccountHead/ItemAccountHead/Acc[@No = '"&sArrAccHead(iCtr)&"']/Item"
										Set tempNode = RtAcc.selectNodes(sExp)
										
										For iLen = 0 to tempNode.length - 1
											sDesc = TempNode.item(iLen).attributes.getnameditem("Desc").value
											dAmt = TempNode.item(iLen).attributes.getnameditem("Amount").value
											
										%>
											<tr>
												<td class="ExcelSerial" align="center"></td>
												<td class="ExcelDisplayCell"  colSpan="2" ><%= sDesc%></td>
												<td class="ExcelDisplayCell" align="right"><%= FormatNumber(dAmt,2,,,0)%></td>
												<td class="ExcelDisplayCell" align="right"></td>
											</tr>
										<%
										Next
									End If
								Next

								Dim ndTemp
								For Each oNodEntry in RtAcc.childNodes
									If oNodEntry.nodeName="TaxAccountHead" then
										Set RtTaxAcc  = oNodEntry
											
											For each oNodTemp in RtTaxAcc.childnodes 
											
											if oNodTemp.hasChildNodes Then
											
											iAccountHead = oNodTemp.getAttribute("No")
											dSumTaxAccHead =  oNodTemp.getAttribute("TotalAmt")
											
											dAccTotal = cdbl(dAccTotal) + cdbl(dSumTaxAccHead)
											
											iItemCtr = iItemCtr + 1

								%>								
										<tr>
												<td class="ExcelSerial" align="center"><%= iItemCtr%>
												</td>
												<td class="ExcelDisplayCell"  colSpan="2"><b><%= getAccountHeadName(iAccountHead)%></b>
												</td>
												<td class="ExcelDisplayCell"></td>
												<td class="ExcelDisplayCell" align="right"><%= FormatNumber(dSumTaxAccHead,2,,,0) %>
												</td>
											</tr>
										<% For each ndTemp in oNodTemp.ChildNodes 
											sDesc = ndTemp.getAttribute("Name")
											dAmt = ndTemp.getAttribute("Amount")
											
											dCrItemValue = 0
											dCrItemValue= ndTemp.getAttribute("CRAmount")
											dDrItemValue = 0
											dDrItemValue= ndTemp.getAttribute("DRAmount")
										
										%>
											<tr>
												<td class="ExcelSerial" align="center">
												</td>
												<td class="ExcelDisplayCell" colSpan="2"><%= sDesc%>
												</td>
												<td class="ExcelDisplayCell" align="right"><%= FormatNumber(dAmt,2,,,0) %>
												</td>
												<td class="ExcelDisplayCell" align="right">
												</td>
											</tr>
										
											
								<% 
											Next
										End if
									Next
									End if

								Next
									
									%>
						
									<%
									'' round off exists
									If blnRoundOff Then
									%>
										<tr>
											<td class="ExcelSerial" align="center"><%= iItemCtr + 1%>
											</td>
											<td class="ExcelDisplayCell"  colSpan="2"><b><%= getAccountHeadName(sRndOffAccHead)%></b>
											</td>
											<td class="ExcelDisplayCell"></td>
											<td class="ExcelDisplayCell" align="right"><%= FormatNumber(dRndOffAmt,2,,,0) %>
											</td>
										</tr>
										<tr>
											<td class="ExcelSerial" align="center">
											</td>
											<td class="ExcelDisplayCell" align="Right"  colSpan="2">
											<% if cdbl(dRndOffAmt) >=0 then %>
											 Round Off Value (Cr)
											 <% else %>
											 Round Off Value (Dr)
											 <% end if%>
											</td>
											<td class="ExcelDisplayCell" align="right"><%= FormatNumber(dRndOffAmt,2,,,0) %>
											</td>
											<td class="ExcelDisplayCell" align="right">
											</td>
										</tr>
									<%
									dAccTotal = cdbl(dAccTotal) + cdbl(dRndOffAmt) 
									Else
										if cdbl(dAccTotal) > 0 then 
											dAccTotal = cdbl(dAccTotal) + cdbl(dRndOffAmt) 
										end if 	
									End if
									%>
									
									<%
									
									
									dToAdjust = 0
									dToPay = cdbl(dAccTotal) - cdbl(dToAdjust)
									
									%>		
											<tr>
												<td class="ExcelSerial" align="right" colspan="4"><b>Total&nbsp;</b>
												</td>
												<td class="ExcelDisplayCell" align="right">
													<input type="text" name="mTotalValue" value="<%= formatnumber(dAccTotal+dTotalitemValue,2,,,0) %>" class="FormElemRead" readonly >
													<input type="hidden" name="hTotalTaxValue" value="<%= formatnumber(dAccTotal,2,,,0) %>" >
												</td>
											</tr>
										<!--
											<tr>
												<td class="ExcelSerial" align="right" colspan="4"><b>Adjust&nbsp;</b>
												</td>
												<td class="ExcelDisplayCell" align="right">
													<input type="text" name="mAdjustValue" value="<%'= formatnumber(dToAdjust,2,,,0) %>" class="FormElemRead" readonly >
												</td>
											</tr>

											<tr>
												<td class="ExcelSerial" align="right" colspan="4"><b>To Pay&nbsp;</b>
												</td>
												<td class="ExcelDisplayCell" align="right">
													<input type="text" name="mToPayValue" value="<%'= formatnumber(dToPay,2,,,0) %>" class="FormElemRead" readonly >
												</td>
											</tr>
											-->
											


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
											        <input type="Button" value="Next"	name="BtnNext" onclick="javascript:Goto_AccItemValuePage()" class="ActionButton" tabindex="3">
 													<input type="reset"  value="Reset"	name="BtnReset" class="ActionButton" tabindex="4">
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
<%
Function popAccountHead(sOrgID,iSelAccHead,sStockType)

Dim dcrs,iAccHeadNo,sAccHeadName,iCCExist,iAHExist,sSqlTemp

Set dcrs = Server.CreateObject("ADODB.RecordSet")

if trim(sStockType) = ""  then sStockType = "S"

sStockType = ucase(trim(sStockType))

'blocked by kalaiselvi on March 30,2006
'to fetch item Acchead from Purchase / Fixed Asset Table
'With dcrs
'	.CursorLocation = 3
'	.CursorType = 3
'	.Source = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode from Inv_Control_OrgAccountHeads A,Acc_M_GLAccountHead B where A.AccountHeadFor='P' and A.OrganisationCode='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
'	.ActiveConnection = con
'	.Open
'End With
'Set dcrs.ActiveConnection = Nothing
'Set iAccHeadNo = dcrs(0)
'Set sAccHeadName = dcrs(1)
'Set iCCExist = dcrs(2)
'Set iAHExist = dcrs(3)

'If not dcrs.EOF then
'	Do while not dcrs.EOF
'		if trim(iSelAccHead) = trim(iAccHeadNo) then
'			Response.Write "<option value="""&trim(iAccHeadNo)&"|"&trim(iCCExist)& "|" &trim(iAHExist)& """ selected>" &trim(sAccHeadName) & "</option>" & vbCr
'		Else
'			Response.Write "<option value="""&trim(iAccHeadNo)&"|"&trim(iCCExist)& "|" &trim(iAHExist)& """>" & trim(sAccHeadName) & "</option>" & vbCr
'		End if
'
'		dcrs.MoveNext
'	Loop
'Else
'	dcrs.Close

	'sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccHeadForInventApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	if trim(sStockType) = "C" then
		sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccHeadForFAApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	else
		sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccheadforPurchaseApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	end if 	
	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSqlTemp
		.ActiveConnection = con
		.Open
	End With
	
	Set dcrs.ActiveConnection = Nothing
	
	Set iAccHeadNo = dcrs(0)
	Set sAccHeadName = dcrs(1)
	Set iCCExist = dcrs(2)
	Set iAHExist = dcrs(3)

	Do while not dcrs.EOF
		if trim(iSelAccHead) = trim(iAccHeadNo) then
			Response.Write "<option value="""&trim(iAccHeadNo)&"|"&trim(iCCExist)& "|" &trim(iAHExist)& """ selected>" & trim(sAccHeadName) & "</option>" & vbCr
		Else
			Response.Write "<option value="""&trim(iAccHeadNo)&"|"&trim(iCCExist)& "|" &trim(iAHExist)& """>" & trim(sAccHeadName) & "</option>" & vbCr
		End if
		dcrs.MoveNext
	Loop
'End if
dcrs.Close

End Function
%>

<%
Function getAccountHeadName(iAccHead)

Dim objAcc,sSql
Set objAcc = Server.CreateObject("Adodb.Recordset")

sSql = "Select AccountDescription,AccountHeadCode from Acc_M_GLAccountHead where AccountHead=" & iAccHead & ""
with objAcc
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
end with
set objAcc.ActiveConnection = nothing

If not objAcc.EOF then
	getAccountHeadName = objAcc(0)
Else
	getAccountHeadName = iAccHead  
End if 
objAcc.Close  

End function
%>
