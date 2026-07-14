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

<SCRIPT type="text/plain" data-itms-legacy-client-script="1">

Dim sInvAgainst,sInvRefNum,bFlag
Dim TaxRoot,RootNode,oNodTemp,ItemRoot, dBasicVal,dTBasicVal
'--------------------------------------------------------------------------
function SetDate()
	sDate    = document.formname.hInvDate.value
	sMaxDate = document.formname.hMaxDate.value
	sMinDate = document.formname.hMinDate.value
	document.formname.ctldate.SetMinDate = sMinDate
	document.formname.ctldate.SetMaxDate = sMaxDate
end function
'--------------------------------------------------------------------------

function MinDate()
	sDate = document.formname.ctldate.GetDate
	sMinDate = document.formname.hMinDate.value
	sMaxDate = document.formname.hMaxDate.value

	If dateDiff("d",sDate,document.formname.hMinDate.value) > 0 or  dateDiff("d",sDate,document.formname.hMaxDate.value) < 0 then
		alert("Date Should be within the Financial Year  "& sMinDate&" to " & sMaxDate )
		document.formname.ctldate.SetDate = document.formname.hMaxDate.value
		exit Function
	end if
end function
'--------------------------------------------------------------------------
'*************************
Function PopulatePartyType()
	Dim ndRoot
	if document.formname.hPartyCode.value<>"" then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		    'alert(document.formname.hPartyCode.value)
			objhttp.open "GET","../../Common/PartySubType.asp?ParCode="&document.formname.hPartyCode.value&"&OrgCode="&document.formname.hOrgID.value,false
			objhttp.send
			'alert(objhttp.responseText)
			if trim(objhttp.responseXML.xml)<>"" then
				PartySubTypeData.loadXML(objhttp.responseXML.xml)
				set ndRoot = PartySubTypeData.documentElement

				if ndRoot.hasChildNodes() then
					if ndRoot.childNodes.length=1 then
						document.formname.cmbPartyType.length = 0
					else
						document.formname.cmbPartyType.length = 1
					end if
					for each ndChild in ndRoot.childNodes
						document.formname.cmbPartyType.length =document.formname.cmbPartyType.length +1
						sTemp = split(ndChild.getAttribute("SubType"),"|")
						document.formname.cmbPartyType(document.formname.cmbPartyType.length-1).value = sTemp(1)&"|"&sTemp(0)
						document.formname.cmbPartyType(document.formname.cmbPartyType.length-1).text = ndChild.text
					next
				end if 'if ndRoot.hasChildNodes() then
			else
				alert(objhttp.responseText)
			end if
	end if 'if document.formname.hPartyCode.value<>"" then
End Function
'*****************************************
function init(sSelRcptNo,sOrgID)
	if trim(document.formname.hSuppInvDate.value) <> "" then
		document.formname.ctlDate.setdate = document.formname.hSuppInvDate.value
	else
		document.formname.ctlDate.setdate = date()
	end if
	'alert("start init")
	if sSelRcptNo <> "0" then
		'SpnRcptCode.innerText = left(document.formname.hRcptCode.value,len(document.formname.hRcptCode.value)-1)
    End if
		sFlag = document.formname.hFlag.value

'		displayItemDetail sOrgID,sFlag
'		getTaxDet()

		Set ObjRoot = InvoiceDet.documentElement
		'alert(ObjRoot.xml)

		for each Node1 in ObjRoot.ChildNodes
			if Node1.nodeName = "InvoiceHeader" then
				for each Node2 in Node1.ChildNodes
					if Node2.NodeName = "Header" then
						sPurType		= Node2.GetAttribute("PurchaseType")
						'document.formname.ctlDate.setdate = Node2.GetAttribute("SuppInvDt")
						document.formname.txtSuppInvNo.value = Node2.GetAttribute("SuppInvNo")

						'alert(document.formname.hPSubType.value)
						'alert(document.formname.hPartyType.value)
					    For icnt = 0 to document.formname.cmbPartyType.length - 1
					        if document.formname.cmbPartyType(iCnt).value = document.formname.hPSubType.value & "|" & document.formname.hPartyType.value then
					            document.formname.cmbPartyType.selectedIndex = iCnt
					        end if
					    Next
						'document.formname.cmbPartyType.value = document.formname.hPSubType.value & "|" & document.formname.hPartyType.value
						document.formname.txtPartyName.value = sPartyName

					end if 'if Node2.NodeName = "Header" then
				next 'for each Node2 in Node1.ChildNodes
			end if 'if Node1.nodeName = "InvoiceHeader" then
		next ' for each Node1 in ObjRoot.ChildNodes
    
		document.all.ImgDeleteIcon.disabled=true
	
'	alert(sPurType)
		if trim(sPurType) <> "0" then

			additem
			'popTax
		else
			PurchaseTypeWise
			'PopTax1
		end if
	

end function
'***********************************************************
Function RefType_Click()
Dim ndInvRoot,ndType
Dim ndInvHeader,ndHead,ndInvItemDet,ndItem
Dim ItemNode,MaterialNode,EntryNode
Dim sRefType,sPartyCode,sOrgID,sarrValue
Dim RcptNo,ForUnit,Flag,iItemEntryNo
set ndInvRoot = InvoiceDet.documentElement

sOrgID = document.formname.hOrgID.value
if trim(sOrgID)="" or IsNull(sOrgID) then
		sOrgID = document.formname.hOrgID.value
end if
nFlag=1
iStock="N"
bAddButton= "N"
sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
sPartyCode = document.formname.hPartyCode.value
sDispItem = 0
    'RefTypeSelection sRefType,sOrgID,sItemType,sPartyCode,iStock,nFlag,bAddButton,sDispItem
    RefTypeSelection sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem
    'alert(OutData.xml)
	if trim(sRefType)<>"N" then
		set Root = OutData.documentElement
		if Root.hasChildNodes() then
			For each RefNode in Root.childNodes
				RcptNo	= RefNode.getAttribute("ReferenceNo")
				ForUnit = sOrgID
				Flag	= ""
				
				if trim(RefNode.getAttribute("Remarks"))<>"" then
					PartyCode = split(RefNode.getAttribute("Remarks"),"-")(0)
					spnSuppName.innerText = split(RefNode.getAttribute("Remarks"),"-")(1)
				end if  'if trim(RefNode.getAttribute("Remarks"))<>"" then
				SpnRcptCode.innerText = trim(RefNode.getAttribute("ReferenceCode"))&" - "& trim(RefNode.getAttribute("ReferenceDate"))

				document.formname.hOrgID.value = ForUnit
				document.formname.hPartyCode.value = PartyCode
				document.formname.hRecNo.value = RcptNo
				document.formname.hRefType.value=sRefType
				document.formname.hRefDate.value = RefNode.getAttribute("ReferenceDate")

			next

				PopulatePartyType()

			set objhttp = CreateObject("Microsoft.XMLHTTP")
			objhttp.open "GET","InvItemPopulate.asp?RefType="& sRefType &"&hRecNo="& document.formname.hRecNo.value&"&OrgID="&sOrgID&"&InvNo="&document.formname.hInvNo.value,false
			objhttp.send
'			alert(objhttp.responseText)

			if trim(objhttp.responseXML.xml)<>"" then
				InvoiceDet.loadXML(objhttp.responseXML.xml)
			else
				alert(objhttp.responseText)
			end if
		end if
	else

		if trim(spnSuppName.innerText)="" then
			alert("Select the Party")
			exit function
		end if
		set Root = OutData.documentElement

		if ndInvRoot.hasChildNodes() then
			For each ndInvHead in ndInvRoot.childNodes
				ndInvRoot.removeChild(ndInvHead)
			Next
		end if 'if ndInvRoot.hasChildNodes() then

		Set ndInvHeader = InvoiceDet.createElement("InvoiceHeader")
		ndInvRoot.appendChild ndInvHeader

		Set ndHead = InvoiceDet.createElement("Header")
			ndHead.setAttribute "OrgID",document.formname.hOrgID.value
			ndHead.setAttribute "Party",spnSuppName.innerText
			ndHead.setAttribute "PurchaseType",""
			ndHead.setAttribute "Currency",""
			ndHead.setAttribute "InvAgainst",""
			ndHead.setAttribute "RefNum",""
			ndHead.setAttribute "PartyCode",""
			ndHead.setAttribute "PartyType",""
			ndHead.setAttribute "PartySubType",""
			ndHead.setAttribute "CurrencyNo",""
			ndHead.setAttribute "DespatchMode",""
			ndHead.setAttribute "PaymentMode",""
			ndHead.setAttribute "PayTerms",""
			ndHead.setAttribute "IssueBank",""
			ndHead.setAttribute "BenificiaryBank",""
			ndHead.setAttribute "PricingBasis",""
			ndHead.setAttribute "Transporter",""
			ndHead.setAttribute "LoadingPort",""
			ndHead.setAttribute "DestPort",""
			ndHead.setAttribute "Remarks",""
			ndHead.setAttribute "SuppInvNo",""
			ndHead.setAttribute "SuppInvDt",""
			ndHead.setAttribute "TransporterFlag",""
			ndHead.setAttribute "PoNo",""
			ndHead.setAttribute "ConfNum",""
			ndHead.setAttribute "InvoiceFlag",""
			ndHead.setAttribute "InvValue","0"
			ndHead.setAttribute "RoundOff","0"
			ndHead.setAttribute "SuppCode",""
			ndHead.setAttribute "ItemType",""
			ndHead.setAttribute "InvoiceNumber",document.formname.hInvNo.value
		ndInvHeader.appendChild ndHead

		set ndInvItemDet = InvoiceDet.createElement("ItemDetails")
		ndInvRoot.appendChild ndInvItemDet

		if Root.hasChildNodes() then
		iItemEntryNo = 0
			for each ItemNode in Root.ChildNodes
				if ItemNode.nodeName="Item" then
					iItemEntryNo = iItemEntryNo + 1
					set ndItem = InvoiceDet.createElement("Item")
						ndItem.setAttribute "ItemCode",ItemNode.getAttribute("ItemCode")
						ndItem.setAttribute "ClassificationCode",ItemNode.getAttribute("ClassCode")
						ndItem.setAttribute "ItmDescription",ItemNode.getAttribute("ItemName")
						ndItem.setAttribute "Uom",ItemNode.getAttribute("StoresUoM")
						ndItem.setAttribute "Qty","0"
						ndItem.setAttribute "Rate","0"
						ndItem.setAttribute "DisPer","0"
						ndItem.setAttribute "DisAmount","0"
						ndItem.setAttribute "NettBasic","0"
						ndItem.setAttribute "UomDesc",ItemNode.getAttribute("StoresUoM")
						ndItem.setAttribute "EntryNo",iItemEntryNo
						ndItem.setAttribute "RatePerQtyUoM","0"
						ndItem.setAttribute "SourceEntryNo",""
						ndItem.setAttribute "PurchaseType",""
						ndItem.setAttribute "Amount",""
						ndItem.setAttribute "ItemValue","0"
						ndItem.setAttribute "ItemRate","0"
						ndItem.setAttribute "RateUOM",""
						ndItem.setAttribute "StockType","S"
						ndItem.setAttribute "VAT",""
						ndItem.setAttribute "AttributeList",ItemNode.getAttribute("AttributeList")
					ndInvItemDet.appendChild ndItem
				elseif ItemNode.nodeName="Materials" then
					for each EntryNode in ItemNode.childNodes
						if EntryNode.nodeName="Entry" then
							iItemEntryNo = iItemEntryNo + 1
								set ndItem = InvoiceDet.createElement("Item")
									ndItem.setAttribute "ItemCode","0"
									ndItem.setAttribute "ClassificationCode","0"
									ndItem.setAttribute "ItmDescription",EntryNode.getAttribute("ItemName")
									ndItem.setAttribute "Uom","NOS"
									ndItem.setAttribute "Qty","0"
									ndItem.setAttribute "Rate","0"
									ndItem.setAttribute "DisPer","0"
									ndItem.setAttribute "DisAmount","0"
									ndItem.setAttribute "NettBasic","0"
									ndItem.setAttribute "UomDesc","NOS"
									ndItem.setAttribute "EntryNo",iItemEntryNo
									ndItem.setAttribute "RatePerQtyUoM","0"
									ndItem.setAttribute "SourceEntryNo",""
									ndItem.setAttribute "PurchaseType",""
									ndItem.setAttribute "Amount",""
									ndItem.setAttribute "ItemValue","0"
									ndItem.setAttribute "ItemRate","0"
									ndItem.setAttribute "RateUOM",""
									ndItem.setAttribute "StockType","S"
									ndItem.setAttribute "VAT",""
									ndItem.setAttribute "AttributeList",""
								ndInvItemDet.appendChild ndItem
						end if ' if EntryNode.nodeName="Entry" then
					next
				end if ' if ItemNode.nodeName="Item" then
			next
		end if 'if Root.hasChildNodes() then
	end if 'if trim(sRefType)<>"N" then
getTaxDet()
End Function
'************************************************************************
'----------------------------------------------------------------------------
function AddItem()
Dim oRow,headerCell,Root,iCtr,ItemNode,sExp,i,Node,InvQty,Rate,DisPer,DisAmount
Dim NettBasic,ItemCode,ClassCode,TaxMode,TaxVal,TaxPer,TaxName
Dim sCatCode,sTaxCode,UomCode,dTotItemValue,sPurchaseType

ClearTable
Set Root=InvoiceDet.documentElement
iCtr=1

sPurchaseType = document.Formname.cmbPurType.value

'alert(Root.xml)
if Root.hasChildNodes then
	dTotItemValue = 0

	Set ItemNode=Root.Selectnodes("//ItemDetails/Item[@PurchaseType="&sPurchaseType&"] ")
	for i = 0 to ItemNode.Length - 1
			ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
			ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
			EntNo = ItemNode.Item(i).Attributes.getNamedItem("EntryNo").value
			InvQty = ItemNode.Item(i).Attributes.getNamedItem("Qty").value
			Rate=FormatNumber(ItemNode.Item(i).Attributes.getNamedItem("Rate").value,5,,,0)
			DisPer=ItemNode.Item(i).Attributes.getNamedItem("DisPer").value
			DisAmount=ItemNode.Item(i).Attributes.getNamedItem("DisAmount").value
			NettBasic=ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value
			UomCode=ItemNode.Item(i).Attributes.getNamedItem("Uom").value
			nRatePerQtyUoM =ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value
			nItemValue =ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value
			sRateUOM =ItemNode.Item(i).Attributes.getNamedItem("RateUOM").value
			sVAT =ItemNode.Item(i).Attributes.getNamedItem("VAT").value
			dTotItemValue = cdbl(dTotItemValue)+ cdbl(nItemValue)

			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)
			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtSerA"& ClassCode &"A"& ItemCode &"A"& EntNo&""" value="""&iCtr&""" size=""1"" class=""FormelemRead"" style=""text-align=center"" READONLY>" )
			headerCell.appendChild(oText)
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""checkbox""  name=""chkDeleteA"& ClassCode &"A"& ItemCode &"A"& EntNo&""" value=""Y""  class=""FormElem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=ItemNode.Item(i).Attributes.getNamedItem("ItmDescription").value
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Left"

			set headerCell=oRow.insertCell()
			headercell.innerHtml = UomCode + "<input type=""hidden"" name=""hRateUOM"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value=" & sRateUOM & " >"
			headercell.innerHtml = headercell.innerHtml + "<input type=""hidden"" name=""hRatePerQtyUoM"& ClassCode &"Z"& ItemCode &"Z"& EntNo&""" value=" & nRatePerQtyUoM  & " >"
			headerCell.width = "10"
			headerCell.className="ExcelFieldCell"
			headerCell.align="center"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtInvQty"& ClassCode &"Z"& ItemCode &"Z"& EntNo&""" value="& InvQty &" onBlur=""DisplayAmount('Q','"& ItemCode &"','"& ClassCode &"','"& EntNo &"',this)"" style=""text-align: Right"" size=""10"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelInputCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtInvRate"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& FormatNumber(Rate,5,,,0) &" onBlur=""DisplayAmount('R','"& ItemCode &"','"& ClassCode &"','"& EntNo &"',this)""  style=""text-align: Right"" size=""11"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headercell.innerHtml = headercell.innerHtml + "&nbsp;&nbsp;<a><Img border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" onclick=""SetRateUOM('"& trim(ClassCode) & "','" & trim(ItemCode) &"','" & trim(EntNo) &"','" & UomCode & "','SP','Y')"" alt=""Select Rate UOM"" style=""cursor:hand""></a>"
			headerCell.width = "100"
			headerCell.className="ExcelInputCell"
			headerCell.align="center"

			nAmt = cdbl(InvQty) * cdbl(Rate)

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtValue"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value=" & FormatNumber(nAmt,2,,,0)  & " size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="FormelemRead"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtDisPer"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& formatnumber(DisPer,2,,,0) &" onBlur=""DisplayAmount('D','"& ItemCode &"','"& ClassCode &"','"& EntNo &"',this)"" style=""text-align: Right"" size=""6"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelInputCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtDisAmount"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& FormatNumber(DisAmount,2,,,0)  &" Readonly maxlength=13  style=""text-align: Right"" class=""FormelemRead"" size=""11"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtNetValue"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& FormatNumber(NettBasic,2,,,0)  &" size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" readonly name=""txtItemRateDisplay" & ClassCode & "Z" & ItemCode &"Z"& EntNo & """ value=" & FormatNumber(nItemValue,2,,,0) & " size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			if sVAT = "Y" then
				set oText = document.createElement("<input type=""checkbox"" name=""chkVat" & ClassCode & "Z" & ItemCode &"Z"& EntNo & """ CHECKED class=""Formelem"">")
			else
				set oText = document.createElement("<input type=""checkbox"" name=""chkVat" & ClassCode & "Z" & ItemCode &"Z"& EntNo & """ class=""Formelem"">")
			end if
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			'set headerCell=oRow.insertCell()
			'set oText = document.createElement("<select size=""1"" name=""D16"" class=""Formelem""><option>Purchase Type</option></select>")
			'headerCell.appendChild(oText)
			'headerCell.width = "10"
			'headerCell.className="ExcelDisplayCell"
			'headerCell.align="center"

			iCtr=iCtr+1

			nTotalBasicValue =  CDbl(nTotalBasicValue) + CDbl(nAmt)
			nTotalDiscValue  =  CDbl(nTotalDiscValue) + CDbl(DisAmount)
			nTotalNettBasic  =  CDbl(nTotalNettBasic) + CDbl(NettBasic)
	next
	'To display Total
	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "6"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Total</B>"
	headerCell.align="Right"


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  readonly  name=""txtBasicValue"" value=" & FormatNumber(nTotalBasicValue,2,,,0) & " size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Center"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" readonly  name=""txtDisValue"" value=" & FormatNumber(nTotalDiscValue,2,,,0) & " size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Center"


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  readonly name=""txtAmount"" value =" & FormatNumber(nTotalNettBasic,2,,,0) & " size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Center"
	for i = 0 to ItemNode.Length - 1
		EntNo = ItemNode.Item(i).Attributes.getNamedItem("EntryNo").value
		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text"" readonly name=""txtTotalItemRate" & sPurchaseType &"Z"& EntNo & """ value =" & FormatNumber(dTotItemValue,2,,,0) & " size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelDisplayCell"
		headerCell.align="Center"
	Next
	
	set headerCell=oRow.insertCell()
		headerCell.width = "10"
		headerCell.align="center"
		headerCell.className="ExcelDisplayCell"

	Set ItemNode=Root.Selectnodes("//TaxDetails")
	ItemNode.item(0).Attributes(0).value = nTotalBasicValue
	ItemNode.item(0).Attributes(1).value = nTotalNettBasic

	nTotalTaxValue = 0
	'To display tax details
	Set ItemNode=Root.Selectnodes("//TaxDetails/Tax")

	for i = 0 to ItemNode.Length - 1

		TaxName=ItemNode.Item(i).text
		sCatCode=trim(ItemNode.Item(i).Attributes.getNamedItem("CatCode").value )
		sTaxCode=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxCode").value )
		TaxMode=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxMode").value )

		sTaxDisp = ""
		if trim(TaxMode) = "K" then
			sTaxDisp = "/Pack"
		elseif trim(TaxMode) = "Q" then
			sTaxDisp = "/Qty"
		elseif trim(TaxMode) = "P" then
			sTaxDisp = "%"
		end if

		TaxPer=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxValue").value )
		TaxVal=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxAmount").value )

		if trim(ItemNode.Item(i).Attributes.getNamedItem("CatCode").value) <> "0" and trim(ItemNode.Item(i).Attributes.getNamedItem("TaxCode").value) <> "0" then
			if trim(ItemNode.Item(i).Attributes.getNamedItem("AccHead").value ) = "0" then
				nTotalTaxValue = cdbl(nTotalTaxValue) + cdbl(TaxVal)
			end if
		end if

		sOrgID = document.formname.hOrgID.value

		set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

		if trim(TaxMode)<>"K" then
			set headerCell=oRow.insertCell()
			headerCell.colspan = "8"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml=TaxName
			headerCell.align="Right"
		else
			set headerCell=oRow.insertCell()
			headerCell.colspan = "8"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml=TaxName
			headerCell.align="Right"
			headerCell.innerHtml =headerCell.innerHtml+"<a href='#' onClick=Packvaluechange('"+ CStr(sOrgId) + "','" + cstr(sPurchaseType) + "','" + cstr(sCatCode)+"','" + cstr(sTaxCode)+ "') 'class='ExcelDisplayLink' ><img border='0' src='../../assets/images/iTms icons/DetailsIcon.gif' align='center' width='11' height='11' alt='View Computation Mode Details'> </a>"
		end if ' if trim(TaxMode)<>"K" then

		set headerCell=oRow.insertCell()
		if trim(TaxMode) = "F" or trim(TaxMode) = "K" then
			headerCell.ClassName = "ExcelDisplayCell"
			headerCell.Align		= "left"
		else
			set oText = document.createElement("<input type=""text"" name=""txtTaxPer" & sCatCode & sTaxCode & """ value="""& TaxPer &""" size=""6"" onChange=CalcTax() style=""text-align: Right"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.innerhtml=headerCell.innerhtml + sTaxDisp
			headerCell.ClassName = "ExcelInputCell"
			headerCell.Align		= "left"
		end if


		set headerCell=oRow.insertCell()

		if trim(TaxMode)="F" then
			set oText = document.createElement("<input type=""text"" name=""txtTaxPer" & sCatCode & sTaxCode & """ value="""& TaxVal &""" size=""6"" onChange=CalcTax() style=""text-align: Right"" class=""Formelem"">")
			headerCell.appendChild(oText)
			set oText = document.createElement("<input type=""hidden"" name=""txtTaxValue" & sCatCode &"Z"& sTaxCode & """  size=""11"" value=0 onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" class=""Formelem"">")
			headerCell.appendChild(oText)
		elseif trim(TaxMode)="K" then
			set oText = document.createElement("<input type=""hidden"" name=""txtTaxPer" & sCatCode & sTaxCode & """  size=""6"" style=""text-align: Right"" class=""Formelem"" value=0>")
			headerCell.appendChild(oText)
			set oText = document.createElement("<input type=""text"" name=""txtTaxValue" & sCatCode &"Z"& sTaxCode & """ value="""& TaxVal &""" size=""11"" onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" ReadOnly class=""FormelemRead"">")
			headerCell.appendChild(oText)
		else
			set oText = document.createElement("<input type=""text"" name=""txtTaxValue" & sCatCode &"Z"& sTaxCode & """ value="""& TaxVal &""" size=""11"" onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" ReadOnly class=""FormelemRead"">")
			headerCell.appendChild(oText)
		end if

		headerCell.width = "10"
		headerCell.align="center"
		if trim(TaxMode)="F" then
			headerCell.className="ExcelInputCell"
		else
			headerCell.className="ExcelDisplayCell"
		end if

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text""  readonly  name=""txtSubTaxValue" & sCatCode & sTaxCode & """ value=" & formatnumber(nTotalNettBasic,2,,,0) & " size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelAverageCell"
		headerCell.align="center"
		
		set headerCell=oRow.insertCell()
		headerCell.width = "10"
		headerCell.align="center"
		headerCell.className="ExcelDisplayCell"
		
	next

	Set ItemNode=Root.Selectnodes("//TaxDetails")
	ItemNode.item(0).Attributes(2).value = nTotalTaxValue
	ItemNode.item(0).Attributes(3).value = nTotalNettBasic

end if
	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Rounded Off</B>("
	set oText = document.createElement("<input type=""radio"" value=""Y"" checked=true name=""rdRndOff"" onclick=""RoundOffInv()"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml=headerCell.innerhtml + "Yes"
	set oText = document.createElement("<input type=""radio"" value=""N"" name=""rdRndOff"" onclick=""RoundOffInv()"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml = headerCell.innerhtml + "No)"
	headerCell.align="Right"


	dRoundOffValue = round(nTotalNettBasic,0) - nTotalNettBasic
	nTotalNettBasic = round(nTotalNettBasic,0)


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" NAME=""txtRoundOff"" ReadOnly onBlur=""AssignRoundOffValue()"" style=""text-align: Right"" size=""11""  class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"
	
	set headerCell=oRow.insertCell()
		headerCell.width = "10"
		headerCell.align="center"
		headerCell.className="ExcelDisplayCell"

	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Invoice Value</B>"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  readonly  NAME=""txtInvValue"" value=" & FormatNumber(nTotalNettBasic,2,,,0) & " style=""text-align: Right"" size=""11"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"
	
	set headerCell=oRow.insertCell()
		headerCell.width = "10"
		headerCell.align="center"
		headerCell.className="ExcelDisplayCell"

	document.all.ImgDeleteIcon.disabled= false


	'note : the following block is used to calculate ItemRate based on InvoiceRate Column
	'******* block ***
	if Root.hasChildNodes then
		Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
		for i = 0 to ItemNode.Length - 1
			ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
			ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
			EntNo=ItemNode.Item(i).Attributes.getNamedItem("EntryNo").value
			UomCode=ItemNode.Item(i).Attributes.getNamedItem("Uom").value

			SetRateUOM trim(ClassCode),trim(ItemCode),trim(EntNo),UomCode,"SP","N"
		next
	end if 'if Root.hasChildNodes then
	'******* block ***

end function

'-------------------------------------------------------------------------------------------------

function PurchaseTypeWise()
Dim oRow,headerCell,Root,iCtr,ItemNode,sExp,i,Node,InvQty,Rate,DisPer
Dim DisAmount,NettBasic,ItemCode,ClassCode,TaxMode,TaxVal,TaxPer,TaxName
Dim TaxNode,PurType,j,nTempLastItemCtr,dTotItemValue,nTotal,nInvValue,nSubTotal

ClearTable
Set Root=InvoiceDet.documentElement
iCtr=1
nTempLastItemCtr = 0
'alert(Root.xml)
if Root.hasChildNodes then
	dTotItemValue = 0
	Set TaxNode=Root.Selectnodes("//TaxDetails")
	for j = 0 to TaxNode.Length - 1
		PurType=TaxNode.Item(j).Attributes.getNamedItem("PurchaseType").value
		Set ItemNode=Root.Selectnodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")
		for i = 0 to ItemNode.Length - 1
			nTempLastItemCtr = nTempLastItemCtr + 1
		Next
	Next
end if

if Root.hasChildNodes then
	nTotal = 0
	nInvValue = 0

	Set TaxNode=Root.Selectnodes("//TaxDetails")
	for j = 0 to TaxNode.Length - 1

		PurType=TaxNode.Item(j).Attributes.getNamedItem("PurchaseType").value
		nTotalTaxValue= TaxNode.Item(j).Attributes.getNamedItem("TotalTax").value
		nSubTotal= TaxNode.Item(j).Attributes.getNamedItem("SubTotal").value

		nInvValue = cdbl(nInvValue) + cdbl(nSubTotal)

		nTotalBasicValue = 0
		nTotalDiscValue  = 0
		nTotalNettBasic  = 0


		dTotItemValue = 0

		Set ItemNode=Root.Selectnodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")
		for i = 0 to ItemNode.Length - 1

				ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
				ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
				EntNo = ItemNode.Item(i).Attributes.getNamedItem("EntryNo").value
				InvQty = ItemNode.Item(i).Attributes.getNamedItem("Qty").value
				Rate=FormatNumber(ItemNode.Item(i).Attributes.getNamedItem("Rate").value,5,,,0)
				DisPer=ItemNode.Item(i).Attributes.getNamedItem("DisPer").value
				DisAmount=ItemNode.Item(i).Attributes.getNamedItem("DisAmount").value
				NettBasic=ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value
				UomCode=ItemNode.Item(i).Attributes.getNamedItem("Uom").value
				nRatePerQtyUoM =ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value
				nItemValue =ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value
				sRateUOM =ItemNode.Item(i).Attributes.getNamedItem("RateUOM").value

				dTotItemValue	= cdbl(dTotItemValue)+ cdbl(nItemValue)
				nTotal			= cdbl(nTotal)+ cdbl(nItemValue)

				set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtSerA"& ClassCode &"A"& ItemCode  &"A"& EntNo &""" value="""&iCtr&""" size=""1"" class=""FormelemRead"" style=""text-align=center"" READONLY>" )
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox""  name=""chkDeleteA"& ClassCode &"A"& ItemCode  &"A"& EntNo &""" value=""Y""  class=""FormElem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"

				set headerCell=oRow.insertCell()
				headerCell.innerHTML=ItemNode.Item(i).Attributes.getNamedItem("ItmDescription").value
				headerCell.className="ExcelDisplayCell"
				headerCell.align="Left"


				set headerCell=oRow.insertCell()
				headercell.innerHtml = UomCode + "<input type=""hidden"" name=""hRateUOM"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value=" & sRateUOM  & " >"
				headercell.innerHtml = headercell.innerHtml + "<input type=""hidden"" name=""hRatePerQtyUoM"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value=" & nRatePerQtyUoM & " >"
				headerCell.width = "10"
				headerCell.className="ExcelFieldCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtInvQty"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& InvQty &" onBlur=""DisplayPurAmount('Q','"& ItemCode &"','"& ClassCode &"','"& EntNo &"',this)"" style=""text-align: Right"" size=""10"" class=""Formelem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"
				headerCell.align="center"


				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtInvRate"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& FormatNumber(Rate,5,,,0) &" onBlur=""DisplayPurAmount('R','"& ItemCode &"','"& ClassCode &"','"& EntNo &"',this)""  style=""text-align: Right"" size=""11"" class=""Formelem"">")
				headerCell.appendChild(oText)
				headercell.innerHtml = headercell.innerHtml + "&nbsp;&nbsp;<a><Img border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" onclick=""SetRateUOM('"& trim(ClassCode) & "','" & trim(ItemCode) &"','"& trim(EntNo) &"','" & UomCode & "','MP','Y')"" alt=""Select Rate UOM"" style=""cursor:hand""></a>"
				headerCell.width = "100"
				headerCell.className="ExcelInputCell"
				headerCell.align="center"


				nAmt = cdbl(InvQty) * cdbl(Rate)

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtValue"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value=" & FormatNumber(nAmt,2,,,0) & " size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="FormelemRead"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtDisPer"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& FormatNumber(DisPer,2,,,0) &" onBlur=""DisplayPurAmount('D','"& ItemCode &"','"& ClassCode &"','"& EntNo &"',this)"" style=""text-align: Right"" size=""6"" class=""Formelem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtDisAmount"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& FormatNumber(DisAmount,2,,,0) &" ReadOnly maxlength=13  style=""text-align: Right"" class=""FormelemRead"" size=""11"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtNetValue"& ClassCode &"Z"& ItemCode &"Z"& EntNo &""" value="& FormatNumber(NettBasic,2,,,0) &" size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"


				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" readonly name=""txtItemRateDisplay" & ClassCode & "Z" & ItemCode &"Z"& EntNo & """ value=" & FormatNumber(nItemValue,2,,,0) & " size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox"" name=""chkVat" & ClassCode & "Z" & ItemCode &"Z"& EntNo & """ class=""Formelem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				iCtr=iCtr+1

				nTotalBasicValue =  CDbl(nTotalBasicValue) + CDbl(nAmt)
				nTotalDiscValue  =  CDbl(nTotalDiscValue)  + CDbl(DisAmount)
				nTotalNettBasic  =  CDbl(nTotalNettBasic)  + CDbl(NettBasic)

		next

		if ItemNode.Length>0 then

			'To display Total
			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.colspan = "6"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml="<B>Total</B>"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" ReadOnly name=""txtBasicValue"& PurType &""" size=""11"" value=" & FormatNumber(nTotalBasicValue,2,,,0) & " style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"

			set headerCell=oRow.insertCell()
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text""  ReadOnly  name=""txtDisValue"& PurType &""" size=""11"" value=" & FormatNumber(nTotalDiscValue,2,,,0) & " style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text""  ReadOnly  name=""txtAmount"& PurType &""" size=""11"" value=" & FormatNumber(nTotalNettBasic,2,,,0) & " style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text""  ReadOnly  name=""txtTotalItemRate"& PurType &"Z"& EntNo &""" size=""11"" value=" & FormatNumber(dTotItemValue,2,,,0) & "  style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"

			'Adding Total And subTotal
			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.colspan = "9"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml="Total Tax Value"
			set oText = document.createElement("<img name=""imgTaxDet"& PurType &""" border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" alt=""View Tax Details"" onclick=""ShowTaxDet('"& PurType  &"')"" width=""15"" height=""15"" style=""cursor:hand"">")
			headerCell.appendChild(oText)

			headerCell.align="Right"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" NAME=""txtTot"& PurType &""" style=""text-align: Right"" value="& FormatNumber(nTotalTaxValue,2,,,0) & " size=""11"" style=""text-align: Right""  class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelInputCell"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" NAME=""txtTaxDisplay"& PurType &""" style=""text-align: Right"" size=""11"" readonly class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Right"

			'set headerCell=oRow.insertCell()
			'headerCell.width = "10"
			'headerCell.className="ExcelDisplayCell"
			'headerCell.align="Right"

			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.colspan = "9"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml="Sub Total"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text""  ReadOnly  NAME=""txtSubTot"& PurType &""" value=" & FormatNumber(nSubTotal,2,,,0)  & " style=""text-align: Right"" size=""11"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Right"

		end if
	next
end if

'Adding InvoiceValue and Roundoff
	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Rounded Off</B>("

	set oText = document.createElement("<input type=""radio"" value=""Y"" checked=true onclick=""RoundOffInv()"" name=""rdRndOff"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml=headerCell.innerhtml + "Yes"

	set oText = document.createElement("<input type=""radio"" value=""N"" onclick=""RoundOffInv()"" name=""rdRndOff"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml = headerCell.innerhtml + "No)"
	headerCell.align="Right"


	'nRoundOffValue = 0
	'nInvValue = 0
	'if Root.hasChildNodes then
	'	Set TaxNode=Root.Selectnodes("//InvoiceHeader/Header")
	'	for j = 0 to TaxNode.Length - 1
	'		nRoundOffValue=TaxNode.Item(j).Attributes.getNamedItem("RoundOff").value
	'		nInvValue=TaxNode.Item(j).Attributes.getNamedItem("InvValue").value
	'		exit for
	'	next
	'end if 'if Root.hasChildNodes then

	nRoundOffValue = round(nInvValue,0) - nInvValue
	nInvValue = round(nInvValue,0)

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" NAME=""txtRoundOff"" ReadOnly onBlur=""AssignRoundOffValue()"" style=""text-align: Right"" size=""11""  class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	'set headerCell=oRow.insertCell()
	'headerCell.width = "10"
	'headerCell.className="ExcelDisplayCell"
	'headerCell.align="Right"

	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Invoice Value</B>"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  ReadOnly  NAME=""txtInvValue"" value=" & FormatNumber(nInvValue,2,,,0) &  " style=""text-align: Right"" size=""11"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  ReadOnly  Name=""mTxtTotal""  value=" & FormatNumber(nTotal,2,,,0) &  " size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	document.all.ImgDeleteIcon.disabled= false


	'note : the following block is used to calculate ItemRate based on InvoiceRate Column
	'******* block ***
	if Root.hasChildNodes then

		Set TaxNode=Root.Selectnodes("//TaxDetails")
		for j = 0 to TaxNode.Length - 1

			PurType=TaxNode.Item(j).Attributes.getNamedItem("PurchaseType").value

			Set ItemNode=Root.Selectnodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")
			for i = 0 to ItemNode.Length - 1

				ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
				ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
				UomCode=ItemNode.Item(i).Attributes.getNamedItem("Uom").value
				EntNo = ItemNode.Item(i).Attributes.getNamedItem("EntryNo").value

				SetRateUOM trim(ClassCode),trim(ItemCode),trim(EntNo),UomCode,"MP","N"
			next
		next
	end if 'if Root.hasChildNodes then
	'******* block ***

end function
'-------------------------------------------------------------------------------------------
function SetRateUOM(iClassCode,iItemCode,EntNo,sUOM,sPurchaseTypeNature,sShowWindow)
Dim sOrgID,sRateUOM

Dim nRatePerQtyUOM

sOrgID = document.formname.hOrgID.value

set obj = eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode)& "Z" & trim(EntNo) )
nRatePerQtyUOM = obj.value

set obj = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode)& "Z" & trim(EntNo) )
sRateUOM = obj.value

if trim(sRateUOM) = "" then sRateUOM = sUOM

if trim(sShowWindow) = "Y" then
	Set ndSample= showModalDialog("invPurInvEntry_RateUomPop.asp?hOrgID=" & trim(sOrgID) & "&hClassCode="+trim(iClassCode)+"&hItemCode="+trim(iItemCode)+"&hEntNo="+trim(EntNo)+"&hUOM=" + trim(sUOM)+"&hRateUOM=" + trim(sRateUOM) +"&hRatePerQtyUOM="+trim(nRatePerQtyUOM),"","dialogHeight:250px;dialogWidth:300px;status:no")
	'alert(ndSample.xml)

	set Root = ndSample
	'alert(Root.getAttribute("RateUOM"))
	'alert(Root.getAttribute("RatePerQtyUoM"))

	sRateUOM = ""
	nRatePerQtyUoM = 0.00

	if trim(Root.getAttribute("RateUOM")) <> "" then
		sRateUOM = Root.getAttribute("RateUOM")
		nRatePerQtyUoM = Root.getAttribute("RatePerQtyUoM")
	else
		sRateUOM = sUOM
		nRatePerQtyUoM = eval("document.formname.txtInvRate"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value
	end if
else
	' this part is used to calculate ItemRate while loading
	if trim(eval("document.formname.hRateUOM"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value) <> "" then
		sRateUOM = eval("document.formname.hRateUOM"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value
		nRatePerQtyUoM = eval("document.formname.hRatePerQtyUoM"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value
	else
		sRateUOM = sUOM
		nRatePerQtyUoM = eval("document.formname.txtInvRate"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value
	end if
end if 'if trim(sShowWindow) = "Y" then

set obj = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo) )
obj.value = sRateUOM
set obj = eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo) )
obj.value = nRatePerQtyUoM

'alert(sRateUOM)
'alert(nRatePerQtyUoM)

getQtyUoMRate sOrgID,iClassCode,iItemCode,EntNo,eval("document.formname.hRateUoM"&iClassCode&"Z"&iItemCode&"Z"&EntNo),sUOM

set obj = eval("document.formname.txtInvRate"&iClassCode&"Z"&iItemCode&"Z"&EntNo)
'alert(Obj.value)
'alert(sPurchaseTypeNature)
if UCase(Trim(sPurchaseTypeNature)) = "SP" then 'single purchase type
	DisplayAmount "R",iItemCode,iClassCode,EntNo,obj
else	' item wise purchase type
	DisplayPurAmount "R",iItemCode,iClassCode,EntNo,obj
end if
end function
'-------------------------------------------------------------------------------------------

Function getQtyUoMRate(ORGID,CLASSCODE,ITEMCODE,ENTNO,OBJ,sUOM)

Dim iOptToBaseRate,iOptToBaseOperator,QTYUOM,RATE,iCtr,dQTY,dRatePerQtyUoM,dRateTotalQtyUoM

'QTYUOM = eval("document.formname.cmbUoM"&CLASSCODE&"Z"&ITEMCODE).value
QTYUOM = sUOM
arrUoM = split(QTYUOM,":")
QTYUOM = arrUoM(0)

RATE = eval("document.formname.txtInvRate"&CLASSCODE&"Z"&ITEMCODE&"Z"&ENTNO).value
dQTY = eval("document.formname.txtInvQty"&CLASSCODE&"Z"&ITEMCODE&"Z"&ENTNO).value

RATEUOM = OBJ.value
if trim(RATEUOM) <> "" then
	arrUoM = split(RATEUOM,":")
	RATEUOM = arrUoM(0)
end if 'if trim(RATEUOM) <> "" then

'alert(QTYUOM)
'alert(RATEUOM)

'if ucase(CLASSCODE) = "TEMP" then CLASSCODE = 0

if trim(RATE) = "" then RATE = 0

'if (CLASSCODE = "0" or CLASSCODE="TEMP") and (trim(RATEUOM) <> trim(QTYUOM)) then
'	msgbox "Quantity UoM and Rate UoM should be same for Temporary Items",0,"Order"
'	obj.focus()
'	exit function
'end if

if (CLASSCODE = "0" or CLASSCODE="TEMP") then
	dRatePerQtyUoM = RATE
else
	If trim(RATEUOM) <> trim(QTYUOM) Then
		dRatePerQtyUoM = getRatePerQtyUoM(ORGID,CLASSCODE,ITEMCODE,ENTNO,QTYUOM,RATEUOM,RATE)
		'alert("test ....")
	Else
		dRatePerQtyUoM = RATE
	End If
end if

'alert(dRatePerQtyUoM)

eval("document.formname.txtInvRate"&CLASSCODE&"Z"&ITEMCODE&"Z"&ENTNO).value  = formatnumber(dRatePerQtyUoM,5,,,0)

End function
'-------------------------------------------------------------------------------------------
Function DisplayAmount(sPara,iItemCode,iClassCode,EntNo,objText)

dim dBasicTotal,dDisTotal,dNetTotal,dRate,dQty,dDisPer

	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
	next

	For Each oNodTemp in ItemRoot.childNodes
		sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value )
	Next

	'alert(sCurrentItemPurType)
	if trim(sCurrentItemPurType) = "" or trim(sCurrentItemPurType) = "0" then
		Alert("Enter Purchase Details")
		exit function
	end if

	dBasicTotal	= 0
	dDisTotal	= 0
	dNetTotal	= 0

	dQty = 0
	dRate = 0.00
	dDisPer = 0.00

	if sPara = "Q" then dQty = objText.value
	if sPara = "R" then dRate = objText.value
	if sPara = "D" then dDisPer = objText.value

	if trim(objText.value)<>"" then

		if IsNumeric(objText.value) then
			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.getAttribute("EntryNo") = trim(EntNo) and oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode then

					if sPara = "Q" then oNodTemp.SetAttribute "Qty",dQty
					if sPara = "R" then
						oNodTemp.SetAttribute "Rate",FormatNumber(dRate,5,,,0)
						if trim(eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo)).value) <> "" then
							oNodTemp.SetAttribute "RatePerQtyUoM",	eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo)).value
							sRateUOM = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo)).value
							if trim(sRateUOM) <> "" then
								Arr1 = Split(sRateUOM,":")
								sRateUOM = Arr1(0)
							end if
							oNodTemp.SetAttribute "RateUOM", sRateUOM
						else
							oNodTemp.SetAttribute "RatePerQtyUoM",FormatNumber(dRate,5,,,0)
							oNodTemp.SetAttribute "RateUOM", oNodTemp.getAttribute("Uom")
						end if
					end if 'if sPara = "R" then
					if sPara = "D" then oNodTemp.SetAttribute "DisPer",dDisPer


					dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
					dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
					dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

					if dQty > 0 and dRate > 0  then
						dBasicAmount=CDbl(dQty)*CDbl(dRate)

						dDisAmount = Round((dDisPer/100)* dBasicAmount,2)

						dNetAmount = Round(dBasicAmount - dDisAmount,2)

						oNodTemp.SetAttribute "DisAmount",dDisAmount
						oNodTemp.SetAttribute "NettBasic",dNetAmount

						dBasicAmount=FormatNumber(Round(dBasicAmount,2),2,,,0)
						dDisAmount=FormatNumber(dDisAmount,2,,,0)
						dNetAmount=FormatNumber(dNetAmount,2,,,0)

						eval("document.formname.txtValue"&iClassCode&"Z"&iItemCode&"Z"&trim(EntNo)).value=dBasicAmount
						eval("document.formname.txtDisAmount"&iClassCode&"Z"&iItemCode&"Z"&trim(EntNo)).value=dDisAmount
						eval("document.formname.txtNetValue"&iClassCode&"Z"&iItemCode&"Z"&trim(EntNo)).value=dNetAmount
					end if 'if dQty > 0 and dRate > 0  then


				end if
			next

			For Each oNodTemp in ItemRoot.childNodes

				dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
				dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
				dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

				if dQty > 0 and dRate > 0  then
					dBasicAmount=CDbl(dQty)*CDbl(dRate)

					dDisAmount = Round((dDisPer/100)* dBasicAmount,2)

					dNetAmount = Round(dBasicAmount - dDisAmount,2)

					dBasicTotal	= dBasicTotal + dBasicAmount
					dDisTotal	= dDisTotal + dDisAmount
					dNetTotal	= dNetTotal + dNetAmount
				end if 'if dQty > 0 and dRate > 0  then

			next

			document.formname.txtBasicValue.value	= FormatNumber(Round(dBasicTotal,2),2,,,0)
			document.formname.txtDisValue.value		= FormatNumber(Round(dDisTotal,2),2,,,0)
			document.formname.txtAmount.value		= FormatNumber(Round(dNetTotal,2),2,,,0)

			popTax

			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.getAttribute("EntryNo") = trim(EntNo) and oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode then
					CalcItemValueForEachItem iItemCode,iClassCode,EntNo
					if sPara = "R" then
						objText.value= FormatNumber(trim(objText.value),5,,,0)
					elseif sPara = "D" then
						objText.value= FormatNumber(trim(objText.value),2,,,0)
					end if
				end if
			next

		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if 'if IsNumeric(objText.value) then

	end if 'if trim(objText.value)<>"" then
End Function
'-------------------------------------------------------------

Function DisplayPurAmount(sPara,iItemCode,iClassCode,EntNo,objText)
dim sCurrentItemPurType
dim dBasicTotal,dDisTotal,dNetTotal,dRate,dQty,dDisPer


	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
	next

	dBasicTotal	= 0
	dDisTotal	= 0
	dNetTotal	= 0

	dQty = 0
	dRate = 0.00
	dDisPer = 0.00

	if sPara = "Q" then dQty = objText.value
	if sPara = "R" then dRate = objText.value
	if sPara = "D" then dDisPer = objText.value

	if trim(objText.value)<>"" then

		if IsNumeric(objText.value) then
			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.getAttribute("EntryNo") = trim(EntNo) and oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode then

					sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value )

					if sPara = "Q" then oNodTemp.SetAttribute "Qty",dQty
					if sPara = "R" then
						oNodTemp.SetAttribute "Rate",FormatNumber(dRate,5,,,0)
						if trim(eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo)).value) <> "" then
							oNodTemp.SetAttribute "RatePerQtyUoM",	eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo)).value
							sRateUOM = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode) & "Z" & trim(EntNo)).value
							if trim(sRateUOM) <> "" then
								Arr1 = Split(sRateUOM,":")
								sRateUOM = Arr1(0)
							end if
							oNodTemp.SetAttribute "RateUOM", sRateUOM
						else
							oNodTemp.SetAttribute "RatePerQtyUoM",FormatNumber(dRate,5,,,0)
							oNodTemp.SetAttribute "RateUOM", oNodTemp.getAttribute("Uom")
						end if
					end if 'if sPara = "R" then
					if sPara = "D" then oNodTemp.SetAttribute "DisPer",dDisPer


					dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
					dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
					dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

					if dQty > 0 and dRate > 0  then
						dBasicAmount=CDbl(dQty)*CDbl(dRate)

						dDisAmount = Round((dDisPer/100)* dBasicAmount,2)

						dNetAmount = Round(dBasicAmount - dDisAmount,2)

						oNodTemp.SetAttribute "DisAmount",dDisAmount
						oNodTemp.SetAttribute "NettBasic",dNetAmount

						dBasicAmount=FormatNumber(Round(dBasicAmount,2),2,,,0)
						dDisAmount=FormatNumber(dDisAmount,2,,,0)
						dNetAmount=FormatNumber(dNetAmount,2,,,0)


						eval("document.formname.txtValue"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value=dBasicAmount
						eval("document.formname.txtDisAmount"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value=dDisAmount
						eval("document.formname.txtNetValue"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value=dNetAmount



					end if 'if dQty > 0 and dRate > 0  then


				end if
			next

			For Each oNodTemp in ItemRoot.childNodes

				If sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value ) then

					dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
					dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
					dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

					if dQty > 0 and dRate > 0  then
						dBasicAmount=CDbl(dQty)*CDbl(dRate)

						dDisAmount = Round((dDisPer/100)* dBasicAmount,2)

						dNetAmount = Round(dBasicAmount - dDisAmount,2)

						dBasicTotal	= dBasicTotal + dBasicAmount
						dDisTotal	= dDisTotal + dDisAmount
						dNetTotal	= dNetTotal + dNetAmount
					end if 'if dQty > 0 and dRate > 0  then

				end if ' if sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value )

			next

			eval("document.formname.txtBasicValue" &sCurrentItemPurType).value = FormatNumber(Round(dBasicTotal,2),2,,,0)
			eval("document.formname.txtDisValue" &sCurrentItemPurType).value = FormatNumber(Round(dDisTotal,2),2,,,0)
			eval("document.formname.txtAmount" &sCurrentItemPurType).value = FormatNumber(Round(dNetTotal,2),2,,,0)



			popTax1

			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.getAttribute("EntryNo") = trim(EntNo) and oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode then
					CalcItemValueForEachItem iItemCode,iClassCode,EntNo
					if sPara = "R" then
						objText.value= FormatNumber(trim(objText.value),5,,,0)
					elseif sPara = "D" then
						objText.value= FormatNumber(trim(objText.value),2,,,0)
					end if
				end if
			next

		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if 'if IsNumeric(objText.value) then

	end if 'if trim(objText.value)<>"" then
End Function
'-------------------------------------------------------------------------------------------
FUNCTION setTaxPercentage(sCatCode,sTaxCode,objText)
	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			FOR EACH Node2 in oNodTemp.ChildNodes
				if Node2.NodeName = "Header" then
					sPurchaseType	= Node2.getAttribute("PurchaseType")
				end if 'if Node2.NodeName = "Header" then
			next
		end if

		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For Each oNodTemp in TaxRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then
					oNodTemp.Attributes.Item(4).nodeValue=objText.value
					popTax
					CalculateItemValue sPurchaseType
					exit function
				end if
			Next
		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if
	end if
END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION setTaxAmount(sCatCode,sTaxCode,objText)

	Dim sTaxCatType

	Set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			FOR EACH Node2 in oNodTemp.ChildNodes
				if Node2.NodeName = "Header" then
					sPurchaseType	= Node2.getAttribute("PurchaseType")
				end if 'if Node2.NodeName = "Header" then
			next
		end if

		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	Next

	dInvAmount=document.formname.txtAmount.value
	dBasicTotal=document.formname.txtBasicValue.value
	dTotal	=document.formname.txtAmount.value

	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For Each oNodTemp in TaxRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then

					sTaxMode = oNodTemp.Attributes.Item(2).nodeValue
					sFormula = oNodTemp.Attributes.Item(3).nodeValue
					sTaxCatType = oNodTemp.getAttribute("TaxCategoryType")

					if sTaxMode="F" then
						'if sTaxCetType <> "C" and trim(objtext.value) < 0 then
						'	MsgBox ("Enter Numeric Value")
						'	objText.select
						'	exit function
						'Else
							oNodTemp.Attributes.Item(4).nodeValue=objText.value
							oNodTemp.Attributes.Item(5).nodeValue=objText.value
						'End if
					elseif sTaxMode="P" then
						'alert(sTaxMode)
						'alert(sFormula)
						'alert(dBasicTotal)
						'alert(dTotal)
						'alert(objText.value)
						'oNodTemp.Attributes.Item(4).nodeValue = calPercentage(sFormula,dBasicTotal,dTotal,objText.value)

					else ' K & Q case
						oNodTemp.Attributes.Item(4).nodeValue=objText.value
						oNodTemp.Attributes.Item(5).nodeValue=objText.value
					end if
					popTax
					CalculateItemValue sPurchaseType
					exit function
				end if
			Next

		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if
	end if
END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION popTax()
dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax
dim nDisplayTotal

	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	dInvAmount=document.formname.txtAmount.value
	dBasicTotal=document.formname.txtBasicValue.value
	dTotal	=document.formname.txtAmount.value

	if trim(dInvAmount) = "" then dInvAmount = 0
	if trim(dBasicTotal) = "" then dBasicTotal = 0
	if trim(dTotal) = "" then dTotal = 0

	nDisplayTotal = CDbl(dTotal)

	For Each oNodEntry in TaxRoot.childNodes

		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		sFormula=oNodEntry.Attributes.Item(3).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue

		If sCatCode <> "0" and sTaxCode <> "0" then		' Except round off node - 09 Aug 04

			' Added on 16-Jun-04 - should test
			if sTaxMode="P" then
				dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue,0)
				eval("document.formname.txtTaxPer"&sCatCode&sTaxCode).value=dTaxValue
			elseif sTaxMode = "Q" then
				dTax = cdbl(dTaxvalue) * cdbl(Qtysum)
			elseif sTaxMode = "K" then
				sTotpackvalue = 0
				sExp = "//Tax[@CatCode = "&sCatCode&" and @TaxCode ="&sTaxCode&"]/Taxpack"
				set sNode = TaxRoot.Selectnodes(sExp)
				If sNode.Length > 0 then
					For IMCtr = 0 to sNode.Length - 1
						'Msgbox sNode.Item(iMCtr).Attributes.Item(3).nodevalue
						sTotpackvalue = cdbl(sTotpackvalue) + cdbl(sNode.Item(iMCtr).Attributes.Item(3).nodevalue)
					Next
				Else
					sTotpackvalue = cdbl(sTotpackvalue) + cdbl(oNodEntry.Attributes.Item(5).nodevalue)
				end if
				dTax=sTotpackvalue
			else
				dTax=dTaxvalue
			end if

			dTRndoff = oNodEntry.getAttribute("Rndoff")

			if trim(dTRndoff) = "1" then
				dtax = RndOff(dTax)
			End if


			dTax=FormatNumber(dTax,2,,,0)
			oNodEntry.Attributes.Item(5).nodeValue=dTax

			dInvAmount=dInvAmount+CDbl(dTax)

			eval("document.formname.txtTaxValue"&sCatCode&"Z"&sTaxCode).value=dTax

			nDisplayTotal = nDisplayTotal + dTax
			eval("document.formname.txtSubTaxValue"&sCatCode&sTaxCode).value=FormatNumber(nDisplayTotal,2,,,0)

		end if
	Next

	dRoundedInvvalue = RndOff(dInvamount)
	dRoundedoff = Round(cdbl(dRoundedInvvalue) - cdbl(dInvamount),2)

	dRoundedoff = FormatNumber(dRoundedoff,2,,,0)
	dRoundedInvvalue=FormatNumber(dRoundedInvvalue,2,,,0)

	'alert(dRoundedoff)
	if cdbl(dRoundedoff) <> 0.00 then
		document.formname.rdRndOff(0).Checked=true
		document.formname.txtRoundOff.value = dRoundedoff
		document.formname.txtInvValue.value = dRoundedInvvalue
	else
		document.formname.rdRndOff(1).Checked=true
		document.formname.txtRoundOff.value="0.00"
		document.formname.txtInvValue.value = formatnumber(dInvamount,2,,,0)
	end if

	' set value for round off node - 09 Aug 04
	For Each oNodEntry in TaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue

		If sCatCode = "0" and sTaxCode = "0" then
			oNodEntry.setAttribute "TaxValue",document.formname.txtRoundOff.value
			oNodEntry.setAttribute "TaxAmount",document.formname.txtRoundOff.value
			'msgbox document.formname.txtRoundOff.value
		End if
	Next

	TaxRoot.Attributes.Item(0).nodeValue=dBasicTotal
	TaxRoot.Attributes.Item(1).nodeValue=dTotal
	'TaxRoot.Attributes.Item(2).nodeValue=dRoundedInvvalue
	'TaxRoot.Attributes.Item(3).nodeValue=dRoundedoff
	TaxRoot.setAttribute "TotalTax", dInvamount-dTotal
	TaxRoot.setAttribute "SubTotal", dInvamount
'''	TaxRoot.setAttribute "InvAmtWithoutRoundOff", dInvamount
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set TempNode=oNodTemp
			exit for
		end if
	next

	for Each oNodTemp in TempNode.childNodes
		'oNodTemp.setAttribute "InvValue",dRoundedInvvalue
		'oNodTemp.setAttribute "RoundOff",dRoundedoff
		oNodTemp.setAttribute "InvValue",document.formname.txtInvValue.value
		oNodTemp.setAttribute "RoundOff",document.formname.txtRoundOff.value
	next

END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION CalculateTax(sFormula,dBValue,dDValue,dPercentage,sPurType)
dim saTemp,dTaxAmount,iCounter,TaxRoot
dim oNodTemp,iTemp
dim saTemp1

set RootNode=InvoiceDet.documentElement
if trim(sPurType) ="0" then
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
			exit for
		end if
	next
else
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			if trim(oNodTemp.getAttribute("PurchaseType")) = trim(sPurType) then
				set TaxRoot=oNodTemp
				exit for
			end if
		end if
	next
end if 'if trim(sPurType) ="0" then


saTemp=Split(sFormula,",")
if trim(saTemp(0)) = "BV" then
	dTaxAmount = dBValue
	iTemp = 1
elseif trim(saTemp(0))="BD" then
	dTaxAmount = dDValue
	iTemp = 1
else
	dTaxAmount = 0
	iTemp = 0
end if
for iCounter = iTemp to UBound(saTemp)
	saTemp1=Split(trim(saTemp(iCounter)),"#")
	For Each oNodTemp in TaxRoot.childNodes
		if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
			dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(5).nodeValue)
			'alert("testing ....")
			'alert(oNodTemp.Attributes.Item(5).nodename)
		end if
	next
next
'CalculateTax= formatnumber(dTaxAmount*(cdbl(dPercentage)/100),2,,,0)

if trim(dPercentage)<>"" then
	CalculateTax=dTaxAmount*(cdbl(dPercentage)/100)
else
	CalculateTax=dTaxAmount
end if

END FUNCTION
'-------------------------------------------------------------------------------------------

Function ClearTable()
	dim j
	for j=2 to document.all.tblItemDet.rows.length - 1
		document.all.tblItemDet.deleteRow(2)
	next
End Function
'-------------------------------------------------------------------
function getTaxDet()
	Dim ndEntry,ndTax,sTaxName,sFormName,iTaxCode,iCatCode,iFormNo,dPerAmt,sMode
	Dim hPurchaseType,objhttp,sForUnit,objhttp1
	dim RtTax,root,DeleteNode,ItemNode
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root=InvoiceDet.documentElement

	'To add Tax details node
	Set DeleteNode = Root.Selectnodes("//TaxDetails")
	for i=0 to (DeleteNode.Length-1)
		root.RemoveChild(DeleteNode.Item(i))
	next
	sForUnit = document.formname.hOrgID.value
	hPurchaseType = document.Formname.cmbPurType.value

	if hPurchaseType=0 then
		document.all.imgPurchaseDet.disabled=false
	else
		document.all.imgPurchaseDet.disabled=true
	end if
	nInvNo=document.formname.hInvNo.value
	sRcptNum = document.formname.hRcptNum.value

	objhttp.Open "GET","XMLGetTaxDetails.asp?PurType="&hPurchaseType&"&ForUnit="&sForUnit&"&InvNo="&nInvNo&"&RcptNum="&sRcptNum,false
	objhttp.send

	If objhttp.responsexml.xml <> "" then
		'msgbox "objhttp.responseXML.xml"+objhttp.responseXML.xml
		TaxFormData.loadXML objhttp.responseXML.xml
		Set RtTax = TaxFormData.documentElement
		If RtTax.hasChildNodes Then
			Set Root=InvoiceDet.documentElement
			For Each ndEntry in RtTax.childNodes
				Root.appendchild ndEntry
			next
		end if
	end if

	'To set Puchasetype for each Item
	if Root.hasChildNodes then

		Set ItemNode = Root.Selectnodes("//InvoiceHeader/Header")
		for i = 0 to ItemNode.Length - 1
			ItemNode.Item(i).Attributes.getNamedItem("PurchaseType").value= hPurchaseType
		next

		Set ItemNode = Root.Selectnodes("//ItemDetails/Item")
		for i = 0 to ItemNode.Length - 1
			ItemNode.Item(i).Attributes.getNamedItem("PurchaseType").value= hPurchaseType
		next

		Set ItemNode = Root.Selectnodes("//TaxDetails/Tax")
		for i = 0 to ItemNode.Length - 1
			ItemNode.Item(i).setAttribute "ItemValue", "0"
		next
	end if

	set objhttp1 = CreateObject("Microsoft.XMLHTTP")
	objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=AmdNewInvItemValue", false
	objhttp1.send InvoiceDet.XMLDocument

	additem()

end function
'------------------------------------------------------------------

Function showTaxFormPopUp(sOrgId,sInvDt)
Dim ndSample,iPurTypeVal,sPurType,sPartyName,rootTax,ndEntry

Set rootTax = TaxFormData.documentElement

If document.formname.cmbPurType.value  = "0" Then
	msgbox "Select Purchase Type"
	document.formname.cmbPurType.focus()
	Exit Function
Else
	iPurTypeVal = document.formname.cmbPurType.selectedIndex
	sPurType	= document.formname.cmbPurType(document.formname.cmbPurType.selectedIndex).text
	sPartyName	= document.formname.txtPartyName.value
	Set ndSample= showModalDialog("invPurTaxFormPop.asp?OrgID="+sOrgId+"&PurType="+cstr(iPurTypeVal)+"&PurTypeName="+cstr(sPurType)+"&Party="+sPartyName+"&InvDt='"+cstr(sInvDt)+"'","","dialogHeight:500px;dialogWidth:900px;status:no")
	If ndSample.hasChildNodes Then
		For each ndEntry in ndSample.childNodes
			rootTax.appendchild ndEntry
		Next
	End If
End If
End Function
'----------------------------------------------------------------------------------------
Function ShowPurchaseDet(sOrgID,sRefNum)
dim ndTemp,objhttp1,NoOfItems
set RootNode=InvoiceDet.DocumentElement
set objhttp1 = CreateObject("Microsoft.XMLHTTP")
objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=AmdNewInvItemValue", false
objhttp1.send InvoiceDet.XMLDocument
'alert "saved"
'Set ndTemp =showModalDialog("popItemPurType.asp?Org="+ sOrgID +"&ItemType="+ sItemType +"&RefNum="+sRefNum,InvoiceDet,"","dialogHeight:435px;dialogWidth:700px;Status:no;help:no")
Set ndTemp =showModalDialog("popItemPurType.asp",InvoiceDet,"dialogHeight:400px;dialogWidth:450px;status:no;help:no")
set InvoiceDet.DocumentElement=ndTemp
'alert( InvoiceDet.xml)
PurchaseTypeWise
popTax1
end function

'-----------------------------------------------------------------------------------------
Function ShowTaxDet(PurType)
'msgBox "In ShowTax"
dim ndTemp,ItemNode
Set InvoiceDet.DocumentElement = showModalDialog("popTaxDetails.asp?PurType="+ PurType,InvoiceDet,"dialogLeft:0px;dialogTop:0Px;dialogHeight:400px;dialogWidth:500px;status:no")
popTax1
CalculateItemValue(PurType)
end function

'-----------------------------------------------------------------------------------------

Function getReferenceDetail()
bFlag = False
sInvAgainst = ""
sInvRefNum = ""
if SpnRcptCode.innerText <> "" then
	bFlag = True
	sInvAgainst = "Receipt"
	'sInvRefNum = document.formname.cmbReceipt.value
	sInvRefNum = document.formname.hRcptno.value
''	msgbox sInvRefNum
End if
End Function
'------------------------------------------------------------------------------------------
Function displayItemDetail(sOrgMD,sFlag)
Dim objHttp,Root,ndHeader,ndItem
Dim sPartyName, dQty
Dim iClassCode,iItemCode,iQtyRecd,sClassDesc,sItemDesc

getReferenceDetail()

If bFlag = False Then
	msgbox "Select Reference Number",0,"Invoice"
	exit function
End if
''msgbox sFlag

Set objhttp = CreateObject("MSXML2.XMLHTTP")
objhttp.Open "GET","XMLGetInvItem.asp?Flag="&sFlag&"&sOrgID="&sOrgID&"&InvoiceAgainst="& sInvAgainst& "&RefNum=" & sInvRefNum ,false
objhttp.send


If objhttp.responseXML.xml <> "" then
	OutData.loadXML objhttp.responseXML.xml
	Set Root = OutData.documentElement
	if Root.HaschildNodes() then

	'' To fill in Party type select box based on party code selection
	For Each ndHeader In Root.childNodes
		if ndHeader.nodename = "PartyType" then
			document.formname.cmbPartyType.length = 0
			For each ndTemp in ndHeader.childnodes
				sPartyType = ndTemp.getAttribute("Type")
				spartySubType = ndTemp.getAttribute("SubTypeCode")
				sPartySubtypeName = ndTemp.getAttribute("SubTypeName")

				document.formname.cmbPartyType.length = document.formname.cmbPartyType.length+1

				document.formname.cmbPartyType.options(document.formname.cmbPartyType.length-1).text = sPartySubtypeName
				document.formname.cmbPartyType.options(document.formname.cmbPartyType.length-1).Value = cstr(spartySubType) + "|" + cstr(sPartyType)

			Next

		End if
	Next

	For Each ndHeader In Root.childNodes

		if ndHeader.nodename = "Header" then

			document.formname.hPartyCode.value = ndHeader.Attributes.Item(3).nodeValue
			sPartyName = ndHeader.Attributes.Item(4).nodeValue
			'document.formname.hPartyType.value = ndHeader.Attributes.Item(5).nodeValue
			'document.formname.hPSubType.value = ndHeader.Attributes.Item(6).nodeValue
			document.formname.cmbPartyType.value = ndHeader.Attributes.Item(5).nodeValue&"|"&ndHeader.Attributes.Item(6).nodeValue
			document.formname.txtPartyName.value = sPartyName

			if trim(ndHeader.Attributes.Item(7).nodeValue) <> "" and trim(ndHeader.Attributes.Item(7).nodeValue)<>"0" then
				document.formname.txtSuppInvNo.value  =  ndHeader.Attributes.Item(7).nodeValue
				document.formname.txtSuppInvDT.value  =  ndHeader.Attributes.Item(8).nodeValue

			else
				document.formname.txtSuppInvNo.readOnly = false
				'document.formname.txtSuppInvDT.readOnly = false
				document.formname.txtSuppInvNo.className = "formelem"
				'document.formname.txtSuppInvDT.className = "formelem"
				'document.formname.ctlDate.setdate = date()
				document.formname.ctlDate.enable = true
			end if

		End if
	Next
	end if
End if
End Function
'-------------------------------------------------------------------------------------------
'=====================================================================================================
'-------------------------------------------------------------------------------------------
FUNCTION popTax1()
dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax,TaxNode,i,j,PurType,ItemNd,TotInvAmount
dim nDisplayTotal,dRndTotInvAmount
set RootNode=Invoicedet.documentElement
TotInvAmount=0
Set TaxNode=RootNode.SelectNodes("//TaxDetails")
for i=0 to TaxNode.length-1
	PurType=TaxNode.item(i).Attributes.getnamedItem("PurchaseType").value
	set ItemNd=RootNode.SelectNodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")

	dInvAmount=eval("document.formname.txtAmount"&PurType).value
	dBasicTotal=eval("document.formname.txtBasicValue"&PurType).value
	dTotal	=eval("document.formname.txtAmount"&PurType).value

	if trim(dInvAmount) = "" then dInvAmount = 0
	if trim(dBasicTotal) = "" then dBasicTotal = 0
	if trim(dTotal) = "" then dTotal = 0

	nDisplayTotal = 0
	nDispTax = 0
'''	nDisplayTotal = CDbl(dTotal)
	set TaxRoot=TaxNode.item(i)
	For Each oNodEntry in TaxRoot.childNodes

		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		sFormula=oNodEntry.Attributes.Item(3).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue
		'alert(dTaxValue)

		If sCatCode <> "0" and sTaxCode <> "0" then		' Except round off node - 09 Aug 04

			' Added on 16-Jun-04 - should test
			if sTaxMode="P" then
				dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue,PurType)
			elseif sTaxMode = "Q" then
				dTax = cdbl(dTaxvalue) * cdbl(Qtysum)
			elseif sTaxMode = "K" then
				sTotpackvalue = 0
				sExp = "//Tax[@CatCode = "&sCatCode&" and @TaxCode ="&sTaxCode&"]/Taxpack"
				set sNode = TaxRoot.Selectnodes(sExp)
				''alert taxroot.xml
				If sNode.Length > 0 then
					For IMCtr = 0 to sNode.Length - 1
						'Msgbox sNode.Item(iMCtr).Attributes.Item(3).nodevalue
						sTotpackvalue = cdbl(sTotpackvalue) + cdbl(sNode.Item(iMCtr).Attributes.Item(3).nodevalue)
					Next
				end if
				dTax=sTotpackvalue
			else
				dTax=dTaxvalue
			end if

			dTRndoff = oNodEntry.getAttribute("Rndoff")

			if trim(dTRndoff) = "1" then
				'alert(" dTRndoff = " + dTRndoff)

				'dtax = Round(dTax,0)
				'alert( dtax)
				dtax = RndOff(dTax)
			End if


			dTax=FormatNumber(dTax,2,,,0)
			oNodEntry.Attributes.Item(5).nodeValue=dTax

			dInvAmount=dInvAmount+CDbl(dTax)
			'alert(dTax)

			nDisplayTotal = nDisplayTotal + dTax

			if trim(oNodEntry.getAttribute("AccHead")) <> "0" then
				nDispTax = nDispTax + dTax
			end if
		end if
	Next

	'msgbox " dInvamount = " & dInvamount
	dRoundedInvvalue = RndOff(dInvamount)
	dRoundedoff = Round(cdbl(dRoundedInvvalue) - cdbl(dInvamount),2)

	dRoundedoff = FormatNumber(dRoundedoff,2,,,0)
	dRoundedInvvalue=FormatNumber(dRoundedInvvalue,2,,,0)

'''	document.formname.txtRoundOff.value = dRoundedoff



	' set value for round off node - 09 Aug 04
	For Each oNodEntry in TaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue

		If sCatCode = "0" and sTaxCode = "0" then
			oNodEntry.setAttribute "TaxValue",document.formname.txtRoundOff.value
			oNodEntry.setAttribute "TaxAmount",document.formname.txtRoundOff.value
			'msgbox document.formname.txtRoundOff.value
		End if
	Next

'''	document.formname.txtInvValue.value=dRoundedInvvalue
	TaxRoot.Attributes.Item(0).nodeValue=dBasicTotal
	TaxRoot.Attributes.Item(1).nodeValue=dTotal
'''	TaxRoot.Attributes.Item(2).nodeValue=dRoundedInvvalue
'''	TaxRoot.Attributes.Item(3).nodeValue=dRoundedoff
'''	TaxRoot.setAttribute "InvAmtWithoutRoundOff", dInvamount
	TaxRoot.setAttribute "TotalTax", FormatNumber(nDisplayTotal,2,,,0)
	TaxRoot.setAttribute "SubTotal", dInvamount

	eval("document.formname.txtTot" & PurType).value=FormatNumber(nDisplayTotal,2,,,0)
	eval("document.formname.txtTaxDisplay" & PurType).value=FormatNumber(nDispTax,2,,,0)

	eval("document.formname.txtSubTot" & PurType).value=FormatNumber(dInvAmount,2,,,0)
	TotInvAmount=TotInvAmount+dInvAmount
next

dRndTotInvAmount = RndOff(TotInvAmount)
dRoundedoff = Round(cdbl(dRndTotInvAmount) - cdbl(TotInvAmount),2)

dRoundedoff = FormatNumber(dRoundedoff,2,,,0)
dRndTotInvAmount=FormatNumber(dRndTotInvAmount,2,,,0)

'alert(dRoundedoff)
if cdbl(dRoundedoff) <> 0.00 then
	document.formname.rdRndOff(0).Checked=true
	document.formname.txtRoundOff.value = dRoundedoff
	document.formname.txtInvValue.value=dRndTotInvAmount
else
	document.formname.rdRndOff(1).Checked=true
	document.formname.txtRoundOff.value = "0.00"
	document.formname.txtInvValue.value= Formatnumber(TotInvAmount,2,,,0)
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set TempNode=oNodTemp
			exit for
		end if
	next

	for Each oNodTemp in TempNode.childNodes
		'oNodTemp.setAttribute "InvValue",dRoundedInvvalue
		'oNodTemp.setAttribute "RoundOff",dRoundedoff
		oNodTemp.setAttribute "InvValue",document.formname.txtInvValue.value
		oNodTemp.setAttribute "RoundOff",document.formname.txtRoundOff.value
	next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
END FUNCTION
'---------------------------------------------------------------------------------
function Next_Click(sPara)
	dim objhttp1
	dim ItemNd,Qty,Rate,i,DisPer,Node,Nd,InvVal,nRetValue

	nRetValue = 0

	If trim(document.formname.txtSuppInvNo.value)  = "" Then
		msgbox "Enter Supplier Invoice No.",0,"Invoice"
		document.formname.txtSuppInvNo.focus()
		Exit Function
	elseif document.formname.cmbBillType.selectedIndex = 0 Then
		MsgBox "Select Bill Type "
		document.formname.cmbBillType.focus()
		exit function
	elseIf datediff("d",document.formname.ctldate.getdate,document.formname.hCurrDate.value) < 0 then
		MsgBox "Invoice date cannot be greater than the current date"
		exit function
	elseif document.formname.cmbPurType.selectedIndex = 0 Then
		MsgBox "Select Purchase Type "
		document.formname.cmbPurType.focus()
		exit function
	elseif document.formname.cmbInvCat.selectedIndex = 0 Then
		MsgBox "Select Category "
		document.formname.cmbInvCat.focus()
		exit function
    elseif document.formname.cmbPartyType.selectedIndex = 0 then
        MsgBox "Select Party Type"
        document.formname.cmbPartyType.focus()
        exit function
	End If


	set RootNode=Invoicedet.documentElement

	set ItemNd=RootNode.SelectNodes("//ItemDetails/Item")

	for i=0 to ItemNd.Length-1
		iClassCode = cdbl(ItemNd.Item(i).Attributes.getNamedItem("ClassificationCode").value)
		iItemCode =  cdbl(ItemNd.Item(i).Attributes.getNamedItem("ItemCode").value)
		EntNo =  cdbl(ItemNd.Item(i).Attributes.getNamedItem("EntryNo").value)
		Qty = cdbl(ItemNd.Item(i).Attributes.getNamedItem("Qty").value)
		Rate = cdbl(ItemNd.Item(i).Attributes.getNamedItem("Rate").value)
		DisPer = cdbl(ItemNd.Item(i).Attributes.getNamedItem("DisPer").value)
		if Qty <= 0 then
			msgbox "Invoice Quantity should be greater than 0"
			nRetValue = 1
		elseif Rate<=0 then
			msgbox "Invoice Rate should be greater than 0"
			nRetValue = 1
		elseif DisPer < 0 then
			msgbox "Discount Percentage should not be less than 0"
			nRetValue = 1
		end if
		if nRetValue = 1 then exit for

		if trim(ItemNd.Item(i).Attributes.getNamedItem("RateUOM").value)= "" then
			ItemNd.Item(i).SetAttribute "RateUOM", ItemNd.Item(i).Attributes.getNamedItem("Uom").value
		end if

		ItemNd.Item(i).SetAttribute "RatePerQtyUoM",eval("document.formname.txtInvRate" & trim(iClassCode) & "Z" & trim(iItemCode)& "Z" & trim(EntNo) ).value
	next

	'msgbox rootnode.xml
	if nRetValue = 0 then
		set TempNode=RootNode.SelectNodes("//InvoiceHeader/Header")
		if TempNode.Length > 0 then
			TempNode.item(0).SetAttribute "InvValue",document.formname.txtInvValue.value
			TempNode.item(0).SetAttribute "RoundOff",document.formname.txtRoundOff.value
			TempNode.item(0).SetAttribute "Remarks",document.formname.mTextAreaRemarks.value

			TempNode.item(0).SetAttribute "SuppInvNo",document.formname.txtSuppInvNo.value
			TempNode.item(0).SetAttribute "SuppInvDt",document.formname.ctldate.getdate

			TempNode.item(0).SetAttribute "InvCategory",document.formname.cmbInvCat.value
			TempNode.item(0).SetAttribute "BillType",document.formname.cmbBillType.value
			sArrValue= split(document.formname.cmbPartyType(document.formname.cmbPartyType.selectedIndex).value,"|")
			TempNode.item(0).SetAttribute "PartyType", sarrValue(1)
			TempNode.item(0).SetAttribute "PartySubType",sArrValue(0)
		end if
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	end if 'if nRetValue = 0 then


	if sPara = "V" then
		Next_Click = true
		 exit function
	end if


	CalcAmountForLastItem

	AddRoundOffNode document.formname.txtRoundOff.value,document.formname.txtInvValue.value
'	alert(InvoiceDet.xml)
'	exit function

	set objhttp1 = CreateObject("Microsoft.XMLHTTP")
	objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=AmdNewInvItemValue", false
	objhttp1.send InvoiceDet.XMLDocument


	if nRetValue = 0 then
		Next_Click = true
	else
		Next_Click = false
		exit function
	end if


	document.formname.action = "AmdInvPurInvoiceEntryAccDetails.asp"
	document.formname.submit()


end function

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Function CalcItemValueForEachItem(iItemCode,iClassCode,EntNo)
	Dim ObjRoot,Node1,nSumOfItemRate,sItemPurType,nTotal
	Dim sPurchaseType


	Set ObjRoot = InvoiceDet.documentElement
	'alert(ObjRoot.xml)

	For Each oNodTemp in ObjRoot.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set Node=oNodTemp
			exit for
		end if
	next
	for each Nd in Node.ChildNodes
		dInvAmount		= Nd.Attributes.getNamedItem("InvValue").Value
		sPurchaseType	= Nd.Attributes.getNamedItem("PurchaseType").Value
	next

	nSumOfItemRate = 0.00
	nTotal = 0.00
	'alert(	"dInvAmount = "& trim(dInvAmount))
	dTemp=0
	sItemPurType = ""
	for each Node1 in ObjRoot.ChildNodes


		if Node1.nodeName = "ItemDetails" then

			dTemp=0

			if sItemPurType = "" then
				For each Node2 in Node1.ChildNodes
					If Node2.getAttribute("EntryNo") = EntNo and Node2.getAttribute("ItemCode") = iItemCode and Node2.getAttribute("ClassificationCode") = iClassCode then
						sItemPurType=Node2.getAttribute("PurchaseType")
					End if 'If Node2.getAttribute("ItemCode") = iItemCode and Node2.getAttribute("ClassificationCode") = iClassCode then
				Next 'For each Node2 in Node1.ChildNodes
			end if 'if sItemPurType = "" then

			'alert(sItemPurType)


			iItemCount = 0
			nPos = 0
			if trim(sItemPurType) <> "" then
				For each Node2 in Node1.ChildNodes
					If trim(Node2.getAttribute("PurchaseType")) = trim(sItemPurType) then
						Node2.setAttribute "Amount", Node2.getAttribute("NettBasic")
						iItemCount = iItemCount + 1
						nPos = Node2.getAttribute("EntryNo")
					End if 'If trim(Node2.getAttribute("PurchaseType")) = trim(sItemPurType) then
				Next 'For each Node2 in Node1.ChildNodes
			end if 'if sItemPurType = "" then


			For each Node2 in Node1.ChildNodes



				If sItemPurType = Node2.getAttribute("PurchaseType") then


					dItemValue=Node2.getAttribute("Amount")

					'msgbox " first = "& trim(dItemValue)

					if cdbl(dItemValue) > 0 then
						set TaxNodes=ObjRoot.SelectNodes("//TaxDetails[@PurchaseType="& sItemPurType &"]")

						if TaxNodes.Length > 0 then

							set oNodTax=TaxNodes.Item(0)

							dTotBasicVal	= oNodTax.getAttribute("Basicvalue")
							dTotNetVal		= oNodTax.getAttribute("NettValue")
							dInvAmount		= oNodTax.getAttribute("SubTotal")

							'alert(	"dInvAmount = "& trim(dInvAmount))

							'Msgbox "dTotBasicVal = " & Trim(dTotBasicVal)
							'Msgbox "dTotNetVal = "  & Trim(dTotNetVal)


								For Each oNodEntry in oNodTax.childNodes


									sCatCode	=	oNodEntry.getAttribute("CatCode")
									sTaxCode	=	oNodEntry.getAttribute("TaxCode")
									sTaxMode	=	oNodEntry.getAttribute("TaxMode")
									sFormula	=	oNodEntry.getAttribute("TaxFormula")
									dTaxValue	=	oNodEntry.getAttribute("TaxValue")

									'alert("dTaxValue = " & trim(dTaxValue) )


									if sTaxMode="F" then
										dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * cdbl(dItemValue)
									elseif sTaxMode="P" then
										dTax=ItemValueCalculateTax(oNodTax,sFormula,dItemValue,dItemValue,dTaxValue,sItemPurType)
									else
										dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * cdbl(dItemValue)
									end if

									saTemp= Split(CStr(dTax),".")
									if UBound(saTemp)>0 then
										oNodEntry.setAttribute "ItemValue", saTemp(0)&"."&left(saTemp(1),2)
									else
										oNodEntry.setAttribute "ItemValue", saTemp(0)
									end if



									'msgbox "dTax = " & trim(dTax)


								Next


								For Each oNodEntry in oNodTax.childNodes
									if cint(oNodEntry.getAttribute("AccHead"))=0 then
										'alert(dItemValue)
										'alert(oNodEntry.getAttribute("ItemValue"))
										if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
											dItemValue=CDbl(dItemValue)+CDbl(oNodEntry.getAttribute("ItemValue"))
										end if
									else
										if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
											dTemp = CDbl(dTemp) + CDbl(oNodEntry.getAttribute("TaxAmount"))
										end if
									end if
								next

							dItemTotal=CDbl(dItemValue)+CDbl(dItemTotal)
							Node2.setAttribute "Amount", dItemValue

							'msgbox " aaa dItemValue = " & trim(dItemValue)
					' block begin
					'set objhttp1 = CreateObject("Microsoft.XMLHTTP")
					'objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=NewInvItemValue", false
					'objhttp1.send InvoiceDet.XMLDocument

					'alert("saved.....")
					'block end

						end if 'if TaxNodes.Length > 0 then
					end if 'if cdbl(dItemValue) > 0 then
				end if 'If sItemPurType = Node2.getAttribute("PurchaseType") then

			Next'FOR EACH Node2 in Node1.ChildNodes

			if nPos > 0 then

				nPos = nPos  - 1

				if  iItemCount > 0 then
					dItemTotal = CDbl(dTemp/iItemCount) + CDbl(dItemTotal)

					'alert("dItemTotal = " + trim(dItemTotal))

					dTemp = Node1.childNodes(nPos).getAttribute("Amount")
					''dTemp = CDbl(dTemp) + Round(CDbl(dInvAmount)-CDbl(dItemTotal),2)
					dTemp = CDbl(dTemp) + CDbl(dInvAmount)-CDbl(dItemTotal)

					Node1.childNodes(nPos).SetAttribute "Amount",dTemp

					'alert(" dTemp = "& trim(dTemp) )
				end if 'if  iItemCount > 0 then
				'' block 1
			end if 'if nPos > 0 then

			For each Node2 in Node1.ChildNodes

				dProdRate		= Node2.getAttribute("Rate")
				iQty			= Node2.getAttribute("Qty")
				nTempRate		= Node2.getAttribute("Rate")
				dItemNetBasic	= Node2.getAttribute("NettBasic")
				dItemValue		= Node2.getAttribute("Amount")

				If sItemPurType = Node2.getAttribute("PurchaseType") then
					if cdbl(dItemValue) > 0 then
						''Item Rate
						dItemRate = cdbl(dItemValue) / cdbl(iQty)

						'' roundoff to 4 decimals added on 13 Sep 04 to avoid Invoice value - stock value mismatch
						dItemRate = round(dItemRate,4)
						dItemValue = FormatNumber(dItemValue,2,,,0)

						Node2.setAttribute "ItemValue", dItemValue
						Node2.setAttribute "ItemRate", dItemRate


						If iQty > 0 and nTempRate > 0 then
							If dItemRate > 0 then
								'eval("document.formname.txtItemRateDisplay"&Node2.getAttribute("ClassificationCode")&"Z"&Node2.getAttribute("ItemCode")).value=FormatNumber(dItemRate,4,,,0)
								'nSumOfItemRate = cdbl(nSumOfItemRate) + cdbl(dItemRate)
								eval("document.formname.txtItemRateDisplay"&Node2.getAttribute("ClassificationCode")&"Z"&Node2.getAttribute("ItemCode")&"Z"&Node2.getAttribute("EntryNo")).value=dItemValue
								nSumOfItemRate = cdbl(nSumOfItemRate) + cdbl(dItemValue)
							End If 'If dItemRate > 0 then
						End If 'If iQty > 0 and nTempRate > 0 then


					end if 'if cdbl(dItemValue) > 0 then
				end if 'If sItemPurType = Node2.getAttribute("PurchaseType") then


				if dItemValue > 0 and iQty > 0 and dProdRate > 0 then
					nTotal = cdbl(nTotal) + cdbl(Node2.getAttribute("Amount"))
				end if

			Next 'For each Node2 in Node1.ChildNodes

		end if 'if Node1.nodeName = "ItemDetails" then

	next 'for each Node1 in ObjRoot.ChildNodes

	if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
		eval("document.formname.txtTotalItemRate" + sPurchaseType + "Z" + EntNo).value	= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
	else

		for each NodeTax in ObjRoot.ChildNodes
			if NodeTax.nodeName = "TaxDetails" then
				sPType= NodeTax.getAttribute("PurchaseType")
				nTotal = cdbl(nTotal) + cdbl( eval("document.formname.txtTaxDisplay" & sPType).value )
			end if
		next

		eval("document.formname.txtTotalItemRate" + sItemPurType + "Z" + EntNo).value		= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
		document.formname.mTxtTotal.value = FormatNumber(RndOff(nTotal),2,,,0)
	end if

End Function


'---------------------------------------------------------------------------------
Function CalculateItemValue(sPassPurType)
	Dim ObjRoot,Node1
	Dim nSumOfItemRate,nTotal
	Dim sPurchaseType,sItemPurType

	Set ObjRoot = InvoiceDet.documentElement
	'alert(ObjRoot.xml)

	nTotal = 0.00
	for each Node1 in ObjRoot.ChildNodes

		if Node1.nodeName = "InvoiceHeader" then

			FOR EACH Node2 in Node1.ChildNodes
				if Node2.NodeName = "Header" then
					sPurchaseType	= Node2.getAttribute("PurchaseType")
					dTotInvVal		= Node2.getAttribute("InvValue")
				end if 'if Node2.NodeName = "Header" then
			next
		end if 'if Node1.nodeName = "InvoiceHeader" then


		if Node1.nodeName = "ItemDetails" then


			dTemp=0
			iItemCount = 0
			nPos = 0

			if trim(sPassPurType) <> "" then
				For each Node2 in Node1.ChildNodes

					If trim(Node2.getAttribute("PurchaseType")) = trim(sPassPurType) then
						Node2.setAttribute "Amount", Node2.getAttribute("NettBasic")
						iItemCount = iItemCount + 1
						nPos = Node2.getAttribute("EntryNo")
					End if 'If trim(Node2.getAttribute("PurchaseType")) = trim(sPassPurType) then

				Next 'For each Node2 in Node1.ChildNodes

			end if 'if trim(sPassPurType) <> "" then

			For each Node2 in Node1.ChildNodes

				sItemPurType=Node2.getAttribute("PurchaseType")

				if trim(sPassPurType) = trim(sItemPurType) then

					dItemValue=Node2.getAttribute("Amount")

					set TaxNodes=ObjRoot.SelectNodes("//TaxDetails[@PurchaseType="& sPassPurType &"]")

					if TaxNodes.Length > 0 then

						set oNodTax=TaxNodes.Item(0)

						dTotBasicVal	= oNodTax.getAttribute("Basicvalue")
						dTotNetVal		= oNodTax.getAttribute("NettValue")
						dInvAmount		= oNodTax.getAttribute("SubTotal")

						if cdbl(dTotBasicVal) > 0 and cdbl(dTotNetVal) > 0 then

							For Each oNodEntry in oNodTax.childNodes

								sCatCode	=	oNodEntry.getAttribute("CatCode")
								sTaxCode	=	oNodEntry.getAttribute("TaxCode")
								sTaxMode	=	oNodEntry.getAttribute("TaxMode")
								sFormula	=	oNodEntry.getAttribute("TaxFormula")
								dTaxValue	=	oNodEntry.getAttribute("TaxValue")

								if sTaxMode="F" then
									dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * dItemValue
								elseif sTaxMode="P" then
									dTax=ItemValueCalculateTax(oNodTax,sFormula,dItemValue,dItemValue,dTaxValue,sItemPurType)
								else
									dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * cdbl(dItemValue)
								end if

								saTemp= Split(CStr(dTax),".")

								if UBound(saTemp)>0 then
									oNodEntry.setAttribute "ItemValue", saTemp(0)&"."&left(saTemp(1),2)
								else
									oNodEntry.setAttribute "ItemValue", saTemp(0)
								end if
							Next

							For Each oNodEntry in oNodTax.childNodes
								if cint(oNodEntry.getAttribute("AccHead"))=0 then	'' considering Acc head = 0 only
									if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
										dItemValue=CDbl(dItemValue)+CDbl(oNodEntry.getAttribute("ItemValue"))
									end if
								else
									if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
										dTemp = CDbl(dTemp) + CDbl(oNodEntry.getAttribute("TaxAmount"))
									end if
								end if
							next
						end if 'if cdbl(dTotBasicVal) > 0 and cdbl(dTotNetVal) > 0 then
						dItemTotal=CDbl(dItemValue)+CDbl(dItemTotal)
						Node2.setAttribute "Amount", dItemValue

						'msgbox "dItemValue = " & trim(dItemValue)

					end if 'if TaxNodes.Length > 0 then
				end if 'if trim(sPassPurType) = trim(sItemPurType) then

			Next'FOR EACH Node2 in Node1.ChildNodes

			if nPos > 0 then
				nPos = nPos - 1
				'For Last Item

				if  iItemCount > 0 then
					dItemTotal = CDbl(dTemp/iItemCount) + CDbl(dItemTotal)

					'alert("after calc = " & trim(dItemTotal) )

					dTemp = Node1.childNodes(nPos).getAttribute("Amount")
					dTemp = CDbl(dTemp) + CDbl(dInvAmount)-CDbl(dItemTotal)
					Node1.childNodes(nPos).SetAttribute "Amount",dTemp

					'alert(" final 123 = " & trim(dTemp) )
				end if 'if  iItemCount > 0 then
				'block 1
			end if 'if nPos > 0 then

			For each Node2 in Node1.ChildNodes

				dProdRate		= Node2.getAttribute("Rate")
				iQty			= Node2.getAttribute("Qty")
				dItemNetBasic	= Node2.getAttribute("NettBasic")
				dItemValue		= Node2.getAttribute("Amount")
				sItemPurType	= Node2.getAttribute("PurchaseType")


				if trim(sPassPurType) = trim(sItemPurType) then
					''Item Rate
					dItemRate = cdbl(dItemValue) / cdbl(iQty)
					'alert(dItemRate)

					'' roundoff to 4 decimals added on 13 Sep 04 to avoid Invoice value - stock value mismatch
					dItemRate = round(dItemRate,4)
					dItemValue = FormatNumber(dItemValue,2,,,0)

					Node2.setAttribute "ItemValue", dItemValue
					Node2.setAttribute "ItemRate", dItemRate

					iItemCode = Node2.getAttribute("ItemCode")
					iClassCode = Node2.getAttribute("ClassificationCode")
					EntNo = Node2.getAttribute("EntryNo")
					If iQty > 0 and dProdRate > 0 then
						If dItemRate > 0 then

							eval("document.formname.txtItemRateDisplay"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value=dItemValue
							nSumOfItemRate = cdbl(nSumOfItemRate) + cdbl(dItemValue)

						End If 'If dItemRate > 0 then
					End If 'If iQty > 0 and dProdRate > 0 then

					if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
						eval("document.formname.txtTotalItemRate" + sPurchaseType + "Z" + EntNo).value	= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					else
						eval("document.formname.txtTotalItemRate" + sItemPurType + "Z" + EntNo).value		= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					end if
				end if 'if trim(sPassPurType) = trim(sItemPurType) then

				'alert(dItemValue)
				if dItemValue > 0 and iQty > 0 and dProdRate > 0 then
					nTotal = cdbl(nTotal) + cdbl(dItemValue)
				end if

				'msgbox "dProdRate = "& trim(dProdRate)
			Next

		end if 'if Node1.nodeName = "ItemDetails" then

	next 'for each Node1 in ObjRoot.ChildNodes

	'alert("before = " & trim(nTotal) )

	if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then

	else
		for each NodeTax in ObjRoot.ChildNodes
			if NodeTax.nodeName = "TaxDetails" then
				sPType= NodeTax.getAttribute("PurchaseType")
				'alert("data = " & trim( eval("document.formname.txtTaxDisplay" & sPType).value )   )
				nTotal = cdbl(nTotal) + cdbl( eval("document.formname.txtTaxDisplay" & sPType).value )
				'alert("for = " + trim(nTotal) )
			end if
		next

		nTotal = cdbl(nTotal) + cdbl(document.formname.txtRoundOff.value)
		document.formname.mTxtTotal.value = FormatNumber(nTotal,2,,,0)
	end if


End Function
'--------------------------------------------------------------------------------------------
Function CalcAmountForLastItem()

	Dim ObjRoot,Node1,sRoundOffAccHead
	Dim nSumOfItemRate,nTotal,nRoundOffValue
	Dim sPurchaseType,sItemPurType,sPrevItemPurchaseType


	sRoundOffAccHead = trim(document.formname.hRoundOffAccHead.value)

	Set ObjRoot = InvoiceDet.documentElement
	'alert(ObjRoot.xml)

	nTotal = 0.00
	for each Node1 in ObjRoot.ChildNodes

		if Node1.nodeName = "InvoiceHeader" then

			FOR EACH Node2 in Node1.ChildNodes
				if Node2.NodeName = "Header" then
					dTotInvVal		= Node2.getAttribute("InvValue")
					sPurchaseType	= Node2.getAttribute("PurchaseType")
					nRoundOffValue	= Node2.getAttribute("RoundOff")
				end if 'if Node2.NodeName = "Header" then
			next
		end if 'if Node1.nodeName = "InvoiceHeader" then


		if Node1.nodeName = "ItemDetails" then

			iItemCount = 0

			if Node1.hasChildNodes then
				iItemCount = Node1.ChildNodes.length
				nLastItemCode  = Node1.childNodes(iItemCount - 1).getAttribute("ItemCode")
				nLastClassCode = Node1.childNodes(iItemCount - 1).getAttribute("ClassificationCode")
				nLastEntNo     = Node1.childNodes(iItemCount - 1).getAttribute("EntryNo")
			end if

			if sRoundOffAccHead	 = "0" then
				For each Node2 in Node1.ChildNodes


					if trim(Node2.getAttribute("EntryNo")) = trim(nLastEntNo) and trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) then

						'Node2.setAttribute "Amount", Node2.getAttribute("NettBasic")
						dItemValue=Node2.getAttribute("Amount")

						if document.formname.RdRndOff(0).Checked=true then
						'	if CDbl(nRoundOffValue) < 0 then
						'		dItemValue = CDbl(dItemValue) -( -1 * CDbl(nRoundOffValue) )
						'	else
						'		dItemValue = CDbl(dItemValue) + CDbl(nRoundOffValue)
						'	end if
							dItemValue = CDbl(dItemValue) + CDbl(nRoundOffValue)
						else
						'	if CDbl(nRoundOffValue) < 0 then
						'		dItemValue = CDbl(dItemValue) + ( -1 * CDbl(nRoundOffValue) )
						'	else
						'		dItemValue = CDbl(dItemValue) - CDbl(nRoundOffValue)
						'	end if
							dItemValue = CDbl(dItemValue) - CDbl(nRoundOffValue)
						end if 'if document.formname.RdRndOff(0).Checked=true then

						Node2.setAttribute "Amount", dItemValue
					end if 'if trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) then

				Next'FOR EACH Node2 in Node1.ChildNodes
			end if 'if sRoundOffAccHead	 = "0" then

			sPrevItemPurchaseType = ""
			For each Node2 in Node1.ChildNodes


				dItemValue		= Node2.getAttribute("Amount")
				if eval("document.formname.chkVat"&Node2.getAttribute("ClassificationCode")&"Z"&Node2.getAttribute("ItemCode")&"Z"&Node2.getAttribute("EntryNo")).checked then
					Node2.setAttribute "VAT", "Y"
				else
					Node2.setAttribute "VAT", "N"
				end if

				if trim(Node2.getAttribute("EntryNo")) = trim(nLastEntNo) and trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) then
					iQty			= Node2.getAttribute("Qty")
					dItemNetBasic	= Node2.getAttribute("NettBasic")

					sItemPurType	= Node2.getAttribute("PurchaseType")

					if trim(sPrevItemPurchaseType) = "" or trim(sPrevItemPurchaseType) <> trim(sItemPurType) then
						nSumOfItemRate = 0.00
						sPrevItemPurchaseType = sItemPurType
					end if

					''Item Rate
					dItemRate = cdbl(dItemValue) / cdbl(iQty)
					'alert(dItemRate)

					'' roundoff to 4 decimals added on 13 Sep 04 to avoid Invoice value - stock value mismatch
					dItemRate = round(dItemRate,4)
					dItemValue = FormatNumber(dItemValue,2,,,0)

					Node2.setAttribute "ItemValue", dItemValue
					Node2.setAttribute "ItemRate", dItemRate

					iItemCode = Node2.getAttribute("ItemCode")
					iClassCode = Node2.getAttribute("ClassificationCode")
					EntNo = Node2.getAttribute("EntryNo")
					If iQty > 0 and dItemRate > 0 then
						If dItemRate > 0 then

							eval("document.formname.txtItemRateDisplay"&iClassCode&"Z"&iItemCode&"Z"&EntNo).value=dItemValue
							nSumOfItemRate = cdbl(nSumOfItemRate) + cdbl(dItemValue)

						End If 'If dItemRate > 0 then
					End If 'If iQty > 0 and dItemRate > 0 then

					if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
						eval("document.formname.txtTotalItemRate" + sPurchaseType + "Z" + EntNo).value	= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					else
						eval("document.formname.txtTotalItemRate" + sItemPurType + "Z" + EntNo).value		= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					end if
				end if 'if trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) then

			'	alert(dItemValue)

				nTotal = cdbl(nTotal) + cdbl(dItemValue)

			Next

		end if 'if Node1.nodeName = "ItemDetails" then

	next 'for each Node1 in ObjRoot.ChildNodes

	if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
	else
		for each NodeTax in ObjRoot.ChildNodes
			if NodeTax.nodeName = "TaxDetails" then
				sPType= NodeTax.getAttribute("PurchaseType")
				nTotal = cdbl(nTotal) + cdbl( eval("document.formname.txtTaxDisplay" & sPType).value )
			end if
		next

		document.formname.mTxtTotal.value	= FormatNumber(nTotal,2,,,0)
	end if

End Function
'--------------------------------------------------------------------------------------------
Function ItemValueCalculateTax(oNodTaxRoot,sFormula,dBValue,dDValue,dPercentage,sItemPurType)
dim saTemp,dTaxAmount,iCounter,iTemp,TaxRoot
dim oNodTemp
dim saTemp1


if trim(sItemPurType) ="0" then
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
else
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			if trim(oNodTemp.getAttribute("PurchaseType")) = trim(sItemPurType) then
				set TaxRoot=oNodTemp
			end if
		end if
	next
end if 	'if trim(sItemPurType) ="0" then

saTemp=Split(sFormula,",")
if trim(saTemp(0))="BV" then
	dTaxAmount=dBValue
	iTemp=1
elseif trim(saTemp(0))="BD" then
	dTaxAmount=dDValue
	iTemp=1
else
	dTaxAmount=0
	iTemp=0
end if
'alert(sFormula)
for iCounter=iTemp to UBound(saTemp)
	saTemp1=Split(trim(saTemp(iCounter)),"#")
	For Each oNodTemp in oNodTaxRoot.childNodes
		if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
			'dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(7).nodeValue)

			'dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(5).nodeValue)
			'alert(oNodTemp.Attributes.Item(5).nodeValue)

			dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.getAttribute("ItemValue") )
		end if
	next
next

if trim(dPercentage)<>"" then
	ItemValueCalculateTax=dTaxAmount*(cdbl(dPercentage)/100)
else
	ItemValueCalculateTax=dTaxAmount
end if

End function
'-------------------------------------------------------------------------------------------
Function AssignRoundOffValue()

if document.formname.RdRndOff(0).Checked=true then
	if CDbl(document.formname.txtRoundOff.value) <> 0 then
		set RootNode=InvoiceDet.documentElement
		set TempNode=RootNode.SelectNodes("//InvoiceHeader/Header")
		if TempNode.Length > 0 then
			if CDbl(document.formname.txtRoundOff.value) <> cdbl(TempNode.item(0).Attributes.GetNamedItem("RoundOff").value) then

				InvVal=cdbl(TempNode.item(0).Attributes.GetNamedItem("InvValue").value)
				'substract existing round off value and add new round off value
				InvVal = CDbl(InvVal) - cdbl(TempNode.item(0).Attributes.GetNamedItem("RoundOff").value) + CDbl(document.formname.txtRoundOff.value)
				TempNode.item(0).SetAttribute "InvValue",InvVal
				TempNode.item(0).SetAttribute "RoundOff",document.formname.txtRoundOff.value
				RoundOffInv()
			end if
		end if 'if TempNode.Length > 0 then
	end if 	'if CDbl(document.formname.txtRoundOff.value) <> 0 then
end if 'if document.formname.RdRndOff(0).Checked=true then
end Function
'-------------------------------------------------------------------------------------------
Function RoundOffInv()
dim TempNode,InvVal,RndOffVal,blnRoundOffNodeExist,RndOffAmt

set RootNode=InvoiceDet.documentElement
set TempNode=RootNode.SelectNodes("//InvoiceHeader/Header")
if TempNode.Length > 0 then
	sPurchaseType	= TempNode.item(0).Attributes.GetNamedItem("PurchaseType").value
	InvVal=cdbl(TempNode.item(0).Attributes.GetNamedItem("InvValue").value)
	RndOffVal=cdbl(TempNode.item(0).Attributes.GetNamedItem("RoundOff").value)
	RndOffAmt = RndOffVal
end if

if document.formname.RdRndOff(0).Checked=true then
	'InvVal = CDbl(InvVal) +  CDbl(RndOffVal)
else
	InvVal = CDbl(InvVal) -  CDbl(RndOffVal)
	RndOffVal = 0.00
end if
document.formname.txtRoundOff.value = FormatNumber(RndOffVal,2,,,0)
document.formname.txtInvValue.value = FormatNumber(InvVal,2,,,0)

if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
else
	document.formname.mTxtTotal.value = FormatNumber(InvVal,2,,,0)
end if

end function
'---------------------------------------------------------------------------------
Sub AddRoundOffNode(nPassRndOffAmt,nInvVal)

Dim Node1,TempNode,NewElem

Dim blnRoundOffNodeExist


	set RootNode=InvoiceDet.documentElement

	For each Node1 in RootNode.ChildNodes

		If trim(Node1.NodeName) = "InvoiceHeader" then
			for each TempNode in Node1.ChildNodes
				if trim(TempNode.NodeName) = "Header" then
					TempNode.setAttribute "InvValue",nInvVal
				end if
			next
		end if

		If trim(Node1.NodeName) = "TaxDetails" then

			'find round of tax node is exist or not , if not exist add it
			blnRoundOffNodeExist = false
			for each TempNode in Node1.ChildNodes
				set NewElem = TempNode
				if TempNode.getAttribute("CatCode") = "0" and TempNode.getAttribute("TaxCode") = "0" then
					blnRoundOffNodeExist = true
				end if
			Next

			if not blnRoundOffNodeExist then
				set NewElem = InvoiceDet.CreateElement("Tax")
				Node1.AppendChild NewElem
			end if 	 'end if 'if not blnRoundOffNodeExist then

			NewElem.setAttribute "CatCode","0"
			NewElem.setAttribute "TaxCode","0"
			NewElem.setAttribute "TaxMode","0"
			NewElem.setAttribute "TaxFormula","0"
			NewElem.setAttribute "TaxValue",nPassRndOffAmt
			NewElem.setAttribute "TaxAmount",nPassRndOffAmt
			NewElem.setAttribute "AccHead",document.formname.hRoundOffAccHead.value
			NewElem.setAttribute "Formnumber","0"
			NewElem.setAttribute "TransAmt","0"
			NewElem.setAttribute "ItemValue","0"
			NewElem.text = "ROUND OFF"

			'alert("nPassRndOffAmt = " + trim(nPassRndOffAmt) )
		End If 'If trim(Node1.NodeName) = "TaxDetails" then

	Next 'For each Node1 in RootNode.ChildNodes

end sub
'---------------------------------------------------------------------------------
Function ShowPrefDet()
	dim TempNode
	dim Curr,Mod1,Mop,IssueBank,PayTerm,Bop,Transporter
	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set TempNode=oNodTemp
			exit for
		end if
	next
	for Each oNodTemp in TempNode.childNodes
		Curr=oNodTemp.Attributes.getnameditem("Currency").value
		Mod1=oNodTemp.Attributes.getnameditem("DespatchMode").value
		Mop=oNodTemp.Attributes.getnameditem("PaymentMode").value
		nIssueBank=oNodTemp.Attributes.getnameditem("IssueBank").value
		nBenefitBank=oNodTemp.Attributes.getnameditem("BenificiaryBank").value
		PayTerm=oNodTemp.Attributes.getnameditem("PayTerms").value
		Bop=oNodTemp.Attributes.getnameditem("PricingBasis").value
		Transporter=oNodTemp.Attributes.getnameditem("Transporter").value
		nLoadPort=oNodTemp.Attributes.getnameditem("LoadingPort").value
		nDestPort=oNodTemp.Attributes.getnameditem("DestPort").value
	next
	set TempNode = showModalDialog("InvPurInvoiceEntryPref.asp?Mod1="+ Mod1 +"&Mop="+ Mop +"&IssueBank="+trim(nIssueBank)+"&PayTerm="+PayTerm +"&Bop="+Bop +"&Transporter="+Transporter+"&BenefitBank="+trim(nBenefitBank) + "&LoadPort=" + trim(nLoadPort) + "&DestPort=" + trim(nDestPort) ,TempNode,"dialogLeft:0px;dialogTop:0Px;dialogHeight:250px;dialogWidth:600px;status:no")
end function

'---------------------------------------------------------------------------------

Function DeleteItems()
	dim root,sExp,ItemNode,itr,iItemDel,iClassDel,objSel,objSer
	set root = InvoiceDet.DocumentElement
	sExp ="//ItemDetails/Item"
	Set ItemNode = Root.Selectnodes(sExp)
	if ItemNode.Length > 0 then
		for itr = 0 to ItemNode.Length - 1
			iItemDel = ItemNode.item(itr).Attributes.getNamedItem("ItemCode").value
			iClassDel = ItemNode.item(itr).Attributes.getNamedItem("ClassificationCode").value
			iEntDel  = ItemNode.item(itr).Attributes.getNamedItem("EntryNo").value
			set objSel = eval("document.formname.chkDeleteA"&CStr(iClassDel)&"A"&CStr(iItemDel)&"A"&CStr(iEntDel))
			set objSer = eval("document.formname.txtSerA"&CStr(iClassDel)&"A"&CStr(iItemDel)&"A"&CStr(iEntDel))
			if objSel.checked then DeleteItem objSel,objSer.value + 2
		next
	end if
End Function

'-------------------------------------------------------------------------------

Function DeleteItem(sobj,iRow)
	dim itr,iItem,iClass,arrTemp,root,sExp,DeleteNode,oNode,iser,iItemDel,iClassDel,objSer,PurType,i,TaxNodes,Node,ItemRoot
	itr = 0

	'alert(window.event.srcElement.sourceIndex)
	arrTemp = split(sobj.name,"A")
	iItem = arrTemp(2)
	iClass = arrTemp(1)
	iEntNo = arrTemp(3)
	set root = InvoiceDet.DocumentElement
	if root.hasChildNodes then
		For Each oNodTemp in Root.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
		next

		sExp = "//Item[@ItemCode = "&iItem&" and @ClassificationCode = "&iClass&" and @EntryNo = "& iEntNo &"]"
		Set DeleteNode = ItemRoot.Selectnodes(sExp)
		if DeleteNode.Length > 0 then

			PurType=DeleteNode.Item(0).Attributes.getNamedItem("PurchaseType").value
			Set oNode = ItemRoot.RemoveChild(DeleteNode.Item(0))
			if document.formname.cmbPurType.value=0 and PurType<>"0" then
				sExp = "//TaxDetails"
				Set TaxNodes = Root.Selectnodes(sExp)
				for i=0 to TaxNodes.Length-1
					if TaxNodes.Item(i).Attributes.getNamedItem("PurchaseType").value=Purtype then
						exit for
						msgbox "i="&i
					end if
				next
				iRow=iRow + 3*i
				document.all.tblItemDet.deleteRow(iRow - 1)
			else
				document.all.tblItemDet.deleteRow(iRow - 1)
			end if
		end if

		sExp = "//ItemDetails/Item[@PurchaseType = "& PurType &"]"
		set ItemNode= Root.Selectnodes(sExp)
		'msgbox "Purtype="&PurType
		'msgbox "Len="&ItemNode.Length
		if ItemNode.Length=0 then
			sExp = "//TaxDetails[@PurchaseType = "& PurType &"]"
			Set DeleteNode = Root.Selectnodes(sExp)
			if DeleteNode.Length > 0 then
				if document.formname.cmbPurType.value=0 and PurType<>"0" then
					document.all.tblItemDet.deleteRow(iRow - 1)
					document.all.tblItemDet.deleteRow(iRow - 1)
					document.all.tblItemDet.deleteRow(iRow - 1)
				else
					for i=2 to (document.all.tblItemDet.rows.length - 1)
						document.all.tblItemDet.deleteRow(2)
					next
				end if
				Set oNode = root.RemoveChild(DeleteNode.Item(0))
			end if

		end if

		iser = 0
		set TaxNodes=Root.SelectNodes("//TaxDetails")
		for i= 0 to TaxNodes.Length-1
			PurType=TaxNodes.item(i).Attributes.GetNamedItem("PurchaseType").value
			sExp ="//ItemDetails/Item[@PurchaseType = "& PurType &"]"
			Set ItemNode = Root.Selectnodes(sExp)

			if ItemNode.Length > 0 then
				for itr = 0 to ItemNode.Length - 1
					iser = iser + 1
					iItemDel = ItemNode.item(itr).Attributes.getNamedItem("ItemCode").value
					iClassDel = ItemNode.item(itr).Attributes.getNamedItem("ClassificationCode").value
					iEntDel = ItemNode.item(itr).Attributes.getNamedItem("EntryNo").value
					set objSer = eval("document.formname.txtSerA"&CStr(iClassDel)&"A"&CStr(iItemDel)&"A"&CStr(iEntDel))
					objSer.value = iser
				next
			end if
		next
		if TaxNodes.Length=0 then
			for i=2 to (document.all.tblItemDet.rows.length - 1)
				document.all.tblItemDet.deleteRow(2)
			next
		else
			if document.formname.cmbPurType.value=0 and PurType<>"0" then
				setValues1
			else
				SetValues
			end if
		end if
	end if
	'alert(OutData.xml)
End Function

'=========================================================================
FUNCTION setValues1()
	dim dNetTotal,dDisTotal,dAmount,dRate,dQty,dDisAmount,i,PurType,j,ItemNd
	set RootNode=Invoicedet.documentElement
	Set TaxNode=RootNode.SelectNodes("//TaxDetails")
	for i=0 to TaxNode.length-1
		PurType=TaxNode.item(i).Attributes.getnamedItem("PurchaseType").value
		set ItemNd=RootNode.SelectNodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")

		dNetTotal = 0
		dDisTotal = 0
		dTBasicVal = 0

		for j=0 to ItemNd.length-1
			set oNodTemp=ItemNd.item(j)
			dDisTotal=CDbl(dDisTotal)+CDbl(oNodTemp.Attributes.Item(7).nodeValue)
			dQty=oNodTemp.Attributes.Item(4).nodeValue

			'' For Alt. UoM
			'dRate=oNodTemp.Attributes.Item(5).nodeValue
			'dRate= oNodTemp.Attributes.getNamedItem("RatePerQtyUoM").value
			dRate= oNodTemp.Attributes.getNamedItem("Rate").value

			dTBasicVal = dTBasicVal + (CDbl(dQty)*CDbl(dRate))
			dAmount=(CDbl(dQty)*CDbl(dRate))-CDbl(oNodTemp.Attributes.Item(7).nodeValue)
			dNetTotal=CDbl(dNetTotal)+dAmount
		next
		eval("document.formname.txtDisValue" &PurType).value=FormatNumber(Round(dDisTotal,2),2,,,0)
		eval("document.formname.txtAmount" &PurType).value=FormatNumber(Round(dNetTotal,2),2,,,0)
		eval("document.formname.txtBasicValue" &PurType).value=FormatNumber(Round(dTBasicVal,2),2,,,0)
	next
	popTax1
END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION setValues()
dim dNetTotal,dDisTotal,dAmount,dRate,dQty,dDisAmount
	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
	next

	dNetTotal = 0
	dDisTotal = 0
	dTBasicVal = 0

			For Each oNodTemp in ItemRoot.childNodes
				dDisTotal=CDbl(dDisTotal)+CDbl(oNodTemp.Attributes.Item(7).nodeValue)
				dQty=oNodTemp.Attributes.Item(4).nodeValue

				'' For Alt. UoM
				'dRate=oNodTemp.Attributes.Item(5).nodeValue
				'dRate = oNodTemp.Attributes.getNamedItem("RatePerQtyUoM").value
				dRate= oNodTemp.Attributes.getNamedItem("Rate").value

				dTBasicVal = dTBasicVal + (CDbl(dQty)*CDbl(dRate))
				dAmount=(CDbl(dQty)*CDbl(dRate))-CDbl(oNodTemp.Attributes.Item(7).nodeValue)
				dNetTotal=CDbl(dNetTotal)+dAmount
			next
			document.formname.txtDisValue.value=FormatNumber(Round(dDisTotal,2),2,,,0)
			document.formname.txtAmount.value=FormatNumber(Round(dNetTotal,2),2,,,0)
			document.formname.txtBasicValue.value=FormatNumber(Round(dTBasicVal,2),2,,,0)
'			msgbox dTBasicVal
			popTax
END FUNCTION
'-------------------------------------------
Function CalcTax()

Dim nChk,dInvAmount,BasTot,NetTot,sTaxCode,sCatCode,sFormula,dTotal,dTaxvalue
dim sRoot,node,node1,TaxFound,sTaxMode,dBasicTotal,sTotpackvalue,iMCtr,iPurType

set sRoot=InvoiceDet.DocumentElement

set objNetTotal = eval("document.formname.txtAmount")
set objBasTotal = eval("document.formname.txtBasicValue")
set objInvTotal = eval("document.formname.txtAmount")


if objBasTotal.value="" then
	bastot=0
else
	bastot=objBasTotal.value
end if
if objNetTotal.value="" then
	NetTot=0
else
	NetTot=objNetTotal.value
end if


TaxFound=0
iPurType = 0
if sRoot.hasChildNodes() then
	for each node in sRoot.childNodes
		if strcomp(node.nodeName,"InvoiceHeader")=0 then
			for each ndChild in node.childNodes
				if strcomp(ndChild.nodeName,"Header")=0 then
					iPurType =  ndChild.getAttribute("PurchaseType")
					exit for
				end if
			next
		end if
	next
end if
if sroot.haschildNodes() then
	for each node in sRoot.childnodes
		if node.nodename="TaxDetails" then
			TaxFound=1
			exit for
		end if
	next
end if

if TaxFound=0 then
	dim NewElem
	set NewElem= InvoiceDet.CreateElement("TaxDetails")
	sRoot.AppendChild NewElem
	NewElem.setAttribute "BasicValue" , objBasTotal.value
	NewElem.setAttribute "NetValue" , objNetTotal.value
	NewElem.setAttribute "InvValue" , objNetTotal.value
	NewElem.setAttribute "RoundOff" , 0
else
	set sRoot=InvoiceDet.DocumentElement
	for each newElem in sroot.childnodes
		if newelem.nodename="TaxDetails" then
			node.Attributes.Item(0).nodeValue = objbastotal.Value
			node.Attributes.Item(1).nodeValue = objNetTotal.Value

			exit for
		end if
	next
end if

dBasicTotal=ObjBasTotal.value
if sroot.haschildNodes() then
	for each node in sRoot.childnodes
		dInvAmount=0
		if node.nodename="TaxDetails" then

			dinvAmount=node.Attributes.Item(1).nodeValue
			dTotal=node.Attributes.Item(1).nodeValue

			if dInvAmount=0 then exit function

			if node.haschildnodes then
				for each node1 in node.childnodes

					iCtr=node1.Attributes.Item(0).nodeValue & node1.Attributes.Item(1).nodeValue
					dim pos,TaxPer
					sTaxMode =node1.Attributes.Item(2).nodeValue
					sCatCode =node1.Attributes.Item(0).nodeValue
					sTaxCode =node1.Attributes.Item(1).nodeValue

					set ObjTaxPer		= eval("document.formname.txtTaxPer"&cstr(sCatCode)&CStr(sTaxCode))
					pos=instr(1,(objTaxPer.value),"%")
					if pos>0 then
						TaxPer=mid(objtaxPer.value,1,pos-1)
					else
						taxper=cdbl(objTaxPer.value)
					end if

					dTaxValue=cdbl(TaxPer)
					sFormula=node1.Attributes.Item(3).nodeValue
					dTRndoff = node1.Attributes.Item(7).nodeValue


					nChk = 0
					if isnull(ObjTaxPer.value)  then
							alert("Enter Tax Percentage")
							ObjTaxPer.focus()
							nChk = 1
							exit function
						else
							if not IsNumeric(TaxPer) then
								alert("Enter Numeric value")
								ObjTaxPer.focus()
								objtexPer.Select()
								nChk = 1
								exit function
							end if 'if not IsNumeric(ObjQty.value) then
					end if 'if ObjQty.value = "" then
					if IsNumeric(TaxPer) then
						set ObjTaxValue = eval("document.formname.txtTaxValue"+cstr(sCatCode)&"Z"&cstr(sTaxCode))
						set ObjSubTotal = eval("document.formname.txtSubTaxValue"+cstr(sCatCode)&cstr(sTaxCode))

							if sTaxMode ="P" then
								dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue,iPurType)
								eval("document.formname.txtTaxPer"&cstr(sCatCode)&cstr(sTaxCode)).value=dTaxValue & "%"
							elseif sTaxMode="Q" then
								dTax = cdbl(dTaxvalue) * cdbl(Qtysum)
								eval("document.formname.txtTaxPer"&cstr(sCatCode)&cstr(sTaxCode)).value=dTaxValue
							elseif sTaxMode = "K" then
								sTotpackvalue = 0
								sExp = "//Tax[@CatCode = "&sCatCode&" and @TaxCode ="&sTaxCode&"]/Taxpack"
								set sNode = sRoot.Selectnodes(sExp)

								If sNode.Length > 0 then
									For IMCtr = 0 to sNode.Length - 1
										sTotpackvalue = cdbl(sTotpackvalue) + cdbl(sNode.Item(iMCtr).Attributes.Item(3).nodevalue)
									Next
								end if
								dTax=sTotpackvalue
								eval("document.formname.txtTaxPer"&CStr(sCatCode)&cstr(sTaxCode)).value=dTaxValue
							else
								dTax=dTaxvalue
								eval("document.formname.txtTaxPer"&CStr(sCatCode)&cstr(sTaxCode)).value=dTaxValue
							end if

							'MsgBox dTax

							if trim(dTRndoff) = "1" then
								dtax = Rndoff(dTax)
							End if

							dInvAmount=dInvAmount+CDbl(dTax)
							objTaxValue.value=FormatNumber(dtax,2,-1,-1,0)
							objSubTotal.value=FormatNumber(dInvAmount,2,-1,-1,0)
							node1.Attributes.Item(5).nodeValue = trim(objtaxValue.value)
							set ObjTaxPer		= eval("document.formname.txtTaxPer"&CStr(sCatCode)&cstr(sTaxCode))
							pos=instr(1,(objTaxPer.value),"%")
							if pos>0 then
								TaxPer=mid(objtaxPer.value,1,pos-1)
							else
								taxper=cdbl(objTaxPer.value)
							end if
							node1.Attributes.Item(4).nodeValue = taxPer
							objtaxper.value=taxper
						'end if
					end if
					node.attributes.item(2).nodevalue=dInvAmount
					CalcTotal()
				next
			else
				CalcTotal()
			end if 'if node.haschildnodes then

			exit for

		end if 'if node.nodename="TaxDetails" then
	next
end if 'if sroot.haschildNodes() then
end Function
'----------------------------------------------------------------------
Function CalcTotal()
	Dim dInvAmount
	dim sRoot,node
	set sRoot=InvoiceDet.DocumentElement
	for each node in sRoot.childnodes
		if node.nodename="TaxDetails" then
			dinvAmount=node.Attributes.Item(2).nodeValue
			set objInvValue = eval("document.formname.txtInvValue")
			objInvValue.value=FormatNumber(RndOff(dInvAmount),2,,,0)

			set objRoundOff = eval("document.formname.txtRoundOff")
			objRoundOff.value = FormatNumber((cdbl(objInvValue.value)-dInvAmount) ,2,-1,0,0)
			node.Attributes.Item(2).nodeValue = trim(objinvValue.value)
			node.Attributes.Item(3).nodeValue = trim(objRoundOff.value)
			exit for
		end if
	next
end Function
'--------------------------------------------------------------------------
Function Packvaluechange(sOrgid,sPurType,sTcode,sTpcode)
	Dim Returnvalue,sTemp

		sTemp = sPurType&":"&sOrgid&":"&sTcode&":"&sTpcode

		Returnvalue = showModalDialog ("PurInvTaxPackage.asp?Invcode="&sTemp,InvoiceDet,"dialogHeight:370px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")

		If Returnvalue = "S" then
			CalcTax
		End if
End Function
'-------------------------------------------------------------------------------------------------
'=========================================================================
function ViewItemValues()
	if Next_Click("V") then
		'showModalDialog "invPurItemValuePop.asp","","","dialogHeight:435px;dialogWidth:700px;Status:no;help:no"
		Set Root=InvoiceDet.documentElement
		'alert(Root.xml)
		showModalDialog "InvPurItemValueViewPop.asp",InvoiceDet,"dialogHeight:435px;dialogWidth:700px;Status:no;help:no"
	end if
end function
'--------------------------------------------------------------------------
Function SelectAll()
	dim i, objCheck, Flag
	if document.formname.cmbPurType.selectedIndex = 0 Then
		MsgBox "Select Purchase Type "
		document.formname.cmbPurType.focus()
		exit function
	end if
	if document.formname.ChkAll.checked then
		Flag = True
	else
		Flag = False
	end if

	set root = InvoiceDet.DocumentElement
	sExp ="//ItemDetails/Item"
	Set ItemNode = Root.Selectnodes(sExp)
	if ItemNode.Length > 0 then
		for itr = 0 to ItemNode.Length - 1
			iItemDel = ItemNode.item(itr).Attributes.getNamedItem("ItemCode").value
			iClassDel = ItemNode.item(itr).Attributes.getNamedItem("ClassificationCode").value
			iEntDel = ItemNode.item(itr).Attributes.getNamedItem("EntryNo").value
			set objCheck = eval("document.formname.chkVat"&CStr(iClassDel)&"Z"&CStr(iItemDel)&"Z"&CStr(iEntDel))
			objCheck.Checked = Flag
		next
	end if
End Function

</script>

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
