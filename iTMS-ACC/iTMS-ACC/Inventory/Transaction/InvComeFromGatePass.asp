<%	option explicit	%>
<%
	'Program Name				:	InvComeFromGatepass.asp
	'Module Name				:	Gate Pass - Service
	'Author Name				:	Ragavendran R
	'Created On					:	August 26,2010
	'Modified On				:
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
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>

<!-- #include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/PurchaseTermsConditions.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/PurChkItemSpecPack.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<!---XML Data Island---->
<script type="application/xml" data-itms-xml-island="1" id="TempData"> <Root/> </script>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ITEMData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="TaxFormData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData">
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTaxData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="InvoiceDet" data-src="<%="../temp/transaction/InvItemValue_PUR_"&Session.SessionID&".xml"%>"></script>

<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/calcAlternateUoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>

<%

'declaring variables
Dim iParTypeID,sParTypeName,sParType,sTraUnit,sRecptAg,sParSubType
Dim rsItem,dcrs,oDom,objDom2,Root,objFs,sAccRoot,sTempNode,SubNode,SubNode1,rsTemp
Dim newElem1
Dim iGPNo,iActivityNo,iPartyCode,iClassCode,iItemCode,iQtyRecd,iEntryNo
Dim sOrgID,sRoundOffHead,sSql,sPartyName,sClassDesc,sItemDesc,UoMCode,sItemType
Dim dOrderRate,dItemValue,sExp,sRateType

set rsTemp = Server.CreateObject("ADODB.RecordSet")
Set rsItem = Server.CreateObject("ADODB.Recordset")
Set dcrs = Server.CreateObject("ADODB.RecordSet")
set objDOM2 = server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iGPNo = trim(Request.QueryString("iGPNo"))
sOrgID = trim(Request.QueryString("ForUnit"))
sRateType = trim(Request.QueryString("RateType"))

iActivityNo = "10"		'Activity Number for "RECEIPT INVOICE"

dOrderRate=0
dItemValue=0

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



sSql = "Select GatePassNo,OrganisationCode,InvoiceType,ForInvoiceNo,PartyCode,TypeOfItems,MarkedOn,DCCode,"&_
	   "NoofPacks,PackingType,Transport,TakenBy,DeliveryBy,RefType from FORGATEPASSHEADER where GatePassNo = "& iGPNo  &""
dcrs.open sSql,con
if not dcrs.eof then
	iPartyCode = dcrs(4)
	sItemType = dcrs(5)
end if
dcrs.close



sSql = "Select PartyName from App_M_PartyMaster where PartyCode=" & trim(iPartyCode) & ""
'Response.Write sSql
With rsItem
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End With
Set rsItem.ActiveConnection = nothing
if not rsItem.EOF then
	sPartyName = rsItem(0)
End if
rsItem.Close

sSql = "Select PartyType,PartySubType from VWORGParty where PartyCode = "& trim(iPartyCode)
'Response.Write sSql
rsItem.Open sSql,con
if not rsItem.EOF then
	sParType = trim(rsItem(0))
	sParSubType = trim(rsItem(1))
end if
rsItem.Close

	Set SubNode= oDOM.createElement("InvoiceHeader")
	Root.appendChild SubNode
	Set SubNode1=oDOM.createElement("ItemDetails")
	Root.appendChild SubNode1

	Set NewElem1 = oDOM.createElement("Header")
	newElem1.setAttribute "OrgID", sOrgID
	newElem1.setAttribute "Party", sPartyName
	newElem1.setAttribute "PurchaseType", ""
	newElem1.setAttribute "Currency",""
	newElem1.setAttribute "InvAgainst", "Receipt"
	newElem1.setAttribute "RefNum", ""
	newElem1.setAttribute "PartyCode", iPartyCode
	newElem1.setAttribute "PartyType",sParType
	newElem1.setAttribute "PartySubType",sParSubType
	newElem1.setAttribute "CurrencyNo", ""
	newElem1.setAttribute "DespatchMode",""
	newElem1.setAttribute "PaymentMode",""
	newElem1.setAttribute "PayTerms",""
	newElem1.setAttribute "IssueBank",""
	newElem1.setAttribute "BenificiaryBank",""
	newElem1.setAttribute "PricingBasis",""
	newElem1.setAttribute "Transporter",""
	newElem1.setAttribute "LoadingPort",""
	newElem1.setAttribute "DestPort",""
	newElem1.setAttribute "Remarks",""
	newElem1.setAttribute "SuppInvNo",""
	newElem1.setAttribute "SuppInvDt",""
	newElem1.setAttribute "TransporterFlag",""
	newElem1.setAttribute "PoNo",""
	newElem1.setAttribute "ConfNum",""
	newelem1.setAttribute "InvoiceFlag",""
	newelem1.setAttribute "InvValue",0
	newelem1.setAttribute "RoundOff",0
	newelem1.setAttribute "SuppCode",""
	newelem1.setAttribute "ItemType",""
	SubNode.appendChild NewElem1

	'==============================================================================
	sSql ="Select isNull(ClassificationCode,0),isNull(ItemCode,0),isNull(Quantity,0),"&_
	" isNull(Description,''),isNull(InvoicedUoM,'') from FORGATEPASSDETAILS where GatePassNo = "& iGPNo

	'Response.Write ssql
		With rsItem
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		End With

		Set rsItem.ActiveConnection = nothing

		Set iClassCode = rsItem(0)
		Set iItemCode= rsItem(1)
		Set iQtyRecd = rsItem(2)
		Set sItemDesc = rsItem(3)
		Set UomCode=rsItem(4)
	'	set sStockType = rsItem(7)
	'	set nActItemRate = rsItem(8)

		iEntryNo = 0
		If not rsItem.EOF then
		Do While Not rsItem.EOF

			'Response.Write "<p> sRefNum = "& sRefNum
			'' to add Qty Validation for Receipt

				sSql = "Select ItemDescription from vwItem where ItemCode ="& iItemCode &"  and ClassificationCode = "& iClassCode &""
				rsTemp.Open sSql,con
				if not rsTemp.EOF then
					sItemDesc = trim(rsTemp(0))
				end if
				rsTemp.Close


					iEntryNo = iEntryNo + 1
					Set newElem1 = oDOM.createElement("Item")

					newElem1.setAttribute "ItemCode", iItemCode
					newElem1.setAttribute "ClassificationCode", iClassCode
					newElem1.setAttribute "ItmDescription", sItemDesc
					newElem1.setAttribute "Uom", UomCode
					newElem1.setAttribute "Qty", iQtyRecd
					newElem1.setAttribute "Rate","0"
					newElem1.setAttribute "DisPer", "0"
					newElem1.setAttribute "DisAmount", "0"
					newElem1.setAttribute "NettBasic", "0"
					newElem1.setAttribute "UomDesc", UomCode
					newElem1.setAttribute "EntryNo", iEntryNo
					newElem1.setAttribute "RatePerQtyUoM", "0"
					newElem1.setAttribute "SourceEntryNo", iEntryNo
					newElem1.setAttribute "PurchaseType", ""
					newElem1.setAttribute "Amount", "0"
					newElem1.setAttribute "ItemValue", "0"
					newElem1.setAttribute "ItemRate", "0"
					newElem1.setAttribute "RateUOM", ""
					newElem1.setAttribute "StockType",""
					newElem1.setAttribute "VAT", ""
					SubNode1.appendChild NewElem1
			rsItem.MoveNext
		Loop
		End if
		rsItem.Close
		'Root.appendChild SubNode1
		'*******************************************************************************************************
		oDOM.save server.MapPath("..\temp\transaction\InvItemValue_PUR_"&Session.SessionID&".xml")

%>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/invComeFromGatePassModern.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="addItem()">

<form method="POST" name="formname" action>
	<input type=hidden name="hItemtype" value="<%=sItemType%>">
	<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
	<input type="hidden" name="hPartyCode" value="<%=iPartyCode%>">
	<input type="hidden" name="hCurrDate" value="<%=FormatDate(Date)%>">
	<input type="hidden" name="hPartyType" value="<%%>">
	<input type="hidden" name="hPSubType" value="<%%>">
	<input type="hidden" name="txtSuppInvDt" value="">
	<Input type="hidden" name="hConfNum" value="<%%>" >
	<input type="hidden" name="hFlag" value="<%%>">
	<input type="hidden" name="hRcptno" value="<%%>">
	<input type="hidden" name="hRcptCode" value="<%%>">
	<input type="hidden" name="txtPartyName" value="">
	<input type="hidden" name="hRoundOffAccHead" value="<%=sRoundoffHead%>">

	<input type="hidden" name="hSuppInvDate" value="<%%>">
	<input type="hidden" name="hGPNo" value="<%=iGPNo%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Purchase Invoice Entry - Service Return
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
												<td class="FieldCell" valign="top">Supplier Name
												</td>
												<td class="FieldCellSub" colspan="4">
													<span class="Dataonly" id="spnSuppName"><%=sPartyName%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Supplier Type
												</td>
												<td class="FieldCellSub" colspan="4">
													<select size="1" class="Formelem" name="cmbPartyType">
															<option value="0" selected>Select</option>
															<%populatePartyTypeForParty sParType,sParSubType %>
													</select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Supplier Inv. No. / Dt.
												</td>
												<td class="FieldCellSub" colspan="4">
													<table>
													<tr><td>
													<input type="text" name="txtSuppInvNo" size="30" class="FormElem" align=top value="<%%>">
													</td><td>
 													<% ' Function Call to Insert Date Picker
													Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
													<td>
													<!--	<select size="1" class="Formelem" name="cmbBillType">
															<option value="0" selected>Select Bill type</option>
															<option value="P">Credit bill</option>
															<option value="C">Cash bill</option>
														</select>-->
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
														<option  value="0">--------------ITEMWISE-------------</option>
														<%

															popSelPurTypeFull("0")
														%>
													</select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Currency
												</td>
												<td class="FieldCellSub"><select size="1" name="D29" class="Formelem">
														<%
														populateCurrency
								  						'popSelCurrency(Curr1)
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
														%>
														<Option value="<%=dcrs(0)%>"><%=dcrs(1)%></Option>
														<%
														dcrs.MoveNext
														loop
														dcrs.Close
													%>


													</select>
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Narration
												</td>
												<td class="FieldCellSub" colspan=4>
													<textarea id="txtNarration" class="FormElem" rows=2 cols=50></textarea>
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
										<div class="frmBody" id="frm2" style="width: 585; height:242;">
											<table border="0" cellspacing="1" class="ExcelTable" width="1000" id="tblItemDet">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
														<a href="#"><img name="ImgDeleteIcon" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onClick="DeleteItems()" alt="Delete Selected Item"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description
														<a><img name=imgPurchaseDet border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" disabled=true alt="Enter Purchase Details" onclick="ShowPurchaseDet('<%=sOrgID%>','<%=sItemType%>')" width="15" height="15" style="cursor:hand"></a>
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
													<td class="ExcelHeaderCell" align="center" rowspan="2">Item Value
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
												<td><textarea rows="2" name="mTextAreaRemarks" cols="93" class="FormElem"></textarea>
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
