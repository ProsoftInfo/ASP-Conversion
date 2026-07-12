<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalReturnEntry2.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 01, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId
Dim dPreCrValue,Temparr,iSalTrNo,dSalInvAmount,sFlag,sTransNo,dSalInvVal,sRetVal
Dim sAmdTy,iSelBook


' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sInvoiceNo=Request.Form("selInvoiceNo")
sBookName=Request.Form("hBookName")
Temparr = Split(sInvoiceNo,":")
'sInvoiceNo = Temparr(0)
'dPreCrValue = Temparr(1)
'Response.write sInvoiceNo
iInvNo = sInvoiceNo
sFlag = Request.Form("sFlag")
sUserid = getUserID()

sInvoiceNo = Request("hTransNo")
sTransNo = Request("hTransNo")
sAmdTy = Request("AmdType")
'Response.Write sTransNo

'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")
sRetVal = GetVouchXML(sInvoiceNo)
oDOM.Load server.MapPath(sRetVal)

set oNodRoot=oDOM.documentElement

for each oNodHeader in oNodRoot.childNodes
	if oNodHeader.nodeName="Header" then
		for Each oNodEntry in  oNodHeader.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.text
			end if

			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue&"&nbsp;-&nbsp;"&oNodEntry.Attributes.Item(1).nodeValue
				iSalTrNo = oNodEntry.Attributes.Item(5).nodeValue
			end if

			if oNodEntry.nodeName="Book" then
				iSelBook = oNodEntry.Attributes.Item(0).nodeValue
			end if

		next
	end if

	sQuery = "Select AmountReceived,AmountReceivable From Acc_T_CreatedReceivables Where CreatedTransNo = "&iSalTrNo
	objRs.Open sQuery,Con
	If Not objRs.EOF Then
		dPreCrValue = objRs(0)
		dSalInvAmount = objRs(1)
	End IF
	objRs.Close

	if oNodHeader.nodeName="Details" then
		set oNodDeatils=oNodHeader
	end if
	if oNodHeader.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodHeader
	end if
	if oNodHeader.nodeName="AgentDetails" then
		set oNodtemp=oNodRoot.removeChild(oNodHeader)
	end if




next

sInvValue=oNodTaxRoot.Attributes.Item(2).nodeValue
dim dInvAmount
dInvAmount = sInvValue

sQuery = "Select VoucherAmount From Acc_T_CreatedVoucherHeader WHere CreatedTransNo = "&iSalTrNo

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	dSalInvVal = objRs(0)
End IF
objRs.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="TaxData" data-src="<%="../temp/transaction/Voucher Entry_CNAmd_"&Session.SessionID&".xml"%>"></script>
<script SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<script SRC="../../scripts/cancel.js"></SCRIPT>
<script >
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function toNumber(value) {
	var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
	return isNaN(parsed) ? 0 : parsed;
}

function formField(name) {
	var form = document.formname;
	return form && (form.elements[name] || form[name]) || null;
}

function xmlObject(name) {
	var element;
	if (window.ITMSModernCompat) {
		window.ITMSModernCompat.upgradeXmlIslands(document);
	}
	element = document.getElementById(name);
	return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
}

function xmlDocument(name) {
	var object = xmlObject(name);
	return object && (object.XMLDocument || object._doc || object) || null;
}

function xmlRoot(name) {
	var doc = xmlDocument(name);
	return doc && doc.documentElement || null;
}

function selectNodes(context, expression) {
	var doc;
	var found;
	var nodes = [];
	if (!context) {
		return nodes;
	}
	if (typeof context.selectNodes === "function") {
		return Array.prototype.slice.call(context.selectNodes(expression));
	}
	doc = context.nodeType === 9 ? context : context.ownerDocument;
	if (!doc || !doc.evaluate) {
		return nodes;
	}
	found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
	for (var i = 0; i < found.snapshotLength; i += 1) {
		nodes.push(found.snapshotItem(i));
	}
	return nodes;
}

function childElements(node, nodeName) {
	var result = [];
	var wanted = nodeName ? String(nodeName).toLowerCase() : "";
	for (var i = 0; node && i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
			result.push(node.childNodes[i]);
		}
	}
	return result;
}

function attr(node, name) {
	var attribute = node && node.attributes && node.attributes.getNamedItem(name);
	return attribute ? attribute.value : "";
}

function setAttr(node, name, value) {
	if (node && node.setAttribute) {
		node.setAttribute(name, value == null ? "" : String(value));
	}
}

function dateValue(control) {
	if (control && typeof control.GetDate === "function") {
		return control.GetDate();
	}
	if (control && typeof control.getDate === "function") {
		return control.getDate();
	}
	return control ? control.value : "";
}

function postXml(url, doc) {
	var request = new XMLHttpRequest();
	request.open("POST", url, false);
	request.send(doc);
	return request.responseText || "";
}

function SaveXML() {
	var form = document.formname;
	var oldInvValue = toNumber(form.hTotinvVal.value);
	var newInvValue = toNumber(form.txtInvValue.value);
	var preCrValue = toNumber(form.hPreCrValue.value);
	var rootNode = xmlRoot("TaxData");
	var responseText;
	var checkVal;
	var saleInvoices;
	var books;

	if (oldInvValue < newInvValue + preCrValue) {
		alert("Returned Invoice Value Should be less than the Invoiced Value ");
		form.txtInvValue.focus();
		return false;
	}

	childElements(rootNode, "Details").forEach(function (details) {
		setAttr(details, "VouDate", dateValue(form.ctlDate));
	});

	selectNodes(rootNode, "//Details/Entry").forEach(function (entry) {
		var entryNo = attr(entry, "No");
		var amountField = formField("txtAmount" + entryNo);
		if (amountField) {
			setAttr(entry, "Amount", amountField.value);
		}
	});

	if (form.optApprove && form.optApprove[0] && form.optApprove[0].checked === true) {
		checkVal = "Y";
		if (form.selUserId.selectedIndex === 0) {
			alert("Select Approver ");
			form.selUserId.focus();
			return false;
		}
	} else {
		checkVal = "N";
	}

	saleInvoices = selectNodes(rootNode, "//SaleInvoice");
	if (saleInvoices.length !== 0) {
		setAttr(saleInvoices[0], "Approval", checkVal);
		setAttr(saleInvoices[0], "Approver", form.selUserId.value);
		setAttr(saleInvoices[0], "CrTransNo", form.hdTransNo.value);
	}

	books = selectNodes(rootNode, "//Book");
	if (books.length !== 0) {
		setAttr(books[0], "BookId", form.selBook.value);
		books[0].textContent = form.selBook.options[form.selBook.selectedIndex].text;
	}

	responseText = postXml("XMLSave.asp?Mod=CNAmd&Name=Voucher Entry", xmlDocument("TaxData"));
	if (trim(responseText) !== "") {
		alert(responseText);
		return false;
	}
	if (trim(form.hFlag.value) === "True") {
		form.action = "AmdAccCrNtGenerate.asp";
	}
	form.submit();
	return true;
}

function EnbApp(sObj) {
	if (sObj.value === "Y") {
		document.formname.selUserId.disabled = false;
	} else {
		document.formname.selUserId.selectedIndex = 0;
		document.formname.selUserId.disabled = true;
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="VouCNSalRetAdjAmd2.asp">
<Input type="hidden" name="hdTransNo" value="<%=iSalTrNo%>">
<Input type="hidden" name="hPreCrValue" value="<%=dPreCrValue%>">

<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hTransNo" value="<%=sTransNo%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Return Credit
          Note 		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="90">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<tr><td align="center">Adjustments</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<tr><td align="center">Voucher</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <!--tr>
                            <td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: pointer" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
                    <font size="4" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell">
                    &nbsp;
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Unit </td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sOrgName%></span></td>
                                                    <td class="FieldCellSub" width="100"><p align="left">Date</p></td>
                                                    <td class="FieldCellSub" width="145">
                                                 <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Party
                                                      Name</td>
                                                    <td class="FieldCell" colspan="3">  <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>

                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Invoice No-Date</td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sInvoiceNo%> </span></td>
                                                    <td class="FieldCellSub" width="100">Invoice Value</td>
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=FormatNumber(dSalInvVal,2,,,0)%> </span></td>
                                                        </tr>

                                                        <tr>
															<td class="FieldCellSub" width="100">Credit Note Value</td>
															<td class="FieldCell" width="200">  <span class="DataOnly"><%=FormatNumber(dPreCrValue,2,,,0)%> </span></td>
                                                         </tr>
                                                         <%IF CStr(sAmdTy) = "A" Then %>
																<tr>
																	<td class="FieldCellSub" width="100">Book</td>
																	<td class="FieldCell" colspan="3">
																		<select size="1" name="selBook" class="FormElem">
																		<%
																			sQuery = "Select BookNumber,BookName From VwOrgBookNames Where  "&_
																					 "OUDefinitionID = '"&sOrgId&"' And BookCode = '07' Order By BookName "
																			With objRs
																				.CursorLocation = 3
																				.CursorType = 3
																				.Source = sQuery
																				.ActiveConnection = Con
																				.Open
																			End With
																			Set objRs.ActiveConnection = Nothing
																			Do While Not objRs.EOF
																				IF CStr(iSelBook) = CStr(objRs(0)) Then
																		%>
																					<Option Value="<%=objRs(0)%>" Selected><%=objRs(1)%></Option>
																		<%		Else%>
																					<Option Value="<%=objRs(0)%>"><%=objRs(1)%></Option>
																		<%
																				End IF
																				objRs.MoveNext
																			Loop
																			objRs.Close

																		%>
																		 </select>
																	 </td>
																</tr>
															<%End IF%>
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            <tr>
                <td></td>
                <td valign="top" width="100%">
                <div class="frmBody" id="frm2" style="width: 582; height:320;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
    <td class="ExcelHeaderCell" align="center" colspan="3">Invoice</td>
    <td class="ExcelHeaderCell" align="center" colspan="2">Returned</td>

        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="75">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="75">Value</td>
    <td class="ExcelHeaderCell" align="center" width="75">Discount</td>
    <td class="ExcelHeaderCell" align="center">Qty</td>
    <td class="ExcelHeaderCell" align="center">Amount</td>

        </tr>
<%
	Dim sSalQty
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.GetNamedItem("No").value
		sDescription=oNodEntry.Attributes.GetNamedItem("PayTo").value
		sAmount=oNodEntry.Attributes.GetNamedItem("Amount").value
		sQty=oNodEntry.Attributes.GetNamedItem("Qty").value
		sValue=oNodEntry.Attributes.GetNamedItem("ActValue").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value
		oNodEntry.Attributes.GetNamedItem("Rate").value=CDbl(oNodEntry.Attributes.GetNamedItem("Amount").value)/CDbl(sQty)
		dTotal=CDbl(dTotal)+CDbl(sAmount)

		sQuery = "Select InvoicedQuantity From Acc_T_CreatedVoucherDetails "&_
				 "Where CreatedTransNo = "&iSalTrNo&" And VoucherEntryNumber = "&iSno&" "

		objRs.Open sQuery,Con
		If Not objRs.EOF Then
			sSalQty = objRs(0)
		End IF
		objRs.Close

%>

	<Input type="hidden" name=hQty<%=iSno%> value="<%=sSalQty%>">
	<Input type="hidden" name=hInvVal<%=iSno%> value="<%=sValue%>">

    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sQty,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sValue,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sDiscount,2,,,0)%></td>
    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtqty<%=iSno%>1" onBlur="setQty(this,'<%=iSno%>')" value="<%=FormatNumber(sQty,3,,,0)%>" class="Formelem" size="13"></td>

    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iSno%>" onBlur="setTotal(this,'<%=iSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13"></td>

        </tr>
<%
	next
%>


        <tr>
    <td align="center" colspan="2"></td>
    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
     <td align="right" colspan="2"></td>
     <td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotal" value="<%=FormatNumber(dTotal,2,,,0)%>" class="Formelem" size="13"></td>
        </tr>



<%
	Dim sCheckExp,CheckNode
'	dim dInvAmount
'	dInvAmount=dTotal
	Dim iCheck
	For Each oNodEntry in oNodTaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.GetNamedItem("CatCode").value
		sTaxCode=oNodEntry.Attributes.GetNamedItem("TaxCode").value
		sTaxMode=oNodEntry.Attributes.GetNamedItem("TaxMode").value
		sFormula=oNodEntry.Attributes.GetNamedItem("TaxFormula").value
		dTaxValue=oNodEntry.Attributes.GetNamedItem("TaxValue").value


		If sCatCode = "0" and sTaxCode = "0" and sTaxMode = "0" Then
			iRndOff = 0
		Else
			sCheckExp = "//TaxDetails/Tax[@CatCode="&sCatCode&" and @TaxCode="&sTaxCode&" and @RoundOff]"
			Set CheckNode = oNodRoot.selectNodes(sCheckExp)
			IF CheckNode.length <> 0 Then
				iRndOff = oNodEntry.Attributes.GetNamedItem("RoundOff").value
			Else
				iRndOff = 0
			End IF
		End If
		sTaxName=oNodEntry.Text
		If iRndOff = 1 Then
			dTax = FormatNumber(Round(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,0),2,,,0)
		Else
			dTax = FormatNumber(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,2,,,0)
		End If
%>
			<tr>
				<td align="center" colspan="2"></td>
				<td class="ExcelSerial" align="center" colspan="3"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="P" then %>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="Formelem">&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  onBlur="setTaxAmount('<%=sCatCode%>','<%=sTaxCode%>',this)"size="11" class="Formelem"></td>
				    </tr>
			<%
	next
oDOM.save server.MapPath("../temp/transaction/Voucher Entry_CNAmd_"&Session.SessionID&".xml")

%>

		<input type="hidden" name="hTotinvVal" value="<%=dSalInvAmount%>">
        <tr>
        <td align="center" colspan="2"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="Formelem">
    </td>
        </tr>
        <tr>
        </tr>
        <tr>
			<td align="left" class="FieldCell" colspan="2" valign="Top">
				Approval
			</td>
			<td align="center" class="FieldCellSub" colspan="2">
			<Input type="radio" name="optApprove" checked value="Y" onClick="EnbApp(this)"> Yes &nbsp;&nbsp;&nbsp;
			<Input type="radio" name="optApprove" value="N" onClick="EnbApp(this)"> No &nbsp;&nbsp;&nbsp;
			</td>
        </tr>
        <tr>
			<td align="left" class="FieldCell" colspan="2" valign="Top">
				Immediate Approver
			</td>
			<td align="center" class="FieldCellSub" colspan="2">
			<select size="1" name="selUserId" class="FormElem">
              <option value="I">Immediate Approver</option>
                <%=populateEmployeeWithVal(sUserId)%>
                    </select>
			</td>
        </tr>
            </table>
                </div>
                </td>
                <td></td>
                            </tr>
                                                        <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                                        </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                                <input type="button" value="Next" name="B2" class="ActionButton" onClick="SaveXML()" >
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" onClick="Cancel('CreditVouchers.ASP')" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
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
</html>
