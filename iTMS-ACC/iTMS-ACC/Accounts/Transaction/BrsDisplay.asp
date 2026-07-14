<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BrsDisplay.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 27,2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->


<%

'XML DOM Variables
Dim oDOM,nodHeader,Root,sQuery,objrs
dim sBookName,dPassBalance,sPassCrDr,dBookBalance,sBookCrDr,sFromDt,sToDt
Dim sRecVal,sAccName,iBookNo,sQry,iAccHead,sOrgId
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml")

set Root=oDOM.documentElement
Set objrs = Server.CreateObject("ADODB.RecordSet")
sOrgId=root.Attributes.Item(0).nodeValue
iBookNo = root.Attributes.Item(1).nodeValue
sBookName= root.Attributes.Item(2).nodeValue
sFromDt= root.Attributes.Item(3).nodeValue
sToDt= root.Attributes.Item(4).nodeValue
dPassBalance= root.Attributes.Item(5).nodeValue
sPassCrDr= root.Attributes.Item(6).nodeValue
dBookBalance= root.Attributes.Item(7).nodeValue
sBookCrDr= root.Attributes.Item(8).nodeValue

sRecVal = Request("hChkVal")
'Response.Write "sChkVal="&sChkVal
sQuery="select AccountHead,AccountDescription from Acc_M_GLAccountHead where AccountHead=(select BankChargesHead from Acc_M_BankDetails where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='02' and BookNumber="&iBookNo&")"
'Response.Write sQuery
with objRs
.CursorLocation = 3
.CursorType = 3
.Source = sQuery
.ActiveConnection = con
.Open
end with
set objRs.ActiveConnection = nothing

if not objRs.EOF then
	iAccHead=objRs(0)
	sAccName=objRs(1)
end if
objrs.close
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>BRS Display</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml"%>">
</script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script >
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function field(name) {
	var lower = String(name).toLowerCase();
	var form = document.formname;
	if (form.elements[name]) {
		return form.elements[name];
	}
	for (var i = 0; i < form.elements.length; i += 1) {
		if (String(form.elements[i].name).toLowerCase() === lower) {
			return form.elements[i];
		}
	}
	return null;
}

function fields(name) {
	var item = field(name);
	if (!item) {
		return [];
	}
	if (item.length != null && !item.tagName) {
		return Array.prototype.slice.call(item);
	}
	return [item];
}

function dateControl() {
	return field("ctlDate") || document.getElementById("ctlDate");
}

function setDateEnabled(enabled) {
	var control = dateControl();
	if (control) {
		control.disabled = !enabled;
	}
}

function xmlRoot(value) {
	if (!value) {
		return null;
	}
	if (typeof value === "string") {
		return new DOMParser().parseFromString(value, "text/xml").documentElement;
	}
	if (value.nodeType === 1) {
		return value;
	}
	if (value.documentElement) {
		return value.documentElement;
	}
	if (value.XMLDocument && value.XMLDocument.documentElement) {
		return value.XMLDocument.documentElement;
	}
	if (value._doc && value._doc.documentElement) {
		return value._doc.documentElement;
	}
	return null;
}

function childElements(node) {
	var nodes = [];
	for (var i = 0; node && i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1) {
			nodes.push(node.childNodes[i]);
		}
	}
	return nodes;
}

function firstChildElement(node, name) {
	var wanted = String(name).toLowerCase();
	var nodes = childElements(node);
	for (var i = 0; i < nodes.length; i += 1) {
		if (String(nodes[i].nodeName).toLowerCase() === wanted) {
			return nodes[i];
		}
	}
	return null;
}

function openModernDialog(url, args, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	} else {
		window.open(url, "_blank", "height=350,width=600,resizable=no,status=no");
	}
}

function BankCharges() {
	var enabled = document.formname.C1.checked === true;
	var crdr = fields("selCRDR");
	document.formname.hChkVal.value = enabled ? "True" : "False";
	setDateEnabled(enabled);
	for (var i = 0; i < crdr.length; i += 1) {
		crdr[i].disabled = !enabled;
	}
	document.formname.txtNarration.disabled = !enabled;
	document.formname.txtAmount.disabled = !enabled;
	if (!enabled) {
		document.formname.hVouDate.value = "";
		document.formname.hAmt.value = "";
		document.formname.hCrDr.value = "";
	}
}

function applyBrsPopupReturn(outValue) {
	var root = xmlRoot(outValue);
	var entries = childElements(root);
	if (!root) {
		return;
	}
	document.formname.hVouDate.value = root.getAttribute("VouDate") || "";
	for (var i = 0; i < entries.length; i += 1) {
		var node = entries[i];
		var accHead;
		var narration;
		if (String(node.nodeName).toLowerCase() === "entry") {
			document.formname.hAmt.value = node.getAttribute("Amount") || "";
			document.formname.hCrDr.value = node.getAttribute("CRDR") || "";
			accHead = firstChildElement(node, "AccHead");
			narration = firstChildElement(node, "Narration");
			if (accHead) {
				document.formname.hAccNo.value = accHead.getAttribute("No") || "";
			}
			if (narration) {
				document.formname.hNarr.value = narration.textContent || "";
			}
		}
	}
}

function BrsPopup() {
	var insDet = document.formname.hInsDet.value;
	if (document.formname.C1.checked !== true) {
		document.formname.hChkVal.value = "False";
		return;
	}
	document.formname.hChkVal.value = "True";
	openModernDialog("BrsCommEntry.asp?InsDet=" + encodeURIComponent(insDet), "", "dialogHeight:350px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No", applyBrsPopupReturn);
}

function selectedCrDr() {
	var crdr = fields("selCRDR");
	for (var i = 0; i < crdr.length; i += 1) {
		if (crdr[i].checked) {
			return crdr[i].value;
		}
	}
	return "";
}

function controlDateValue() {
	var control = dateControl();
	if (control && typeof control.getDate === "function") {
		return control.getDate();
	}
	if (control && typeof control.GetDate === "function") {
		return control.GetDate();
	}
	return control ? control.value : "";
}

function finalDone(bFlag) {
	if (trim(bFlag) === "P") {
		if (document.formname.C1.checked !== true && !window.confirm("Bank Charges Not Applicable.Do U want to Continue?")) {
			return;
		}
		if (document.formname.C1.checked === true) {
			document.formname.hVouDate.value = controlDateValue();
			if (document.formname.txtAmount.value === "") {
				alert("Enter Amount");
				return;
			}
			document.formname.hAmt.value = document.formname.txtAmount.value;
			document.formname.hCrDr.value = selectedCrDr();
		}
		document.formname.action = "BrsGenerate.asp";
		document.formname.submit();
	}
}

function CheckRecon() {
	setDateEnabled(false);
}

function SelChk() {
	var values = [];
	var count = parseInt(document.formname.hCtr.value, 10) || 0;
	for (var i = 1; i < count; i += 1) {
		var item = field("Chk" + i);
		if (item && item.checked) {
			values.push(item.value);
		}
	}
	document.formname.txtNarration.value = values.join(",");
}
</script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="CheckRecon()">

<form method="POST" name="formname">
<input type="hidden" name="hRecVal" Value="<%=sRecVal%>">
<input type="hidden" name="hChkVal" value="False">
<input type="hidden" name="hVouDate" value="">
<input type="hidden" name="hAccNo" value="<%=iAccHead%>">
<input type="hidden" name="hAmt" value="">
<input type="hidden" name="hCrDr" value="">
<input type="hidden" name="hNarr" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Bank Reconciliation</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Instruments
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
                                    <tr>
                                    	<td align="center">Reconciled List</td>
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
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
												<DIV class=frmBody id=frm1 style="width: 585; height:140;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"></td>
                                        <td class="ExcelHeaderCell" align="center">Instrument No. - Date</td>
                                        <td class="ExcelHeaderCell" align="center" width="150">Paid/Recd </td>
                                        <td class="ExcelHeaderCell" align="center" width="80">Receipt</td>
                                        <td class="ExcelHeaderCell" align="center" width="80">Payment</td>
                                        <td class="ExcelHeaderCell" align="center" width="50">Date</td>
                                            </tr>
<%
dim sVouNo,sInstDet,dAmont,sTransType,sClearedOn,sPayRec,iVouNo,iInsDet
dim dPayTotal,dRecTotal,iSNo,dTotal,sExp,tempNode,iCounter,iTransNo,iCtr
iCtr=1

sExp="//Voucher"
set tempNode=Root.selectNodes(sExp)
for iCounter=0 to tempNode.length-1
	iSNo = tempNode.item(iCounter).Attributes.Item(0).nodeValue
	iTransNo = tempNode.item(iCounter).Attributes.Item(1).nodeValue
	sVouNo=tempNode.item(iCounter).Attributes.Item(3).nodeValue
	sInstDet=tempNode.item(iCounter).Attributes.Item(4).nodeValue
	sPayRec=tempNode.item(iCounter).Attributes.Item(5).nodeValue
	dAmont=tempNode.item(iCounter).Attributes.Item(6).nodeValue
	sTransType=tempNode.item(iCounter).Attributes.Item(7).nodeValue
	sClearedOn=tempNode.item(iCounter).Attributes.Item(8).nodeValue

	if tempNode.item(iCounter).Attributes.Item(2).nodeValue="Y" then

	'Response.Write sVouNo & "<br>"
	iVouNo = iVouNo &","& sVouNo
	iInsDet = iInsDet &","& sInstDet
%>
                                            <tr>
                                        <td class="ExcelSerial" align="center"><%=iCtr%></td>
                                        <td class="ExcelDisplayCell" align="right"><p align="left">
                                        <input type="Checkbox" name="Chk<%=iCtr%>" value="<%=sInstDet%>" onclick="SelChk()" checked >
                                        </td>
                                        <td class="ExcelDisplayCell" align="right"><p align="left"><%=sInstDet%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=sPayRec%></td>
                                        <td class="ExcelDisplayCell" align="right">
                                        <%if sTransType="R" then Response.Write FormatNumber(dAmont,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right">
                                        <%if sTransType="P" then Response.Write FormatNumber(dAmont,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=sClearedOn%></td>
                                            </tr>
<%	iCtr=CInt(iCtr)+1

	else

		if sTransType="P" then
			dPayTotal=CDbl(dPayTotal)+ CDbl(dAmont)
		else
			dRecTotal=CDbl(dRecTotal)+ CDbl(dAmont)
		end if

	end if
	'iCtr=CInt(iCtr)+1
next
iVouNo = mid(iVouNo,2)
iInsDet = mid(iInsDet,2)
%>
   <input type="hidden" name="hVouNo" value="<%=iVouNo%>">
   <input type="hidden" name="hInsDet" value="<%=iInsDet%>">
   <input type="hidden" name="hTransNo" value="<%=iTransNo%>">
   <input type="hidden" name="hCtr" value="<%=iCtr%>">
                                                </table>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly">
                                <tr>
									<td class="MiddlePack" colspan="3"> </td>
                                </tr>
                                <tr>
                                <!--td class="FieldCellSub" >Bank Charges Applicable</td>
								<td class="FieldCell"><input type="checkbox" name="ChkBanChrg" class="formelem" value="" onclick="BrsPopup()"></td-->
                                </tr>
									<tr>
												<td class="FieldCell">Book Balance
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=formatnumber(dBookBalance,2,,,0)%>&nbsp;<%=sBookCrDr%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Passbook Balance as on <%=sToDt%>
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=formatnumber(dPassBalance,2,,,0)%>&nbsp;<%=sPassCrDr%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Cheque / DD deposited but yet to be cleared
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=FormatNumber(dRecTotal,2,,,0)%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Cheque / DD issued but yet to be cleared
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=FormatNumber(dPayTotal,2,,,0)%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell" valign="top">Reason
												</td>
												<td class="FieldCellSub"><textarea rows="2" name="S1" cols="40" class="FormElem"></textarea>
												</td>
											</tr>

										</table>
									</td>
									<td align="center" class="ClearPixel" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" width="5" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%"><center>
										<div align="left">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td>
														<table cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td class="GroupTitleLeft" width="10">&nbsp;
																</td>
																<td class="GroupTitle" width="120">
																	<p align="center">
																	<input type="checkbox" name="C1" value="ON" class="FormElem" onclick="BankCharges()">
 																	Bank Charges
																</td>
															</center><td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>

													</table>
												</td>
											</tr>

											<tr>
												<td class="GroupTable"><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="MiddlePack" colspan="3">
															</td>
														</tr>

														<tr>
															<td class="ClearPixel" width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
															<td class="FieldCell" width="100%"></center>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td class="FieldCell">Book Name
																	</td>
																	<td class="FieldCellSub">
																		<span class="DataOnly"><%=sBookName%>&nbsp;</span>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell">Charges Account Head
																	</td>
																	<td class="FieldCellSub">
																		<span class="DataOnly"><%=sAccName%>&nbsp;</span>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell">Charge Type
																	</td>
																	<td class="FieldCellSub">
																		<input type="radio" value="C" checked name="selCRDR" class="FormElem" disabled>
 																		Receipts
																		<input type="radio" name="selCRDR" value="D" class="FormElem" disabled>
 																		Payments
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell">Charges Date
																	</td>
																	<td class="FieldCellSub">
																	 <% ' Function Call to Insert Date Picker
																		Response.Write InsertDatePicker("ctlDate")
																		%>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell" valign="top">Narration
																	</td>
																	<td class="FieldCellSub"><textarea rows="2" name="txtNarration" cols="60" class="FormElem" Disabled ><%=iInsDet%></textarea>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell" valign="top">Amount
																	</td>
																	<td class="FieldCellSub">
																		<input type="text" name="txtAmount" size="10" class="FormElem" disabled>
																	</td>
																</tr>

															</table>
														</td>
														<td class="ClearPixel" width="5">
															<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
														</td>
													</tr>

													<tr>
														<td class="MiddlePack" width="267" colspan="3">
														</td>
													</tr>

													</table>
												</td>
											</tr>

										</table>
										</div>
									</td>
									<td align="center" class="ClearPixel" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
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
                                                        <!--input type="button" value="Save & Print" onClick="finalDone('P')" name="B2" class="ActionButtonX">
                                                        <input type="button" value="Done" name="B6" onClick="finalDone('S')" class="ActionButton"-->
                                                        <% IF trim(sRecVal) = "BR" then %>
															<input type="button" value="Save" onClick="finalDone('P')" name="B2" class="ActionButtonX">
														<% Else %>
															<input type="button" value="Save" onClick="finalDone('P')" name="B2" class="ActionButtonX">
                                                        <% End IF %>
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
</form>
</BODY>
</HTML>
