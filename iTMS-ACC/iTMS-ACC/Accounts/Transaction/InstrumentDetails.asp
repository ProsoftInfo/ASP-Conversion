<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	InstrumentDetails.asp
	'Module Name				:	Fixed Deposit(Transaction)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Jun 15,2005
	'Modified By				:	S.Maheswari
	'Modified On				:	Jun 13,2008
	'Modified By				:   UmaMaeswari S
	'Tables Used				:	04 April 2010
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Bank Voucher - Instrument Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<xml id="VoucherData" >
<Root></Root>
</xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT>
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function field(form, name) {
	var lower = String(name).toLowerCase();
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

function radioValue(form, name) {
	var item = form.elements[name];
	var items = item && item.length != null && !item.tagName ? item : [item];
	for (var i = 0; i < items.length; i += 1) {
		if (items[i] && items[i].checked) {
			return items[i].value;
		}
	}
	return "";
}

function xmlDocument(name) {
	var island = window[name] || document.getElementById(name);
	if (island && island.XMLDocument) {
		return island.XMLDocument;
	}
	if (island && island._doc) {
		return island._doc;
	}
	if (island && island.documentElement) {
		return island;
	}
	return new DOMParser().parseFromString("<Root/>", "text/xml");
}

function clearNamedChildren(root, nodeName) {
	var nodes = [];
	for (var i = 0; root && i < root.childNodes.length; i += 1) {
		if (root.childNodes[i].nodeType === 1 && String(root.childNodes[i].nodeName).toLowerCase() === String(nodeName).toLowerCase()) {
			nodes.push(root.childNodes[i]);
		}
	}
	for (var n = 0; n < nodes.length; n += 1) {
		root.removeChild(nodes[n]);
	}
}

function controlDate(name) {
	var control = field(document.formname, name) || document.getElementById(name);
	if (control && typeof control.getDate === "function") {
		return control.getDate();
	}
	if (control && typeof control.GetDate === "function") {
		return control.GetDate();
	}
	return control ? control.value : "";
}

function setControlDate(name, value) {
	var control = field(document.formname, name) || document.getElementById(name);
	if (!control) {
		return;
	}
	if (typeof control.setDate === "function") {
		control.setDate(value);
	} else if (typeof control.SetDate === "function") {
		control.SetDate(value);
	} else {
		control.value = value;
	}
}

function CheckSubmit() {
	var form = document.formname;
	var doc;
	var root;
	var elem;
	var saveRequest;
	var updateRequest;
	var response;
	if (trim(form.hStatusFlag.value) === "True") {
		alert("This Cheque is Already Used.Updation is Not Possible");
		return;
	}
	if (form.txtStartNo.value === "") {
		alert("Enter Cheque Start No");
		form.txtStartNo.focus();
		return;
	}
	if (form.txtEndNo.value === "") {
		alert("Enter Cheque End No");
		form.txtEndNo.focus();
		return;
	}
	if (parseInt(form.txtStartNo.value, 10) > parseInt(form.txtEndNo.value, 10)) {
		alert("Cheque End No. should be greater than Start No.");
		form.txtEndNo.focus();
		return;
	}
	doc = xmlDocument("VoucherData");
	root = doc.documentElement;
	clearNamedChildren(root, "InstrumentDetails");
	elem = doc.createElement("InstrumentDetails");
	elem.setAttribute("EntryNo", form.hEntNo.value);
	elem.setAttribute("UnitId", form.hUnitId.value);
	elem.setAttribute("UnitName", form.hUnitName.value);
	elem.setAttribute("BookId", form.hBookId.value);
	elem.setAttribute("BookName", (field(form, "hBookName") || { value: "" }).value);
	elem.setAttribute("AccType", form.hAccType.value);
	elem.setAttribute("AccNo", form.hAccNo.value);
	elem.setAttribute("IssuedOn", controlDate("ctlIssueDate"));
	elem.setAttribute("StartNo", form.txtStartNo.value);
	elem.setAttribute("EndNo", form.txtEndNo.value);
	elem.setAttribute("DrawnOn", form.txtDrawnOn.value);
	elem.setAttribute("PayAt", form.txtPayAt.value);
	elem.setAttribute("Status", radioValue(form, "ChkStatus"));
	root.appendChild(elem);

	saveRequest = new XMLHttpRequest();
	saveRequest.open("POST", "XMLSave.asp?Name=BankInsDet&Mod=BA", false);
	saveRequest.setRequestHeader("Content-Type", "text/xml");
	saveRequest.send(new XMLSerializer().serializeToString(doc));

	updateRequest = new XMLHttpRequest();
	updateRequest.open("GET", "BankInsDetailsUpdate.asp", false);
	updateRequest.send(null);
	response = trim(updateRequest.responseText);
	if (response === "") {
		alert("Bank  Instrument Details Updated Successfully");
		window.close();
	} else {
		alert(response);
	}
}

function FnInit() {
	if (document.formname.hIssueDate.value !== "") {
		setControlDate("ctlIssueDate", document.formname.hIssueDate.value);
	} else {
		setControlDate("ctlIssueDate", new Date());
	}
}
</SCRIPT>
<%
Dim sUnitId,sBookId,sUnitName,sBookName,sAccNo,sAccType,sChkN,sChkH,sChkC
Dim sQry,objrs,objrs1,oDOM,iEntNo,sDrawnOn,sPayAt,iStartNo,iEndNo,dtIssueDate,sStatus
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objrs = Server.CreateObject("ADODB.Recordset")
Set objrs1 = Server.CreateObject("ADODB.Recordset")
sUnitId = Request("UnitId")
sBookId = Request("BookId")
sUnitName = Request("UnitName")
sBookName = Request("BookName")
sAccType  = Request("AccType")
sAccNo    = Request("AccNo")
sDrawnOn  = Request("DrwOn")
sPayAt    = Request("PayAt")
IF trim(sAccType) = "S" then sAccType = ""

'Response.Write "sBookId="&sBookId
sQry = "Select EntryNo,DrawnOn,PayableAt,StartNo,EndNo,convert(VarChar,DateOfIssue,103),Status from "&_
	   "Acc_R_BankInstrumentDetails where OUDefinitionID ='"&sUnitId&"' and BookNumber = '"&sBookId&"' "
'Response.Write sQry
objrs.Open sQry,con
IF Not objrs.EOF then
	iEntNo		= objrs(0)
	'sDrawnOn	= objrs(1)
	'sPayAt		= objrs(2)
	iStartNo	= objrs(3)
	iEndNo		= objrs(4)
	dtIssueDate	= objrs(5)
	sStatus		=  objrs(6)
End IF
objrs.Close
IF sStatus	= "" or sStatus	= "N" then sChkN = "Checked"
IF sStatus	= "H" then sChkH = "Checked"
IF sStatus	= "C" then sChkC = "Checked"
dim sStatusFlag
sStatusFlag = False
'Validation for updating ins details
'Response.Write iEntNo &"-----"
IF trim(iEntNo) <> "" and trim(iEntNo) <> "0" then
	sQry = "Select InstrumentEntryNo from Acc_R_BankInstrumentUsage where EntryNo = "&iEntNo
	objrs.Open sQry,con
	Do while not  objrs.EOF
		sQry = "Select Status from Acc_R_BankInstrumentUsage where EntryNo = "&iEntNo&" and InstrumentEntryNo = "&objrs(0)
		'Response.Write sQry
		objrs1.Open sQry,con
		IF Not objrs1.EOF then
			If objrs1(0) = "U" then sStatusFlag = True
		End If
		objrs1.Close
	objrs.MoveNext
	loop
	objrs.Close
End IF 'IF trim(iEntNo) <> "" and trim(iEntNo) <> "0" then
'Response.Write sStatusFlag
'Set Root = VoucherData.documentElement
%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit()" >

<form method="POST" name="formname">
<input type="hidden" name="hUnitId" value="<%=sUnitId%>">
<input type="hidden" name="hBookId" value="<%=sBookId%>">
<input type="hidden" name="hUnitName" value="<%=sUnitName%>">
<input type="hidden" name="hBookNAme" value="<%=sBookName%>">
<input type="hidden" name="hAccType" value="<%=sAccType%>">
<input type="hidden" name="hAccNo" value="<%=sAccNo%>">
<input type="hidden" name="hEntNo" value="<%=iEntNo%>">
<input type="hidden" name="hIssueDate" value="<%=dtIssueDate%>">
<input type="hidden" name="hStatusFlag" value="<%=sStatusFlag%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Instrument
          Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0">

													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
                                                <table cellpadding="0" cellspacing="0">
                                            <tr>
                                        <td>
                                        <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                <td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p>
                                </td>
                                <td class="GroupTitle" width="126" align="center"><p align="center">Instrument Details</td>
                                <td class="GroupTitleRight"><p align="left">&nbsp;</td>
                                    </tr>
                                        </table>
                                        </td>
                                            </tr>
                                            <tr>
                                        <td class="GroupTable">
                                        <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                <td class="MiddlePack" colspan="4"><p align="left"></td>
                                    </tr>
                                    <!--<tr>
                                <td class="FieldCellSub"><p align="left">Unit</p>
                                </td>
                                <td class="FieldCellSub" colspan="3"><p align="left"><span id="spUnit" class="DataOnly">
                                <%=sUnitName%> </span>
                                </td>

                                    </tr>-->
                                    <tr>
										<td class="FieldCellSub"><p align="left">Book
                                          Name
										</td>
										<td class="FieldCellSub" colspan="3"><p align="left"><span id="spBook" class="DataOnly">
                                        <%=sBookName%></span> </p>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Account Type
										</td>
										<td class="FieldCellSub" colspan="3"><span id="spAccType" class="DataOnly"><%=sAccType%></span>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Account Number
										</td>
										<td class="FieldCellSub" colspan="3"><span id="spAccNo" class="DataOnly"><%=sAccNo%></span>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Cheque Book
                                          Issued On
										</td>
										<td class="FieldCellSub" colspan="3">
										<% Response.Write InsertDatePicker("ctlIssueDate") %>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Cheque Start No
										</td>
										<td class="FieldCellSub"><input type="text" name="txtStartNo" value="<%=iStartNo%>" size="15" maxlength="6" class="FormElem">
										</td>
										<td class="FieldCellSub">Cheque End No
										</td>
										<td class="FieldCellSub"><input type="text" name="txtEndNo" Value="<%=iEndNo%>" size="15" maxlength="6" class="FormElem">
										</td>
									</tr>

                                    <tr>
										<td class="FieldCellSub">Drawn On
										</td>
										<td class="FieldCellSub"><input type="text" name="txtDrawnOn" Value="<%=sDrawnOn%>" size="20" class="FormElem" disabled>
										</td>
										<td class="FieldCellSub">Payable At
										</td>
										<td class="FieldCellSub"><input type="text" name="txtPayAt" value="<%=sPayAt%>" size="20" class="FormElem" disabled>
										</td>
                                    </tr>
                                    <tr>
										<td class="FieldCellSub">Status
										</td>
										<td class="FieldCellSub" colspan="3"><input type="radio" value="N" name="ChkStatus" class="FormElem" <%=sChkN%>>
                                          New&nbsp; <input type="radio" name="ChkStatus" value="H" class="FormElem"  <%=sChkH%>>
                                          Hold&nbsp; <input type="radio" name="ChkStatus" value="C" class="FormElem"  <%=sChkC%>>&nbsp;
                                          Closed
										</td>
                                    </tr>
                                    <tr>
										<td class="FieldCellSub">Last Used
                                          Cheque No
										</td>
										<td class="FieldCellSub" colspan="3"><span id="spChkLastNo" >&nbsp;</span>&nbsp;<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" width="11" alt="" height="11" valign="top" style="cursor: pointer" onClick="">
										</td>
                                    </tr>

                                    <tr>
                                <td class="MiddlePack" colspan="4"><p align="left"></td>
                                    </tr>
                                        </table>
                                        </td>
                                            </tr>
                                                </table>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
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
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
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
