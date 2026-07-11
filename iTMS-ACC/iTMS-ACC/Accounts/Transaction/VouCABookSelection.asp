<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCashBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 21, 2002
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
<!--#include file="../../include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<!-- XML Data Island -->
<XML ID="UnitBookData">
<Book/>
</XML>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/VouSelection.js"></SCRIPT>
<SCRIPT>
var unitBookDoc = null;

function parseXml(text) {
	return new DOMParser().parseFromString(text || "<Book/>", "text/xml");
}

function loadBookXml(text) {
	var island = window.UnitBookData || document.getElementById("UnitBookData");
	if (island && typeof island.loadXML === "function") {
		island.loadXML(text);
		unitBookDoc = island.XMLDocument || island._doc || island;
		return unitBookDoc.documentElement || island.documentElement;
	}
	unitBookDoc = parseXml(text);
	return unitBookDoc.documentElement;
}

function bookRoot() {
	var island = window.UnitBookData || document.getElementById("UnitBookData");
	if (unitBookDoc && unitBookDoc.documentElement) {
		return unitBookDoc.documentElement;
	}
	if (island && island.documentElement) {
		return island.documentElement;
	}
	if (island && island.XMLDocument) {
		return island.XMLDocument.documentElement;
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

function attr(node, index) {
	return node && node.attributes && node.attributes[index] ? node.attributes[index].nodeValue : "";
}

function selectedBookNode() {
	var nodes = childElements(bookRoot());
	for (var i = 0; i < nodes.length; i += 1) {
		if (attr(nodes[i], 0) === document.formname.selBook.value) {
			return nodes[i];
		}
	}
	return null;
}

function setBookFields() {
	var node = selectedBookNode();
	if (node) {
		document.formname.hBookAccHead.value = attr(node, 2);
		document.formname.hBookOtherUnit.value = attr(node, 3);
	}
	document.formname.hBookName.value = document.formname.selBook.options[document.formname.selBook.selectedIndex].text;
	document.formname.horgName.value = document.formname.selUnitId.options[document.formname.selUnitId.selectedIndex].text;
	return node;
}

function DisplayBook(objUnit) {
	var xhr;
	var iUnitNo;
	var root;
	var nodes;
	var selBook = document.formname.selBook;
	selBook.options.length = 1;
	if (objUnit.selectedIndex !== 0) {
		iUnitNo = objUnit.options[objUnit.selectedIndex].value;
		xhr = new XMLHttpRequest();
		xhr.open("GET", "XMLGetOrgBook.asp?BkCode=01&orgID=" + encodeURIComponent(iUnitNo), false);
		xhr.send(null);
		if (xhr.responseText) {
			root = loadBookXml(xhr.responseText);
			nodes = childElements(root);
			for (var i = 0; i < nodes.length; i += 1) {
				selBook.options[selBook.options.length] = new Option(attr(nodes[i], 1), attr(nodes[i], 0));
			}
		}
	}
}

function VouCreate() {
	var node;
	if (validate()) {
		node = setBookFields();
		if (node && attr(node, 4) === "C" && document.formname.selVouType.value === "C") {
			alert("Book balance is in Credit cannot make Payment");
			document.formname.selVouType.focus();
			return false;
		}
		document.formname.action = "VouCAEntry.asp";
		document.formname.submit();
	}
}

function VouAmend() {
	if (validate()) {
		if (document.formname.txtVouNo.value === "") {
			alert("Select Voucher Number ");
			return;
		}
		setBookFields();
		document.formname.hTransNo.value = document.formname.hTransNo.value;
		document.formname.action = "VouCAAmdEntry.asp";
		document.formname.submit();
	}
}

function VouDel() {
	if (validate()) {
		if (document.formname.txtVouNo.value === "") {
			alert("Select Voucher Number ");
			return;
		}
		setBookFields();
		document.formname.hTransNo.value = document.formname.hTransNo.value;
		document.formname.action = "VouCADelDisplay.asp";
		document.formname.submit();
	}
}

function VouView() {
	if (validate()) {
		if (document.formname.txtVouNo.value === "") {
			alert("Select Voucher Number ");
			return;
		}
		setBookFields();
		document.formname.hTransNo.value = document.formname.hTransNo.value;
		document.formname.action = "VouCAView.asp";
		document.formname.submit();
	}
}

function validate() {
	if (document.formname.selUnitId.selectedIndex < 1) {
		alert("Select Unit");
		return false;
	}
	if (document.formname.selBook.selectedIndex < 1) {
		alert("Select Cash Book");
		return false;
	}
	if (document.formname.selVouType.selectedIndex < 1) {
		alert("Select Voucher type");
		return false;
	}
	return true;
}
</script>



</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hBookAccHead" value="">
<input type="hidden" name="hBookOtherUnit" value="">

<input type="hidden" name="horgName" value="">
<input type="hidden" name="hTransNo" value="">
<input type="hidden" name="hAmendTy" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Cash Voucher
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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">

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
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="108">Organization </td>
                            <td class="FieldCell">
                             <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationListDB%>
                              </select>
                                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Book</td>
                            <td class="FieldCell">
                            <select size="1" name="selBook" class="FormElem">
                        <option value="S">Select Book</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Voucher Type</td>
                            <td class="FieldCell">
                            <select size="1" name="selVouType" class="FormElem">
							<option value="S">Select Voucher Type</option>
							<option value="C">Payment</option>
							<option value="D">Receipt</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Voucher Number</td>
                            <td class="FieldCell">
                            <input type="text" name="txtVouNo" size="20" class="FormElem" readonly>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                           <a href="javascript:popVoucherNo('C','01','CA',document.formname.selUnitId.value,document.formname.selVouType.value,document.formname.selBook.value)">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:popVoucherNo('A','01','CA',document.formname.selUnitId.value,document.formname.selVouType.value,document.formname.selBook.value)">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers"></a>
                            </td>

                                </tr>
                                    </table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Create" onClick="VouCreate()" name="btnCreate" class="ActionButton" >
                                                                <input type="button" value="View" name="B7" onClick="VouView()" class="ActionButton">
                                                                <input type="button" value="Amendment" name="btnAmend" onClick="VouAmend()"  class="ActionButtonX">
                                                                <input type="button" value="Delete" name="btnDel" onClick="VouDel()" class="ActionButton">
                                                                <input type="reset" value="Reset" name="B10" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="10" class="BottomPack" colspan="3">
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
