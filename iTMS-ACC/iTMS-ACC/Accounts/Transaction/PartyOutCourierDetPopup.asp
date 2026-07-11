<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PartyOutstandingPrevReminder.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	UmaMaheswari S
	'Created On					:	09 April 2011
	'Modified By				:	
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Reminder Preview</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=10">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<XML id="GenReminder"><Root/></XML>
<XML id="OutData"><Root Done=""/></XML>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script>
(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || {};
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlRoot(value) {
		var object = typeof value === "string" ? window[value] || document[value] : value;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDocument(value) {
		var object = typeof value === "string" ? window[value] || document[value] : value;
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function serializeXml(value) {
		var doc = xmlDocument(value);
		var root = xmlRoot(value);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function firstReminderNode(root) {
		if (!root) {
			return null;
		}
		if (String(root.nodeName).toLowerCase() === "reminder") {
			return root;
		}
		return root.getElementsByTagName ? root.getElementsByTagName("Reminder")[0] : null;
	}

	function selectedSendBy() {
		var radios = field("radSendBy");
		var items = radios && radios.length != null && !radios.tagName ? Array.prototype.slice.call(radios) : radios ? [radios] : [];
		var selected = items.filter(function (item) {
			return item.checked;
		})[0];
		return selected ? selected.value : "";
	}

	window.CloseWindow = function () {
		window.close();
	};

	window.Submit = function () {
		var source = window.dialogArguments;
		var reminder;
		var xhr;
		var outRoot;
		ensureCompat();
		reminder = firstReminderNode(xmlRoot(source));
		if (reminder) {
			reminder.setAttribute("SENDBY", selectedSendBy());
			reminder.setAttribute("NAME", trim(field("txtCouComName") && field("txtCouComName").value));
			reminder.setAttribute("ID", trim(field("txtCouTransID") && field("txtCouTransID").value));
			reminder.setAttribute("ADDRESS", trim(field("txtCouComAddress") && field("txtCouComAddress").value));
		}
		xhr = new XMLHttpRequest();
		xhr.open("POST", "GenReminderInsert.asp", false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(serializeXml(source));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
		} else {
			alert("Remonder Generated");
			outRoot = xmlRoot("OutData");
			if (outRoot) {
				outRoot.setAttribute("Done", "Y");
			}
		}
		window.close();
		return false;
	};
}(window, document));
</script>
<script>
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.xmlIsland("OutData");
});
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">	
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">

	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Courier Details
		
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
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
								<td valign="top" width="100%">
                                <!--<div class="frmBody" id="frm2" style="width: 640; height:310;">-->
									<TABLE BORDER="0" CELLSPACING=1 CELLPADDING=0 WIDTH=100% class="ExcelTable">
										<tr>
											<td class="FieldCellSub">To Send By</td>
											<td class="FieldCell">
												<Input type="Radio" name="radSendBy" Value="C">Courier
												<Input type="Radio" name="radSendBy" Value="E">E-Mail
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Courier Company Name</td>
											<td class="FieldCell">
												<Input type="Text" name="txtCouComName" value="" class="FormElem">
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Courier TransactionID</td>
											<td class="FieldCell">
												<Input type="Text" name="txtCouTransID" value="" class="FormElem">
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Address</td>
											<td class="FieldCell" colspan="2">
												<Textarea type="Text" name="txtCouComAddress" value="" class="FormElem" cols="40"></Textarea>
											</td>
										</tr>
									</TABLE>
                                 <!--</div>-->
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
													<input type="button" value="Done" class="ActionButton" onclick="Submit()">
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
