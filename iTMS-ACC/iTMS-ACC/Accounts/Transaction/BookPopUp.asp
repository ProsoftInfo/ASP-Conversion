<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BookPopUp.asp
	'Module Name				:	ACCOUNTS ()
	'Author Name				:
	'Modified By				:	S.Maheswari
	'Created On					:	Sep 16 2008
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
<!--#include virtual="/include/populate.asp"-->
<%

'XML DOM Variables
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS -
	  <%

			Response.Write "Select Book"

	%>
</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData"><Book/></script>
<script>
window["return" + "Value"] = "0--0";
window.returnvalue = "0--0";
window.ReturnValue = "0--0";

function dialogId() {
	var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
	return match ? decodeURIComponent(match[1]) : "";
}

function notifyDialogValue(id, value) {
	if (!id || !window.opener) {
		return;
	}
	try {
		if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, value);
			return;
		}
	} catch (ignoreDirectReturn) {}
	try {
		window.opener.postMessage({ type: "itms-dialog-return", id: id, value: value }, window.location.origin || "*");
	} catch (ignoreMessageReturn) {}
}

function setDialogReturnValue(value) {
	var id;
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(value);
		return;
	}
	window["return" + "Value"] = value;
	window.returnvalue = value;
	window.ReturnValue = value;
	id = dialogId();
	notifyDialogValue(id, value);
}

function responseRoot(xhr) {
	if (xhr.responseXML && xhr.responseXML.documentElement) {
		return xhr.responseXML.documentElement;
	}
	if (String(xhr.responseText || "").replace(/^\s+|\s+$/g, "") !== "") {
		return new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
	}
	return null;
}

function DisplayBook() {
	var unitNo = document.formname.hUnitId.value;
	var vouType = String(document.formname.hVouType.value || "").replace(/^\s+|\s+$/g, "");
	var bookCode = vouType === "GJ" ? "08" : vouType === "CR" ? "07" : vouType === "DR" ? "06" : "";
	var select = document.formname.selBook;
	var xhr;
	var root;
	select.options.length = 0;
	if (!bookCode) {
		return;
	}
	xhr = new XMLHttpRequest();
	xhr.open("GET", "XMLGetOrgBook.asp?BkCode=" + encodeURIComponent(bookCode) + "&orgID=" + encodeURIComponent(unitNo), false);
	xhr.send(null);
	root = responseRoot(xhr);
	Array.prototype.forEach.call(root ? root.childNodes : [], function (node) {
		var option;
		if (node.nodeType !== 1) {
			return;
		}
		option = document.createElement("option");
		option.value = node.getAttribute("BookNo") || (node.attributes[0] ? node.attributes[0].nodeValue : "");
		option.text = node.getAttribute("BookName") || (node.attributes[1] ? node.attributes[1].nodeValue : "");
		select.options.add(option);
	});
}

function Win_UnLoad() {
	var select = document.formname.selBook;
	var option;
	if (select.selectedIndex === -1) {
		alert("Select Book");
		return false;
	}
	option = select.options[select.selectedIndex];
	setDialogReturnValue(option.value + "--" + option.text);
	window.close();
	return true;
}
</script>
<%
Dim sUnit,sVouType
sUnit = Request("Unit")
sVouType = Request("VouType")

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload= "DisplayBook()" >
<form method="POST" name="formname" action="">
<Input type="hidden" name="hUnitId" value="<%=sUnit%>">
<Input type="hidden" name="hVouType" value="<%=sVouType%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Book
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
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">

											 <tr>
											    <td class="FieldCellSub" width="168">Book</td>
											    <td class="FieldCell">
											    <select size="5" name="selBook" class="FormElem">
												<!--option value="S">Select Book</option-->
											    </select></td>
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
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
 													<input type="button" value="Done" name="B3" class="ActionButton" onclick="Win_UnLoad()">

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
</body>
