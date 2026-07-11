<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtReqBlankChqEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January  31, 2003
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
<!--#include file="../../include/populate.asp"-->
<%
dim sOrgId,sOrgName,sAccCode,sAccName
'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("hUnitName")
sOrgId = session("organizationcode")
sOrgName = session("orgshortname")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="AccHeadData">
<account/>
</XML>
<XML id="PartyData"><Root></Root></XML>
<XML id="TempXMLData"><Root></Root></XML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/VouTransactions.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT>
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function formField(name) {
	return document.formname.elements[name] || document.formname.elements[String(name).toLowerCase()] || null;
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

function islandRoot(name) {
	return xmlRoot(window[name] || document.getElementById(name));
}

function firstEntry(root) {
	for (var i = 0; root && i < root.childNodes.length; i += 1) {
		if (root.childNodes[i].nodeType === 1 && String(root.childNodes[i].nodeName).toLowerCase() === "entry") {
			return root.childNodes[i];
		}
	}
	return null;
}

function openModernDialog(url, args, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	} else {
		window.open(url, "_blank", "height=500,width=420,resizable=no,status=no");
	}
}

function runSelectionDialog(programName, query, args, features, done) {
	openModernDialog("../../Common/" + programName + "?" + query, args, features, function (outValue) {
		var root = xmlRoot(outValue);
		var action = trim(root && root.getAttribute("Action")).toUpperCase();
		var passQuery = trim(root && root.getAttribute("PassQuery"));
		if (!root || action === "CLOSE") {
			return;
		}
		if (action !== "DONE" && passQuery !== "") {
			runSelectionDialog(programName, passQuery, args, features, done);
			return;
		}
		done(root);
	});
}

function resetAccount() {
	document.formname.hAccountCode.value = "";
	document.formname.hAccountName.value = "";
	document.formname.txtPayTo.value = "";
}

function selAccountHead(objAcc) {
	var selectedText;
	if (objAcc.selectedIndex > 0) {
		if (objAcc.selectedIndex > 1) {
			selectedText = objAcc.options[objAcc.selectedIndex].text;
			showPartyHead(document.formname.hUnitId.value, objAcc.value + "?" + selectedText);
		} else {
			showGLHead(document.formname.hUnitId.value);
		}
	} else {
		resetAccount();
	}
}

function showGLHead(sOrgId) {
	var sizeInfo = GetWindowSizeForPopup("5").split(":");
	var programName = sizeInfo[0];
	var features = "dialogHeight:" + sizeInfo[1] + "px;dialogWidth:" + sizeInfo[2] + "px;Status:No";
	var args = window.TempXMLData || document.getElementById("TempXMLData");
	runSelectionDialog(programName, "orgID=" + encodeURIComponent(sOrgId) + "&BookId=00&BookNo=", args, features, function (root) {
		var entry = firstEntry(root);
		var retVal;
		var accRoot;
		var headerNode;
		if (!entry) {
			document.formname.selAcctype.selectedIndex = 0;
			return;
		}
		retVal = [0, 1, 2, 3, 4, 5, 6].map(function (index) {
			return entry.getAttribute("RetField" + index) || "";
		}).join(":");
		if (typeof window.GetGlHeadXml === "function") {
			window.GetGlHeadXml(retVal);
		}
		accRoot = islandRoot("AccHeadData");
		for (var i = 0; accRoot && i < accRoot.childNodes.length; i += 1) {
			if (accRoot.childNodes[i].nodeType === 1) {
				headerNode = accRoot.childNodes[i];
			}
		}
		if (headerNode) {
			document.formname.hAccountCode.value = headerNode.getAttribute("No") || "";
			document.formname.hAccountName.value = (headerNode.getAttribute("Name") || "") + "&nbsp;";
			document.formname.txtPayTo.value = headerNode.getAttribute("Name") || "";
		} else {
			document.formname.selAcctype.selectedIndex = 0;
		}
	});
}

function showPartyHead(sOrgId, sPartyType) {
	var sizeInfo = GetWindowSizeForPopup("2").split(":");
	var programName = sizeInfo[0];
	var features = "dialogHeight:" + sizeInfo[1] + "px;dialogWidth:" + sizeInfo[2] + "px;Status:No";
	var args = window.PartyData || document.getElementById("PartyData");
	runSelectionDialog(programName, "orgid=" + encodeURIComponent(sOrgId) + "&Party=" + encodeURIComponent(sPartyType), args, features, function (root) {
		var entry = firstEntry(root);
		var partyName;
		var partyCode;
		if (!entry) {
			document.formname.selAcctype.selectedIndex = 0;
			return;
		}
		partyCode = entry.getAttribute("RetField1") || "";
		partyName = entry.getAttribute("RetField0") || "";
		document.formname.hAccountCode.value = sPartyType + "?" + partyCode;
		document.formname.hAccountName.value = partyName;
		document.formname.txtPayTo.value = partyName;
	});
}

function checksubmit() {
	var accountType = formField("selAcctype");
	if (document.formname.hAccountCode.value === "") {
		alert("Select Account Head");
		if (accountType) {
			accountType.focus();
		}
		return;
	}
	if (trim(document.formname.txtReason.value) === "") {
		alert("Enter Reason");
		document.formname.txtReason.select();
		return;
	}
	if (ValidateAmount(document.formname.txtAmount.value) === false) {
		document.formname.txtAmount.select();
		return;
	}
	if (document.formname.selUserId.selectedIndex === 0) {
		alert("Select Approver ");
		document.formname.selUserId.focus();
		return;
	}
	document.formname.submit();
}

function ValidateAmount(dAmount) {
	var amount = Number(String(dAmount).replace(/,/g, ""));
	if (trim(dAmount) === "") {
		alert("Amount Cannot be blank");
		return false;
	}
	if (!isFinite(amount)) {
		alert("Enter Numeric values for Amount");
		return false;
	}
	if (amount > 9999999999.99) {
		alert("Amount should be  < 9999999999.99");
		return false;
	}
	return true;
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="PmtReqBlankChqInsert.asp">
<input type="hidden" name="hUnitName" value="">
<input type="hidden" name="hUnitId" value="<%=sOrgId%>">
<input type="hidden" name="hAccountCode" value="">
<input type="hidden" name="hAccountName" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Blank Cheque Payment
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<!--<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td class="TabCell" valign="bottom" align="center" width="110">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Request Selection</td>
									</tr>
								  </table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="132">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								   <tr>
									  <td align="center">Requisition Details</td>
									</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                &nbsp;
								</td>
                            </tr>
						</table>
					</td>
				</tr>-->
				<TR>
					<TD class=TabBody>
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
													<table cellpadding="0" cellspacing="0" width="100%">
														<!--<tr>
															<td class="FieldCell">Unit</td>
															<td><span class="DataOnly"><%=sOrgName%></span></td>
														</tr>-->
														<tr>
															<td class="FieldCell" width="105">Party Type</td>
															<td class="FieldCell">
															    <select size="1" name="selAcctype" class="FormElem" onChange="selAccountHead(this)">
									   							 	<option value="S">Select Account Head</option>
									   							 	<option value="G">General Ledger</option>
															   		 <%populatePartyType(sOrgId)%>
															        </select>
															        </td>
															</tr>
															   <tr>
															<td class="FieldCell" width="105">Pay To</td>
															<td class="FieldCell"> <input type="text" name="txtPayTo" size="45" class="FormElem"></span> </td>
															    </tr>
														<tr>
															<td class=FieldCell width="130" valign="top"> Reason</td>
															<td class="FieldCell">
                                                            <textarea rows="3" name="txtReason" cols="40" class="FormElem"></textarea>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="130"> Amount&nbsp;</td>
															<td  class="FieldCell">
                                                            <input type="text" name="txtAmount" size="15"  value="0" style="text-align:right" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="130"> Immediate Approver&nbsp;</td>
															<td  class="FieldCell">
															    <select size="1" name="selUserId" class="FormElem">
																<option value="I">Immediate Approver</option>
																<%=populateEmployee%>
																    </select>
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
 <input type="button" value="Ok" name="B4" class="ActionButton" onclick="checksubmit()">
  <input type="reset" value="Reset" name="B1" class="ActionButton" >
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
