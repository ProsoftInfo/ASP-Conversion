<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	comAccountGroupTreePopup.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 16,2010
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
<!-- #include File="../../include/GetSettings.asp" -->
<!--#include file="../../include/sessionVerify.asp"-->
<%
	dim sIP
	sIP = GetSettings("IP")
%>
<HTML>
<HEAD>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script type="text/javascript" src="../../scripts/itms-modern-compat.js"></script>
<base target="_self">
<TITLE>Account Heads Tree</TITLE>
<XML id="GLHierarchyData"><Root/></XML>
<script type="text/javascript">
var sGroupCode = "0";
var glHierarchyDocument = null;

function setDisplay(id, visible) {
	var element = document.getElementById(id);
	if (element) {
		element.style.display = visible ? "block" : "none";
	}
}

function SaveGroup() {
	document.formname.action = "AccGroupNamePopupInsert.asp";
	document.formname.submit();
}

function EditGroup() {
	var form = document.formname;
	if (String(form.txtClassEditName.value || "").trim() === "") {
		alert("Enter the Group Name");
		form.txtClassEditName.focus();
		return;
	}
	form.action = "GroupEditPopupUpdate.asp";
	form.submit();
}

function HierarchyGroup() {
	var form = document.formname;
	var root = glHierarchyDocument ? glHierarchyDocument.documentElement : null;
	var groupCount = Number(root ? root.getAttribute("Counter") : form.hGroupCount.value) || 0;
	var order = [];
	for (var i = 0; i < form.elements.length; i += 1) {
		if (form.elements[i].name === "txtOrder") {
			order.push(String(form.elements[i].value));
		}
	}
	for (var counter = 1; counter <= groupCount; counter += 1) {
		if (order.indexOf(String(counter)) === -1) {
			alert("Incorrect Group Order Sequence\nOrder '" + counter + "' is Missing");
			return;
		}
	}
	form.action = "AccGroupOrderUpdatePopup.asp";
	form.submit();
}

function resetActionButtons() {
	var form = document.formname;
	form.btnEdit.disabled = false;
	form.btnCreate.disabled = false;
	form.btnHierarchy.disabled = false;
}

function FinalDone(bflag) {
	var form = document.formname;
	var tree = form.ctlGroupHeadList;

	if (bflag === "C") {
		form.GCode.value = "0";
		setDisplay("divAccGroupTree", true);
		setDisplay("divDone", true);
		setDisplay("divCreateGroup", false);
		setDisplay("divEditGroup", false);
		setDisplay("divHierarchyGroup", false);
		resetActionButtons();
		return;
	}

	form.GCode.value = tree ? tree.GroupValue : "0";
	form.GName.value = tree ? tree.GroupName : "";

	if (String(form.GCode.value) === "0") {
		alert("Select Group ");
		return;
	}

	if (bflag === "N") {
		setDisplay("divAccGroupTree", false);
		setDisplay("divDone", false);
		setDisplay("divEditGroup", false);
		setDisplay("divCreateGroup", true);
		setDisplay("divHierarchyGroup", false);
		form.btnEdit.disabled = true;
		form.btnHierarchy.disabled = true;
	} else if (bflag === "E") {
		form.txtClassEditName.value = form.GName.value;
		setDisplay("divAccGroupTree", false);
		setDisplay("divDone", false);
		setDisplay("divCreateGroup", false);
		setDisplay("divEditGroup", true);
		setDisplay("divHierarchyGroup", false);
		form.btnCreate.disabled = true;
		form.btnHierarchy.disabled = true;
	} else if (bflag === "H") {
		setDisplay("divAccGroupTree", false);
		setDisplay("divDone", false);
		setDisplay("divCreateGroup", false);
		setDisplay("divEditGroup", false);
		setDisplay("divHierarchyGroup", true);
		form.btnCreate.disabled = true;
		form.btnEdit.disabled = true;
		form.hParentGroupCode.value = form.GCode.value;
		PopulateTable(form.GCode.value);
	} else {
		window.returnValue = form.GCode.value + ":" + form.GName.value;
		window.returnvalue = window.returnValue;
		window.close();
	}
}

function ClosePopup() {
	window.returnValue = "";
	window.returnvalue = "";
	window.close();
}

function PopulateTable(sGCode) {
	fetch("GLGetHierarchyDetails.asp?GCode=" + encodeURIComponent(sGCode), { cache: "no-cache", credentials: "same-origin" })
		.then(function (response) {
			if (!response.ok) {
				throw new Error(response.status + " " + response.statusText);
			}
			return response.text();
		})
		.then(function (text) {
			var xmlDoc = new DOMParser().parseFromString(text, "text/xml");
			if (xmlDoc.getElementsByTagName("parsererror").length) {
				throw new Error("Invalid XML");
			}
			glHierarchyDocument = xmlDoc;
			renderHierarchyTable(xmlDoc);
		})
		.catch(function (error) {
			alert("Unable to load hierarchy details: " + error.message);
		});
}

function renderHierarchyTable(xmlDoc) {
	var form = document.formname;
	var root = xmlDoc.documentElement;
	var nodes = root ? root.children : [];
	var table = document.getElementById("tblHierarchy");
	ClearTable();
	if (!root || !table) {
		return;
	}
	form.hGroupCode.value = root.getAttribute("Group") || "";
	form.hGroupCount.value = root.getAttribute("Counter") || "";
	for (var i = 0; i < nodes.length; i += 1) {
		var row = table.insertRow(table.rows.length);
		var snoCell = row.insertCell();
		var nameCell = row.insertCell();
		var hierarchyCell = row.insertCell();
		var orderCell = row.insertCell();
		var orderInput = document.createElement("input");

		snoCell.innerText = nodes[i].getAttribute("SNo") || "";
		snoCell.className = "ExcelDisplayCell";
		nameCell.innerText = nodes[i].getAttribute("GName") || "";
		nameCell.className = "ExcelDisplayCell";
		hierarchyCell.innerText = nodes[i].getAttribute("Hierarchy") || "";
		hierarchyCell.className = "ExcelDisplayCell";
		hierarchyCell.align = "Center";

		orderInput.type = "text";
		orderInput.name = "txtOrder";
		orderInput.className = "FormElem";
		orderInput.value = nodes[i].getAttribute("Hierarchy") || "";
		orderInput.size = 5;
		orderInput.align = "center";
		orderCell.appendChild(orderInput);
		orderCell.className = "ExcelInputCell";
	}
}

function ClearTable() {
	var table = document.getElementById("tblHierarchy");
	if (!table) {
		return;
	}
	while (table.rows.length > 1) {
		table.deleteRow(1);
	}
}
</script>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname">
<input type="hidden" name="GName" value="">
<input type="hidden" name="GCode" value="">
<input type="hidden" name="GroupFlag" value="G">
<input type="hidden" name="hParentGroupCode" value="">
<input type="hidden" name="hGroupCode" value="">
<input type="hidden" name="hGroupCount" value="">
<table border="0" cellspacing="0" cellpadding="0" width="100%">
<tr><td height="1px"></td></tr>
<tr>
	<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td class="PageTitle">GL Group</td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td colspan=2>
	<table class="ExcelTable" width="100%" border=0 cellspacing="0" cellpadding="0">
		<tr>
			<td colspan=2>
				<div>
					<table  border=0 cellspacing="0" cellpadding="0" width=100%>
						<tr>
							<td valign="center" class="CollapseBand" height="25px" align="center">
								<input type="button" value="Create" class="ActionButton" onClick="FinalDone('N')" name="btnCreate">
				                 &nbsp;
				                 <input type="button" value="Edit" class="ActionButton" onClick="FinalDone('E')" name="btnEdit">&nbsp;
				                 <input type="button" value="Hierarchy" class="ActionButton" onClick="FinalDone('H')" name="btnHierarchy">
				    		</td>
						</tr>
					</table>
					<table>
							<tr>
								<td>&nbsp;</td>
								<td>
								<div id="divAccGroupTree" style="width:100%">
									<table>
										<tr>
										<td width="10px">&nbsp;</td>
										<td>
										<OBJECT ID="ctlGroupHeadList" CLASSID="CLSID:355CEAFA-CB06-4345-8384-D0725C8A3048" CODEBASE="../components/GroupHeadList.CAB#version=1,0,0,0" VIEWASTEXT width="263px" height="344px">
										    <param name="_ExtentX" value="6959">
										    <param name="_ExtentY" value="9102">
										    <param name="DSN" value="http://<%=sIP%>/Accounts/components/GetACCGroup.asp">
										    <param name="BackColor" value="9987378">
										    <param name="ForeColor" value="9987378">
										    <param name="GroupValue" value="0">
										    <param name="ListName" value="ACCOUNTS GROUPS">
										    <param name="HeadValue" value="0">
										    <param name="GroupName" value="0">
										    <param name="HeadName" value="0">
										</OBJECT>
										</td>
										<td>
										</td>
										</tr>
										<tr>
										<td width="10px" colspan="3" height="20px"></td>
										</tr>
									</table>
								</div>

								<div id="divCreateGroup"  class="frmbody" style="display:none">
									<table >
										<tr>
										<td align=center class="PageTitle" colspan="3"></td>
										</tr>
										<tr>
										<td width=10 class="middlepack" colspan="3"></td>
										</tr>
										<tr>
										<tr>
											<td class=FieldCellSub width="100px"> Group Name</td>
											<td class="FieldCellSub">
												<input type="text" name="txtClassCreateName" size="30" maxlength=40 class="FormElem">
											</td>
										</tr>
										<tr>
											<td class="ActionCell" colspan="2">
									        <input type="Button" value="Save" name="B2" onClick="SaveGroup()" class="ActionButtonX">&nbsp;
									        <input type="button" value="Cancel" class="ActionButton" onClick="FinalDone('C')" name="btnCreateCancel"></td>
										</tr>
										<tr>
										<td width=10 class="middlepack" colspan="3"></td>
										</tr>
										<tr>
									</table>
								</div>

								<div id="divEditGroup" class="frmbody" style="display:none">
									<table>
										<tr>
										<td class="PageTitle" colspan="3"></td>
										</tr>
										<tr>
										<td width="10px" class="middlepack" colspan="3"></td>
										</tr>
										<tr>

										<tr>
										<td class=FieldCellSub width="100px"> Group Name</td>
										<td class="FieldCellSub">
											<input type="text" name="txtClassEditName" size="30" maxlength=40 class="Formelem">
										</td>
										</tr>
										<tr>
											<td class="ActionCell" colspan="3">
									        <input type="Button" value="Save" name="B2" onClick="EditGroup()" class="ActionButtonX">
									        <input type="button" value="Cancel" class="ActionButton" onClick="FinalDone('C')" name="btnEditCancel">
									        </td>
										</tr>
										<tr>
										<td width=10 class="middlepack" colspan="3"></td>
										</tr>
										<tr>
									</table>
								</div>


								<div id="divHierarchyGroup" class="frmbody" style="display:none">
									<table>
										<tr>
											<td width="100%">
												<table id="tblHierarchy" cellspacing="1px" cellpadding="0" border="0" class="ExcelTable" width="100%">
													<tr>
														<td class="ExcelHeaderCell">S.No</td>
														<td class="ExcelHeaderCell">Group Name</td>
														<td class="ExcelHeaderCell">Hierarchy</td>
														<td class="ExcelHeaderCell">&nbsp;</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
										<td width="10px" class="middlepack"></td>
										</tr>
										<tr>
											<td class="ActionCell">
									        <input type="Button" value="Save" name="B2" onClick="HierarchyGroup()" class="ActionButtonX">
									        <input type="button" value="Cancel" class="ActionButton" onClick="FinalDone('C')" name="btnEditCancel">
									        </td>
										</tr>
										<tr>
										<td width="10px" class="middlepack"></td>
										</tr>
									</table>
								</div>


							</td>
						</tr>
					</table>
					<div id="divDone" class="frmbody">
						<table width=100%>
							<tr>
								<td class="ActionCell" colspan="3">
								     <input type="button" value="Done" class="ActionButton" onClick="FinalDone('D')" name="btn1">
								     <input type="button" value="Close" class="ActionButton" onClick="ClosePopup()" name="btn5">
								</td>
							</tr>
						</table>
					</div>
				</div>
			</td>
		</tr>
	</table>
</td>
</tr>
</table>
</form>
</BODY>
</HTML>

