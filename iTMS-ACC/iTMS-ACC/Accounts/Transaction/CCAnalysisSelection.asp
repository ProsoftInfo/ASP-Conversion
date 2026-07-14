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
	'Program Name				:	CCAnalysisSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 23, 2002
	'Modified On				:   January  23, 2003
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
<%
dim sOrgId,sAccCode,objRs,sQuery,sTransNo,sEntNo
sOrgId=Request("orgid")
sAccCode=Request("AccCode")
sTransNo = Request("TransNo")
sEntNo = Request("EntNo")

'Response.Write sTransNo & sEntNo
Set objRs = Server.CreateObject("ADODB.RecordSet")
Response.Write "<font color=red>"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS Cost Center-Analytical Head Selection</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/Selection.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<root No="0">
	<CostCenter/>
	<Analytical/>
</root>
</script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>

<script>
var dialogCompleted = false;

function childElements(node) {
	var result = [];
	if (!node) {
		return result;
	}
	for (var i = 0; i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1) {
			result.push(node.childNodes[i]);
		}
	}
	return result;
}

function dialogArgumentsValue() {
	var args = window.dialogArguments;
	var match;
	var id;
	if (!args && window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.dialogArgumentsRoot) {
		args = window.ITMSModalReturnCompat.dialogArgumentsRoot();
	}
	if (!args && window.opener && window.opener.__itmsDialogArgs) {
		match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		id = match ? decodeURIComponent(match[1]) : "";
		if (id && Object.prototype.hasOwnProperty.call(window.opener.__itmsDialogArgs, id)) {
			args = window.opener.__itmsDialogArgs[id];
			window.dialogArguments = args;
		}
	}
	return args && args.documentElement ? args.documentElement : args;
}

function Init() {
	var ndRoot = dialogArgumentsValue();
	var children;
	if (!ndRoot || !ndRoot.childNodes) {
		return;
	}
	children = childElements(ndRoot);
	for (var i = 0; i < children.length; i += 1) {
		if (children[i].nodeName === "CostCenter") {
			populateCC(children[i]);
		}
		if (children[i].nodeName === "Analytical") {
			populateAnal(children[i]);
		}
	}
}

function populateCC(ndRoot) {
	var children = childElements(ndRoot);
	var optionValue;
	for (var i = 0; i < children.length; i += 1) {
		for (var j = 0; j < document.formname.selCCFrombox.length; j += 1) {
			optionValue = document.formname.selCCFrombox.options[j].value.split("?")[0];
			if (optionValue === children[i].getAttribute("No")) {
				document.formname.selCCFrombox.options[j].selected = true;
			}
		}
	}
	if (children.length && typeof window.addclick === "function") {
		addclick("selCCTobox", "selCCFrombox", "remove");
	}
}

function populateAnal(ndRoot) {
	var children = childElements(ndRoot);
	var optionValue;
	for (var i = 0; i < children.length; i += 1) {
		for (var j = 0; j < document.formname.selANALFrombox.length; j += 1) {
			optionValue = document.formname.selANALFrombox.options[j].value.split("?")[0];
			if (optionValue === children[i].getAttribute("No")) {
				document.formname.selANALFrombox.options[j].selected = true;
			}
		}
	}
	if (children.length && typeof window.addclick === "function") {
		addclick("selANALTobox", "selANALFrombox", "remove");
	}
}

function clearChildElements(node) {
	var children = childElements(node);
	for (var i = 0; i < children.length; i += 1) {
		node.removeChild(children[i]);
	}
}

function returnSelection(value) {
	dialogCompleted = true;
	window.returnValue = value;
	window.returnvalue = value;
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(value);
	}
	window.close();
}

function ActionDone() {
	var root = AccHeadData.documentElement;
	var ccRoot = childElements(root)[0];
	var analRoot = childElements(root)[1];
	var bFlag = 0;
	var form = document.getElementById("frm1") || document.formname;
	var parts;
	var newElem;
	clearChildElements(ccRoot);
	clearChildElements(analRoot);
	for (var i = 0; i < form.selCCTobox.length; i += 1) {
		bFlag = 1;
		parts = form.selCCTobox.options[i].value.split("?");
		newElem = AccHeadData.createElement("CC");
		newElem.setAttribute("No", (parts[0] || "").trim());
		newElem.setAttribute("Name", form.selCCTobox.options[i].text);
		newElem.setAttribute("ShortName", (parts[1] || "").trim());
		newElem.setAttribute("Ratio", (parts[2] || "").trim());
		newElem.setAttribute("Amount", "0");
		ccRoot.appendChild(newElem);
	}
	for (var j = 0; j < form.selANALTobox.length; j += 1) {
		bFlag = 1;
		parts = form.selANALTobox.options[j].value.split("?");
		newElem = AccHeadData.createElement("Anal");
		newElem.setAttribute("No", (parts[0] || "").trim());
		newElem.setAttribute("Name", form.selANALTobox.options[j].text);
		newElem.setAttribute("ShortName", (parts[1] || "").trim());
		newElem.setAttribute("Ratio", (parts[2] || "").trim());
		newElem.setAttribute("Amount", "0");
		newElem.setAttribute("GroupCode", (parts[3] || "").trim());
		analRoot.appendChild(newElem);
	}
	root.attributes.item(0).nodeValue = bFlag;
	returnSelection(root);
}

function finalcancel() {
	returnSelection(AccHeadData.documentElement);
}
</script>

<script	ID=clientEventHandlersJS>
<!--
function window_onunload()
{
	if (!dialogCompleted) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(AccHeadData.documentElement);
		}
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()" onunload="return window_onunload()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="popuptable">
<form name="formname" id="frm1">
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
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel" height="2">
								</td>
								<td valign="top" align="center" width="100%">

                                    <table border="0" cellspacing="1" width="100%" class="TableOutlineOnly">
                                <tr>
									<td colspan="2" class="TableHeader" width="50%"><p align="center"> Enter few characters to select&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT ID="FormsEditField10" TYPE=TEXT NAME="txtSearch" VALUE="" SIZE=10  ONKEYUP="selectTheItem(this,'selCCFrombox')" class="formelem"></td>
                                </tr>
                                <tr>
									<td width="50%" class="TableHeader" align="center">Select Cost Center</td>
									<td width="50%" class="TableHeader" align="center">Selected Cost Center</td>
                                </tr>
                                <tr>
									<td width="50%" class="TableInput"><p align="center">
									<select size="5" name="selCCFrombox" multiple class="FormElem">
<%
dim iCCHead,sCCCode,sCCDescription,dRatio,iSno

'Selection of all Cost Center codes related to Account head and Unit
'sQuery ="select CostCenterHead,CCHeadCode,CCAccountDescription from VwOrgCostCenter where CostCenterHead not in "&_
'		"(select distinct(CostCenterHead) from VwOrgGLCostCenter where OUDefinitionID='"&sOrgId&"'"&_
'		"  and AccountHead="&sAccCode&") and OUDefinitionID='"&sOrgId&"'"

'************** Selection of Only Related CChead to Accounthead and Unit
sQuery ="select CostCenterHead,CCHeadCode,CCAccountDescription,AllocationRatio "&_
	"from VwOrgGLCostCenter where OUDefinitionID='"&sOrgId&"' and AccountHead="&sAccCode

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

set iCCHead = objRs(0)
set sCCCode = objRs(1)
set sCCDescription = objRs(2)
dRatio= 0

iSno=1
If not objRs.EOF then
	Do While Not objRs.EOF
				Response.Write "<option value="""& iCCHead&"?"&sCCCode&"?"&dRatio & """>" & sCCDescription   &"</option>"
	objRs.MoveNext
	loop
end if
objRs.Close
%>
									</select>
									</td>
                                    <td width="50%" class="TableInput"><p align="center">
										 <select size="5" name="selCCTobox" multiple class="FormElem">

										 </select>
									</td>
                                        </tr>
                                        <tr>

                                    <td class="TableFooter" width="50%"><p align="center"><input type="button" value="Add >>" NAME="add" ONCLICK="addclick('selCCTobox','selCCFrombox','remove')" class="AddButton"  >
									 </td>
                                    <td class="TableFooter" width="50%"><p align="center"><input type="button" value="<< Remove" NAME="remove" ONCLICK="removeclick('selCCTobox','selCCFrombox','remove')" class="AddButton" >
                                    </td>
                                    </tr>
                                        </table>

								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
								</td>
                            </tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
                                    <table border="0" cellspacing="1" width="100%" class="TableOutlineOnly">
                                <tr>
									<td colspan="2" class="TableHeader" width="50%"><p align="center"> Enter few characters to select&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT ID="FormsEditField11" TYPE=TEXT NAME="txtSearch" VALUE="" SIZE=10  ONKEYUP="selectTheItem(this,'selANALFrombox')" class="formelem"></td>
                                </tr>
                                <tr>
									<td width="50%" class="TableHeader" align="center">Select Analytical Head</td>
									<td width="50%" class="TableHeader" align="center">Selected Analytical Head</td>
                                </tr>
                                <tr>
									<td width="50%" class="TableInput"><p align="center">
									<select size="5" name="selANALFrombox" multiple class="FormElem">
<%
dim iAnalHead,sAnalCode,sAnalDescription
DIM sAnalGroupCode,sAnalGroupName

'********* Selected all Anal code for Selected Accounthead and Unit
'sQuery ="select AnalyticalCode,AnalyticalShortName,AnalyticalName,AHGroupCode,AHGroupName "&_
'	"from VwOrgAnalytical where OUDefinitionID='"&sOrgId&"' and "&_
'	"ltrim(str(AnalyticalCode))+ltrim(AHGroupCode)not in  (select distinct(ltrim(str(AnalyticalCode))+ltrim(AHGroupCode)) from VwOrgGLAnalytical "&_
'	" where AccountHead="&sAccCode &" and  OUDefinitionID='"&sOrgId&"')"&_
'	" ORDER BY  AnalyticalName,AHGroupName"

'********* Selects only the Anal code that are related only to Accounthead and Unit
sQuery = "select Distinct AnalyticalCode,AnalyticalShortName,AnalyticalName,AllocationRatio,AHGroupCode,AHGroupName "&_
		 "from VwOrgGLAnalytical where OUDefinitionID='"&sOrgId&"' and AccountHead="&sAccCode &" ORDER BY "&_
		 " AnalyticalName,AHGroupName"


with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

set iAnalHead = objRs(0)
set sAnalCode = objRs(1)
set sAnalDescription = objRs(2)
dRatio= 0

iSno=1
set iAnalHead = objRs(0)
set sAnalCode = objRs(1)
set sAnalDescription = objRs(2)

set sAnalGroupCode= objRs(4)
set sAnalGroupName= objRs(5)

iSno=1
dRatio=0

If not objRs.EOF then
	Do While Not objRs.EOF
		Response.Write "<option value="""& iAnalHead&"?"&sAnalCode&"?"&dRatio&"?"&sAnalGroupCode & """>" & sAnalDescription &":"& sAnalGroupName &"</option>"
		objRs.MoveNext
	loop
end if
objRs.Close
%>
									</select>
									</td>
                                    <td width="50%" class="TableInput"><p align="center">
										 <select size="5" name="selANALTobox" multiple class="FormElem">

										 </select>
									</td>
                                        </tr>
                                        <tr>

                                    <td class="TableFooter" width="50%"><p align="center"><input type="button" value="Add >>" NAME="add" ONCLICK="addclick('selANALTobox','selANALFrombox','remove')" class="AddButton" tabindex="3" >
									 </td>
                                    <td class="TableFooter" width="50%"><p align="center"><input type="button" value="<< Remove" NAME="remove" ONCLICK="removeclick('selANALTobox','selANALFrombox','remove')" class="AddButton" tabindex="3" >
                                    </td>
                                    </tr>
                                        </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
 <input type="button" value="Done" name="next" name="B8" class="ActionButton" onClick="ActionDone()" >
  <input type="button" value="Cancel" name="B8" class="ActionButton" onClick="finalcancel()" >

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
	</form>
</table>

</BODY>
</HTML>
