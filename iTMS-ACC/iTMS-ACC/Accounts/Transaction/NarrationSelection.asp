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
	'Program Name				:	NarrationSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January 04, 2003
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
<%
dim sOrgId,sBookCode,objRs,sQuery,sTemp,sBookNo
sOrgId=Request("orgid")
sTemp=split(Request("BookCode"),"?")
sBookCode=sTemp(0)
sBookNo=sTemp(1)

Set objRs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Frequently Used Narration</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT src="../../scripts/Selection.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
var sNarration = "";
var dialogCompleted = false;

function finishNarrationDialog(value) {
	sNarration = value == null ? "" : String(value);
	dialogCompleted = true;
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(sNarration);
	} else {
		window["return" + "Value"] = sNarration;
		window.returnvalue = sNarration;
	}
	window.close();
}

function CheckSelected() {
	if (document.formname.selNarration.selectedIndex < 0) {
		alert("Select a Narration");
		return;
	}
	finishNarrationDialog(document.formname.selNarration.value);
}

function finalcancel() {
	finishNarrationDialog(sNarration);
}
</script>
<SCRIPT ID=clientEventHandlersJS>
<!--
function window_onunload()
{
	if (!dialogCompleted) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(sNarration);
		} else {
			window["return" + "Value"] = sNarration;
			window.returnvalue = sNarration;
		}
	}
}

//-->

function document_onkeypress(evt)
{
	evt = evt || null;
	if (evt && evt.key === "Escape")
	{
		finalcancel();
	}
}
document.addEventListener("keydown", document_onkeypress);
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onunload="return window_onunload()">
<form method="POST" name="formname" action="CCAnalysisSelection.asp">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="popuptable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Frequently Used Narration
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
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="122">Enter Text</td>
                            <td class="FieldCell">
 <INPUT ID="FormsEditField10" TYPE=TEXT NAME="txtSearch" VALUE="" SIZE=15  ONKEYUP="selectTheItem(this,'selNarration')" class="formelem"></td>

                                </tr>
                                <tr>
                            <td class="FieldCell" width="122" valign="top">Select Narration</td>
                            <td class="FieldCell">
<%
dim iNarrHead,sNarrDescription,sShortDescription
sQuery ="select NarrationDesc,NarrationNumber,NarrationShortDesc from VwOrgFrequentNarration where "&_
	" OUDefinitionID='"&sOrgId&"'and BookCode='"&sBookCode&"' and BookNumber="&sBookNo

'Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if objRs.RecordCount>20 then
%>
 <select size="20" name="selNarration"  onDblclick="CheckSelected()" class="FormElem">
<%else%>
 <select size="<%=objRs.RecordCount%>" name="selNarration"  onDblclick="CheckSelected()" class="FormElem">
<%
end if

set sNarrDescription = objRs(0)
set iNarrHead = objRs(1)
set sShortDescription= objRs(2)
If not objRs.EOF then
	Do While Not objRs.EOF
		Response.Write("<OPTION VALUE="""&sNarrDescription&""">"&_
			""&sNarrDescription&"</OPTION>")
		objRs.MoveNext
	Loop
end if
objRs.Close
%>
 </select>
                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
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
                                                                <input type="button" value="Done" name="B7" onclick="CheckSelected()" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="B8" onClick="finalcancel()" class="ActionButton">
                                                                 <input type="reset" value="Reset" name="B9" class="ActionButton" >
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
</HTML>
