<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParAgentSelectPopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 15,2010
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
<%
	Dim iPartyCode,rsObj,sQuery,sPartyName

	set rsObj = Server.CreateObject("ADODB.Recordset")

	iPartyCode = Request.QueryString("PartyCode")

	sQuery = "Select PartyCode,PartyName from APP_M_PartyMaster where PartyCode="&iPartyCode
	rsObj.Open sQuery,con
	if not rsObj.EOF then
		sPartyName = rsObj(1)
	end if
	rsObj.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self">
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<xml id="OutData"></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript src="../../scripts/Selection.js"></SCRIPT>
<SCRIPT language="vbscript">

'**************************************************
Function DisplayAgent()
dim iUnitNo,iPartyId,iAgentType
dim Root

if document.formname.selUnitId.selectedIndex = 0 then
	alert("Select Unit")
	document.formname.selUnitId.focus
	exit function
end if

iUnitNo= document.formname.selUnitId(document.formname.selUnitId.selectedIndex).value
iPartyId=document.formname.hPartyCode.value
iAgentType=document.formname.selAgentType.value

if document.formname.hPartyCode.value <> "" and document.formname.selUnitId.selectedIndex >0 and  document.formname.selAgentType.selectedIndex >0 then
	document.formname.selFrombox.length =0
	document.formname.selTobox.length =0
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLPartyAgent.asp?Flag=A&AgentType="&iAgentType&"&orgID=" & iUnitNo&"&PartyCode="&iPartyId, false
	objhttp.send

	'alert objhttp.responseText

	if objhttp.responseXML.xml <> "" then
		document.formname.selFrombox.options.length = 0
		document.formname.selTobox.options.length = 0
		OutData.loadXML objhttp.responseXML.xml

		Set Root = OutData.documentElement
			For Each HeaderNode In Root.childNodes
				if HeaderNode.Attributes.getNamedItem("Selected").Value="N" then
					document.formname.selFrombox.length = document.formname.selFrombox.length+1
					document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.text
					document.formname.selFrombox.options(document.formname.selFrombox.length-1).Value = HeaderNode.Attributes.getNamedItem("AgentId").value
				else
					document.formname.selTobox.length = document.formname.selTobox.length+1
					document.formname.selTobox.options(document.formname.selTobox.length-1).text = HeaderNode.text
					document.formname.selTobox.options(document.formname.selTobox.length-1).Value = HeaderNode.Attributes.getNamedItem("AgentId").value
				end if
			next
	else
		alert(objhttp.responseText)
	end if
else
	document.formname.selFrombox.length=0
	document.formname.selTobox.length=0
end if
end Function
'***************************
Function window_onunload()
	window.returnValue="Close"
End Function
</script>
<script language="javascript">
window.__itmsPopupCompat = { type: "agentSelect" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
<script>
function final()
{
		if (document.formname.selUnitId.selectedIndex >0)
		{
			//if (document.formname.selParty.selectedIndex >0)
			//{
				if (document.formname.selAgentType.selectedIndex >0)
				{
					document.formname.next.disabled = true
					finaldone('selTobox','hSelectedValue');
				}
				else
				{
					alert("Select an Agent Type");
					return false;
				}
			//}
			//else
			//{
			//	alert("Select a Party");
			//	return false;
			//}
		}
		else
		{
			alert("Select a Unit");
			return false;
		}

}
function actionreset()
{
	document.formname.reset();
	document.formname.selFrombox.length=0;
	document.formname.selTobox.length=0;
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="ParAgentPopupUpdate.asp">
<input type=hidden name="hSelectedValue" value="">
<input type=hidden name="hPartyCode" value="<%=iPartyCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Party
          Agent Allocation
		</td>
    </tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="90"> Select Unit</td>
                            <td>
                                            <select size="1" name="selUnitId" class="FormElem">
												<OPTION value="0">Select a Unit</option>
													<%populateOrganizationList%>
                                            </select>
                            </td>
                                </tr>

                                <tr>
                            <td class="FieldCell" width="90" valign="top">Party Name</td>
                            <td>
                            <table>
                            <tr><td>
                            <INPUT ID="FormsEditField10" TYPE=TEXT NAME="txtPartyName" VALUE="<%=sPartyName%>" SIZE=59  class="formelemread" readonly>
                            </td></tr>
							</table>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="120"> Select Agent Type</td>
                            <td>
                            <table>
                            <tr><td >
                                <select size="1" name="selAgentType" class="FormElem" >
                                <OPTION value="0">Select Agent Type</option>
                                <OPTION value="CR">Commision Agent</option>
								<OPTION value="DR">Depo Agent </option>
                                </select>
                            </td>
                            <td ><input type="button" name="show" onClick="DisplayAgent()" Value="Show" class="ActionButtonX"> </td>
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
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
                                    <table border="0" cellspacing="1" width="100%" class="TableOutlineOnly">
                                <tr>
									<td colspan="2" class="TableHeader" width="50%"><p align="center"> Enter few characters to select&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT ID="FormsEditField10" TYPE=TEXT NAME="txtSearch" VALUE="" SIZE=10  ONKEYUP="selectTheItem(this,'selFrombox')" class="formelem"></td>
                                </tr>
                                <tr>
									<td width="50%" class="TableHeader" align="center">Select
                                      Agent</td>
									<td width="50%" class="TableHeader" align="center">Selected
                                      Agent</td>
                                </tr>
                                <tr>
									<td width="50%" class="TableInput"><p align="center">
									<select size="5" name="selFrombox" multiple class="FormElem">

									</select>
									</td>
                                    <td width="50%" class="TableInput"><p align="center">
										 <select size="5" name="selTobox" multiple class="FormElem">
										 </select>
									</td>
                                        </tr>
                                        <tr>

                                    <td class="TableFooter" width="50%"><p align="center"><input type="button" value="Add >>" NAME="add" ONCLICK="addclick('selTobox','selFrombox','remove')" class="AddButton" tabindex="3" >
									 </td>
                                    <td class="TableFooter" width="50%"><p align="center"><input type="button" value="<< Remove" NAME="remove" ONCLICK="removeclick('selTobox','selFrombox','remove')" class="AddButton" tabindex="3" >
                                    </td>
                                    </tr>
                                        </table>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
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
                                                                <input type="Button" value="Save" name="next" class="ActionButton" onClick="final()" >
																<input type="Button" value="Reset" name="B1" onClick="actionreset()" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
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
