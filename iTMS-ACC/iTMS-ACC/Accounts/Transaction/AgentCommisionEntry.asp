<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	SalTrOrderhdr.asp
	'Module Name				:	Sales (Transaction - Order Creation)
	'Author Name				:	SUBBIAH.S
	'Created On					:	FEB 08, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				: On click of the Next Button SalTrOrditemDet.asp
	'Procedures/Functions Used	:
	'Internal Variables			:

	'Database					:	SITMS
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim objRs,sQuery
Dim sAgentcode,sAgentname,sCustomerCode,sCustomername,sAgentType
Dim sCurcode,sCurname,sOrgID
sAgentType=Request("AgentType")
sOrgID=Request("OrgID")
Set objRs = Server.CreateObject ("ADODB.Recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" id=OutData><AgentDetails/></script>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/DataValidation.js"></SCRIPT>
<SCRIPT SRC="../../scripts/selection.js"></SCRIPT>
<Script>
function Getname(Val)
{
	 return document.formname.selTobox.options[Val].text
}
function Getpartyname()
{
	 return document.formname.cmbparty.options[document.formname.cmbparty.selectedIndex].text
}
</Script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script>
window.__itmsPopupCompat = { type: "agentCommission" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" onunload="return window_onunload()" MARGINWIDTH="0">

<form method="POST" name="formname" action="" class="PopupTable">
<input type=hidden name="hSelectedValue" value="" >
<%IF CStr(sAgentType) = "1" Then %>
	<input type=hidden name="hParType" value="CR" >
<%Else%>
	<input type=hidden name="hParType" value="DR" >
<%End IF %>
<input type=hidden name="hParSubType" value="<%=sAgentType%>" >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20">
          <p align="center">Agent Selection
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
						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td>
								</td>
								<td valign="top">
                                                          <table border="0" cellspacing="1" width="100%" class="TableOutlineOnly">
                                                    <tr>
                                                <td colspan="2" class="TableHeader" width="50%"><p align="left">Enter few characters to select&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                <input ID="FormsEditField10" NAME="txtSearch" VALUE SIZE="10" ONKEYUP="selectTheItem(this,'selFrombox')" class="formelem"></td>
                                                    </tr>
                                                    <tr>
                                                <td width="50%" class="TableHeader" align="center">Select Agents</td>
                                                <td width="50%" class="TableHeader" align="center">Selected Agents</td>
                                                    </tr>
                                                    <tr>
                                                <td width="50%" class="TableInput"><p align="center">
													<select size="5" name="selFrombox" multiple class="FormElem">
													<%

														IF Cstr(sAgentType) = "1" Then
															sQuery = "Select PartyCode,PartyName from VwOrgParty where OUDefinitionID='"&sOrgID&"'"&_
															" and PartyType='CR' and PartySubType="&sAgentType
														Else
															sQuery = "Select PartyCode,PartyName from VwOrgParty where OUDefinitionID='"&sOrgID&"'"&_
															" and PartyType='DR' and PartySubType="&sAgentType
														End IF
														objRs.Open sQuery,Con
														If Not objRs.EOF then
															Do while Not objRs.EOF
													%>
														<option value="<%=objRs(0)%>"><%=objRs(1)%></option>
													<%
															objRs.Movenext
														loop
														objRs.close
														End if
													%>
													</select>
                                                </td>
                                                <td width="50%" class="TableInput"><p align="center">
													<select size="5" name="selTobox" multiple class="FormElem">

													</select>
                                                </td>
                                                    </tr>
                                                    <tr>
                                                <td class="TableFooter" width="50%"><p align="center">
                                                <input type="button" value="Add &gt;&gt;" NAME="add" ONCLICK="Newadd()" class="AddButton" tabindex="3"></td>
                                                <td class="TableFooter" width="50%"><p align="center">
                                                <input type="button" value="&lt;&lt; Remove" NAME="remove" ONCLICK="Removeadd()" class="AddButton" tabindex="3"></td>
                                                    </tr>
                                                          </table>
								</td>
								<td>
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td>
								</td>
								<td valign="top" width="100%">
											<table border="0" id = "tblBin" cellspacing ="1" class="ExcelTable" width="100%">

                                                                      <tr>
                                                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                                                        <td class="ExcelHeaderCell" align="center">Agent Name</td>
                                                                        <td class="ExcelHeaderCell" align="center">Commission<br>
                                                                        Type</td>
                                                                        <td class="ExcelHeaderCell" align="center">Commission</td>
                                                                      </tr>

                                                                    </table>
								</td>
								<td>
								</td>
                            </tr>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td colspan="3" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Done" name="next" class="ActionButton" onclick = "FinalSubmit()" >
                                                                <input type="button" value="Cancel" name="B5" class="ActionButton" onclick = "FinalCancel()" >
																<input type="reset" value="Reset" name="B4" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td colspan="3" class="BottomPack">
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
</html>
