<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	RepCreationEntry.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 03,2012
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/sessionVerify.asp" -->
<%
    Dim rsTemp
    Dim sAgentName,sAgentShortName,sAgentAddress1,sAgentAddress2,sPartyCode,sQuery,sCity,sPinCode
    Dim sAgentAddress4,sAgentPhone,sAgentFax,sAgentMailID,sExternalOrInternal,sAgentType,sArrTemp
    Dim iAgentEntryID,iRepAreaCode
    
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    Response.Write "<font colo=red>"
    
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self"></base>
<xml id="RepAreaData"><Root></Root></xml>
<xml id="AgentData"><Root></Root></xml>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script SRC="../../scripts/trim.js"></SCRIPT>

<script>
window.__itmsPopupCompat = { type: "representativeCreation" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">

<form method="post" name="formname" action="">
<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
<input type="hidden" name="hAreaCode" value="<%=iRepAreaCode%>">
<input type="hidden" name="hAgentEntryID" value="<%=iAgentEntryID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="middle" class=PageTitle height="20"><p align="center">Representative Creation</p>
		</td>
    </tr>
	<tr>
		<td align="middle" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="middle" colspan="3" class="MiddlePack">
									<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
								</td>
							</tr>
							<tr>
								<td align="middle">
								</td>
								<td valign="top" width="100%">
								<table cellpadding="0" cellspacing="0">
									<tr>
										<td class=FieldCell width="110">Rep. Name</td>
										<td class='FieldCell'>
										<input name="txtAgentName" size="65" class="Formelem" style="LEFT: 0px; TOP: 2px" maxLength=50  value="<%=sAgentName%>"></td>
                     
									</tr>
									<tr>
										<td class=FieldCell width="110"> Rep. Short Name</td>
										<td class='FieldCell'>
										<input name="txtSAgentName" size="10" class="Formelem" maxLength=20 value="<%=sAgentShortName%>"></td>
									</tr>
							<!--		<tr>
                            <td class="FieldCell" width="90"> Agent Type</td>
                            <td class='FieldCell'>
                                <input type=radio name=radIntExt value="I" onclick="ChangeType()">Internal&nbsp;&nbsp;
                                <input type=radio name=radIntExt value="E" onclick="ChangeType()" checked>External&nbsp;&nbsp;
                                <select size="1" name="cmbAgentType" class="FormElem">
									<OPTION value="0">Select </option>
									<OPTION value="I">Indian Agent</option>
									<OPTION value="C">Consolidation Agent</option>
									<OPTION value="L">Clearing Agent</option>
                                </select>
                            </td>
                                </tr>-->
								</table>
								</td>
								<td align="middle">
								</td>
                            </tr>
                            <tr>
								<td align="middle" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="middle">
								</td>
								<td valign="top">
                                                <table cellpadding="0" cellspacing="0">
                                            <tr>
                                        <td>
                                        <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                <td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p>
                                </td>
                                <td class="GroupTitle" width="110" align="middle"><p align="center">Address 
                              Details</p> </td>
                                <td class="GroupTitleRight"><p align="left">&nbsp;</p></td>
                                    </tr>
                                        </table>
                                        </td>
                                            </tr>
                                            <tr>
                                        <td class="GroupTable">
                                        <table cellpadding="0" cellspacing="0">
                                    <tr>
                                <td class="MiddlePack" colspan="5"><p align="left"></p></td>
                                    </tr>
                                    <tr>
                                 
                                <td class="FieldCellSub"><p align="left">Address</p>
                                </td>
                                <td class="FieldCellSub" colspan="4"><p align="left">
                                <input name="txtAddress1" size="81" class="Formelem" maxLength=50 value="<%=sAgentAddress1%>"></p>
                                </td>
                                    </tr>
                                    <tr>
                                <td class="FieldCellSub"><p align="left"></p></td>
                                <td class="FieldCellSub" colspan="4"><p align="left">
                                <input name="txtAddress2" size="81" class="Formelem" maxLength=50 value="<%=sAgentAddress2%>"></p>
                                </td>
                                    </tr>
                                    <tr>
                                        <td class="FieldCellSub"><p align="left">City</p>
                                        </td>
                                        <td class="FieldCellSub"><p align="left">
                                        <input name="txtCity" size="25" class="Formelem" maxLength=40 value="<%=sCity%>"></p>
                                        </td>
                                        <td class="FieldCellSub"></td>
                            	        <td class=FieldCellSub> Rep. Area</td>
								        <td class='FieldCellSub'>
								            <select name="SelRepArea" class="FormElem">
								                <option value="0">Select</option>
								            </select>&nbsp;&nbsp;
								            <input type=button name=btnAddNew value="Add New" class="ActionButtonX" onclick="AddArea()">
								        </td>
									</tr>
                                    <tr>
                                <td class="FieldCellSub"><p align="left">PIN</p>
                                </td>
                                <td class="FieldCellSub"><p align="left">
                                <input name="txtPin" size="7" class="Formelem" maxLength=20 value="<%=sPinCode%>"></p>
                                </td>
                            
                                <td class="FieldCellSub"></td>
                                <td class="FieldCellSub"><p align="left">Phone</p>
                                </td>
                                <td class="FieldCellSub"><p align="left">
                                <input name="txtPhone" size="18" class="Formelem" maxLength=50 value=""></p>
                                </td>
                                    </tr>
                                
                                
                                <td class="FieldCellSub"><p align="left">Fax</p>
                                </td>
                                <td class="FieldCellSub"><p align="left">
                                <input name="txtFax" size="18" class="Formelem" maxLength=50 value="<%=sAgentFax%>"></p>
                                </td>
                                <td class="FieldCellSub"></td>
                                <td class="FieldCellSub">Mobile</td>
                                <td class="FieldCellSub">
                                <input name="txtMobile" size="18" class="Formelem" maxLength=20 value="<%=sAgentPhone%>"></td>
                                
                                    </tr>
                                    <tr>
                                <!--
                                    <td class="FieldCellSub"><p align="left">State</p>
                                </td>
                                <td class="FieldCellSub"><p align="left">
                                <input name="txtState" size="35" class="Formelem" maxLength=40 ></p>
                                </td>
                                <td class="FieldCellSub"></td>
                                <td class="FieldCellSub"><p align="left">Country</p>
                                </td>
                                <td class="FieldCellSub"><p align="left">
                                <input name="txtCountry" size="25" class="Formelem" maxLength=20>
                                </p>
                                </td>-->
                                
                                    </tr>
                                    <tr>
                                <td class="FieldCellSub"><p align="left">E-mail ID</p>
                                </td> 
                                <td class="FieldCellSub"><p align="left">
                                <input name="txtEmail" size="35" class="Formelem" maxLength=100 value="<%=sAgentMailID%>"></p>
                                </td>
                                <!--<td class="FieldCellSub"></td>
                                <td class="FieldCellSub"><p align="left">URL</p>
                                </td>
                                <td class="FieldCellSub"><p align="left">
                                <input name="txtUrl" size="25" class="Formelem" maxLength=50></p>
                                </td>-->
                                    </tr>
                                    <tr>
                                <td class="MiddlePack" colspan="5"><p align="left"></p></td>
                                    </tr>
                                        </table>
                                        </td>
                                            </tr>
                                                </table>
								</td>
								<td align="middle">
								</td>
							</tr>
							<tr>
								<td align="middle" colspan="3" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td align="middle">
									<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
								</td>
								<td valign="top">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td valign="middle" class="ActionCell">
											<p align="center">
                                                <input type="button" value="Save" onClick="checkSubmit()" name="B2" class="ActionButton" tabindex="3" > 
                                                <input type="button" value="Close" name="B3" class="ActionButton" tabindex="3" onclick="window.close()"> 
										</td>
									</tr>
								</table>
								</td>
								<td align="middle">
									<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
								</td>
							</tr>
							<tr>
								<td align="middle" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</TD>
				</TR>
			</TABLE>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
