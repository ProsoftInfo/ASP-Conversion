<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgUnitCreationEntry.asp
	'Module Name				:	Inventory (Organization Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 12, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	orgUnitCreationInsert.asp
	'Procedures/Functions Used	:	populateOrganizationUnit
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
<HTML><HEAD><TITLE>Organization Unit</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../scripts/orgUnitCreate.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="OrgUnitCreateInsert.asp">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Organization Creation</p>
		</td>
    </tr>
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="95">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgCreationEntry.asp">
											<td width="100%" align="center">Organization
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="130">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Organization Units
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="190">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="OrgUnitDefinitionEntry.asp">
									  <td width="100%" align="center">Organization Units Definition</td></a>
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
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td class="PlainCell" valign="top" width="100%">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class=FieldCell> Unit description</td>
															<td class='FieldCellSub'>
																<input type="text" name="txtUnitName" size="30" maxlength=255 class="Formelem">
																<input type="button" value="Add" name="B2" class="ActionButton" tabindex="3" onClick="javascript:checkSubmit()">
																<input type="reset" value="Reset" name="B4" class="ActionButton" tabindex="4" >
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell valign="top"> Existing Units</td>
															<td class='FieldCellSub'><select size="4" name="selOrganizationUnit" class="FormElem">
															<%	'Calling the Function which populates the Organization Units list
																populateOrganizationUnit
															%>
                                                            </select>
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
