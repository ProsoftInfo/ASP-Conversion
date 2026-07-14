<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ParDelSelectType.asp
	'Module Name				:	Accounts (Master - Party Delete Selection)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Mar 17, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:
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
<!--#include virtual="/include/Salpopulate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
	Dim sType,sAgcode,Temparr,Objrs,sQuery,sAgType

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Agent/Party Selection</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<script>
window.__itmsPopupCompat = { type: "partyDeleteSelection" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<Form method="POST" name="formname" action=""  class="PopupTable">

<table border="0" cellspacing="0" cellpadding="0" width="97%">
	<tr>
		<td align="center" class=PageTitle height="20">
          <p align="center">Party Delete Selection
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
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td>
								</td>
								<td valign="top" width="100%">

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
								        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
								          <tr>
								            <!--td class="ExcelHeaderCell" align="center" width="10">S.No</td-->
								            <td class="ExcelHeaderCell" align="center" width="10">Sel</td>
								            <td class="ExcelHeaderCell" align="center">Selection Type</td>

								          </tr>
								          <tr>
								          <!--td class="ExcelHeaderCell" align="center">1</td-->
								          <td class="ExcelHeaderCell" align="center">
											<input type="radio" class="Formelem" value="A" name="optAgsel" >
								          </td>
								          <td class="ExcelDisplayCell" align="left">Deletion From All Units</td>
								          </tr>
								          <tr>
								          <!--td class="ExcelHeaderCell" align="center">1</td-->
								          <td class="ExcelHeaderCell" align="center">
											<input type="radio" class="Formelem" value="S" name="optAgsel" checked >
								          </td>
								          <td class="ExcelDisplayCell" align="left">Deletion From Specific Unit</td>
								          </tr>
								          <tr>
								          <!--td class="ExcelHeaderCell" align="center">1</td-->
								          <td class="ExcelHeaderCell" align="center">&nbsp;</td>
										  <td class="ExcelDisplayCell" align="left">
										  <select size="5" name="selUnitId" class="FormElem" multiple>

											<%populateOrganizationListDB%>
										  </select>

										  </td>
								          </tr>

								        </table>
									</td>
									<td>
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
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onClick="SelType()">
                                                                <input type="button" value="Cancel" name="B4" class="ActionButton" onclick="window.close()">

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
