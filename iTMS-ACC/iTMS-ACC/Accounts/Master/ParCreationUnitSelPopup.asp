<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ParCreationUnitSelPopup.asp
	'Module Name				:	Accounts (Party Creation)
	'Author Name				:	TAJUDEEN S
	'Created On					:	october 29, 2004
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
<!--#include file="../../include/Populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Unit Selection</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "unitSelection" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="" >
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20">
          <p align="center">Unit Selection
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
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
								<td></td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Select Unit</td>
											<td class="FieldCellSub">
												<Select name="selUnit" size=10 Class="FormElem">
												<%
													populateUnit()
												%>
												</Select>
											</td>
										</tr>
                                    </table>
								</td>
								<td></td>
							</tr>
							<tr>
								<td colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell" align="center">
                                                <input type="button" value="Done" name="B2" class="ActionButton" onClick="CheckSubmit()">
                                                <input type="button" value="Close" name="B2" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td colspan="3" class="BottomPack"></td>
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
