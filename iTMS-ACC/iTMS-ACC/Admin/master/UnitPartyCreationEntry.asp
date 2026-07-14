<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	UnitPartyCreationEntry.asp
	'Module Name				:	ADMIN (Unit - Party Creation)
	'Author Name				:	TAJUDEEN S
	'Created On					:	June 15, 2004
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
<!--#include virtual="/include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Unit Party Relation</TITLE>
<SCRIPT type="application/xml" id="XMLData" data-itms-xml-island><ROOT/></SCRIPT>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../Scripts/AdminUnitPartyCompat.js"></SCRIPT>

</HEAD>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="Init()">
	<form method="POST" name="formname" >
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Unit Party Relation
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" bordercolor="#000000">
					<tr>
						<td class="TabBodywithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center">
									</td>
									<td align="left" valign="top" >
									<DIV class=frmBody id=frm1 style="height:300;width=100%">
								        <table id="tblDetails" border="0" cellspacing="1" class="ExcelTable" width=100% >
								           <tr>
										   <td class="ExcelHeaderCell" align="left" Colspan="4">Division</td>
										   </tr>
								           <tr>
										   <td class="ExcelHeaderCell" align="center" width="10"><p align="center">S.No.</td>
										   <td class="ExcelHeaderCell" align="center">Unit Name</td>
										   <td class="ExcelHeaderCell" align="center">Party Name</td>
										   <td class="ExcelHeaderCell" align="center">Party</td>
								           </tr>
								         </table>

									</td>
									</DIV>
									<td align="center">
									</td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="BottomPack">
									</td>
								</tr>
								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="Button" value="Done"name="B2" class="ActionButton" onClick="CheckSubmit()">
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
</body>
</html>
