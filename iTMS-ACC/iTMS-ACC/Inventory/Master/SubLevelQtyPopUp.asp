<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	SubLevelQtyPopUp.asp
	'Module Name				:	INVENTORY (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 21, 2011
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
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
Dim objRs,objRs1,objRs2,iSno
dim sOrgId,sQuery

set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="SubLevelQty" data-itms-xml-island="1"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot Sub Level Entry
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
                                &nbsp;
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly" width="100%">
                                <tr>
                                    <td class="MiddlePack" colspan="3"></td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">
										<Input type="text" Name="txtSubLevelNo" class="FormElem" value="50" size="4" >&nbsp;
									Cones of </td>
									<td class="FieldCellSub">
										<Input type="text" Name="txtSubLevelQty" class="FormElem" value="1.25" size="6">&nbsp;KGS&nbsp;each
										<Input type="Button" name="btnAdd" class="ActionButton" value="Add" onclick="">
									</td>
                                </tr>

                                <tr>
                                     <td class="MiddlePack" colspan="3"></td>
                                </tr>

                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>

                            <tr>
								<td></td>
								<td Width="100%">
									<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center">Serial</td>
												<td class="ExcelHeaderCell" align="center">
													<a href="#"><img style="cursor:hand" border="0" width="11" height="11" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Delete" ></a>
												</td>
												<td class="ExcelHeaderCell" align="center">Sub Level Qty</td>
											</tr>

											<tr>
												<td class=ExcelSerial align=center>1</td>
												<td class=ExcelDisplaycell align=center>
													<Input type="Checkbox" name="chkbox" value="">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerQty" value="" class="FormElem">
												</td>
											</tr>
									</Table>
								</td>
                            </tr>



								</td>
								<td align="center" class="ClearPixel" width="5">
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
													<input type="button" value="Done" name="B3" class="ActionButton" onclick="Save()" >
                                                    <input type="button" value="Close" name="B2" class="ActionButton" onclick="window.close()" >
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
</html>
