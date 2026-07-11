<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BankInsDetails.asp
	'Module Name				:	Fixed Deposit(Transaction)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Jun 15,2005
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include File="../../include/DatabaseConnection.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!--#include File="../../include/IncludeDatePicker.asp" -->
<%
	Dim sQry,rs
	Set rs = Server.CreateObject("ADODB.Recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Party Sub Type</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script>
window.__itmsPopupCompat = { type: "partySubtypeDialog" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">

<form method="POST" name="formname" action="">
<Input Type="hidden" name="hExists" Value="">
<Input Type="hidden" name="hEditNo" Value="">

<input type="hidden" name="hCtr" value="1">
<input type="hidden" name="hInsType" value="C">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Party Sub Type Details
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
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">
														<Tr>
														    <td width="100%">
														        <div id="divCRDRSubType" style="display:block" >
														            <Table border="0" cellspacing="0" cellpadding="0" width="100%">
														            <Tr>
														                <Td width="1%"></td>
														                <TD width="48%"  style="height: 35px" valign="top">
														                        <Table  id="CRTable" border="0" cellspacing="1" cellpadding="0" class="ExcelTable" width="100%">
														                            <tr>
														                                <td class="ExcelHeaderCell" colspan=3 align=center>Creditor Party SubType</td>
														                            </tr>
															                        <Tr>
																                        <td class="ExcelHeaderCell" width="10"></td>
																                        <td class="ExcelHeaderCell" align=center>Sub Type</td>
																                        <td class="ExcelHeaderCell" align=center>Status</td>
															                        </tr>
														                        </Table>
														                    
														                </Td>
														                <Td  width="2%" style="height: 35px"></td>
														                <TD  width="48%" style="height: 35px" valign="Top">
																			
														                        <Table id="DRTable" border="0" cellspacing="1" cellpadding="0" class="ExcelTable" width="100%">
														                            <tr>
														                                <td class="ExcelHeaderCell" colspan=3 align=center>Debtor Party SubType</td>
														                            </tr>
															                        <Tr>
																                        <td class="ExcelHeaderCell" width="10"></td>
																                        <td class="ExcelHeaderCell" align=center>Sub Type</td>
																                        <td class="ExcelHeaderCell" align=center>Status</td>
															                        </tr>
														                        </Table>
														                    
														                </Td>
														                <td  width="1%"></td>
														                </tr>
														            </table>
														        </div>
														    </td>
														</Tr>
                                                        
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
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
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
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
</HTML>
