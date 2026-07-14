<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 16, 2003
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
<!--#include virtual="/include/Accpopulate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData"><Book/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" ID="CommData"><Book/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/VouTransactions.js"></script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/CreditDebitNoteBookSelectionCompat.js"></script>

<script>
ITMSCreditDebitNoteBookSelectionCompat.install({ mode: "debit" });
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="horgName" value="">
<input type="hidden" name="hTransNo" value="">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hVouDetails" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Debit Note Voucher Entry
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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="90">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Adjustments
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<tr><td align="center">Voucher</td>
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
						<table border="0" cellpadding="0" cellspacing="0">
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
                                <tr>
                            <td class="FieldCell" width="168">Organization </td>
                            <td class="FieldCell">
                             <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationListDBWithVal("")%>
                              </select>
                              </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Book</td>
                            <td class="FieldCell">
                            <select size="1" name="selBook" class="FormElem">
                        <option value="S">Select Book</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Select Party Type</td>
                            <td class="FieldCell">
                            <select size="1" name="selPartyType" class="FormElem" onChange="selParty(this)">
								<option value="S">Select Party Type</option>
								</select>
								</td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Party Name</td>
                            <td class="FieldCell" colspan="3"> <input type="text" name="txtPartyName" size="40" class="FormElem"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Voucher Type</td>
                            <td class="FieldCell">
                            <select size="1" name="selVoucherType" class="FormElem" onChange="setInvoiceNo()">
							<option value="S">Select Voucher Type</option>
							<!--<option value="PR">Purchase Returns</option>-->
							<!--option value="OT">Others</option>
							<option value="OI">Other Invoices</option-->
							<option value="OT" Selected>Others</option>
							<option value="OP">Purchase Invoices</option>
							<option value="SI">Sales Invoices</option>

                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Select Reference Number</td>
                            <td class="FieldCell">
                            <select size="1" name="selInvoiceNo" class="FormElem" onchange="SelInvChoice()">
							<option value="S">Select Invoice </option>
                            </select></td>
                            <td class="FieldCell">Or</td>
                            <td class="FieldCell">
                            <select  name="selRefNo" class="FormElem" onchange="SelRefChoice()">
							<option value="S">Select Invoice </option>
                            </select></td>

                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Debit Note Number</td>
                            <td class="FieldCell">
								<input type="text" name="txtVouchNo" size="20" class="FormElem" readonly>
								<!--
								  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<a href="#" onclick="popVoucherNo('C'); return false;">
								<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<a href="#" onclick="popVoucherNo('A'); return false;">
								<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers"></a -->
                            </td>
                                </tr>

                                    </table>
								</td>
								<td align="center" width="5">
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
                                                                <input type="button" value="Create" onClick="VouCreate()" name="btnCreate" class="ActionButton" >
                                                                <!--input type="button" value="View" onClick="VouView()" name="btnView" class="ActionButton" >
                                                                <input type="button" value="Amendment" onClick="VouAmend()" name="btnAmend" class="ActionButtonX" >
                                                                <input type="button" value="Delete" onClick="VouDel()" name="btnDelete" class="ActionButton" -->
                                                                <input type="reset" value="Reset" name="btnReset" class="ActionButton" >
                                                                <input type="button" value="New" onClick="NewPage()" name="btnNew" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" width="5">
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
