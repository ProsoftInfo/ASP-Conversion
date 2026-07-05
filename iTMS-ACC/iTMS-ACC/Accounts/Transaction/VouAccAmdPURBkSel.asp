<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouAccAmdPURBkSel.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	March 01, 2003
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
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
dim objRs,sQuery,sFinPeriod,sFinFun,iTransNo,sRetVal,oDOM,sFormVal

sFinPeriod = Session("FinPeriod")
sFinFun = GetfinYear(sFinPeriod)
If trim(sFinFun) = "True" Then
	Response.Redirect ("../../welcome_Welcome.asp?sFinFun="&sFinFun&"")
End If



iTransNo = Request("TransNo")
sFormVal = Request.Form("hFormVal")
'Response.Write sFormVal
Session("AmdPur") = sFormVal

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")



IF CStr(iTransNo) <> "" Then
	sRetVal = GetVouchXML(iTransNo)
	oDOM.Load server.MapPath(sRetVal)
	oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")
End IF

oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="VoucherData"><Voucher/></XML>
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<XML id="OldVouData" src="<%="../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"%>"></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PurchaseAmdBookSelectionCompat.js"></SCRIPT>

</HEAD>
	<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="DispOldVal()">

<form method="POST" name="formname" action="VouPURAccAmdDetails.asp?AmdType=A">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hBkAccHead" value="">
<input type="hidden" name="hInvDate" value="">
<input type="hidden" name="hVouDate" value="">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Accounted Purchase Voucher Amendment </td>
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
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Invoice Details</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Advance</td>
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
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="108">Organization </td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationList%>
                              </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Purchase Book</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selBook" class="FormElem">
                        <option value="S">Select Book</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Purchase Type&nbsp;</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selPurType" class="FormElem" >
									<option value="0">Select Purchase Type</option>
									<%
										dim sCode,sValue
										sQuery = "Select PurchaseType,PurchaseTypeName from APP_M_PurchaseTypes Where Active = 'Y' "
									  	With objRs
									  		.CursorLocation = 3
									  		.CursorType = 3
									  		.Source = sQuery
									  		.ActiveConnection = con
									  		.Open
									  	End with
									  	Set objRs.Activeconnection = nothing
									  	Set sCode = objRs(0)
									  	Set sValue = objRs(1)
									  	Do while not objRs.EOF
									%>
											<option value="<%Response.Write sCode%>"><%Response.Write sValue%></option>
									<%

											objRs.MoveNext
										Loop
										objRs.Close
								    %>
								</select>
                            </td>
                                </tr>

                                <tr>
                            <td class="FieldCell" width="108">Party Type</td>
                            <td class="FieldCell" colspan="3">
                            	<select size="1" name="selPartyType" class="FormElem" onChange="selAccountHead(this)">
								<option value="A">Select Party Type</option>

								</select>
                              </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Party Name</td>
                            <td class="FieldCell" colspan="3"> <input type="text" name="txtPartyName" size="40" class="FormElem"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Invoice Number</td>
                            <td class="FieldCell">
                            <input type="text" name="txtInvoiceNo" size="20" class="FormElem">
                            <!--
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                           <a href="javascript:popVoucherNo('C')">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:popVoucherNo('A')">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers "></a>-->

                            </td>
                            <td class="FieldCell" width="120"> Invoice Date</td>
                            <td class="FieldCell">
								     <% ' Function Call to Insert Date Picker
						  						Response.Write InsertDatePicker("ctlDate")
						  			 %>
							</td>
                                </tr>

                            <tr>
                            <td class="FieldCell" width="108">Category</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selPurCat" class="FormElem" >
									<option value="0">Select Purchase Category</option>
									<%
										'dim sCode,sValue
										sQuery = "Select CategoryCode,CategoryName From APP_M_InvoiceCategory Where ApplicableFor = 'P' Order By CategoryName "
									  	With objRs
									  		.CursorLocation = 3
									  		.CursorType = 3
									  		.Source = sQuery
									  		.ActiveConnection = con
									  		.Open
									  	End with
									  	Set objRs.Activeconnection = nothing
									  	Do while not objRs.EOF
									%>
											<option value="<%Response.Write objRs(0)%>"><%Response.Write objRs(1)%></option>
									<%

											objRs.MoveNext
										Loop
										objRs.Close
								    %>
								</select></td>
                                </tr>
                                    </table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>

							<tr>
								<td align="center" width="10" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Amend" name="btnCreate" class="ActionButton" onClick="VouCreate()" >
                                                                <!--input type="button" value="View" name="btnView" class="ActionButton" onClick="VouView()" >
                                                                <input type="button" value="Amendment" name="btnAmend" class="ActionButtonX" onClick="VouAmend()">
                                                                <input type="button" value="Delete" name="btnDel" class="ActionButton" onClick="VouDel()" -->
                                                                <input type="reset" value="Reset" name="B5" class="ActionButton"  >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="10" class="BottomPack" colspan="3">
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