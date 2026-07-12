<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GJDayBookCriteria.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	N.Rajkumar
	'Created On					:	15th May 2003
	'Modified By				:	UmaMaheswari s
	'Modified On				:	April 19,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	DBGJView.asp
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
<!--#include file="../../include/DatabaseConnection.asp."-->
<!--#include file="../../include/populate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/sessionVerify.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>GJ - Day Book</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData"><Book/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData"><account/></script>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/GJDayBookCriteriaCompat.js"></SCRIPT>
<%
dim sFinPeriod,sFinTemp,sMaxDate,sMinDate,Da,Mo,Yr,sSelDayBook
Dim sFinFromDate,sFinToDate
sFinPeriod = Session("FinPeriod")
sFinTemp = Split(sFinPeriod,":")
sFinFromDate = "01/04/"& sFinTemp(0)
sFinToDate = "31/03/"&sFinTemp(1)

sSelDayBook  = Request("RadDayBook")
If sSelDayBook = "" Then sSelDayBook = "01"
'Response.Write sMinDate & " *** "& sMaxDate
%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="setdate();DisplayBook()">

<form method="POST" name="formname" action="">
	<Input type="hidden" name="hUnitId" value="<%=Session("organizationcode")%>">
	<Input type="hidden" name="hUnitName" value="<%=Session("orgshortName")%>">
	<Input type="hidden" name="hFDate" value="<%=sFinFromDate%>">
	<Input type="hidden" name="hTDate" value="<%=sFinToDate%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class=PageTitle>
			GJ Day Book
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
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" align=left width="100%">
<TABLE BORDER="0" CELLSPACING=0 CELLPADDING=0>
<!--<TR><TD class="FieldCell"> Organization</TD>
<TD class="FieldCellSub">
                                                           <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationList%>
                              </select>
                              </TD>
</TR>-->
<tr>
<TD class="FieldCell">
GJ Day Book</TD>
<TD class="FieldCellSub">
                                                            <select size="1" name="selBook" class="FormElem" OnChange="SelNew()">
                        <option value="S">Select Book</option>
                            </select></TD>
</tr>
</TABLE>

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
								</td>
								<td valign="top" align="center">
																<table class="BodyTable" cellSpacing="0" cellPadding="2" border="0">
																<tbody>
																<tr>
																	<td class="ExcelHeaderCell">Filter By</td>
																	<td class="ExcelHeaderCell">From&nbsp;&nbsp;</td>
																    <td class="ExcelHeaderCell">To&nbsp;&nbsp;&nbsp;&nbsp;</td>
																    <td class="ExcelHeaderCell"></td>
																</tr>
																<tr>
																    <td class="FieldCellSub">
																		<!--<input onclick="OptSelection()" type="radio" value="VouDate" name="optCriteria"  CHECKED> Voucher Date-->
																		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Voucher Date
																	</td>
																    <td align="left" class="FieldCellSub">
																	   <input type="date" id="ctlVouFromDate" name="ctlVouFromDate" onblur="MinDate()" class="FormElem itms-date-picker" style="width: 89px; height: 20px;">
																	</td>
																    <td align="left" class="FieldCellSub">
																	    <input type="date" id="ctlVouToDate" name="ctlVouToDate" onblur="MinDate()" class="FormElem itms-date-picker" style="width: 89px; height: 20px;">
																	</td>
																    <td align="left" class="FieldCellSub">&nbsp;</td>
																</tr>
																<tr>
																	<td class="FieldCellSub">
																		<!--<input type="radio" value="VouNo" name="optCriteria" onclick="OptSelection()">Voucher	Number&nbsp;-->
																		<input type="checkbox" value="VouNo" name="chkBox1" onclick="OptSelection()">Voucher	Number&nbsp;
																	</td>
																	<td align="left" class="FieldCellSub"><input class="FormElem"  size="11" name="txtNoFrom" Readonly></td>
																	<td align="left" class="FieldCellSub"><input class="FormElem"  size="11" name="txtNoTo" Readonly></td>
																	<td align="left" class="FieldCellSub">&nbsp;</td>
																</tr>
																<tr>
																   <td class="FieldCellSub">
																		<!--<input type="radio" onclick="OptSelection()"  value="Amount" name="optCriteria">-->
																		<input type="Checkbox" onclick="OptSelection()"  value="Amount" name="chkBox2">
																		Amount&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
																	<td align="left" class="FieldCellSub"><input class="FormElem" size="11" Readonly name="txtGAmount"></td>
																	<td align="left" class="FieldCellSub"><input class="FormElem" size="11" Readonly name="txtLAmount"></td>
																	<td align="left" class="FieldCellSub"></td>
																</tr>

																<tr>
																	<td class="FieldCellSub">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																		Account Head</td>
																	<td align="left" class="FieldCellSub">
																		<select class="FormElem" OnChange="SelectAccHead()" size="1" name="SelAccHead">
																		  <option value="0">Select Option</option>
																		  <option value="S" >Select Account Head</option>
																		</select>
																   </td>
																   <!--<td colSpan="2" class="FieldCellSub">
																   <span id="spAccHead" class="DataOnly"></span>&nbsp;
																   </td>-->
																</tr>
																 <tR>
																	<td colSpan="3" class="FieldCellSub">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																		<span id="spAccHead" class="DataOnly"></span>&nbsp;
																   </td>
																</tr>
																<!--<tr>
																	<td vAlign="center" class="ExcelHeaderCell" align="center">Viewed
                                                                      By</td>
																<td vAlign="center" align="center" class="ExcelHeaderCell">From&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell">To&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell"></td>
                                                                </tr>
																<tr>
																	<td vAlign="center" class="FieldCellSub"><input type="radio" value="VouNo" CHECKED name="optCriteria" onclick="OptSelection()">
																	Voucher	Number&nbsp;</td>
																<td vAlign="center" align="left" class="FieldCellSub" width="10"><input class="formelem"  size="11" name="txtNoFrom"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub" width="10"><input class="formelem"  size="11" name="txtNoTo"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"></td>
                                                                </tr>
                                                                <tr>
                                                                  <td vAlign="center" class="FieldCellSub"><input onclick="OptSelection()" type="radio" value="VouDate" name="optCriteria">
                                                                    Voucher Date</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouFromDate")
 %>
 </td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouToDate")
 %>

</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub">

</td>
                                                                </tr>
                                                                <tr>
                                                                  <td vAlign="center" class="FieldCellSub"><input type="radio" onclick="OptSelection()"  value="Amount" name="optCriteria">
                                                                    Amount&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtGAmount"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtLAmount"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"></td>
                                                                </tr>
                                                                <tr>
                                                                  <td vAlign="center" class="FieldCellSub"><input type="radio" onclick="OptSelection()" value="AccHead" name="optCriteria">
                                                                    Account Head</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"><select class="formelem" onchange="SelectAccHead()" disabled size="1" name="SelAccHead">
                                                                      <option value="0">Select Option</option>
                                                                      <option value="S">Selected Account Head</option>
                                                                    </select></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub" colspan="2"><span id="spAccHead" class="DataOnly"></span>&nbsp;</td>
                                                                </tr>-->
                                                              </tbody>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="ActionCell">
                                                <input type="button" value="View" class="ActionButton" onClick="CheckSubmit('S')" >
                                                <input type="button" value="Print" class="ActionButton" onClick="CheckSubmit('P')"  id=button1 name=button1>
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
