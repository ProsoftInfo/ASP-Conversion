<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ProductionNoSeriesEntry.asp
	'Module Name				:	Production (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 25,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	ProductionNoSeriesInsert.asp
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
<HTML><HEAD><TITLE>Production Number Series Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="NoSeries"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="SeriesNoData" data-src="../NoSeries/xmldata/SeriesNumberDetail.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="SeriesList"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="Data"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemWiseData" data-src="xmldata/Itemwise_PackingNo.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"><root></root></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/NoSeriesEntryCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ProductionNoSeriesEntryCompat.js"></SCRIPT>
</HEAD>

<%
	Dim sItemTypeId, sItemTypeName, dcrs

	Set dcrs = Server.CreateObject("ADODB.Recordset")
%>
<BODY leftMargin=0 topMargin=0  onLoad="popSeriesNo()" >
<form method="POST" name="formname" action="ProductionNoSeriesInsert.asp">
<input type=hidden name="hSeriesType" value="">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hActivityName" value="">

<input type=hidden name="hSeriesNo" value="">
<input type=hidden name="hSeriesCode" value="">
<input type=hidden name="hFinFrom" value="">
<input type=hidden name="hFinTo" value="">
<input type=hidden name="hEntryNo" value="">

<input type=hidden name="hEditCheck" value="N">
<input type=hidden name="hTransNo" value="">
<input type=hidden name="hAmnFlag" value="">
<input type=hidden name="hTotEntNo" value="">
<input type=hidden name="hDispCheck" value="N">

<input type="hidden" name="hClassCode" value="">
<input type="hidden" name="hCatCode" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Packing No. Series Allocation</p>
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
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class=FieldCell>Select Unit</td>
											<td class="FieldCellSub">
												<select size="1" name="selUnit" class="FormElem" onChange="document.formname.selNumType.selectedIndex = 0">
													<option value="select">Select</option>
													<%	'Calling the Function which populates Organization Unit list
														populateUnit
													%>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Numbering Type</td>
											<td class="FieldCellSub">
											    <select size="1" name="selNumType" class="FormElem" onChange="DisableChk(this)">
													<option value="select">Select</option>
													<option value="U">Unit Wise</option>
													<!--<option value="I">Item Type Wise</option>-->
													<option value="P">Product Wise</option>
													<option value="C">Classification Wise</option>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Manual Numbering</td>
											<td class="FieldCellSub">
											    <input type="radio" name="radManual" value="Y" class="FormElem" onClick="DisableNoSeries()">Yes &nbsp;
											    <input type="radio" name="radManual" value="N" class="FormElem" checked onClick="EnableNoSeries()">No
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Select Classification</td>
											<td class="FieldCellSub">
											    <span id="txtClass" class="DataOnly">&nbsp;</span>&nbsp;&nbsp;<a href="#" onclick="SelectClassifcation()"><img style="cursor: hand" src="../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Classification"></a>
											</td>
										</tr>
										<!--<tr>
											<td class="FieldCell">Select Item Type</td>
											<td class="FieldCellSub">
											    <select size="1" name="selItemType" class="FormElem">
													<option value="select">Select</option>
												<%
												'	With dcrs
												'		.CursorLocation = 3
												'		.CursorType = 3
												'		.Source = "SELECT ITEMTYPEID, ITEMTYPENAME FROM INV_M_ITEMTYPE"
												'		.ActiveConnection = con
												'		.Open
												'	End With
												'	dcrs.ActiveConnection = Nothing
'
'													Set sItemTypeId = dcrs(0)
'													Set sItemTypeName = dcrs(1)
'
'													Do While Not dcrs.EOF
												%>
														<option value="<%=sItemTypeId%>"><%=sItemTypeName%></option>
												<%
'														dcrs.MoveNext
'													Loop
'													dcrs.Close
												%>
												</select>
												<a href="#" onclick="SelectItem()"><img style="cursor: hand" src="../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Item"></a>
											</td>
										</tr>-->
										<tr>
											<td class=FieldCell> Product Wise <br> (For clssification wise)</td>
											<td class="FieldCell">
												<input type="checkbox" name="chkProductWise" value="1" class="FormElem">
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell> Packing Type</td>
											<td class="FieldCell">
												<input type="checkbox" name="chkPacking" value="1" class="FormElem">
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell> Select No Series</td>
											<td class="FieldCellSub">
												<select size="1" name="selNoSeries" class="FormElem" onchange="DisplayTable()">
													<OPTION value="0">Select</option>
												</select>
                                            </td>
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td align="center" valign="top">
                                         <table id="tblBook" border="0" cellspacing="1" class="ExcelTable" >
                                            <tr>
											<td class="ExcelHeaderCell" align="center" width="10"><p align="center">S.No.</td>
											<td class="ExcelHeaderCell" align="center" width="75">Period</td>
											<td class="ExcelHeaderCell" align="center" width="50">Start No</td>
											<td class="ExcelHeaderCell" align="center" width="100">Prefix</td>
											<td class="ExcelHeaderCell" align="center" width="100">Suffix</td>
                                            </tr>
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
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															<input type="button" value="Submit" name="btnSubmit" class="ActionButton" onClick="validateForm('S')">
															<!--<input type="button" value="Cancel" name="btnCancel" class="ActionButton" onClick="Cancel('../welcome_Production.asp')">-->
															<input type="button" value="Update" name="btnUpdate" class="ActionButton" onClick="validateForm('U')" disabled>
															<input type="Reset" value="Reset" name="btnReset" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
								<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
									<DIV class=frmBody id="DisVoucher" style="width:600; visibility:hidden; height:1;">
										<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="580">
										<tr>
											<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
											<td class="ExcelHeaderCell" align="center">&nbsp</td>
											<td class="ExcelHeaderCell" align="center" >Numbering Type</td>
											<td class="ExcelHeaderCell" align="center">Manual Numbering</td>
											<td class="ExcelHeaderCell" align="center" >Select Item Type</td>
											<td class="ExcelHeaderCell" align="center" >Product Wise</td>
											<td class="ExcelHeaderCell" align="center">Packing Type</td>
											<td class="ExcelHeaderCell" align="center">Series Type</td>
										</tr>
										</table>

									</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
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

