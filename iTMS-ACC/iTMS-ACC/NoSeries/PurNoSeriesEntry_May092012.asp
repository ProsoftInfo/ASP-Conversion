<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PurNoSeriesEntry.asp
	'Module Name				:	Purchase (Master Creation)
	'Author Name				:	Malathi N
	'Modified By				:
	'Created On					:	Sep 29,2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	PurNoSeriesInsert.asp
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
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
	Dim sUnit,sActivity,sInvTy,sAgent,sSalTy,sSerTy,sQuery,Objrs
	Dim sSeriesNo,sSeriesCode

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="NoSeries"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="SeriesNoData" data-src="../NoSeries/xmldata/SeriesNumberDetail.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="SeriesList"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/PurNoSeriesEntryCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" onLoad="popSeriesNo()" MARGINWIDTH="0">

<form method="POST" name="formname" action="PurNoSeriesInsert.asp">
<input type=hidden name="hSeriesType" value="">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hActivityName" value="">
<input type=hidden name="hOptval" value="N">
<input type="hidden" name="hSerNo" value="">

<input type=hidden name="hSeriesNo" value="">
<input type=hidden name="hSeriesCode" value="">
<input type=hidden name="hFinFrom" value="">
<input type=hidden name="hFinTo" value="">
<input type=hidden name="hEntryNo" value="">

<input type=hidden name="hItemType" value="">
<input type=hidden name="hItemValue" value="">

<input type=hidden name="hNumFor" value="">
<input type=hidden name="hTempName" value="">
<input type=hidden name="hEditCheck" value="N">
<input type=hidden name="hTransNo" value="">
<input type=hidden name="hAmnFlag" value="">
<input type=hidden name="hTotEntNo" value="">
<input type=hidden name="hDispCheck" value="N">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">No. Series Allocation</p>
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
											<td class="FieldCell">Activity</td>
											<td class="FieldCellSub">
											    <select size="1" name="selActType" class="FormElem" onchange="FormReset()">
													<option value="0">Select</option>
													<%'Function to populate activities
														 popActivity
													%>
												</select>
											</td>
										</tr>

										<tr>
											<td class=FieldCell> Select Unit</td>
											<td class="FieldCellSub">
												<select size="1" name="selUnit" class="FormElem"  onChange="GetSeriesList('U')">
													<option value="0">Select</option>
													<%	'Calling the Function which populates Organization Unit list
														populateUnit
													%>
												</select>
											</td>
										</tr>

										<tr>
											<td class=FieldCell>For </td>
											<td class="FieldCellSub">
												<INPUT type="checkbox" name="chkFor" value = "D" checked> Domestic
												<INPUT type="checkbox" name="chkFor" value = "E" checked> Import
											</td>
										</tr>

									<!--	<tr>
											<td class=FieldCell>Item Type </td>
											<td class="FieldCellSub">
												<INPUT type="radio" name="optItem" value = "A" onClick="PopupDet(this,'I')" checked> All
												<INPUT type="radio" name="optItem" value = "S" onClick="PopupDet(this,'I')"> Specific
												&nbsp;
												<Input type="text" name="txtItem" size="20" maxlength="50" readonly class="FormElemRead")">
											</td>
										</tr>-->



										<tr>
											<td class=FieldCell> Select No Series</td>
											<td class="FieldCellSub">
												<select size="1" name="selNoSeries" class="FormElem" onChange="DisplayTable()">
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
															<input type="button" value="Done" name="btnSubmit" class="ActionButton" onClick="validateForm('S')" >
															<input type="button" value="Update" name="btnUpdate" class="ActionButton" onClick="validateForm('U')" disabled>
															<input type="button" value="Cancel" name="btnCancel" class="ActionButton" onClick="Cancel('../welcome_Purchase.asp')">
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
							<DIV class=frmBody id="DisVoucher" style="width:500; visibility:hidden; height:1;">
								<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="500">
								<tr>
									<td align="center" class="ExcelHeaderCell" width="100">S.No.</td>
									<td align="center" class="ExcelHeaderCell" width="100">&nbsp</td>
									<td align="center" class="ExcelHeaderCell" width="100">Number For</td>
									<td align="center" class="ExcelHeaderCell" width="100">Item Type</td>
									<td align="center" class="ExcelHeaderCell" width="100">Series Type</td>
								</tr>
								</table>
							</div>
															</td>
															<td align="center" class="ClearPixel" width="5">
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

