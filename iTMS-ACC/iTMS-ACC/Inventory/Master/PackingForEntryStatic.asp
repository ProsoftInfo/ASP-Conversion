<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PackingForEntry.asp
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
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
<script type="application/xml" data-itms-xml-island="1" id="SubLevelQty"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/packingForEntry.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot Entry
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
                                    <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Item Description </td>
									<td class="FieldCellSub">
										<span Id="dataonly">xxxx</span>
									</td>
									<td class="FieldCellsub">Date Of Arrival / Packing</td>
									<td class="FieldCellSub">
										<%Response.Write InsertDatePicker("ctlDate")%>
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Lot No</td>
									<td class="FieldCellSub" colspan="3">
										<select class="FormElem" Name="selLotNo">
											<option value="N">New Lot&nbsp;&nbsp;</option>
										</select>
										&nbsp;[Last Used&nbsp;<font color="red"><b>XXX</b></font>]&nbsp;
										<input type="text" Name="LotNo" size="5" class="FormElem" value="">
										&nbsp;
										<Input type="checkbox" Name="chkAutoSerial" class="FormElem" value="">
										Auto Serial&nbsp;[Last Used&nbsp;<font color="red"><b>XXX</b></font>]&nbsp;
									</td>
                                </tr>

                                <tr>
									<td class="FieldCellsub">Packing Type</td>
									<td class="FieldCellSub" colspan="3">
										<select class="FormElem" Name="selPackType">
											<option value="S">Yarn Bag</option>
										</select>
										&nbsp;No OF Packs&nbsp;
										<Input type="text" Name="txtNoOfPack" size="6" class="FormElem" value="5">
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Gross / Pack</td>
									<td class="FieldCellSub" colspan="3">
										<Input type="text" Name="txtuniSerQty" size="10" class="FormElem" value="62.5">
										<span id="Dataonly">KGS</span>
										&nbsp;with
										<input type="text" Name="txtSubLevelQty" size="5" class="FormElem" value="50">
										<span id="Dataonly">Cones</span>&nbsp;of
										<input type="text" Name="txtSubLevelwt" size="5" class="FormElem" value="1.25">
										<span id="Dataonly">KGS</span>&nbsp;each
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellSub">Applicable Mix</td>
									<td class="FieldCellSub" colspan="2">
										<select class="FormElem" Name="selPackType">
											<option value="S">Mix Name / code</option>
										</select>
									</td>
									<td class="FieldcellSub">
										<Input type="Button" name="btnAdd" class="ActionButton" value="Add" onclick="">
									</td>
                                </tr>

                                <tr>
                                     <td class="MiddlePack" colspan="4"></td>
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
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class="ExcelHeaderCell" align="left" colspan="4">Lot No - Qty [Rate / KGS]</td>
												<td class="ExcelHeaderCell" align="center" colspan="2">Party</td>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10%">Serial</td>
												<td class="ExcelHeaderCell" align="center" >
													<a href="#"><img style="cursor:hand" border="0" width="11" height="11" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Delete" ></a>
												</td>
												<td class="ExcelHeaderCell" align="center" width="20%">Packing Type</td>
												<td class="ExcelHeaderCell" align="center" width="20%">Gross Qty</td>
												<td class="ExcelHeaderCell" align="center" width="20%">Nett Qty</td>
												<td class="ExcelHeaderCell" align="center" width="20%">No Of <span Id=""><b>Cones</b></span></td>
												<td class="ExcelHeaderCell" align="center" width="20%">Serial</td>
												<td class="ExcelHeaderCell" align="center" width="20%">Qty</td>
											</tr>
											<tr>
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class=ExcelDisplaycell align="Left" colspan="4"><b>1 - 312.5 KGS [Rs.125.00]</b></td>
												<td class=ExcelDisplaycell align="Left" colspan="2"><b>1 - 312.5 KGS</b></td>
											</tr>
											<tr>
												<td class=ExcelSerial align=center>1</td>
												<td class=ExcelDisplaycell align=center>
													<Input type="Checkbox" name="chkbox" value="">
												</td>
												<td class="ExcelInputcell" align="center">
													<select class="FormElem" Name="selPackType">
														<option value="S">Yarn Bag</option>
													</select>
												</td>
												<td class="ExcelInputcell" align="Left">
													<Input type="text" name="txtSubLevelQty" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
													<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Sub Level Qty" onclick="ShowSubLevelQty()"></a>
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtQty" value="" class="FormElem" size="10">
												</td>
											</tr>
											<tr>
												<td class="ExcelSerial" align="center">2</td>
												<td class="ExcelDisplaycell" align="center">
													<Input type="Checkbox" name="chkbox" value="">
												</td>
												<td class="ExcelInputcell" align="center">
													<select class="FormElem" Name="selPackType">
														<option value="S">Yarn Bag</option>
													</select>
												</td>
												<td class="ExcelInputcell" align="Left">
													<Input type="text" name="txtSubLevelQty" value="" class="FormElem" size=10>
												</td><td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
													<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Sub Level Qty" onclick="ShowSubLevelQty()"></a>
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align=center>
													<Input type="text" name="txtQty" value="" class="FormElem" size="10">
												</td>
											</tr>
											<tr>
												<td class=ExcelDisplayCell align=right colspan="3"><b>Total</b></td>
												<td class=ExcelDisplayCell align=right><b>xxxx</b></td>
												<td class=ExcelDisplayCell align=right><b>xxxx</b></td>
												<td class=ExcelDisplayCell align=right><b>xxxx</b></td>
												<td class=ExcelDisplayCell align=right colspan="2"></td>
											</tr>
									</Table>
								</td>
                            </tr>

                <!--<div class="frmBody" id="frm2" style="width: 50%; height:150;">
             	</div>-->

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
													<input type="button" value="Save" name="B3" class="ActionButton" onclick="Save()" >
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

