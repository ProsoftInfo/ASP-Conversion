<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	stockLotSerFabPop.asp
	'Module Name				:	Inventory (Opening Stock Lot and Serial Details)
	'Author Name				:	TAJUDEEN S
	'Created On					:	June 1, 2004
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->
<!--#include virtual="/include/UoMDecimal.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock - Quantity Breakup</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/openingLotDetails.js"></SCRIPT>
</HEAD>
<%
	dim STemp, arrTemp, sType, iItem, iClass, sOrgID, iTotQty, sStoName, sStoresUom
	dim iQty, sCheck, iNo, sAltUom, sAltCheck

	sTemp = trim(Request.QueryString("sTemp"))
	'Response.Write "sTemp = "& STemp
	arrTemp = split(sTemp ,":")
	sType	= arrTemp(0)
	iItem	= arrTemp(1)
	iClass	= arrTemp(2)
	sOrgID	= arrTemp(3)
	iTotQty = arrTemp(6)
	sStoName = arrTemp(7)
	sStoresUom = arrTemp(8)
	iQty = arrTemp(10)
	iNo = arrTemp(11)
	sAltUom = arrTemp(12)

	if trim(sAltUom) = "select" then sAltUom = "-"

	sCheck = UoMDecimal(sStoresUom)
	sAltCheck = UoMDecimal(sAltUom)
%>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init('<%=STemp%>','<%=sCheck%>','<%=sAltCheck%>')">
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Quantity Breakup
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
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idItemName"><%=ItemDisplay(iItem,iClass)%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Store -- Bin</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idStoreName"><%=sStoName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idQty"><%=iQty%></span> -
                                                <span class="DataOnly" ><%=sStoresUom%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
											<td class="FieldCell">Quantity Entered</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idQtyEntered">0</span> -
                                                <span class="DataOnly"><%=sStoresUom%>&nbsp;</span>
                                            </td>
                                        </tr>
<!--Begin New-->
										<tr>
											<td align="center" colspan="3" class="MiddlePack"></td>
										</tr>

										<tr>
											<td valign="top" width="100%" colspan="3">
										        <table cellpadding="0" cellspacing="0">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p></td>
																	<td class="GroupTitle" width="160">
																		<p align="center"><input type="checkbox" value="Y" name="chkAltUom" class="FormElem" onClick="CheckEnable()"> Alternate UoM Breakup
										                            </td>
																	<td class="GroupTitleRight"><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<div align="left">
																<table cellpadding="0" cellspacing="0" width="100%">
																	<tr>
																		<td class=MiddlePack colspan="3"> <p align="left"> </td>
																	</tr>

																	<tr>
																		<td class=FieldCell>
																			<table border="0" cellspacing="0" cellpadding="0">
																				<tr>
																					<td class="FieldCell"></td>
																					<td class="FieldCell" width="40">Gross </td>
																					<td class="FieldCellSub">
										                                                <input type="text" name="txtGross" class="FormElem" DISABLED size="13" onkeypress="DoKeyPress('<%=sCheck%>',7,1)" style="text-align:right" onBlur="CheckGrossQty(this)">
										                                                &nbsp;<span class="DataOnly"><%=sAltUom%>&nbsp;</span>
																					</td>
																				</tr>
																				<tr>
																					<td class="FieldCell"></td>
																					<td class="FieldCell">Nett</td>
																					<td class="FieldCellSub">
										                                                <input type="text" name="txtNett" class="FormElem" DISABLED size="13" onkeypress="DoKeyPress('<%=sCheck%>',7,1)" style="text-align:right" onBlur="CheckNettQty(this)">
										                                                &nbsp;<span class="DataOnly"><%=sAltUom%>&nbsp;</span>
																					</td>
																				</tr>
																			</table>
																		</td>
																	</tr>
																</table>
															</div>
														</td>
													</tr>
										        </table>
											</td>
										</tr>
										<tr>
											<td align="center" colspan="3" class="MiddlePack"></td>
										</tr>

<!--End New-->
                                        <tr>
                                            <td class="FieldCell" colspan="2">Quantity&nbsp;
                                                <input type="text" name="txtQty" class="FormElem" size="13" onkeypress="DoKeyPress('<%=sCheck%>',7,1)" style="text-align:right">&nbsp;
                                                <input type="button" value="Add Breakup" name="BtnAdd" class="AddButton" onClick="AddRow()">
                                            </td>
										</tr>
                                        <tr>
                                            <td class="FieldCell" colspan="5"><p align="center"></p></td>
                                        </tr>
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td align="center">
									<div class="frmBody" id="frm2" style="width:100%; height:190;" >
										<table border="0" cellspacing="1" id="tblLot" class="ExcelTable" align="center">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
												<td class="ExcelHeaderCell" align="center">Gross</td>
												<td class="ExcelHeaderCell" align="center">Nett</td>
											</tr>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
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
                                                    <input type="button" value="Done" name="BtnDone" class="ActionButton" onClick="CheckSubmit()" DISABLED>
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

