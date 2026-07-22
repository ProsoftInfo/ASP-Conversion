<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassServiceEntry.asp
	'Module Name				:	Gate Pass - Service
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 05, 2010
	'Modified On				:	Jan 06,2011
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
<!--#include virtual="/include/CheckPrevFinYear.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
	Dim sOrgID,sOrgName,rsObj
	sOrgID = Session("organizationcode")
	set rsObj = Server.CreateObject("ADODB.Recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Gate Pass</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="application/xml" id="UoMData" data-itms-xml-island="1" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" id="OutData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="Data" data-itms-xml-island="1"><ROOT></ROOT></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/gatePassServiceEntry.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
	<form method="POST" name="formname">
	<input type=hidden name="hItem" value="">
	<input type=hidden name="hClass" value="">
	<input type=hidden name="hUnitID" value="<%=sOrgID%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Gate Pass - Service
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center">
									</td>
									<td width="100%" colspan="2">
										<div align="left">
											<table border="0" cellspacing="0" cellpadding="0" width="572">
											<!--	<tr>
													<td class="FieldCellSub">For Unit</td>
													<td class="FieldCellSub">
														<select size="1" name="selUnit" class="FormElem">
															<option value="select">Select</option>
															<%	'Calling the Function which populates Organization Unit list
																populateUnit
															%>
														</select>
													</td>
												</tr>-->

											<!--	<tr>
													<td class="FieldCellSub">Item Type</td>
													<td class="FieldCellSub" valign="top">
                                                        <select size="1" name="selItmType" class="FormElem">
															<option value="select">Select</option>
															<%	'Calling the Function which populates the Item Type list
														'		populateItemType
															%>
														</select>
													</td>
												</tr>-->
												<tr>
													<td class="FieldCellSub">Party</td>
													<td class="FieldCellSub">
														<input type="text" name="txtRefName" value size="60" class="FormElemRead" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Party" width="11" height="11" onClick="popSuppAgent()"></a>
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Item Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtItemDesc" value size="60" class="FormElemRead" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item" width="11" height="11" onClick="SelectItem()"></a>
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Other Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtDesc" maxlength=50 size="60" class="FormElem">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Reason</td>
													<td class="FieldCellSub">
														<input type="text" name="txtReason" maxlength=35 size="60" class="FormElem" value="SENT FOR REPAIRS - TO BE RETURNED">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Quantity & Value</td>
													<td class="FieldCellSub">
														<input type="text" name="txtQty" value size="12" class="FormElem" onkeypress="DoKeyPress('Y',7,3)">&nbsp;
														<select size="1" name="selUOM" class="FormElem">
															<option value="select">Select</option>
															<%
															populateUoM()
															%>
														</select>
														&nbsp;
														<input type="text" name="txtValue" size="12" class="FormElem" onkeypress="DoKeyPress('Y',7,2)" >&nbsp;
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Form JJ Applicable</td>
													<td class="FieldCellSub">
														<input type="checkbox" name="ChkFormJJ" value="Y" class="FormElem" >
														<input type="button" value="Add" name="B3" class="AddButtonX" onclick = "AddDetails()">
													</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack"></td>
								</tr>

								<tr>
									<td align="center"></td>
									<td width="100%" colspan="2">
										<div class="frmBody" id="frm1" style="width: 585; height:230;">
											<table border="0" cellspacing="1" class="ExcelTable" width="585" id=tblDetails>
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">
														<p align="center">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" width=500>
														Item Description
													</td>
													<td class="ExcelHeaderCell" align="center">Quantity</td>
													<td class="ExcelHeaderCell" align="center">UoM</td>
													<td class="ExcelHeaderCell" align="center">Value</td>
													<td class="ExcelHeaderCell" align="center">Form JJ Applicable</td>
													<td class="ExcelHeaderCell" align="center">Reason</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td colspan="4">
										<table border="0" cellspacing="0" cellpadding="0" width="572">
											<tr>
												<td class="FieldCell" Rowspan="3">Remarks</td>
												<td class="FieldCellSub" Rowspan="3" >
													<textarea rows="3" name="txtRemarks" cols="50" class="FormElem"></textarea>
												</td>
												<td class="FieldCell">Transport</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTransport" maxlength=50 size="50" class="FormElem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Taken By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTakenBy" maxlength=50 size="50" class="FormElem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Delivery By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtDeliveryBy" maxlength=50 size="50" class="FormElem">
												</td>
											</tr>
										</table>
									</td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" colspan="2">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<input type="button" value="Save" name="BtnSubmit" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
 													&nbsp;
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="BottomPack">
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

