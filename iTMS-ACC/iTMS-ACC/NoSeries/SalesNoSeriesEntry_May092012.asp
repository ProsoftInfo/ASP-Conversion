<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SalesNoSeriesEntry.asp
	'Module Name				:	Sales (Master Creation)
	'Author Name				:	Subbiah
	'Modified By				:	Manohar Prabhu.R
	'Created On					:	August 07,2003
	'Modified On				:	25 May 2004
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	SalesNoSeriesInsert.asp
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
<%
	Dim sUnit,sActivity,sInvTy,sAgent,sSalTy,sSerTy,sQuery,Objrs
	Dim sSeriesNo,sSeriesCode,Objrs2,oDom,Root,MainNode,EntryNode

	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set Objrs2 = Server.CreateObject("ADODB.RecordSet")
	Set oDom = Server.CreateObject("Microsoft.XMLDOM")
	Set Root = oDom.createElement("Root")
	oDom.appendChild Root
	sQuery = "Select SeriesNo,Description,CounterType,UsedBy, "&_
			 "NumberLength From Ms_NumberSeries "

	With Objrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	End With
	Set Objrs.ActiveConnection = Nothing
	Do While Not Objrs.EOF
		Set MainNode = oDom.createElement("Series")
		MainNode.setAttribute "No", Objrs(0)
		MainNode.setAttribute "Description", Objrs(1)
		MainNode.setAttribute "Type", Objrs(2)
		MainNode.setAttribute "UsedBy", Objrs(3)
		MainNode.setAttribute "NumberLength", Objrs(4)

		sQuery = "Select EntryNo,Period,Number,Prefix,Suffix From "&_
				 "Ms_NumberSeriesEntry Where SeriesNo = "&Objrs(0)&" "

		With Objrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End With
		Set Objrs2.ActiveConnection = Nothing

		Do While Not Objrs2.EOF
			Set EntryNode = oDom.createElement("Entry")
			EntryNode.setAttribute "EntryNo", Objrs2(0)
			EntryNode.setAttribute "Period", Objrs2(0)
			EntryNode.setAttribute "Number", Objrs2(0)
			EntryNode.setAttribute "Prefix", Objrs2(0)
			EntryNode.setAttribute "Suffix", Objrs2(0)

			MainNode.appendChild EntryNode
			Objrs2.MoveNext
		loop
		Objrs2.Close
		Root.appendChild MainNode
		Objrs.MoveNext
	loop
	Objrs.Close

	oDOM.Save server.MapPath("/NoSeries/xmldata/SeriesNumberDetail.xml")




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
<SCRIPT LANGUAGE=javascript SRC="../scripts/SalesNoSeriesEntryCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" onLoad="popSeriesNo();init()" MARGINWIDTH="0">

<form method="POST" name="formname" action="SalesNoSeriesInsert.asp">
<input type=hidden name="hSeriesType" value="">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hActivityName" value="">
<input type=hidden name="hOptval" value="N">
<input type=hidden name="hAgentName" value="">

<input type=hidden name="hSeriesNo" value="">
<input type=hidden name="hSeriesCode" value="">
<input type=hidden name="hFinFrom" value="">
<input type=hidden name="hFinTo" value="">
<input type=hidden name="hEntryNo" value="">

<input type=hidden name="hItemType" value="">
<input type=hidden name="hItemValue" value="">

<input type=hidden name="hInvType" value="">
<input type=hidden name="hInvValue" value="">

<input type=hidden name="hSaleType" value="">
<input type=hidden name="hSaleValue" value="">

<input type=hidden name="hNumFor" value="">

<input type=hidden name="hAgentCode" value="">

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
											    <select size="1" name="selActType" class="FormElem" onChange="Activity_change(this)">
													<option value="0">Select</option>
													<option value="QUT">Quotation</option>
													<!--Added on	: 19/11/2004
													    Reason		: No Series for Order Creation also (Req From LMC)
													-->
													<option value="ORD">Order Creation</option>
													<option value="OCR">Order Confirmation</option>
													<option value="ORP">Order Processing</option>
													<option value="DIS">Despatch Instruction Slip</option>
													<option value="PIS">Production Instruction Slip</option>
													<option value="PUR">Purchase Instruction Slip</option>
													<option value="RIS">Repack Instruction Slip</option>
													<option value="PFO">Proforma Invoice</option>
													<option value="INV">Invoice</option>
													<option value="FJJ">Form JJ</option>
													<option value="DIN">Depot Inward</option>
												</select>
											</td>
										</tr>

										<tr>
											<td class=FieldCell> Select Unit</td>
											<td class="FieldCellSub">
												<select size="1" name="selUnit" class="FormElem" onChange="GetSeriesList('U')">
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
												<INPUT type="checkbox" name="chkFor" value = "E" checked> Export
											</td>
										</tr>

									<!--	<tr>
											<td class=FieldCell>Item Type </td>
											<td class="FieldCellSub">
												<INPUT type="radio" name="optItem" value = "A" onClick="PopupDet(this,'I')" checked> All
												<INPUT type="radio" name="optItem" value = "S" onClick="PopupDet(this,'I')"> Specific
												&nbsp;
												<Input type="text" name="txtItem" size="60" maxlength="150" readonly class="FormElemRead">
											</td>
										</tr>-->

										<tr>
											<td class=FieldCell>Invoice Type </td>
											<td class="FieldCellSub">
												<INPUT type="radio" name="optInv" value = "A" onClick="PopupDet(this,'V')" checked> All
												<INPUT type="radio" name="optInv" value = "S" onClick="PopupDet(this,'V')"> Specific
												&nbsp;
												<Input type="text" name="txtInv" size="60" maxlength="150" readonly class="FormElemRead">
											</td>
										</tr>

										<tr>
											<td class=FieldCell>Sale Type </td>
											<td class="FieldCellSub">
												<INPUT type="radio" name="optSale" value = "A" onClick="PopupDet(this,'S')" checked> All
												<INPUT type="radio" name="optSale" value = "S" onClick="PopupDet(this,'S')"> Specific
												&nbsp;
												<Input type="text" name="txtSale" size="60" maxlength="150" readonly class="FormElemRead">
											</td>
										</tr>

										<tr>
											<td class=FieldCell>Commission Agent /<br>Guarantor wise <br>Numbering </td>
											<td class="FieldCellSub">
												<INPUT type="radio" name="optComm" value = "Y" onClick="PopupDet(this, 'A')"> Yes
												<INPUT type="radio" name="optComm" value = "N" checked onClick="PopupDet(this, 'A')"> No
												<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<Input type="text" name="txtAgent" size="60" maxlength="150" readonly class="FormElemRead" -->
											</td>
										</tr>

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
															<input type="button" value="Cancel" name="btnCancel" class="ActionButton" onClick="Cancel('SalesNoSeriesEntry.asp')">
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
							<DIV class=frmBody id="DisVoucher" style="width:585; visibility:hidden; height:1;">
								<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="970">
								<tr>
									<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
									<td class="ExcelHeaderCell" align="center">&nbsp</td>
									<td class="ExcelHeaderCell" align="center" >Number For</td>
									<td class="ExcelHeaderCell" align="center">Item Type</td>
									<td class="ExcelHeaderCell" align="center" >Invoice Type</td>
									<td class="ExcelHeaderCell" align="center" >Sale Type</td>
									<td class="ExcelHeaderCell" align="center">Agent</td>
									<td class="ExcelHeaderCell" align="center">Series Type</td>
									<td class="ExcelHeaderCell" align="center">Agent Code</td>
									<td class="ExcelHeaderCell" align="center" width="80">Item Values</td>
									<td class="ExcelHeaderCell" align="center" width="120">Invoice Values</td>
									<td class="ExcelHeaderCell" align="center" >Sales Values</td>

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

