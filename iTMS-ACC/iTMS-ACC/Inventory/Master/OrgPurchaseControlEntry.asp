<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgPurchaseControlEntry.asp
	'Module Name				:	Inventory (Organization Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 20, 2002
	'Modified By                :   Ragavendran R
	'Modified On				:   July 21,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	OrgPurchaseControlInsert.asp
	'Procedures/Functions Used	:	populateUnit
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
<%
	dim dcrs,sSql,sLocName,iCtr,sOrgCode,sOrgName
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	iCtr = 1
	
	sOrgCode = Session("organizationcode")
	sOrgName = Session("OrgShortName")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Organization Control Definition - Purchase</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="storageData" src="../xmldata/Storage.xml"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
    function CheckSubmit()
    {
		document.forms[0].action = "OrgPurchaseControlInsert.asp"
		document.forms[0].submit();
    }

</SCRIPT>
<SCRIPT LANGUAGE=javascript>window.ITMS_STORAGE_APPS = "PU";</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/orgStorageDisplay.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hOrgName" value="<%=sOrgName%>">
<input type=hidden name="hOrgCode" value ="<%=sOrgCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Organization Control Definition</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
							    <td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
										<tr><a href="OrgControlDefn.asp">
											<td width="100%" align="center">Org Defn
											</td>
											</a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="125">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="StoreLocations.asp">
											<td width="100%" align="center">
                                                Storage Location
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="MASUOMENTRY.asp">
											<td width="100%" align="center">
                                                UOM
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
							  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
							   <tr><a href="PackingTypes.asp">
								  <td width="100%" align="center">Packing Type</td></a>
								</tr>
							  </table>
							</td>
					        <td class="TabCell" valign="bottom" align="center" width="105">
						          <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
						           <tr><a href="../../NoSeries/InventoryNoSeriesEntry.asp">
							          <td width="100%" align="center">Number Series</td></a>
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
				<tr>
		            <td>
		                <table border="0" cellpadding="0" cellspacing="0" width="100%">
		                    <tr>
		                        <td class="TabCurrentCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Purchase
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgInventoryControlEntry.asp">
											<td width="100%" align="center">
                                                Inventory
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="85">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgInspectionControlEntry.asp">
											<td width="100%" align="center">
                                                Inspection
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgSalesControlEntry.asp">
											<td width="100%" align="center">
                                                Sales
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="OrgManufacturingControlEntry.asp">
									  <td width="100%" align="center">Manufacturing</td></a>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="150">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="MasSubCOntProcess.asp">
									  <td width="100%" align="center">Sub-Contract Process</td></a>
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
								<td align="center" colspan="3" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" class="">
                                        <tr>
											<td align="center" class="MiddlePack" rowspan="2">
												<table cellpadding="0" cellspacing="0" width="300">
													<tr>
														<td>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class='GroupTitleLeft' width="10"><p align="left">&nbsp;</p>
                                                        </td>
														<td class='GroupTitle' width="45"><p align="center">Allow
                                                        </td>
														<td class='GroupTitleRight'><p align="left">&nbsp;
                                                        </td>
													</tr>
												</table>
                                                    </td>
													</tr>
													<tr>
														<td class=GroupTable>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class=MiddlePack colspan="3"> </td>
													</tr>
													<tr>
														<td class=FieldCellSub width="230"> <p align="left"> </td>
														<td class='FieldCellSub'><p align="left">Yes</td>
														<td class='FieldCellSub'><p align="left">&nbsp;No&nbsp;&nbsp; </td>
													</tr>
													<tr>
														<td class=FieldCellSub width="230"> <p align="left"> Receipt / Ordering of Alternate Items</p>
                                                        </td>
														<td class='FieldCellSub'>
                                                        <p align="left">
                                                        <input type="radio" value="1" name="radAlter" class="FormElem"></p>
                                                        </td>
														<td class='FieldCellSub'>
                                                        <p align="left">
                                                        <input type="radio" value="0" name="radAlter" class="FormElem" checked></p>
                                                        </td>
													</tr>
													<tr>
														<td class=FieldCellSub width="230"> <p align="left"> Over Receipts</p>
                                                        </td>
														<td class='FieldCellSub'>
                                                        <p align="left">
                                                        <input type="radio" value="1" name="radOReceipts" class="FormElem"></p>
                                                        </td>
														<td class='FieldCellSub'>
                                                        <p align="left">
                                                        <input type="radio" value="0" name="radOReceipts" class="FormElem" checked></p>
                                                        </td>
													</tr>
													<tr>
														<td class=FieldCellSub width="230"> <p align="left"> Under Receipts</p>
                                                        </td>
														<td class='FieldCellSub'>
                                                        <p align="left">
                                                        <input type="radio" value="1" name="radUReceipts" class="FormElem"></p>
                                                        </td>
														<td class='FieldCellSub'>
                                                        <p align="left">
                                                        <input type="radio" value="0" name="radUReceipts" class="FormElem" checked></p>
                                                        </td>
													</tr>
													<tr>
														<td class=FieldCellSub width="230"> Unordered Receipts
                                                        </td>
														<td class='FieldCellSub'>
                                                        <input type="radio" value="1" name="radUOReceipts" class="FormElem">
                                                        </td>
														<td class='FieldCellSub'>
                                                        <input type="radio" value="0" name="radUOReceipts" class="FormElem" checked>
                                                        </td>
													</tr>
													<tr>
														<td class=FieldCellSub width="230"> MODVAT Credit
                                                        </td>
														<td class='FieldCellSub'>
                                                        <input type="radio" value="1" name="radModvat" class="FormElem">
                                                        </td>
														<td class='FieldCellSub'>
                                                        <input type="radio" value="0" name="radModvat" class="FormElem" checked>
                                                        </td>
													</tr>
												</table>
                                                        </td>
													</tr>
												</table>
											</td>
											<td align="center" class="MiddlePack" width="5" rowspan="2">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                            <p>&nbsp;
											</td>
											<td align="center" class="ClearPixel">

                                                <img border="0" src="../../assets/images/clearpixel.gif" width="9" height="9">

											</td>
                                        </tr>
                                        <tr>
											<td align="left" class="MiddlePack" valign="top">
												<DIV class=frmBody id=frm3 style="width: 280; height:145;">

                                                <table border="0" cellspacing="1" class="ExcelTable" Id ="tblLoc" name="tblLoc">
													<tr>
														<td class="ExcelHeaderCell" align="center">S.No.</td>
														<td class="ExcelHeaderCell" align="center" width="100%">Receiving Storage Location</td>
													</tr>
													<!--tr>
														<td class="ExcelSerial" align="center"></td>
														<td class="ExcelDisplayCell" width="100%"></td>
													</tr-->
                                                </table>
												</div>
											</td>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Save" name="B1" class="ActionButton" onClick="javascript:CheckSubmit()" >
													<input type="button" value="Reset" name="B1" class="ActionButton" onClick="ClearAll()">
													<input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
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

