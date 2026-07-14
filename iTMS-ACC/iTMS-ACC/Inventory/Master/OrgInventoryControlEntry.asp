<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgInventoryControlEntry.asp
	'Module Name				:	Inventory (Organization Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Modified By                :   Ragavendran R
	'Modified On				:   July 21,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	OrgInventoryControlInsert.asp
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
	dim dcrs,sSql,sLocName,iCtr,iOrgIndex,sOrgCode,sOrgName
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	iCtr = 1
	sOrgCode = Session("organizationcode")
	sOrgName = Session("OrgShortName")
	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Organization Control Definition - Inventory</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="storageData" src="../xmldata/Storage.xml"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function checkNumbers(val){
	var valid = "0123456789."
	var temp;

	for (var i=0; i < val.length; i++) {
		temp = "" + val.substring(i, i+1);
		if (valid.indexOf(temp) == "-1") 
			return false;
		else
			return true;
	}
}

function CheckSubmit(){
     if (trimTrue(document.forms[0].txtStock.value) != "" && (!checkNumbers(document.forms[0].txtStock.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtStock.select();
		return false;
	}
	else if (document.forms[0].radABC[0].checked && (trimTrue(document.forms[0].txtA.value) == "")) {
		alert("Enter Classification A Value");
		document.forms[0].txtA.select();
		return false;
	}
	else if (document.forms[0].radABC[0].checked && (!checkNumbers(document.forms[0].txtA.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtA.select();
		return false;
	}
	else if (document.forms[0].radABC[0].checked && (trimTrue(document.forms[0].txtB.value) == "")) {
		alert("Enter Classification B Value");
		document.forms[0].txtB.select();
		return false;
	}
	else if (document.forms[0].radABC[0].checked && (!checkNumbers(document.forms[0].txtB.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtB.select();
		return false;
	}
	else if (document.forms[0].radABC[0].checked && (trimTrue(document.forms[0].txtC.value) == "")) {
		alert("Enter Classification C Value");
		document.forms[0].txtC.select();
		return false;
	}
	else if (document.forms[0].radABC[0].checked && (!checkNumbers(document.forms[0].txtC.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtC.select();
		return false;
	}
	else if (document.forms[0].radFSN[0].checked && (trimTrue(document.forms[0].txtFast.value) == "")) {
		alert("Enter Fast Moving in Days");
		document.forms[0].txtFast.select();
		return false;
	}
	else if (document.forms[0].radFSN[0].checked && (!checkNumbers(document.forms[0].txtFast.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtFast.select();
		return false;
	}
	else if (document.forms[0].radFSN[0].checked && (trimTrue(document.forms[0].txtSlow.value) == "")) {
		alert("Enter Slow Moving in days");
		document.forms[0].txtSlow.select();
		return false;
	}
	else if (document.forms[0].radFSN[0].checked && (!checkNumbers(document.forms[0].txtSlow.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtSlow.select();
		return false;
	}
	else if (document.forms[0].radFSN[0].checked && (trimTrue(document.forms[0].txtNon.value) == "")) {
		alert("Enter Non-Moving in days");
		document.forms[0].txtNon.select();
		return false;
	}
	else if (document.forms[0].radFSN[0].checked && (!checkNumbers(document.forms[0].txtNon.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtNon.select();
		return false;
	}
	else if (!(document.forms[0].chkFIFO.checked || document.forms[0].chkLIFO.checked || document.forms[0].chkWA.checked)) {
		alert("Select Accounting Type");
		document.forms[0].chkFIFO.focus();
		return false;
	}
	else {
		document.forms[0].action = "OrgInventoryControlInsert.asp"
		document.forms[0].submit();
	}
}

</SCRIPT>
<SCRIPT LANGUAGE=javascript>window.ITMS_STORAGE_APPS = "IN";</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/orgStorageDisplay.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="DisplayStoDet()">

<form method="POST" name="formname" action="">
<input type=hidden name="hOrgName" value="<%=sOrgName%>">
<input type=hidden name="hOrgCode" value="<%=sOrgCode%>">
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
		                    <td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgPurchaseControlEntry.asp">
											<td width="100%" align="center">Purchase
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">
                                                Inventory
											</td>
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
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
									    <tr>
											<td width="130" class="FieldCell">Stock Holding Period&nbsp;</td>
											<td class="FieldCell">
												<input type="text" name="txtStock" size="3" maxlength=3 class="FormElem">
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
								</td>
								<td valign="top">
	   							<center>
	                                   <div align="left">
	   								<table cellpadding="0" cellspacing="0" align="left">
	   									<tr>
	   										<td>
	   								<table cellpadding="0" cellspacing="0" width="100%">
	   									<tr>
	   										<td class='GroupTitleLeft' width="10">&nbsp;
	                                           </td>
	   										<td class='GroupTitle' width="115"><p align="center">Valuation Method
	                                           </td>
	   							</center>
	   										<td class='GroupTitleRight'><p align="left">&nbsp;
	                                           </td>
	   									</tr>
	   								</table>
	                                       </td>
	   									</tr>
	   									<tr>
	   										<td class=GroupTable>
	   							<center>
	                                   <div align="left">
	   								<table cellpadding="0" cellspacing="0">
	   									<tr>
	   										<td class=MiddlePack colspan="7"> </td>
	   									</tr>
	   									<tr>
	   										<td class=FieldCellSub width="100">ABC</td>
	   										<td class='FieldCellSub' width="120">
	                                           <input type="radio" value="1" name="radABC" class="FormElem"> Yes&nbsp;
	                                           <input type="radio" value="0" name="radABC" class="FormElem" checked onClick="document.formname.txtA.value = '';document.formname.txtB.value = '';document.formname.txtC.value = ''"> No
	                                        </td>
	   										<td class='FieldCellSub' width="28">A</td>
	   										<td class='FieldCellSub' width="40">
												<input type="text" name="txtA" size="8" maxlength=8 class="FormElem"></td>
	   										<td class='FieldCellSub' width="30">B</td>
	   										<td class='FieldCellSub' width="40">
												<input type="text" name="txtB" size="8" maxlength=8 class="FormElem"></td>
	   										<td class='FieldCellSub' width="30">C</td>
	   										<td class='FieldCellSub' width="40">
												<input type="text" name="txtC" size="8" maxlength=8 class="FormElem"></td>
	   										<td class='FieldCellSub' width="35">
	                                           </td>
	   									</tr>
	   									<tr>
	   										<td class=FieldCellSub width="100">FSN</td>
	   										<td class='FieldCellSub' width="120">
	                                           <input type="radio" value="1" name="radFSN" class="FormElem"> Yes&nbsp;
	                                           <input type="radio" value="0" name="radFSN" class="FormElem" checked onClick="document.formname.txtFast.value = '';document.formname.txtSlow.value = '';document.formname.txtNon.value = ''"> No
	                                        </td>
	   										<td class='FieldCellSub' width="28">Fast</td>
	   										<td class='FieldCellSub' width="40">
												<input type="text" name="txtFast" size="3" maxlength=3 class="FormElem"></td>
	   										<td class='FieldCellSub' width="30">Slow</td>
	   										<td class='FieldCellSub' width="40">
												<input type="text" name="txtSlow" size="3" maxlength=3 class="FormElem"></td>
	   										<td class='FieldCellSub' width="75">Non-Moving</td>
	   										<td class='FieldCellSub' width="35">
	                                           <input type="text" name="txtNon" size="3" maxlength=3 class="FormElem">&nbsp;</td>
	   									</tr>
	   							</center>
	   									<tr>
	   										<td class=FieldCellSub width="110">Accounting Type</td>
	   										<td class='FieldCellSub' width="403" colspan="7">
	                                           <input type="checkbox" name="chkFIFO" value="1">FIFO&nbsp;&nbsp;
	                                           <input type="checkbox" name="chkLIFO" value="1">LIFO&nbsp;&nbsp;
	                                           <input type="checkbox" name="chkWA" value="1">Weighted Average&nbsp;
	                                        </td>
	   									</tr>
	   								</table>
	                               </div>
	                                           </td>
	   									</tr>
	   								</table>
	                               </div><p>&nbsp;
	                                   <p>&nbsp;
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
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
                                        <tr>
											<td valign="top" rowspan="2">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0" width="200" align="left">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="45"><p align="center">Allow
                                                            </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="3"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub> </td>
															<td class='FieldCellSub'><p align="center">Yes</td>
															<td class='FieldCellSub'><p align="center">&nbsp;No </td>
															<td class='FieldCellSub'> </td>
														</tr>
														<tr>
															<td class=FieldCellSub> Replenishment</td>
															<td class='FieldCellSub' align="center">
                                                            <input type="radio" value="1" name="radRep" class="FormElem"></td>
															<td class='FieldCellSub' align="center">
                                                            <input type="radio" value="0" name="radRep" class="FormElem" checked></td>
															<td class='FieldCellSub' align="center">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCellSub> Inter Unit Transfer</td>
															<td class='FieldCellSub' align="center">
                                                            <input type="radio" value="1" name="radUnit" class="FormElem"></td>
															<td class='FieldCellSub' align="center">
                                                            <input type="radio" value="0" name="radUnit" class="FormElem" checked></td>
															<td class='FieldCellSub' align="center">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCellSub> Location Transfer</td>
															<td class='FieldCellSub' align="center">
                                                            <input type="radio" value="1" name="radLoc" class="FormElem"></td>
															<td class='FieldCellSub' align="center">
                                                            <input type="radio" value="0" name="radLoc" class="FormElem" checked></td>
															<td class='FieldCellSub' align="center">
                                                            </td>
														</tr>
												</center>
													</table>
                                                            </td>
														</tr>
													</table>
                                                    </div>
											</td>
											<td valign="top" rowspan="2" class="ClearPixel" width="5">
                                                <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td valign="top" class="ClearPixel" width="100%">
                                                <img border="0" src="../../assets/images/clearpixel.gif" width="10" height="10">
											</td>
                                    </tr>
                                        <tr>
											<td valign="top" width="100%" >
												<DIV class=frmBody id=frm3 style="width: 360;height=95">

                                                <table border="0" cellspacing="1" class="ExcelTable" Id ="tblLoc" name="tblLoc">
                                            <tr>
												<td class="ExcelHeaderCell" align="center">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="100%">Inventory Storage Location</td>
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

