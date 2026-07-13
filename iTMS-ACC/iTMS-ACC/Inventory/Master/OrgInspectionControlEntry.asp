<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgInspectionControlEntry.asp
	'Module Name				:	Inventory (Organization Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 02, 2002
	'Modified By                :   Ragavendran R
	'Modified On				:   July 21,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	OrgInspectionControlInsert.asp
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
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="storageData" src="../xmldata/Storage.xml"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript >
function CheckSubmit(){
    if (!(document.forms[0].chkPO.checked || document.forms[0].chkOO.checked || document.forms[0].chkP.checked || document.forms[0].chkPS.checked)) {
		alert("Select Point Of Inspection");
		document.forms[0].chkPO.focus();
		return false;
	}
	else {
		document.forms[0].action = "OrgInspectionControlInsert.asp"
		document.forms[0].submit();
	}
}

</SCRIPT>
<SCRIPT LANGUAGE=javascript>window.ITMS_STORAGE_APPS = "POI,OI,PI,PSI";</SCRIPT>
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
								<td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgInventoryControlEntry.asp">
											<td width="100%" align="center">
                                                Inventory
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="85">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">
                                                Inspection
											</td>
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
								<td align="center">
								</td>
								<td valign="top">
									<table cellpadding="0" cellspacing="0" width="440">
										<tr>
											<td>
									<table cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class='GroupTitleLeft' width="10"><p align="left">&nbsp;</p>
                                            </td>
											<td class='GroupTitle' width="125"><p align="center">Point of Inspection
                                            </td>
											<td class='GroupTitleRight'><p align="left">&nbsp;
                                            </td>
										</tr>
									</table>
                                        </td>
										</tr>
										<tr>
											<td class=GroupTable>
                                <div align="left">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class=MiddlePack colspan="8"> <p align="left">  </td>
										</tr>
										<tr>
											<td class=FieldCellSub> <p align="left">
                                            <input type="checkbox" name="chkPO" value="1" class="FormElem"> </td>
											<td class='FieldCellSub'><p align="left">Pre-Order Sample </td>
											<td class='FieldCellSub'>
                                            <input type="checkbox" name="chkOO" value="1" class="FormElem"> </td>
											<td class='FieldCellSub'>Out Order </td>
											<td class='FieldCellSub'>
                                            <input type="checkbox" name="chkP" value="1" class="FormElem"> </td>
											<td class='FieldCellSub'>Process </td>
											<td class='FieldCellSub'>
                                            <input type="checkbox" name="chkPS" value="1" class="FormElem"> </td>
											<td class='FieldCellSub'>Post Sale </td>
										</tr>
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
								</td>
								<td valign="top">
									<DIV class=frmBody id=frm2 style="width: 440; height:64;">

                                    <table border="0" cellspacing="1" class="ExcelTable" Id ="tblLoc" name="tblLoc">
									    <tr>
											<td class="ExcelHeaderCell" align="center">S.No.</td>
											<td class="ExcelHeaderCell" align="center" width="100%">Inspection Storage Location</td>
									    </tr>
										<!--tr>
											<td class="ExcelSerial" align="center"></td>
											<td class="ExcelDisplayCell" width="100%"></td>
										</tr-->
                                    </table>
									</div>
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

