<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgControlDefn.asp
	'Module Name				:	Inventory (Organization Control Definition)
	'Author Name				:	Ragavendran R
	'Created On					:	July 21,2011
	'Modified By                :   
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
	dim dcrs,sSql,sLocName,iCtr
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	iCtr = 1
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
<SCRIPT LANGUAGE=javascript SRC="../scripts/orgControlDefn.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hOrgIndex" value="">
<input type=hidden name="hOrgName" value="">
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
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Org Defn
											</td>
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
		                   <td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgPurchaseControlEntry.asp">
											<td width="100%" align="center">
                                                Purchase
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
					<TD class="TabBody">
					    <div style="height:300px"
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
