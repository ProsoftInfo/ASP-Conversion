<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	InventoryNoSeriesAmendEntry.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 20, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	InventoryNoSeriesAmendInsert.asp
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
Dim ObjRs
Dim sQuery
set ObjRs = server.CreateObject("ADODB.Recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Inventory Number Series Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData"><ROOT/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/inventoryNoSeriesAmend.js"></SCRIPT>


</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="InventoryNoSeriesAmendInsert.asp">
<input type=hidden name="hSeriesType" value="">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hActivityName" value="">
<input type="hidden" name="hClassCode" value="">
<input type="hidden" name="hCatCode" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">No. Series Allocation Amendment
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
					<TD class=TabBodyWithTopLine>
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
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class=FieldCell> Select Unit</td>
											<td class="FieldCellSub">
												<select size="1" name="selUnit" class="FormElem" onChange="document.getElementById('selCategory').selectedIndex = 0; document.formname.selActType.selectedIndex = 0; document.formname.hClassCode.value=''; document.formname.hCatCode.value='';">
													<option value="select">Select</option>
													<%	'Calling the Function which populates Organization Unit list
														populateUnit
													%>
												</select>
											</td>
										</tr>
										<!--<tr>
											<td class="FieldCell">Item Type</td>
											<td class="FieldCellSub">
											    <select size="1" name="selItmType" class="FormElem" onChange="document.formname.selActType.selectedIndex = 0">
													<option value="select">Select</option>
													<%	'Calling the Function which populates the Item Type list
														'populateItemTypeWOCap
													%>
												</select>
											</td>
										</tr>-->
										<tr>
											<td class="FieldCell">Classification</td>
											<td class="FieldCellSub">
											    <span id="txtClass" class="DataOnly">&nbsp;</span>&nbsp;&nbsp;<a href="#" onclick="SelectClassifcation()"><img style="cursor: hand" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Classification"></a>
											        <select id="selCategory" class="FormElem">
											        <option value="0">Select Category</option>
											        <%
											            sQuery = "Select CategoryCode,CategoryName from Inv_M_ClassificationCategory Order By CategoryCode"
											            Objrs.open sQuery,con
											            if not Objrs.eof then
											                do while not Objrs.eof
											                    Response.Write "<option value="& Objrs(0) &">"& Objrs(1) &"</option>"
											                    Objrs.movenext
											                loop
											            end if
											            Objrs.close
											        %>
											    </select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Activity</td>
											<td class="FieldCellSub">
											    <select size="1" name="selActType" class="FormElem" onChange="DisplayTable()">
													<option value="select">Select</option>
													<option value="LO">Lot Number</option>
													<option value="MR">MR Number</option>
													<!--option value="IR">Internal Receipts</option>
													<option value="ER">External Receipts</option-->
													<option value="IS">Issue Number</option>
													<option value="PN">Packing Number</option>
													<option value="SL">Sample Label</option>
													<option value="DC">DC - Gate Pass</option>
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
											<td class="ExcelHeaderCell" align="center" width="100">Period</td>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<input type="button" value="Amend" name="B4" class="ActionButton" onClick="validateForm()" >
															<input type="Reset" value="Cancel" name="B2" class="ActionButton" onClick="ClearAll()">
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


