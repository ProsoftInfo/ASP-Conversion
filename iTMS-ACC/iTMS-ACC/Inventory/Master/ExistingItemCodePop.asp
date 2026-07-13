<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ExistingItemCodePop.asp
	'Module Name				:	Inventory (Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 03, 2004
	'Modified By				:	TAJUDEEN S
	'Modified On				:	July 22, 2004
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
<%
	dim sType,arrTemp,sItmType

	arrTemp = split(trim(Request.QueryString("sTemp")),":")
	sType	= arrTemp(0)
	sItmType = arrTemp(1)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Existing Item Codes</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data"><root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/existingItemCodePop.js"></SCRIPT>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type=hidden name=hItmType value="<%=sType%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Existing Item Code Details
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
                                            <td class="FieldCell">Item Type</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly"><%=sItmType%>&nbsp;</span>
                                            </td>
                                        </tr>
										<!--tr>
                                            <td class="FieldCell">Display By</td>
                                            <td class="FieldCell">
												<input type="radio" value="A" name="radDisplay" class="FormElem" onclick="DisableTxtQty(this)">  All Items
                                            </td>
                                        </tr-->
										<input type="hidden" value="A" name="radDisplay">
										<tr>
                                            <td class="FieldCell"></td>
                                            <td class="FieldCell">
												<input type="radio" value="I" name="radDisplay" class="FormElem" onclick="DisableTxtQty(this)">  Search By&nbsp;&nbsp;
												<select size="1" name="selSearchBy" class="FormElem" onChange= "document.formname.txtSearchFor.value=''" disabled>
													<option value="select">Select</option>
													<option value="IC">Item Code</option>
													<option value="IN">Item Name < Starts with ></option>
													<option value="IA">Item Name < Anywhere ></option>
													<option value="CL">Classification</option>
													<option value="DN">Drawing No.</option>
													<option value="CN">Catalogue No.</option>
													<option value="MN">MGR No.</option>
													<option value="PN">Page No.</option>
													<option value="PO">Position No.</option>
												</select>
                                            </td>
                                        </tr>
   										<tr>
                                            <td class="FieldCell"></td>
                                            <td class="FieldCell">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Search For&nbsp;
												<input type="text" name="txtSearchFor" size="20" maxlength=15 class="FormElem">&nbsp;&nbsp;
												<input type="button" value="Search" name="btnSearch" class="AddButtonX" onClick="SearchItem()" disabled>
                                            </td>
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
								<td>
									<div class="frmBody" id="frm2" style="width: 100%; height:140;">
										<table border="0" cellspacing="1" id="tblDetails" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Item Code</td>
												<td class="ExcelHeaderCell" align="center">Description</td>
												<td class="ExcelHeaderCell" align="center">Drawing No.</td>
												<td class="ExcelHeaderCell" align="center">Catalogue No.</td>
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
                                                    <input type="button" value="Close" name="B1" class="ActionButton" onClick="window.close()">
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
