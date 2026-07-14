<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	itmRelateEntry.asp
	'Module Name				:	Inventory (Temporary Item Relation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 17, 2003
	'Modified By				:	TAJUDEEN S
	'Modified On				:	October 16, 2004
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	itmRelateInsert.asp
	'Procedures/Functions Used	:	populateUnits
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
	dim dcrs,iTempItmCode,sDesc,sAddDesc,sShDesc,sIType,sCode,sTempItmname
	'Declaration of Objects

	iTempItmCode = trim(Request.Form("selItem"))
	sTempItmname = trim(Request.Form("hTempItemname"))
	sIType = trim(Request.Form("hItemType"))

	if not iTempItmCode = "" then
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT GENITEMCODE,SHORTDESCRIPTION,ITEMDESCRIPTION,ADDITIONALDESCRIPTION,ITEMTYPEID FROM MS_TEMPORARYITEMMASTER WHERE FINALSTATUS = 'N' AND TEMPITEMCODE = " & iTempItmCode & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			sCode = trim(dcrs(0))
			sShDesc = trim(dcrs(1))
			sDesc = trim(dcrs(2))
			sAddDesc = trim(dcrs(3))
			sIType = trim(dcrs(4))
		end if
		dcrs.Close
	end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Temporary Item Relation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="ItemData" data-itms-xml-island="1"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/RelateItem.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="" onSubmit="return CheckSubmit('R')">
<input type=hidden name="hOrgId" value="<%=Session("organizationcode")%>">
<input type=hidden name="hItemType" value="<%=sIType%>">
<input type=hidden name="hTempItemCode" value="<%=iTempItmCode%>">
<input type=hidden name="hTempItemname" value="<%=Server.HTMLEncode(sTempItmname)%>">
<input type=hidden name="hItemCode" value="">
<input type=hidden name="hClassCode" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Temporary Item Relation with Exisiting</td>
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
                                            <td class="FieldCell">Temporary Item</td>
                                            <td class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sTempItmname%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Item Type</td>
                                            <td class="FieldCellSub" colspan="3">
												<select size="1" name="selIType" class="FormElem" disabled>
													<%	'Calling the Function which populates the Item Type list
														populateItemTypeSelected sIType
													%>
												</select>&nbsp;
												<input type="button" value="Search" class="AddButton" onClick="CheckItem()" id=button1 name=button1>
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
								<td align="center" width="5">
								</td>
								<td valign="top">
									<div class="frmBody" id="frm2" style="width: 100%; height:390;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Item Description</td>
												<td class="ExcelHeaderCell" align="center" width="100">Item Code</td>
												<td class="ExcelHeaderCell" align="center" width="80">Unit</td>
												<td class="ExcelHeaderCell" align="center" width="50">Stores UoM</td>
												<td class="ExcelHeaderCell" align="center" width="20">Relate</td>
                                            </tr>
                                        </table>
                           			</div>
								</td>
								<td align="center">
								</td>
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
													<input type="button" value="Create New" name="B1" class="ActionButtonX" onClick="CheckSubmit('C')" >
                                                    <input type="button" value="Relate" name="B1" class="ActionButton" onClick="CheckSubmit('R')">
                                                    <input type="button" value="Reset" name="B2" class="ActionButton" onclick="ClearAll()">
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

