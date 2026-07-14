<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmBoMEntry.asp
	'Module Name				:	Inventory (Asset Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 08, 2005
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
<!--#include virtual="/include/UoMDecimal.asp"-->
<%
' Declaration of variables
Dim dcrs,iCtr
'Declaration of Objects
iCtr = 0
Set dcrs = Server.CreateObject("ADODB.RecordSet")

dim arrTemp,arrUoM,sUoMCode,sUoM,sItemName,sClassName,sOrgID,sItemType,sOrgName

arrTemp = split(trim(Request.QueryString("sTemp")),":")

sClassName = arrTemp(0)
sItemName = arrTemp(1)
sItemType = arrTemp(2)
sOrgID = Session("organizationcode")
sOrgName = Session("OrgShortName")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - BoM Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmBoMEntry.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hItemCode" value="">
<input type=hidden name="hClassCode" value="">
<input type=hidden name="hItemType"  value="<%=sItemType%>">
<input type=hidden name="hConsumable" value="">
<input type=hidden name="hRowCtr" value="0">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Bill Of Materials
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
								<td width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
										<tr>
										    <td class="FieldCell" valign="top"> Select Item</td>
										    <td class="FieldCellSub">
												<span id="spSelItem" class="DataOnly" >&nbsp;</span>&nbsp;&nbsp;
												<a href="#" onclick="SelectItem()"><img src="../../assets/images/iTMS%20icons/Entry.gif" width=10 height=10></a>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Quantity</td>
											<td class="FieldCellSub">
												<input type="text" name="txtQty" size="12" maxlength=10 class="FormElem" style="text-align:right" onkeypress="DoKeyPress('Y',7,3)">&nbsp;&nbsp;<span class="DataOnly" id="idUoM">&nbsp;</span>
											</td>
										</tr>
                                        <tr>
                                            <td class="FieldCell">Consumable</td>
                                            <td class="FieldCellSub">
												<input type=Checkbox name="ChkConsumable" class="FormElem">
											</td>
                                        </tr>
										<tr>
											<td class="FieldCell">Item Type</td>
											<td class="FieldCell">
												<input type="radio" name="radType" value="F" class="FormElem" Checked>Final Component &nbsp;
												<input type="radio" name="radType" value="A" class="FormElem">Assembly &nbsp;
											</td>
										</tr>
										<tr>
											<td class="FieldCell"></td>
											<td class="FieldCell">
												<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
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
								<td width="100%">
									<div class="frmBody" id="frm2" style="width: 100%; height:130;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">
												    <img src="../../assets/images/iTMS%20icons/DeleteIcon.gif" onclick="DelItem()">
												</td>
												<td class="ExcelHeaderCell" align="center">Item</td>
												<td class="ExcelHeaderCell" align="center" width="100">Quantity</td>
												<td class="ExcelHeaderCell" align="center">Type</td>
												<td class="ExcelHeaderCell" align="center">Consumable</td>
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="window.close()">
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


