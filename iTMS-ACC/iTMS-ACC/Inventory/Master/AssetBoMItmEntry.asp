<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AssetBoMItmEntry.asp
	'Module Name				:	Inventory (Asset Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 17, 2004
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<%
' Declaration of variables
Dim dcrs,iCtr
'Declaration of Objects
iCtr = 0
Set dcrs = Server.CreateObject("ADODB.RecordSet")

dim arrTemp,arrUoM,sUoMCode,sUoM,sItemName,sClassName,sOrgID,sItemType

arrTemp = split(trim(Request.QueryString("sTemp")),":")

sOrgID = arrTemp(0)
sClassName = arrTemp(1)
sItemName = arrTemp(2)

sItemType = "STO"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - BoM Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data">
<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData">
<Root/>
</script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/AssetBoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="fnInit()">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
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
										    <td class="FieldCell" valign="top"> Select Classification</td>
										    <td class="FieldCellSub">
												<select size="1" name="selClass" class="FormElem" onChange="GetItem(this)">
													<option value="select">Select</option>
											<%	'Calling the Function which populates Classification List
												populateClassification
											%>
												</select>
											</td>
										</tr>
										<tr>
										    <td class="FieldCell" valign="top"> Select Item</td>
										    <td class="FieldCellSub">
												<select size="5" name="selItem" class="FormElem" onChange="DisplayUoM(this)">
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Quantity</td>
											<td class="FieldCellSub">
												<input type="text" name="txtQty" size="12" maxlength=10 class="FormElem" style="text-align:right" onkeypress="DoKeyPress('Y',7,3)">
											</td>
										</tr>
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idUoM">&nbsp;</span>
											</td>
                                        </tr>
										<tr>
											<td class="FieldCell">Item Type</td>
											<td class="FieldCell">
												<input type="radio" name="radType" value="F" class="FormElem">Final Component &nbsp;
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
												<td class="ExcelHeaderCell" align="center">Item / Classification</td>
												<td class="ExcelHeaderCell" align="center" width="100">Quantity</td>
												<td class="ExcelHeaderCell" align="center">UoM</td>
												<td class="ExcelHeaderCell" align="center">Type</td>
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

<%
	' Function to populate the Classification list
	Function populateClassification()
		' Declaration of variables
		Dim dcrs,sClass,sClassName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT DISTINCT CLASSIFICATIONCODE,GROUPNAME FROM VWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND ITEMTYPEID = " & Pack(sItemType) & " ORDER BY 2"
			.Source = "SELECT DISTINCT CLASSIFICATIONCODE,GROUPNAME FROM VW_INV_ITEMS WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND ITEMTYPEID = " & Pack(sItemType) & " ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sClass = dcrs(0)
		set sClassName = dcrs(1)

		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sClass)&""">"&trim(sClassName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
