<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmTypOptionPoPEntry.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	January 29, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	ItmTypeAttributeEntry.asp
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Attribute Options</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<script type="application/xml" id="OutDataValue" data-itms-xml-island="1">
<root></root>
</script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmTypeOption.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%
	dim arrTemp,sTypeName,sHdrName,sAttrName

	arrTemp = split(trim(Request.QueryString("sTemp")),":")

	sTypeName = arrTemp(0)
	sHdrName = arrTemp(1)
	sAttrName = arrTemp(2)
%>

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Attribute Options
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
								<td>
                                    <table border="0" cellspacing="0" cellpadding="4" class="TableOutlineOnly">
									    <tr>
											<td class="FieldCell">Type Name</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=sTypeName%>&nbsp;</span>
											</td>
									    </tr>
									    <tr>
											<td class="FieldCell">Header Name</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=sHdrName%>&nbsp;</span>
											</td>
									    </tr>
									    <tr>
											<td class="FieldCell">Attribute</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=sAttrName%>&nbsp;</span>
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
								<td align="center"></td>
								<td width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Name</td>
											<td class="FieldCellSub">
												<input type="text" name="txtValue" size="25" maxlength=30 class="FormElem">&nbsp;
												<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td width="100%">
									<div class="frmBody" id="frm2" style="width:100%;height:90;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="300">Name</td>
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
