<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmAccHead.asp
	'Module Name				:	Inventory
	'Author Name				:	Ragavendran R
	'Created On					:	
	'Modified By				:	
	'Tables Used				:	
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	
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
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/GetSerialDetail.asp" -->

<%
Dim iItemCode
iItemCode = Request.QueryString("ItemCode")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Item Account Head</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<base target="_self" />
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script language="javascript">
<!--
function DoUpload() {
	theFeats = "height=120,width=500,location=no,menubar=no,resizable=no,scrollbars=no,status=no,toolbar=no";
	theUniqueID = (new Date()).getTime() % 1000000000;
	window.open("../../Common/progressbar.asp?ID=" + theUniqueID, theUniqueID, theFeats);
	document.formname.action = "ItmUploadImageInsert.asp?ID=" + theUniqueID;
	document.formname.submit();
}
//-->
</script>
</HEAD>
<BODY>
<form method="POST" name="formname" action="ItmUploadImageInsert.asp" enctype="multipart/form-data">
<input type="hidden" name="hItemCode" value="<%=iItemCode%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Upload Image
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top" class=TabBodyWithTopLine>
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD >
						<table cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="FieldCellSub">Thumbnail</td>
								<td class="FieldCellSub">
									<input type="file" name="imgThumb" class="formelem" />
								</td>
							</tr>
							<tr>
								<td class="FieldCellSub">Blowup</td>
								<td class="FieldCellSub">
									<input type="file" name="imgBlow" class="formelem" />
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
                    <td class="actioncell" align="center">
                        <input type="button" name="btnDone" value="Done" class="ActionButton" onclick="DoUpload()" />
                    </td>
                </tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
