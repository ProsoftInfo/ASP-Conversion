<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasClassificationAttributeEntry.asp	
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 18, 2002
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	MasClassificationAttributeInsert.asp
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Classification Attribute</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/masClassificationAttrCreate.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<% 	dim spGroup,spName,sgPath
	spGroup = trim(Request.Form("pGroup"))
	spName = trim(Request.Form("pName"))
	sgPath = trim(Request.Form("gPath"))
	if not (isNull(spGroup) or isEmpty(spGroup) or spGroup = "") then
%>
<form method="POST" name="formname" action="">
	<table border="0" cellspacing="0" width="100%" cellpadding="0">
	<tr>
        <td class="ExcelHeaderCell" colspan="3"><p align="center">Define Attributes</td>
            </tr>
            <tr>
        <td width="10" colspan="3" class="MiddlePack"></td>
        </tr>
        <tr>
        <td width="5"></td>
        <td>
			<center>
				<table cellpadding="0" cellspacing="0" width="100%" border="0">
					<tr>
						<td>
				<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td class='GroupTitleLeft' width="10">&nbsp;</td>
						<td class='GroupTitle' width="<%=len(spName)+100%>"><p align="center"><%=spName%></td>
			</center>
			<td class='GroupTitleRight'><p align="left">&nbsp;</td>
					</tr>
				</table>
		</td>
		</tr>
		<tr>
			<td class=GroupTable>
			<center>
			<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td class=MiddlePack colspan="2"> </td>
				</tr>
				<tr>
					<td class=FieldCellSub> Name</td>
					<td class='FieldCellSub'><input type="text" name="txtAttrName" size="30" maxlength=30 class="Formelem"></td>
				</tr>
				<tr>
					<td class=FieldCellSub> Data type</td>
					<td class='FieldCellSub'>
						<select size="1" name="selDataType" class="FormElem" onChange="return(checkSelect())">
							<OPTION value = "select">Select</OPTION>
							<OPTION value = "String">String</OPTION>
							<OPTION value = "Numeric">Numeric</OPTION>
						</select>
                    </td>
				</tr>
					</center>
				<tr>
					<td class=FieldCellSub> Length</td>
					<td class='FieldCellSub'><input type="text" name="txtDataLen" size="5" maxlength=3 class="Formelem"></td>
				</tr>
				<tr>
					<td class=FieldCellSub> Decimals</td>
					<td class='FieldCellSub'><input type="text" name="txtDecimal" size="5" maxlength=3 class="Formelem"></td>
				</tr>
				<tr>
					<td class=MiddlePack colspan="2"> </td>
				</tr>
				<tr>
					<td class=ActionCell colspan="2"> <p align="center">
				    <input type="button" value="Save" name="B4" class="ActionButton" onClick="javascript:CheckSubmit()">
				    <input type="reset" value="Reset" name="B5" class="ActionButton"></td>
				</tr>
			</table>
            </td>
		</tr>
	</table>
        </td>
        <td width="5"></td>
        </tr>
        <tr>
			<td width="10" colspan="3" class="BottomPack">
        </td>
        </tr>
            </table>
			</td>
            </tr>
		<tr>
			<td align="center" class="MiddlePack" width="100%" colspan="3">
				<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
			</td>
		</tr>
	</table>
	<INPUT type=hidden value="<%=spGroup%>" name=hpGroup>
	<INPUT type=hidden value="<%=sgPath%>" name=hgPath>
</form>
<%	end if %>
</BODY>
</HTML>