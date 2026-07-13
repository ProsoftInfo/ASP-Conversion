<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasClassificationNameAmendEntry.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 22, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	MasClassificationNameUpdate.asp
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Classification</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/masClassificationUpdate.js"></SCRIPT>
<SCRIPT>
function Init() {
	document.formname.txtClassName.value = document.formname.hClassName.value.split("``").join('"');
}
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init()">
<% 	
	dim iClass,sClassName,arrTemp,sPara
	arrTemp = split(trim(Request.Form("pGroup")),":")
	sClassName = trim(Request.Form("pName"))
	iClass = arrTemp(1)
	sPara =  Request("hPara")
%>

<form method="POST" name="formname" action="" target="bodyFrame">
<input type=hidden name=hItmType value="">
<input type=hidden name=hPara value="<%=sPara%>">
	<table border="0" cellspacing="0" width="100%" cellpadding="0">
		<tr>
			<td class="ExcelHeaderCell" colspan="3"><p align="center">Classification Amendment</td>
		</tr>
		<tr>
			<td width="10" colspan="3" class="MiddlePack"></td>
		</tr>
		<tr>
			<td width="5"></td>
			<td>
				<table cellpadding="0" cellspacing="0" width="100%" border="0">
					<tr>
						<td>
							<table cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td class='GroupTitleLeft' width="10">&nbsp;</td>
									<td class='GroupTitle' width="60"><p align="center">Details</td>
									<td class='GroupTitleRight'><p align="left">&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class=GroupTable>
							<table cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td class=MiddlePack colspan="2"> </td>
								</tr>
								<tr>
									<td class=FieldCellSub width="62"> Name</td>
									<td class='FieldCellSub'>
										<input type="text" name="txtClassName" size="30" value="" maxlength=40 class="Formelem">
										<INPUT type=hidden value="<%=replace(sClassName,"""","``")%>" name="hClassName">
									</td>
								</tr>
								<tr>
									<td class=MiddlePack colspan="2"> </td>
								</tr>
								<tr>
									<td class=ActionCell colspan="2"> <p align="center">
								    <input type="button" value="Amend" name="B2" class="ActionButton" onClick="javascript:CheckSubmit()">
								    <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
								</tr>
							</table>
						</td>
						<td width="5"></td>
					</tr>
				</table>
				<INPUT type=hidden value="<%=iClass%>" name="hpGroup">
			</td>
		</tr>
	</table>
</form>
</BODY>
</HTML>
