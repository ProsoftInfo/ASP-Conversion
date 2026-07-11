<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	CostCenterDet.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Senthil E
	'Modified By				:	Manohar Prabhu.R
	'Created On					:	July 22,2003
	'Modified On				:	Nov 28, 2003
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!--#include file="../../include/populate.asp"-->
<%
	dim sHeadCode,sHeadName,sHeadShortName,objRs2,sAmenType,sAmenTxt,oDOM,sCallFrom
	dim sGCode,sGName,sGFlag
	dim objRs,sQuery
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs2 = Server.CreateObject("ADODB.Recordset")

	sCallFrom = Trim(Request.Form("hCallFrom"))
	'Response.Write "<p>sCallFrom="&sCallFrom
	If sCallFrom="GP" Then
		sGCode = trim(Request.Form("GCode"))
		sGName = trim(Request.Form("GName"))
		sGFlag = trim(Request.Form("GroupFlag"))
	Else
		sHeadCode = trim(Request("hHeadValue"))
		sHeadName = trim(Request("hHeadName"))
	End IF


if not (isNull(sHeadCode) or isEmpty(sHeadCode) or sHeadCode = "") and sCallFrom = "HD" then

	sQuery ="select CCHeadCode from VwOrgCostCenter where  CostCenterHead="&sHeadCode

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	sHeadShortName=objRs(0)
	objRs.Close
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Cost Center Group</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<script>
//For Head
function HeadUpdate()
{
	if (trim(document.formname.txtClassName.value)=="" )
	{
		alert ("Enter Head Name");
		document.formname.txtClassName.focus();
		return false;
	}
	if (trim(document.formname.txtShortName.value)=="" )
	{
		alert ("Enter Head Short Name");
		document.formname.txtShortName.focus();
		return false;
	}
	document.formname.submit();
}

//For Group
function GrpUpdate()
{
	if (trim(document.formname.txtClassName.value)=="" )
	{
		alert ("Enter Group Name");
		document.formname.txtClassName.focus();
		return false;
	}

	document.formname.submit();

}

</script>
<script>
window.__itmsPopupCompat = { type: "costCenterDetails" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
<script src="../../scripts/itms-modern-compat.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<%
If sCallFrom = "HD" Then
	if not (isNull(sHeadCode) or isEmpty(sHeadCode) or sHeadCode = "") then
%>

<form method="POST" name="formname" action="CostCenterDetUpdateCommon.asp" TARGET="bodyFrame">

<INPUT type=hidden value="<%=sHeadCode%>" name="hHeadValue">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
	<table border="0" cellspacing="0" width="100%" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
	<td class="ExcelHeaderCell" colspan="2">Edit&nbsp; Cost Center </td>
	</tr>
	<tr>
	<td width="5">
	</td>
	<td>
	<center>
	<table cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td class="MiddlePack" colspan="2"> </td>
		</tr>
			<tr>
			<td class="FieldCellSub" > Name</td>
			<td class="FieldCellSub">
				<input type="text" name="txtClassName" value="<%=sHeadName%>" size="30" maxlength="40" class="FormElem">
			</td>
		</tr>
		<tr>
			<td class="FieldCellSub" > Short Name</td>
			<td class="FieldCellSub">
				<input type="text" name="txtShortName" size="11" maxlength="10" value="<%=sHeadShortName%>" class="FormElem">
			</td>
		</tr>

		<tr>
			<td class="MiddlePack" colspan="2"> </td>
		</tr>
		<tr>
			<td class="ActionCell" colspan="2">
	        <input type="Button" value="Update" name="B2" onClick="HeadUpdate()" class="ActionButton">
	        <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
		</tr>
	</table>
	</center>

</form>
<%End if

ElseIf sCallFrom = "GP" Then

	if not (isNull(sGCode) or isEmpty(sGCode) or sGCode = "") then	'GroupEditUpdate.asp
	%>

<form method="POST" name="formname" action="CostCenterDetUpdateCommon.asp" TARGET="bodyFrame">
<input type="hidden" name="hGroupCode" value="<%=sGCode%>">
<input type="hidden" name="hGroupFlag" value="<%=sGFlag%>">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

	<table border="0" cellspacing="0" width="100%" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="ExcelHeaderCell" colspan="2">Edit Group Name</td>
	</tr>
	<tr>
		<td width="10" colspan="2" class="MiddlePack">	</td>
	</tr>
	<tr>
	<td width="5">	</td>
	<td>
	<table cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td class="MiddlePack" colspan="2"> </td>
		</tr>
			<tr>
			<td class="FieldCellSub" width="100px"> Group Name</td>
			<td class="FieldCellSub">
				<input type="text" name="txtClassName" size="30" maxlength=40 value="<%=sGName%>" class="FormElem">
			</td>
		</tr>
		<tr>
			<td class="MiddlePack" colspan="2"> </td>
		</tr>
		<tr>
			<td class="ActionCell" colspan="2">
	        <input type="Button" value="Save" name="B2" onClick="GrpUpdate()" class="ActionButton">
	        <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
		</tr>
	</table>
	</td>
	</tr>
    </table>
</form>
<%end if
End IF	'If sCallFrom = "HD" Then%>
</BODY>
</html>
