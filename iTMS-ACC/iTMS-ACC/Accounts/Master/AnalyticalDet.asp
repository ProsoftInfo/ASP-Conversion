<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AnalyticalDet.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	July 22,2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	March 21, 2011
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
	Dim sHeadCode,sHeadName,sHeadShortName,sAmenType,sAmenTxt,sDelOpt,sParentCode
	dim sGCode,sGName,sGFlag,objRs,sQuery,Objrs2,sCallFrom

	sCallFrom = Trim(Request("hCallFrom"))

	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs2 = Server.CreateObject("ADODB.Recordset")

	If sCallFrom = "GP" Then

		sGCode = trim(Request.Form("GCode"))
		sGName = trim(Request.Form("GName"))
		sGFlag = trim(Request.Form("GroupFlag"))

	Elseif sCallFrom = "HD" Then

		sHeadCode = trim(Request("hHeadValue"))
		sHeadName = trim(Request("hHeadName"))
		sParentCode = trim(Request("hGroupValue"))

		if not (isNull(sHeadCode) or isEmpty(sHeadCode) or sHeadCode = "") then

			sQuery ="select AnalyticalShortName from VwOrgAnalytical where  AnalyticalCode="&sHeadCode
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

	End IF	'If sCallFrom = "GP" Then

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Analytical Group</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<script>
function AnalGroupUpdate()
{
	if (trim(document.formname.txtClassName.value)=="" )
	{
		alert ("Enter Group Name");
		document.formname.txtClassName.focus();
		return false;
	}

	document.formname.submit();
}

function AnalHeadUpdate()
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
	//document.formname.hCallType.value = sTemp
	document.formname.B2.disabled = true
	document.formname.submit();
}

</script>
<script src="../../scripts/itms-modern-compat.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<%	If  sCallFrom = "GP" Then

		if not (isNull(sGCode) or isEmpty(sGCode) or sGCode = "") then	'GroupEditUpdate.asp
		%>

		<form method="POST" name="formname" action="AnalyticalDetUpdateCommon.asp" TARGET="bodyFrame">
		<input type="hidden" name="hGroupCode" value="<%=sGCode%>">
		<input type="hidden" name="hGroupFlag" value="<%=sGFlag%>">
		<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

		<table border="0" cellspacing="0" width="100%" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="ExcelHeaderCell" colspan="2">Edit Group Name</td>
		</tr>
		<tr>
			<td width="10px" colspan="2px" class="MiddlePack">	</td>
		</tr>
		<tr>
		<td width="5px">	</td>
		<td>
		<table cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td class=MiddlePack colspan="2px"> </td>
			</tr>
				<tr>
				<td class=FieldCellSub width="100px"> Group Name</td>
				<td class="FieldCellSub">
					<input type="text" name="txtClassName" size="30" maxlength="40" value="<%=sGName%>" class="FormElem">
				</td>
			</tr>
			<tr>
				<td class=MiddlePack colspan="2"> </td>
			</tr>
			<tr>
				<td class=ActionCell colspan="2">
		        <input type="Button" value="Save" name="B2" onClick="AnalGroupUpdate()" class="ActionButton">
		        <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
			</tr>
		</table>
		</td>
		</tr>
	    </table>
	</form>
	<%
		end if	'if not (isNull(sGCode) or isEmpty(sGCode) or sGCode = "") then

	Else

		if not (isNull(sHeadCode) or isEmpty(sHeadCode) or sHeadCode = "") then	'ANALEditUpdate.asp
	%>
	<form method="POST" name="formname" action="AnalyticalDetUpdateCommon.asp" TARGET="bodyFrame">

		<INPUT type=hidden value="" name="hCallType">
		<INPUT type=hidden value="<%=sHeadCode%>" name="hHeadValue">
		<INPUT type=hidden value="<%=sParentCode%>" name="hParentCode">
		<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

	<table border="0" cellspacing="0" width="100%" cellpadding="0">
	<tr>
	<td class="ExcelHeaderCell" colspan="2"><p align="center">Edit&nbsp; Analytical Head</td>
	</tr>
	<tr>
	<td width="5px">
	</td>
	<td>
	<center>
	<table cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td class=MiddlePack colspan="2"> </td>
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
			<td class=ActionCell colspan="2">
	        <input type="Button" value="Update" name="B2" onClick="AnalHeadUpdate()" class="ActionButton">
	        <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
		</tr>
	</table>
	</center>

</form>
<%
	end if 'if not (isNull(sHeadCode) or isEmpty(sHeadCode) or sHeadCode = "") then

End IF 'If sCallFrom = "GP" Then
%>
</BODY>
