<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ANALGroupNameEntry.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Senthil E
	'Created On					:	January 19,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	CCGroupNameEntryInsert.asp
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
<HTML><HEAD><TITLE>Anlytical Group</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<script>
function SaveAnalGroup()
{
	if (trim(document.formname.txtClassName.value)=="" )
	{
		alert ("Enter Group Name");
		document.formname.txtClassName.focus();
		return false;
	}
	document.formname.B2.disabled = true
	document.formname.submit();
}


function SaveAnalHead()
{
	var sUnit,bFlag
	if (trim(document.formname.txtClassName.value)=="" )
	{
		alert ("Enter Analytical Name");
		document.formname.txtClassName.focus();
		return false;
	}
	if (trim(document.formname.txtShortName.value)=="" )
	{
		alert ("Enter Analytical Short Name");
		document.formname.txtShortName.focus();
		return false;
	}
	document.formname.B2.disabled = true
	document.formname.submit();
}
</script>
<script>
window.__itmsPopupCompat = { type: "analGroupAndHeadCreation" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
<script src="../../scripts/itms-modern-compat.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%
	Dim objRs
	Dim spClass,sCallFrom ,sUnitID,sQuery,sAnalName,sAnalShortName,sAnalCode

	Set objRs = Server.CreateObject("ADODB.RecordSet")

	spClass = trim(Request.Form("GCode"))
	sCallFrom = Trim(Request.QueryString("Data"))
	sUnitID = Trim(Session("organizationcode"))

	If sCallFrom = "GP" Then

	if not (isNull(spClass) or isEmpty(spClass) or spClass = "") then	'ANALGroupNameInsert.asp
%>

<form method="POST" name="formname" action="AnalGroupAndHeadNameInsert.asp" TARGET="bodyFrame">
	<table border="0" cellspacing="0" width="100%" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
	<td class="ExcelHeaderCell" colspan="3">Define New Group Name</td>
	</tr>
	<tr>
	<td width="10px" colspan="3" class="MiddlePack">
	</td>
	</tr>
	<tr>
	<td width="5ps">
	</td>
	<td>
	<center>
	<table cellpadding="0" cellspacing="0" width="100%" border="0">
		<tr>
			<td>
	<table cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td class="GroupTitleLeft" width="10px">&nbsp;
	        </td>
			<td class="GroupTitle" width="128px">Group Details
	        </td>
	</center>
			<td class="GroupTitleRight">&nbsp;
	        </td>
		</tr>
	</table>
	    </td>
		</tr>
		<tr>
			<td class="GroupTable">
	<center>
	<table cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td class="MiddlePack" colspan="2"> </td>
		</tr>
			<tr>
			<td class="FieldCellSub" width="62px"> Name</td>
			<td class="FieldCellSub">
				<input type="text" name="txtClassName" size="30" maxlength="40" class="FormElem">
			</td>
		</tr>
		<tr>
			<td class=MiddlePack colspan="2"> </td>
		</tr>
		<tr>
			<td class="ActionCell" colspan="2">
	        <input type="Button" value="Save" name="B2" onClick="SaveAnalGroup()" class="ActionButton">
	        <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
		</tr>
	</table>
	        </td>
	<td width="5px">
	</td>
		</tr>
	</table>
	<INPUT type=hidden value="<%=spClass%>" name="hpGroup">
	<INPUT type=hidden value="<%=sCallFrom%>" name="hCallFrom">
</form>
<%
	end if

Elseif sCallFrom="HD" Then

	if not (isNull(spClass) or isEmpty(spClass) or spClass = "") then	'ANALNameInsert.asp
%>

<form method="POST" name="formname" action="AnalGroupAndHeadNameInsert.asp" TARGET="bodyFrame">
<INPUT type=hidden value="<%=spClass%>" name="hParentCode">
<INPUT type=hidden value="<%=sUnitID%>" name="hUnitID">
<INPUT type=hidden value="<%=sCallFrom%>" name="hCallFrom">
	<table border="0" cellspacing="0" width="100%" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
	<td class="ExcelHeaderCell" colspan="2">Define Analytical Head</td>
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
			<td class="FieldCellSub" > Existing Head</td>
			<td class="FieldCellSub">
				<select class="FormElem" name="selAnalHead" onChange="setAnalName()">
				<option value="0?0">New</option>
<%
	sQuery = " SELECT distinct AnalyticalName ,AnalyticalShortName,AnalyticalCode FROM Acc_M_AnalyticalHeads where "&_
			 " AnalyticalCode not in (SELECT AnalyticalCode from Acc_R_OrgAnalyticalHeads where AHGroupCode='"&spClass &"') order by AnalyticalName"
	'Response.Write sQuery

	with objRs
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	set sAnalShortName=objRs("AnalyticalShortName")
	set sAnalName=objRs("AnalyticalName")
	set sAnalCode=objRs("AnalyticalCode")

	while Not objRs.EOF
		Response.Write "<option value="""&sAnalCode&"?"&sAnalShortName&""">"&sAnalName&"</option>"
		objRs.MoveNext
	wend
%>
				</select>
			</td>
		</tr>
			<tr>
			<td class="FieldCellSub" > Name</td>
			<td class="FieldCellSub">
				<input type="text" name="txtClassName" size="30" maxlength=40 class="FormElem">
			</td>
		</tr>
		<tr>
			<td class="FieldCellSub" > Short Name</td>
			<td class="FieldCellSub">
				<input type="text" name="txtShortName" size="11" maxlength=10 class="FormElem">
			</td>
		</tr>
		<tr>
			<td class=MiddlePack colspan="2"> </td>
		</tr>
		<tr>
			<td class="ActionCell" colspan="2">
	        <input type="Button" value="Save" name="B2" onClick="SaveAnalHead()" class="ActionButton">
	        <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
		</tr>
	</table>
	</center>

</form>
<%	end if

End IF	'If sCallFrom = "GP" Then%>
</BODY>
</HTML>


