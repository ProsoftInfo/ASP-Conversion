<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GLAccHeadTree.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 21,2011
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
<!--#include file="../../include/DatabaseConnection.asp."-->
<!--#include file="../../include/populate.asp"-->
<!-- #include File="../../include/GetSettings.asp" -->
<%
dim sIP,sOrgid
sIP = GetSettings("IP")
sOrgid = Session("organisationcode")
 %>
<HTML>
<HEAD>
<TITLE>iTms Group-Head List View</TITLE>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="text/javascript" src="../../scripts/itms-modern-compat.js"></script>
<script type="text/javascript">
function getGroupHeadList() {
	return document.formname.ctlGroupHeadList;
}

function RefreshData() {
	var form = document.formname;
	var orgId = String(form.hOrgCode.value || "").trim();
	var tree = getGroupHeadList();
	if (orgId && tree) {
		tree.DSN = "http://" + form.hIP.value + "/Accounts/components/GetACCGroup.asp?flag=B&orgid=" + encodeURIComponent(orgId);
	}
}

function CreateNewHead() {
	document.formname.action = "GLCreate_Edit_AccHeadDet.asp";
	document.formname.submit();
}

function EditAccHead() {
	var form = document.formname;
	var tree = getGroupHeadList();
	if (tree && Number(tree.HeadValue) > 0) {
		form.hHeadValue.value = tree.HeadValue;
		form.hHeadName.value = tree.HeadName;
		form.action = "GLCreate_Edit_AccHeadDet.asp";
		form.submit();
	} else {
		alert("Select Head");
		if (tree && tree.focus) {
			tree.focus();
		}
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="RefreshData()">
<form method="POST" name="formname">
<input type="hidden" name="hIP" value="<%=sIP%>">
<input type="hidden" name="hHeadValue" value="0">
<input type="hidden" name="hHeadName" value="">
<input type="hidden" name="hOrgCode" value="<%=sOrgID%>">

<table border="0" cellspacing="0" cellpadding="0" width=100%>
<tr>
	<td align="center" class=PageTitle height="20"><p align="center">GL Account Head
	</td>
</tr>
<tr>
	<td align="center" class="TopPack">
	</td>
</tr>
<tr>
<td>
<div id="ctlGroupHeadList" data-itms-tree-control data-width="263px" data-height="353px"
	data-dsn="http://<%=sIP%>/Accounts/components/GetEditACCGroup.asp?flag=B&orgid=<%=sOrgid%>"
	data-list-name="ACCOUNT GROUPS"></div>
</td>
</tr>
<tr>
<tr>
	<td align="center" class="TopPack">
	</td>
</tr>
<td valign="top">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
	    <tr>
			<td valign="middle" class="ActionCell" align=center>
			    <input type="button" value="Create New" onclick="CreateNewHead()" class="ActionButtonX" name="btn1" >
                &nbsp;&nbsp;&nbsp;<input type="button" value="Edit" onclick="EditAccHead()" class="ActionButtonX" name="btn1" >
			</td>
		</tr>
	</table>
</td>
</tr>
</table>
</form>
</BODY>
</HTML>

