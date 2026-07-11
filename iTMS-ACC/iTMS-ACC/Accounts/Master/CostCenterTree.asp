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
	'Program Name				:	CostCenterTree.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Senthil E
	'Created On					:	July 22,2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	March 19, 2011
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
dim sIP,sOrgID
dim sModule

sIP = GetSettings("IP")
sOrgID = Session("organizationcode")
%>
<HTML>
<HEAD>
<TITLE>iTms Group-Head List View</TITLE>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script>
var sCallFrom = "";

function getGroupHeadList() {
	return document.formname.ctlGroupHeadList;
}

function RefreshData() {
	var form = document.formname;
	var tree = getGroupHeadList();
	if (tree) {
		tree.DSN = "http://" + form.hIP.value + "/Accounts/components/GetCGroup.asp?flag=B&orgid=" + encodeURIComponent(form.hUnitId.value);
	}
}

function CreateGP() {
	var form = document.formname;
	var tree = getGroupHeadList();
	sCallFrom = "GP";
	form.GCode.value = tree ? tree.GroupValue : "0";
	form.GName.value = tree ? tree.GroupName : "";
	form.action = "CCGroupAndHeadNameCreation.asp?sCallFrom=" + sCallFrom;
	form.submit();
}

function CreateHD() {
	var form = document.formname;
	var tree = getGroupHeadList();
	sCallFrom = "HD";
	if (!tree || tree.GroupValue === "0") {
		alert("Select A Group");
		return;
	}
	form.GCode.value = tree.GroupValue;
	form.GName.value = tree.GroupName;
	form.action = "CCGroupAndHeadNameCreation.asp?sCallFrom=" + sCallFrom;
	form.submit();
}

function Edit() {
	var form = document.formname;
	var tree = getGroupHeadList();
	var editType = tree && Number(tree.HeadValue) > 0 ? "HD" : "GP";
	form.hCallFrom.value = editType;

	if (editType === "HD") {
		if (tree && Number(tree.HeadValue) > 0) {
			form.hHeadValue.value = tree.HeadValue;
			form.hHeadName.value = tree.HeadName;
			form.action = "CostCenterDet.asp";
			form.submit();
		} else {
			alert("Select Head");
			if (tree && tree.focus) {
				tree.focus();
			}
		}
		return;
	}

	if (tree && Number(tree.GroupValue) > 0) {
		form.GCode.value = tree.GroupValue;
		form.GName.value = tree.GroupName;
		form.action = "CostCenterDet.asp";
		form.submit();
	} else {
		alert("Select Group");
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="RefreshData()">
<form method="POST" name="formname" action="" target="Detail">
<input type="hidden" name="hIP" value="<%=sIP%>">
<input type="hidden" name="hUnitId" value="<%=sOrgID%>">
<input type="hidden" name="hHeadValue" value="0">
<input type="hidden" name="hHeadName" value="">
<input type="hidden" name="GName" value="">
<input type="hidden" name="GCode" value="">
<input type="hidden" name="GroupFlag" value="C">
<input type="hidden" name="hCallFrom" value="">

<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="10px"></td>
<td>
<TABLE BORDER="0" CELLSPACING=0 CELLPADDING=0>
</TABLE>
</td>
</tr>
<tr>
<td width="10px">&nbsp;</td>
<td>
<div id="ctlGroupHeadList" data-itms-tree-control data-width="263px" data-height="353px"
	data-dsn="http://<%=sIP%>/Accounts/components/GetCGroup.asp"
	data-list-name="COST CENTER"></div>
</td>
</tr>
<tr>
	<td align="center">
		<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
	</td>
	<td valign="top">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td class="ActionCell">
					<input type="button" value="CrGroup" onClick="CreateGP()" class="ActionButtonX"  name="btn1">&nbsp;&nbsp;
					<input type="button" value="CrHead" onClick="CreateHD()" class="ActionButtonX"  name="btn2">&nbsp;&nbsp;
                    <input type="button" value="Edit" onclick="Edit()" class="ActionButton" name="btn3" >
				</td>
			</tr>
		</table>
	</td>
</tr>
</table>
</form>
</BODY>
</HTML>

