<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	comClassificationTree.asp
	'Module Name				:	Inventory (Master)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 16, 2002
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
<!--#include virtual="/include/GetSettings.asp"-->
<%
	dim sIP
	sIP = GetSettings("IP")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Classification Tree</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<style type="text/css">
html,
body {
	height: 100%;
	overflow: hidden;
}
.classification-tree-table {
	height: 100%;
}
</style>
</HEAD>
<BODY leftMargin="0" topMargin="0" MARGINHEIGHT="0" MARGINWIDTH="0">
	<table border="0" cellspacing="0" cellpadding="0" width=100% class="classification-tree-table">
		<tr>
	        <td class="ExcelHeaderCell">Classification Hierarchy</td>
		</tr>
		<tr>
			<td >
				<div id="ctlClassificationTree" data-itms-tree-control data-tree-kind="classification"
					data-dsn="../Components/GetGroup.asp" data-width="100%" data-height="calc(100vh - 21px)"></div>
            </td>
		</tr>
	</table>
</BODY>
</HTML>

