<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	comItemClassificationTree.asp
	'Module Name				:	Inventory (Master)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 18, 2002
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
<!-- #include File="../../include/GetSettings.asp" -->
<%
	dim sIP
	sIP = GetSettings("IP")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Classification Tree</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
</HEAD>
<BODY leftMargin=15 topMargin=5 MARGINHEIGHT="0" MARGINWIDTH="0">
<div id="ctlCategoryTree" data-itms-tree-control data-tree-kind="item-classification"
	data-dsn="../Components/GetCategoryGroup.asp" data-itype="NO:NO:NO"
	data-width="552px" data-height="340px"></div>
</BODY>
</HTML>

