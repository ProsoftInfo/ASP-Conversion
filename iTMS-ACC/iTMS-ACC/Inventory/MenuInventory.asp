<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MenuInventory.asp
	'Module Name				:	Inventory (Menu)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 01, 2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populatemenu.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Inventory Menu</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardMenu.css" TYPE="text/css">

<body topmargin=0 leftmargin=0>
<table border="0" cellspacing="0" cellpadding="0" height="100%" width="100%">
  <tr valign=top>
    <td class=TextDisplay>
		<%populatemenu "4"%>
    </td>
  </tr>

</table>
</body>
</html>
