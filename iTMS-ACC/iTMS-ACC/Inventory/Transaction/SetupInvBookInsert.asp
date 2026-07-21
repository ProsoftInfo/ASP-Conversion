<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SetupInvBookInsert.asp
	'Module Name				:	Inventory
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 06,2014
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	SetupInvBooks.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
Dim oDOM
set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.load(Request)
oDOM.save server.MapPath("../XMLData/BookSetup.xml")
%>

