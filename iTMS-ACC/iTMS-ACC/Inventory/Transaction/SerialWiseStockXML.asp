<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MatConsDetailsInsert.asp
	'Module Name				:	Inventory (Material Consumption)
	'Author Name				:	KUMAR K A
	'Created On					:	MAY 09, 2003
	'Modified By				:	
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	MatConsDetailsEntry.asp
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
<!--#include file="../../include/populate.asp"-->
<%
'XML DOM Variables
dim newxml,RootNode,HeaderNode,LotNode,SerialNode,objfs,objTxt,MixNode

' Create our DOM Document Objects
Set newxml = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
'set objTxt = objfs.CreateTextFile(server.MapPath("..\Temp\Transaction\MatConsInsert.txt"))
newxml.async = false
newxml.load(Request)
newxml.Save server.MapPath("../temp/transaction/SRLWISESTK"&Session.SessionID&".xml")
%>

