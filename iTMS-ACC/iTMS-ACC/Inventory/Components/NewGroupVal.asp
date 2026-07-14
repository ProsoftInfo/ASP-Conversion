<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	NewGroupVal.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 18, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	Component(Tree view for Classification)
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
<%
dim sClassCode,dcrs,iItemCode

sClassCode = trim(Request("Groupcode"))
'sClassCode = "1"
	set dcrs = Server.CreateObject("ADODB.Recordset")

	with dcrs
		.Source = "SELECT ITEMCODE FROM INV_M_ITEMGROUP WHERE CLASSIFICATIONCODE = " & sClassCode & " AND LEAFNODE = 1"
		.ActiveConnection = con
		.Open
	end with
	set iItemCode = dcrs(0)
	if not dcrs.EOF then
		iItemCode = iItemCode
		response.ContentType = "text/xml"
		response.write "<?xml version='1.0'?>" &vbCrLf
		response.write "<Root>" & vbCrLf
		response.write "<Item Code=""" & trim(iItemCode) & """/>" &vbCrLf
		response.write "</Root>"
	else
		response.ContentType = "text/xml"
		response.write "<?xml version='1.0'?>" &vbCrLf
		response.write "<Root>" & vbCrLf
		response.write "<Item Code='-1'/>" &vbCrLf
		response.write "</Root>"
	end if
	dcrs.Close
	con.close
	set con = nothing

%>

