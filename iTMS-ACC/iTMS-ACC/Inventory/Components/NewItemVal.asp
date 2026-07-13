<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	NewItemVal.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 30, 2004
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

<!-- #include File="../../include/DatabaseConnection.asp" -->
<%
	dim sClassCode,dcrs

	sClassCode = trim(Request("Groupcode"))
	
	set dcrs = Server.CreateObject("ADODB.Recordset")

	with dcrs
		.Source = "SELECT ISNULL(CHILDCOUNT,0) FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & sClassCode & ""
		.ActiveConnection = con
		.Open
	end with
	if not dcrs.EOF then
		response.ContentType = "text/xml"
		response.write "<?xml version='1.0'?>" &vbCrLf
		response.write "<Root>" & vbCrLf
		response.write "<Item Counts=""" & cint(dcrs(0)) & """/>" &vbCrLf
		response.write "</Root>"
	end if
	dcrs.Close
	con.close
	set con = nothing

%>

