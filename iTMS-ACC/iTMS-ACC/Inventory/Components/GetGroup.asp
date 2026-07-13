<%@ Language=VBScript %>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetGroup.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 16, 2002
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
	dim dcrs,sSql,sGCode,sGName,sGChild,sPGroup
	dim OutData,newElem,newElem1

	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")

	with dcrs
		.Source = "SELECT GROUPCODE,GROUPNAME,CHILDCOUNT,PARENTGROUP FROM INV_M_CLASSIFICATION ORDER BY GROUPCODE"
		.ActiveConnection = con
		.Open
	end with
	set sGCode = dcrs(0)
	set sGName = dcrs(1)
	set sGChild = dcrs(2)
	set sPGroup = dcrs(3)

	Set newElem = OutData.createElement("Root")

	do while not dcrs.EOF
		'response.write "<Group Code=""" & trim(sGCode) & """ Name=""" & replace(replace(trim(sGName),"&","&amp;"),"<","&lt;") & """ ChildCount=""" & trim(sGChild)& """ ParentGroup=""" & trim(sPGroup)& """/>" &vbCrLf

		Set newElem1 = OutData.createElement("Group")

		newElem1.setAttribute "Code", trim(sGCode)
		newElem1.setAttribute "Name", trim(sGName)
		newElem1.setAttribute "ChildCount", trim(sGChild)
		newElem1.setAttribute "ParentGroup", trim(sPGroup)

		newElem.appendChild newElem1

		dcrs.MoveNext
	loop
	dcrs.close

	OutData.appendChild newElem

	Response.ContentType="text/xml"
	Response.Write OutData.xml

	con.close
	set con = nothing
%>

