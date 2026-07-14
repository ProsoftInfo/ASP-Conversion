<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppRoleCreationInsert.asp
	'Module Name				:	Admin (Role Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	December 08, 2010
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
<!--#include virtual="/include/populate.asp"-->

<%
Dim dcrs,sSql,iInternalRoleID,sRoleDesc,sPassType,objDOM,Root,iRoleID

'sRoleDesc = trim(Request.Form("txtDesc"))

Set dcrs = Server.CreateObject("ADODB.RecordSet")

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
objDOM.async = False
objDOM.load(Request)
	
Set Root     = objDOM.documentElement
con.beginTrans

sPassType = Root.Attributes.getNamedItem("TYPE").value
sRoleDesc = Root.Attributes.getNamedItem("ROLEDESC").value
iRoleID   = Root.Attributes.getNamedItem("ROLEID").value

If sPassType = "CRN" Then

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(ROLEID)+1,1) FROM MS_ROLES"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iInternalRoleID = trim(dcrs(0))
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ROLEID FROM MS_ROLES WHERE LOWER(ROLEDESCRIPTION) = " & Pack(lcase(sRoleDesc)) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if dcrs.EOF then

		sSql = "INSERT INTO MS_ROLES(ROLEID,ROLEDESCRIPTION) VALUES( " &_ 
			" " & iInternalRoleID & "," & Pack(sRoleDesc) & ")"
		'Response.Write sSql & "<BR>"
		con.Execute sSql
	else
	
	end if
	dcrs.Close
	
Else	'If sPassType = "CRN" Then
		
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ROLEID FROM MS_ROLES WHERE LOWER(ROLEDESCRIPTION) = " & Pack(lcase(sRoleDesc)) & " AND ROLEID <>" & iRoleID
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if dcrs.EOF then
		sSql = "UPDATE MS_ROLES SET ROLEDESCRIPTION=" & Pack(sRoleDesc) & " WHERE ROLEID=" & iRoleID
		'Response.Write sSql & "<BR>"
		con.Execute sSql
	else

	end if
	dcrs.Close
	
End IF	'If sPassType = "CRN" Then

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
