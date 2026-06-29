<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	RoleDeletion.asp
	'Module Name				:	Admin 
	'Author Name				:	UmaMaheswari S
	'Created On					:	January 14, 2011
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "ApplicationRoles.asp"
		}
		else {
			alert(strr);
			window.location.href = "ApplicationRoles.asp"
		}
	}
//-->
</SCRIPT>
<%
Dim dcrs,dcrs1,sSql,nRoleID

nRoleID = Request.QueryString("RoleID")
con.beginTrans

Set dcrs  = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")



	sSql =  " SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_ROLEACTIVITY "&_
			" WHERE RoleID = "& nRoleID &" "
	'Response.Write sSql	
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with

	set dcrs1.ActiveConnection = nothing
		
	If dcrs1.EOF Then

		sSql = "DELETE FROM Ms_Roles WHERE  ROLEID = "& nRoleID &" "
		'Response.Write "<p>sql="&sSql
		con.Execute sSql
	
	%>
		<BODY onLoad = "msgbox('Selected Role has been deleted Successfully','Y')">
	<%
	else
	%>
		<BODY onLoad = "msgbox('Role is not deleted because already its mapped to User','N')">
	<%
	end if
dcrs1.Close 

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
