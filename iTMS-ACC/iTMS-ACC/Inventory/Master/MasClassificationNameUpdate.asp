<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasClassificationNameUpdate.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 22, 2003
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
	function msgbox(strr,flag,sPara) {
		if (flag == "Y") {
			alert(strr);
			if (sPara == "A"){			
			window.location.href = "MasClassificationAmendEntry.asp"
			}
			else
			{
			window.location.href = "MasClassificationEntry.asp"
			}
		}
		else {
			alert(strr);
			document.formname.target = "body"
			window.history.back(1)
		}
		
	}
//-->
</SCRIPT>

<%
dim dcrs,sSql

dim sClassName,spGroupCode,sPara

spGroupCode = trim(Request.Form("hpGroup"))
sClassName = trim(Request.Form("txtClassName"))
sPara = Request("hPara")

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT GROUPCODE FROM INV_M_CLASSIFICATION WHERE LOWER(GROUPNAME) = " & Pack(lcase(sClassName)) & ""
	.ActiveConnection = con
	.Open
end with
' Response.Write dcrs.Source
' Response.End 
set dcrs.ActiveConnection = nothing
if dcrs.EOF then
    sSql = "UPDATE INV_M_CLASSIFICATION SET GROUPNAME = " & Pack(sClassName) & " WHERE GROUPCODE = " & spGroupCode & ""
    ' Response.Write sSql
    con.Execute sSql
%>
	<BODY onLoad = "msgbox('Classification has been Updated Successfully','Y','sPara')">
<%
else
%>
	<BODY onLoad = "msgbox('Classification Already Exists','N','sPara')">
<%
End If

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
<form method="POST" name="formname" action="" target="bodyFrame">
</FORM>