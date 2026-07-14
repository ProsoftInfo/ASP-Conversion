<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppActivityUpdate.asp
	'Module Name				:	Admin (Activity Creation)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	December 17, 2003
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
Dim dcrs,sSql,sTemp,sArr,nPracticeCode,nProcessCode

sTemp = Request.QueryString("sPassData")
sArr  = Split(sTemp,":")
nPracticeCode = Trim(sArr(0))
nProcessCode = Trim(sArr(1))

con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")

sSql = " DELETE FROM MS_APPLICATIONPROCESS WHERE APPLICATIONCODE = "& nPracticeCode&" AND PROCESSCODE IN ("& nProcessCode&") "
'Response.Write "<p>sql="& sSql
con.Execute sSql	   

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
