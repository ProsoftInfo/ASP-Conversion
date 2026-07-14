<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	INVMASUOMDelete.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	RAGAVENDRAN
	'Created On					:	MARCH 16,2010
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
'XML DOM Variables
Dim oDOM,newElem,Root,objfs,tempNode,bFlag,dcrs
Dim sUOMCode,sSql
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
set dcrs= server.CreateObject("ADODB.Recordset")
sUOMCode = trim(Request.QueryString("sTemp"))

con.beginTrans

		sSql = "Delete from MS_UNITOFMEASUREMENT WHERE UOMCODE = " & Pack(sUOMCode) & ""
		con.Execute sSql

'Response.Write sSql 
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
