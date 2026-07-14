<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	CategoryDeletionUpdate.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 10, 2004
	'Modified  By               :   Ragavendran R
	'Modified On				:   Jul 22,2011
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
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
</SCRIPT>
<%
Response.Write "<font color=red>"
'XML DOM Variables
Dim oDOM,CategoryNode,Root,objfs,DeleteNode,oNode,sExp,sExp1

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,sSql
dim sCatCode,sCatName,arrTemp
arrTemp = split(trim(Request.Form("selCategory")),"|")

sCatCode = trim(arrTemp(0))
sCatName = trim(Request.Form("hCatName"))

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT GROUPCATEGORY FROM INV_M_CLASSIFICATION WHERE GROUPCATEGORY = " & Pack(sCatCode) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
if dcrs.EOF then
	sSql = "DELETE INV_M_CLASSIFICATIONCATEGORY WHERE CATEGORYCODE = " & Pack(sCatCode) & ""
	'Response.Write sSql & "<BR>"
	con.Execute sSql
	
End If

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	Response.End 
	con.CommitTrans
end if

con.close
set con = nothing

Response.Redirect "MasCategoryAddEditPop.asp"
%>
