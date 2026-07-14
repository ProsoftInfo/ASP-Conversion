<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ClassDeletionUpdate.asp
	'Module Name				:	Inventory (Classification Deletion)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 10, 2004
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
dim oDOM,RootNode,ClassNode,sSql,objfs
dim iPGroup,iClass,sExp,iCtr

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

Set objfs = Server.CreateObject("Scripting.FileSystemObject")

oDOM.async = false
oDOM.load(Request)

Set RootNode = oDOM.documentElement

con.BeginTrans

sExp ="//DETAILS/CLASSIFICATION [ @DELETE = 'Y']"
Set ClassNode = RootNode.Selectnodes(sExp)
for iCtr =  0 to ClassNode.length - 1
	iClass = ClassNode.Item(iCtr).Attributes.getNamedItem("CCODE").Value
	iPGroup = ClassNode.Item(iCtr).Attributes.getNamedItem("PGROUP").Value

	sSql = "DELETE INV_M_CLASSIFICATIONATTRIBUTES WHERE GROUPCODE = " & iClass & ""
	con.Execute sSql

'	sSql = "DELETE INV_M_ORGCLASSIFICATION WHERE CLASSIFICATIONCODE = " & iClass & ""
'	con.Execute sSql

	sSql = "DELETE INV_M_STORAGECLASSIFICATION WHERE CLASSIFICATIONCODE = " & iClass & ""
	con.Execute sSql

	sSql = "DELETE INV_M_CLASSIFICATION WHERE GROUPCODE = " & iClass & ""
	con.Execute sSql

	if iClass <> iPGroup then
		sSql = "UPDATE INV_M_CLASSIFICATION SET CHILDCOUNT = (CHILDCOUNT - 1) WHERE GROUPCODE = " & iPGroup & ""
		'Response.Write sSql
		con.Execute sSql
	end if
next

if con.Errors.count <> 0 then
	dim iErrCounter
	con.RollbackTrans
	for iErrCounter=0 to con.Errors.count
		Response.Write con.Errors(iErrCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	if objfs.FileExists(Server.MapPath("../temp/master/ClassDelete.xml")) then
		objfs.DeleteFile server.MapPath("../temp/master/ClassDelete.xml")
	end if

	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing


%>

