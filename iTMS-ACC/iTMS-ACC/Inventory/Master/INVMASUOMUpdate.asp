<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	INVMASUOMUpdate.asp
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
Dim sUOMName,sUOmShName,sDecimalAllowed,sUOMCode,sSql,sArrValue,sArr
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
set dcrs= server.CreateObject("ADODB.Recordset")
sArr = trim(Request.QueryString("sTemp"))
sArrValue= split(sArr,":")
sUOMCode = sArrValue(0)
sUOmShName = sArrValue(1)
sUOMName = sArrValue(2)
sDecimalAllowed = sArrValue(3)

con.beginTrans

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	'.Source = "SELECT UOMCODE FROM MS_UNITOFMEASUREMENT WHERE LOWER(UOMSHORTDESCRIPTION) = " & Pack(lcase(sUOmShName)) & " AND DECIMALALLOWED = " & Pack(sDecimalAllowed) & ""
	.Source = "SELECT UOMCODE FROM MS_UNITOFMEASUREMENT WHERE LOWER(UOMSHORTDESCRIPTION) = " & Pack(lcase(sUOmShName)) 
	.ActiveConnection = con
	.Open
end with
'Response.Write "<p> " & dcrs.Source 
set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		sSql = "UPDATE MS_UNITOFMEASUREMENT SET UOMDESCRIPTION = " & Pack(ucase(sUOMName)) & "," &_
			"UOMSHORTDESCRIPTION = " & Pack(ucase(sUOmShName)) & ",DECIMALALLOWED = " & Pack(sDecimalAllowed) & " " &_
			"WHERE UOMCODE = " & Pack(sUOMCode) & ""
		con.Execute sSql
	else
		sSql = "INSERT INTO MS_UNITOFMEASUREMENT (UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION,DECIMALALLOWED,NOOFDECIMALS)"&_
				"VALUES ('"& sUOMCode &"','"& sUOMName &"','"& sUOmShName&"','"&sDecimalAllowed&"',0)"
		con.execute sSql
	end if
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
