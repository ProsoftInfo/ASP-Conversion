<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasCategoryAddEditPopInsert.asp	
	'Module Name				:	Inventory 
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 16, 2003
	'Modified On				:   Ragavendran R
	'Tables Used				:   Jul 22,2011
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
<%
'XML DOM Variables
Dim oDOM,newElem,Root,objfs,tempNode,bFlag,sChkFlag,sEligible

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,sSql
dim sCatCode,sCatName,sCatShName,arrTemp
arrTemp = split(trim(Request.Form("selCategory")),"|")

sCatCode = trim(arrTemp(0))
sCatName = trim(Request.Form("txtCatName"))
sCatShName = trim(Request.Form("txtCatShName"))
sChkFlag = Request("hFlag")
sEligible = Request("chkEligible")
'Response.Write "sChkFlag ="& sChkFlag 
Response.write "<font color=red>"
'Response.write "sEligible="& sEligible

if trim(sEligible)="" or IsNull(sEligible) then sEligible = "0"

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

If trim(sChkFlag) = "N" then 'New Category
	sCatCode = Request.Form("txtCatCode")
	sSQL = "Select IsNull(Max(CategoryCode),0)+1 from INV_M_CLASSIFICATIONCATEGORY "
	dcrs.open sSql,con
	if not dcrs.eof then
	    sCatCode = dcrs(0)
	end if 
	dcrs.close
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT CATEGORYCODE FROM INV_M_CLASSIFICATIONCATEGORY WHERE LOWER(CATEGORYNAME) = " & Pack(lcase(sCatName)) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if dcrs.EOF then
		sSql = "INSERT INTO INV_M_CLASSIFICATIONCATEGORY (CATEGORYCODE,CATEGORYNAME,CATEGORYSHORTNAME,EligibleForWebStore) VALUES" &_
			"(" & Pack(ucase(sCatCode)) & "," & Pack(ucase(sCatName)) & "," & Pack(ucase(sCatShName)) & ","& pack(sEligible) &")"
		'Response.Write sSql & "<BR>"
		con.Execute sSql 
	End If
Else 'If trim(sChkFlag) = "A" then
	with dcrs
		.CursorLocation = 3
		.CursorType = 3	
		.Source = "SELECT CATEGORYCODE FROM INV_M_CLASSIFICATIONCATEGORY WHERE LOWER(CATEGORYNAME) = " & Pack(lcase(sCatName)) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sSql = "UPDATE INV_M_CLASSIFICATIONCATEGORY SET CATEGORYNAME = " & Pack(ucase(sCatName)) & ", " &_
			"CATEGORYSHORTNAME = " & Pack(ucase(sCatShName)) & ",EligibleForWebStore="& pack(sEligible) &" WHERE CATEGORYCODE = " & Pack(sCatCode) & ""
		'Response.Write sSql & "<BR>"
		con.Execute sSql
	End If
End If 'If trim(sChkFlag) = "N" then
'Response.End 
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	'Response.end
	con.CommitTrans
end if

con.close
set con = nothing
Response.Redirect "MasCategoryAddEditPop.asp"
%>
