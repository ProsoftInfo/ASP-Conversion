<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasClassificationNameInsert.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 18, 2002
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
			if (confirm("Do You want to define another Classification"))
				window.location.href = "MasClassificationEntry.asp"
			else
				window.location.href = "../welcome_Inventory.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>

<%
'XML DOM Variables
Dim oDOM,newElem,Root,objfs

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,sSql,iCount

dim sClassName,sCatCode,sItemType,spGroup,spGroupCode,arrpGroupCode

sClassName = trim(Request.Form("txtClassName"))
sCatCode = trim(Request.Form("selCategory"))
sItemType = trim(Request.Form("hItmType"))
spGroup = trim(Request.Form("hpGroup"))

if Trim(sItemType)="" or IsNull(sItemType) then sItemType="NAP"

if not len(spGroup) = 3 then
	arrpGroupCode = split(spGroup,":")
	spGroupCode = arrpGroupCode(1)
else
	spGroupCode = 1
end if

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ISNULL(MAX(GROUPCODE),0) + 1 FROM INV_M_CLASSIFICATION WHERE (GROUPCODE = (SELECT ISNULL(MAX(GROUPCODE), 0) FROM INV_M_CLASSIFICATION))"
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
set iCount = dcrs(0)
iCount = iCount
dcrs.close
if cint(iCount) = 1 then
	sSql = "INSERT INTO INV_M_CLASSIFICATION (GROUPCODE,GROUPNAME,PARENTGROUP,CHILDCOUNT," &_
		"ATTRIBUTECOUNT,GROUPCATEGORY,ITEMTYPEID) VALUES " &_
		"(" & iCount & "," & Pack(sClassName) & ",1,0,0," & Pack(sCatCode) & "," & Pack(sItemType) & ")"
'	rESPONSE.wRITE SsQL
	con.Execute sSql

    sSql = "UPDATE INV_M_CLASSIFICATION SET PARENTGROUP = (SELECT GROUPCODE FROM INV_M_CLASSIFICATION)"
    con.Execute sSql
%>
	<BODY onLoad = "msgbox('Classification has been Created Successfully','Y')">
<%
else
	with dcrs
			.CursorLocation = 3
			.CursorType = 3
		.Source = "SELECT GROUPNAME FROM INV_M_CLASSIFICATION WHERE LOWER(GROUPNAME) = " & Pack(lcase(sClassName)) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
    If Not dcrs.EOF Then
%>
	<BODY onLoad = "msgbox('Classification Name Already Exists','N')">
<%
	Response.End
    End If
    dcrs.Close

    If Trim(spGroup) = "GRP" Then
		sSql = "INSERT INTO INV_M_CLASSIFICATION (GROUPCODE,GROUPNAME,PARENTGROUP,CHILDCOUNT," &_
			"ATTRIBUTECOUNT,GROUPCATEGORY,ITEMTYPEID) VALUES " &_
			"(" & iCount & "," & Pack(sClassName) & "," & iCount & ",0,0," & Pack(sCatCode) & "," & Pack(sItemType) & ")"
'Response.write ssQl
		con.Execute sSql
%>
	<BODY onLoad = "msgbox('Classification has been Created Successfully','Y')">
<%
    Else
		sSql = "INSERT INTO INV_M_CLASSIFICATION (GROUPCODE,GROUPNAME,PARENTGROUP,CHILDCOUNT," &_
			"ATTRIBUTECOUNT,GROUPCATEGORY,ITEMTYPEID) VALUES " &_
			"(" & iCount & "," & Pack(sClassName) & "," & spGroupCode & ",0,0,NULL," & Pack(sItemType) & ")"
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT COUNT(PARENTGROUP) FROM INV_M_CLASSIFICATION WHERE PARENTGROUP = " & CInt(spGroupCode) & " AND GROUPCODE <> " & CInt(spGroupCode) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set iCount = dcrs(0)
		If Not dcrs.EOF Then
            sSql = "UPDATE INV_M_CLASSIFICATION SET CHILDCOUNT = " & iCount & " WHERE GROUPCODE = " & CInt(spGroupCode)
            con.Execute sSql
        End If
%>
	<BODY onLoad = "msgbox('Classification has been Created Successfully','Y')">
<%
    End If
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
</FORM>