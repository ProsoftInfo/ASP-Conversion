<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgManufacturingControlInsert.asp	
	'Module Name				:	Inventory (Organization Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 03, 2002
	'Modified By                :   Ragavendran R
	'Modified On				:   July 21,2011
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

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag,org) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "OrgControlDefn.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
dim dcrs,sSql
dim sOrgCode,sPO,sOO,iOrgIndex,sOrgName

sOrgCode = trim(Request.Form("hOrgCode"))
sOrgName = trim(Request.Form("hOrgName"))
sPO = trim(Request.Form("chkIn"))
sOO = trim(Request.Form("chkEx"))

if isnull(sPO) or IsEmpty(sPO) or sPO = "" then sPO = "0"
if isnull(sOO) or IsEmpty(sOO) or sOO = "" then sOO = "0"

con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT OUDEFINITIONID FROM INV_CONTROL_ORGMANUFACTURING WHERE OUDEFINITIONID = " & Pack(sOrgCode) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if dcrs.EOF then
	sSql = "INSERT INTO INV_CONTROL_ORGMANUFACTURING (OUDEFINITIONID,INTERNALMANUFACTURING," &_
		" EXTERNALMANUFACTURING) VALUES" &_
		"(" & Pack(sOrgCode) & "," & Pack(sPO) & "," & Pack(sOO) & ")"
	'Response.Write sSql & "<BR>"
	con.Execute sSql
	
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Manufacturing \nhas been defined Successfully','Y','1')">
<%
else
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Manufacturing \nhas been already defined','N','1')">
<%end if

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
