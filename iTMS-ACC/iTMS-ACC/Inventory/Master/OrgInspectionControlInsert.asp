<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgInspectionControlInsert.asp	
	'Module Name				:	Inventory (Organization Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Modified By                :   Ragavendran R
	'Modified On				:   July 21,2011
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

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag,org) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "OrgSalesControlEntry.asp";
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
dim sOrgCode,sPO,sOO,sP,sPS,iOrgIndex,sOrgName

sOrgCode = trim(Request.Form("hOrgCode"))
sOrgName = trim(Request.Form("hOrgName"))
sPO = trim(Request.Form("chkPO"))
sOO = trim(Request.Form("chkOO"))
sP = trim(Request.Form("chkP"))
sPS = trim(Request.Form("chkPS"))

if isnull(sPO) or IsEmpty(sPO) or sPO = "" then sPO = "0"
if isnull(sOO) or IsEmpty(sOO) or sOO = "" then sOO = "0"
if isnull(sP) or IsEmpty(sP) or sP = "" then sP = "0"
if isnull(sPS) or IsEmpty(sPS) or sPS = "" then sPS = "0"

con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT OUDEFINITIONID FROM INV_CONTROL_ORGINSPECTION WHERE OUDEFINITIONID = " & Pack(sOrgCode) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if dcrs.EOF then
	sSql = "INSERT INTO INV_CONTROL_ORGINSPECTION (OUDEFINITIONID,PREORDERINSPECTION,OUTORDERINSPECTION," &_
		" PROCESSINSPECTION,POSTSALEINSPECTION) VALUES" &_
		"(" & Pack(sOrgCode) & "," & Pack(sPO) & "," & Pack(sOO) & "," &_
		"" & Pack(sP) & "," & Pack(sPS) & ")"
	'Response.Write sSql & "<BR>"
	con.Execute sSql
	
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Inspection \nhas been defined Successfully','Y','1')">
<%
else
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Inspection \nhas been already defined','N','1')">
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
