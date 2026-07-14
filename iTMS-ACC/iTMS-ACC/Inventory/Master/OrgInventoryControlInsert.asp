<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgInventoryControlInsert.asp	
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
			window.location.href = "OrgInspectionControlEntry.asp";
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
dim sOrgCode,iStock,sABC,iA,iB,iC,iOrgIndex,sOrgName
dim sFSN,iFast,iSlow,iNoN,sFIFO,sLIFO,sWA,sRep,sUnit,sLocation

sOrgCode = trim(Request.Form("hOrgCode"))
sOrgName = trim(Request.Form("hOrgName"))
iStock = trim(Request.Form("txtStock"))

if iStock = "" then iStock = "NULL"

sABC = trim(Request.Form("radABC"))
iA = trim(Request.Form("txtA"))
iB = trim(Request.Form("txtB"))
iC = trim(Request.Form("txtC"))
sFSN = trim(Request.Form("radFSN"))
iFast = trim(Request.Form("txtFast"))
iSlow = trim(Request.Form("txtSlow"))
iNoN = trim(Request.Form("txtNon"))
sFIFO = trim(Request.Form("chkFIFO"))
sLIFO = trim(Request.Form("chkLIFO"))
sWA = trim(Request.Form("chkWA"))
sRep = trim(Request.Form("radRep"))
sUnit = trim(Request.Form("radUnit"))
sLocation = trim(Request.Form("radLoc"))

if sABC = "0" then
	iA = "0"
	iB = "0"
	iC = "0"
end if

if sFSN = "0" then
	iFast = "0"
	iSlow = "0"
	iNoN = "0"
end if

if isnull(sFIFO) or IsEmpty(sFIFO) or sFIFO = "" then sFIFO = "0"
if isnull(sLIFO) or IsEmpty(sLIFO) or sLIFO = "" then sLIFO = "0"
if isnull(sWA) or IsEmpty(sWA) or sWA = "" then sWA = "0"

if isnull(iA) or IsEmpty(iA) or iA = "" then iA = "0"
if isnull(iB) or IsEmpty(iB) or iB = "" then iB = "0"
if isnull(iC) or IsEmpty(iC) or iC = "" then iC = "0"

if isnull(iFast) or IsEmpty(iFast) or iFast = "" then iFast = "0"
if isnull(iSlow) or IsEmpty(iSlow) or iSlow = "" then iSlow = "0"
if isnull(iNoN) or IsEmpty(iNoN) or iNoN = "" then iNoN = "0"

Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT OUDEFINITIONID FROM INV_CONTROL_ORGINVENTORY WHERE OUDEFINITIONID = " & Pack(sOrgCode) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

con.beginTrans

if dcrs.EOF then
	sSql = "INSERT INTO INV_CONTROL_ORGINVENTORY (OUDEFINITIONID,ALLOWFIFOVALUATION,ALLOWLIFOVALUATION," &_
		" ALLOWWAVALUATION,ALLOWREPLENISHMENT,ALLOWINTERUNITTRANSFER,ALLOWLOCATIONTRANSFER," &_
		" STOCKHOLDINGPERIOD,ABCCLASSIFICATIONEXIST,CLASSIFICATIONAVALUE,CLASSIFICATIONBVALUE," &_
		" CLASSIFICATIONCVALUE,FSNCLASSIFICATIONEXIST,FASTMOVINGDAYS,SLOWMOVINGDAYS,NONMOVINGDAYS) VALUES" &_
		"(" & Pack(sOrgCode) & "," & Pack(sFIFO) & "," & Pack(sLIFO) & "," &_
		"" & Pack(sWA) & "," & Pack(sRep) & "," & Pack(sUnit) & "," &_
		"" & Pack(sLocation) & "," & iStock & "," & Pack(sABC) & "," &_
		"" & iA & "," & iB & "," & iC & "," &_
		"" & Pack(sFSN) & "," & iFast & "," & iSlow & "," & iNoN & ")"
	'Response.Write sSql & "<BR>"
	con.Execute sSql
	
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Inventory \nhas been defined Successfully','Y','1')">
<%
else
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Inventory \nhas been already defined','N','1')">
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
