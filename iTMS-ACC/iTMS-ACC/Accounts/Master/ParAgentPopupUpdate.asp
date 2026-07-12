<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParAgentPopupUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 19, 2010
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	
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
<%

dim Temp,arrAgent,iCounter,iCode,iRecCount
dim sQuery,objRs,sParCode,sUnitId,iAgentType

' Create our DOM Document Objects

Set objRs = Server.CreateObject("ADODB.RecordSet")

sUnitId=trim(Request.Form("selUnitId"))
sParCode=trim(Request.Form("hPartyCode"))
iAgentType=trim(Request.Form("selAgentType"))

Temp=trim(Request.Form("hSelectedValue"))
arrAgent=Split(Temp,":")

sQuery="delete APP_R_AgentOrgParty where OUDefinitionID='"&sUnitId&"'"&_
			" and PartyCode="&sParCode&" and AgentSubType = 1"
			
con.Execute(sQuery)			

for iCounter=0 to UBound(arrAgent)
	sQuery="INSERT INTO APP_R_AgentOrgParty(OUDefinitionID,AgentCode,AgentSubType,PartyCode)"&_
			" VALUES('"&sUnitId&"',"&trim(arrAgent(iCounter))&",1,"&sParCode&")"
			
'Response.Write sQuery &"<br><br>"

con.Execute(sQuery)
next	
			
%>

<HTML>
<base target="_self">
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script>
window.__itmsPopupCompat = {
	type: "autoClose",
	message: "Agent Related successfully",
	returnValue: "Done"
};
</script>
<script src="../../scripts/PopupModernCompat.js"></script>

<BODY onLoad = "init()">
<form name="formname">
<input type=hidden name="hPartyCode" value="<%=sParCode%>">
</form>
</BODY>
<HTML>
