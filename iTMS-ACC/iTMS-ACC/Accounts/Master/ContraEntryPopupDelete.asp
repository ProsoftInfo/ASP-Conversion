<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ContraEntryPopupDelete.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 11,2010
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
Dim sQuery,iCnt
dim sOrgid,iFromHead,iToHead,iArrToHead
Dim Objrs,sDisplay

sOrgid=trim(Request.QueryString("OrgCode"))
iFromHead=trim(Request.QueryString("FromHead"))
iToHead=trim(Request.QueryString("ToHead"))

iArrToHead = split(iToHead,",")

Set Objrs = Server.CreateObject("ADODB.RecordSet")

for iCnt = 0 to UBound(iArrToHead)
	sQuery = "Delete from Acc_M_ContraEntries where FromAccountHead = "& iFromHead &" and ToAccountHead = "& iArrToHead(iCnt)
	con.Execute(sQuery)
next
		
%>
