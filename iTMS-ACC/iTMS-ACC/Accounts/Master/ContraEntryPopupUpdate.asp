<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ContraEntryPopupUpdate.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
Dim sQuery,iCnt,iArrCnt,Objrs,sDisplay
Dim sOrgid,iFromHead,iToHead,iArrToHead

 
sOrgid=trim(Request.QueryString("OrgCode"))
iFromHead=trim(Request.QueryString("FromHead"))
iToHead=trim(Request.QueryString("ToHead"))

Set Objrs = Server.CreateObject("ADODB.RecordSet")

iArrToHead = Split(iToHead,",")
'Response.Write "<p>OrgCode="&sOrgid
'Response.Write "iFromHead = "& iFromHead 
'Response.Write "iToHead = "& iToHead

Con.beginTrans

if trim(iFromHead)<>"" and trim(iFromHead)<>"0" then
    for iArrCnt=0 to UBound(iArrToHead)
	    sQuery = "Select Count(1) From Acc_M_ContraEntries Where OUDefinitionID = '"&sOrgid&"' "&_
			     "and FromAccountHead = "&iFromHead&" and ToAccountHead = "&iArrToHead(iArrCnt)
	    Objrs.Open sQuery,con
	    IF Not Objrs.EOF Then
		    iCnt = Objrs(0)
	    End IF
	    Objrs.Close
    	
	    IF CStr(iCnt) = "0" Then
			    sQuery="INSERT Acc_M_ContraEntries(OUDefinitionID, FromAccountHead, ToAccountHead) "&_
					    "VALUES('"&sOrgid&"',"&iFromHead&","&iArrToHead(iArrCnt)&")"
			    'Response.Write sQuery
				con.Execute(sQuery)
    		
	    End IF
    next		
else
    Response.Write "Map Account Head for the Book"
end if
'con.RollbackTrans
'Response.End 
Con.CommitTrans
%>
