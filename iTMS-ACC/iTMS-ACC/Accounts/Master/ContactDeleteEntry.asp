<%@ Language=VBScript %>
<%	option explicit	
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ContactDeleteEntry.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	Sep 29,2011
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<%

Dim sQuery,iContactNo,iExistingPartyCode,iExistingPartyContactNo
Dim objRs

Set objRs = Server.CreateObject("ADODB.RecordSet")

iContactNo = Request.QueryString("hContactNo")

Con.BeginTrans

'Response.Write "<p> " & iContactNo
if trim(iContactNo)<>"" then

	iExistingPartyCode = 0
	iExistingPartyContactNo = 0
	
	sQuery = "Select PartyCode,PartyContactNo from APP_M_Contacts where ContactNumber="& iContactNo
	objRs.Open sQuery,con 
	if not objRs.EOF then
		iExistingPartyCode = objRs(0)
		iExistingPartyContactNo = objRs(1)
	end if 
	objRs.Close 
						
	if trim(iExistingPartyCode) <> "0" then
	
		sQuery ="Delete from APP_M_PartyContactPersons where PartyCode="& iExistingPartyCode & " and ContactNo = " & iExistingPartyContactNo & ""
		'Response.Write "<p>"&sQuery
		con.execute sQuery
		
		sQuery ="Update APP_M_PartyContactPersons set ContactNo =  ContactNo -1 where PartyCode="& iExistingPartyCode & " and ContactNo > " & iExistingPartyContactNo  & ""
		Response.Write "<p>"&sQuery
		con.execute sQuery
		
	end if 'if trim(iExistingPartyCode) <> "0" then
	
	sQuery ="Delete from APP_M_Contacts where ContactNumber="& iContactNo
	'Response.Write "<p>"&sQuery
	con.execute sQuery
	
	
end if ' if trim(iContactNo)<>"" then

	
		
'con.rollbacktrans
'Response.End 

Response.Clear 
con.committrans
%>
