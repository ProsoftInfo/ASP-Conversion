<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLContraEntAccChk.asp
	'Module Name				:	ACCOUNTS (TRANSACTION)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	08/10/2004
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
	Dim iAccHead,sQuery,sOrgid,Objrs,iConVal,iBookAccHead
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	iAccHead = Request("AccHead")
	sOrgid = Request("orgID")
	iBookAccHead = Request("BkAccHd")
	
	sQuery = "Select FromAccountHead From Acc_M_ContraEntries Where OUDefinitionID = '"&sOrgid&"' "&_
			 "and FromAccountHead = "&iBookAccHead&" and ToAccountHead = "&iAccHead&" "
			 
	Objrs.Open sQuery,Con
	IF Not Objrs.EOF Then
		iConVal = Objrs(0)
	Else
		iConVal = 0
	End if
	Objrs.Close
	
	Response.Write iConVal
			 
	
%>