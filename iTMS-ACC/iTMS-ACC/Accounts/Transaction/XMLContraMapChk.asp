<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLContraMapChk.asp
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%	
	Dim iAccHead,sQuery,sOrgid,Objrs,iConVal,iBookAccHead,sTemp,sTemparr,sBkChk
	Dim iBkCnt,sRetVal
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	
	sTemp = Request("sValue")
	sTempArr = Split(sTemp,"-")
	
	iBookAccHead = sTemparr(0)
	iAccHead = sTemparr(1)
	sOrgid = sTemparr(2)
	
	sQuery = "Select Count(1) From VWorgBookNames Where BookAccountHead = "&iAccHead&" and "&_
		 "OUDefinitionID = '"&sOrgID&"' and BookCode Not In (04,05) "
		 
	'Response.Write sQuery &"<br>"
		 
	Objrs.Open sQuery,Con
	IF Not Objrs.Eof Then
		iBkCnt = Objrs(0)
	Else
		iBkCnt = 0
	End IF
	objrs.close
	
	'Response.Write iBkCnt
	
	IF Cstr(iBkCnt) = "0" Then
		sRetVal = "T" 'Entry Cannot be allowed Selected Acc Head is Not Mapped to an Book
	Else
		sRetVal = "F" 'Selected Acc Head is Mapped to a Book Contra to be Checked
	End IF
	
	
	IF Cstr(sRetVal) = "F" Then 'Contra To be Checked
		sQuery = "Select Count(1) From Acc_M_ContraEntries Where OUDefinitionID = '"&sOrgid&"' "&_
				 "and FromAccountHead = "&iBookAccHead&" and ToAccountHead = "&iAccHead&" "

		'Response.Write sQuery

		Objrs.Open sQuery,Con
		IF Not Objrs.EOF Then
			iConVal = Objrs(0)
		Else
			iConVal = 0
		End if
		Objrs.Close
		
		IF Cstr(iConVal) = "0"  Then
			sRetVal = "F" 'Contra is Not Mapped Entry Has to be Blocked
		Else
			sRetVal = "T" 'Contra is Mapped Entry can allowed
		End IF
	Else
		sRetVal = "T" 'Account Head Not a Contra Itself Entry Allowed.
	End IF
	
	Response.Write sRetVal
			 
	
%>
