<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->

<%
	Dim Objrs,sQuery,sOrgId,iParCode,sInvDate,sTemp,Temparr,sParTy,sParSubTy
	Dim sInvNo,sRetVal,sBookCode,sPayVal,sUnitID,sFinPeriod,sTemparr
	Dim sSTDate,sEndDate
	sTemp = Request("sValue")
	Temparr = Split(sTemp,"?")
	sParTy = Temparr(0)
	sParSubTy = Temparr(1)
	iParCode = Temparr(3)
	sInvNo = Trim(Temparr(4))
	sInvDate = Trim(Temparr(5))
	sBookCode = Temparr(6)
	sUnitID = Trim(Temparr(7))
	sFinPeriod = Session("FinPeriod")
	sTemparr = Split(sFinPeriod,":")
	sStDate = "01/04/"&Trim(sTemparr(0))
	sEndDate = "31/03/"&Trim(sTemparr(1))
	
	'Response.write sTemp
	
	sPayVal = sInvNo&"-"&sInvDate
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	
	IF CStr(sBookCode) = "04" Then
		sRetVal = CheckPurInv()
	Elseif CStr(sBookCode) = "05" Then
		sRetVal = CheckSalInv()
	Else
		sRetVal = "W" 'Wrong Code
	End IF
	
	Response.Write sRetVal
	
	
	
	
%>

<%
	Function CheckPurInv()
		sQuery = "Select CreatedTransNo From Acc_T_CreatedVoucherHeader Where  "&_
				 "PartyType = '"&sParTy&"' and PartySubType = "&sParSubTy&" and  "&_
				 "PartyCode = "&iParCode&" and PayToRecdFrom = '"&sPayVal&"' and BookCode = '04' "&_
				 "and OUDefinitionID = '"&sUnitID&"' "&_
				 "and Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sStDate&"',103) "&_
				 "and Convert(datetime,VoucherDate,103) <= Convert(datetime,'"&sEndDate&"',103) "
				 

				 
		Objrs.Open sQuery,Con
		IF Not Objrs.EOF Then
			CheckPurInv = "N"
		Else
			CheckPurInv = "C"
		End IF		
		Objrs.Close
	End Function
%>

<%
	Function CheckSalInv()
		sQuery = "Select CreatedTransNo From Acc_T_CreatedVoucherHeader Where  "&_
				 "PartyType = '"&sParTy&"' and PartySubType = "&sParSubTy&" and  "&_
				 "PartyCode = "&iParCode&" and CreatedVoucherNo = '"&sInvNo&"' and "&_
				 "Convert(Char,VoucherDate,103) = Convert(Char,'"&sInvDate&"',103) and BookCode = '05' "&_
				 "and OUDefinitionID = '"&sUnitID&"' "&_
				 "and Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sStDate&"',103) "&_
				 "and Convert(datetime,VoucherDate,103) <= Convert(datetime,'"&sEndDate&"',103) "
				 
		'Response.Write sQuery
				 
		Objrs.Open sQuery,Con
		IF Not Objrs.EOF Then
			CheckSalInv = "N"
		Else
			CheckSalInv = "C"
		End IF		
		Objrs.Close
	End Function
%>
