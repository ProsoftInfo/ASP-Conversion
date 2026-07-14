<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->

<%
	dim Objrs,sMonYr,sMon,sYr,sLastDay,sTemp,sQuery,dBookBal,sBookNo
	Dim sPassCrDr,iMaxTrNo,sUnit
	
	Set objRs = Server.CreateObject("ADODB.Recordset")
	
	sMonYr = Request("MonYr")
'	Response.Write "<p>sMonYr="&sMonYr
	sTemp = Split(sMonYr,":")
	sMonYr = sTemp(0)
	sBookNo = sTemp(1)
	sUnit = sTemp(2)
	sMon = Right(sMonYr,2)
	sYr = Left(sMonYr,4)
	
	IF CInt(sMon) = 4 or CInt(sMon) = 6 or CInt(sMon) = 9 or CInt(sMon) = 11 Then
		sLastDay = "30"
	ElseIF CInt(sMon) = 2 Then
		IF CInt(sYr) Mod 2 = 0 Then
			sLastDay = "29"
		Else
			sLastDay = "28"
		End IF
	Else
		sLastDay = "31"
	End IF
	
	sLastDay = sLastDay&"/"&sMon&"/"&sYr
	
	'sQuery = "Select isNull(Max(BrsTransactionNo),0) From Acc_T_BankReconciliation Where "&_
	'		 "Month(BRSPeriodTo) = "&sMon&" and bookno = "&sBookNo&" "
	
	sQuery = "Select isNull(Max(B.BrsTransactionNo),0) From Acc_T_BankReconciliation B, "&_
			 "Acc_T_VoucherHeader H Where Month(B.BRSPeriodTo) = "&sMon&" and B.Bookno = "&sBookNo&" "&_
			 "And B.BRSTransactionNo = H.BRSTransactionNo And H.OUDefinitionID = '"&sUnit&"' "
			 
		
'	Response.Write sQuery
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	IF Not Objrs.EOF Then
		iMaxTrNo = Objrs(0)
	Else
		iMaxTrNo = 0
	End IF
	Objrs.Close
	
	sQuery = "Select BankBalance,isNull(CDIndication,'C') From Acc_T_BankReconciliation Where  "&_
			 "Convert(Char,BRSPeriodTo,103) = '"&sLastDay&"'  and isNull(BookNo,0) = "&sBookNo&" "&_
			 "and BrsTransactionNo = "&iMaxTrNo&" "
	
	
'	Response.write sQuery
	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	IF Not Objrs.EOF Then
		dBookBal = Trim(Objrs(0))
		sPassCrDr = Trim(Objrs(1))
	Else
		dBookBal = 0
		sPassCrDr = "C"
	End IF
	Objrs.Close
	
	dBookBal = CDbl(dBookBal)
	'Response.Write dBookBal
	
	
	IF CStr(sPassCrDr) = "D" Then
		dBookBal = Cdbl(dBookBal) * -1
	End IF
	
	'dBookBal = FormatNumber(dBookBal,2,,,0)
	Response.Write dBookBal
		
	
%>
