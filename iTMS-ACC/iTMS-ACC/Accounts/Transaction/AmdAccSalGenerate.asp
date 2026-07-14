<%@ Language="VBScript" %>
<% option explicit %>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AmdAccSalGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Maheshwari  S.
	'Created On					:	Nov 14, 2006
	'Modified By				:	R.Manohar Prabhu
	'Modified On				:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	ACCSALVOUCHERS.ASP	
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
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/AccPopulate.asp"-->
<!--#include virtual="/Accounts/Transaction/GenerateSalVou.asp"-->
<%
Dim rs1,rs2,iTransNo,sQry,sOrgId,sHCrDr,sHAccHeadCode,sPAccHead,sRecNo,sAmtRec,sTAmt
Dim sMon,sYear,sMonYr,dDrAmt,dCrAmt,sBkCode,sBkNo,sAmount,sAdjAmt,sPayNo,sAmtPaid
Dim sTbGLVal,sOldCrNo,sOldAccNo,iAccTrNo,sVoucDate,iParCode,sParTy,sParSubTy,sParCrDrIndi
Dim sHAmt,sHAccHead,sFormVal,sFrmApp

iTransNo = Request("hTransNo")
sFormVal = Session("AmdSal")
Session("AmdSal") = " "

Set rs1 = server.CreateObject("ADODB.RecordSet")
set rs2 = server.CreateObject("ADODB.RecordSet")
	
sQry = "Select CreatedVoucherNo,VoucherNumber From Acc_T_VoucherHeader Where CreatedTransNo = "&iTransNo&" "
rs1.Open sQry,Con
IF Not rs1.EOF Then
	sOldCrNo = rs1(0)
	sOldAccNo = rs1(1)
End IF
rs1.Close

sQry = "Select isNull(OtherApplnTransNo,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = " & iTransNo
rs1.Open sQry,Con
IF Not rs1.EOF Then
	sFrmApp = rs1(0)
End IF
rs1.Close



Con.BeginTrans


sQry = "Select TransactionNumber,BookNumber,OUDefinitionID,Convert(Char,VoucherDate,103), "&_
	   "Month(VoucherDate),Year(VoucherDate),VoucherAmount,PartyCode,PartyType, "&_ 
	   "PartySubType,CrDrIndication From Acc_T_VoucherHeader Where CreatedTransNo = "&iTransNo
	   
Response.write sQry &"<br><br>"
rs1.Open sQry,Con
IF Not rs1.EOF Then
	iAccTrNo = rs1(0)
	sBkNo = rs1(1)
	sOrgId = rs1(2)
	sVoucDate = rs1(3)
	sMon = rs1(4)
	sYear = rs1(5)
	sAmount = rs1(6)
	iParCode = rs1(7)
	sParTy = rs1(8)
	sParSubTy = rs1(9)
	sParCrDrIndi = rs1(10)	
End IF
rs1.Close
sBkCode = "05"

IF Len(sMon) = 1 Then
	sMonYr = "0"&sMon&sYear
Else
	sMonYr = sMon&sYear
End IF

IF CStr(iAccTrNo) <> "" Then

	sQry = "SELECT T.AccUnitAccountHead, T.Amount, T.TransCrDrIndication,  "&_
		   "V.SummaryPosting FROM Acc_T_VoucherDetails T INNER JOIN VwOrgGLHeads V ON  "&_
		   "T.AccUnitAccountHead = V.AccountHead AND T.AccountingUnit = V.OUDefinitionID "&_
		   "WHERE T.TransactionNumber = "&iAccTrNo
		   
	Response.Write sQry &"<br>"
	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQry
		.Open
	End WIth
	Set rs1.ActiveConnection = Nothing
	Do While Not rs1.EOF
		sHAmt = rs1("Amount")
		sHCrDr = rs1("TransCrDrIndication")
		sHAccHead = rs1("AccUnitAccountHead")
		
		IF CStr(rs1("SummaryPosting")) = "1" Then 'If it is Marked For Summary 
			sQry =  "UPDATE Acc_T_GLTransactions Set Amount = (Amount - "&sHAmt&")  Where Accounthead =  "&sHAccHead&" and BookCode = '05' "&_
					"and BookNumber = "&sBkNo&" and OUDefinitionID = '"&sOrgId&"' and TransactionNumber = 0 "&_
					"and TransCrDrIndication = '"&sHCrDr&"' and Month(VoucherDate) = "&sMon&" and Year(VoucherDate) = "&sYear&" "
		Else
			sQry =  "UPDATE Acc_T_GLTransactions Set Amount = (Amount - "&sHAmt&") Where "&_
					"Accounthead = "&sHAccHead&" and TransactionNumber = "&iAccTrNo&" "&_
					"And Amount - "&sHAmt&" >= 0 "
		End IF
		
		Response.Write sQry &"<br><br>"
		Con.Execute sQry
		
		IF CStr(sHCrDr) = "C" Then
			sQry = "UPDATE Acc_T_GLAccTransactAmt SET MonthCrAmount = (MonthCrAmount - "&sHAmt&") "&_
				   "Where Accounthead = "&sHAccHead&" and OUDefinitionID = "&sOrgId&" "&_ 
				   "and MonthYear = "&sMonYr&" "
		Else
			sQry = "UPDATE Acc_T_GLAccTransactAmt SET MonthDrAmount = (MonthDrAmount - "&sHAmt&") "&_
				   "Where Accounthead = "&sHAccHead&" and OUDefinitionID = "&sOrgId&" "&_ 
				   "and MonthYear = "&sMonYr&" "
		End IF	
		
		Response.Write sQry &"<br><br>"
		Con.Execute sQry
		
		rs1.MoveNext
	Loop
	rs1.Close
	Response.Write "**************** Removal of Tables Voucher Details is Over ***************** <br><br>"
	'****************** Removal of Tables Voucher Details is Over *************************************************************
	sQry = "SELECT T.AccountHead,T.TransCrDrIndication,T.TaxAmount,V.SummaryPosting "&_
		   "FROM Acc_T_VoucherTaxDetails T INNER JOIN VwOrgGLHeads V ON  "&_
		   "T.AccountHead = V.AccountHead WHERE V.OUDefinitionID = '"&sOrgId&"' "&_
		   "AND T.TransactionNumber = "&iAccTrNo
	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQry
		.Open
	End WIth
	Set rs1.ActiveConnection = Nothing
	Do While Not rs1.EOF
		sHAmt = rs1("TaxAmount")
		sHCrDr = rs1("TransCrDrIndication")
		sHAccHead = rs1("AccountHead")
		
		IF CStr(rs1("SummaryPosting")) = "1" Then 'If it is Marked For Summary 
			sQry =  "UPDATE Acc_T_GLTransactions Set Amount = (Amount - "&sHAmt&")  Where Accounthead =  "&sHAccHead&" and BookCode = '05' "&_
					"and BookNumber = "&sBkNo&" and OUDefinitionID = '"&sOrgId&"' and TransactionNumber = 0 "&_
					"and TransCrDrIndication = '"&sHCrDr&"' and Month(VoucherDate) = "&sMon&" and Year(VoucherDate) = "&sYear&" "
		Else
			sQry =  "UPDATE Acc_T_GLTransactions Set Amount = (Amount - "&sHAmt&") Where "&_
					"Accounthead = "&sHAccHead&" and TransactionNumber = "&iAccTrNo
		End IF
		
		Response.Write sQry &"<br><br>"
		Con.Execute sQry
		
		IF CStr(sHCrDr) = "C" Then
			sQry = "UPDATE Acc_T_GLAccTransactAmt SET MonthCrAmount = (MonthCrAmount - "&sHAmt&") "&_
				   "Where Accounthead = "&sHAccHead&" and OUDefinitionID = "&sOrgId&" "&_ 
				   "and MonthYear = "&sMonYr&" "
		Else
			sQry = "UPDATE Acc_T_GLAccTransactAmt SET MonthDrAmount = (MonthDrAmount - "&sHAmt&") "&_
				   "Where Accounthead = "&sHAccHead&" and OUDefinitionID = "&sOrgId&" "&_ 
				   "and MonthYear = "&sMonYr&" "
		End IF	
		
		Response.Write sQry &"<br><br>"
		Con.Execute sQry
		
		rs1.MoveNext
	Loop
	rs1.Close
	Response.Write "**************** Removal of Tables VoucherTaxDetails is Over ***************** <br><br>"	    
	'****************** Removal of Tables VoucherTaxDetails is Over *************************************************************
	sQry = "Select AccountHead From VwOrgPartyType Where OUDefinitionID = '"&sOrgId&"' "&_
		   "and PartyType = '"&sParTy&"' and PartySubType = "&sParSubTy
	rs1.Open sQry,Con
	Response.Write sQry &"<br><br>"
	IF Not rs1.EOF Then
		sPAccHead = rs1(0)
	End IF
	rs1.Close

	sQry =  "UPDATE Acc_T_GLTransactions Set Amount = (Amount - "&sAmount&")  Where Accounthead =  "&sPAccHead&" and BookCode = '05' "&_
			"and BookNumber = "&sBkNo&" and OUDefinitionID = '"&sOrgId&"' and TransactionNumber = 0 "&_
			"and TransCrDrIndication = '"&sParCrDrIndi&"' and Month(VoucherDate) = "&sMon&" and Year(VoucherDate) = "&sYear&" "

	Response.Write sQry &"<br><br>"
	Con.Execute sQry

	IF CStr(sParCrDrIndi) = "C" Then
		sQry = "UPDATE Acc_T_GLAccTransactAmt SET MonthCrAmount = (MonthCrAmount - "&sAmount&") "&_
			   "Where Accounthead = "&sPAccHead&" and OUDefinitionID = "&sOrgId&" "&_ 
			   "and MonthYear = "&sMonYr&" "
	Else
		sQry = "UPDATE Acc_T_GLAccTransactAmt SET MonthDrAmount = (MonthDrAmount - "&sAmount&") "&_
			   "Where Accounthead = "&sPAccHead&" and OUDefinitionID = "&sOrgId&" "&_ 
			   "and MonthYear = "&sMonYr&" "
	End IF	
		
	Response.Write sQry &"<br><br>"
	Con.Execute sQry

	IF CStr(sParCrDrIndi) = "C" Then
		sQry = "UPDATE Acc_T_PartyTransactAmt SET MonthCrAmount = (MonthCrAmount - "&sAmount&") where PartyType = '"&sParTy&"' "&_
			   "and PartySubType = "&sParSubTy&" and PartyCode = "&iParCode&" "&_
			   "and OUDefinitionID = "&sOrgId&"  and MonthYear = "&sMonYr&" "
	Else
		sQry = "UPDATE Acc_T_PartyTransactAmt SET MonthDrAmount = (MonthDrAmount - "&sAmount&") where PartyType = '"&sParTy&"' "&_
			   "and PartySubType = "&sParSubTy&" and PartyCode = "&iParCode&" "&_
			   "and OUDefinitionID = "&sOrgId&"  and MonthYear = "&sMonYr&" "
	End IF	
		
	Response.Write sQry &"<br><br>"
	Con.Execute sQry

	sQry = "Delete From Acc_T_PartyTransactions Where TransactionNumber = "&iAccTrNo
	Con.Execute sQry

	sQry = "Delete From Acc_T_GLTransactions Where TransactionNumber = "&iAccTrNo
	Con.Execute sQry

	'****************** Removing of Old Enteries is Over **********************************************************************
	'///////////////////// Deletion of Enteries Starts Here\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	sQry = "Select PayablesNumber,AmountPaid from Acc_T_CreatedPybleAdjDet where CreatedTransNo = "&iTransNo&" "&_
		   " and Adjusttype <> 'Null' "
	with rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQry
		.ActiveConnection = Con
		.Open
	End with
	Do while not rs1.EOF 
		sPayNo = rs1(0)
		sAmtPaid = rs1(1)
		
		sQry = "Select Advanceadjusted From Acc_T_AdvancePayments Where CreatedAdvanceNo = "&sPayNo&" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = Con
			.Open
		End with
		If not rs2.EOF then
			sAdjAmt = rs2(0)
		End If 
		rs2.Close
			
		sQry = "UPDATE Acc_T_AdvancePayments SET AdvanceAdjusted = ("&sAmtPaid&" - "&sAdjAmt&") where CreatedAdvanceNo = "&sPayNo&" "
		Con.Execute sQry 
				
		sQry = "Select Advanceadjusted From Acc_T_CreatedAdvances Where CreatedAdvanceNo = "&sPayNo&" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = Con
			.Open
		End with
		If not rs2.EOF then
			sAdjAmt = rs2(0)
		End If 
		rs2.Close
				
		sQry = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted = ("&sAmtPaid&" - "&sAdjAmt&") where CreatedAdvanceNo = "&sPayNo&" "
		Con.Execute sQry 
		rs1.movenext
	loop
	rs1.Close
	'======================For Receivables=======================
	'sQry = "Select CreatedAdvanceNo,AmountAdjusted From Acc_T_CreatedAdvanceAdj  "&_
	'	   "Where CreatedTransno = "&iTransNo

	sQry = "Select C.CreatedAdvanceNo,C.AmountAdjusted,A.AdvanceNumber From  "&_
		   "Acc_T_CreatedAdvanceAdj C,Acc_T_AdvancePayments A Where  "&_
		   "C.CreatedTransno = "&iTransNo&" and C.CreatedAdvanceNo = A.CreatedAdvanceNo "
		   
	Response.Write sQry &"<br>"
	With rs1
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQry
		.Open
	End With
	Set rs1.ActiveConnection = Nothing
	Do While Not rs1.EOF
		sQry =  "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = AdvanceAdjusted - "&rs1(1)&"  "&_
				"Where CreatedAdvanceNo = "&rs1(0)&" "
		Con.Execute sQry
				
		
		
		sQry =  "Update Acc_T_AdvancePayments Set AdvanceAdjusted = AdvanceAdjusted - "&rs1(1)&"  "&_
				"Where AdvanceNumber = "&rs1(2)&" "
				
		Response.Write sQry &"<br>"
		Con.Execute sQry
				
		rs1.MoveNext
	Loop
	rs1.Close
			
			
	sQry = "Delete from Acc_T_PybleAdjustmentDetails where PaidByTransactionNo = "&iAccTrNo&" "
	con.execute sQry
			
	sQry = "Delete from Acc_T_RcvblAdjustmentDetails where RecdByTransactionNo = "&iAccTrNo&" "
	con.execute sQry
			
	sQry = "Delete from Acc_T_Receivables where TransactionNumber = "&iAccTrNo&" "
	Response.Write sQry &"<br>"
	con.execute sQry
				
	sQry = "Delete from Acc_T_VoucherTaxDetails where TransactionNumber =  "&iAccTrNo&" "
	con.execute sQry
			
	sQry = "Delete from Acc_T_VoucherDetails where TransactionNumber = "&iAccTrNo&" "
	con.execute sQry
									
	sQry = "Delete from Acc_T_VoucherHeader where TransactionNumber = "&iAccTrNo&" "
	con.execute sQry
			
	sQry = "Delete From Acc_T_GLTransSummaryBreakup where TransactionNumber = "&iAccTrNo&" "
	con.execute sQry
					
	sQry = "Delete from Acc_T_CreatedPybleAdjDet where CreatedTransNo =  "&iTransNo&" "
	con.execute sQry
				
	sQry = "Delete from Acc_T_CreatedRcvbleAdjDet where CreatedTransNo =  "&iTransNo&" "
	con.execute sQry
			
	sQry = "Delete from  Acc_T_CreatedReceivables where CreatedTransNo  = "&iTransNo&" "
	con.execute sQry
				
	sQry = "Delete from Acc_T_CreatedVoucherTaxDet where CreatedTransNo =  "&iTransNo&" "
	con.execute sQry
				
	sQry = "Delete from Acc_T_CreatedVoucherDetails where CreatedTransNo =  "&iTransNo&" "
	con.execute sQry
				
	sQry = "Delete from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iTransNo&" "
	con.execute sQry 

	sQry = "Delete from Acc_T_CreatedAdvanceAdj where CreatedTransno = "&iTransNo
	con.execute sQry

	'========================== Deletion is Over ============================================
	CreateSalInv sOldCrNo,sOldAccNo
	'=========================== Creation Of Sales Invoice is Over ==========================
	sTAmt = 0
	dDrAmt = 0
	dCrAmt = 0
	sTAmt = CheckTBGL(sOrgId)
			
	sTbGLVal = Split(sTAmt,":")
	dDrAmt = sTbGLVal(0)
	dCrAmt = sTbGLVal(1)
			
	IF CStr(dDrAmt) = CStr(dCrAmt) Then
		con.CommitTrans
		'con.RollbackTrans
		
		'Response.Redirect "ACCSALVOUCHERS.ASP?hFormVal="&sFormVal
		Response.Redirect "MANAGESALESVOUCHERS.ASP"
		
	Else
		con.RollbackTrans
		Response.write  "<b>Debit Amount and Credit Amount does'nt Match Transaction Rolled back </b><br><br>"
		Response.write  "DR AMOUNT " & dDrAmt &"<br>"
		Response.write  "CR AMOUNT " & dCrAmt &"<br>"
		Response.Write  "<b>DIFFERENCES " & FormatNumber(Round(CDbl(dDrAmt) - CDbl(dCrAmt),2),2,,,0) &"</b><br>"
		Response.End 
	End IF
Else
	Response.Clear
	Response.Write "Error On Admendment <br>"
	Response.End
End IF
%>
