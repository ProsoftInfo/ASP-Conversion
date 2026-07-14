<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AdjustVouWithBillUpdate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 28, 2002
	'Modified On				:   January  23, 2003
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
<!--#include virtual="/include/Accpopulate.asp"-->
<%
Dim objRS,oDOM,Root,sExp,TempNode,sQuery,iAdvCloseNo,iCtr,iAdvNo,objRS2,iParCode
Dim iCrRecNo,iAccRecNo,sAdjType,iBillTrNo,iCrTrNo,iTrNo,dVouAmt,dBillAmt,sVouDate
Dim sUnitID,dCrDrAmt,nOldAdjAmt,sCheckAlreadyAdjustOrNot,nBillDetCheck,nToBeAdjustAmt
Dim nAdjustAlreadyAmount,nTotalAmtToAdjust,sTemp,sPartyType,sPartySubType,nPartyCode
Dim iCrAdvNo,sCrVouNo,sAdjNarr,blnInsertAdvanceTable,blnExistAdvNo,iExistAdvNo
Dim iExistCrAdvNo

iCrTrNo = Request.Form("hCrTrNo")
iTrNo = Request.Form("hTrNo")
dVouAmt = Request.Form("hVouAmt")
sVouDate = Request.Form("hVouDate")
iAdvNo = Request.Form("hAdvNo")
iParCode = Request.Form("hParCode")
sUnitID = Request.Form("hUnitID")
dCrDrAmt = Request.Form("hDrCrAmt")
sCrVouNo = Request.Form("hVouNumber")

IF CStr(iTrNo) = "" Then
	iTrNo = 0
	sVouDate = "30/10/2007"
	dVouAmt = dCrDrAmt
End IF

Set objRs  = server.CreateObject("adodb.recordset")
Set objRs2  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../temp/transaction/Bill Closing_Det_"&Session.SessionID&".xml")
Set Root=oDOM.documentElement
sExp = "//PayRec/Doc"
Set TempNode = Root.selectNodes(sExp)

Con.BeginTrans
IF CStr(iAdvNo) <> 0 Then
	sQuery = "Select Distinct B.BillTrNo,B.PayRecNo,H.BookCode,B.BillAdjAmt From Acc_T_BillsClosingDetails B, "&_
			 "Acc_T_VoucherHeader H Where B.BillTrNo = H.TransactionNumber  "&_
			 "And B.AdvCloseNo = "&iAdvNo
			 
	Response.Write sQuery &"<br><br>"
	With objRS
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set objRS.ActiveConnection = Nothing
	Do While Not objRS.EOF
		IF CStr(objRS("BookCode")) = "05" or CStr(objRS("BookCode")) = "06"  Then
			sQuery = "Select CreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&objRS("PayRecNo")
			
			Response.Write sQuery &"<br><br>"
			objRS2.Open sQuery,con
			IF Not objRS2.EOF Then
				iCrRecNo = objRS2("CreatedReceivable")
			Else
				iCrRecNo = 0
			End IF
			objRS2.Close
			
			sQuery = "Update Acc_T_Receivables Set AmountReceived = AmountReceived - "&objRS("BillAdjAmt")&" Where "&_
					 "ReceivableNumber = "&objRS("PayRecNo")
					 
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
			
			sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = AmountReceived - "&objRS("BillAdjAmt")&" Where "&_
					 "ReceivableNumber = "&iCrRecNo
					 
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
					 
			sQuery = "Delete FROM Acc_T_RcvblAdjustmentDetails WHERE ReceivableNumber = "&objRS("PayRecNo")&" "&_
					  "And AdjustType = 'C' "
			
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
			
			sQuery = "Delete FROM Acc_T_CreatedRcvbleAdjDet WHERE ReceivableNumber = "&iCrRecNo&" "&_
					  "And AdjustType = 'C' "
					  
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
				 
			
		End IF
		
		IF CStr(objRS("BookCode")) = "04" or CStr(objRS("BookCode")) = "07"  Then
			sQuery = "Select CreatedPayablesNumber From Acc_T_Payables Where PayablesNumber = "&objRS("PayRecNo")
			objRS2.Open sQuery,con
			IF Not objRS2.EOF Then
				iCrRecNo = objRS2("CreatedPayablesNumber")
			Else
				iCrRecNo = 0
			End IF
			objRS2.Close
			
			sQuery = "Update Acc_T_Payables Set AmountPaid = AmountPaid - "&objRS("BillAdjAmt")&" Where "&_
					 "PayablesNumber = "&objRS("PayRecNo")
			Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
			
			sQuery = "Update Acc_T_CreatedPayables Set AmountPaid = AmountPaid - "&objRS("BillAdjAmt")&" Where "&_
					 "PayablesNumber = "&iCrRecNo
			Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
					 
			sQuery = "Delete FROM Acc_T_PybleAdjustmentDetails WHERE PayablesNumber = "&objRS("PayRecNo")&" "&_
					 "And AdjustType = 'C' "
			Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
			
			sQuery = "Delete FROM Acc_T_CreatedPybleAdjDet WHERE PayablesNumber = "&iCrRecNo&" "&_
					  "And AdjustType = 'C' "
			Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
				 
			
		End IF
		objRS.MoveNext
	Loop
	objRS.Close
	
	
	sQuery = "Select BillTrNo,PayRecNo,BillAdjAmt From Acc_T_BillsClosingDetails Where "&_
			 "BillType In('P','R') And AdvCloseNo = "&iAdvNo
				 
	With objRS
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = COn
		.Open
	End With
	Set objRS.ActiveConnection = Nothing
	Do While Not objRS.EOF
		sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = AdvanceAdjusted - "&objRS("BillAdjAmt")&" "&_
				 "Where CreatedAdvanceNo = "&objRS("BillTrNo")
					 
		Response.Write sQuery &"<br><br>"
		Con.Execute sQuery
			
		sQuery = "Update Acc_T_AdvancePayments Set AdvanceAdjusted = AdvanceAdjusted - "&objRS("BillAdjAmt")&" "&_
				 "Where AdvanceNumber = "&objRS("PayRecNo")
					 
		Response.Write sQuery &"<br><br>"
		Con.Execute sQuery
			
		objRS.MoveNext
	Loop
	objRS.Close
		
	
	
	sQuery = "Delete FROM Acc_T_BillsClosingDetails Where AdvCloseNo = "&iAdvNo
	Response.Write sQuery&"<br><br>"
	Con.Execute sQuery
End IF

'Response.End

'============================================================================================
'============================================================================================
'==================================== Creation Starts Here ==================================
'============================================================================================
sQuery = "Select isNull(Max(AdvCloseNo),0) + 1 From Acc_T_BillsClosingDetails "
objRS.Open sQuery,con
IF Not objRS.EOF Then
	iAdvCloseNo = objRS(0)
End IF
objRS.Close

IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		nBillDetCheck = ""
		sCheckAlreadyAdjustOrNot = ""
		blnInsertAdvanceTable = False
		
		iCrRecNo = TempNode.Item(iCtr).Attributes.getNamedItem("No").Value
		iAccRecNo = TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		sAdjType = TempNode.Item(iCtr).Attributes.getNamedItem("AdjType").Value
		dBillAmt = TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value
		nBillDetCheck = TempNode.Item(iCtr).Attributes.getNamedItem("Check").Value
		
		If nBillDetCheck <> "" Then
			sCheckAlreadyAdjustOrNot = TempNode.Item(iCtr).Attributes.getNamedItem("CheckAdjust").Value
		End If

		Response.Write "<p>sAdjType="&sAdjType  
		
		IF CStr(sAdjType) = "I" Then
			sQuery = "Select TransactionNumber From Acc_T_Receivables Where ReceivableNumber = "&iAccRecNo
			objRS.Open sQuery,con
			IF Not objRS.EOF Then
				iBillTrNo = objRS("TransactionNumber")
			End IF
			objRS.Close
			
			'Added Newly
			If sCheckAlreadyAdjustOrNot = "Y" and nBillDetCheck <> "" Then
				
				nOldAdjAmt = cdbl("0")
				
				sQuery = "Select isNull(AmountReceived,0) FROM Acc_T_CreatedRcvbleAdjDet WHERE ReceivableNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				objRS.Open sQuery,con
				If Not objrs.EOF Then
					nOldAdjAmt = objrs(0)
				End IF
				objrs.Close 
			
				Response.Write "<p>nOldAdjAmt="&nOldAdjAmt
			
				sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = (AmountReceived - "&nOldAdjAmt&") Where ReceivableNumber = "&iCrRecNo &" "
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				sQuery = "Update Acc_T_Receivables Set AmountReceived = (AmountReceived - "&nOldAdjAmt&") Where ReceivableNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
			
				sQuery = "DELETE FROM Acc_T_CreatedRcvbleAdjDet WHERE ReceivableNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
			
				sQuery = "DELETE FROM Acc_T_RcvblAdjustmentDetails WHERE ReceivableNumber = "&iAccRecNo&" and RecdByTransactionNo = "& iTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
				
			End IF
			
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then			
				
				If cdbl(dBillAmt) > "0.00" Then
					
					sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, ReceivedOn,  "&_
							 "AmountReceived, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iCrRecNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
			
					sQuery = "INSERT INTO Acc_T_RcvblAdjustmentDetails (ReceivableNumber, RecdByTransactionNo, ReceivedOn,  "&_
							 "AmountReceived, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iAccRecNo&", "&iTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
					
				End IF	'If dBillAmt <> "0" Then
				
			End IF	'If sCheckAlreadyAdjustOrNot <> "" Then	
			
		End IF
		
		IF CStr(sAdjType) = "D" Then
			sQuery = "Select TransactionNumber From Acc_T_Receivables Where ReceivableNumber = "&iAccRecNo
			objRS.Open sQuery,con
			IF Not objRS.EOF Then
				iBillTrNo = objRS("TransactionNumber")
				iCrRecNo = objRS("TransactionNumber")
			End IF
			objRS.Close
			
			'Added Newly
			'If sCheckAlreadyAdjustOrNot = "Y" and nBillDetCheck <> "" Then
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then
			
				If sCheckAlreadyAdjustOrNot = "N" Then sCheckAlreadyAdjustOrNot = ""
				
				nOldAdjAmt = cdbl("0")
				Response.Write "<p>iCrRecNo="&iCrRecNo
				sQuery = "Select isNull(AmountReceived,0) FROM Acc_T_Receivables WHERE ReceivableNumber = "&iAccRecNo&" and TransactionNumber = "& iCrRecNo &" "
				Response.Write "<p>sQuery = "&squery
				objRS.Open sQuery,con
				If Not objrs.EOF Then
					nOldAdjAmt = objrs(0)
				End IF
				objrs.Close 
			
				Response.Write "<p>nOldAdjAmt="&nOldAdjAmt
			
				sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = (AmountReceived - "&nOldAdjAmt&") Where ReceivableNumber = "&iCrRecNo &" "
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				sQuery = "Update Acc_T_Receivables Set AmountReceived = (AmountReceived - "&nOldAdjAmt&") Where ReceivableNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
			
				sQuery = "DELETE FROM Acc_T_CreatedRcvbleAdjDet WHERE ReceivableNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
			
				sQuery = "DELETE FROM Acc_T_RcvblAdjustmentDetails WHERE ReceivableNumber = "&iAccRecNo&" and RecdByTransactionNo = "& iTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
				
			End IF
			
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then			
				
				If cdbl(dBillAmt) > "0.00" Then
					
					sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, ReceivedOn,  "&_
							 "AmountReceived, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iCrRecNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
			
					sQuery = "INSERT INTO Acc_T_RcvblAdjustmentDetails (ReceivableNumber, RecdByTransactionNo, ReceivedOn,  "&_
							 "AmountReceived, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iAccRecNo&", "&iTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
					
				End IF	'If dBillAmt <> "0" Then
				
			End IF	'If sCheckAlreadyAdjustOrNot <> "" Then	
			
		End IF
		
		IF CStr(sAdjType) = "D" and 1= 2 Then
			
			sQuery = "Select TransactionNumber From Acc_T_Receivables Where ReceivableNumber = "&iAccRecNo
			objRS.Open sQuery,con
			IF Not objRS.EOF Then
				iBillTrNo = objRS("TransactionNumber")
			End IF
			objRS.Close
			'Added Newly
			
			If sCheckAlreadyAdjustOrNot = "Y" and nBillDetCheck <> "" Then
			
				nOldAdjAmt = cdbl("0")
				
				sQuery = "Select isNull(AmountPaid,0)FROM Acc_T_CreatedPybleAdjDet WHERE PayablesNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				objRS.Open sQuery,con
				If Not objrs.EOF Then
					nOldAdjAmt = objrs(0)
				End IF
				objrs.Close 
			
				Response.Write "<p>nOldAdjAmt="&nOldAdjAmt
			
				sQuery = "Update Acc_T_CreatedPayables Set AmountPaid = (AmountPaid - "&nOldAdjAmt&") Where PayablesNumber = "&iCrRecNo &" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				sQuery = "Update Acc_T_Payables Set AmountPaid = (AmountPaid - "&nOldAdjAmt&") Where PayablesNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
			
				sQuery = "DELETE FROM Acc_T_CreatedPybleAdjDet WHERE PayablesNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
			
				sQuery = "DELETE FROM Acc_T_PybleAdjustmentDetails WHERE PayablesNumber = "&iAccRecNo&" and PaidByTransactionNo = "& iTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
			
			End IF	'If sCheckAlreadyAdjustOrNot = "Y" 
			Response.Write "<p>AAAAAAAAAAAAAAAAAiTrNo = "&iTrNo
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then			
				
				If cdbl(dBillAmt) > "0.00" Then
					sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn,  "&_
							 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iCrRecNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write "<p>"&sQuery&"<br><br>"
					Con.Execute sQuery
			
					sQuery = "INSERT INTO Acc_T_PybleAdjustmentDetails (PayablesNumber, PaidByTransactionNo, PaidOn,  "&_
							 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iAccRecNo&", "&iTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
				End IF
			End If
		End IF
		
		IF CStr(sAdjType) = "PI" Then
			
			sQuery = "Select TransactionNumber From Acc_T_Payables Where PayablesNumber = "&iAccRecNo
			objRS.Open sQuery,con
			IF Not objRS.EOF Then
				iBillTrNo = objRS("TransactionNumber")
			End IF
			objRS.Close
			
			'Added Newly
			
			If sCheckAlreadyAdjustOrNot = "Y" and nBillDetCheck <> "" Then
			
				nOldAdjAmt = cdbl("0")
				
				sQuery = "Select isNull(AmountPaid,0)FROM Acc_T_CreatedPybleAdjDet WHERE PayablesNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				objRS.Open sQuery,con
				If Not objrs.EOF Then
					nOldAdjAmt = objrs(0)
				End IF
				objrs.Close 
			
				Response.Write "<p>nOldAdjAmt="&nOldAdjAmt
			
				sQuery = "Update Acc_T_CreatedPayables Set AmountPaid = (AmountPaid - "&nOldAdjAmt&") Where PayablesNumber = "&iCrRecNo &" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				sQuery = "Update Acc_T_Payables Set AmountPaid = (AmountPaid - "&nOldAdjAmt&") Where PayablesNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
			
				sQuery = "DELETE FROM Acc_T_CreatedPybleAdjDet WHERE PayablesNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
			
				sQuery = "DELETE FROM Acc_T_PybleAdjustmentDetails WHERE PayablesNumber = "&iAccRecNo&" and PaidByTransactionNo = "& iTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
			
			End IF	'If sCheckAlreadyAdjustOrNot = "Y" 
			
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then			
				
				If cdbl(dBillAmt) > "0.00" Then
					
					sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn,  "&_
							 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iCrRecNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
			
					sQuery = "INSERT INTO Acc_T_PybleAdjustmentDetails (PayablesNumber, PaidByTransactionNo, PaidOn,  "&_
							 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iAccRecNo&", "&iTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'C', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
					
				End IF	'If dBillAmt <> "0" Then
				
			End IF	'If sCheckAlreadyAdjustOrNot <> "" Then			
			
		End IF
		
		IF CStr(sAdjType) = "C" Then
			sQuery = "Select TransactionNumber From Acc_T_Payables Where PayablesNumber = "&iAccRecNo
			objRS.Open sQuery,con
			IF Not objRS.EOF Then
				iBillTrNo = objRS("TransactionNumber")
			End IF
			objRS.Close
			
			sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn,  "&_
					 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
					 "VALUES ("&iCrRecNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
					 "'C', 1) "
					 Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
			
			sQuery = "INSERT INTO Acc_T_PybleAdjustmentDetails (PayablesNumber, PaidByTransactionNo, PaidOn,  "&_
					 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
					 "VALUES ("&iAccRecNo&", "&iTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
					 "'C', 1) "
					 Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
		
		End IF
		
		IF CStr(sAdjType) = "P" Then
			
			'sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAccRecNo
			sQuery="select isnull(PayablesNumber,0) from Acc_T_CreatedPayables where CreatedTransNo = "& iCrTrNo
			'Response.Write "<p>"&sQuery
			objRS.Open sQuery,con
			IF Not objRS.EOF Then
				iCrRecNo = objRS(0)
			End IF
			objRS.Close
			
			squery = " select isNull(PayablesNumber,0),TransactionNumber from Acc_T_Payables where createdPayablesNumber  = "&iCrRecNo&" "
			objRS2.Open squery,con
			If Not objRS2.EOF Then
				iAccRecNo = objrs2(0)
				iTrNo= objrs2(1)
			End IF
			objRS2.Close
			
			'Added Newly
			'Response.Write "<p>sCheckAlreadyAdjustOrNot="&sCheckAlreadyAdjustOrNot
			'Response.Write "<p>nBillDetCheck="&nBillDetCheck
			
			'If sCheckAlreadyAdjustOrNot = "Y" and nBillDetCheck <> "" Then
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then
			
				nOldAdjAmt = cdbl("0")
				If sCheckAlreadyAdjustOrNot = "N" Then sCheckAlreadyAdjustOrNot = ""
				
				sQuery = "Select isNull(AmountPaid,0)FROM Acc_T_CreatedPybleAdjDet WHERE CreatedTransNo = "& iCrTrNo &" "
				'Response.Write "<p>squery="&squery
				objRS.Open sQuery,con
				If Not objrs.EOF Then
					nOldAdjAmt = objrs(0)
				End IF
				objrs.Close 
				
				Response.Write "<p>nOldAdjAmt="&nOldAdjAmt
			
				If cdbl(nOldAdjAmt) > 0 Then
					sQuery = "Update Acc_T_CreatedPayables Set AmountPaid = (AmountPaid + "&nOldAdjAmt&") Where PayablesNumber = "&iCrRecNo &" and CreatedTransNo = "& iCrTrNo &" "
					Response.Write "<p>"&sQuery&"<br><br>"
					Con.Execute sQuery
			
					sQuery = "Update Acc_T_Payables Set AmountPaid = (AmountPaid + "&nOldAdjAmt&") Where createdPayablesNumber = "&iCrRecNo
					Response.Write "<p>"&sQuery&"<br><br>"
					Con.Execute sQuery
			
					sQuery = "DELETE FROM Acc_T_CreatedPybleAdjDet WHERE PayablesNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
					Response.Write "<p>"&sQuery &"<br><br>"
					con.Execute sQuery
			
					sQuery = "DELETE FROM Acc_T_PybleAdjustmentDetails WHERE PayablesNumber = "&iCrRecNo&" "
					Response.Write "<p>"&sQuery &"<br><br>"
					con.Execute sQuery
				End IF 
				
			End IF	'If sCheckAlreadyAdjustOrNot = "Y"
			'Response.Write "<p>iAccRecNo="&iAccRecNo
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then			
				
				If cdbl(dBillAmt) > "0.00" Then
			
					sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn,  "&_
							 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iCrRecNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'J', 1) "
							 Response.Write "<p>"&sQuery&"<br><br>"
					Con.Execute sQuery
					
					 
					sQuery = "INSERT INTO Acc_T_PybleAdjustmentDetails (PayablesNumber, PaidByTransactionNo, PaidOn,  "&_
							 "AmountPaid, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iAccRecNo&", "&iTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'J', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
				End IF
			End IF
		End IF
		
		IF CStr(sAdjType) = "R" Or CStr(sAdjType) = "C" Then
			
			'sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAccRecNo
			sQuery="select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "& iCrTrNo
			Response.Write "<p>sql="&sQuery
			objRS.Open sQuery,con
			IF Not objRS.EOF Then
				iCrRecNo = objRS(0)
			End IF
			objRS.Close
			
			squery = " select isNull(ReceivableNumber,0),TransactionNumber from Acc_T_Receivables where createdReceivable  = "&iCrRecNo&" "
			objRS2.Open squery,con
			If Not objRS2.EOF Then
				iAccRecNo = objrs2(0)
				iTrNo= objrs2(1)
			End IF
			objRS2.Close
			
			'Added Newly
			If sCheckAlreadyAdjustOrNot = "Y" and nBillDetCheck <> "" Then
				
				nOldAdjAmt = cdbl("0")
				
				sQuery = "Select isNull(AmountReceived,0) FROM Acc_T_CreatedRcvbleAdjDet WHERE ReceivableNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				objRS.Open sQuery,con
				If Not objrs.EOF Then
					nOldAdjAmt = objrs(0)
				End IF
				objrs.Close 
			
				Response.Write "<p>nOldAdjAmt="&nOldAdjAmt
			
				sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = (isNull(AmountReceived,0) + "&nOldAdjAmt&") Where ReceivableNumber = "&iCrRecNo &" "
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				sQuery = "Update Acc_T_Receivables Set AmountReceived = (isNull(AmountReceived,0) + "&nOldAdjAmt&") Where ReceivableNumber = "&iCrRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				sQuery = "DELETE FROM Acc_T_CreatedRcvbleAdjDet WHERE ReceivableNumber = "&iCrRecNo&" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
			
				sQuery = "DELETE FROM Acc_T_RcvblAdjustmentDetails WHERE ReceivableNumber = "&iAccRecNo&" and RecdByTransactionNo = "& iTrNo &" "
				Response.Write "<p>"&sQuery &"<br><br>"
				con.Execute sQuery
				
			End IF
			
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then
				If CDbl(dBillAmt) > "0.00" Then
					sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, ReceivedOn,  "&_
							 "AmountReceived, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iCrRecNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'J', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
			
					sQuery = "INSERT INTO Acc_T_RcvblAdjustmentDetails (ReceivableNumber, RecdByTransactionNo, ReceivedOn,  "&_
							 "AmountReceived, AdjustType, VoucherEntryNumber) "&_
							 "VALUES ("&iAccRecNo&", "&iTrNo&", Convert(Datetime,'"&sVouDate&"',103), "&dBillAmt&", "&_
							 "'J', 1) "
							 Response.Write sQuery&"<br><br>"
					Con.Execute sQuery
				End IF'If CDbl(dBillAmt) > "0.00" Then
			End IF
			
		End IF
		
		IF CStr(sAdjType) = "P" Or CStr(sAdjType) = "R" Then
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then			
				If cdbl(dBillAmt) > "0.00" Then
					sQuery = "INSERT INTO Acc_T_BillsClosingDetails (CreatedTransNo, TransactionNo, AdvCloseNo, " & _
							 "VoucherAmt, BillTrNo, PayRecNo, BillAmt, BillAdjAmt, CreatedBy, CreatedOn,PartyCode,OUDefinitionID,BillType) " & _
							 "VALUES ("&iCrTrNo&", "&iTrNo&", "&iAdvCloseNo&", "&dVouAmt&", "&iCrRecNo&","&iAccRecNo&", "&dBillAmt&", "&dBillAmt&", "&getUserID()&", getDate(),"&iParCode&", "&_
							 "'"&sUnitID&"','"&sAdjType&"') "
				End IF
			End If 'If sCheckAlreadyAdjustOrNot <> "" Then			
		Else
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" Then			
				If cdbl(dBillAmt) > "0.00" Then
				sQuery = "INSERT INTO Acc_T_BillsClosingDetails (CreatedTransNo, TransactionNo, AdvCloseNo, " & _
						 "VoucherAmt, BillTrNo, PayRecNo, BillAmt, BillAdjAmt, CreatedBy, CreatedOn,PartyCode,OUDefinitionID,BillType) " & _
						 "VALUES ("&iCrTrNo&", "&iTrNo&", "&iAdvCloseNo&", "&dVouAmt&", "&iBillTrNo&","&iAccRecNo&", "&dBillAmt&", "&dBillAmt&", "&getUserID()&", getDate(),"&iParCode&", "&_
						 "'"&sUnitID&"','"&sAdjType&"') "
				End IF
			End If 'If sCheckAlreadyAdjustOrNot <> "" Then			
				 
		End IF
		Response.Write sQuery
		Con.Execute sQuery
		
		IF CStr(sAdjType) = "I" Then
			
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" and cdbl(dBillAmt) > "0.00" Then			
				nToBeAdjustAmt = cdbl(nToBeAdjustAmt) + cdbl(dBillAmt)
				blnInsertAdvanceTable = True
				
				'sQuery = "Update Acc_T_Receivables Set AmountReceived = AmountReceivable Where ReceivableNumber = "&iAccRecNo
				sQuery = "Update Acc_T_Receivables Set AmountReceived = (AmountReceivable+"& dBillAmt&" )Where ReceivableNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				'sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = AmountReceivable Where ReceivableNumber = "&iCrRecNo
				sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = (AmountReceivable+"&dBillAmt&") Where ReceivableNumber = "&iCrRecNo
				Response.Write sQuery&"<br><br>"
				Con.Execute sQuery
			End IF
			
		End IF
		
		IF CStr(sAdjType) = "PI" or CStr(sAdjType) = "C" Then
			
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" and dBillAmt > "0.00" Then			
				
				nToBeAdjustAmt = cdbl(nToBeAdjustAmt) + cdbl(dBillAmt)
				blnInsertAdvanceTable = True
				
				'sQuery = "Update Acc_T_Payables Set AmountPaid = AmountPayable Where PayablesNumber = "&iAccRecNo
				sQuery = "Update Acc_T_Payables Set AmountPaid = (AmountPaid + "&dBillAmt&") Where PayablesNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				'sQuery = "Update Acc_T_CreatedPayables Set AmountPaid = AmountPayable Where PayablesNumber = "&iCrRecNo
				sQuery = "Update Acc_T_CreatedPayables Set AmountPaid = (AmountPaid + "&dBillAmt&") Where PayablesNumber = "&iCrRecNo &" and CreatedTransNo = "& iCrTrNo &" "
				Response.Write sQuery&"<br><br>"
				Con.Execute sQuery
				
			End IF	'If sCheckAlreadyAdjustOrNot <> "" Then
			
		End IF
		
		IF CStr(sAdjType) = "D" Then
			'sQuery = "Update Acc_T_Receivables Set AmountReceived = AmountReceivable Where ReceivableNumber = "&iAccRecNo
			'Response.Write sQuery&"<br><br>"
			'Con.Execute sQuery
			
			'sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = AmountReceivable Where ReceivableNumber = "&iCrRecNo
			'Response.Write sQuery&"<br><br>"
			'Con.Execute sQuery
			
			If sCheckAlreadyAdjustOrNot <> "" and nBillDetCheck <> "" and cdbl(dBillAmt) > "0.00" Then
				'nToBeAdjustAmt = cdbl(nToBeAdjustAmt) + cdbl(dBillAmt)
				'blnInsertAdvanceTable = True
				
				'sQuery = "Update Acc_T_Receivables Set AmountReceived = AmountReceivable Where ReceivableNumber = "&iAccRecNo
				'sQuery = "Update Acc_T_Receivables Set AmountReceived = (AmountReceivable+"& dBillAmt&" )Where ReceivableNumber = "&iAccRecNo
				sQuery = "Update Acc_T_Receivables Set AmountReceived = (AmountReceived+"& dBillAmt&" )Where ReceivableNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
				
				sQuery = "Select CreatedReceivable From Acc_T_Receivables where ReceivableNumber ="&iAccRecNo
				objRS.Open squery,con
				If Not objrs.EOF Then
					iCrRecNo = objRS("CreatedReceivable")
				End IF
				objrs.Close
			
				'sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = AmountReceivable Where ReceivableNumber = "&iCrRecNo
				'sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = (AmountReceivable+"&dBillAmt&") Where ReceivableNumber = "&iCrRecNo
				sQuery = "Update Acc_T_CreatedReceivables Set AmountReceived = (AmountReceived+"&dBillAmt&") Where ReceivableNumber = "&iCrRecNo
				Response.Write sQuery&"<br><br>"
				Con.Execute sQuery
			End IF
						
		End IF
		
		IF CStr(sAdjType) = "C" and 1= 2 Then
			sQuery = "Update Acc_T_Payables Set AmountPaid = AmountPayable Where PayablesNumber = "&iAccRecNo
			Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
			
			sQuery = "Update Acc_T_CreatedPayables Set AmountPaid = AmountPayable Where PayablesNumber = "&iCrRecNo
			Response.Write sQuery&"<br><br>"
			Con.Execute sQuery
			
		End IF
		
		IF CStr(sAdjType) = "P" Then
				
			nToBeAdjustAmt = cdbl(nToBeAdjustAmt) + cdbl(dBillAmt)
			'If cdbl(dBillAmt) > 0 Then	
			If nBillDetCheck <> "" and cdbl(dBillAmt) > 0 Then	
				'sQuery = "Update Acc_T_AdvancePayments Set AdvanceAdjusted = AdvancePaid Where AdvanceNumber = "&iAccRecNo
				sQuery = "Update Acc_T_AdvancePayments Set AdvanceAdjusted = (isnull(AdvanceAdjusted,0)+"& dBillAmt &") Where AdvanceNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				'sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = AdvancePaid Where CreatedAdvanceNo = "&iCrRecNo
				sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = (isnull(AdvanceAdjusted,0)+"& dBillAmt &") Where CreatedAdvanceNo = "&iCrRecNo
				Response.Write sQuery&"<br><br>"
				Con.Execute sQuery
			End IF			
		End IF
		
		IF CStr(sAdjType) = "R" Then
			If cdbl(dBillAmt) > 0 Then	
				sQuery = "Update Acc_T_AdvancePayments Set AdvanceAdjusted = (isNull(AdvanceReceived,0)+ "& dBillAmt &") Where AdvanceNumber = "&iAccRecNo
				Response.Write "<p>"&sQuery&"<br><br>"
				Con.Execute sQuery
			
				sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = (isNull(AdvanceReceived,0)+ "& dBillAmt &") Where CreatedAdvanceNo = "&iCrRecNo
				Response.Write sQuery&"<br><br>"
				Con.Execute sQuery
			End IF'If cdbl(dBillAmt) > 0 Then	
		End IF
		'Response.Write "<p>sCheckAlreadyAdjustOrNot="&sCheckAlreadyAdjustOrNot & "---"& nBillDetCheck	
		If sCheckAlreadyAdjustOrNot = "" and nBillDetCheck = "" Then			
			'Response.Write "<p>AmtToAdjustBBBB="&dBillAmt
			nAdjustAlreadyAmount = cdbl(nAdjustAlreadyAmount) + cdbl(dBillAmt)
		End IF
		
	Next
	'Response.Write "<p><Font color=red>sAdjType="&sAdjType
End IF

Response.Write "<p>sAdjType="&sAdjType

If sAdjType <> "" AND sAdjType <> "P" and sAdjType <> "R" and sAdjType <> "D" Then
	Response.Write "<p>VoucherAmount="&dVouAmt
	Response.Write "<p>nToBeAdjustAmt="&nToBeAdjustAmt
	Response.Write "<p>nAdjustAlreadyAmount="&nAdjustAlreadyAmount

	nTotalAmtToAdjust = cdbl(nToBeAdjustAmt) + cdbl(nAdjustAlreadyAmount)
	Response.Write "<p>nTotalAmtToAdjust="&nTotalAmtToAdjust


	If sAdjType = "P" or sAdjType = "R" or sAdjType = "D" Then
		sExp = "//PayRec/Party"
		Set TempNode = Root.selectNodes(sExp)
		If TempNode.length > 0 Then
			sPartyType =  TempNode.item(0).Attributes.getNamedItem("ParType").value
			sPartySubType = TempNode.item(0).Attributes.getNamedItem("ParSubType").value
			nPartyCode = TempNode.item(0).Attributes.getNamedItem("ParCode").value
		End IF
	Else
		sExp = "//PayRec/AccHead"
		Set TempNode = Root.selectNodes(sExp)
		If TempNode.length > 0 Then
			sTemp = split(TempNode.item(0).Attributes.getNamedItem("No").value,"?")
			sPartyType =  sTemp(0)
			sPartySubType = sTemp(1)
			nPartyCode = sTemp(3)
		End IF
	End IF

	'If cdbl(dVouAmt) >= cdbl(nTotalAmtToAdjust) and blnInsertAdvanceTable = True Then
	If cdbl(dVouAmt) >= cdbl(nTotalAmtToAdjust) Then
	
		blnExistAdvNo = False
			
		sQuery = "Select AdvanceNumber,CreatedAdvanceNo From Acc_T_AdvancePayments Where TransactionNumber = "& iTrNo &" and CreatedTransNo = "& iCrTrNo &" "
		'Response.Write "<p>sql="&sQuery&"<br><br>"
	
		objRs.Open sQuery,con
		IF Not objRs.EOF Then
			blnExistAdvNo = True
			iExistAdvNo = objRs(0)
			iExistCrAdvNo = objrs(1)
		End IF
		objRs.Close
		Response.Write "<p>blnExistAdvNo="&blnExistAdvNo
		sAdjNarr = "Adjusted to Purchase Invoice No "&sCrVouNo&" DT: "& sVouDate  &" Amt: "& nTotalAmtToAdjust
	
		If blnExistAdvNo Then
			If nAdjustAlreadyAmount = "" Then
				
				If nTotalAmtToAdjust <> "0" Then
				
					sQuery = " update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&nTotalAmtToAdjust &_
							 " where AdvanceNumber = "&iExistAdvNo
					Response.Write sQuery &"<br><br>"
					con.execute(sQuery)
			
					sQuery = " update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&nTotalAmtToAdjust &_
							 " where CreatedAdvanceNo = "&iExistCrAdvNo
			
					Response.Write sQuery &"<br><br>"
					con.execute(sQuery)
			
				
				sQuery = "INSERT INTO ACC_T_CREATEDADVANCEADJ (CREATEDADVANCENO, CREATEDTRANSNO, ADJUSTEDON,  "&_
						 "AMOUNTADJUSTED, NARRATION) VALUES "&_
						 "("& iExistCrAdvNo &", "&iCrTrNo&", Convert(Datetime,'"&sVouDate &"',103), "&nTotalAmtToAdjust&", '"&sAdjNarr&"') "

				'Response.Write sQuery &"<br><br>"
				'con.execute(sQuery)
				End IF
			End IF
		
			If nAdjustAlreadyAmount <> "" and nToBeAdjustAmt = "" Then
				sQuery = " update Acc_T_AdvancePayments set AdvancePaid=isnull(AdvancePaid,0)+"&nTotalAmtToAdjust &_
						 " where AdvanceNumber = "&iExistAdvNo
				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
		
				sQuery = " update Acc_T_CreatedAdvances set AdvancePaid=isnull(AdvancePaid,0)+"&nTotalAmtToAdjust &_
						 " where CreatedAdvanceNo = "&iExistCrAdvNo
		
				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
			End IF
	
			If nAdjustAlreadyAmount <> "" and nToBeAdjustAmt <> "" Then
				nTotalAmtToAdjust = cdbl(dVouAmt)- cdbl(nToBeAdjustAmt)
				
				sQuery = " update Acc_T_AdvancePayments set AdvancePaid ="& nTotalAmtToAdjust &"  " &_
						 " where AdvanceNumber = "&iExistAdvNo
				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
		
				sQuery = " update Acc_T_CreatedAdvances set AdvancePaid ="& nTotalAmtToAdjust &" " &_
						 " where CreatedAdvanceNo = "&iExistCrAdvNo
		
				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
				
			End If
		Else
			sQuery = "Select isNull(Max(AdvanceNumber),0)+1 from Acc_T_AdvancePayments "
			objRs.Open sQuery,Con
			If Not objRs.EOF Then
				iAdvNo = objRs(0)
			End IF
			objRs.Close
	
			sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
			objRs.Open sQuery,Con
			If Not objRs.EOF Then
				iCrAdvNo = objRs(0)
			End IF
			objRs.Close
		
			If cdbl(nToBeAdjustAmt) > "0" Then
		
				squery = " INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
						 " PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
						 " VALUES("&iAdvNo&", "& iTrNo &",'"&Session("organizationcode")&"','"&sPartyType &"',"&sPartySubType&","&_
						 " "&nPartyCode&","&dVouAmt&",("&dVouAmt&"-"&CDbl(nToBeAdjustAmt)&"),NULL,NULL,"&iCrAdvNo&","&iCrTrNo &")"
				Response.Write "<p>squery="&sQuery
				con.Execute sQuery
				
				'	sQuery = " INSERT INTO ACC_T_CREATEDADVANCEADJ (CREATEDADVANCENO, CREATEDTRANSNO, ADJUSTEDON,  "&_
				'		" AMOUNTADJUSTED, NARRATION) VALUES "&_
				'		 " ("&iCrAdvNo&", "&iCrTrNo&", Convert(Datetime,'"&sVouDate&"',103), ("&dVouAmt&"-"&CDbl(nToBeAdjustAmt)&"), '"&sAdjNarr&"') "
				'	Response.Write "<p>squery="&sQuery
				'	con.Execute sQuery
				
				squery = " INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
					 " PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
					 " VALUES("&iCrAdvNo&", "& iCrTrNo &",'"&Session("organizationcode")&"','"&sPartyType&"',"&sPartySubType&","&_
					 " "&nPartyCode&","&dVouAmt&",("&dVouAmt&"-"&CDbl(nToBeAdjustAmt)&"),NULL,NULL)"
				Response.Write "<p>squery="&sQuery
				con.Execute sQuery
			End IF	'If nToBeAdjustAmt <> "0" Then
		
		End IF	'If blnExistAdvNo Then
	End IF
End If 'If sAdjType <> "P" or sAdjType <> "R" Then

Con.RollbackTrans
Response.End

if con.Errors.count <> 0 then
	Response.Clear
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	Con.RollbackTrans
	Response.End
else
'	Con.RollbackTrans
'	Response.End 
	Response.Clear 
	Con.CommitTrans
End IF

IF CStr(iAdvNo) = "0" Then
	Response.Redirect("MANAGEBANKVOUCHERS.ASP")
Elseif sAdjType = "D" or sAdjType = "P" Then
	Response.Redirect("MANAGEPURCHASEVOUCHERS.ASP")
Elseif CStr(iTrNo) <> "0" Then
	Response.Redirect("BILLSCLOSEVIEW.ASP")
Elseif CStr(iTrNo) = "0" and CStr(iAdvNo) = "0" Then
	Response.Redirect("../reports/PayReceivableSelection.asp")
Else
	Response.Redirect("BILLSCLOSEVIEW.ASP")
End IF
	
%>