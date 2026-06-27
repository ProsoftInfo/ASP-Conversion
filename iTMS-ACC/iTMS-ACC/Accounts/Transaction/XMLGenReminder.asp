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
	'Program Name				:	XMLGenReminder.asp
	'Module Name				:	Accounts (Reports)
	'Modified By				:	UmaMaheswari S
	'Modified On				:	09th April 2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include file="../../include/Accpopulate.asp"-->
<%
dim objRs,objRs2,objRs3,sQuery,saTemp,dTotalBal,dTotAmtPaid
dim sOrgId,sOrgName,sTillDate,sPartySubType,sSubTypeName,iPartyCode,sVouDate
dim sPartyName,sGetVal,sPartyType,iInvoiceNo,sInvoiceDate,dAmount,dAmtPaid,dBalAmt,iSNo,iNoOfDays
Dim sTransType,sTransName,iDueDays,iCrTransNo
Dim iPartyCheck,dTotRec,dParOpenAmt,dParCloseAmt,sCheck,sOpenCD,sCloseCd
Dim objDOM,Root,PartyNode,DetNode,CourNode

iSNo=0
set objRs  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
set objRs3  = server.CreateObject("adodb.recordset")


sOrgId = session("organizationcode")
set objDOM = Server.CreateObject("Microsoft.XMLDOM")
set Root = objDOM.CreateElement("Reminder")

Root.setAttribute("SENDBY"),""
Root.setAttribute("NAME"),""
Root.setAttribute("ID"),""
Root.setAttribute("ADDRESS"),""

Dim nJJ,iPartyCodeArr,sPartySubTypeArr,sPartyTypeArr,sPartyDataArr
sGetVal = Request("sData")
sPartyDataArr = split(sGetVal,":")
iPartyCodeArr = split(sPartyDataArr(0),",")
sPartyTypeArr = split(sPartyDataArr(1),",")
sPartySubTypeArr = split(sPartyDataArr(2),",")

For nJJ = LBound(iPartyCodeArr) To UBound(iPartyCodeArr)

	iPartyCode = iPartyCodeArr(nJJ)
	sPartyType = sPartyTypeArr(nJJ)
	sPartySubType = sPartySubTypeArr(nJJ)
	
	IF CStr(iPartyCode) <> "0" Then
		sPartyName = GetPartyName(iPartyCode)
	End IF
	
	set PartyNode = objDOM.createElement("Party")
	Root.appendchild PartyNode

	PartyNode.setAttribute("CODE"),iPartyCode 
	PartyNode.setAttribute("TYPE"),sPartyType  
	PartyNode.setAttribute("SUBTYPE"),sPartySubType
	PartyNode.setAttribute("NAME"),sPartyName

	IF CStr(iPartyCode) = "0" Then

		sQuery = "Select O.PartyCode,P.PartyName From APP_R_OrgParty O,App_M_PartyMaster P Where PartySubType = "&sPartySubType&" "&_
				 "and PartyType = '"&sPartyType&"' and OUDefinitionID = '"&sOrgId &"' and O.PartyCode = P.PartyCode Order By P.PartyName "
				 
		With Objrs2
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = Con
			.Source = sQuery
			.Open
		End With

		iSNo = 1

		Set objRs2.ActiveConnection = Nothing

		Do While Not objRs2.EOF 
			iPartyCode = objRs2(0)
			sPartyName = objRs2(1)
			iPartyCheck = "1"
			
			sCheck = "F"
		
			sQuery = "Select P.AdvanceReceived - isNull(P.AdvanceAdjusted,0),H.TransactionType, Convert(Char,H.VoucherDate,103),H.CrDrIndication, isNull(P.AdvanceAdjusted,0), "&_
					 "P.AdvanceReceived, VoucherNumber,Convert(Char,VoucherDate,103),H.TransactionNumber "&_
					 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
					 "P.PartySubType = "&sPartySubType&" and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
					 "H.TransactionNumber = P.TransactionNumber and P.AdvanceReceived is Not Null and P.AdvanceReceived - isNull(P.AdvanceAdjusted,0) > 0"&_
					 "and H.TransactionType IN ('SJR','DNR') "
					 
					 
			with objRs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection=con
				.Source =sQuery
				.Open
			End with
		
			set objRS.ActiveConnection =nothing
		
			Do While Not objRs.EOF
				
				set DetNode = objDOM.CreateElement("DETAILS")
				DetNode.setAttribute("SNO"),iSNo
				DetNode.setAttribute("DOCTYPE"),trim(objRs(1))
				DetNode.setAttribute("INVOICENO"),Trim(objRs(6))
				DetNode.setAttribute("DATE"),trim(objRs(7))
				DetNode.setAttribute("ACCOUNTEDON"),trim(objRs(2))
				DetNode.setAttribute("AMOUNT"),FormatNumber(objRs(5),2,,,0)
				DetNode.setAttribute("AMOUNTPAIDTILLDATE"),FormatNumber(objRs(4),2,,,0)
				DetNode.setAttribute("BALANCE"),FormatNumber(objRs(0),2,,,0)
				DetNode.setAttribute("NOOFDAYSOUT"),""
				DetNode.setAttribute("NOOFDAYSOVER"),""
				DetNode.setAttribute("TRANSACTIONNO"),objrs(8)
				DetNode.setAttribute "SELECT","N"
				PartyNode.appendChild DetNode
				
					
				iSNo = iSNo + 1
				objRs.MoveNext
			loop
			objRs.Close	 
	
			sQuery="Select PartyInvoiceNumber,convert(char,PartyInvoiceDate,103),AmountReceivable,"&_
					" AmountReceived,0,"&_
					" convert(char,VoucherDate,103),isNull(TransactionNumber,0) from Acc_T_Receivables"&_
					" where OuDefinitionId='"&sOrgId &"' and PartyCode="&iPartyCode &" "&_
					" and PartySubType="&sPartySubType&"  and partyType='"&sPartyType &"' "&_
					" and AmountReceivable>AmountReceived"
					

			'Response.Write "2="& sQuery
			with objRs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection=con
				.Source =sQuery
				.Open
			End with
			set objRS.ActiveConnection =nothing
		
		

			while not objRs.EOF
				iSNo =iSNo +1 
				sCheck = "T"
				iInvoiceNo=objRs(0)
				sInvoiceDate=objRs(1)
				dAmount=objRs(2)
				dAmtPaid=objRs(3)
				iNoOfDays=objRs(4)
				sVouDate=objRs(5)
				dBalAmt=cdbl(dAmount)-cdbl(dAmtPaid)
				
				sQuery = "Select TransactionType,VoucherNumber,Convert(Char,VoucherDate,103),CreatedTransNo From Acc_T_VoucherHeader Where TransactionNumber = "&objRs(6)&" and TransactionType IN ('SJR','DNR') " 
				
				with objRs3
					.CursorLocation =3
					.CursorType =3
					.ActiveConnection=con
					.Source =sQuery
					.Open
				End with
				set objRs3.ActiveConnection =nothing
				IF Not objRs3.EOF Then
					sTransType = objRs3(0)
					iInvoiceNo = objRs3(1)
					sInvoiceDate = objRs3(2)
					iCrTransNo = objRs3(3)
				End IF
				objRs3.Close
				
				IF CStr(sTransType) = "SJR" Then
					iDueDays = GetDueDays(iCrTransNo,iNoOfDays,objRs(6),sInvoiceDate)			
				Else
					iDueDays = 0
				End IF
				
				set DetNode = objDOM.CreateElement("DETAILS")
				DetNode.setAttribute("SNO"),iSNo
				DetNode.setAttribute("DOCTYPE"),trim(sTransType)
				DetNode.setAttribute("INVOICENO"),trim(iInvoiceNo)
				DetNode.setAttribute("DATE"),trim(sInvoiceDate)
				DetNode.setAttribute("ACCOUNTEDON"),trim(sVouDate)
				DetNode.setAttribute("AMOUNT"),FormatNumber(dAmount,2,,,0)
				DetNode.setAttribute("AMOUNTPAIDTILLDATE"),FormatNumber(dAmtPaid,2,,,0)
				DetNode.setAttribute("BALANCE"),FormatNumber(dBalAmt,2,,,0)
				DetNode.setAttribute("NOOFDAYSOUT"),iNoOfDays
				DetNode.setAttribute("NOOFDAYSOVER"),iDueDays
				DetNode.setAttribute("TRANSACTIONNO"),objrs(6)
				DetNode.setAttribute "SELECT","N"
				PartyNode.appendChild DetNode
				
				objRs.MoveNext 
			Wend
			objRs.Close
		
			
			objRs2.MoveNext
		loop
		objRs2.Close
		
		
	Else

		Dim sCrDrInvNo,sCrDrInvDate
		iSNo = 1
		Dim sQuery1
		
		sQuery = "Select ClosingAmount,OpeningAmount From Acc_T_PartyOpeningAmt Where PartyType = '"&sPartyType&"' and  "&_
				     "PartySubType = "&sPartySubType&" and PartyCode = "&iPartyCode&" and OUDefinitionID = '"&sOrgId&"' "
				     
			'Response.Write "3="& sQuery
			Objrs2.Open sQuery,Con
			IF Not Objrs2.EOF Then
				dParCloseAmt = objRs2(0)
				dParOpenAmt = Objrs2(1)
			Else
				dParCloseAmt = 0
				dParOpenAmt = 0
			End IF
			Objrs2.Close
			
			
		iPartyCheck = "2"
		IF CStr(Trim(sPartySubType)) = "" Then
		
			'Taking the Party Opening AMount	
			
			
			sQuery1 = "Select O.PartyCode,P.PartyName,O.PARTYSUBTYPE From APP_R_OrgParty O,App_M_PartyMaster P Where "&_
				      " PartyType = '"&sPartyType&"' and OUDefinitionID = '"&sOrgId &"' and O.PartyCode = "& iPartyCode &" and O.PartyCode = P.PartyCode "
			
			With Objrs2
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = Con
				.Source = sQuery1
				.Open
			End With

			Set objRs2.ActiveConnection = Nothing
			Do While Not objRs2.EOF
				sPartySubType = sPartySubType &","&Trim(Objrs2(2))
			objRs2.MoveNext
			loop
			objRs2.Close
			sPartySubType = Mid(Trim(sPartySubType),2)
		End IF
				
			sQuery = "Select P.AdvanceReceived - isNull(P.AdvanceAdjusted,0),H.TransactionType, Convert(Char,H.VoucherDate,103),H.CrDrIndication, isNull(P.AdvanceAdjusted,0), "&_
					 "P.AdvanceReceived, VoucherNumber,Convert(Char,VoucherDate,103) ,H.TransactionNumber "&_
					 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
					 "P.PartySubType in("&sPartySubType&") and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
					 "H.TransactionNumber = P.TransactionNumber and P.AdvanceReceived is Not Null and P.AdvanceReceived - isNull(P.AdvanceAdjusted,0) > 0"&_
					 " and H.TransactionType IN ('SJR','DNR')  "
				
			'Response.Write "5="& sQuery
			
					 
			with objRs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection=con
				.Source =sQuery
				.Open
			End with
			set objRS.ActiveConnection =nothing
			Do While Not objRs.EOF			
				
				set DetNode = objDOM.CreateElement("DETAILS")
				DetNode.setAttribute("SNO"),iSNo
				DetNode.setAttribute("DOCTYPE"),trim(objRs(1))
				DetNode.setAttribute("INVOICENO"),trim(objRs(6))
				DetNode.setAttribute("DATE"),trim(objRs(7))
				DetNode.setAttribute("ACCOUNTEDON"),trim(objRs(2))
				DetNode.setAttribute("AMOUNT"),FormatNumber(objRs(5),2,,,0)
				DetNode.setAttribute("AMOUNTPAIDTILLDATE"),FormatNumber(objRs(4),2,,,0)
				DetNode.setAttribute("BALANCE"),FormatNumber(objRs(0),2,,,0)
				DetNode.setAttribute("NOOFDAYSOUT"),""
				DetNode.setAttribute("NOOFDAYSOVER"),""
				DetNode.setAttribute("TRANSACTIONNO"),objrs(8)
				DetNode.setAttribute "SELECT","N"
				PartyNode.appendChild DetNode
				
				iSNo = iSNo + 1
				
				objRs.MoveNext
			loop
			objRs.Close	 
			
			iSNo = iSNo - 1
				
			'objRs2.MoveNext
			'Loop		
			'End IF
			'objRs2.close
		dim iRcvbleNo
		
		sQuery = "Select PartyInvoiceNumber,convert(char,PartyInvoiceDate,103),AmountReceivable,"&_
				 " AmountReceived,0,"&_
				 " convert(char,VoucherDate,103),isNull(TransactionNumber,0),ReceivableNumber from Acc_T_Receivables"&_
				 " where OuDefinitionId='"&sOrgId &"' and PartyCode="&iPartyCode&" "&_
				 " and PartySubType IN ("&sPartySubType&")  and partyType='"&sPartyType &"' "&_
				 " and AmountReceivable>AmountReceived "
				 

		'Response.Write "4="& sQuery
		
		with objRs
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sQuery
			.Open
		End with
		set objRS.ActiveConnection =nothing

		while not objRs.EOF
			iSNo =iSNo +1 
			iInvoiceNo=objRs(0)
			sInvoiceDate=objRs(1)
			dAmount=objRs(2)
			dAmtPaid=objRs(3)
			iNoOfDays=objRs(4)
			sVouDate=objRs(5)
			iRcvbleNo=objRs(7)
			dBalAmt=cdbl(dAmount)-cdbl(dAmtPaid)
			
			sQuery = "Select TransactionType,VoucherNumber,Convert(Char,VoucherDate,103),CreatedTransNo From Acc_T_VoucherHeader Where TransactionNumber = "&objRs(6)&" and TransactionType IN ('SJR','DNR') " 
				
				'Response.Write "<BR><BR>A="&sQuery
				with objRs3
					.CursorLocation =3
					.CursorType =3
					.ActiveConnection=con
					.Source =sQuery
					.Open
				End with
				set objRs3.ActiveConnection =nothing
				IF Not objRs3.EOF Then
					sTransType = objRs3(0)
					iInvoiceNo = objRs3(1)
					sInvoiceDate = objRs3(2)
					iCrTransNo = objRs3(3)
				End IF
				objRs3.Close
				
				IF CStr(sTransType) = "SJR" Then
					iDueDays = GetDueDays(iCrTransNo,iNoOfDays,objRs(6),sInvoiceDate)			
				Else
					iDueDays = 0
				End IF
				
				set DetNode = objDOM.CreateElement("DETAILS")
				DetNode.setAttribute("SNO"),iSNo
				DetNode.setAttribute("DOCTYPE"),trim(sTransType)
				DetNode.setAttribute("INVOICENO"),trim(iInvoiceNo)
				DetNode.setAttribute("DATE"),trim(sInvoiceDate)
				DetNode.setAttribute("ACCOUNTEDON"),trim(sVouDate)
				DetNode.setAttribute("AMOUNT"),FormatNumber(dAmount,2,,,0)
				DetNode.setAttribute("AMOUNTPAIDTILLDATE"),FormatNumber(dAmtPaid,2,,,0)
				DetNode.setAttribute("BALANCE"),FormatNumber(dBalAmt,2,,,0)
				DetNode.setAttribute("NOOFDAYSOUT"),iNoOfDays
				DetNode.setAttribute("NOOFDAYSOVER"),iDueDays
				DetNode.setAttribute("TRANSACTIONNO"),objrs(6)
				DetNode.setAttribute "SELECT","N"
				PartyNode.appendChild DetNode
		
			objRs.MoveNext 
		Wend
		objRs.Close
		
		dTotalBal = abs(CDbl(dTotalBal) - CDbl(dTotRec))
		
	End IF
Next


objDOM.appendChild Root
'objDOM.save server.MapPath("../Temp/Transaction/GenReminderXML.xml")

Response.ContentType = "text/xml"
Response.Write objDOM.xml


Function GetDueDays(iCrTrNo,iOutStdDays,iAccTrNo,sBillDate)
		Dim sQuery,sObjDueRs,iOthAppNo,iPayTerms,iPayNoDays,iPayCount,iTotalDueDay
		Dim iDurPer,iDurDays,sPayTillDate,sObjPayRs,iPayRecNo,iParPayAmt,iPayInvAmt
		Dim iPayAmtToCome 
		Set sObjDueRs = Server.CreateObject("ADODB.RecordSet")
		Set sObjPayRs = Server.CreateObject("ADODB.RecordSet")
		
		sQuery = "Select isNull(OtherApplnTransNo,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iCrTrNo
		'Response.Write sQuery
		sObjDueRs.Open sQuery,Con
		IF Not sObjDueRs.EOF Then
			iOthAppNo = sObjDueRs(0)
		End IF
		sObjDueRs.Close
		
		IF CStr(iOthAppNo) = "0" Then
			GetDueDays = 0
		Else
			sQuery = "Select InvPymtTerms From Sal_T_InvoiceHeader Where SaleTransactionNo = " & iOthAppNo
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayTerms = sObjDueRs(0)
			End IF
			sObjDueRs.Close
			
			'iPayTerms = 2
			
			sQuery = "Select Count(1) From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
			
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayCount = sObjDueRs(0)
			End IF
			sObjDueRs.Close
			
			
			
			IF iPayCount = 1 Then
				sQuery = "Select DueDay From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayNoDays = sObjDueRs(0)
				End IF
				sObjDueRs.Close
				
				'iPayNoDays = 300
				
				'Response.Write iAccTrNo &"   " & iPayNoDays &"<br>"
				iTotalDueDay = CDbl(iOutStdDays) - CDbl(iPayNoDays)
			Else
				iPayAmtToCome = 0
				sQuery = "Select ReceivableNumber,AmountReceivable From Acc_T_Receivables Where TransactionNumber = "&iAccTrNo
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayRecNo = sObjDueRs(0)
					iPayInvAmt = sObjDueRs(1)
				End IF
				sObjDueRs.Close
				
				sQuery = "Select DueDay,DuePercent From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				Do While Not sObjDueRs.EOF
					iDueDays = sObjDueRs(0)
					iDurPer = sObjDueRs(1)
					
					'Response.Write iPayInvAmt &"  " & iDurPer &"<br>"
					
					iPayAmtToCome = Cdbl((CDbl(iPayInvAmt) * CDbl(iDurPer))/100) + CDbl(iPayAmtToCome)
					
					
					sQuery = "Select Top 1 Convert(Varchar,DateAdd(day,"&iDueDays&",Convert(Datetime,'"&sBillDate&"',103)),103) From Acc_T_RcvblAdjustmentDetails  "
					'Response.Write sQuery
					sObjPayRs.Open sQuery,Con
					
					IF Not sObjPayRs.EOF Then
						sPayTillDate = sObjPayRs(0)
					End IF
					sObjPayRs.Close
					
					sQuery = "Select Sum(AmountReceived) From Acc_T_RcvblAdjustmentDetails Where ReceivableNumber = "&iPayRecNo&" and "&_
							 "Convert(Datetime,ReceivedOn,103) <=  Convert(Datetime,'"&sPayTillDate&"',103)  "
							
					'Response.Write sQuery &"<br>"
					sObjPayRs.Open sQuery,Con 
					IF sObjPayRs.EOF Then
						iParPayAmt = sObjPayRs(0)
					End IF
					sObjPayRs.Close
					
					'iParPayAmt = 119000
					'Response.Write iParPayAmt &"  " & iPayAmtToCome &"<br>"
					
					IF CDbl(iParPayAmt) >= CDbl(iPayAmtToCome) Then
						iTotalDueDay = 0
					Else
						iTotalDueDay = CDbl(iOutStdDays) - CDbl(iDueDays)
					End IF
					
					'Response.Write iTotalDueDay &"<br>"
					
					
							 
					sObjDueRs.MoveNext
				Loop
				sObjDueRs.Close
			End IF
		End IF
		
		IF iTotalDueDay <0 Then
			iTotalDueDay = 0
		End IF
		
		IF CStr(iTotalDueDay) = "" Then
			iTotalDueDay = 0
		End IF
		
		GetDueDays = iTotalDueDay
	End Function
%>