<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PRNBankReceiptPrint.asp
	'Module Name				:	Accounts
	'Author Name				:	SANJAI KUMAR
	'Modified By				:   S.Maheswari
	'Created On					:	3 August 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/GetOrganization.asp"-->
<!--#include virtual="/include/PrintFunctions.asp"-->
<%
'------------------------Declaration Constants -----------------------------
Dim sPrHeader,sPrFooter,sPgMargin,sPgBreak,iPgLineNo,iRecCount,sDisplayHead
Dim aiHeaderColWidth(4,9),isNo,iPageNo,sTemp,sText,sRetVal,sAccName
Dim objFSO, objTxt,iPgTot
dim iTotPagNo
sPrHeader=""
sPrFooter=""
sPgMargin="       "
sPgBreak= FormattPrint("PAGESKIP","")
iPgLineNo=72

Dim iTransNo,sPartyName,oDom,Root,sOrgId,sOrgName,iBookCode,sBookName,sVouCherType,sShtOrgId
Dim sVouDate,sVoucherDate,sPagetitle2,sExp,tempNode,sChkPayTo,sChkParName,iParCnt,iGLCnt
Dim ConsPgHeader1,sQuery,objRs,objRs1,sBankInsDet,EntryNode,HeaderNode,sPayTo,sAccNo,sAccType
Dim sInstrType,sBankInsNo, sBankInsName,sBankInsDrawnOn,sBankInsDt,sAuthorisedSign
Dim iVouNo,sVouType,iPartyCode,iCreatedBy,sCreatedOn,sVouStatus,sAmount,iVouAmount,iTotAdjAmt,iTotAdvAmount
Dim sTranIndication,sTranEntryIndication,iEntryNo,iHeadOfAcc,iHeadOfAccAmt,sNarr
Dim iPartyCtrlAcc,sHeadOfAccName,sEmpName, sAddressLine1, sAddressLine2,sCity,iAdjRecCount
Dim iAccUnitPartyCode, sVouNarration, sPayableAt, sUserName,sBankInsAmt,sDataHead
Dim sBillNum,sBillDate,ArrTemp
Dim iLineNo,i,sTempCtr

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
iTransNo=Request("Value")

sRetVal = GetVouchXML(iTransNo)

oDOM.Load server.MapPath(sRetVal)
'Response.end
'oDOM.load server.MapPath("../XmlData/Voucher/" & iTransNo & ".xml")

	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Transaction/"&Session.SessionID&"PRNReceipt.txt"))


'objtxt.write iTransNo
set Root=oDOM.documentElement

sOrgId=Root.Attributes.Item(0).nodeValue
sOrgName=Root.Attributes.Item(1).nodeValue
iBookCode=Root.Attributes.Item(2).nodeValue
sBookName=Root.Attributes.Item(3).nodeValue
sVouCherType=Root.Attributes.Item(4).nodeValue

sShtOrgId = Right(sOrgId,2)

IF Len(Trim(iBookCode)) = 1 Then
	iBookCode = "0"&Trim(iBookCode)
End IF
sVouDate = Root.Attributes.Item(5).nodeValue
sVoucherDate = sVouDate


IF CStr(sVouCherType) = "D" Then
	sPagetitle2 = "CHEQUE RECEIPT VOUCHER"
Else
	sPagetitle2 = "CHEQUE PAYMENT VOUCHER"
End IF

sExp = "//Entry"
Set tempNode = Root.selectNodes(sExp)
IF TempNode.Length <> 0 Then
	sChkPayTo = TempNode.Item(0).Attributes.Item(2).nodeValue
End IF

sExp = "//Entry/AccHead[@Type=""P""]"
Set tempNode = Root.selectNodes(sExp)
IF TempNode.Length <> 0 Then
	sChkParName = TempNode.Item(0).Attributes.Item(3).nodeValue
Else
	sChkPayTo = ""
End IF

sExp = "//AccHead[@Type=""P""]"
Set tempNode = Root.selectNodes(sExp)
iParCnt = tempNode.length

sExp = "//AccHead[@Type=""G""]"
Set tempNode = Root.selectNodes(sExp)
iGLCnt = tempNode.length




sQuery = "Select G.AccountsGroupName From Acc_M_AccountGroups G,Acc_R_ApplicableAccountHeads H, "&_
		 "Acc_R_OrgGLAccountHead O Where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='"&iBookCode&"' and  "&_
		 "H.BookAccountHead = O.AccountHead and O.AccountsGroupCode = G.AccountsGroupCode "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
end With

Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	sAccName = objRs(0)
End IF
objRs.Close

sExp = "//BankInstrumentDet"
Set tempNode = Root.selectNodes(sExp)
IF tempNode.length <> 0 Then
	IF Cstr(Trim(tempNode.Item(0).Attributes.getNamedItem("InsType").value)) = "C" Then
			sBankInsDet = "Cheque "
		Elseif Cstr(Trim(tempNode.Item(0).Attributes.getNamedItem("InsType").value)) = "D" Then
			sBankInsDet = "Demand Draft "
		Elseif Cstr(Trim(tempNode.Item(0).Attributes.getNamedItem("InsType").value)) = "B" Then
			sBankInsDet = "Bankers Cheque "
		Else
			sBankInsDet = "Telegraphic Transfer "
	End IF

	sBankInsDet = sBankInsDet&" No: "
	sBankInsDet = sBankInsDet& tempNode.Item(0).Attributes.getNamedItem("InsNo").value
	sBankInsDet = sBankInsDet& " "
	sBankInsDet = sBankInsDet& "Dt: "& tempNode.Item(0).Attributes.getNamedItem("InsDate").value
	sBankInsDet = sBankInsDet& " Drawn On "

	sBankInsDet = sBankInsDet & tempNode.Item(0).Attributes.getNamedItem("DrawnOn").value
End IF


For each EntryNode in Root.childNodes
	if EntryNode.nodeName="Entry" then
		sPayTo = EntryNode.Attributes.Item(2).nodeValue
	end if
	for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
			sAccNo = Split(HeaderNode.attributes.item(0).nodeValue,"?")
			sAccType=HeaderNode.Attributes.Item(4).nodeValue
		End IF
	next
Next


sQuery = "select Distinct CreatedVoucherNo,TransactionType,OUDefinitionID,PayToRecdFrom," &_
		"PartyCode,CreatedBy,Convert(varchar,CreatedOn,103),CreatedVouchStatus," &_
		"BankInstrumentType,BankInstrumentNo,Convert(varchar,BankInstrumentDate,103),DrawnOnBank, PayableAt,VoucherAmount,VoucherNarration, " &_
		"Reverse(SubString(Reverse(VoucherNarration),1,10)),SubString(VoucherNarration,1,abs(Len(isnull(VoucherNarration,''))-11)),Convert(Varchar(10),VoucherDate,103) VoucherDate " & _
		" from VW_Created_BankVoucherView where CreatedTransNo="& iTransNo

'Response.write sQuery
'Response.end

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iVouNo = objRs(0)
	sVouType = objRs(1)
	sOrgId = objRs(2)
	sPayTo = objRs(3)
	iPartyCode = objRs(4)
	iCreatedBy = objRs(5)
	sCreatedOn = objRs(6)
	sVouStatus = objRs(7)
	sInstrType = objRs(8)
	'sBankInsNo = objRs(9)
	'sBankInsDt = trim(objRs(10))
	'sPayableAt = objRs(11)
	'sBankInsDrawnOn = objRs(12)
	'sAmount = objRs(13)
	iVouAmount = objRs(13)
	sVouNarration = objRs(14)
	sBillNum = objRs(16)
	sBillDate = objRs(15)
	sVoucherDate = objRs("VoucherDate")

'Response.Write sVouNarration
'Response.End
	'Response.Write sVouType
	'If sInstrType = "C" Then
	'	sBankInsName = "Cheque "
	'Elseif sInstrType = "D" Then
	'	sBankInsName = "Demand Draft "
	'ElseIf sInstrType = "B" Then
	'	sBankInsName = "Bankers Cheque "
	'Else
	'	sBankInsName = "Telegraphic Transfer "
	'End If
end if
objRs.Close
'----------Newly added on July 2nd 2008 by S.Maheswari to fetch bank details from Acc_T_CreatedVoucherInstrumentDet table instead of taking from CreatedVoucherHeader table----
sQuery = "Select BankInstrumentNo,convert(Varchar,BankInstrumentDate,103),BankInstrumentType,InstrumentAmount,PayableAt,DrawnOnBank "&_
		" from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo = "&iTransNo&" Order by InstrumentEntryNo "
 ' Response.Write sQuery
 
objRs.Open sQuery,con
If Not objRs.EOF then
	Do while not objRs.EOF
		sBankInsNo		= sBankInsNo&","&objRs(0)
		sBankInsDt		= sBankInsDt&","&objRs(1)
		IF trim(sTemp) <> trim(objRs(2)) then
			sBankInsName	= sBankInsName&","&objRs(2)
		Else

			sBankInsName	= sBankInsName
		End IF
		sBankInsAmt	    = sBankInsAmt&","&objRs(3)
		sBankInsDrawnOn = sBankInsDrawnOn&","&objRs(5)
		sTemp = objRs(2)
		objRs.MoveNext
	loop
End If
objRs.Close
sBankInsNo		= mid(sBankInsNo,2)
sBankInsDt		= mid(sBankInsDt,2)
sBankInsName	= mid(sBankInsName,2)
sBankInsAmt	    = mid(sBankInsAmt,2)
sBankInsDrawnOn = mid(sBankInsDrawnOn,2)
'==================================================================================================

'

If trim(sVouStatus) = "010104" Then
	sQuery = "select VoucherNumber from ACC_T_VoucherHeader where CreatedTransNo="&iTransNo
	with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
	end with
	if not 	objRs.EOF then
		iVouNo = objRs(0)
	End If
	objRs.Close
End If
set objRs.ActiveConnection = nothing
If trim(sVouType) = "BAP" Then
	sTranIndication = "C"
	sTranEntryIndication = "D"
Else
	sTranIndication = "D"
	sTranEntryIndication = "C"
End If
sQuery = "Select VoucherEntryNumber,AccUnitAccountHead,Amount,VoucherNarration,isNull(AccUnitPartyCode,0) from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"'"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
'Response.write sQuery
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iEntryNo = objRs(0)
	iHeadOfAcc = objRs(1)
	iHeadOfAccAmt = FormatNumber(objRs(2),2,,,-2)
	sNarr = objRs(3)
	iPartyCtrlAcc = objRs(4)
end if
objRs.Close
'Response.Write iPartyCtrlAcc
IF iHeadOfAcc <> "" Then
	sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
Else
	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
End If
'Response.write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sHeadOfAccName = objRs(0)
end if
objRs.Close

'

If iPartyCtrlAcc <> "0" Then
	sQuery = "SELECT PARTYNAME,ADDRESSLINE1,ADDRESSLINE2,CITY FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc

	With Objrs
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.Open
	End With
	set objRs.ActiveConnection = nothing
	IF not objRs.EOF  Then
		sPartyName = Trim(Objrs(0))
		sAddressLine1 = Trim(objRs(1))
		sAddressLine2 = Trim(objRs(2))
		sCity = Trim(objRs(3))
	End If
	objRs.Close
End If

sQuery = "SELECT LoginId,UserName FROM DCS_User WHERE InternalUserID ="&iCreatedBy

With Objrs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
set objRs.ActiveConnection = nothing
IF not objRs.EOF  Then
	sEmpName = Trim(Objrs(0))
	sUserName = Trim(objrs(1))
End IF
objRs.Close




'No and Date
aiHeaderColWidth(0,0)=76
aiHeaderColWidth(0,1)=10

'Receivedwiththanks,thesumofrupees
aiHeaderColWidth(1,0)=14
aiHeaderColWidth(1,1)=70

'Particulars
aiHeaderColWidth(2,0)=12
aiHeaderColWidth(2,1)=3
aiHeaderColWidth(2,2)=2
aiHeaderColWidth(2,3)=20
aiHeaderColWidth(2,4)=2
aiHeaderColWidth(2,5)=7
aiHeaderColWidth(2,6)=2
aiHeaderColWidth(2,7)=12
aiHeaderColWidth(2,8)=7
aiHeaderColWidth(2,9)=18

'Rs,Prepared,Sign
aiHeaderColWidth(3,0)=16
aiHeaderColWidth(3,1)=21
aiHeaderColWidth(3,2)=8
aiHeaderColWidth(3,3)=15
aiHeaderColWidth(3,4)=3
aiHeaderColWidth(3,5)=22


isNo=0
iPageNo=1
iLineNo = 0
	dim nReceiptNo,dReceiptDate,sReceivedFrom,sAmountInWords,iSerialNo,sRefNo,sRefDate
	dim iAmtAdjusted,sPaymentBy,iChequeNo,dChequeDate,sChequeAt,iAmount,sPreparedBy,sCheckedBy



	'Assigning Hardcoded values for Variables
	nReceiptNo				= iVouNo
	dReceiptDate			= sVoucherDate
	sReceivedFrom			= sPayTo
	sAmountInWords		= AmountWords(iVouAmount)
	iSerialNo					= 1
	sRefNo					= sVouNarration
	sRefDate					= sBankInsDt
	iAmtAdjusted			= iAmount
	sPaymentBy			= sBankInsName
	iChequeNo				= sBankInsNo
	dChequeDate			= sBankInsDt
	sChequeAt				= sBankInsDrawnOn
	'iAmount					= iHeadOfAccAmt
	sPreparedBy			= sEmpName & "-" & sCreatedOn
	sCheckedBy				= "xCheckedBy"
	sAuthorisedSign		= sOrgName  '"xAuthorizedSign"

	sTemp = ""
	Header()



	'Particulars

	'Response.write iPartyCtrlAcc &"======="
	


	'''''''''''''''''''''''''''Blocked by S.Maheswari on Nov 1st 2008 '''''''''''''''''''
	'sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")							'
	'sTemp = sTemp & myAlign(iSerialNo,aiHeaderColWidth(2,1),"L")						'
	'sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")								'
	'IF Cstr(Trim(iPartyCtrlAcc)) <> "0" Then											'
	'	sTemp = sTemp & myAlign(sBillNum,11,"L")										'
	'	sTemp = sTemp & myAlign(" ",5,"L")												'
	'	sTemp = sTemp & myAlign(sBillDate,10,"L")										'
	'Else																				'
	'	sTemp = sTemp & myAlign(sVouNarration,26,"L")									'
	'End if																				'
																						'
	'sTemp = sTemp & myAlign("",4,"L")													'
	'sTemp = sTemp & myAlign(Formatnumber(sAmount,2,,,0),aiHeaderColWidth(2,7),"L")		'
	'sTemp = sTemp & myAlign("",aiHeaderColWidth(2,8)-2,"L")							'
	'sTemp = sTemp & myAlign(sBankInsNo,aiHeaderColWidth(2,9),"L")						'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'------------------Newly added by S.Maheswari on Nov 1st 2008 ----------------------

	dim iRcvbleNo,iRCtr,sBkInsNoflag,sBkInsDtflag,sBkInsPayAtflag,sBkInsDrOnflag,iCtr
	
	sQuery = "Select ReceivableNumber,IsNull(AmountReceived,0) from acc_T_createdRcvbleAdjDet where CreatedTransNo = "& iTransNo
	objRs.Open sQuery,con
	iRCtr = 0
	iCtr = 0
	iTotAdjAmt = 0
	 If objRs.EOF then
		sQuery = "select isNull(VoucherNarration,''),Amount from Acc_T_CreatedVoucherDetails where  CreatedTransNo = "&iTransNo
		'Response.Write "<p> " & sQuery
		'Response.End 
		objRs1.Open sQuery,con
		if not objRs1.EOF then

			sVouNarration = trim(objRs1(0))
			iAmount = objRs1(1)

			if trim(sReceivedFrom) = "" then
				ArrTemp = BreakString(sVouNarration,25)
			else
				ArrTemp = BreakString("Adv.Received:",25) 
			end if 	
			
	
			iAdjRecCount = UBOUND(ArrTemp)
			For i = 0 to UBOUND(ArrTemp)
				IF trim(ArrTemp(i)) <> "" then
					iCtr = iCtr + 1
					sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
					sTemp = sTemp & myAlign("",aiHeaderColWidth(2,1),"L")
					sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")
 					sTemp = sTemp & myAlign(ArrTemp(i),25,"L")
 					sTemp = sTemp & myAlign(" ",1,"L")
					sTemp = sTemp & myAlign(" ",3,"L")
					If iCtr = 1 Then
						sTemp = sTemp & myAlign(Formatnumber(iAmount,2,,,0),13,"R")
						sTemp = sTemp & myAlign(" ",5,"L")
					else
						sTemp = sTemp & myAlign(" ",13,"R")
						sTemp = sTemp & myAlign(" ",5,"L")
					End IF

					If iCtr = 1 Then
						if sBankInsNo="0" then
							sTemp = sTemp & myAlign("",20,"L")
						else
							sTemp = sTemp & myAlign(sBankInsNo,20,"L")
						end if
						sBkInsNoflag = True
					elseIf iCtr = 2 Then
						sTemp = sTemp & myAlign(sBankInsDt,20,"L")
						sBkInsDtflag = True
					elseIf iCtr = 3 Then
						sTemp = sTemp & myAlign(sPayableAt,20,"L")
						sBkInsPayAtflag = True
					elseIf iCtr = 4 Then
						sTemp = sTemp & myAlign(sBankInsDrawnOn,20,"L")
						sBkInsDrOnflag = True
					end if
					sTemp = sTemp & vbCrLf
					iLineNo = iLineNo +1
					'Response.Write sNarr(i)&":::"&iLineNo&"<BR>"
				End IF
			Next
		end if
		objRs1.Close
		
		sText = sText & sTemp
		sText = sText & vbCrLf	
		iLineNo = iLineNo +1
		sTemp = ""
	else

		sTempCtr = 0

		sTemp = ""
		iAdjRecCount = objRs.RecordCount

		do while not objRs.eof
		 
			iRCtr = iRCtr + 1
			sTempCtr = sTempCtr + 1

			if sTempCtr > 7 then
				 
				
				
				sText = sText & sTemp
			 	sText = sText & vbCrLf & vbCrLf & vbCrLf & vbCrLf 
			 	sText = sText & MYALIGN("Contnd....",95,"R") & vbCrLf
			 	sText = sText & vbCrLf & vbCrLf & vbCrLf & vbCrLf
			 	sTemp  = ""
			 	iLineNo = 1
			 	iPageNo = iPageNo + 1
			 	Header()
			 	sTempCtr = 1
		  
		   end if


			iRcvbleNo = objRs(0)
			sQuery = "Select isNull(PartyInvoiceNumber,''),Convert(varchar,isnull(PartyInvoiceDate,''),103),AmountReceived from Acc_T_CreatedReceivables where ReceivableNumber = "& iRcvbleNo
			objRs1.Open sQuery,con


			If Not objRs1.EOF then
				sBillNum  = objrs1(0)
				sBillDate = objrs1(1)
				sAmount   = objrs(1) 'added on 23rd June 2009 instead of objrs1(2)

				iTotAdjAmt = iTotAdjAmt + CDbl(sAmount)

				sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
				sTemp = sTemp & myAlign(iRCtr,aiHeaderColWidth(2,1),"L")
				sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")

	 			sTemp = sTemp & myAlign(sBillNum,11,"L")
	 			sTemp = sTemp & myAlign(" ",5,"L")
				sTemp = sTemp & myAlign(sBillDate,10,"L")
	 			sTemp = sTemp & myAlign(" ",3,"L")
				sTemp = sTemp & myAlign(Formatnumber(sAmount,2,,,0),13,"R")
				sTemp = sTemp & myAlign(" ",5,"L")

				If iRCtr = 1 Then
					if sBankInsNo="0" then
						sTemp = sTemp & myAlign("",20,"L")
					else
						sTemp = sTemp & myAlign(sBankInsNo,20,"L")
					end if
					sBkInsNoflag = True
				elseIf iRCtr = 2 Then
					sTemp = sTemp & myAlign(sBankInsDt,20,"L")
					sBkInsDtflag = True
				elseIf iRCtr = 3 Then
					sTemp = sTemp & myAlign(sPayableAt,20,"L")
					sBkInsPayAtflag = True
				elseIf iRCtr = 4 Then
					sTemp = sTemp & myAlign(sBankInsDrawnOn,20,"L")
					sBkInsDrOnflag = True
				end if
			
				Objrs.MoveNext
				
				If objrs.EOF Then
					iTotAdvAmount = (iVouAmount-iTotAdjAmt)
					sTempCtr = sTempCtr + 1
					'iTotAdvAmount = 1
					Select Case iAdjRecCount

					Case 1 :

						'sTemp = sTemp & myAlign(sBankInsNo,20,"L")
						'sBkInsNoflag = True
						'===============Print Advance================
						If iTotAdvAmount > 0 Then
							sTemp = sTemp & vbCrLf
							iLineNo = iLineNo +1
							if trim(sReceivedFrom) = "" then
								sDataHead = sVouNarration 
							else
								sDataHead = "Adv. Received:"
							end if
							
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
							sTemp = sTemp & myAlign(iRCtr+1,aiHeaderColWidth(2,1),"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")

							sTemp = sTemp & myAlign(sDataHead,16,"L")
							'sTemp = sTemp & myAlign(" ",5,"L")
							sTemp = sTemp & myAlign(" ",10,"L")
							sTemp = sTemp & myAlign(" ",3,"L")
							sTemp = sTemp & myAlign(Formatnumber((iVouAmount-iTotAdjAmt),2,,,0),13,"R")
							sTemp = sTemp & myAlign(" ",5,"L")
							'===============Print Advance Ends ===========

							sTemp = sTemp & myAlign(sBankInsDt,20,"L")
							sBkInsDtflag = True
						End If
					Case 2 :

						'sTemp = sTemp & myAlign(sBankInsDt,20,"L")
						'sBkInsDtflag = True
						'===============Print Advance================
						If iTotAdvAmount > 0 Then
							sTemp = sTemp & vbCrLf
							iLineNo = iLineNo +1
							if trim(sReceivedFrom) = "" then
								sDataHead = sVouNarration 
							else
								sDataHead = "Adv. Received:"
							end if 	
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
							sTemp = sTemp & myAlign(iRCtr+1,aiHeaderColWidth(2,1),"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")

							sTemp = sTemp & myAlign(sDataHead,16,"L")
							'sTemp = sTemp & myAlign(" ",5,"L")
							sTemp = sTemp & myAlign(" ",10,"L")
							sTemp = sTemp & myAlign(" ",3,"L")
							sTemp = sTemp & myAlign(Formatnumber((iVouAmount-iTotAdjAmt),2,,,0),13,"R")
							sTemp = sTemp & myAlign(" ",5,"L")
							'===============Print Advance Ends ===========

							sTemp = sTemp & myAlign(sPayableAt,20,"L")
							sBkInsPayAtflag = True

							
						End If

					Case 3 :

						'sTemp = sTemp & myAlign(sPayableAt,20,"L")
						'sBkInsPayAtflag = True
						'===============Print Advance================
						If iTotAdvAmount > 0 Then
							sTemp = sTemp & vbCrLf
							iLineNo = iLineNo +1
							if trim(sReceivedFrom) = "" then
								sDataHead = sVouNarration 
							else
								sDataHead ="Adv. Received:"
							end if 	
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
							sTemp = sTemp & myAlign(iRCtr+1,aiHeaderColWidth(2,1),"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")

							sTemp = sTemp & myAlign(sDataHead,16,"L")
							'sTemp = sTemp & myAlign(" ",5,"L")
							sTemp = sTemp & myAlign(" ",10,"L")
							sTemp = sTemp & myAlign(" ",3,"L")
							sTemp = sTemp & myAlign(Formatnumber((iVouAmount-iTotAdjAmt),2,,,0),13,"R")
							sTemp = sTemp & myAlign(" ",5,"L")
							'===============Print Advance Ends ===========

							sTemp = sTemp & myAlign(sBankInsDrawnOn,20,"L")
							sBkInsDrOnflag = True

							iLineNo = iLineNo +1
						End If

					Case 4 :

						'sTemp = sTemp & myAlign(sBankInsDrawnOn,20,"L")
						'sBkInsDrOnflag = True
						'===============Print Advance================
						If iTotAdvAmount > 0 Then
							sTemp = sTemp & vbCrLf
							iLineNo = iLineNo +1
							if trim(sReceivedFrom) = "" then
								sDataHead = sVouNarration 
							else
								sDataHead = "Adv.Received:"
							end if 
								
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
							sTemp = sTemp & myAlign(iRCtr+1,aiHeaderColWidth(2,1),"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")

							sTemp = sTemp & myAlign(sDataHead,16,"L")
							'sTemp = sTemp & myAlign(" ",5,"L")
							sTemp = sTemp & myAlign(" ",10,"L")
							sTemp = sTemp & myAlign(" ",3,"L")
							sTemp = sTemp & myAlign(Formatnumber((iVouAmount-iTotAdjAmt),2,,,0),13,"R")
							sTemp = sTemp & myAlign(" ",5,"L")
							'===============Print Advance Ends ===========

							'sBkInsDrOnflag = True
						
						End If

					Case Else :
						if sTempCtr > 7 And iTotAdvAmount > 0 then
							sText = sText & sTemp
							sText = sText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
							sTemp  = ""
							iLineNo = 1
							Header()
							sTempCtr = 1


							If iTotAdvAmount > 0 Then
								'sTemp = sTemp & vbCrLf
								if trim(sReceivedFrom) = "" then
									sDataHead = sVouNarration 
								else
									sDataHead = "Adv.Received:"
								end if 	
								
								sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
								sTemp = sTemp & myAlign(iRCtr+1,aiHeaderColWidth(2,1),"L")
								sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")

								sTemp = sTemp & myAlign(sDataHead,16,"L")
								'sTemp = sTemp & myAlign(" ",5,"L")
								sTemp = sTemp & myAlign(" ",10,"L")
								sTemp = sTemp & myAlign(" ",3,"L")
								sTemp = sTemp & myAlign(Formatnumber((iVouAmount-iTotAdjAmt),2,,,0),13,"R")
								sTemp = sTemp & myAlign(" ",5,"L")
								'===============Print Advance Ends ===========

								'sBkInsDrOnflag = True

								'iLineNo = iLineNo +1
							End If


							sTemp = sTemp & myAlign(sBankInsNo,20,"L") & vbCrLf

							iLineNo = iLineNo + 1


							sTemp = sTemp & myAlign(" ",74,"L") & myAlign(sBankInsDt,21,"L") & vbCrLf
							iLineNo = iLineNo + 1

							sTemp = sTemp & myAlign(" ",74,"L") & myAlign(sPayableAt,20,"L") & vbCrLf
							iLineNo = iLineNo + 1

							sTemp = sTemp & myAlign(" ",74,"L") & myAlign(sBankInsDrawnOn,20,"L") & vbCrLf
							iLineNo = iLineNo + 1

						end if 'if sTempCtr > 7 And iTotAdvAmount > 0 then

						'If iRCtr = 1 Then
						'	sTemp = sTemp & myAlign(sBankInsNo,20,"L")
						'	sBkInsNoflag = True
						'elseIf iRCtr = 2 Then
						'	sTemp = sTemp & myAlign(sBankInsDt,20,"L")
						'	sBkInsDtflag = True
						'elseIf iRCtr = 3 Then
						'	sTemp = sTemp & myAlign(sPayableAt,20,"L")
						'	sBkInsPayAtflag = True
						'elseIf iRCtr = 4 Then
						'	sTemp = sTemp & myAlign(sBankInsDrawnOn,20,"L")
						'	sBkInsDrOnflag = True
						'end if

					End Select
					objRs.MovePrevious
				Else
					objRs.MovePrevious
				End If
				
				sTemp = sTemp & vbCrLf

				iLineNo = iLineNo +1


			End If
			objRs1.Close

			objrs.MoveNext
		loop
	end if 	
	objrs.Close

	'Print the balance Advance amount after adjustment
	If sBkInsNoflag <> True OR iAdjRecCount = 0 Then
	
		if trim(sReceivedFrom) = "" then
			sDataHead = sVouNarration 
		else
			sDataHead = "Adv. Received:"
		end if 
		
		
			
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
		sTemp = sTemp & myAlign(iRCtr+1,aiHeaderColWidth(2,1),"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")

		sTemp = sTemp & myAlign(sDataHead,16,"L")
		'sTemp = sTemp & myAlign(" ",5,"L")
		sTemp = sTemp & myAlign(" ",10,"L")
		sTemp = sTemp & myAlign(" ",3,"L")
		sTemp = sTemp & myAlign(Formatnumber((iVouAmount-iTotAdjAmt),2,,,0),13,"R")
		sTemp = sTemp & myAlign(" ",5,"L")
		'sTemp = sTemp & vbCrLf
		'iLineNo = iLineNo +1
	End If


	IF sBkInsNoflag <> True OR iAdjRecCount = 0 then
		If iAdjRecCount = 0 Then
			if sBankInsNo="0" then
				sTemp = sTemp & myAlign("",20,"L") & vbCrLf
			else
				sTemp = sTemp & myAlign(sBankInsNo,20,"L") & vbCrLf
			end if
		Else
			if sBankInsNo="0" then
				sTemp = sTemp & myAlign(" ",74,"L") & myAlign(sBankInsNo,20,"L") & vbCrLf
			else
				sTemp = sTemp & myAlign(" ",74,"L") & myAlign("",20,"L") & vbCrLf
			end if
		End If
		iLineNo = iLineNo + 1
	End IF
	IF sBkInsDtflag <> True OR iAdjRecCount = 0 then
		sTemp = sTemp & myAlign(" ",74,"L") & myAlign(sBankInsDt,21,"L") & vbCrLf
		iLineNo = iLineNo + 1
	End IF
	IF sBkInsPayAtflag <> True OR iAdjRecCount = 0 then
		sTemp = sTemp & myAlign(" ",74,"L") & myAlign(sPayableAt,20,"L") & vbCrLf
		iLineNo = iLineNo + 1
	End IF

	IF sBkInsDrOnflag <> True OR iAdjRecCount = 0 then
		sTemp = sTemp & myAlign(" ",74,"L") & myAlign(sBankInsDrawnOn,20,"L") & vbCrLf
		iLineNo = iLineNo + 1
	End IF
'Response.write iLineNo
'Response.end
	'---------------------------------------------------------------------------------------

	for i = iLineNo to 25 ' 27
		sTemp = sTemp & vbCrLf
		iLineNo = iLineNo + 1
	Next
	'Rs,Prepared,Sign
	sTemp = sTemp & vbCrLf
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+11,"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,1),"L")
	sTemp = sTemp & myAlign("",7,"L")
	sTemp = sTemp & myAlign(sUserName,aiHeaderColWidth(3,3),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L")
'
	sText = sText & sTemp
	sText = sText &" " & vbCrLf
	sTemp = ""

	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+11,"L")
	sTemp = sTemp & myAlign(FormatNumber(iVouAmount,2,,,0),aiHeaderColWidth(3,1),"L")
	sTemp = sTemp & myAlign("",7,"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,3),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L")

	sText = sText & sTemp
	sText = sText &" " & vbCrLf
	sTemp = ""

	'Rs,Prepared,Sign
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,1),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,2)+10,"L")
	sTemp = sTemp & myAlign(sUserName,aiHeaderColWidth(3,3),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L")

	sText = sText & sTemp
	sText = sText &" " & vbCrLf & vbCrLf & vbCrLf & vbCrLf& vbCrLf& vbCrLf
	sTemp = ""

	'sText = sText & chr(12)
	objTxt.write sText
	Response.Redirect "../../Components/FormattPrintNew.asp?server=server&filepath=/Accounts/temp/transaction/"&Session.SessionID & "PRNReceipt.txt&exitpath=/Accounts/reports/PRNBankReceiptPrint.asp&frame=_parent"

%>
<%

	Function Header()
		'Blank Lines
		sText = sText & ""
		stext = stext & chr(10) 'Line Feed Character
		'stext = stext & " " & vbcrlf & " " & vbcrlf & " " & vbcrlf
		
		
		
		sQuery = "Select ReceivableNumber,IsNull(AmountReceived,0) from acc_T_createdRcvbleAdjDet where CreatedTransNo = "& iTransNo
		'Response.Write sQuery 
		objRs1.Open sQuery,con
		iAdjRecCount = 0
		do while not objRs1.EOF   
		
			iAdjRecCount  = iAdjRecCount + 1
		objRs1.MoveNext 
		loop
		
		
			'Response.Write iAdjRecCount
		objRs1.Close 
		iTotPagNo = Round(cdbl(iAdjRecCount)/7)
		 'Response.Write "iTotPagNo="&iTotPagNo
		' Response.End 
		If iTotPagNo > 1 then 
			'sTemp = sTemp & replace(iPgTot,"PGTOT","Page "& iPageNo &" of "& iTotPagNo)
			stext = stext & myAlign("Page "& iPageNo &" of "& iTotPagNo,95,"R") 
		Else
			stext = stext & " " & vbcrlf & " " & vbcrlf & " " & vbcrlf
		End IF
		
		
		'iPgTot="PGTOT"
		
		stext = stext & " " & vbcrlf & " " & vbcrlf

		'Number and date

		'sTemp = sTemp & myAlign("LINE NO ="&iLineNo&"****",30,"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(0,0)+9,"L")
		sTemp = sTemp & myAlign(nReceiptNo,aiHeaderColWidth(0,1),"L")

		sText = sText & sTemp
		sText = sText &" " & vbCrLf  &" " & vbcrlf
		sTemp = ""
		sTemp = sTemp & myAlign("",aiHeaderColWidth(0,0)+9,"L")
		sTemp = sTemp & myAlign(dReceiptDate,aiHeaderColWidth(0,1),"L")

		sText = sText & sTemp
		sText = sText &" " & vbCrLf &" " & vbcrlf &" " & vbcrlf
		sTemp = ""

		'Receivedwiththanks,thesumofrupees
		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+aiHeaderColWidth(1,0)+16,"L")
		if trim(sReceivedFrom) = "" then 
			sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(1,1)-aiHeaderColWidth(1,0),"L")
		else
			sTemp = sTemp & myAlign(sReceivedFrom,aiHeaderColWidth(1,1)-aiHeaderColWidth(1,0),"L")
		end if 	
		sText = sText & sTemp
		stext = stext & " " & vbcrlf
		sTemp = ""

		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+aiHeaderColWidth(1,0)+16,"L")
		sTemp = sTemp & myAlign(sAddressLine1,aiHeaderColWidth(1,1)-aiHeaderColWidth(1,0),"L")
		sText = sText & sTemp
		stext = stext & " " & vbcrlf
		sTemp = ""

		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+aiHeaderColWidth(1,0)+16,"L")
		sTemp = sTemp & myAlign(sAddressLine2,aiHeaderColWidth(1,1)-aiHeaderColWidth(1,0),"L")
		sText = sText & sTemp
		stext = stext & " " & vbcrlf
		sTemp = ""

		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+aiHeaderColWidth(1,0)+16,"L")
		sTemp = sTemp & myAlign(sCity,aiHeaderColWidth(1,1)-aiHeaderColWidth(1,0),"L")
		sText = sText & sTemp
		stext = stext & " " & vbcrlf & " " & vbCrLf &" " & vbcrlf
		sTemp = ""

		sAmountInWords = Replace(sAmountInWords,"Rupees"," ")
		sAmountInWords = Trim(sAmountInWords)

		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+7,"L")
		sTemp = sTemp & myAlign(sAmountInWords,aiHeaderColWidth(1,1),"L")

		sText = sText & sTemp
		sText = sText &" " & vbCrLf &" " & vbcrlf  &" " & vbcrlf   &" " & vbcrlf
		sTemp = ""
		iLineNo = iLineNo + 19
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+10,"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,1),"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,3)+aiHeaderColWidth(2,6),"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,4)+4,"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,7),"L")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(2,8),"L")
		sTemp = sTemp & myAlign(sPaymentBy,aiHeaderColWidth(2,9),"L")

		sText = sText & sTemp
		sText = sText &" " & vbCrLf
		iLineNo = iLineNo +1
		sTemp = ""
	End Function
'================================= USER DEFINED FUNCTIONS ================================='
'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function centerAlign(str1,width)
		dim diff,strlen,val, i, str, newstr, blank
			str = str1
			strlen = len(str)
			diff = width - strlen
			for i=0 to (diff-1)/2
				blank = blank & " "
			next
			newstr = blank & str & blank
		centerAlign = newstr
	end function
	'------------------------End OF myAlign Function----------------------------
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<TITLE>No Records </TITLE>
</HEAD>
<BODY BGCOLOR="#CCCCCC" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<TABLE BORDER=0 ALIGN=CENTER CELLSPACING=0 CELLPADDING=0 NOF=LY>
    <TR >
		<TD height="20"><%=sText%>&nbsp;</TD>
		<TD WIDTH=549 ><P ALIGN=CENTER><B>
		<FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT">No Records Found</FONT></B></TD>
	</TR>
	<TR>
        <TD height="20">&nbsp;</TD>
        <TD WIDTH=549 ><P ALIGN=CENTER><B>
        <FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT"><a href="#" onclick="window.history.back(1); return false;">Back</a></FONT></B></TD>
    </TR>
</Table>
</HTML>

