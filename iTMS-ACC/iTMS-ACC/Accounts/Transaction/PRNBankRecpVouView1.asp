<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name					:	PRNBankRecpVouView1.asp
	'Module Name					:
	'Author Name					:	S.MAHESWARI
	'Modified By					:
	'Created On						:	29 July 2008
	'Modified On					:   23 Sep 2008
	'Tables Used					:
	'Temporary Tables				:
	'Temporary Files				:
	'Input Parameter				:	None
	'Connects To					:
	'Procedures/Functions Used		:
	'Internal Variables				:
	'Database						: iTMS_At_KSS_Test
	'Queries Used					:
	'Counters						:
	'String							:
	'Boolean						:
	'Object Holders					:
	'Description					:
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/PrintFunctions.asp"-->
<!--#include File="../../include/GetOrganization.asp"-->

			<%
'------------------------Declaration Constants -----------------------------
dim aiHeaderColWidth(5,9),objFSO,objTxt
dim sTextOut,sTempStr,sPaidTo,iAmount
dim sDetails, sReceivedPayment,sPreparedBy,sCheckedBy,sPassedBy,dcrs,iPageNo,sRetVal,i
dim sTemp,sHeadOfAcc,sAccFlag,sOnlyPayVou

set		dcrs		= server.CreateObject("adodb.recordset")
set		objFSO	= Server.CreateObject("Scripting.FileSystemObject")
set		objTxt	= objFSO.CreateTextFile(server.MapPath("../temp/Transaction/"&Session.SessionID&"_BankVoucher_View.txt"))
		sTempStr	= ""

			'XML DOM Variables
			Dim oDOM,Root,objRs,sQuery,objRsTemp
			dim sNarration,sAccount,sAddtional,iSno,sNarr
			dim dTotal,sOrgId,sAccHead,sType
			dim EntryNode,HeaderNode,dAmount,sAccNo,sPartyName
			dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
			dim iTransNo,iBkHeadCode,bOtherUnit,iTdsAmount,sExp,TempNode,sBankInsDet
			Dim sInstrType,sBankInsNo, sBankInsName,sBankInsDrawnOn,sBankInsDt
			Dim sAddress1, sAddress2,sCity,sState,sPostcode,sTranIndication,sTranEntryIndication
			Dim iCreatedBy,sCreatedOn,sVouStatus,sEmpName,iPartyCtrlAcc,sAdjType,sDetFlag
			Dim iEntryNo,iHeadOfAcc,sHeadOfAccName,iHeadOfAccAmt,iPartyCode,iEntryAmt,iCtr,iAccNo
			Dim iNetAmtPaid,iTotRecovered,sNarrFlag,sAddnFlag,iTotAddnlAmt,sAdjFlag,sAdjOn,iTotAdj,sBankInsAmt
			' Create our DOM Document Objects
			Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
			Dim iOne,iThree,iNoOfLinesCtr,iFlag12,iFlag34,sDrwOn,sPayAt

			iTransNo=Request("Value")
			'Response.Write iTransNo
			iFlag12 = false
			iFlag34 = false
			iNoOfLinesCtr = 0

			'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
			sRetVal = GetVouchXML(iTransNo)
			oDOM.Load server.MapPath(sRetVal)

			set Root=oDOM.documentElement
			sOrgId=Root.Attributes.Item(0).nodeValue
			sOrgName=Root.Attributes.Item(1).nodeValue
			iBookCode =Root.Attributes.Item(2).nodeValue
			sBookName =Root.Attributes.Item(3).nodeValue
			sVouType=Root.Attributes.Item(4).nodeValue
			sVoucDate=Root.Attributes.Item(5).nodeValue
			iBkHeadCode=Root.Attributes.Item(6).nodeValue

			iVouNo=Root.Attributes.Item(9).nodeValue
			sApprove=Root.Attributes.Item(7).nodeValue

			set objRs = Server.CreateObject("ADODB.Recordset")
			set objRsTemp = Server.CreateObject("ADODB.Recordset")

			sQuery="select OtherUnitTransaction from vwOrgBookNames where OUDefinitionID = '" & sOrgId & "'"&_
			"and BookNumber="&iBookCode&" and BookCode='02'"

			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with

			if not 	objRs.EOF then
				bOtherUnit=objRs(0)
			else
				bOtherUnit=0
			end if
			objRs.Close
			'Newly added on 10th Feb 2009 by S.Maheswari to fetch account no. from Acc_M_BankDetails table
			sQuery = "Select isnull(AccountNo,'') from Acc_M_BankDetails where  BookNumber="&iBookCode&" and BookCode='02'"
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with

			if not 	objRs.EOF then
				iAccNo=objRs(0)
			end if
			objRs.Close
			'==============================Voucher View header queries starts here ============================
			sQuery = "select Distinct CreatedVoucherNo,TransactionType,OUDefinitionID,PayToRecdFrom," &_
					"PartyCode,CreatedBy,Convert(varchar,CreatedOn,103),CreatedVouchStatus," &_
					"BankInstrumentType,BankInstrumentNo,Convert(varchar,BankInstrumentDate,103),PayableAt,DrawnOnBank" &_
					" from VW_Created_BankVoucherView where CreatedTransNo="& iTransNo
				'	Response.Write sQuery
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
				'sBankInsDt = objRs(10)
				'sBankInsDrawnOn = objRs(12)

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

			'Newly added on July 2nd 2008 by S.Maheswari to fetch bank details from Acc_T_CreatedVoucherInstrumentDet table instead of taking from CreatedVoucherHeader table
			sQuery = "Select BankInstrumentNo,convert(Varchar,BankInstrumentDate,103),BankInstrumentType,InstrumentAmount,PayableAt,DrawnOnBank "&_
					" from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo = "&iTransNo&" Order by InstrumentEntryNo "
			 'Response.Write sQuery
			 'Response.ENd

			objRs.Open sQuery,con
			If Not objRs.EOF then
				Do while not objRs.EOF
					sBankInsNo		= sBankInsNo&","&objRs(0)
					sBankInsDt		= sBankInsDt&","&objRs(1)
					sBankInsName	= sBankInsName&","&objRs(2)
					sBankInsAmt	    = sBankInsAmt&","&objRs(3)
					sBankInsDrawnOn = sBankInsDrawnOn&","&objRs(5)
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

			If trim(sVouStatus) = "010104" Then
				sQuery = "Select VoucherNumber from ACC_T_VoucherHeader where CreatedTransNo="&iTransNo
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
					sTranIndication = "D"
					sTranEntryIndication = "C"
				Else
					sTranIndication = "C"
					sTranEntryIndication = "D"
				End If
			sQuery = "Select VoucherEntryNumber,AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranIndication&"' "
			 
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with

			set objRs.ActiveConnection = nothing
			if not objRs.EOF then
				iEntryNo = objRs(0)
				iHeadOfAcc = objRs(1)
				iHeadOfAccAmt = FormatNumber(objRs(2),2,,,0)
				sNarr = objRs(3)
				iPartyCtrlAcc = objRs(4)
			end if
			objRs.Close
			
			sQuery = "Select Count(1) From Acc_T_CreatedVoucherHeader H,Acc_T_CreatedVoucherDetails D "&_
					 "Where H.CreatedTransNo = D.CreatedTransNo And H.TransactionType = 'BAP' " & _
					 "And D.TransCrDrIndication = 'C' And H.CreatedTransNo = "&iTransNo
			objRs.Open sQuery,con
			IF Not objRs.EOF Then
				sTemp = objRs(0)
			Else
				sTemp = 0
			End IF
			objRs.Close
			
			IF CStr(sTemp) = "0" Then
				sQuery = "Select VoucherAmount From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo
				objRs.Open sQuery,con
				IF Not objRs.EOF Then
					iHeadOfAccAmt = objRs(0)
					sOnlyPayVou = "True"
				End IF
				objRs.Close
			Else
				sOnlyPayVou = "False"
			End IF
			
			
			
			
			
			
  			'Response.Write iPartyCtrlAcc &"<BR>"
			IF iHeadOfAcc <> "" Then
				sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
			Else
				If trim(iPartyCtrlAcc) <> "" then 	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
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
					sHeadOfAcc = objRs(0)
				end if
			objRs.Close
			'Response.Write "sHeadOfAccName="&sHeadOfAcc & vbCrLf
			sQuery = "select OrgUnitDescription,Address1,Address2,City,State,PostCode from DCS_OrganizationUnitDefinitions where OUDefinitionID='"&sOrgId&"'"
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection = nothing
				if not 	objRs.EOF then
					sOrgName = objRs(0)
					sAddress1 = objRs(1)
					sAddress2 = objRs(2)
					sCity = objRs(3)
					sState = objRs(4)
					sPostcode = objRs(5)
				else
					sOrgName = ""
					sAddress1 = ""
					sAddress2 = ""
					sCity = ""
					sState = ""
					sPostcode = ""
				end if
			objRs.Close
			If iPartyCode <> "0" and iPartyCode <> ""  then
				sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCode
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
					End IF
				objRs.Close
			End If

					
					sQuery = "SELECT LoginId FROM DCS_User WHERE InternalUserID ="&iCreatedBy
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
	                End IF
                    objRs.Close

			'==============================Voucher View header queries ends here ============================

				sQuery = "Select isNull(BankInstrumentType,''),ISNull(BankInstrumentNo,''),isNull(BankInstrumentDate,''), "&_
						 "isNull(PayableAt,'') From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo
				with objRs
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with

				if not 	objRs.EOF then
					'sPayTo = objRs(3)
					sAccNo = objRs(1)
					sType = objRs(0)
				End IF
				objRs.Close
				IF sType = "P" Then
					sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&sAccNo(3)
					With Objrs
						.ActiveConnection = con
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQuery
						.Open
					End With
					Set Objrs.ActiveConnection = nothing

					IF not objRs.EOF  Then
						sPartyName = Trim(Objrs(0))
					End IF
					objRs.Close
				End IF

'No and Date
aiHeaderColWidth(0,0)=67
aiHeaderColWidth(0,1)=16

'Head of Account, Amount
aiHeaderColWidth(1,0)=14
aiHeaderColWidth(1,1)=70

'Paid To, Rupees
aiHeaderColWidth(2,0)=14
aiHeaderColWidth(2,1)=46
aiHeaderColWidth(2,2)=7
aiHeaderColWidth(2,3)=16

'NameOfBank, Details
aiHeaderColWidth(3,0)=3
aiHeaderColWidth(3,1)=20
aiHeaderColWidth(3,2)=3
aiHeaderColWidth(3,3)=21
aiHeaderColWidth(3,4)=80
aiHeaderColWidth(3,5)=10
aiHeaderColWidth(3,6)=1
aiHeaderColWidth(3,7)=15
aiHeaderColWidth(3,8)=53
aiHeaderColWidth(3,9)=10


'Prepared by, Checked by, Passed by
aiHeaderColWidth(4,0)=3
aiHeaderColWidth(4,1)=16
aiHeaderColWidth(4,2)=3
aiHeaderColWidth(4,3)=18
aiHeaderColWidth(4,4)=2
aiHeaderColWidth(4,5)=18

'------------------------End of Declaration Constants ----------------------
iPageNo = 1
Header(iPageNo)
Function Header(iPageNo)
'iPageNo=1
	if sType="P" then
		sPaidTo			= sPartyName
	else
		sPaidTo			= sPayTo
	end if
	iNoOfLinesCtr = 0
	'Blank Lines
	sTextOut = sTextOut & " " & vbcrlf & Vbcrlf
	iNoOfLinesCtr = iNoOfLinesCtr + 2
	'Number and date

	 sTextOut =  sTextOut & myAlign("",aiHeaderColWidth(0,0)+6,"L")

	'Response.Write  sHeadOfAcc
	IF Cstr(sVouStatus) = "010104" Then
		 sTextOut =  sTextOut & myAlign(iVouNo,aiHeaderColWidth(0,1),"L")
	Else
		 sTextOut =  sTextOut & myAlign(" ",aiHeaderColWidth(0,1),"L")
	End IF
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 1

	sTextOut = sTextOut & " " & vbcrlf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(0,0)+6,"L")
	sTextOut = sTextOut  & myAlign(sVoucDate,aiHeaderColWidth(0,1),"L")
'
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 2

	'Blank Lines
	sTextOut = sTextOut & " " & vbcrlf & " " & vbcrlf
	'Head of Account
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+8,"L")
	'sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(1,1),"L")
	sTextOut = sTextOut  & myAlign(sHeadOfAcc,aiHeaderColWidth(1,1),"L")
'
	sTextOut = sTextOut & vbCrLf & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 3
	IF iPageNo > 1  then
		sTextOut = sTextOut & vbCrLf
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+7,"L")
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,1),"L")
		'
		sTextOut = sTextOut & vbCrLf

		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+7,"L")
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,1),"L")
	'
		'sTextOut = sTextOut & vbCrLf

		iNoOfLinesCtr = iNoOfLinesCtr + 3
		'PaidTo,Rs
		sTextOut = sTextOut & vbCrLf & Vbcrlf
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(2,0)+7,"L")
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(2,1),"L")
		sTextOut = sTextOut  & myAlign("",5,"R")
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(2,2)+2,"R")
		'
		sTextOut = sTextOut & vbCrLf
		sTextOut = sTextOut & vbCrLf
		iNoOfLinesCtr = iNoOfLinesCtr + 4

		sBInsDtFlag = True
		sBInsNoFlag = True
		sAccFlag = True
		sTextOut = sTextOut & vbCrLf & vbCrLf & vbCrLf
		iNoOfLinesCtr = iNoOfLinesCtr + 2
	End If
End Function

	Dim sDesc2,sDesc,iWidth
	sDesc	= AmountWords(replace(iHeadOfAccAmt,",",""))
	If Len(sDesc) > 50 Then
		For i = 1 to 50
			If Mid(sDesc,50-i,1) = " " Then
				iWidth = 50-i
			Exit For
			End if
		Next
		sDesc2 = Mid(sDesc,iWidth+1,Len(sDesc))
		sDesc =  Mid(sDesc,1,50-i)
	End If

	sTextOut = sTextOut & vbCrLf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+8,"L")
	sTextOut = sTextOut  & myAlign(sDesc,aiHeaderColWidth(1,1),"L")
	'
	sTextOut = sTextOut & vbCrLf

	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+8,"L")
	sTextOut = sTextOut  & myAlign(sDesc2,aiHeaderColWidth(1,1),"L")
'
	'sTextOut = sTextOut & vbCrLf

	iNoOfLinesCtr = iNoOfLinesCtr + 3
	'PaidTo,Rs
	sTextOut = sTextOut & vbCrLf & Vbcrlf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(2,0)+8,"L")
	sTextOut = sTextOut  & myAlign(sPayTo,aiHeaderColWidth(2,1),"L")
	sTextOut = sTextOut  & myAlign("",5,"R")
	sTextOut = sTextOut  & myAlign(FormatNumber(iHeadOfAccAmt,2,,,0),13,"R")
	'
	sTextOut = sTextOut & vbCrLf
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 4


'	'Details, ReceivedPayment
	'Name Of Bank :
	'blocked on jul 30th
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sTextOut = sTextOut  & formattprint("CONDENSESTART","") 'Condensed formatt
'sTextOut = sTextOut  & " " & vbCrLf
'iNoOfLinesCtr = iNoOfLinesCtr + 1

%>

<%
'Fetch the ADDITIONAL PAYMENT / RECEIPT ENTRIES
'check from this


dim sFlagTotAmt,sum,sBInsNoFlag


sQuery = "Select VoucherNarration from Acc_T_CreatedVoucherDetails where createdTransNo = "&iTransNo&" "
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
iOne = objRs.RecordCount
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sNarr = objRs(0)
end if
objRs.Close
IF sBankInsDrawnOn <> "" then
	IF trim(sBankInsDrawnOn) = "AB" or trim(sBankInsDrawnOn) = "ANDHRA BANK"  then
		sDrwOn = "ANDHARA BANK"
		sPayAt = "TIRUPUR"
	'Else
	'	sTemp = split(sBankInsDrawnOn,",")
	'	sDrwOn = trim(sTemp(0))
	'	sPayAt = trim(sTemp(1))
	End IF

	sTextOut = sTextOut  & vbCrLf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5)+4,"L") '31 Space
	sTextOut = sTextOut  & myAlign(trim(sDrwOn),aiHeaderColWidth(3,3)-4,"L") & vbCrLf '21
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5)+4,"L") '31 Space
	sTextOut = sTextOut  & myAlign(trim(sPayAt),aiHeaderColWidth(3,3)-4,"L")
Else
	sTextOut = sTextOut  & vbCrLf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3),"L") & vbCrLf '21
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3),"L")
End If

	sTextOut = sTextOut  & myAlign("",4,"L")
	iNoOfLinesCtr =  iNoOfLinesCtr + 3
	sum = len(sNarr) / 80
	sum = Round(sum)

	IF sNarr <> "" then
		IF len(sNarr) > 80 then
			sNarr = BreakString(sNarr,80)
			for i = 0 to UBOUND(sNarr)
				sTextOut = sTextOut  & myAlign(" ",2,"L")
				sTextOut = sTextOut  & myAlign(UCASE(trim(sNarr(i))),80,"L")& vbCrLf  '80
				iNoOfLinesCtr = iNoOfLinesCtr + 1

				IF sum = i then
					sNarr = ""
					Exit for
				End If
				IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
					BankDetails(iNoOfLinesCtr)
				else
					sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space
					sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+4,"L")  '21
				End IF
			next
		else
			sTextOut = sTextOut  & myAlign(" ",2,"L")
			sTextOut = sTextOut  & myAlign(UCASE(trim(sNarr)),aiHeaderColWidth(3,4),"L") & vbCrLf

			IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
			   	BankDetails(iNoOfLinesCtr)
			else
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+4,"L")  '21
			End IF
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		End IF
			sNarrFlag = True
	End IF

'Fetch the Receipts / Payments entries
'===================================RECOVERIES===============================================
dim sRecFlag,sBInsDtFlag,sRcvdFlag,sNetFlag,sTotAmtFlag,sInsideRec,sAlTransType
sInsideRec = "N"
IF Cstr(sTranEntryIndication) = "C" Then
	sAlTransType = "D"
Else
	sAlTransType = "C"
End IF

sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> "&iEntryNo&" and TransCrDrIndication='"&sTranEntryIndication&"'"
iCtr = 1
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
iThree = objRs.RecordCount

set objRs.ActiveConnection = nothing

if not 	objRs.EOF then
	sInsideRec = "Y"
	iCtr = 1

	IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
		BankDetails(iNoOfLinesCtr)
	End If

	IF sRecFlag <> True then
		sTextOut = sTextOut  & myAlign(" ",2,"L") '21 Space
		sTextOut = sTextOut  & myAlign("Recoveries",aiHeaderColWidth(3,1)+1,"L") '21 Space
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+aiHeaderColWidth(3,3)+6,"L")

		If sTotAmtFlag <> True then
			sTextOut = sTextOut  & myAlign(" ",2,"L")'16 Space
			sTextOut = sTextOut  & myAlign("Total Amount Rs.",aiHeaderColWidth(3,1)-4,"L")'16 Space
			sTextOut = sTextOut  & myAlign(iHeadOfAccAmt,aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			sTotAmtFlag = True
		Else
			sTextOut = sTextOut  & vbCrLf
		End If
		iNoOfLinesCtr = iNoOfLinesCtr + 1
		sRecFlag = True
	End IF
	iCtr = 0

	Do While not objRs.EOF
		iHeadOfAcc = objRs(0)
		iEntryAmt = cdbl(objRs(1))
		iTotRecovered = cdbl(iTotRecovered) + cdbl(objRs(1))
		sAddnFlag = True
		sDetFlag = True
		iCtr = iCtr + 1
		If iHeadOfAcc <> "" Then
			sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
		Else
			sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
		End If

		with objRsTemp
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		Set objRsTemp.ActiveConnection = Nothing
		sHeadOfAccName = ""
		if not 	objRsTemp.EOF then
			sHeadOfAccName = objRsTemp(0)
		end if
		objRsTemp.Close


		'If iNoOfLinesCtr = 21 then sTextOut = sTextOut  & myAlign("",54,"L")
		If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
			BankDetails(iNoOfLinesCtr)
		End If
		'sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5)+10,"L") '40 Space
		If iPageNo > 1  then
			If iCtr = 1 then sTextOut = sTextOut  & myAlign("",56,"L")
			sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
			sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
			sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R")& vbCrLf '13 Space
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		Else
			sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
			sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
			sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R") '13 Space
		End IF

		If iCtr = 1 and sRcvdFlag <> True then
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+2,"L")
			sTextOut = sTextOut  & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,1)-4,"L")
			sTextOut = sTextOut  & myAlign("XXXXXXXXXXXXX",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			sRcvdFlag = True
		End If

		If iCtr = 2 and sNetFlag <> True then
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+2,"L")
			IF CStr(sOnlyPayVou) = "False" Then
				sTextOut = sTextOut  & myAlign("Net Amount Rs.",aiHeaderColWidth(3,1)-4,"L")
				sTextOut = sTextOut  & myAlign("YYYYYYYYYYYYY",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			Else
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			End IF
			sNetFlag = True
		End If

		'Response.write iNoOfLinesCtr
		If iCtr > 2 then sTextOut = sTextOut  &  vbCrLf
		iNoOfLinesCtr = iNoOfLinesCtr + 1
		If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
			BankDetails(iNoOfLinesCtr)
		Else
			sTextOut = sTextOut  & myAlign("",56,"L")
		End If

		IF iNoOfLinesCtr > 28 then
			iPageNo = iPageNo + 1
			iCtr = 0
			Header(iPageNo)
		End IF
	 objRs.MoveNext
	Loop
end if
objRs.Close

'Response.end

IF Cstr(sInsideRec) = "N" Then

	sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> "&iEntryNo&" and TransCrDrIndication='"&sAlTransType&"'"
	iCtr = 1
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	iThree = objRs.RecordCount

	set objRs.ActiveConnection = nothing
	if not 	objRs.EOF then
		iCtr = 1

		If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
			Response.write "16"
				'Response.end
			BankDetails(iNoOfLinesCtr)
		'Else
			'sTextOut = sTextOut  & myAlign("",56,"L")
		End IF
		
		IF sRecFlag <> True then
			'IF iNoOfLinesCtr <> 21 then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5)+4,"L")
			'IF sBInsNoFlag <> True then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+3,"L") '19 Space

			sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
			sTextOut = sTextOut  & myAlign("Additional Payments",aiHeaderColWidth(3,1)+1,"L") '21 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+aiHeaderColWidth(3,3)+6,"L")

			'This is been Blocked by Manohar on 14/10/08 From the CAll from KSS that for the Payments voucher only payment entry is there then the following is not needed.
			If sTotAmtFlag <> True then
				sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
				'sTextOut = sTextOut  & myAlign("Total Amount Rs.",aiHeaderColWidth(3,1)-4,"L")'16 Space
				'sTextOut = sTextOut  & myAlign(iHeadOfAccAmt,aiHeaderColWidth(3,5)+3,"R") & vbCrLf
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")'16 Space
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
				sTotAmtFlag = True
			Else
				sTextOut = sTextOut  & vbCrLf
			End If
			iNoOfLinesCtr = iNoOfLinesCtr + 1
			sRecFlag = True
		End IF
		iCtr = 0

		Do While not objRs.EOF
			iHeadOfAcc = objRs(0)
			iEntryAmt = cdbl(objRs(1))
			iPartyCtrlAcc = objRs(3)
			iTotRecovered = 0
			'iTotRecovered = cdbl(iTotRecovered) + cdbl(objRs(1))
			sAddnFlag = True
			sDetFlag = True
			iCtr = iCtr + 1
			If iHeadOfAcc <> "" Then
				sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
			Else
				sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
			End If
			with objRsTemp
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			Set objRsTemp.ActiveConnection = Nothing
			sHeadOfAccName = ""
			if not 	objRsTemp.EOF then
				sHeadOfAccName = objRsTemp(0)
			end if
			objRsTemp.Close

			'Response.write iNoOfLinesCtr
			'If iCtr > 2 then sTextOut = sTextOut  &  vbCrLf
			'iNoOfLinesCtr = iNoOfLinesCtr + 1
			If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
				BankDetails(iNoOfLinesCtr)
		'	else
		'		sTextOut = sTextOut  & myAlign("",56,"L")
			End IF
			 
			'sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+7,"L")
			If iPageNo > 1  then
				If iCtr = 1 then sTextOut = sTextOut  & myAlign("",56,"L")
				sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
				sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
				sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R")& vbCrLf '13 Space
				iNoOfLinesCtr = iNoOfLinesCtr + 1
			Else
				sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
				sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
				sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R") '13 Space
			End IF

			'This is been Blocked by Manohar on 14/10/08 From the CAll from KSS that for the Payments voucher only payment entry is there then the following is not needed.
			If iCtr = 1 and sRcvdFlag <> True then
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+2,"L")
				'sTextOut = sTextOut  & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,1)-4,"L")
				'sTextOut = sTextOut  & myAlign("XXXXXXXXXXXXX",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
				sRcvdFlag = True
			End If
			
			'This is been Blocked by Manohar on 14/10/08 From the CAll from KSS that for the Payments voucher only payment entry is there then the following is not needed.
			If iCtr = 2 and sNetFlag <> True then
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+2,"L")
				IF CStr(sOnlyPayVou) = "False" Then
					sTextOut = sTextOut  & myAlign("Net Amount Rs.",aiHeaderColWidth(3,1)-4,"L")
					sTextOut = sTextOut  & myAlign("YYYYYYYYYYYYY",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
				Else
					sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")
					sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
				End IF
				'sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")
				'sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
				sNetFlag = True
			End If

			If iCtr > 2 then sTextOut = sTextOut  &  vbCrLf
			iNoOfLinesCtr = iNoOfLinesCtr + 1
			If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
				Response.write "2"
				'Response.end
				BankDetails(iNoOfLinesCtr)
			Else
				sTextOut = sTextOut  & myAlign("",56,"L")
			End If

			IF iNoOfLinesCtr > 28 then
				iPageNo = iPageNo + 1
				iCtr = 0
				Header(iPageNo)
			End IF
		 objRs.MoveNext
		Loop
	end if
	objRs.Close

End IF
'Response.end
'**********************************************************************************
'Fetch the Receivable Adjustments
sQuery = "Select ReceivableNumber,Convert(varchar,ReceivedOn,103),AmountReceived,AdjustType from Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
'iCtr = 1

'sTextOut = sTextOut &"2nd Test entry "
'Response.Write sTextOut
'Response.End

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then

	IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
		Response.write "3"
				'Response.end
		BankDetails(iNoOfLinesCtr)
	End If
	'IF sRecFlag <> True then
		'IF iNoOfLinesCtr <> 21 then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5)+4,"L")
		'IF sBInsNoFlag <> True then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+3,"L") '19 Space
		sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
		sTextOut = sTextOut  & myAlign("Adjustments",aiHeaderColWidth(3,1)+1,"L") '20 Space
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+aiHeaderColWidth(3,3)+6,"L")
		'

		If sTotAmtFlag <> True then
			sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
			sTextOut = sTextOut  & myAlign("Total Amount Rs.",aiHeaderColWidth(3,1)-4,"L")'16 Space
			sTextOut = sTextOut  & myAlign(iHeadOfAccAmt,aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			sTotAmtFlag = True
		Else
			sTextOut = sTextOut  & vbCrLf
		End If
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	'	sRecFlag = True
	'End IF
	iCtr = 0
	Do While not objRs.EOF
		iHeadOfAcc = objRs(0)
		sAdjOn = objRs(1)
		iEntryAmt = cdbl(objRs(2))
		sAdjType = objRs(3)
		iTotAdj = iTotAdj + iEntryAmt
		sDetFlag = True
		iCtr = iCtr + 1
		sQuery = "Select Narration from ACC_T_CreatedReceivables where Receivablenumber = "&iHeadOfAcc
		with objRsTemp
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		Set objRsTemp.ActiveConnection = Nothing
		sHeadOfAccName = ""
		if not 	objRsTemp.EOF then
			sHeadOfAccName = objRsTemp(0)
		end if
		objRsTemp.Close

		If trim(sAdjType) = "A" Then sHeadOfAccName = "Less Advance Receipts "

		If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
			Response.write "4"
				'Response.end
			BankDetails(iNoOfLinesCtr)

		End If

		'If iNoOfLinesCtr <> 21 then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,8)-4,"L")
		'sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+5,"L")
		If iPageNo > 1  then
			If iCtr = 1 then sTextOut = sTextOut  & myAlign("",56,"L")
			sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
			sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
			sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R")& vbCrLf '13 Space
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		Else
			sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
			sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
			sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R") '13 Space
		End IF

		If iCtr = 1 and sRcvdFlag <> true then
			sTextOut = sTextOut  & myAlign(" ",2,"L") '2 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0),"L")
			sTextOut = sTextOut  & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,1)-4,"L")
			sTextOut = sTextOut  & myAlign("XXXXXXXXXXXXX",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			sRcvdFlag = True
		End If

		If iCtr = 2 and sNetFlag <> True then
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0),"L")
			IF CStr(sOnlyPayVou) = "False" Then
				sTextOut = sTextOut  & myAlign("Net Amount Rs.",aiHeaderColWidth(3,1)-4,"L")
				sTextOut = sTextOut  & myAlign("YYYYYYYYYYYYY",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			Else
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			End IF
			sNetFlag = True
		End If

		If iCtr > 2 then sTextOut = sTextOut  &  vbCrLf
		iCtr = iCtr + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
		If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
			Response.write "5"
				'Response.end
			BankDetails(iNoOfLinesCtr)
		Else
			sTextOut = sTextOut  & myAlign("",56,"L")
		End If

		IF iNoOfLinesCtr > 28 then
			iPageNo = iPageNo + 1
			iCtr = 0
			Header(iPageNo)
		End IF
	objRs.MoveNext
	Loop
end if
objRs.Close

'Response.Write Len(sTempStr)
'Response.Write "00"
'Response.End

'
'sTextOut = sTextOut &"3nd Test entry "

'********************************************************************************************************************************************************************
'Fetch the Payable Adjustments
sQuery = "select PayablesNumber,Convert(varchar,PaidOn,103),AmountPaid,AdjustType from Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo

iCtr = 1
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then


	IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
		Response.write "6"
				'Response.end
		BankDetails(iNoOfLinesCtr)
	'Else
	'	sTextOut = sTextOut & myAlign("",56,"L")

	End If
	'IF sRecFlag <> True then
	'IF iNoOfLinesCtr <> 21 then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5)+4,"L")
'	IF sBInsNoFlag <> True then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+3,"L") '19 Space

	sTextOut = sTextOut & myAlign("Adjustments",aiHeaderColWidth(3,1)+1,"L") '20 Space
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,1)+aiHeaderColWidth(3,3)+6,"L")

	If sTotAmtFlag <> True then
		sTextOut = sTextOut  & myAlign("Total Amount Rs.",aiHeaderColWidth(3,1)-4,"L")'16 Space
		sTextOut = sTextOut  & myAlign(iHeadOfAccAmt,aiHeaderColWidth(3,5)+3,"R") & vbCrLf
		sTotAmtFlag = True
	Else
		sTextOut = sTextOut  & vbCrLf
	End If
	iNoOfLinesCtr = iNoOfLinesCtr + 1
	'Response.Write sTempStr & vbCrLf
	'Response.Write iNoOfLinesCtr
	'Response.End
	IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
		Response.write "7"
			'	Response.end
		BankDetails(iNoOfLinesCtr)
	End If
	iCtr = 0
	Do While not objRs.EOF
		iHeadOfAcc = objRs(0)
		sAdjOn = objRs(1)
		iEntryAmt = cdbl(objRs(2))
		sAdjType = objRs(3)
		iTotAdj = iTotAdj + iEntryAmt
		sDetFlag = True
		iCtr = iCtr + 1
		sQuery = "select Narration from ACC_T_CreatedPayables where Payablesnumber ="&iHeadOfAcc
		with objRsTemp
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		Set objRsTemp.ActiveConnection = Nothing
		sHeadOfAccName = ""
		if not 	objRsTemp.EOF then
			sHeadOfAccName = objRsTemp(0)
		end if
		objRsTemp.Close

		If trim(sAdjType) = "A" Then sHeadOfAccName = "Less Advance Payments "

		'If iNoOfLinesCtr = 23 then sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,8)+3,"L")
		'sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+5,"L")
		'Response.Write iPageNo & vbCrLf
		'Response.Write "CTR="&iCtr & vbCrLf
		If iPageNo > 1  then
			If iCtr = 1 then sTextOut = sTextOut  & myAlign("",56,"L")
			sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
			sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R")& vbCrLf '13 Space
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		Else
			sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(3,1)+aiHeaderColWidth(3,1)+aiHeaderColWidth(3,5),"L") '50 Space
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)-1,"L")
			sTextOut = sTextOut  & myAlign(Formatnumber(iEntryAmt,2,,,0),aiHeaderColWidth(3,7)-2,"R") '13 Space
		End IF
		If iCtr = 1 and sRcvdFlag <> true then
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0),"L")
			sTextOut = sTextOut  & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,1)-4,"L")
			sTextOut = sTextOut  & myAlign("XXXXXXXXXXXXX",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			sRcvdFlag = True
		End If
		'Response.Write sTempStr & vbCrLf
		'Response.Write "CTR="&iCtr & sNetFlag


		If iCtr = 2 and sNetFlag <> True then
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0),"L")
			IF CStr(sOnlyPayVou) = "False" Then
				sTextOut = sTextOut  & myAlign("Net Amount Rs.",aiHeaderColWidth(3,1)-4,"L")
				sTextOut = sTextOut  & myAlign("YYYYYYYYYYYYY",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			Else
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")
				sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
			End IF
			sNetFlag = True
		End If

		If iCtr > 2 then sTextOut = sTextOut  &  vbCrLf

		'Response.write iNoOfLinesCtr

		iNoOfLinesCtr = iNoOfLinesCtr + 1
		If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22  or iNoOfLinesCtr = 24 then
			Response.write "8"
				'Response.end
			BankDetails(iNoOfLinesCtr)
		Else
			sTextOut = sTextOut  & myAlign("",56,"L")
		End If

		IF iNoOfLinesCtr > 28 then
			iPageNo = iPageNo + 1
			iCtr = 0
			Header(iPageNo)
		End IF

	objRs.MoveNext
	Loop
end if
objRs.Close
'Response.end
'**********************************************************************************************************
IF sDetFlag = True then
	If sRcvdFlag <> true then
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,3)+aiHeaderColWidth(3,1)+6,"L")
		sTextOut = sTextOut  & myAlign("(-)Recovered",aiHeaderColWidth(3,1)-4,"L")
		sTextOut = sTextOut  & myAlign(FormatNumber(iTotRecovered,2,,,0),aiHeaderColWidth(3,5)+3,"R") & vbCrLf
		sRcvdFlag = True
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	End If
	
	
	IF sNetFlag <> true  then
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,3)+aiHeaderColWidth(3,1)+10,"L")
		IF CStr(sOnlyPayVou) = "False" Then
			sTextOut = sTextOut  & myAlign("Net Amount Rs.",aiHeaderColWidth(3,1)-4,"L")
			sTextOut = sTextOut  & myAlign(FormatNumber((iHeadOfAccAmt-iTotRecovered),2,,,0),aiHeaderColWidth(3,5)+3,"R") & vbCrLf
		Else
			sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,1)-4,"L")
			sTextOut = sTextOut  & myAlign(" ",aiHeaderColWidth(3,5)+3,"R") & vbCrLf
		End IF
		sNetFlag = True
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	End If
End IF
If trim(iTotRecovered) = "" then iTotRecovered = 0
Response.Write iHeadOfAccAmt & "<br>"
Response.Write iTotRecovered &"=="
sTemp = CDbl(iHeadOfAccAmt) - CDbl(iTotRecovered)

sTextOut  = Replace(sTextOut,"XXXXXXXXXXXXX",myAlign(FormatNumber(iTotRecovered,2,,,0),aiHeaderColWidth(3,5)+3,"R"))
sTextOut  = Replace(sTextOut,"YYYYYYYYYYYYY",myAlign(FormatNumber((sTemp),2,,,0),aiHeaderColWidth(3,5)+3,"R"))

'sTextOut = sTextOut  & vbCrLf
'iNoOfLinesCtr = iNoOfLinesCtr + 1

If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
	Response.write "9"
				'Response.end
	BankDetails(iNoOfLinesCtr)
	sTextOut = sTextOut  & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 1
End If

'Response.write "<br>" & iNoOfLinesCtr & "<br>"

For i = iNoOfLinesCtr to 28
	If iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then
		Response.write "10"
				'Response.end
		BankDetails(iNoOfLinesCtr)
	End If
	sTextOut = sTextOut  & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 1
	iPageNo = iPageNo + 1
Next
'Response.end
sTextOut = sTextOut & formattprint("CONDENSEEND","") 'Condensed formatt
'Response.End
sTextOut = sTextOut & Vbcrlf
sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,0)+7,"L")

IF Cstr(sVouStatus) <> "010104" Then
	sTextOut = sTextOut & myAlign(sEmpName&"/" & iVouNo & "-"&sCreatedOn,aiHeaderColWidth(3,7)+4,"L")
Else
	sTextOut = sTextOut & myAlign(sEmpName&"-"&sCreatedOn,aiHeaderColWidth(3,7),"L")
End IF
'sTextOut = sTextOut & chr(12)
sTextOut = sTextOut & Vbcrlf
sTextOut = sTextOut & Vbcrlf
sTextOut = sTextOut & Vbcrlf
sTextOut = sTextOut & Vbcrlf
sTextOut = sTextOut & Vbcrlf
sTextOut = sTextOut & Vbcrlf
sTextOut = sTextOut & Vbcrlf

'sTextOut = sTextOut & Vbcrlf

	objTxt.write sTextOut
	Response.Redirect("../../Components/FormattPrintNew.asp?server=server&filepath=/accounts/temp/Transaction/"&Session.SessionID&"_BankVoucher_View.txt&exitpath=/accounts/reports/VouBAView.asp&frame=_parent")
	'Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/Transaction/"&Session.SessionID&"_BankVoucher_View.txt&exitpath=/accounts/reports/VouBAView.asp&frame=_parent")
%>
<%
'================================= USER DEFINED FUNCTIONS ================================='
'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function myAlign(val1,alen,str1)
	dim vlen,k,str2,val
		val=val1
		IF len(val) then vlen = CInt(len(val))
		if (vlen > alen) then
			val = Mid(val,1,alen-1)
			vlen = CInt(len(val))
		end if
		k = (alen - vlen)
		if alen = vlen then
		   str2 = val
		     myAlign = str2
		else if (str1="L") then
			str2 = val & String(k," ")
			myAlign = str2
		    else if (str1 = "R") then
			         str2 = String(k," ") & val
			         myAlign = str2
		          end if
		    end if
		end if
	end function
	'------------------------End OF myAlign Function----------------------------

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
Function BankDetails(iLineNo)
dim sAccFlag
	IF iLineNo = 20 and sAccFlag <> True then
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5)+4,"L") '31 Space
		sTextOut = sTextOut  & myAlign(iAccNo,aiHeaderColWidth(3,3)-4,"L")
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+1,"L")'2 Space
		sAccFlag = True
	End If

	'Response.write iLineNo
	'Response.end

	IF iLineNo = 22  and sBInsNoFlag <> true then
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5)+4,"L") '31 Space
		if trim(sBankInsNo)="0" then
			sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)-4,"L") '14
		else
			sTextOut = sTextOut  & myAlign(sBankInsNo,aiHeaderColWidth(3,3)-4,"L") '14
		end if
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+1,"L")'2 Space
		sBInsNoFlag = True

	End If


	If iNoOfLinesCtr = 24 and sBInsDtFlag <> True then
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5)+4,"L") '31 Space
		sTextOut = sTextOut  & myAlign(sBankInsDt,aiHeaderColWidth(3,3)-4,"L") '22
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+1,"L")'2 Space
		sBInsDtFlag = True

	End If
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<TITLE>iTMS -
		  <%
			if sVouType="BAP" then
				Response.Write "Cheque Payment Voucher"
			else
				Response.Write "Cheque Receipt Voucher"
			end if
		%>
 </TITLE>
</HEAD>
<BODY BGCOLOR="#CCCCCC" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<TABLE BORDER=0 ALIGN=CENTER CELLSPACING=0 CELLPADDING=0 NOF=LY>
    <TR >
		<TD height="20">&nbsp;</TD>
		<TD WIDTH=549 ><P ALIGN=CENTER><B>
		<FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT">No Records Found</FONT></B></TD>
	</TR>
	<TR>
        <TD height="20">&nbsp;</TD>
        <TD WIDTH=549 ><P ALIGN=CENTER><B>
        <FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT"><a href="javascript:window.history.back(1)">Back</a></FONT></B></TD>
    </TR>
</Table>
</HTML>

