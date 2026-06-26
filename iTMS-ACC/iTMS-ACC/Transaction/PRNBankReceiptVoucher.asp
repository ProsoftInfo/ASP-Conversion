<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PRNBankReceiptVoucher.asp
	'Module Name				:
	'Author Name				:	KUMAR K A	
	'Created On					:	10 Dec 2007
	'Modified By				:	S.MAHESWARI
	'Modified On				:   24 Sep 2008	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'Connects To				:	BankVouchView_San.asp
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/PrintFunctions.asp"-->
<!--#include File="../../include/GetOrganization.asp"-->

			<%
'------------------------Declaration Constants -----------------------------
dim aiHeaderColWidth(5,9),objFSO,objTxt
dim sTextOut,sTempStr, nCashVoucherNo,dCashVoucherDate,sHeadOfAccount,sAmount,sPaidTo,iAmount,iAccAmount
dim sDetails, sReceivedPayment,sPreparedBy,sCheckedBy,sPassedBy,dcrs,iPageNo,sRetVal

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
			Dim iCreatedBy,sCreatedOn,sVouStatus,sEmpName,iPartyCtrlAcc,sAdjType
			Dim iEntryNo,iHeadOfAcc,sHeadOfAccName,iHeadOfAccAmt,iPartyCode,iEntryAmt,iCtr
			Dim iNetAmtPaid,iTotRecovered,sNarrFlag,sAddnFlag,iTotAddnlAmt,sAdjFlag,sAdjOn,iTotAdj
			' Create our DOM Document Objects
			Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
			Dim iOne,iThree,iNoOfLinesCtr,iFlag12,iFlag34,objRs1
			Dim sPartyCode,sPartyAddress1,sPartyAddress2, sPartyCity, sPartyPinCode,sParty, sPartyAddString, sPartyAddSplt,i,iSpltCnt
			dim iAccHead,sAccHeadName
			iTransNo=Request("Value")
			'Response.Write iTransNo
			iFlag12 = false
			iFlag34 = false
			iNoOfLinesCtr = 0
			i = 0
			sPartyCode = "":sPartyAddress1="":sPartyAddress2="": sPartyCity="": sPartyPinCode=""
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
			set objRs1 = Server.CreateObject("ADODB.Recordset")
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

			'==============================Voucher View header queries starts here ============================

			sQuery = "select Distinct CreatedVoucherNo,TransactionType,OUDefinitionID,PayToRecdFrom," &_
					"PartyCode,CreatedBy,Convert(varchar,CreatedOn,103),CreatedVouchStatus," &_
					"BankInstrumentType,BankInstrumentNo,Convert(varchar,BankInstrumentDate,103),PayableAt,DrawnOnBank" &_
					" from VW_Created_BankVoucherView where CreatedTransNo="& iTransNo

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
				sBankInsNo = objRs(9)
				sBankInsDt = objRs(10)
				sBankInsDrawnOn = objRs(12)

				If sInstrType = "C" Then
					sBankInsName = "Cheque "
				Elseif sInstrType = "D" Then
					sBankInsName = "Demand Draft "
				ElseIf sInstrType = "B" Then
					sBankInsName = "Bankers Cheque "
				Else
					sBankInsName = "Telegraphic Transfer "
				End If
			end if
			objRs.Close

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
			sQuery = "Select VoucherEntryNumber,AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"'"
			'Response.Write sQuery &"<BR>"  
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with

			set objRs.ActiveConnection = nothing
				if not 	objRs.EOF then
					iEntryNo = objRs(0)
					iHeadOfAcc = objRs(1)
					iHeadOfAccAmt = FormatNumber(objRs(2),2,,,-2)
					sNarr = objRs(3)
					iPartyCtrlAcc = objRs(4)
				end if
			objRs.Close
			'Response.Write "iHeadOfAcc="&iHeadOfAcc

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

			'===================================================================================================
			'  To Find Out the Party Details
			IF iPartyCtrlAcc <> "" Then
				sQuery = "SELECT partyname,AddressLine1,APP_M_PartyMaster.AddressLine2,City,Pincode FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
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
						sPartyCode = iPartyCtrlAcc
						sParty = objrs(0)
						sPartyAddress1 = objrs(1)
						sPartyAddress2 = objrs(2)
						sPartyCity = objrs(3)
						sPartyPinCode = objrs(4)
					end if
				objRs.Close
				sPartyAddString = sParty
				If sPartyAddress1 <> "" Then
					sPartyAddString = sPartyAddString & "~" & sPartyAddress1
				End If
				If sPartyAddress2 <> "" Then
					sPartyAddString = sPartyAddString & "~" & sPartyAddress2
				End If
				If sPartyCity <> "" Then
					sPartyAddString = sPartyAddString & "~" & sPartyCity
				End If
				If sPartyPinCode <> "" Then
					sPartyAddString = sPartyAddString & "-" & sPartyPinCode
				End If
				
				sPartyAddSplt = Split(sPartyAddString,"~")
				iSpltCnt = uBound(sPartyAddSplt)
			End IF
			
			'===================================================================================================

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

			If iPartyCode <> "0" Then
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
			sQuery = "SELECT LoginId FROM DCS_User WHERE InternalUserId ="&iCreatedBy
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
			sExp = "//BankInstrumentDet"
			Set tempNode = Root.selectNodes(sExp)
				For each EntryNode in Root.childNodes
					if EntryNode.nodeName="Entry" then
						sPayTo = EntryNode.Attributes.Item(2).nodeValue
					end if
					For each HeaderNode in EntryNode.childNodes
						if HeaderNode.nodeName="AccHead" then
							sAccNo = Split(HeaderNode.attributes.item(0).nodeValue,"?")
							sType = HeaderNode.Attributes.Item(4).nodeValue
						end if
						if HeaderNode.nodeName="Narration" then
							sNarr= HeaderNode.text
						end if
					next
				next
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
				'Newly added by Maheswari on Sep 25 th 2008
				sQuery = "Select IsNull(AccUnitAccountHead,0),Amount from VW_Created_BankVoucherView where createdtransno = "& iTransNo 
				objRs.Open sQuery,con
				do while not objRs.EOF 
					iAccHead = objrs(0)
					iAccAmount = objRs(1)
					If 	trim(iAccHead) <> "0" then 
						sQuery = "Select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iAccHead
						objRs1.Open sQuery,con
						IF not objRs1.EOF then 
							sAccHeadName = objrs1(0) 	
						End IF
						objRs1.Close
					 
					End If
					objRs.MoveNext 
				loop
				objrs.Close
				'Response.Write iAccAmount
				'Response.End 
				
'=============================================================================================================================================
'Party Code
aiHeaderColWidth(0,1)=28

'Head of Account, Amount
aiHeaderColWidth(1,0)=40
aiHeaderColWidth(1,1)=1
aiHeaderColWidth(1,2) = 28

'Paid To, Rupees
aiHeaderColWidth(2,0)=14
aiHeaderColWidth(2,1)=46
aiHeaderColWidth(2,2)=7
aiHeaderColWidth(2,3)=16

'NameOfBank, Details
aiHeaderColWidth(3,0)=3
aiHeaderColWidth(3,1)=25
aiHeaderColWidth(3,2)=1
aiHeaderColWidth(3,3)=12
aiHeaderColWidth(3,4)=1
aiHeaderColWidth(3,5)=10
aiHeaderColWidth(3,6)=1
aiHeaderColWidth(3,7)=8
aiHeaderColWidth(3,8)=1
aiHeaderColWidth(3,9)=20

'Prepared by, Checked by, Passed by
aiHeaderColWidth(4,0)=3
aiHeaderColWidth(4,1)=16
aiHeaderColWidth(4,2)=3
aiHeaderColWidth(4,3)=18
aiHeaderColWidth(4,4)=2
aiHeaderColWidth(4,5)=18

'------------------------End of Declaration Constants ----------------------
iPageNo=1

	'Assigning Hardcoded values for Variables
	nCashVoucherNo		= "CashVoucherNo"
	dCashVoucherDate	= "CashVoucherDate"
	sHeadOfAccount		= "HeadOfAccount"
	sAmount					= "Amount"
		if sType="P" then
			sPaidTo					= sPartyName
		else
			sPaidTo					= sPayTo
		end if
	'iAmount					= "Amount"
	sDetails					= "Details"
	sReceivedPayment	= "ReceivedPayment"
	sPreparedBy			= "PreparedBy"
	sCheckedBy				= "CheckedBy"
	sPassedBy				= "PassedBy"

	'Blank Lines
	'sTextOut = sTextOut & " " & vbcrlf  & " " & vbcrlf & " " & vbcrlf  & " " & vbcrlf
	
	
	'Company  Name (Added for time being-have to change)
	'Response.Write 	sOrgName  & vbcrlf  Response.Write 	sAddress1 & vbcrlf  	Response.Write 	sAddress2  & vbcrlf  	Response.Write 	sCity  & vbcrlf  
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(0,1)-10,"L")
	
	'sTextOut = sTextOut & myAlign("",8,"L")
	'str=Chr(27) & "W1" & str & Chr(27) + "W0"
	sTextOut = sTextOut & chr(27) & "P" & chr(27) & "W1" & myAlign(sOrgName,aiHeaderColWidth(1,0),"L") &chr(27) & "P" & chr(27) & "W0" & vbCrLf
	sTextOut = sTextOut 
	sTextOut = sTextOut & myAlign("",8,"L") & centerAlign(trim(sAddress1),70)& vbCrLf
	sTextOut = sTextOut & myAlign("",8,"L") & centerAlign(trim(sAddress2)&","&trim(sCity),70)& vbCrLf
	
	sTextOut = sTextOut & " " & vbcrlf & vbcrlf    
	iNoOfLinesCtr = iNoOfLinesCtr + 3
	'Response.End 
	'PartyCode
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(0,1)+6,"L")
	sTextOut = sTextOut & myAlign("PARTY CODE: " & sPartyCode,aiHeaderColWidth(0,1),"L") & vbCrLf
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 2
	' Party Address,VoucherNo,Date,Cheque No
	'Line 1
	sTextOut = sTextOut & myAlign("",8,"L")
	sTextOut = sTextOut & myAlign("TO",aiHeaderColWidth(1,0),"L")
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
	sTextOut = sTextOut & myAlign("DATE       :"&sVoucDate,aiHeaderColWidth(1,2),"L")
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 1
	'Line 2
	If i < iSpltCnt Then
		sTextOut = sTextOut & myAlign("",8,"L")	 
		sTextOut = sTextOut & Space(3) & myAlign(sPartyAddSplt(i),aiHeaderColWidth(1,0)-2,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1),"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	Else
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,0),"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	End If

	'Line 3
	If i < iSpltCnt Then
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & Space(3) & myAlign(sPartyAddSplt(i),aiHeaderColWidth(1,0)-2,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
		sTextOut = sTextOut & myAlign("VOUCHER NO.:"&iVouNo,aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	Else
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,0),"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
		sTextOut = sTextOut & myAlign("VOUCHER NO.:"&iVouNo,aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	End If

	'Line 4
	If i < iSpltCnt Then
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & Space(3) & myAlign(sPartyAddSplt(i),aiHeaderColWidth(1,0)-2,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	Else
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,0),"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	End If

	'Line 5
	If i < iSpltCnt Then
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & Space(3) & myAlign(sPartyAddSplt(i),aiHeaderColWidth(1,0)-2,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
		sTextOut = sTextOut & myAlign("CHEQUE NO. :"&sBankInsNo,aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	Else
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,0),"L")
		sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,1)+5,"L")
		sTextOut = sTextOut & myAlign("CHEQUE NO. :"&sBankInsNo,aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		i = i + 1
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	End If
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 1	
	

	'To Print String Content
	sTextOut = sTextOut & myAlign("",7,"L")
	sTempStr = "Dear Sir,"
	sTextOut = sTextOut & myAlign(sTempStr,10,"L") & vbCrLf
	sTextOut = sTextOut & vbCrLf
	sTextOut = sTextOut & myAlign("",7,"L")
	sTempStr = "Following are the details for the payment made by our cheque attached."
	sTextOut = sTextOut & myAlign(sTempStr,80,"L") & vbCrLf
	sTextOut = sTextOut & myAlign("",7,"L")
	sTempStr = "Kindly send your Official Stamped Receipt:"
	sTextOut = sTextOut & myAlign(sTempStr,80,"L") & vbCrLf
	sTempStr = ""
	iNoOfLinesCtr = iNoOfLinesCtr + 4	
	'To Print Bill Details
	sTextOut = sTextOut & myAlign("",7,"L")
	sTextOut = sTextOut & string(76,"-") & vbCrLf
	sTextOut = sTextOut & myAlign("",8,"L")
	sTextOut = sTextOut & myAlign("BILL/DEBIT NOTE NO.&DATE",aiHeaderColWidth(3,1),"L")
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,2),"L")
	sTextOut = sTextOut & myAlign("BILL AMOUNT",aiHeaderColWidth(3,3),"L")
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,4),"L")
	sTextOut = sTextOut & myAlign("DEBIT AMOUNT",aiHeaderColWidth(3,5)+2,"L")
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,6),"L")
	sTextOut = sTextOut & myAlign("TAX",aiHeaderColWidth(3,7),"L")
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,8),"L")
	sTextOut = sTextOut & myAlign("REMARKS",aiHeaderColWidth(3,9),"L")
	sTextOut = sTextOut & vbCrLf
	sTextOut = sTextOut & myAlign("",7,"L")
	sTextOut = sTextOut & String(76,"-") & vbCrLf
		
	'Amount
	
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(1,0),"L")
	'sTextOut = sTextOut & myAlign(AmountWords(replace(iHeadOfAccAmt,",","")),aiHeaderColWidth(1,1),"L")
	'sTempStr = AmountWords(iHeadOfAccAmt)
	'sTextOut = sTextOut & myAlign(sTempStr,aiHeaderColWidth(1,1),"L")
	 
	sTextOut = sTextOut & sTempStr
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 4
	sTempStr = ""
	
	 
%>
<%
Dim iPayableNo, iCreatedTransNo, iCreatedVoucherNo,sVoucherDate,iVoucherAmount, iDebitAmt, iTax, sRemarks, iEntFlag
	iDebitAmt = 0  : iTax = 0 : sRemarks = ""

			'To Fetch Payable Adjustment Details
			sQuery = "SELECT DISTINCT PayablesNumber FROM Acc_T_CreatedPybleAdjDet WHERE CreatedTransNo ="&iTransNo

			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection = nothing
			If Not objRs.EOF Then
			While Not objrs.EOF
				iPayableNo = objrs(0)
				sQuery = "SELECT CreatedTransNo FROM Acc_T_CreatedPayables WHERE PayablesNumber = " & iPayableNo
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with

				set objRs1.ActiveConnection = nothing
				If Not objRs1.EOF Then
					iCreatedTransNo = objrs1(0)
				End If
				objRs1.Close

				sQuery = "SELECT CreatedVoucherno,CONVERT(CHAR,VoucherDate,103),VoucherAmount FROM Acc_T_CreatedVoucherHeader WHERE createdTransNo = " & iCreatedTransNo
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing

				If Not objRs1.EOF Then
						iCreatedVoucherNo = objRs1(0)
						sVoucherDate = objrs1(1)
						iVoucherAmount = objRs1(2)
						iEntFlag = True
						sTextOut = sTextOut & myAlign("",8,"L")
						sTextOut = sTextOut & myAlign(iCreatedVoucherNo,aiHeaderColWidth(3,1)-11,"L")
						sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,2),"L")
						sTextOut = sTextOut & myAlign(sVoucDate,aiHeaderColWidth(3,3),"L")
						sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,4),"L")
						sTextOut = sTextOut & myAlign(FormatNumber(iVoucherAmount,2,,0),aiHeaderColWidth(3,5),"R")
						sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,6),"L")
						sTextOut = sTextOut & vbCrLf
						iNoOfLinesCtr = iNoOfLinesCtr + 1
				End If
				objRs1.Close
			objrs.MoveNext
			Wend
			End If
			objrs.Close

			'  To Print Receivable Details
			sQuery = "SELECT DISTINCT ReceivableNumber FROM Acc_T_CreatedRcvbleAdjDet WHERE CreatedTransNo ="&iTransNo

			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection = nothing
			If Not objRs.EOF Then
			While Not objrs.EOF
				iPayableNo = objrs(0)
				sQuery = "SELECT CreatedTransNo FROM Acc_T_CreatedReceivables WHERE ReceivableNumber = " & iPayableNo
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with

				set objRs1.ActiveConnection = nothing
				If Not objRs1.EOF Then
					iCreatedTransNo = objrs1(0)
				End If
				objRs1.Close

				sQuery = "SELECT CreatedVoucherno,CONVERT(CHAR,VoucherDate,103),VoucherAmount FROM Acc_T_CreatedVoucherHeader WHERE createdTransNo = " & iCreatedTransNo
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing
				If Not objRs1.EOF Then
						iCreatedVoucherNo = objRs1(0)
						sVoucherDate = objrs1(1)
						iVoucherAmount = objRs1(2)
						iEntFlag = True
						'Response.Write "iCreatedVoucherNo="&iCreatedVoucherNo&"<BR>"
						sTextOut = sTextOut & myAlign("",8,"L")
						sTextOut = sTextOut & myAlign(iCreatedVoucherNo,aiHeaderColWidth(3,1)-11,"L")
						sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,2),"L")
						sTextOut = sTextOut & myAlign(sVoucDate,aiHeaderColWidth(3,3),"L")
						sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,4),"L")
						sTextOut = sTextOut & myAlign(FormatNumber(iVoucherAmount,2,,0),aiHeaderColWidth(3,5),"R")
						sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,6),"L")
						sTextOut = sTextOut & vbCrLf
						iNoOfLinesCtr = iNoOfLinesCtr + 1
				End If
				objRs1.Close
			objrs.MoveNext
			Wend
			End If
			objrs.Close

'Changed by Maheswari

	sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> 0 and TransCrDrIndication='"&sTranEntryIndication&"'"
	'Response.Write sQuery
	'Response.End
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
		Do While not objRs.EOF
			iHeadOfAcc = objRs(0)
			iEntryAmt = cdbl(objRs(1))
			iTotRecovered = iTotRecovered + iEntryAmt

'			iEntFlag = True

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


			If Not sNarrFlag Then

				sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
				sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
				sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
				'sTempStr = sTempStr & myAlign(UCase(sNarr),aiHeaderColWidth(3,3),"L")
				sTextOut = sTextOut & sTempStr
				sTextOut = sTextOut & vbCrLf
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sTempStr = ""
			sNarrFlag = True
			End If
					Response.Write sTempStr
					'response.End
							If iOne < 1 and iThree > 0 and not iFlag12 and iCtr < 2  Then
								If iCtr = 1 Then
									'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")

									'sTempStr = sTempStr & myAlign("Recoveries",aiHeaderColWidth(3,3),"L")&vbCrLf

									'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
									'sTempStr = sTempStr & myAlign(iHeadOfAcc,aiHeaderColWidth(3,3),"L")
									'sTempStr = sTempStr & myAlign(iEntryAmt,aiHeaderColWidth(3,9),"L")
									'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5),"L")
									'sTempStr = sTempStr & myAlign("111222222",aiHeaderColWidth(3,9),"L")
									'sTextOut = sTextOut & sTempStr
									'sTextOut = sTextOut & vbCrLf
									'sTempStr = ""
									'iNoOfLinesCtr = iNoOfLinesCtr + 1
									IF iEntFlag = True then
										'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
									 	'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5)-1,"L")
									'	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
										'sTempStr = sTempStr & myAlign("Recoveries",aiHeaderColWidth(3,3),"L") '1
										sTextOut = sTextOut & myAlign("",8,"L")
										sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0)+aiHeaderColWidth(3,0),"L")
										sTempStr = sTempStr & myAlign("Total",aiHeaderColWidth(3,1),"L")
										sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
										sTempStr = sTempStr & myAlign(iHeadOfAccAmt,aiHeaderColWidth(3,9)-6,"R")
										sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
										sTempStr = sTempStr & myAlign("Adv : ",aiHeaderColWidth(3,5)-3,"L")
										sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
										sTempStr = sTempStr & myAlign("0.00",aiHeaderColWidth(3,5),"R")
										'sTempStr = sTempStr & myAlign("111222222",aiHeaderColWidth(3,9),"L")
										sTextOut = sTextOut & sTempStr
										sTextOut = sTextOut & vbCrLf
										sTempStr = ""
										iNoOfLinesCtr = iNoOfLinesCtr + 1
										iFlag12 = true
									End IF
								End If
								if not iFlag34 then
								
									IF iEntFlag = True then
									
										IF trim(iAccAmount)  <> "" then 
											'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
											'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5)-2,"L")
											sTextOut = sTextOut & myAlign("",8,"L")
											sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0)+aiHeaderColWidth(3,0),"L")
											'sTempStr = sTempStr & myAlign("(-)RECOVERED",aiHeaderColWidth(3,3),"L")
											sTempStr = sTempStr & myAlign(sAccHeadName,aiHeaderColWidth(3,1),"L")
											sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
											sTempStr = sTempStr & myAlign(FormatNumber(iAccAmount,2,,,-2),aiHeaderColWidth(3,9)-6,"R")									
											sTextOut = sTextOut & sTempStr
											sTextOut = sTextOut & vbCrLf
											sTempStr = ""
											iNoOfLinesCtr = iNoOfLinesCtr + 1
										End IF
									End if
								End if
							else
											If iCtr = 1 Then
												sTextOut = sTextOut & myAlign("",8,"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
												sTempStr = sTempStr & myAlign("Recoveries",aiHeaderColWidth(3,3),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5),"R")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,9),"L")
												sTextOut = sTextOut & sTempStr
												sTextOut = sTextOut & vbCrLf
												sTempStr = ""
												iNoOfLinesCtr = iNoOfLinesCtr + 1
											End If
												sTextOut = sTextOut & myAlign("",8,"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
												sTempStr = sTempStr & myAlign(sHeadOfAccName,aiHeaderColWidth(3,3),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
												sTempStr = sTempStr & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(3,5),"R")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
												sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,9),"L")
												sTextOut = sTextOut & sTempStr
												sTextOut = sTextOut & vbCrLf
												sTempStr = ""
												iNoOfLinesCtr = iNoOfLinesCtr + 1
							end If
		objRs.MoveNext
		iCtr = iCtr + 1
		Loop
	end if
	objRs.Close

	iNetAmtPaid = iHeadOfAccAmt + iTotAddnlAmt - iTotRecovered

															If sNarrFlag Then
															IF iEntFlag = True then
																'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
																'sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5)-2,"L")
																sTextOut = sTextOut & myAlign("",8,"L")
																sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0)+aiHeaderColWidth(3,0),"L")
																sTempStr = sTempStr & myAlign("Net Amount",aiHeaderColWidth(3,1),"L")
																sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
																sTempStr = sTempStr & myAlign(FormatNumber(iTotRecovered,2,,,-2),aiHeaderColWidth(3,9)-6,"R")
																sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
																sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
																sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
																sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,9),"L")
																sTextOut = sTextOut & sTempStr
																sTextOut = sTextOut & vbCrLf
																sTempStr = ""
																iNoOfLinesCtr = iNoOfLinesCtr + 1
'END Fetch the Recoveries
																End If
															End If


															'Fetch the Receivable Adjustments
															sQuery = "select ReceivableNumber,Convert(varchar,ReceivedOn,103),AmountReceived,AdjustType from Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
															'Response.Write sQuery
															'Response.End
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
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	sAdjOn = objRs(1)
																	iEntryAmt = cdbl(objRs(2))
																	sAdjType = objRs(3)
																	iTotAdj = iTotAdj + iEntryAmt
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close

														'Fetch the Payable Adjustments
															sQuery = "select PayablesNumber,Convert(varchar,PaidOn,103),AmountPaid,AdjustType from Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
															'Response.Write sQuery
															'Response.End
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
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	sAdjOn = objRs(1)
																	iEntryAmt = cdbl(objRs(2))
																	sAdjType = objRs(3)
																	iTotAdj = iTotAdj + iEntryAmt

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

																	If Not sNarrFlag Then
																		sTextOut = sTextOut & myAlign("",8,"L")
																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
																		sTempStr = sTempStr & myAlign(UCase(sNarr),aiHeaderColWidth(3,3),"L")''																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5),"R")
																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
																		sTempStr = sTempStr & myAlign("000111111",aiHeaderColWidth(3,7),"L")
																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
																		sTempStr = sTempStr & myAlign("000222222",aiHeaderColWidth(3,9),"L")
																		sTextOut = sTextOut & sTempStr
																		sTextOut = sTextOut & vbCrLf
																		sTempStr = ""
																		iNoOfLinesCtr = iNoOfLinesCtr + 1
																	sNarrFlag = True
																	End If
																	If Not sAdjFlag Then
																		If  iOne=0 and iThree=0  and not iFlag12 then
																			sTextOut = sTextOut & myAlign("",8,"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
																			sTempStr = sTempStr & myAlign("Adjustments",aiHeaderColWidth(3,3),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5),"R")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
																			sTempStr = sTempStr & myAlign("111111111",aiHeaderColWidth(3,7),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
																			sTempStr = sTempStr & myAlign("111222222",aiHeaderColWidth(3,9),"L")
																			sTextOut = sTextOut & sTempStr
																			sTextOut = sTextOut & vbCrLf
																			sTempStr = ""
																			iFlag12 = true
																			iNoOfLinesCtr = iNoOfLinesCtr + 1
																		Else
																			sTextOut = sTextOut & myAlign("",8,"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
																		sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
																			sTempStr = sTempStr & myAlign("Adjustments",aiHeaderColWidth(3,3),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,5),"R")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
																			sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,9),"L")
																			sTextOut = sTextOut & sTempStr
																			sTextOut = sTextOut & vbCrLf
																			sTempStr = ""
																			iNoOfLinesCtr = iNoOfLinesCtr + 1
																		End if

																		sAdjFlag = True
																	End If
															If trim(sAdjType) = "A" Then sHeadOfAccName = "Less Advance Payments "
																If  iOne=0 and iThree=0  and not iFlag34 and iCtr=1 then
																	sTextOut = sTextOut & myAlign("",8,"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
																	sTempStr = sTempStr & myAlign(sHeadOfAccName,aiHeaderColWidth(3,3),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
																	sTempStr = sTempStr & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(3,5),"R")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
																	sTempStr = sTempStr & myAlign("111333333",aiHeaderColWidth(3,7),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
																	sTempStr = sTempStr & myAlign("111444444",aiHeaderColWidth(3,9),"L")
																	sTextOut = sTextOut & sTempStr
																	sTextOut = sTextOut & vbCrLf
																	sTempStr = ""
																	iFlag34 = true
																	iNoOfLinesCtr = iNoOfLinesCtr + 1
																Else
																	sTextOut = sTextOut & myAlign("",8,"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,0),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,1),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,2),"L")
																	sTempStr = sTempStr & myAlign(sHeadOfAccName,aiHeaderColWidth(3,3),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,4),"L")
																	sTempStr = sTempStr & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(3,5),"R")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,6),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,7),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,8),"L")
																	sTempStr = sTempStr & myAlign("",aiHeaderColWidth(3,9),"L")
																	sTextOut = sTextOut & sTempStr
																	sTextOut = sTextOut & vbCrLf
																	sTempStr = ""
																	iNoOfLinesCtr = iNoOfLinesCtr + 1
																End if

																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close

'********************************************
	'Response.Write iNoOfLinesCtr
	'Response.End 
	For i = iNoOfLinesCtr to 35
		sTextOut = sTextOut & vbCrLf 
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	Next	
	'sTextOut = sTextOut & chr(12)
	sTextOut = Replace(sTextOut, "000111111","Total Amount")
	sTextOut = Replace(sTextOut, "000222222", FormatNumber((iHeadOfAccAmt + iTotAddnlAmt),2,,,-2))
	'sTextOut = Replace(sTextOut, "111111111","(-) Recovered")
	sTextOut = Replace(sTextOut, "111111111",sAccHeadName)
	
	'sTextOut = Replace(sTextOut, "111222222", FormatNumber(iAmount,2,,,-2))
	sTextOut = Replace(sTextOut, "111333333","Net Amount  ")
	sTextOut = Replace(sTextOut, "111444444", FormatNumber(iNetAmtPaid,2,,,-2))
	
	
	
	'Amount in words
	sAmount = MID(AmountWords(iHeadOfAccAmt),8)
	iCtr = 0
	'sTextOut = sTextOut & myAlign(AmountWords(replace(iHeadOfAccAmt,",","")),aiHeaderColWidth(1,1),"L")
	sTextOut = sTextOut & myAlign("",10,"L")
	'sTextOut = sTextOut & myalign("RUPEES :",10,"L")
	sAmount = "RUPEES :" &sAmount 
	sAmount =  BreakString(sAmount,30)
	For i = 0 to UBOUND(sAmount)
		IF i > 0 then sTempStr = sTempStr & myAlign("",10,"L")	
		sTempStr = sTempStr & sAmount(i)
		sTempStr = sTempStr & myAlign(sTempStr,aiHeaderColWidth(1,1),"L") 
		iCtr = iCtr + 1	
		If iCtr = 1 then sTempStr = sTempStr & myalign("",10,"L") & myalign("For "&sOrgName,len(sOrgName)+5,"L")					
		sTempStr = sTempStr& vbCrLf 
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	Next
	
	sTextOut = sTextOut & sTempStr
	'sTextOut = sTextOut & vbCrLf
	sTextOut = sTextOut & myAlign("",6,"L")	
	sTextOut = sTextOut & myalign("",50,"L") & myalign("Authorised Signatory",25,"L")						
	sTempStr = "" 
	sTextOut = sTextOut & vbCrLf & vbCrLf
	sTextOut = sTextOut & vbCrLf
	
	iNoOfLinesCtr = iNoOfLinesCtr + 3
	'*******************Cheque details**************************
	If 1 = 3 then 'Condition added to block the following cheque details print
		sTextOut = sTextOut & myAlign("",66,"L")	
		sTextOut = sTextOut & myAlign(sVoucDate,aiHeaderColWidth(1,2),"L")
		sTextOut = sTextOut & vbCrLf
		sTextOut = sTextOut & myAlign("",8,"L")
		sTextOut = sTextOut & myAlign("",5,"L")
		sTextOut = sTextOut & myAlign(sPartyAddSplt(0),aiHeaderColWidth(1,0),"L")
		sTextOut = sTextOut & vbCrLf & vbCrLf & vbCrLf	
		sTextOut = sTextOut & myAlign("",5,"L")	
		sTextOut = sTextOut & myalign("",8,"L")		
		sTextOut = sTextOut & myalign(MID(AmountWords(iHeadOfAccAmt),8),70,"L")		
		sTextOut = sTextOut & vbCrLf
		sTextOut = sTextOut & myAlign("",65,"L")
		sTextOut = sTextOut & myAlign(""&FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(1,2)-8,"L")
	End If
	'**************************************************************
	objTxt.write sTextOut
	Response.Redirect("../../Components/FormattPrintNew.asp?server=server&filepath=/accounts/temp/Transaction/"&Session.SessionID&"_BankVoucher_View.txt&exitpath=/accounts/reports/VouBAView.asp&frame=_parent")
%>
<%
'================================= USER DEFINED FUNCTIONS ================================='
'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function myAlign(val1,alen,str1)
	dim vlen,k,str2,val
		val=val1
		IF val <> "" then vlen = CInt(len(val))
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
'	'------------------------End OF myAlign Function----------------------------
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

