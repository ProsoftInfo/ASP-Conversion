<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccVouGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Modified By				:	Manohar Prabhu.R
	'Created On					:	January  31, 2003
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
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<%
Dim oDOM,nodHeader,Root,objRs,sQuery,sExp
DIM EntryNode,HeaderNode,nodANL,newElem,TempNode,bAddFlag,nodAnl2
DIM sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
DIM sOrgId,sBookNo,sVouType,iVouNo,sVouName
DIM sVouCode,sApprove,sVoucDate,sAccUnit,sGrpId
DIM dTotal,sTransType,dCRAmt,dDRAmt
DIM sAccType,sAccCode,sEntryType,sEntryno,iTransNo
DIM sDocType,sVouStatus,sPayTo,iPartyReceiptNo,bPartyReceiptFlag
Dim iAdvNo,iTdsPer,objDom,sTemp2,sTempVouNo,sChkVal,sTemparr,sPrVouNo,sPrVouDt
DIM iSeriesNo,iSeriesCode,iCreatedVouNo,iCreatedTransNo,sAccHeadCode,iTdsAmount
Dim iFrmAccCode,iToAccCode,iConChkCnt,sFrmCrDrTy,iFrmBookNo,iFrmBookCode
Dim iMon,iYear,sRevConChk,sErrChk,AdvNode,sConInsChk,dTBDRAmt,dTBCRAmt
Dim iFrmTransNo,sFrmTransTy,sVouEntChk,sParCheck,sTdsCheck,sTdsType
Dim sNewNarr,sSelInvNo,sSelInvDate,sFormVal,sSelVouTy,sRedirTy,sVouYrMon
Dim sTDSNarr,TDSFlag,iAccHeadNo,iTDSEntryAmt,iTDSEntryPer,sConNarr,sLastDate
Dim iPayRecAmt,sFormula ,sVouEntryno,sPartyCrDr,sVouMonYr,iContAccCode,sConTrType

Dim sRetVal,dTotalTdsAmt,sInternalNo


'=============Variables used for contra check
Dim iHeadAccHead,iEntAccHead,iCtr,iEntUnit,iConCount,sHdFlag
Dim sToBookCode,sToBookNo,sToBookName,sNewVouType,sNewVouchNo
Dim sHeadAccName,sEntAccHeadName,iHeadUnit,iCrTransNo
Dim iHeadUnitName,AccNode,NarrNode,sContraCheck,sNewAccVouNo,iNewTransNo
Dim sNewTransType,OldNarrNode,iEntNo
'=================================================================
sRevConChk = "F"
sErrChk = "F"

sVouName=Request("hVouName")
sVouCode=Request("hVouCode")

iCreatedTransNo=Request("hTransNo")
sFormVal = Request("hFormVal")
bPartyReceiptFlag=false
sSelVouTy = Request("vouType")
iVouNo = Session("TrVouNo")
sRedirTy = Request("RedirTy")

Session("TrVouNo") = ""
'Response.Write "sVoucode=" & sVouCode
'Response.Write iCreatedTransNo &"<Br>"
'Response.Write "sVouName="&sVouName
'Response.End

if sVouName="CA" then
	sVouCode="01"
elseif sVouName="BA" then
	sVouCode ="02"
else
	sVouCode="08"
end if


sTdsCheck = "N"
sVouStatus="010104" 'Voucher Accounted

SET objRs  = server.CreateObject("adodb.recordset")
SET oDOM = Server.CreateObject("Microsoft.XMLDOM")
SET objDom = Server.CreateObject("Microsoft.XMLDOM")

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")
sRetVal = GetVouchXML(iCreatedTransNo)
'Response.Write sRetVal
oDOM.Load server.MapPath(sRetVal)




SET Root=oDOM.documentElement

'Response.Write iCreatedTransNo

sQuery = "Select Top 1 CreatedTransNo From Acc_T_VoucherHeader Where CreatedTransNo = "&iCreatedTransNo
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouEntChk = "Y" 'Voucher Already Accounted
Else
	sVouEntChk = "N" 'Voucher Not Accounted
End IF

'sVouEntChk = "N"

objRs.Close
'Response.Write sVouEntChk
'Response.End
sExp = "//AccHead[@Type=""P""]"
Set tempNode = Root.selectNodes(sExp)
IF tempNode.length <> 0 Then
	sParCheck = "Y"
else
	sParCheck = "N"
End IF


'Response.Write sParCheck


IF CStr(sVouEntChk) = "N" Then
	sOrgId=Root.Attributes.Item(0).nodeValue
	sBookNo =Root.Attributes.Item(2).nodeValue
	sVouType=Root.Attributes.Item(4).nodeValue
	sVoucDate=Root.Attributes.Item(5).nodeValue
	sAccHeadCode=Root.Attributes.Item(6).nodeValue
	iCreatedVouNo=Root.Attributes.Item(9).nodeValue



	sPayTo=Replace(Root.childNodes(0).Attributes.Item(2).nodeValue,"'","''")
	Session("BookName") = Root.Attributes.Item(3).nodeValue

	iFrmAccCode = sAccHeadCode
	sFrmCrDrTy = sVouType
	iFrmBookNo = sBookNo
	iFrmBookCode = sVouCode
	iMon = Trim(Mid(sVoucDate,4,2))
	iYear = Year(sVoucDate)
	iMon = CInt(iMon)
	iYear = CInt(iYear)
	'Response.Clear

	sVoucDate = Trim(sVoucDate)

	sVouMonYr = Trim(Mid(sVoucDate,4,2))
	sVouYrMon = Trim(Right(sVoucDate,4))
	sVouMonYr = sVouMonYr & Trim(Right(sVoucDate,4))
	sVouYrMon = sVouYrMon & Trim(Mid(sVoucDate,4,2))



	sLastDate = GetLastDayMonYr(sVouYrMon)
	'Response.Write sVouMonYr &"<br><br>"
	'Response.Write sLastDate
	'Response.End

	FOR EACH EntryNode IN Root.childNodes
		IF EntryNode.nodeName="Entry" THEN
			sPayTo=Replace(EntryNode.Attributes.Item(2).nodeValue,"'","''")
			IF EntryNode.Attributes.Item(1).nodeValue="C" THEN
				dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			ELSE
				dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			END IF
		END IF
	NEXT

	IF StrComp(sVouType,"D")=0 THEN
		dTotal=dCRAmt-dDRAmt
		IF dTotal >0 THEN
			sTransType=sVouName&"R"
		ELSE
			dTotal=Abs(dTotal)
			sVouType="C"
			sTransType=sVouName&"P"
		END IF
	ELSEIF StrComp(sVouType,"C")=0 THEN
		dTotal=dDRAmt-dCRAmt
		IF dTotal >0 THEN
			sTransType=sVouName&"P"
		ELSE
			dTotal=Abs(dTotal)
			sVouType="D"
			sTransType=sVouName&"R"
		END IF
	ELSE
		dTotal=0
		sTransType="GJR"
	END IF

	Select Case CStr(sTransType)
		Case "CAP"	sConTrType = "CSP"
		Case "CAR"	sConTrType = "CSR"
		Case "BAP"	sConTrType = "BSP"
		Case "BAR"	sConTrType = "BSR"
		Case "GJR"	sConTrType = "GJS"
	End Select
	con.BeginTrans
	iTdsAmount = 0
	Dim iTdsEntryType
	sExp = "//TDS"
	Set tempNode = Root.selectNodes(sExp)
	IF tempNode.length <> 0 Then
		For iCtr = 0 To tempNode.length - 1
			iTdsAmount = iTdsAmount + CDbl(tempNode.Item(iCtr).Attributes.getNamedItem("PayRecAmount").Value)
		Next
	Else
		iTdsAmount = 0
	End IF

	dTotalTdsAmt = iTdsAmount

	dTotal = CDbl(dTotal) - CDbl(iTdsAmount)



	sQuery = "Select isNull(OtherApplnTableName,'') From Acc_T_CreatedVoucherHeader  "&_
			 "Where CreatedTransNo = "&iCreatedTransNo&" "

	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sChkVal = Trim(objRs(0))
	End IF
	objRs.Close

	sChkVal = ""

	'Response.Write sChkVal &"<br>"

	IF Len(sChkVal) <> 0 Then
		sTemparr = Split(sChkVal,":")

		sQuery = "Select VoucherNumber,Convert(Char,VoucherDate,103) From Acc_T_VoucherHeader Where CreatedTransNo = "&sTemparr(1)&" "
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			sPrVouNo = objRs(0)
			sPrVouDt = objRs(1)
		End IF
		objRs.Close

		sQuery = "Select VoucherNumber From Acc_T_VoucherHeader Where OUDefinitionID = '"&sOrgId&"' and "&_
				 "BookCode = '"&sVouCode&"' and BookNumber = "&sBookNo&" and  VoucherNumber Like '"&sPrVouNo&"%' and  "&_
				 "Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sPrVouDt&"',103) "

		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With

		Set objRs.ActiveConnection = Nothing
		IF Not objRs.EOF Then
			objRs.MoveLast
			sTempVouNo = objRs(0)
		End IF
		objRs.Close

		sTemp2 = Trim(Right(sTempVouNo,3))
		IF CStr(Mid(sTemp2,1,1)) = "-" Then
			sTemp = CStr(Mid(sTemp2,2))
			sTemp = CDbl(sTemp) + 1
			IF Len(sTemp) = 1 Then
				sTemp = "0"&sTemp
			End IF
			sTempVouNo = Mid(sTempVouNo,1,Len(sTempVouNo)-3)
			sTempVouNo = sTempVouNo&"-"&sTemp
		Else
			sTempVouNo = sTempVouNo&"-01"
		End IF

		'Response.Write sTempVouNo
		iVouNo = sTempVouNo
		'con.BeginTrans
	Else
		'check for contra entry

		iHeadAccHead = Root.Attributes.getNamedItem("BookAcchead").value
		iHeadUnit = Root.Attributes.getNamedItem("UnitNo").value
		iHeadUnitName = Root.Attributes.getNamedItem("UnitName").value

		sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where AccountHead = "&iHeadAccHead&" "
		'Response.Write sQuery
		iSno = 0
		objRs.Open sQuery,con
		IF Not objRs.EOF Then
			sHeadAccName = objRs(0)
		End IF
		objRs.Close
		sHdFlag = "F"
		sContraCheck = "F"
		sExp = "//Entry[@AccUnit="""&iHeadUnit&"""]"
		Set tempNode = Root.selectNodes(sExp)
		'Response.Write tempNode.length
	'================================= XML Constrauction Starts Here ==================================
		IF tempNode.length <> 0 Then
			'Set oDom = Server.CreateObject("Microsoft.XMLDOM")
			For iCtr = 0 To tempNode.length -1
				iEntNo = Trim(tempNode.Item(iCtr).Attributes.Item(0).nodeValue)
				For Each HeaderNode in tempNode.Item(iCtr).childNodes
					IF HeaderNode.nodeName = "AccHead" Then
						IF HeaderNode.Attributes.Item(4).nodeValue = "G" Then
							iEntAccHead = HeaderNode.Attributes.Item(0).nodeValue
							sEntAccHeadName = HeaderNode.Attributes.Item(3).nodeValue
							sQuery = "select count(1) from Acc_M_ContraEntries where OUDefinitionID = '"&iHeadUnit&"' and FromAccountHead = "&iHeadAccHead&" and ToAccountHead = "&iEntAccHead&" "

							With objRs
								.CursorLocation = 3
								.CursorType = 3
								.Source = sQuery
								.ActiveConnection = con
								.Open
							End With
							Set objRs.ActiveConnection = Nothing
							If Not objRs.EOF Then
								iConCount = objRs(0)
							End IF
							objRs.Close

							IF CInt(iConCount) <> 0 Then
								sContraCheck = "T"
							End IF 'Contra Count Check

						End IF 'GL Account Head Check
					End IF 'Accout Head Type Check
				Next 'Same Unit Loop
			Next
		End IF
		If trim(sContraCheck) = "T"  Then
			iVouNo = GetBookReceiptNo(sVouCode,sOrgId,-1,sVouType) 'Added on 28th Jan 09 for Contra entry
		Else
			IF strcomp(sVouType,"D")=0 THEN
				sQuery = "select DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
						 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
			ELSE
				sQuery = "select CrSeriesNo,CrSeriesCode from Acc_M_BookNumberSeries where "&_
						 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
			END IF

			'Response.Write "xxx"& sQuery
			objRs.open sQuery,con
				iSeriesNo=objRs(0)
				iSeriesCode=objRs(1)
			objRs.close()

			'con.BeginTrans
			IF CStr(iVouNo) = "" Then
				iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
			End IF
		End If

	End IF

	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close
	DIM sInsType,sInsNo,sInsDate,sPayableat,sDrawnOn

	FOR EACH EntryNode IN Root.childNodes
		IF EntryNode.nodeName="Entry" THEN
			sAmount=EntryNode.Attributes.Item(3).nodeValue
		End If
	Next
	iTdsAmount = 0

'    When The TDS is been Deducted For the Party The TDS Amount Will get Deducted With the Vouicher amount
'and Same Will get Inserted for Bank
	'dTotal = dTotal - 	iTdsAmount
	'Response.end

	'Response.Write "<p>sVouCode"&sVouCode
	IF sVouCode="02" THEN

		sQuery = "SELECT ISNULL(BANKINSTRUMENTTYPE,'C'),ISNULL(BANKINSTRUMENTNO,0), "&_
				 "ISNULL(PAYABLEAT,''),CONVERT(CHAR,ISNULL(BANKINSTRUMENTDATE,VOUCHERDATE),103), "&_
				 "ISNULL(DRAWNONBANK,'') FROM ACC_T_CREATEDVOUCHERHEADER WHERE CREATEDTRANSNO ="&iCreatedTransNo

		'Response.Write sQuery

		objRs.open sQuery,con
			sInsType=objRs(0)
			sInsNo=objRs(1)
			sPayableat=objRs(2)
			sInsDate=objRs(3)
			sDrawnOn=objRs(4)
		objRs.Close

		iFrmTransNo = iTransNo
		sFrmTransTy = sTransType
		sConNarr = "Summary Entry For "&sVouMonYr


		sQuery= "Insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType, "&_
				"PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount, "&_
				"PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus, "&_
				"BankInstrumentType,BankInstrumentNo,PayableAt,BankInstrumentDate,DrawnOnBank) values "&_
				"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"', "&_
				"NULL,"&sAccHeadCode&",'"&iVouNo&"','"&iCreatedVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103), "&dTotal&_
				",'"&sPayTo&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"', "&_
				"'"&sInsType&"','"&sInsNo&"','"&sPayableat&"',convert(datetime,'"&sInsDate&"',103),'"&sDrawnOn&"') "

	ELSEIF sVouType="" or sVouCode="08" THEN
		dTotal = 0
		sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				"PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
				"CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"&_
				"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				"NULL,NULL,'"&iVouNo&"','"&iCreatedVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				",NULL,"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"

	ELSEIF sVouCode="01" THEN ' Added on 4th Apr 2011 to assign the transaction no when Contra entry is from a Cash Book
		iFrmTransNo = iTransNo
		sFrmTransTy = sTransType

		sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				"PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
				"PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"&_
				"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				"NULL,"&sAccHeadCode&",'"&iVouNo&"','"&iCreatedVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
			",'"&sPayTo&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"
	ELSE
		sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				"PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
				"PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"&_
				"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				"NULL,"&sAccHeadCode&",'"&iVouNo&"','"&iCreatedVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				",'"&sPayTo&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"


	END IF

	'Response.Write "<br>" & sQuery & "<BR>"
	con.execute(sQuery)

	'-----------------------PROCESS ENTRY NODES-------------------------------

'---------PROCESS THE CHILD NODES OF ENTRIES FOR DETAIL TABLE UPDATION----
	FOR EACH EntryNode IN Root.childNodes
		IF EntryNode.nodeName="Entry" THEN
			sEntryno=EntryNode.Attributes.Item(0).nodeValue
			sAmount=EntryNode.Attributes.Item(3).nodeValue
			sEntryType=EntryNode.Attributes.Item(1).nodeValue
			sAccUnit=EntryNode.Attributes.Item(4).nodeValue
			'sGrpId=EntryNode.Attributes.getNamedItem("GroupId").value



			sExp = "//Entry[@No="&sEntryno&" and @TdsAmount]"
			Set tempNode = Root.selectNodes(sExp)

			IF CStr(sVouCode) = "01" or CStr(sVouCode) = "02" or CStr(sVouCode) = "08" Then
				IF tempNode.length <> 0 Then
					iTdsAmount = EntryNode.Attributes.Item(6).nodeValue
					iTdsPer = EntryNode.Attributes.Item(8).nodeValue
				Else
					iTdsAmount = 0
					iTdsPer = 0
				End IF
			Else
				iTdsAmount = 0
				iTdsPer = 0
			End IF

			IF sEntryType="C" THEN
				dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			ELSE
				dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			END IF
			IF Cstr(iTdsAmount) = "" Then
				iTdsAmount = 0
				iTdsPer = 0
			End IF
			sTDSNarr = "TDS Entry"



'********************************** TDS Enteris  Starts Here ********************************

			sVouEntryno = sEntryno + 1
			sExp = "//Entry[@No="&sEntryno&"]/AccHead"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.length <> 0 Then
				sAccCode=TempNode.Item(0).Attributes.Item(0).nodeValue
				sAccType=TempNode.Item(0).Attributes.Item(4).nodeValue
			End IF


			FOR EACH HeaderNode IN EntryNode.childNodes

				'Response.Write "<b>Inside EntryNode </b>"
				sInternalNo = 0

				'Added by Maheshwari on Mar 3rd 2007 for TDS Entries
				IF HeaderNode.nodeName="TDS" THEN
					'Response.Write "<br><br>*************** TDS Entry Starts Here **************************************<br><br>"
					sTdsCheck = "Y"
					iAccHeadNo	 = HeaderNode.getAttribute("AccHeadCode")
					iTDSEntryAmt = FormatNumber(HeaderNode.getAttribute("TDSAmount"),2,,,0)
					iTDSEntryPer = HeaderNode.getAttribute("TdsPercentage")
					iPayRecAmt	 = FormatNumber(HeaderNode.getAttribute("PayRecAmount"),2,,,0)
					sFormula	 = HeaderNode. getAttribute("Formula")

					If trim(iTDSEntryPer) = "" then
						iTDSEntryPer = "0"
					End If

					IF CStr(sEntryType) = "C" Then
						sTdsType = "D"
					Else
						sTdsType = "C"
					End IF


					sGrpId = 0
					sVouEntryno = sVouEntryno + 1
					sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
					sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
					sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage,PackCode,ReceiptNo) values ("
					sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
					sQuery=sQuery& ","&sVouEntryno&","&iAccHeadNo&",NULL,NULL,NULL,"
					sQuery=sQuery&" '"&sTDSNarr&"',"&iPayRecAmt&",'"&sTdsType&"',"&iTDSEntryAmt&","&iTDSEntryPer&","&sGrpId&",'"&sInternalNo&"')"

					'Response.Write "<br>" & sQuery & "<BR>"
					con.execute(sQuery)


					' Response.Write "Account Head " & sAccCode &"<br><br><br>"

					'sTemp=Split(sAccCode,"?")


				END IF 'IF HeaderNode.nodeName="TDS" THEN

				IF 	HeaderNode.nodeName="Narration" THEN
						sNarration=Replace(HeaderNode.text,"'","''")
				END IF 'End of Check for Narration Node
			NEXT

			'Response.Write "Account Head " & sAccCode &"<br><br><br>"
			sTemp=Split(sAccCode,"?")

			'Response.Write "<BR>******** End PartyTransaction Table ********<BR>"
			'Response.Write "Total <b>" & dTotalTdsAmt &"</b>"

			sTDSNarr = sTDSNarr &" For the Voucher " & iVouNo &" Dt: " & sVoucDate

			'Response.Write "<b>Tds AMount</b> " & UBound(sTemp)
			'Response.Write "<p>Account Head " & sAccCode &"<br><br><br>"
			IF CDbl(dTotalTdsAmt) > 0 and UBound(sTemp) > 2 Then

				sInternalNo = 0
				IF CStr(sVouType) = CStr(sEntryType) Then
					'Response.Write "Calling From 1 <br><br><br>"
					sInternalNo = GetBookReceiptNo(sVouCode,sOrgId,0,"D") 'Added on 28th Jan 09 for Contra entry
				Else
					sInternalNo = 0
				End IF


				sVouEntryno = sVouEntryno + 1
				sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
				sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage,PackCode,ReceiptNo) values ("
				sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery=sQuery& ","&sVouEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(3)&","&sTemp(1)&", "
				sQuery=sQuery&" '"&sTDSNarr&"',"&dTotalTdsAmt&",'"&sEntryType&"',"&iTDSEntryAmt&","&iTDSEntryPer&","&sGrpId&",'"&sInternalNo&"')"

				'Response.Write "<br>" & sQuery & "<BR>"
				con.execute(sQuery)
			End IF
'Response.Write "++++++++++++ TDS Entry Ends Here ++++++++++++++++++++++++<br><br>"
'********************************** TDS Enteris  Ends Here ********************************
'----------------------------DETAIL TABLE UPDATION-----------------------------------------



		IF CStr(sAccCode) <> "0" Then
			sInternalNo = 0

			'Response.Write sVouType &"<b>=========>>></b>" & sEntryType &"<br><br>"

			IF CStr(sVouType) = CStr(sEntryType) Then

				'Response.Write "Calling From 2 <br><br><br>"
				sInternalNo = GetBookReceiptNo(sVouCode,sOrgId,0,"D") 'Added on 28th Jan for contra entry
			Else
				sInternalNo = 0
			End IF
			sVouEntryno = sVouEntryno + 1
			IF StrComp(sAccType,"G")=0  THEN
				sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
				sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage,ReceiptNo) values ("
				sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
				sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&", "&iTdsPer&",'"&sInternalNo&"' )"
			ELSE
				IF sEntryType="C" THEN
					bPartyReceiptFlag=true
				END IF
				sTemp=Split(sAccCode,"?")

				IF CStr(sTdsCheck) = "Y" Then
					sAmount = dTotal
				End IF

				sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
				sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage,ReceiptNo) values ("
				sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery=sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
				sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&", "&iTdsPer&",'"&sInternalNo&"' )"
			END IF

			'Response.Write "<br>" & sQuery & "<BR>"
			con.execute(sQuery)
		End IF

		'-----------------------END OF DETAIL TABLE UPDATION----------------------
		DIM sCCGroup,sAddCode,sAddGroupCode,sAddRatio,sAddAmount,dAddTotal
		Dim sAdjType,dAddAmt,dSubAmt,iChkAdj

		dAddTotal=0
		dAddAmt = 0
		dSubAmt = 0
		iChkAdj = 0
		'--------PROCESS CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION----
			FOR EACH HeaderNode IN EntryNode.childNodes
		'----------------------PROCESS PAYABLE / RECEVIABLE NODES ----------------

				'Response.Write "============================================================== "
				'Response.Write "<br><br><br>"
				'Response.Clear
				IF 	HeaderNode.nodeName="PayRec" THEN
					FOR EACH  nodANL IN HeaderNode.childNodes
						sAddCode=nodANL.Attributes.getNamedItem("PayableNo").value
						sAddAmount=nodANL.Attributes.getNamedItem("AmtToAdjust").value
						sDocType=nodANL.Attributes.getNamedItem("DocType").value
						sAdjType = nodANL.Attributes.getNamedItem("AdjType").value
						sSelInvNo = nodANL.Attributes.getNamedItem("InvNo").value
						sSelInvDate = nodANL.Attributes.getNamedItem("InvDate").value


						sSelInvNo = Replace(sSelInvNo,"Advance Receipts","ADV REC")
						sSelInvNo = Replace(sSelInvNo,"Advance Payments","ADV PAY")
						sSelInvNo = Replace(sSelInvNo,"PUR INV No","P IN")
						sSelInvNo = Replace(sSelInvNo,"Sales Inv","S IN")
						sSelInvNo = Replace(sSelInvNo,"DEBIT NOTE NO","D No")
						sSelInvNo = Replace(sSelInvNo,"Credit Note No","C No")
						sSelInvNo = Replace(sSelInvNo,"Cr Note For Purchase Inv","C No")
						sSelInvNo = Replace(sSelInvNo,"Credit Notes","C No")



						IF InStr(1,Cstr(sSelInvNo),"Purchase") > 0 Then
							sNewNarr = sNewNarr & Trim(Mid(sSelInvNo,15)) &", "
						ElseIF InStr(1,Cstr(sSelInvNo),"SALE") > 0 Then
							sNewNarr = sNewNarr & Trim(Mid(sSelInvNo,5)) &", "
						ElseIF InStr(1,Cstr(sSelInvNo),"PUR") > 0 Then
							sNewNarr = sNewNarr & Trim(Mid(sSelInvNo,4)) &", "
						Else
							sNewNarr = sNewNarr & sSelInvNo &", "
						End IF



						'IF InStr(1,Cstr(sSelInvNo),"Dt") = 0 Then
						'	IF CStr(sSelInvDate) <> ""  Then
						'		sNewNarr = sNewNarr & "DT:"& Trim(sSelInvDate) &", "
						'	End IF
						'End IF


						'Response.Write sNewNarr &"<br><br>"


						IF CStr(sAccType) = "P" and CStr(Left(sAccCode,2)) = "CR" Then
							'dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)
							Select Case CStr(sAdjType)
								Case "PI"
									dAddAmt = CDbl(dAddAmt) + CDbl(sAddAmount)
								Case "C"
									dAddAmt = CDbl(dAddAmt) + CDbl(sAddAmount)
								Case "I"
									dSubAmt = CDbl(dSubAmt) + CDbl(sAddAmount)
								Case "D"
									dSubAmt = CDbl(dSubAmt) + CDbl(sAddAmount)
								Case "P"
									dSubAmt = CDbl(dSubAmt) + CDbl(sAddAmount)
							End Select

							'Response.Write dAddTotal &"<br><br>"
							dAddTotal=CDbl(dAddAmt)-CDbl(dSubAmt)

						ElseIF CStr(sAccType) = "P" and CStr(Left(sAccCode,2)) = "DR" Then
							Select Case CStr(sAdjType)
								Case "I"
									dAddAmt = CDbl(dAddAmt) + CDbl(sAddAmount)
								Case "D"
									dAddAmt = CDbl(dAddAmt) + CDbl(sAddAmount)
								Case "PI"
									dSubAmt = CDbl(dSubAmt) + CDbl(sAddAmount)
								Case "C"
									dSubAmt = CDbl(dSubAmt) + CDbl(sAddAmount)
								Case "R"
									dSubAmt = CDbl(dSubAmt) + CDbl(sAddAmount)
							End Select

							dAddTotal=CDbl(dAddAmt)-CDbl(dSubAmt)
						Else
							dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)
						End IF

						'dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)
						IF CStr(sAdjType) <> "P" and CStr(sAdjType) <> "R" Then
							IF CDbl(sAddAmount) >0 THEN
								'IF sDocType="C" THEN
								''-----------------------UPDATE PAYABLE TABLE------------------------------
								'	sQuery="insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
								'		"PaidOn,AmountPaid) Values ("&sAddCode&","&iTransNo&","&_
								'		"getdate(),"&sAddAmount&")"
								'		Response.Write sQuery
								'	con.execute(sQuery)
								'ELSEIF sDocType="D" THEN
								'-----------------------UPDATE RECEIVABLE TABLE---------------------------
								'	sQuery="insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
								'		"ReceivedOn,AmountReceived) Values ("&sAddCode&","&iTransNo&","&_
								'		"getdate(),"&sAddAmount&")"
								'	con.execute(sQuery)
								'END IF

								IF CStr(sAdjType) = "I" Then
									sQuery = "Select ReceivableNumber From Acc_T_Receivables Where OUDefinitionID = '"&sOrgId&"'  "&_
											 "and ReceivableNumber = "&sAddCode&" and PartyType = '"&sTemp(0)&"' and  "&_
											 "PartySubType = "&sTemp(1)&" and PartyCode = "&sTemp(3)&"  "
									objRs.Open sQuery,Con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
												 "ReceivedOn,AmountReceived) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
										'con.execute(sQuery)
									Else
										sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
												 "PaidOn,AmountPaid) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
										'con.execute(sQuery)
									End IF
									objRs.Close
									con.execute(sQuery)
								ElseIF CStr(sAdjType) = "PI" Then
									sQuery = "Select PayablesNumber From Acc_T_Payables Where OUDefinitionID = '"&sOrgId&"'  "&_
											 "and PayablesNumber = "&sAddCode&" and PartyType = '"&sTemp(0)&"' and  "&_
											 "PartySubType = "&sTemp(1)&" and PartyCode = "&sTemp(3)&"  "
											 'Response.Write "<p>"& sQuery
									objRs.Open sQuery,Con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
												 "PaidOn,AmountPaid) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
										'con.execute(sQuery)
									Else
										sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
												 "ReceivedOn,AmountReceived) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
	'									con.execute(sQuery)
									End IF
									objRs.Close
									con.execute(sQuery)
								Elseif CStr(sAdjType) = "D" Then
									sQuery = "Select ReceivableNumber From Acc_T_Receivables Where OUDefinitionID = '"&sOrgId&"'  "&_
											 "and ReceivableNumber = "&sAddCode&" and PartyType = '"&sTemp(0)&"' and  "&_
											 "PartySubType = "&sTemp(1)&" and PartyCode = "&sTemp(3)&"  "
									objRs.Open sQuery,Con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
												 "ReceivedOn,AmountReceived) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
										'con.execute(sQuery)
									Else
										sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
												 "PaidOn,AmountPaid) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
										'con.execute(sQuery)
									End IF
									objRs.Close
									con.execute(sQuery)
								Elseif CStr(sAdjType) = "C" Then
									sQuery = "Select PayablesNumber From Acc_T_Payables Where OUDefinitionID = '"&sOrgId&"'  "&_
											 "and PayablesNumber = "&sAddCode&" and PartyType = '"&sTemp(0)&"' and  "&_
											 "PartySubType = "&sTemp(1)&" and PartyCode = "&sTemp(3)&"  "
									objRs.Open sQuery,Con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
												 "PaidOn,AmountPaid) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
										'con.execute(sQuery)
									Else
										sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
												 "ReceivedOn,AmountReceived) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&")"
										'Response.Write sQuery & "<br>"
										'con.execute(sQuery)
									End IF
									objRs.Close
									con.execute(sQuery)
								End IF
							END IF
						Else
							sQuery = "UPDATE Acc_T_AdvancePayments SET AdvanceAdjusted = isNull(AdvanceAdjusted,0) + "&sAddAmount&"  "&_
									 "WHERE AdvanceNumber = "&sAddCode&" "
						'	Response.Write sQuery & "<br>"
							con.execute(sQuery)

							IF CStr(sAdjType) = "P" Then
								sQuery = "Insert Into Acc_T_PybleAdjustmentDetails(PayablesNumber, "&_
										 "PaidByTransactionNo,PaidOn,AmountPaid,AdjustType)  "&_
										 "Values("&sAddCode&","&iTransNo&",Convert(Datetime,'"&sVoucDate&"',103),"&sAddAmount&",'A') "

								'Response.Write sQuery
								Con.Execute sQuery
							Elseif CStr(sAdjType) = "R" Then
								sQuery = "Insert Into Acc_T_RcvblAdjustmentDetails(ReceivableNumber, "&_
										 "RecdByTransactionNo,ReceivedOn,AmountReceived,AdjustType)  "&_
										 "Values("&sAddCode&","&iTransNo&",Convert(datetime,'"&sVoucDate&"',103),"&sAddAmount&",'A') "
								'Response.Write sQuery
								Con.Execute sQuery
							End IF
						End IF
					NEXT
				END IF
		'-------------END OF PROCESSING PAYABLE / RECEVIABLE NODES ---------------
			NEXT


		'-END OF PROCESSING CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION-
		DIM iCrAdvNo
		'-------------PROCESS FOR ADVANCE TABLE UPDATION -------------------------
			IF sAccType="P" AND CDbl(sAmount)> CDbl(dAddTotal) THEN

				'sExp = "//Entry[@No="&sEntryno&"]/AccHead"
				'Set AdvNode = Root.selectNodes(sExp)
				'IF AdvNode.length <> 0 Then
				'	IF AdvNode.Item(0).Attributes.getNamedItem("Type").Value = "P" Then
				'		iCrAdvNo = AdvNode.Item(0).Attributes.getNamedItem("AdvNo").Value
				'	Else
						iCrAdvNo = 0
				'	End IF
				'End IF

		'------------- Included on 04/05/2004 for created advance details -------------------------
				IF CStr(iCrAdvNo) = "0" Then
					sQuery = "Select CreatedAdvanceNo from Acc_T_CreatedAdvances where CreatedTransNo = "&iCreatedTransNo
					objRs.Open sQuery,Con
					If Not objRs.EOF Then
						iCrAdvNo = objRs(0)
					Else
						iCrAdvNo = 0
					End IF
					objRs.Close
				End IF
		'------------- End -------------------------

				'IF CStr(iCrAdvNo) = "0" Then

					sQuery = "Select isNull(Max(AdvanceNumber),0)+1 from Acc_T_AdvancePayments "
					objRs.Open sQuery,Con
					If Not objRs.EOF Then
						iAdvNo = objRs(0)
					End IF
					objRs.Close

					IF sVouType="C" THEN
						IF CDbl(dAddAmt) > 0 THEN
							sQuery="INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
								"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
								" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								""&sTemp(3)&","&sAmount&","&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
							'	Response.Write sQuery & "<br><br><br>"
								con.execute(sQuery)
						Else
							sQuery="INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
								"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
								" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								""&sTemp(3)&","&sAmount&","&CDbl(sAmount)&",NULL,NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
							'	Response.Write sQuery & "<br><br><br>"
								con.execute(sQuery)
						End IF

					ELSEIF sVouType="D" THEN
						IF CDbl(dAddAmt) > 0 THEN
							sQuery="INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
									"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
									" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
									""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
							'		Response.Write sQuery & "<br><br><br>"
									con.execute(sQuery)
						Else
							sQuery="INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
								"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
								" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)&",NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
							'	Response.Write sQuery & "<br><br><br>"
								con.execute(sQuery)
						End IF


					END IF
				'End IF
			END IF
		'-------------END OF PROCESSING FOR ADVANCE TABLE UPDATION ---------------
		END IF
	NEXT
	'------------------END OF PROCESSING ENTRY NODES--------------------------

'Response.Write "<br><br><br>************ Checking ********************************<br><br>"
'Response.End
'Response.Write "<br><br><br>************ Checking ********************************<br><br>"



	IF CStr(sNewNarr) <> "" Then
		IF CStr(sInsType) = "C" Then
			sInsType = "CH.NO: "&sInsNo
		Elseif CStr(sInsType) = "D" Then
			sInsType = "DD.NO: "&sInsNo
		Elseif CStr(sInsType) = "B" Then
			sInsType = "BANK CH.NO: "&sInsNo
		Elseif CStr(sInsType) = "T" Then
			sInsType = "TT.NO: "&sInsNo
		Elseif CStr(sInsType) = "W" Then
			sInsType = "CASH: "&sInsNo
		End IF

		'IF CStr(sInsType) <> "" Then
			'sNewNarr = sInsType&" DT: "&Trim(sInsDate)&", "&sNewNarr
		'	sNewNarr = sInsType&", "&sNewNarr
		'End IF
		sNewNarr = sNewNarr&" - "&sNarration

		sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "

		'Response.Write sQuery &"<br><br>"
		Con.Execute sQuery

		sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "
		'Response.Write sQuery &"<br><br>"
		Con.Execute sQuery

		sQuery = "Update Acc_T_VoucherDetails Set VoucherNarration = '"&sNewNarr&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "
		'Response.Write sQuery &"<br><br>"
		Con.Execute sQuery



	End IF



	'------------------END OF PROCESSING ENTRY NODES--------------------------

	'================== Contra Entry Details Starts Here =====================================================



		iHeadAccHead = Root.Attributes.getNamedItem("BookAcchead").value
		iHeadUnit = Root.Attributes.getNamedItem("UnitNo").value
		iHeadUnitName = Root.Attributes.getNamedItem("UnitName").value
		'In case of contra the created Voucher number is made same as original voucher no
		sNewVouchNo = Root.Attributes.getNamedItem("CreatedVoucherNo").value
'Response.write "<br>sNewVouchNo="&sNewVouchNo&"<br>"
		sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where AccountHead = "&iHeadAccHead&" "
		iSno = 0
		objRs.Open sQuery,con
		IF Not objRs.EOF Then
			sHeadAccName = objRs(0)
		End IF
		objRs.Close
		sHdFlag = "F"
		sContraCheck = "F"
		sExp = "//Entry[@AccUnit="""&iHeadUnit&"""]"
		Set tempNode = Root.selectNodes(sExp)
		'Response.Write tempNode.length
	'================================= XML Constrauction Starts Here ==================================
		IF tempNode.length <> 0 Then
			'Set oDom = Server.CreateObject("Microsoft.XMLDOM")
			For iCtr = 0 To tempNode.length -1
				iEntNo = Trim(tempNode.Item(iCtr).Attributes.Item(0).nodeValue)
				For Each HeaderNode in tempNode.Item(iCtr).childNodes
					IF HeaderNode.nodeName = "AccHead" Then
						IF HeaderNode.Attributes.Item(4).nodeValue = "G" Then
							iEntAccHead = HeaderNode.Attributes.Item(0).nodeValue
							sEntAccHeadName = HeaderNode.Attributes.Item(3).nodeValue
							sQuery = "select count(1) from Acc_M_ContraEntries where OUDefinitionID = '"&iHeadUnit&"' and FromAccountHead = "&iHeadAccHead&" and ToAccountHead = "&iEntAccHead&" "

							With objRs
								.CursorLocation = 3
								.CursorType = 3
								.Source = sQuery
								.ActiveConnection = con
								.Open
							End With
							Set objRs.ActiveConnection = Nothing
							If Not objRs.EOF Then
								iConCount = objRs(0)
							End IF
							objRs.Close

							IF CInt(iConCount) <> 0 Then
								iSno = iSno + 1
								sQuery = "Select BookCode,BookNumber,BookName From Acc_R_ApplicableAccountHeads Where "&_
										 "BookAccountHead = "&iEntAccHead&" and OUDefinitionID = '"&iHeadUnit&"' "
								objRs.Open sQuery,Con
								IF Not objRs.EOF Then
									sToBookCode = Trim(objRs(0))
									sToBookNo = Trim(objRs(1))
									sToBookName = Trim(objRs(2))
								Else
									sToBookCode = "0"

									sToBookNo = "0"
									sToBookName = "0"
								End IF
								objRs.Close

								IF CStr(sToBookCode) <> "0" and CStr(sToBookNo) <> 0 Then
									IF CStr(sHdFlag) = "F" Then 'This to check taht only one time the
										sHdFlag = "T"			'Header Node gets created.
										IF CStr(sToBookCode) = "02" Then
											IF CStr(sVouType) = "D" Then
												sNewTransType = "BAP"
												sNewVouType = "C"
											Else
												sNewTransType = "BAR"
												sNewVouType = "D"
											End IF
										Elseif CStr(sToBookCode) = "01" Then
											IF CStr(sVouType) = "D" Then
												sNewTransType = "CAP"
												sNewVouType = "C"
											Else
												sNewTransType = "CAR"
												sNewVouType = "D"
											End IF
										End IF

										IF strcomp(sNewVouType,"D") = 0 THEN
											sQuery = "select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
												     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sToBookCode&"' and BookNumber= "&sToBookNo
										ELSE
											sQuery = "select CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
												     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sToBookCode&"' and BookNumber= "&sToBookNo
										END IF

										'Response.Write sQuery
										objRs.open sQuery,con
										if not objRs.EOF then
											iSeriesNo=objRs(0)
											iSeriesCode=objRs(1)
									    end if 'if not objRs.EOF then
										objRs.close()

										'sNewVouchNo = GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)  'Blocked on 31st Jan 2009 to make contra entry voucher no same

										IF strcomp(sNewVouType,"D")=0 THEN
											sQuery="select DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
												"OUDefinitionID='"&sOrgId&"' and BookCode='"&sToBookCode&"' and BookNumber= "&sToBookNo
										ELSE
											sQuery="select CrSeriesNo,CrSeriesCode from Acc_M_BookNumberSeries where "&_
												"OUDefinitionID='"&sOrgId&"' and BookCode='"&sToBookCode&"' and BookNumber= "&sToBookNo
										END IF

										objRs.open sQuery,con
										if not objRs.EOF then
											iSeriesNo=objRs(0)
											iSeriesCode=objRs(1)
										end if'	if not objRs.EOF then
										objRs.close()

										'sNewAccVouNo = GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate) 'Blocked on 31st Jan 2009 to make contra entry voucher no same
										sNewAccVouNo = iVouNo ' Added to make contra entry voucher no as same
										'In case of contra the created Voucher number is made same as original voucher no
										'sNewVouchNo = tmpCRVouNo

										sQuery = "Select isNull(Max(CreatedTransNo),0) + 1 From Acc_T_CreatedVoucherHeader "
										objRs.open sQuery,con
											iCrTransNo = objRs(0)
										objRs.close()

										sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
										objRs.open sQuery,con
											iNewTransNo = objRs(0)
											'Response.Write "<p> New TransNo = "&  iNewTransNo
										objRs.Close


										Set Root = objDom.createElement("voucher")
										Root.setAttribute "UnitNo", iHeadUnit
										Root.setAttribute "UnitName", iHeadUnitName
										Root.setAttribute "BookNo", sToBookNo
										Root.setAttribute "BookName", sToBookName
										Root.setAttribute "CRDR", sNewVouType
										Root.setAttribute "VouDate", sVoucDate
										Root.setAttribute "BookAcchead", iEntAccHead
										Root.setAttribute "Approver","Y"
										Root.setAttribute "TransNo",iNewTransNo
										Root.setAttribute "VoucherNo", sNewAccVouNo
										Root.setAttribute "CreatedTransNo",iCrTransNo
										Root.setAttribute "CreatedVoucherNo", sNewVouchNo
										objDom.appendChild Root

									End IF 'Flag Check

									sExp = "//Entry[@No="&iEntNo&"]/Narration"
									Set OldNarrNode = Root.selectNodes(sExp)


									Set EntryNode = objDom.createElement("Entry")
									EntryNode.setAttribute "No", iSno
									EntryNode.setAttribute "CRDR", sVouType
									EntryNode.setAttribute "Payto", sHeadAccName
									EntryNode.setAttribute "Amount", dTotal
									EntryNode.setAttribute "AccUnit", iHeadUnit
									EntryNode.setAttribute "AccName", iHeadUnitName
									EntryNode.setAttribute "TdsAmount", "0.00"
									EntryNode.setAttribute "TDSElgi", "0"
									EntryNode.setAttribute "TdsPercentage", "0.00"

									Set AccNode = objDom.createElement("AccHead")
									AccNode.setAttribute "No", iHeadAccHead
									AccNode.setAttribute "CostCenter", "0"
									AccNode.setAttribute "Analytical", "0"
									AccNode.setAttribute "Name", sHeadAccName
									AccNode.setAttribute "Type", "G"
									AccNode.setAttribute "TransFlag", "W"
									EntryNode.appendChild AccNode

									Dim sNewNarr2

									'Response.Write sInsNo &" ========= "

									IF CStr(sInsType) = "W" and Cstr(sVouType) = "D" Then
										sNewNarr2 = "Ch.No: "
										sNewNarr2 = sNewNarr2 & sInsNo
										sNewNarr2 = sNewNarr2 &" / "
										sNewNarr2 = sNewNarr2 & "Cash Deposited "
									Elseif CStr(sInsType) = "W" and Cstr(sVouType) = "C" Then
										sNewNarr2 = "Ch.No: "
										sNewNarr2 = sNewNarr2 & sInsNo
										sNewNarr2 = sNewNarr2 &" / "
										sNewNarr2 = sNewNarr2 & "Cash Withdrawn "
									Else
										sNewNarr2 = ""
									End IF

									Set NarrNode = objDom.createElement("Narration")
									IF OldNarrNode.length <> 0 Then
										sNewNarr2 = OldNarrNode.Item(0).text &" " & sNewNarr2
										NarrNode.text = sNewNarr2
									Else
										sNewNarr2 = sNarration &" " & sNewNarr2
										NarrNode.text = sNewNarr2
									End IF
									EntryNode.appendChild NarrNode

									Root.appendChild EntryNode

									sContraCheck = "T"
								End IF 'Contra Count Check
							End IF 'BookCode and BookNo Check
						End IF 'GL Account Head Check
					End IF 'Accout Head Type Check
				Next 'Same Unit Loop
			Next
			'oDom.appendChild Root
			objDom.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-Voucher_Bank.xml")

	'================================= XML Constrauction Starts Here ==================================
		End IF

	'IF the xml is created is Over and ContraCheck is also True then the following will take place.

	'Response.Write "Contra Check " & sContraCheck &"<br>"

	Dim sRevVal


	IF CStr(sContraCheck) = "T" Then
		Set objDom = Server.CreateObject("Microsoft.XMLDOM")
		objDom.Load server.MapPath("../Temp/Transaction/"&Session.SessionID&"-Voucher_Bank.xml")

		set Root=objDom.documentElement

		sOrgId = Root.Attributes.getNamedItem("UnitNo").value
		sBookNo = Root.Attributes.getNamedItem("BookNo").value
		sVouType = Root.Attributes.getNamedItem("CRDR").value
		sVoucDate = Root.Attributes.getNamedItem("VouDate").value
		sAccHeadCode = Root.Attributes.getNamedItem("BookAcchead").value
		Root.Attributes.getNamedItem("Approver").value=sApprove
		sPayTo=Replace(Root.childNodes(0).Attributes.getNamedItem("Payto").value,"'","''")

		iToAccCode = sAccHeadCode

		FOR EACH EntryNode IN Root.childNodes
			IF EntryNode.nodeName = "Entry" Then
				IF EntryNode.Attributes.getNamedItem("CRDR").value="C" THEN
					dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.getNamedItem("Amount").value)
				ELSE
					dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.getNamedItem("Amount").value)
				END IF
			End IF
		NEXT

		sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
				 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"&_
				 "("&iCrTransNo&",'"&iHeadUnit&"','"&sToBookCode&"',"&sToBookNo&",'"&sNewTransType&"',"&_
				 "NULL,"&sAccHeadCode&",'"&sNewVouchNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				 ",'"&sPayTo&"','"&sNewVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

		'Response.Write "in="&sQuery &"<BR><BR>"
		Con.Execute sQuery
		'ContraTransactionNumber field and Update Query added on 21 st Feb 2009 by S.Maheswari
		sQuery = "insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				 "PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
				 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus,"&_
				 "BankInstrumentType,BankInstrumentNo,PayableAt,BankInstrumentDate,DrawnOnBank,ContraTransactionNumber) values"&_
				 "("&iNewTransNo&",'"&iHeadUnit&"','"&sToBookCode&"',"&sToBookNo&",'"&sNewTransType&"',"&_
				 "NULL,"&sAccHeadCode&",'"&sNewAccVouNo&"','"&sNewVouchNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				 ",'"&sPayTo&"','"&sNewVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"',"&_
				 "'"&sInsType&"','"&sInsNo&"','"&sPayableat&"',convert(datetime,'"&sInsDate&"',103),'"&sDrawnOn&"',"&iTransNo&")"

		'Response.Write "Test = "&sQuery &"<BR><BR>"
		Con.Execute sQuery
		'-----------------------------------------------------------------------------------------------------------------------------
		sQuery = "Update Acc_T_VoucherHeader set ContraTransactionNumber  = "&iNewTransNo&" where TransactionNumber = "& iTransNo&" "
		'Response.Write "Update="&sQuery &"<BR><BR>"
		Con.Execute sQuery
		'-----------------------------------------------------------------------------------------------------------------------------

		FOR EACH EntryNode IN Root.childNodes
			IF  EntryNode.nodeName="Entry" THEN
				sEntryno=EntryNode.Attributes.getNamedItem("No").value
				sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
				sAmount=EntryNode.Attributes.getNamedItem("Amount").value
				sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value
				IF CStr(sVouCode) = "01" or CStr(sVouCode) = "02" or CStr(sVouCode) = "08" Then
					iTdsAmount = EntryNode.Attributes.getNamedItem("TdsAmount").value
					iTdsPer = EntryNode.Attributes.getNamedItem("TdsPercentage").value
				Else
					iTdsAmount = 0
					iTdsPer = 0
				End IF
		'---------PROCESS THE CHILD NODES OF ENTRY NODE FOR DETAIL TABLE UPDATION-------
				FOR EACH HeaderNode IN EntryNode.childNodes
					IF HeaderNode.nodeName="AccHead" THEN
							sAccCode=HeaderNode.Attributes.getNamedItem("No").value
							sAccType=HeaderNode.Attributes.getNamedItem("Type").value
					END IF 'End of Check for Account head Node

					IF 	HeaderNode.nodeName="Narration" THEN
							sNarration=Replace(HeaderNode.text,"'","''")
					END IF 'End of Check for Narration Node
				NEXT
			'-------------END OF PROCESSING CHILD NODES OF ENTRY NODE---------------------
			'----------------------------DETAIL TABLE UPDATION-------------------------
				IF StrComp(sAccType,"G")=0 THEN
					sQuery = "insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
					sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
					sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
					sQuery = sQuery& iCrTransNo&",'"&sAccUnit&"'"
					sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
					sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&iTdsAmount&","&iTdsPer&")"

					'Response.Write "<p>"& sQuery& "<BR><BR>"
					Con.Execute sQuery


					sQuery = "insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
					sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
					sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage) values ("
					sQuery = sQuery& iNewTransNo&",'"&sAccUnit&"'"
					sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
					sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&", "&iTdsPer&")"
					'Response.Write "<p>"& sQuery& "<BR><BR>"
					con.execute(sQuery)

					'This will Remove the Extra Entry Entered in GL_Transactions for the Bank/Cash Receipt/Payment entry

					sQuery = "Delete  Acc_t_gltransactions where oudefinitionid = '"&sOrgId&"' "&_
							 "and bookcode = '"&sToBookCode&"' and booknumber = "&sToBookNo&" and transactiontype = '"&sNewTransType&"' "&_
							 "and accounthead = "&sAccCode&"  and TransactionNumber = "&iNewTransNo&" "

					'===========================================================================================================
					Con.Execute sQuery

				END IF
			End IF
		Next 'Entry Node For Details Tabel Insertion

		'If Both From and To Accounthead is marked as summary then the Debit/Credit value
		'will chaged to from account head

		'Response.Clear
		'Response.Write "====================================="
		Dim iRecAff
		sQuery = "Select SummaryPosting From VwOrgGLHeads Where Accounthead IN ("&iFrmAccCode&","&iToAccCode&") "&_
				 "and OUDefinitionID = '"&sOrgId&"' and SummaryPosting = '1' "

		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		iConChkCnt = objRs.RecordCount
		objRs.Close

		IF CStr(iConChkCnt) = "2" Then 'If Both Account heads are summary entry
			sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&dTotal&" Where BookCode = '"&sToBookCode&"' and  "&_
					 "BookNumber = "&sToBookNo&" and TransactionNumber = 0 and TransCrDrIndication = '"&sFrmCrDrTy&"' and  "&_
					 "AccountHead = "&iFrmAccCode&" and Amount - "&dTotal&" >= 0 and Month(VoucherDate) = "&iMon&" and "&_
					 "Year(VoucherDate) = "&iYear&" "

			'Response.Write sQuery &"<br>"
			Con.Execute sQuery,iRecAff


			'Response.Write iRecAff &"<br>"

			IF CStr(iRecAff) <> "0" Then

				'sQuery = "Update Acc_T_GLTransSummaryBreakup Set VoucherAmount = VoucherAmount - "&dTotal&" Where  "&_
				'		 "BookCode = '"&sToBookCode&"' and BookNumber = "&sToBookNo&" and   "&_
				'		 "AccountHead = "&iFrmAccCode&" and VoucherAmount - "&dTotal&" >= 0 and Convert(Char,VoucherDate,103)  "&_
				'		 "= Convert(Char,'"&sVoucDate&"',103) "

				'Response.Write sQuery &"<br>"
				'Con.Execute sQuery,iRecAff
				'Response.Write iRecAff &"<br>"

				'sQuery = "Update Acc_T_GLTransSummaryBreakup Set VoucherAmount = VoucherAmount + "&dTotal&" Where  "&_
				'		 "BookCode = '"&sVouCode&"' and BookNumber = "&iFrmBookNo&" and  "&_
				'		 "AccountHead = "&iFrmAccCode&" and Convert(Char,VoucherDate,103)  "&_
				'		 "= Convert(Char,'"&sVoucDate&"',103) "

				'Response.Write sQuery &"<br>"
				'Con.Execute sQuery,iRecAff
				'Response.Write iRecAff &"<br>"


				'sQuery = "Update Acc_T_GLTransactions Set Amount = Amount + "&dTotal&" Where BookCode = '"&sVouCode&"' and  "&_
				'		 "BookNumber = "&iFrmBookNo&" and TransactionNumber = 0 and TransCrDrIndication = '"&sFrmCrDrTy&"' and "&_
				'		 "AccountHead = "&iToAccCode&" "

				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount + "&dTotal&" Where BookCode = '"&sVouCode&"' and  "&_
						 "BookNumber = "&iFrmBookNo&" and TransactionNumber = 0 and TransCrDrIndication = '"&sFrmCrDrTy&"' and "&_
						 "AccountHead = "&iFrmAccCode&" and Month(VoucherDate) = "&iMon&" and Year(VoucherDate) = "&iYear&" "


				'Response.Write sQuery &"<br>"
				Con.Execute sQuery,iRecAff
				'Response.Write iRecAff &"<br>"
			End IF

			sQuery = "Select TransactionNumber From Acc_T_GLTransSummaryBreakup Where TransactionNumber = "&iTransNo&" "
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iConChkCnt = 1
			Else
				iConChkCnt = 0
			End IF
			objRs.Close

			'Response.Write "<br><br><br>"
			'Response.Write "<p>iFrmTransNo="&iFrmTransNo &"<br><br><br>"

			IF Cstr(iConChkCnt) = "0" Then
				sQuery = "INSERT INTO Acc_T_GLTransSummaryBreakup (OUDefinitionID, AccountHead, BookCode, BookNumber,  "&_
						 "TransactionType, TransactionNumber, VoucherDate, VoucherEntryNumber, VoucherNo,  "&_
						 "VoucherAmount, TransCrDrIndi) "&_
						 "VALUES ('"&sOrgId&"', "&iFrmAccCode&", '"&iFrmBookCode&"', "&iFrmBookNo&",  "&_
						 "'"&sFrmTransTy&"', "&iFrmTransNo&", Convert(datetime,'"&sVoucDate&"',103), 1, "&iFrmTransNo&", "&dTotal&", '"&sFrmCrDrTy&"') "

				'Response.Write sQuery &"<br><br>"
				Con.Execute sQuery

				sQuery = "Delete Acc_T_GLTransSummaryBreakup Where TransactionNumber = "&iNewTransNo&" and VoucherEntryNumber = 1 "
				Con.Execute sQuery

			End IF


		End IF




		objDom.Load server.MapPath("../Temp/Transaction/"&Session.SessionID&"-Voucher_Bank.xml")
		'objDom.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

		FOR EACH EntryNode IN Root.childNodes
			IF  EntryNode.nodeName="Entry" THEN
				sEntryno=EntryNode.Attributes.getNamedItem("No").value
				sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
				sAmount=EntryNode.Attributes.getNamedItem("Amount").value
				sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value

				sExp="//Entry[@No='"&sEntryno&"']/AccHead"
				set tempNode=Root.selectNodes(sExp)
				sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

				set nodANL=objDom.createElement("Root")
				nodANL.setAttribute "TransNo",iTransNo
				nodANL.setAttribute "EntryNo",sEntryno
				nodANL.setAttribute "UnitCode", sAccUnit
				nodANL.setAttribute "GlHead",sAccCode
				nodANL.setAttribute "ACTFlag","C"

				set nodANL2=objDom.createElement("Root")
				nodANL2.setAttribute "TransNo",iTransNo
				nodANL2.setAttribute "EntryNo",sEntryno
				nodANL2.setAttribute "UnitCode", sAccUnit
				nodANL2.setAttribute "GlHead",sAccCode
				nodANL2.setAttribute "ACTFlag","A"

				sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
				set tempNode=Root.selectNodes(sExp)
				if tempNode.length >0 then
					sCheckForProc = "1"
					set HeaderNode=tempNode.item(0).cloneNode(true)
					nodANL.appendChild(HeaderNode)
					nodANL2.appendChild(HeaderNode)
					bAddFlag=true
				end if

				sExp="//Entry[@No='"&sEntryno&"']/Analytical"
				set tempNode=Root.selectNodes(sExp)
				if tempNode.length >0 then
					set HeaderNode=tempNode.item(0).cloneNode(true)
					nodANL.appendChild(HeaderNode)
					nodANL2.appendChild(HeaderNode)

					bAddFlag=true
				end if
				if bAddFlag then

				 ' Set adoConn = Server.CreateObject("ADODB.Connection")
				  ' adoConn.ConnectionString = con
				  ' adoConn.CursorLocation = 3
				  ' adoConn.Open

				   sQuery = "Proc_VouCCANALUpdate"

				   Set adoCmd = Server.CreateObject("ADODB.Command")
				   Set adoCmd.ActiveConnection =con
				   adoCmd.CommandText = sQuery
				   adoCmd.CommandType = 4 'adCmdStoredProc
				   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)
				   'Set adoRS = adoCmd.Execute()
				   adoCmd.Execute()
				 '  adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL2.xml),nodANL2.xml)
				  ' 'Set adoRS = adoCmd.Execute()
				  ' adoCmd.Execute()

				  ' adoConn.close
				  ' adoConn = Nothing
				  ' adoCmd = Nothing

				end if
			end if
			bAddFlag = false
		NEXT

		'======== This is Check if From Acc Head and To Acc Head is Mapped for Contra
		'======== Transactions 1 - 2 and 2 - 1.

		sQuery = "Select OUDefinitionID from Acc_M_ContraEntries where OUDefinitionID = '"&iHeadUnit&"' and FromAccountHead = "&iEntAccHead&" and ToAccountHead = "&iHeadAccHead&" "
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			sRevConChk = "T"
		Else
			sRevConChk = "F"
		End IF
		objRs.Close


	End IF 'Contra Check

	'================== Contra Entry Details Ends Here =====================================================


	'------------------UPDATE RECEIPT NO--------------------------
	'IF (sVouType<>"" or sVouCode<>"08") and bPartyReceiptFlag=true THEN
	'	sQuery="select isnull(ReceiptNo,0)+1 from Acc_M_BookNumberSeries where OUDefinitionID='"&sOrgId&"'"
'
'
	'	objRs.open sQuery,con
	'		iPartyReceiptNo=objRs(0)
	'	objRs.close()
'
	'	sQuery="update Acc_M_BookNumberSeries set  ReceiptNo="&iPartyReceiptNo&"  where OUDefinitionID='"&sOrgId&"'"
'
	'	con.execute(sQuery)
'
	'	sQuery="update Acc_T_VoucherHeader set  ReceiptNo="&iPartyReceiptNo&"  where TransactionNumber="&iTransNo
	'	con.execute(sQuery)
'

	' END IF

	'------------------UPDATE CREATED VOUCHER DETAILS-------------------------
	sQuery="update Acc_T_CreatedVoucherHeader set  CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iCreatedTransNo

	'Response.Clear
	'Response.Write sQuery &"<br><br>"
	con.execute(sQuery)

	Dim sNewNarr1,sTempVar,sAllNarr,sTempNarr

	sExp = "//Entry"
	Set EntryNode = Root.selectNodes(sExp)
	For iConCount = 0 To EntryNode.length - 1
		sNewNarr1 = ""

		'Response.Write iCreatedTransNo
	'================= For Sales Invoices ========================================
		iEntNo = EntryNode.Item(iConCount).Attributes.Item(0).nodeValue
		sExp = "//Entry[@No="&iEntNo&"]/PayRec/Doc[@AdjType=""I""]"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length = 1 Then
			sNewNarr1 ="INV No:"
			IF Cstr(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value) <> "" Then
				sTempVar = Split(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value,":")
				IF UBound(sTempVar) > 0 Then
					sNewNarr1 = sNewNarr1&sTempVar(1)
					sNewNarr1 = Replace(sNewNarr1," Dt",",")
				Else
					sNewNarr1 = sNewNarr1&Mid(sTempVar(0),22,Len(sTempVar(0))-10)
					sNewNarr1 = Replace(sNewNarr1," Dt",",")
				End IF
			Else
				sQuery = "Select Narration From Acc_T_Receivables Where CreatedReceivable = "&TempNode.Item(0).Attributes.getNamedItem("No").Value
				Objrs.Open sQuery,Con
				IF Not Objrs.Eof Then
					sTempNarr = Objrs(0)
				Else
					sTempNarr = ""
				End IF
				Objrs.close
				TempNode.Item(0).Attributes.getNamedItem("InvNo").Value = sTempNarr
				sTempVar = Split(sTempNarr,":")
				IF UBound(sTempVar) > 0 Then
					sNewNarr1 = sNewNarr1&sTempVar(1)
					sNewNarr1 = Replace(sNewNarr1," Dt",",")
				Else
					sNewNarr1 = sNewNarr1&Mid(sTempVar(0),22,Len(sTempVar(0))-10)
					sNewNarr1 = Replace(sNewNarr1," Dt",",")
				End IF

			End IF
		Elseif TempNode.length > 1 Then
			sNewNarr1 ="INV No:"
			For iCtr = 0 To TempNode.length - 1
				sTempVar = Split(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":")
				IF UBound(sTempVar) > 0 Then
					sNewNarr1 = sNewNarr1&" "&Trim(sTempVar(1))
				Else
					sNewNarr1 = sNewNarr1&Mid(sTempVar(0),22,Len(sTempVar(0))-10)
				End IF
			Next
			sNewNarr1 = Replace(sNewNarr1," Dt",",")
		End IF



	'================= For Purchase Invoices ========================================
		sExp = "//Entry[@No="&iEntNo&"]/PayRec/Doc[@AdjType=""PI""]"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length = 1 Then
			sNewNarr1 ="INV No:"
			sTempVar = Split(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value,":")
			IF UBound(sTempVar) > 0 Then
				sNewNarr1 = sNewNarr1&sTempVar(1)
				sNewNarr1 = Replace(sNewNarr1," Dt",",")
			Else
				sNewNarr1 = sNewNarr1&Mid(sTempVar(0),36,Len(sTempVar(0))-10)
				sNewNarr1 = Replace(sNewNarr1," Dt",",")
			End IF
		Elseif TempNode.length > 1 Then
			sNewNarr1 ="INV No:"
			For iCtr = 0 To TempNode.length - 1
				sTempVar = Split(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":")
				IF UBound(sTempVar) > 0 Then
					sNewNarr1 = sNewNarr1&" "&Trim(sTempVar(1))
					sNewNarr1 = Replace(sNewNarr1," Dt",",")
				Else
					sNewNarr1 = sNewNarr1&Mid(sTempVar(0),36,Len(sTempVar(0))-10)
					sNewNarr1 = Replace(sNewNarr1," Dt",",")
				End IF
			Next

		End IF

	'================= For Debit Notes ========================================

		sExp = "//Entry[@No="&iEntNo&"]/PayRec/Doc[@AdjType=""D""]"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length = 1 Then
			sNewNarr1 = sNewNarr1 &"DN No:"
			Response.Write "<p>Data="&TempNode.Item(0).Attributes.getNamedItem("InvNo").Value
			sTempVar = Split(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value,":")
			If UBound(sTempVar) >= 1 Then
				sNewNarr1 = sNewNarr1&sTempVar(1)
			Else
				sNewNarr1 = sNewNarr1&sTempVar(0)
			End IF
			sNewNarr1 = sNewNarr1&","
		Elseif TempNode.length > 1 Then
			sNewNarr1 = sNewNarr1 &"DN No:"
			For iCtr = 0 To TempNode.length - 1
				sTempVar = Split(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":")
				sNewNarr1 = sNewNarr1&" "&Trim(sTempVar(1))
			Next
			sNewNarr1 = sNewNarr1&","
		End IF

	'================= For Credit Notes ========================================

		sExp = "//Entry[@No="&iEntNo&"]/PayRec/Doc[@AdjType=""C""]"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length = 1 Then
			sNewNarr1 = sNewNarr1 &"CN No:"
			'sTempVar = Split(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value,":")
			sNewNarr1 = sNewNarr1&Mid(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value,13)
			sNewNarr1 = sNewNarr1&","
		Elseif TempNode.length > 1 Then
			sNewNarr1 = sNewNarr1 &"DN No:"
			For iCtr = 0 To TempNode.length - 1
				'sTempVar = Split(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":")
				sNewNarr1 = sNewNarr1&" "&Mid(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,13)
			Next
			sNewNarr1 = sNewNarr1&","

		End IF

	'================= For Advance Receipts ========================================
		sExp = "//Entry[@No="&iEntNo&"]/PayRec/Doc[@AdjType=""R""]"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length = 1 Then
			sNewNarr1 = sNewNarr1 &"RCP No:"
			'sTempVar = Split(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value,":")
			sNewNarr1 = sNewNarr1&TempNode.Item(0).Attributes.getNamedItem("No").Value
			sNewNarr1 = sNewNarr1&","
		Elseif TempNode.length > 1 Then
			sNewNarr1 = sNewNarr1 &"RCP No:"
			For iCtr = 0 To TempNode.length - 1
				'sTempVar = Split(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":")
				sNewNarr1 = sNewNarr1&" "&TempNode.Item(iCtr).Attributes.getNamedItem("No").Value
			Next
			sNewNarr1 = sNewNarr1&","
		End IF

	'================= For Advance Payments ========================================
		sExp = "//Entry[@No="&iEntNo&"]/PayRec/Doc[@AdjType=""P""]"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length = 1 Then
			sNewNarr1 = sNewNarr1 &"PAY No:"
			'sTempVar = Split(TempNode.Item(0).Attributes.getNamedItem("InvNo").Value,":")
			sNewNarr1 = sNewNarr1&TempNode.Item(0).Attributes.getNamedItem("No").Value
			sNewNarr1 = sNewNarr1&","
		Elseif TempNode.length > 1 Then
			sNewNarr1 = sNewNarr1 &"PAY No:"
			For iCtr = 0 To TempNode.length - 1
				'sTempVar = Split(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":")
				sNewNarr1 = sNewNarr1&" "&TempNode.Item(iCtr).Attributes.getNamedItem("No").Value
			Next
			sNewNarr1 = sNewNarr1&","
		End IF



		sExp = "//Entry[@No="&iEntNo&"]/Narration"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			sNarration = TempNode.Item(0).Text
		End IF

		IF CStr(sNewNarr1) = "" Then
			sNewNarr = sNarration
		Else
			sNewNarr = sNewNarr1 &"-"&sNarration
		End IF

		'sAllNarr = sAllNarr &","&sNewNarr

		sNewNarr = Replace(sNewNarr,"'","''")


		'Response.Write sNewNarr &"<br>"




		IF CStr(sParCheck) = "Y"  and CStr(iFrmBookCode) = "01" Then
			sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNarration&"' Where TransactionNumber = "&iTransNo&" and VoucherEntryNumber = "&Abs(iEntNo-1)
			Con.Execute sQuery

		Elseif CStr(sParCheck) = "Y"  and CStr(iFrmBookCode) = "02" Then
			sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNewNarr&"' Where TransactionNumber = "&iTransNo&"  and VoucherEntryNumber = "&Abs(iEntNo-1)

			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery

			sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewNarr&"' Where TransactionNumber = "&iTransNo&"  and VoucherEntryNumber = "&iEntNo
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery


			sQuery = "Update Acc_T_VoucherDetails Set VoucherNarration = '"&sNewNarr&"' Where  "&_
					 "TransactionNumber = "&iTransNo&" and VoucherEntryNumber = "&iEntNo

			Con.Execute sQuery
		End IF

	Next

	'Con.RollbackTrans
	'Response.End
	Dim adoCmd
	IF CStr(sRevConChk) = "T" Then 'If The Wise versa Contra Check is True Then Call Stored Proc
		'Set adoConn = Server.CreateObject("ADODB.Connection")
		'adoConn.ConnectionString = con
		'adoConn.CursorLocation = 3
		'adoConn.Open
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection = Con

		sQuery = "GLContraUpdate"
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection =Con
		adoCmd.CommandText = sQuery
		adoCmd.CommandType = 4 'adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@FromTransNo",3,1,3,iTransNo)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ToTransNo",3,1,3,iNewTransNo)
		adoCmd.Execute()
	End IF



	'Response.Clear

	'Con.RollbackTrans
	'Response.End



	Dim sTbGLVal,sTempVal,sDiffVal
	sTbGLVal = CheckTBGL(sOrgId)
	sTempVal = Split(sTbGLVal,":")
	dTBDRAmt = sTempVal(0)
	dTBCRAmt = sTempVal(1)

	dTBDRAmt = FormatNumber(dTBDRAmt,2,,,0)
	dTBCRAmt = FormatNumber(dTBCRAmt,2,,,0)

	IF CStr(Trim(sTempVal(2))) = "T" Then
		sDiffVal = "In Trial Balance "
	Else
		sDiffVal = "In Ledger  "
	End IF

	'Response.Write "Debit Note="&dTBDRAmt
	'Response.Write "Credit Note="&dTBCRAmt

'	Con.RollBackTrans
'	Response.End
	IF CDbl(dTBDRAmt) <> CDbl(dTBCRAmt) Then
		con.RollbackTrans

		sErrChk = "F"
		If Response.Buffer Then Response.Clear
		Response.write "<!DOCTYPE html><html><head><title>Credit Debit Mismatch!</title>"&_
			"<style>"&_
			"html,body{height:100%;margin:0;font-family:Arial,Helvetica,sans-serif;background:#f6f8fb;color:#1f2933;}"&_
			"body{display:flex;align-items:center;justify-content:center;}"&_
			".msg{max-width:620px;margin:24px;padding:30px 36px;border:1px solid #f0b8b8;background:#fff7f7;text-align:center;box-shadow:0 8px 24px rgba(31,41,51,.12);}"&_
			".msg h1{margin:0 0 12px;font-size:24px;color:#b42318;font-weight:bold;}"&_
			".msg p{margin:0 0 18px;font-size:15px;line-height:1.5;color:#4b5563;}"&_
			".amounts{display:inline-block;min-width:320px;text-align:left;font-size:15px;}"&_
			".row{display:flex;justify-content:space-between;gap:28px;padding:7px 0;border-top:1px solid #f2d6d6;}"&_
			".row strong{color:#111827;}"&_
			".row.diff strong{color:#b42318;}"&_
			".where{margin-top:16px;font-weight:bold;color:#7f1d1d;}"&_
			"</style></head><body><div class=""msg"">"&_
			"<h1>Credit Debit Mismatch!</h1>"&_
			"<p>Debit Amount and Credit Amount does not match. Transaction rolled back.</p>"&_
			"<div class=""amounts"">"&_
			"<div class=""row""><span>DR AMOUNT</span><strong>"&dTBDRAmt&"</strong></div>"&_
			"<div class=""row""><span>CR AMOUNT</span><strong>"&dTBCRAmt&"</strong></div>"&_
			"<div class=""row diff""><span>DIFFERENCES</span><strong>"&FormatNumber(Round(CDbl(dTBDRAmt) - CDbl(dTBCRAmt),2),2,,,0)&"</strong></div>"&_
			"</div>"
		IF CStr(sTempVal(2)) = "L" Then
			Response.write  "<div class=""where"">Differences In : Ledger</div>"
		Elseif CStr(sTempVal(2)) = "T" Then
			Response.write  "<div class=""where"">Differences In : Trial Balance</div>"
		End IF
		Response.write "</div></body></html>"
		'Response.Write "<h3>Debit and Credit is Not Matching </h3><br>"
		'Response.Write "<b>Debit Amount  :--> </b>" & dTBDRAmt &"<br>"
		'Response.Write "<b>Credit Amount :--> </b>" & dTBCRAmt &"<br>"
		'Response.Write "<b>Differences   :--> </b>" & sDiffVal &"<br>"
		'Response.Write "<h3> Voucher Not Created " &"<br>"
		Response.End
	Else

		Con.CommitTrans
		sErrChk = "T" 'No Errors
	End IF










	'Con.RollBackTrans
	'Response.End




	if con.Errors.count <>0 THEN
		con.RollbackTrans
		FOR iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) &"<br>"
		NEXT
		'Redirect to Error Handling System
	ELSE

		'sRetVal = GetVouchXML(iTransNo)
		'oDOM.Load server.MapPath(sRetVal)

		'oDOM.Load server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")
		'Root.Attributes.getNamedItem("TransNo").value=iTransNo
		'Root.Attributes.getNamedItem("VoucherNo").value=iVouNo
		'Set Root = oDOM.documentElement

		'Root.setAttribute "CreatedTransNo", iCreatedTransNo
		'Root.setAttribute "CreatedVoucherNo", iCreatedVouNo

		'Set newElem  = oDOM.createAttribute("CreatedTransNo")
		'newElem.value = iCreatedTransNo
		'Root.setAttributeNode(newElem)

		'Set newElem  = oDOM.createAttribute("CreatedVoucherNo")
		'newElem.value = iCreatedVouNo
		'Root.setAttributeNode(newElem)

		'This is Been Blocked

		'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")

		'Response.Clear
		'Response.Write sRedirTy & sVouCode
		'Response.End




		IF CStr(sRedirTy) = "A"  and CStr(sVouCode) = "01" Then
			Response.Redirect "ACCCASHVOUCHERS.ASP"
			Response.End
		Elseif CStr(sRedirTy) = "A"  and CStr(sVouCode) = "02" Then
			Response.Redirect "ACCBANKVOUCHERS.ASP"
			Response.End
		Elseif CStr(sRedirTy) = "A"  and CStr(sVouCode) = "08" Then
			Response.Redirect "ACCGJVOUCHERS.ASP"
			Response.end
		End IF


		'Response.End
		if sVouName="CA" then
			'Response.Redirect ("AccVoucherList.asp?optCriteria=Exist&selUnitId="&sOrgId&"&selVoucher="&iFrmBookCode&"&selBook="&iFrmBookNo)
			Response.Redirect ("CashVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy&"&ACTN="&Session("ACTN"))
		elseif sVouName="BA" then
			Response.Redirect ("BankVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy&"&ACTN="&Session("ACTN"))
		else
			Response.Redirect ("GJVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
		end if

	END IF
	Else
		if sVouName="CA" then
			'Response.Redirect ("AccVoucherList.asp?optCriteria=Exist&selUnitId="&sOrgId&"&selVoucher="&iFrmBookCode&"&selBook="&iFrmBookNo)
			Response.Redirect ("CashVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
		elseif sVouName="BA" then
			Response.Redirect ("BankVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
		else
			Response.Redirect ("GJVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
		end if
	End IF

%>

<%
	Function GetBookReceiptNo(sBookCode,sUnit,iBookNo,sVouType)
		'for Internal Receipt: iBookNo = 0
		'for Contra Entry : iBookNo = -1

		'Response.Clear
		Dim iReciptNo,sRecrs,iRecSerNo,iRecSerCode,sLastYear
		sLastYear = Right(Session("FinPeriod"),4)
		sLastYear = sLastYear&"03"

		Set sRecrs = Server.CreateObject("ADODB.RecordSet")
		If sVouType = "D" Then
			sQuery = "Select DrSeriesNo,DrSeriesCode From Acc_M_BookNumberSeries  " & _
					 "Where BookCode = '"&sBookCode&"' And OUDefinitionID = '"&sUnit&"' " & _
					 "And BookNumber = "&iBookNo&" And FinPeriod = '"&Session("FinPeriod")&"' " 'ibookno is added on 28th jan 09 to use same function for both internal receipt and contra entry
		Else
			sQuery = "Select CrSeriesNo,CrSeriesCode From Acc_M_BookNumberSeries  " & _
								 "Where BookCode = '"&sBookCode&"' And OUDefinitionID = '"&sUnit&"' " & _
					 "And BookNumber = "&iBookNo&" And FinPeriod = '"&Session("FinPeriod")&"' "
		End If
		Response.Write sQuery &"<br><br><br>"

		sRecrs.Open sQuery,Con
		IF Not sRecrs.EOF Then
			iRecSerNo = sRecrs(0)
			iRecSerCode = sRecrs(1)
		End IF
		sRecrs.Close

		sQuery = "Select Number,isNull(Prefix,''),isNull(Suffix,'') From APP_R_NoSeriesModuleEntry Where SeriesNo = "&iRecSerNo&" and " & _
				 "SeriesCode = "&iRecSerCode&" And OUDefinitionID = '"&sUnit&"' " & _
				 "And Period = '"&sLastYear&"' "

		Response.Write sQuery &"<br><br><br>"
		sRecrs.Open sQuery,Con
		IF Not sRecrs.EOF Then
			'Response.Write sRecrs("Number") &"<br>"
			iReciptNo = sRecrs(1) & sRecrs("Number") & sRecrs(2)
		End IF
		sRecrs.Close

		sQuery = "Update APP_R_NoSeriesModuleEntry Set Number = Number + 1 Where  " & _
				 "SeriesNo = "&iRecSerNo&" And SeriesCode = "&iRecSerCode&" And OUDefinitionID = '"&sUnit&"' " & _
				 "And Period = '"&sLastYear&"' "

		Con.Execute sQuery

		IF CStr(iReciptNo) = "" Then
			iReciptNo = 1
		End IF

		GetBookReceiptNo = iReciptNo
		'Response.Clear
		Response.Write "Receipt Number is " & iReciptNo &"<br><br><br>"
		'Response.End
	End Function
%>
