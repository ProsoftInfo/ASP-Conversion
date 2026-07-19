<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 28, 2002
	'Modified BY				:	RAGAVENDRAN R
	'Modified On				:   FEB 16,2010
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/CommonFunctions.asp"-->
<%

'XML DOM Variables

Dim oDOM,nodHeader,Root,objRs,sQuery,sExp,objAttr
dim EntryNode,HeaderNode,nodANL,newElem,tempNode,bAddFlag,sInsChk
Dim iSaletransno,sAccTransno,sCommissiontype,sCurrcode,iCommissionvalue,sAgentcode
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp,sGrpId
dim sOrgId,sBookNo,sVouType,sVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,sPayTo,iTdsAmount,iTdsPer
dim iSeriesNo,iSeriesCode,iAdvNo,sApprover
dim sAccHeadCode,sSelVouDate,sCrnoTrno,iOldCrNo
Dim AdvNode,iCtr,sCallType,sCodeTy,iTdsGroupID
Dim iTdsValue,iTDSEntryAmt,iTDSEntryPer,iPayRecAmt,sFormula,sTDSNarr,TDSFlag,sVouEntryno,iAccHeadNo
Dim BType,sTdsRndOff,sCallFromFD,sAction
Dim objDOM,domroot,pagenode,sCallFrom,sVoucherXmlPath,iCounter
'=============Variables used for contra check
Dim iHeadAccHead,iEntAccHead,iConCount,sHdFlag
Dim sHeadAccName,sEntAccHeadName,iHeadUnit,iCrTransNo
Dim iHeadUnitName,AccNode,NarrNode,sContraCheck
'=================================================================


BType	 = Request("hVouType")
sVouCode = Request("hVouCode")
sVouName = Request("hVouName")
sApprove = Request("optApprove")
sInsChk = Request("hInsVou")
sApprover = Request.Form("selUserId")
sSelVouDate = Request.Form("hSelVouDate")
iOldCrNo = Request.Form("hTransNo")
sCallType = Request("CallType")
sVouNo = Session("CrVouNo")
iTdsGroupID = Request.Form("SelTDSGrp")
'Session("CrVouNo") = ""
sCallFrom = Request.Form("hInvCallFrom")
'Response.Write sVouNo
sAction = Session("ACTN")
IF CStr(iTdsGroupID) = "" Then
	iTdsGroupID = 0
End IF


IF CStr(sApprover) = "" Then
	sApprover = getUserID()
End IF

bAddFlag=false

IF CStr(sApprove) = "Y" Then
	sVouStatus="010101" 'Crearted For Approval
Else
	sVouStatus="010103" 'Crearted For Accounting
End IF

set objRs  = server.CreateObject("adodb.recordset")

If Request.Form.Count = 0 then
	BType	 = Request.QueryString("hVouType")
	sVouCode = Request.QueryString("hVouCode")
	sVouName = Request.QueryString("hVouName")
	sApprove = Request.QueryString("optApprove")
	sInsChk = Request.QueryString("hInsVou")
	sApprover = Request.QueryString("selUserId")
	sSelVouDate = Request.QueryString("hSelVouDate")
	iOldCrNo = Request.QueryString("hTransNo")
	sCallType = Request.QueryString("CallType")
	iTdsGroupID = Request.QueryString("SelTDSGrp")
	sCallFrom = Request.QueryString("hInvCallFrom")
	sCallFromFD = Request.QueryString("FromFD")
	IF CStr(iTdsGroupID) = "" Then
		iTdsGroupID = 0
	End IF


	IF CStr(sApprover) = "" Then
		sApprover = getUserID()
	End IF


end if
'Response.Write "<p> " &  Request.QueryString("hVouCode")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sVoucherXmlPath = server.MapPath("../temp/transaction/Voucher Entry_"&sVouName&"_"&Session.SessionID&".xml")
If Not oDOM.Load(sVoucherXmlPath) Then
	If Response.Buffer Then Response.Clear
	Response.Write "Unable to load voucher XML file for voucher save. Please add at least one entry and try Save again."
	If oDOM.parseError.errorCode <> 0 Then
		Response.Write "<br>" & Server.HTMLEncode(oDOM.parseError.reason)
	End If
	Response.End
End If

set Root=oDOM.documentElement
If Root Is Nothing Then
	If Response.Buffer Then Response.Clear
	Response.Write "Unable to read voucher XML data for voucher save. Please add at least one entry and try Save again."
	Response.End
End If

sOrgId=Root.Attributes.getNamedItem("UnitNo").value
sBookNo =Root.Attributes.getNamedItem("BookNo").value
sVouType=Root.Attributes.getNamedItem("CRDR").value
sVoucDate=Root.Attributes.getNamedItem("VouDate").value
sAccHeadCode=Root.Attributes.getNamedItem("BookAcchead").value

Root.Attributes.getNamedItem("Approver").value=sApprove

FOR EACH EntryNode IN Root.childNodes
	IF EntryNode.nodeName = "Entry" Then
		sPayTo=Replace(EntryNode.Attributes.getNamedItem("Payto").value,"'","''")
		IF EntryNode.Attributes.getNamedItem("CRDR").value="C" THEN
			dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.getNamedItem("Amount").value)
		ELSE
			dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.getNamedItem("Amount").value)
		END IF
	End IF
NEXT
'Response.Write dCRAmt & ">>>"&dDRAmt &sVouType&"<BR>"
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

iTdsAmount = 0
Dim sTempVouNo,sTemparr,sTemp2
sExp = "//TDS"
Set tempNode = Root.selectNodes(sExp)
IF tempNode.length <> 0 Then
	For iCtr = 0 To tempNode.length - 1
		iTdsAmount = iTdsAmount + CDbl(tempNode.Item(iCtr).Attributes.getNamedItem("PayRecAmount").Value)
	Next
Else
	iTdsAmount = 0
End IF

dTotal = CDbl(dTotal) - CDbl(iTdsAmount)

con.BeginTrans

IF CStr(sInsChk) <> "Y" Then
	'===================Contra check============================

		'Contra check routine
		iHeadAccHead = Root.Attributes.getNamedItem("BookAcchead").value
		iHeadUnit = Root.Attributes.getNamedItem("UnitNo").value
		iHeadUnitName = Root.Attributes.getNamedItem("UnitName").value

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
							'Response.Write sQuery
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
							'Response.Write "<p>CountValue="& iConCount
							'Response.End
							IF CInt(iConCount) <> 0 Then
								sContraCheck = "T"
							End IF 'Contra Count Check

						End IF 'GL Account Head Check
					End IF 'Accout Head Type Check
				Next 'Same Unit Loop
			Next
		End IF
		If trim(sContraCheck) = "T"  Then
			sVouNo = GetBookReceiptNo(sVouCode,sOrgId,-1,sVouType) 'Added on 28th Jan 09 for Contra entry
		Else
			sTemparr = Split(sBookNo,"-")
			sBookNo = sTemparr(0)

			IF strcomp(sVouType,"D") = 0 THEN
				sQuery = "select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
						 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
			ELSE
				sQuery = "select CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
						 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
			END IF
		    Response.write sQuery

			with objRs
				.ActiveConnection = con
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.Open
			end with

			If Not objRs.EOF then
				iSeriesNo=objRs(0)
				iSeriesCode=objRs(1)
			End IF
			objRs.close()
			'con.BeginTrans

			if trim(iSeriesNo) ="" or IsNull(iSeriesNo) then
			    Response.clear
			    Dim sBookName
			    sBookName = GetAccBookName(sVouCode,sBookNo)
			    Response.write "<h1>Number Series is not created for <font color=red>"& sBookName &"</font></h1>"
			    Response.end
			end if
			IF CStr(sVouNo) = "" Then
				sVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
			End IF

			sCrnoTrno = ""
		End If 'End contra check if

Else
	'con.BeginTrans
	sVouNo = Trim(Request.Form("txtVouNo"))
	sCrnoTrno = sVouNo&":"&iOldCrNo

	sQuery = "Select CreatedVoucherNo From Acc_T_CreatedVoucherHeader Where OUDefinitionID "&_
			 "= '"&sOrgId&"' and BookCode = '"&sVouCode&"' and BookNumber = "&sBookNo&" and  "&_
			 "CreatedVoucherNo Like '"&sVouNo&"%' and  "&_
			 "Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sSelVouDate&"',103) "
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
	sVouNo = sTempVouNo
End IF

'Response.Write dTotal
'Response.End

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iTransNo=objRs(0)
objRs.Close

'Response.Write "dTotal="&dTotal

IF sVouType="" or sVouCode="08" THEN
	sQuery="insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
	"PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
	"CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,TDSGroupID) values"&_
	"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
	"NULL,NULL,'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal &_
	",NULL," & sApprover & ",getdate(),NULL,'"& sVouStatus & "',"&iTdsGroupID&")"
ELSE
sQuery="insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
	"PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
	"PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,OtherApplnTableName,TDSGroupID) values"&_
	"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
	"NULL,"&sAccHeadCode&",'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
	",'"&sPayTo&"','"&sVouType&"',"&sApprover&",getdate(),NULL,'"&sVouStatus&"','"&sCrnoTrno&"',"&iTdsGroupID&")"

END IF

'Response.Write sQuery& "<BR>"
'Response.End
con.execute(sQuery)
'Con.RollbackTrans
'Response.End
iTdsAmount = 0
iTdsPer = 0
Dim iTdsEntryType
'-----------------------PROCESS ENTRY NODES-------------------------------
FOR EACH EntryNode IN Root.childNodes
	IF  EntryNode.nodeName="Entry" THEN
		'Response.Write "====Entry====="
		sEntryno=EntryNode.Attributes.getNamedItem("No").value
		sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
		sAmount=EntryNode.Attributes.getNamedItem("Amount").value
		sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value
		Response.Write "sEntryno = "& sEntryno

		IF CStr(sVouCode) <> "08" Then
			Set objAttr = EntryNode.Attributes.getNamedItem("GroupId")
			If objAttr Is Nothing Then
				sGrpId = iTdsGroupID
				If CStr(sGrpId) = "" Then sGrpId = "0"
			Else
				sGrpId = objAttr.value
			End If

			IF CStr(sEntryType) = "D" Then
				iTdsEntryType = "C"
			Else
				iTdsEntryType = "D"
			End IF

			IF CStr(sVouCode) = "01" or CStr(sVouCode) = "02" or CStr(sVouCode) = "08" Then
				Set objAttr = EntryNode.Attributes.getNamedItem("TdsAmount")
				If objAttr Is Nothing Then
					iTdsAmount = 0
				Else
					iTdsAmount = objAttr.value
				End If
				Set objAttr = EntryNode.Attributes.getNamedItem("TdsPercentage")
				If objAttr Is Nothing Then
					iTdsPer = 0
				Else
					iTdsPer = objAttr.value
				End If
			Else
				iTdsAmount = 0
				iTdsPer = 0
			End IF
			IF Cstr(iTdsAmount) = "" Then
				iTdsAmount = 0
				iTdsPer = 0
			End if
			'Response.Write   "<BR>****"& sAmount &"-"& iTdsAmount & "***<BR>"
			'sAmount = Cdbl(sAmount) - Cdbl(iTdsAmount)
			'Response.Write " sAmount = "& sAmount &"<BR><BR>"
			sTDSNarr = "TDS Entry"
			TDSFlag = "Y"
		Else
			sTDSNarr = ""
			TDSFlag = "N"
			iTdsAmount = 0
			iTdsPer = 0
			sGrpID = "0"
		End IF


	'---------PROCESS THE CHILD NODES OF ENTRY NODE FOR DETAIL TABLE UPDATION----
		sVouEntryno = sEntryno + 1
		FOR EACH HeaderNode IN EntryNode.childNodes
			IF HeaderNode.nodeName="AccHead" THEN
				sAccCode = HeaderNode.Attributes.getNamedItem("No").value
				sAccType = HeaderNode.Attributes.getNamedItem("Type").value

			END IF 'End of Check for Account head Node
			'Added by Maheshwari on Mar 2nd 2007 for TDS Entries
			IF HeaderNode.nodeName="TDS" THEN
				'Response.Write HeaderNode.nodeName &"<BR>"
				iAccHeadNo	 = HeaderNode.getAttribute("AccHeadCode")
				iTDSEntryAmt = HeaderNode.getAttribute("TDSAmount")
				iTDSEntryPer = HeaderNode.getAttribute("TdsPercentage")
				iPayRecAmt	 = HeaderNode.getAttribute("PayRecAmount")
				sFormula	 = HeaderNode.getAttribute("Formula")
				sTdsRndOff   = HeaderNode.getAttribute("TdsRndOff")

				'Response.Write "iTDSEntryPer="&iTDSEntryPer &"<BR>"
				sQuery =" Insert into Acc_T_CreatedVoucherDetails(CreatedTransNo,AccountingUnit,VoucherEntryNumber,"&_
						" AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"&_
						" VoucherNarration,Amount,TransCrDrIndication,TDSonAmount,PackCode,TDSFlag,TDSPercentage,TDSRoundOff) "&_
						" values("&iTransNo&",'"&sAccUnit&"',"&sVouEntryno&","&iAccHeadNo&",NULL,NULL,NULL, "&_
						" "&Pack(sTDSNarr)&","&iPayRecAmt&",'"&iTdsEntryType&"',"&iTDSEntryAmt&","&sGrpId&",'"&TDSFlag&"',"&iTDSEntryPer&",'"&sTdsRndOff&"')"
						'Response.Write "<BR>TDS="& sQuery & "<BR><BR>"
						con.execute(sQuery)
				sVouEntryno = sVouEntryno + 1
			END IF 'IF HeaderNode.nodeName="TDS" THEN


			IF 	HeaderNode.nodeName="Narration" THEN
					sNarration=Replace(HeaderNode.text,"'","''")
			END IF 'End of Check for Narration Node
		NEXT
	'-------------END OF PROCESSING CHILD NODES OF ENTRY NODE---------------------
	'----------------------------DETAIL TABLE UPDATION-------------------------
		IF StrComp(sAccType,"G")=0 THEN
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSRoundOff) values ("
			sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" "&Pack(sNarration)&","&sAmount&",'"&sEntryType&"',"&iTdsAmount&",'"&sTdsRndOff&"')"
		ELSE
			sTemp=Split(sAccCode,"?")
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication, TDSonAmount,TDSRoundOff) values ("
			sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
			sQuery=sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
			sQuery=sQuery&" "&Pack(sNarration)&","&sAmount&",'"&sEntryType&"',"&iTdsAmount&",'"&sTdsRndOff&"')"
		END IF
		Response.Write sQuery& "<BR>"
		con.execute(sQuery)
	'-----------------------END OF DETAIL TABLE UPDATION----------------------

	DIM sCCGroup,sAddCode,sAddGroupCode,sAddRatio,sAddAmount,dAddTotal,sAdjTy,iCrAdvNo
	Dim dAddAmt,dSubAmt
	dAddTotal=0
	'Response.Clear
	'--------PROCESS CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION----
		FOR EACH HeaderNode IN EntryNode.childNodes
	'----------------------PROCESS PAYABLE / RECEVIABLE NODES ----------------
			IF 	HeaderNode.nodeName="PayRec" THEN
				FOR EACH  nodANL IN HeaderNode.childNodes
					sAddCode=nodANL.Attributes.getNamedItem("No").value
					Response.Write sAddCode &" ============== <br><br>"
					sAddAmount=nodANL.Attributes.getNamedItem("AmtToAdjust").value
					sDocType=nodANL.Attributes.getNamedItem("DocType").value
					sAdjTy = nodANL.Attributes.getNamedItem("AdjType").value
					iAdvNo = nodANL.Attributes.getNamedItem("PayableNo").value

					sCodeTy = ""
					'=====================================================
					IF CStr(Trim(sAdjTy)) = "D" Then
					    'sQuery = "Select DrCreatedReceivable,isNull(CreatedReceivable,0) From Acc_T_Receivables Where ReceivableNumber = "&sAddCode&" "
						sQuery = "Select DrCreatedReceivable,isNull(CreatedReceivable,0) From Acc_T_Receivables Where CreatedReceivable =  "& sAddCode
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							sAddCode = objRs(0)
							if Trim(sAddCode)="" or Trim(sAddCode)="0" then
							    sAddCode = objRs(1)
							end if
							sCodeTy = "R"
						End IF
						objRs.Close
					Elseif CStr(Trim(sAdjTy)) = "C" Then
						'sQuery = "Select CRCreatedPayable,isNull(CreatedPayablesNumber,0) From Acc_T_Payables Where PayablesNumber = "&sAddCode&" "
						sQuery = "Select CRCreatedPayable,isNull(CreatedPayablesNumber,0) From Acc_T_Payables Where CrCreatedPayable = "&sAddCode&" "
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							sAddCode = objRs(0)
							if Trim(sAddCode)="" or Trim(sAddCode)="0" then
							    sAddCode = objRs(1)
							end if
							sCodeTy = "P"
						End IF
						objRs.Close
					End IF
					'======================================================
					IF CStr(sAccType) = "P" and CStr(Left(sAccCode,2)) = "CR" Then
						'dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)
						Select Case CStr(sAdjTy)
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

						dAddTotal=CDbl(dAddAmt)-CDbl(dSubAmt)

					ElseIF CStr(sAccType) = "P" and CStr(Left(sAccCode,2)) = "DR" Then
						Select Case CStr(sAdjTy)
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

					Response.Write sCodeTy &"<br><br>"

					IF CStr(sAdjTy) <> "P" and CStr(sAdjTy) <> "R" Then
						IF CDbl(sAddAmount) > 0 THEN
							IF sDocType="C" and CStr(sCodeTy) = "" THEN
							'-----------------------UPDATE PAYABLE TABLE------------------------------
								sQuery = "Select PayablesNumber From Acc_T_CreatedPayables Where PayablesNumber = "&sAddCode&" "&_
										 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "

								Response.Write sQuery &"<br><br>"
								objRs.Open sQuery,con
								IF Not objRs.EOF Then
									sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
											 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								Else
									sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
											 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								End IF
								objRs.Close

								Response.Write sQuery &"<br><br>"
								con.execute(sQuery)
							ELSEIF sDocType="D"  and CStr(sCodeTy) = "" THEN
							'-----------------------UPDATE RECEIVABLE TABLE---------------------------
								sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where ReceivableNumber = "&sAddCode&" "&_
										 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "

								Response.Write sQuery &"<br><br>"
								objRs.Open sQuery,con
								IF Not objRs.EOF Then
									sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
											 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								Else
									sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
											 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								End IF
								objRs.Close

								Response.Write sQuery &"<br><br>"
								con.execute(sQuery)
							Elseif CStr(sCodeTy) = "R" Then
								sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where ReceivableNumber = "&sAddCode&" "&_
										 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "

								Response.Write sQuery &"<br>===============================<br>"
								objRs.Open sQuery,con
								IF Not objRs.EOF Then
									sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
											 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								Else
									sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
											 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								End IF
								objRs.Close
								Response.Write sQuery &"<br><br>"
								con.execute(sQuery)


							Elseif CStr(sCodeTy) = "P" Then
								sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where ReceivableNumber = "&sAddCode&" "&_
										 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "
										 Response.Write sQuery &"<br><br>"
								objRs.Open sQuery,con
								IF Not objRs.EOF Then
									sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
											 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								Else
									sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
											 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								End IF
								objRs.Close

								Response.Write sQuery &"<br><br>"
								con.execute(sQuery)

							END IF
						END IF
					ElseIF CStr(sAdjTy) = "P" Then
						sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo&" "
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							iCrAdvNo = objRs(0)
						End IF
						objRs.Close

						sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted = isNull(AdvanceAdjusted,0) + "&sAddAmount&" WHERE  "&_
								 "CreatedAdvanceNo = "&iCrAdvNo&" AND PartyType = '"&sTemp(0)&"' AND PartySubType = "&sTemp(1)&"  "&_
								 "AND PartyCode = "&sTemp(3)&" "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

						sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, "&_
								 "PaidOn, AmountPaid, AdjustType,VoucherEntryNumber) "&_
								 "VALUES ("&iCrAdvNo&", "&iTransNo&", Convert(datetime,getDate(),103), "&sAddAmount&", 'A',"&sEntryno&") "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

					ElseIF CStr(sAdjTy) = "R" Then
						sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo&" "
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							iCrAdvNo = objRs(0)
						End IF
						objRs.Close

						sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted = isNull(AdvanceAdjusted,0) + "&sAddAmount&" WHERE  "&_
								 "CreatedAdvanceNo = "&iCrAdvNo&" AND PartyType = '"&sTemp(0)&"' AND PartySubType = "&sTemp(1)&"  "&_
								 "AND PartyCode = "&sTemp(3)&" "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

						sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, "&_
								 "ReceivedOn, AmountReceived, AdjustType,VoucherEntryNumber) "&_
								 "VALUES ("&iCrAdvNo&", "&iTransNo&", Convert(datetime,getDate(),103), "&sAddAmount&", 'A',"&sEntryno&") "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

					End IF
				NEXT
			END IF
	'-------------END OF PROCESSING PAYABLE / RECEVIABLE NODES ---------------
		NEXT
	'-END OF PROCESSING CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION-
	'-------------PROCESS FOR ADVANCE TABLE UPDATION -------------------------
		IF sAccType="P" AND CDbl(sAmount)> CDbl(dAddTotal) THEN
			IF sVouType="C" THEN
				sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
				objRs.Open sQuery,Con
				If Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close

				sExp = "//Entry[@No="&sEntryno&"]/AccHead"
				Set AdvNode = Root.selectNodes(sExp)
				IF AdvNode.length <> 0 Then
					Set newElem = oDOM.createAttribute("AdvNo")
					newElem.value = iAdvNo
					AdvNode.Item(0).setAttributeNode(newElem)
				End IF

				IF Cdbl(dAddAmt) <> 0 Then
					sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
						"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
						" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
						""&sTemp(3)&","&sAmount&","&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,NULL)"
						Response.Write sQuery
						con.execute(sQuery)
				Else
					sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
						"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
						" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
						""&sTemp(3)&","&sAmount&","&CDbl(sAmount)&",NULL,NULL)"
						Response.Write sQuery
						con.execute(sQuery)

				End IF
			ELSEIF sVouType="D" THEN
				sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
				objRs.Open sQuery,Con
				If Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close

				sExp = "//Entry[@No="&sEntryno&"]/AccHead"
				Set AdvNode = Root.selectNodes(sExp)
				IF AdvNode.length <> 0 Then
					Set newElem = oDOM.createAttribute("AdvNo")
					newElem.value = iAdvNo
					AdvNode.Item(0).setAttributeNode(newElem)
				End IF

				IF Cdbl(dAddAmt) <> 0 Then
					sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
							"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
							" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
							""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)- CDbl(dAddTotal)&",NULL)"
							'Response.Write sQuery
							con.execute(sQuery)
				Else
					sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
						"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
						" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
						""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)&",NULL)"
						'Response.Write sQuery
					con.execute(sQuery)
				End IF

			END IF
		END IF
	'-------------END OF PROCESSING FOR ADVANCE TABLE UPDATION ---------------
	END IF
	'-------------END OF ENTRY NODE CHECK ---------------
NEXT

IF CStr(sApprove) = "Y" Then
	sQuery = "insert into Acc_T_VouchersForApproval(CreatedTransNo,ApprovalLevel,ToBeApprovedBy)"&_
			 "Values("&iTransNo&",1,"&sApprover&")"
	con.execute(sQuery)
End IF

'IF the Voucher is only of the Bank then the following will get updated.

Dim sInsNo,sInsDate,sInsType,sInsPayat,sInsDrawat,sOption,iEntNo,iInsEntNo,i,iSlNo,iBkInsAmt
IF CStr(sVouCode) = "02" Then
	sExp = "//BankInstrumentDet"
	Set tempNode = Root.selectNodes(sExp)
	 'Response.Write tempNode.length
	IF tempNode.length <> 0 Then
		For i = 0 to tempNode.length -1
			iSlNo    = tempNode.Item(i).Attributes.getNamedItem("SlNo").value
			sInsType = tempNode.Item(i).Attributes.getNamedItem("InsType").value
			sInsNo = tempNode.Item(i).Attributes.getNamedItem("InsNo").value
			sInsDate = tempNode.Item(i).Attributes.getNamedItem("InsDate").value
			sInsPayat = tempNode.Item(i).Attributes.getNamedItem("PayAt").value
			sInsDrawat = tempNode.Item(i).Attributes.getNamedItem("DrawnOn").value
			iBkInsAmt = tempNode.Item(i).Attributes.getNamedItem("InsAmt").value
			sOption    = tempNode.Item(i).Attributes.getNamedItem("Option").value
			'Response.Write BType &"**"&sOption

			IF trim(BType) = "P" then
				IF trim(sOption) = "Y" then
				 	sTemp = split(sInsNo,"-")
				 	iEntNo		= sTemp(0)
				 	iInsEntNo	= stemp(1)
				 	sInsNo		= sTemp(2)

				 	sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
				 			 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,BankInstrumentEntryNo,InstrumentEntryNo1,InstrumentAmount)"&_
				 			 " Values ("&iTransNo&","&iSlNo&",'"&sInsType&"','"&sInsNo&"',convert(datetime,'"&sInsDate&"',103),"&_
				 			 "'"&sInsPayat&"','"&sInsDrawat&"',"&iEntNo&","&iInsEntNo&","&iBkInsAmt&" ) "
				 	Response.Write "1="&sQuery &"<BR><BR>"
					con.execute(sQuery)
				 	'sQuery	= "update Acc_T_CreatedVoucherHeader set CreatedVouchStatus='"&sVouStatus&"'," &_
					'		  "BankInstrumentType='"&sInsType&"',BankInstrumentNo='"&sInsNo&"',PayableAt='"&sInsPayat&"',"&_
					'		  "BankInstrumentDate=convert(datetime,'"&sInsDate&"',103),DrawnOnBank='"&sInsDrawat&"',"&_
					'		  "BankInstrumentEntryNo="&iEntNo&",InstrumentEntryNo="&iInsEntNo&" where CreatedTransNo="&iTransNo
					'Response.Write sQuery &"<BR><BR>"
					'con.execute(sQuery)
					sQuery = "Update Acc_R_BankInstrumentUsage set CreatedTransNo  = "&iTransNo&",Status = 'U' where EntryNo = "&iEntNo&" "&_
							" and InstrumentEntryNo = "&iInsEntNo&" and InstrumentNo = '"&sInsNo&"' "
					Response.Write sQuery&"<BR><BR>"
					con.execute(sQuery)
				Else
					sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
				 			 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,InstrumentAmount)"&_
				 			 " Values ("&iTransNo&","&iSlNo&",'"&sInsType&"','"&sInsNo&"',convert(datetime,'"&sInsDate&"',103),"&_
				 			 "'"&sInsPayat&"','"&sInsDrawat&"',"&iBkInsAmt&") "
				 	Response.Write "<BR>2= "&sQuery &"<BR><BR>"
					con.execute(sQuery)

					'sQuery	= "update Acc_T_CreatedVoucherHeader set CreatedVouchStatus='"&sVouStatus&"'," &_
					'		  "BankInstrumentType='"&sInsType&"',BankInstrumentNo='"&sInsNo&"',PayableAt='"&sInsPayat&"',"&_
					'		  "BankInstrumentDate=convert(datetime,'"&sInsDate&"',103),DrawnOnBank='"&sInsDrawat&"'"&_
					'		  " where CreatedTransNo="&iTransNo
					'Response.Write sQuery &"<BR><BR>"
					'con.execute(sQuery)

				End IF
			ElseIF trim(BType) = "R" then

				sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
						 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,InstrumentAmount)"&_
						 " Values ("&iTransNo&","&iSlNo&",'"&sInsType&"','"&sInsNo&"',convert(datetime,'"&sInsDate&"',103),"&_
						 "'"&sInsPayat&"','"&sInsDrawat&"',"&iBkInsAmt&") "
				'Response.Write "<BR> "&sQuery &"<BR><BR>"
				con.execute(sQuery)
			End IF
		Next
	End IF

End IF


'Con.RollbackTrans
'Response.End
'------------------END OF PROCESSING ENTRY NODES--------------------------
Dim iPayTot,iRecTot,iAdvTot,sErrCheck,iAdjAdvTot
iRecTot = 0
iPayTot = 0
iAdvTot = 0
iAdjAdvTot = 0
sErrCheck = "N" 'No Errors found
'*************** Start of Checking adjusted Bill is Not Been Adjusted By Some Other **********************
sQuery = "Select Count(1) From Acc_T_CreatedPybleAdjDet WHere  "&_
		 "CreatedTransNo = "&iTransNo&" and AdjustType is NULL"
objRs.Open sQuery,Con
If Not objRs.EOF Then
	iPayTot = objRs(0)
End if
objRs.Close

sQuery = "Select Count(1) From Acc_T_CreatedRcvbleAdjDet WHere  "&_
		 "CreatedTransNo = "&iTransNo&" and AdjustType is NULL "
objRs.Open sQuery,Con
If Not objRs.EOF Then
	iRecTot = objRs(0)
End if
objRs.Close

sQuery = "Select Count(1) From Acc_T_CreatedPybleAdjDet WHere  "&_
		 "CreatedTransNo = "&iTransNo&" and AdjustType is NOT NULL"
objRs.Open sQuery,Con
If Not objRs.EOF Then
	iAdvTot = objRs(0)
End if
objRs.Close

sQuery = "Select Count(1) From Acc_T_CreatedRcvbleAdjDet WHere  "&_
		 "CreatedTransNo = "&iTransNo&" and AdjustType is NOT NULL "
objRs.Open sQuery,Con
If Not objRs.EOF Then
	iAdvTot = iAdvTot + objRs(0)
End if
objRs.Close

IF CDbl(iAdvTot + iPayTot + iRecTot) <> 0 Then 'atleast one adjustment is made.
	IF CStr(iPayTot) <> "0" Then
		sQuery = "Select Count(1) From Acc_T_CreatedPybleAdjDet C,Acc_T_CreatedPayables P "&_
				 "Where P.PayablesNumber = C.PayablesNumber and C.CreatedTransno = "&iTransNo&" and  "&_
				 "P.AmountPayable >= P.AmountPaid and C.AdjustType is NULL "
		objRs.Open sQuery,con
		IF Not objRs.EOF and CStr(iPayTot) <> CStr(objRs(0)) Then
			sErrCheck = "Y" 'Error Found
		End if
		objRs.Close
	End IF

	IF CStr(iRecTot) <> "0" Then
		sQuery = "Select Count(1) From Acc_T_CreatedRcvbleAdjDet C,Acc_T_CreatedReceivables P "&_
				 "Where P.ReceivableNumber = C.ReceivableNumber and C.CreatedTransno = "&iTransNo&" and  "&_
				 "P.AmountReceivable >= P.AmountReceived and C.AdjustType is NULL "
		Response.write sQuery
		objRs.Open sQuery,con
		IF Not objRs.EOF and CStr(iRecTot) <> CStr(objRs(0)) Then
			sErrCheck = "Y" 'Error Found
		End if
		objRs.Close
	End IF

	IF CStr(iAdvTot) <> "0" Then
		sQuery = "Select Count(1) From Acc_T_CreatedRcvbleAdjDet C,Acc_T_CreatedAdvances P "&_
				 "Where C.ReceivableNumber = P.CreatedAdvanceNo and C.CreatedTransno = "&iTransNo&" and  "&_
				 "P.ActualVoucherAmount >= P.AdvanceAdjusted and C.AdjustType is NOT NULL "
		objRs.Open sQuery,con
		IF Not objRs.EOF Then
			iAdjAdvTot = objRs(0)
		End if
		objRs.Close
	End IF

	IF CStr(iAdvTot) <> "0" Then
		sQuery = "Select Count(1) From Acc_T_CreatedPybleAdjDet C,Acc_T_CreatedAdvances P "&_
				 "Where C.PayablesNumber = P.CreatedAdvanceNo and C.CreatedTransno = "&iTransNo&" and  "&_
				 "P.ActualVoucherAmount >= P.AdvanceAdjusted and C.AdjustType is NOT NULL "
		objRs.Open sQuery,con
		IF Not objRs.EOF Then
			iAdjAdvTot = Cdbl(iAdjAdvTot) + Cdbl(objRs(0))
		End if
		objRs.Close
	End IF

	IF CStr(iAdjAdvTot) <> CStr(iAdvTot) Then
		sErrCheck = "Y"
	End If

'	IF CStr(sErrCheck) = "Y" Then 'Error is Found
'		'Con.CommitTrans
'		Con.RollbackTrans
'		Response.Redirect "AccErrDisp.asp?BookCode="&sVouCode
'		Response.End
'	End IF

End IF
'**************** End of To Checking adjusted Bill is Not Been Adjusted By Some Other **********************

if sCallFrom = "SC" then

	Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDOM.async = false
		objDOM.load(server.MapPath("../temp/transaction/VoucherSalCommDet_"&Session.SessionID&".xml"))

	' Create our DOM Document Objects

	Set domRoot = objDOM.documentElement

		For Each pagenode in domroot.Childnodes
			sAgentcode = Pagenode.Attributes.Item(0).Nodevalue
			sCommissiontype = Pagenode.Attributes.Item(1).Nodevalue
			sCurrcode = Pagenode.Attributes.Item(2).Nodevalue
			iCommissionvalue = Pagenode.Attributes.Item(3).Nodevalue
			iSaletransno = Pagenode.Attributes.Item(4).nodevalue
			sAccTransno = Pagenode.Attributes.Item(5).nodevalue

			sQuery="update Sal_T_AdditionalAgents set CommissionToPay=0	where AccTransactionNo="&sAccTransno&" and AgentCode="&sAgentcode
			'Response.Write sQuery
			con.execute(sQuery)
		Next

end if ' if sCallFrom = "SC" then


'Con.RollbackTrans
'Response.End
con.CommitTrans

'Response.End



'-----------------------PROCESS ENTRY NODES FOR CC/ANAL NODES-------------
Dim sCheckForProc

FOR EACH EntryNode IN Root.childNodes
	IF  EntryNode.nodeName="Entry" THEN
		sEntryno=EntryNode.Attributes.getNamedItem("No").value
		sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
		sAmount=EntryNode.Attributes.getNamedItem("Amount").value
		sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value

		sExp="//Entry[@No='"&sEntryno&"']/AccHead"
		set tempNode=Root.selectNodes(sExp)
		sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

		set nodANL=oDOM.createElement("Root")
		nodANL.setAttribute "TransNo",iTransNo
		nodANL.setAttribute "EntryNo",sEntryno
		nodANL.setAttribute "UnitCode", sAccUnit
		nodANL.setAttribute "GlHead",sAccCode
		nodANL.setAttribute "ACTFlag","C"

		sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
		set tempNode=Root.selectNodes(sExp)
		if tempNode.length >0 then
			sCheckForProc = "1"
			set HeaderNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(HeaderNode)
			bAddFlag=true
		end if

		sExp="//Entry[@No='"&sEntryno&"']/Analytical"
		set tempNode=Root.selectNodes(sExp)
		if tempNode.length >0 then
			set HeaderNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(HeaderNode)
			bAddFlag=true
		end if
		if bAddFlag then
		  Dim adoConn

		   Set adoConn = Server.CreateObject("ADODB.Connection")
		   adoConn.ConnectionString = con
		   adoConn.CursorLocation = 3
		   adoConn.Open

		   sQuery = "Proc_VouCCANALUpdate"



		   Dim adoCmd
		   Set adoCmd = Server.CreateObject("ADODB.Command")
		   Set adoCmd.ActiveConnection =adoConn
		   adoCmd.CommandText = sQuery
		   adoCmd.CommandType = 4 'adCmdStoredProc
		   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)

		   Dim adoRS
		   Set adoRS = adoCmd.Execute()

		end if
	end if
	bAddFlag = false
NEXT






'------------------END OF PROCESSING ENTRY NODES--------------------------


if con.Errors.count <>0 then
	for iCounter=0 to con.Errors.count - 1
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else

	Set newElem  = oDOM.createAttribute("TransNo")
	newElem.value = iTransNo
	Root.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("VoucherNo")
	newElem.value = sVouNo
	Root.setAttributeNode(newElem)



	'Response.End

	'This been Blocked.
	''blocked by Ragav on Jan13,2012
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
	''end

	'---
	If Trim(sCallFromFD) <> "Y" then

		if trim(sCallType)="CN" then Response.Redirect ("CREDITNOTETOCREATE.ASP")

		IF CStr(Trim(sCallType)) <> "A" Then
			select case sVouName
				Case "CA" Response.Redirect ("VouCADisplay.asp?TransNo="&iTransNo&"&Approver="&sApprover)
				Case "BA" Response.Redirect ("VouBADisplay.asp?TransNo="&iTransNo&"&Approver="&sApprover)
				Case "GJ" Response.Redirect ("VouGJDisplay.asp?TransNo="&iTransNo&"&Approver="&sApprover)
			end select
		Else
			Response.Redirect "AccVouGenerate.Asp?hTransNo="&iTransNo&"&hVouName="&sVouName&"&RedirTy=A"
		End IF

	End IF
	'---

end if
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
			sQuery = "Select CreatedDrSeriesNo,CreatedDrSeriesCode From Acc_M_BookNumberSeries  " & _
					 "Where BookCode = '"&sBookCode&"' And OUDefinitionID = '"&sUnit&"' " & _
					 "And BookNumber = "&iBookNo&" And FinPeriod = '"&Session("FinPeriod")&"' " 'ibookno is added on 28th jan 09 to use same function for both internal receipt and contra entry
		Else
			sQuery = "Select CreatedCrSeriesNo,CreatedCrSeriesCode From Acc_M_BookNumberSeries  " & _
								 "Where BookCode = '"&sBookCode&"' And OUDefinitionID = '"&sUnit&"' " & _
					 "And BookNumber = "&iBookNo&" And FinPeriod = '"&Session("FinPeriod")&"' "
		End If
		'Response.Write sQuery &"<br><br><br>"

		sRecrs.Open sQuery,Con
		IF Not sRecrs.EOF Then
			iRecSerNo = sRecrs(0)
			iRecSerCode = sRecrs(1)
		End IF
		sRecrs.Close

		sQuery = "Select Number,isNull(Prefix,''),isNull(Suffix,'') From APP_R_NoSeriesModuleEntry Where SeriesNo = "&iRecSerNo&" and " & _
				 "SeriesCode = "&iRecSerCode&" And OUDefinitionID = '"&sUnit&"' " & _
				 "And Period = '"&sLastYear&"' "

		'Response.Write sQuery &"<br><br><br>"
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
