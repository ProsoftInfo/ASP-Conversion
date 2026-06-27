<%

Function CreateVou(sVouCode,sVouName,sCrVouNo,iVouNo,iOldBkNo,sOldVDate,iOldCrTransNo,iOldAccTransNo,sCrConVouNo,sConVouNo,sNorToContra)
	Dim oDOM,nodHeader,Root,objRs,sQuery,sExp,dTotal,sTransType,dCRAmt,dDRAmt
	dim EntryNode,HeaderNode,nodANL,newElem,tempNode,bAddFlag,sInsChk
	dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp,objDom
	dim sOrgId,sBookNo,sVouType,sVouNo,sApprove,sVoucDate,sAccUnit
	dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo,sDocType,sVouStatus
	dim iAdvNo,sApprover,sPayTo,iTdsAmount,iTdsPer,AdvNode,iCtr,sCallType,sCodeTy
	Dim sAccHeadCode,sSelVouDate,sCrnoTrno,iOldCrNo,sTempVouNo,sTemparr,sTemp2
	DIM sInsType,sInsNo,sInsDate,sPayableat,sDrawnOn,dAddAmt,dSubAmt
	DIM sCCGroup,sAddCode,sAddGroupCode,sAddRatio,sAddAmount,dAddTotal,sAdjTy,iCrAdvNo
	Dim sInsPayat,sInsDrawat,sCheckForProc,iCreatedVouNo,iCreatedTransNo
	Dim iHeadAccHead,iEntAccHead,iEntUnit,iConCount,sHdFlag,sToBookCode,sToBookNo
	Dim sToBookName,sNewVouType,sNewVouchNo,sNewTransType,OldNarrNode,iEntNo,iTdsGroupID
	Dim sHeadAccName,sEntAccHeadName,iHeadUnit,iCrTransNo,sNewAccVouNo,iNewTransNo
	Dim iHeadUnitName,AccNode,NarrNode,sContraCheck,sRevVal,sNewNarr,sSelInvNo,sSelInvDate
	Dim sAdjType,bPartyReceiptFlag,iFrmTransNo,sFrmTransTy,iSeriesNo,iSeriesCode,iToAccCode
	Dim iFrmAccCode,iConChkCnt,sRetValChk,sRevConChk,iCrSeriesNo,iCrSeriesCode,sNoSerTy
	Dim sFrmCrDrTy,iMon,iYear,AdjRoot,AdjDom,iToBeCrAdvNo,sCheckAdjRootVal,objfs
	Dim iTdsValue,iTDSEntryAmt,iTDSEntryPer,iPayRecAmt,sFormula,sTDSNarr,TDSFlag,sVouEntryno,iAccHeadNo
	Dim sGrpid,sTdsCheck,sTdsType,dTotalTdsAmt
	Dim sConInsChk,iContAccCode,sVouMonYr,sLastDate,sVouYrMon,sConTrType,sConNarr
	Dim sNoChFlag,sErrChk,iFrmBookNo,sMonYear,oDOMNew
	Dim sOtherApplnTableName,nOtherApplnTransNo,nFromApplication,dtClearedOn
	Dim sInternalNo

	'Response.Clear

	'sVouCode = Request("hVouCode")
	'sVouName = Request("hVouName")
	sRevConChk = "F"
	sErrChk = "F"
	'Response.Write iOldAccTransNo
	sApprove = "Y"
	sApprover = getUserID()
	bAddFlag=false
	sVouStatus="010104" 'Crearted For Accounting
	sTdsCheck = "N"

		sNoChFlag = False 'Denotes that Only Old Number Series Number Will be Used

	Set objRs  = server.CreateObject("adodb.recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objDom = Server.CreateObject("Microsoft.XMLDOM")
	Set AdjDom = Server.CreateObject("Microsoft.XMLDOM")
	
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set oDOMNew = Server.CreateObject("Microsoft.XMLDOM")
	
	oDOM.Load server.MapPath("../temp/transaction/Voucher Entry_"&sVouName&"_"&Session.SessionID&".xml")
	Set Root = oDOM.documentElement
	
	iHeadAccHead = Root.Attributes.getNamedItem("BookAcchead").value
	iHeadUnit = Root.Attributes.getNamedItem("UnitNo").value
	iHeadUnitName = Root.Attributes.getNamedItem("UnitName").value
	
	sCheckAdjRootVal = "N"
	if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher ADJAMD_"&sVouName&"_"&Session.SessionID&".xml")) then
		AdjDom.Load server.MapPath("../temp/transaction/Voucher ADJAMD_"&sVouName&"_"&Session.SessionID&".xml")
		Set AdjRoot = AdjDom.documentElement
		sCheckAdjRootVal = "Y"
	End IF
 
	dTotalTdsAmt = 0
	sOrgId = Root.Attributes.getNamedItem("UnitNo").value
	sBookNo = Root.Attributes.getNamedItem("BookNo").value
	sVouType = Root.Attributes.getNamedItem("CRDR").value
	sVoucDate = Root.Attributes.getNamedItem("VouDate").value
	sAccHeadCode = Root.Attributes.getNamedItem("BookAcchead").value
	
	
	sOtherApplnTableName	= Root.Attributes.getNamedItem("OtherApplnTableName").value
	nOtherApplnTransNo		= Root.Attributes.getNamedItem("OtherApplnTransNo").value
	nFromApplication		= Root.Attributes.getNamedItem("FromApplication").value
	dtClearedOn				= Root.Attributes.getNamedItem("ClearedOn").value
	
	
		 
	
	if trim(sOtherApplnTableName) = "" then 
		sOtherApplnTableName= "NULL"
	else
		sOtherApplnTableName= "'" & sOtherApplnTableName & "'"
	end if 		
	
	if trim(nOtherApplnTransNo) = "" then nOtherApplnTransNo= "NULL"
	
	if trim(nFromApplication) = "" then 
		nFromApplication= "NULL"
	else
		nFromApplication= "'" & nFromApplication & "'"
	end if 	
	if trim(dtClearedOn) = "" then 
		dtClearedOn= "NULL"
	else
		dtClearedOn = "convert(datetime,'" & dtClearedOn & "',103)"
	end if 	
	
	'Response.Write "<p> dtClearedOn = "  & dtClearedOn
	'Response.End
	
	
	iTdsGroupID = Request.Form("SelTDSGrp")
	If iTdsGroupID = "" Then iTdsGroupID = 0 'added on 23June 2009
	sQuery = "Select M.CounterType From Acc_M_BookNumberSeries B,Ms_NumberSeries M "&_
			 "Where M.SeriesNo = B.DrSeriesNo And B.BookCode = '"&sVouCode&"' And "&_
			 "B.BookNumber = "&sBookNo&" And B.OUDefinitionID = '"& sOrgId &"' "
Response.Write "<p>"& sQuery
	objRs.open sQuery,con
	if not objRs.EOF then
		sNoSerTy = Trim(objRs(0))
	end if
	objRs.close()

	sVouMonYr = Trim(Mid(sVoucDate,4,2))
	sVouYrMon = Trim(Right(sVoucDate,4))
	sVouMonYr = sVouMonYr & Trim(Right(sVoucDate,4))
	sVouYrMon = sVouYrMon & Trim(Mid(sVoucDate,4,2))
	sLastDate = GetLastDayMonYr(sVouYrMon)
	iFrmBookNo = sBookNo
	'Response.Write sNoSerTy &"<br><br>"

	IF CStr(sBookNo) <> CStr(iOldBkNo) Then
	    Response.Write "<p>Book Num is Changed <p>"
	    if Trim(sCrConVouNo)="0" then
		    IF strcomp(sVouType,"D")=0 THEN
			    sQuery = "select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
					     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
		    ELSE
			    sQuery = "select CrSeriesNo,CrSeriesCode,CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
					     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
		    END IF
		    'Response.Write "sQuery="&sQuery
		    objRs.open sQuery,con
		    if not objrs.EOF then
			    iSeriesNo=objRs(0)
			    iSeriesCode=objRs(1)
			    iCrSeriesNo=objRs(2)
			    iCrSeriesCode=objRs(3)
			end if'if not objrs.EOF then
		    objRs.close()

		    iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
		    'Response.Write "<p>iVouNo="&iVouNo
		    sCrVouNo = GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sVoucDate)
		    'Response.Write "<p>sCrVouNo="&sCrVouNo
		else
		    sCrVouNo  = sCrConVouNo
	        iVouNo = sConVouNo
		end if 'if Trim(sCrConVouNo)<>"0" then
		
	ElseIF Mid(sVoucDate,4,2) <> Mid(sOldVDate,4,2) and CStr(sNoSerTy) <> "Y"    Then 'IF The Voucher Date is been changed from month to month
		Response.Write "Inside " &"<br><br>"
		Response.Write "<p>sVoucDate = "& sVoucDate 
		Response.Write "<p>OldVouCDate = "& sOldVDate
		Response.Write "<p>"
		IF strcomp(sVouType,"D")=0 THEN
			sQuery = "select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
					 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
		ELSE
			sQuery = "select CrSeriesNo,CrSeriesCode,CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
					 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
		END IF

		Response.Write sQuery &"<br><br>"
		objRs.open sQuery,con
			iSeriesNo=objRs(0)
			iSeriesCode=objRs(1)
			iCrSeriesNo=objRs(2)
			iCrSeriesCode=objRs(3)
		objRs.close()

		iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
		sCrVouNo = GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sVoucDate)
    else
    
        Response.Write " <p> Else Part CrTransNO ="& sCrVouNo 
        Response.Write "<p> VouCode= "& sVouCode 
        
        
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
		    sCrVouNo = GetBookReceiptNoCrVou(sVouCode,sOrgId,-1,sVouType) 'Added on 28th Jan 09 for Contra entry
			iVouNo = GetBookReceiptNoVou(sVouCode,sOrgId,-1,sVouType) 'Added on 28th Jan 09 for Contra entry
		Else
		    
		    IF strcomp(sVouType,"D")=0 THEN
			    sQuery = "select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
					     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
		    ELSE
			    sQuery = "select CrSeriesNo,CrSeriesCode,CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
					     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
		    END IF

		    Response.Write sQuery &"<br><br>"
		    objRs.open sQuery,con
			    iSeriesNo=objRs(0)
			    iSeriesCode=objRs(1)
			    iCrSeriesNo=objRs(2)
			    iCrSeriesCode=objRs(3)
		    objRs.close()

		    iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
		    sCrVouNo = GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sVoucDate)
		End If
		
	End IF

	'Response.Write iVouNo &"   " & sCrVouNo


	IF CStr(sCrConVouNo) = "0" Then
		sNoChFlag = True 'Denotes New Number Series Number will be Used
	End IF
	
        
    
Response.Write "<p> Cr Vou No = "& sCrVouNo  



	iFrmAccCode = sAccHeadCode

	Root.Attributes.getNamedItem("CreatedVoucherNo").value = sCrVouNo
	Root.Attributes.getNamedItem("VoucherNo").value = iVouNo
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
	Dim iTdsEntryType
	sExp = "//TDS"
	Set tempNode = Root.selectNodes(sExp)
	IF tempNode.length <> 0 Then
		For iCtr = 0 To tempNode.length - 1
			Response.Write CDbl(tempNode.Item(iCtr).Attributes.getNamedItem("PayRecAmount").Value) &"<br><br>"
			iTdsAmount = iTdsAmount + CDbl(tempNode.Item(iCtr).Attributes.getNamedItem("PayRecAmount").Value)
		Next
	Else
		iTdsAmount = 0
	End IF

	dTotalTdsAmt = iTdsAmount
	'Response.Write iTdsAmount &"<br><br>"
	'Response.Write "dTotal = " & dTotal &"<br><br>"



	dTotal = CDbl(dTotal) - CDbl(iTdsAmount)

	'Response.Write "dTotal = " & dTotal &"<br><br>"

	Select Case CStr(sTransType)
		Case "CAP"	sConTrType = "CSP"
		Case "CAR"	sConTrType = "CSR"
		Case "BAP"	sConTrType = "BSP"
		Case "BAR"	sConTrType = "BSR"
		Case "GJR"	sConTrType = "GJS"
	End Select


	'Con.BeginTrans
	sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCreatedTransNo=objRs(0)
	objRs.Close


	sFrmCrDrTy = sVouType
	iMon = Trim(Mid(sVoucDate,4,2))
	iYear = Year(sVoucDate)
	iMon = CInt(iMon)
	iYear = CInt(iYear)


	IF sVouType="" or sVouCode="08" THEN
		sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
				 "CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,TDSGroupID) values"&_
				 "("&iCreatedTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				 "NULL,NULL,'"&sCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal &_
				 ",NULL," & sApprover & ",getdate(),NULL,'"& sVouStatus & "',"&iTdsGroupID&")"
	ELSE
		sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
				 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,OtherApplnTableName,TDSGroupID) values"&_
				 "("&iCreatedTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				 "NULL,"&sAccHeadCode&",'"&sCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				 ",'"&sPayTo&"','"&sVouType&"',"&sApprover&",getdate(),NULL,'"&sVouStatus&"','"&sCrnoTrno&"',"&iTdsGroupID&")"
	END IF
	Dim i
	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)
	dim sOption,iInsEntNo,sAction,RootNew,sBKInsType,sBkInsNo,sBkInsDate,sBkPayableat,sBkDrawnOn,iBkInsAmt,iBankInsEntNo


	Response.Write "<br>%%%%%%%%%%%%%%"& sVouCode &"%%%%%%%%%%%%%%<br>"
	iTransNo =   iOldCrTransNo
	IF sVouCode = "02" Then
		sExp = "//BankInstrumentDet"
		Set tempNode = Root.selectNodes(sExp)

		'sQuery = "Delete from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo  = "&iCreatedTransNo&" "
		sQuery = "Delete from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo  = "&iTransNo&" "
		Response.Write "<BR><BR>" & sQuery
		con.execute(sQuery)
		
		iCtr = 1
		Response.Write "<b>len="& tempNode.length &"</b><BR>"
		For i = 0 to tempNode.length -1
			
			sBKInsType		= tempNode.Item(i).Attributes.getNamedItem("InsType").Value
			sBkInsNo		= tempNode.Item(i).Attributes.getNamedItem("InsNo").Value
			sBkInsDate		= tempNode.Item(i).Attributes.getNamedItem("InsDate").Value
			iBkInsAmt		= tempNode.Item(i).Attributes.getNamedItem("InsAmt").Value
			sBkPayableat	= tempNode.Item(i).Attributes.getNamedItem("PayAt").Value
			sBkDrawnOn		= tempNode.Item(i).Attributes.getNamedItem("DrawnOn").Value
			sOption			= tempNode.Item(i).Attributes.getNamedItem("Option").value
			sAction			= tempNode.Item(i).Attributes.getNamedItem("Action").value
			
						
			iBankInsEntNo = "0"		
			if tempNode.Item(i).Attributes.length >= 10 then
				iBankInsEntNo	= tempNode.Item(i).Attributes.getNamedItem("BankInsEntNo").value
			end if 	

			

		'**************************Newly Added by Maheswari on 17th June 2008********************************

			IF trim(iBankInsEntNo) <> "0"  then
				sTemp	  = split(iBankInsEntNo,"-")
				iEntNo	  = sTemp(0)
				iInsEntNo = sTemp(1)
				sBkInsNo  = sTemp(2)
			End If
			'Response.Write "sInsNo="&sInsNo  &"<BR>"
			'Response.Write "Option="&sOption&"<BR>"

			If trim(sOption) = "Y"  then
				IF trim(sAction) = "C" then
					sQuery = "Update Acc_R_BankInstrumentUsage set Status = 'C' where  CreatedTransNo  = "&iCreatedTransNo &" "&_
						 " and EntryNo = "&iEntNo&"  and InstrumentEntryNo = "&iInsEntNo&" and InstrumentNo = '"&sBkInsNo &"' "
					Response.Write "<BR><BR>" & sQuery
					con.execute(sQuery)
				ElseIF trim(sAction) = "R" then
					sQuery = "Update Acc_R_BankInstrumentUsage set CreatedTransNo  = 0,Status = 'N' where  CreatedTransNo  = "&iCreatedTransNo&" "&_
							"and EntryNo = "&iEntNo&" and InstrumentEntryNo = "&iInsEntNo&" and InstrumentNo = '"&sBkInsNo&"' "
					Response.Write "<BR><BR>" & sQuery
					con.execute(sQuery)
				End IF
				'Response.Write "<B>Insertion in Acc_T_CreatedVoucherInstrumentDet</B>"&"<BR>"
				sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
						 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,BankInstrumentEntryNo,InstrumentEntryNo1,InstrumentAmount,ClearedOn)"&_
						 " Values ("&iCreatedTransNo &","&iCtr&",'"&sBkInsType&"','"&sBkInsNo&"',convert(datetime,'"&sBkInsDate&"',103),"&_
						 "'"&sBkPayableat&"','"&sBkDrawnOn&"',"&iEntNo&","&iInsEntNo&","&iBkInsAmt&"," & dtClearedOn & " ) "
				Response.Write "<BR><BR>" & sQuery
				con.execute(sQuery)

			Else
				sQuery = "Update Acc_R_BankInstrumentUsage set CreatedTransNo  = "&iCreatedTransNo&",Status = 'U' where  InstrumentNo = '"&sBkInsNo&"' "
						'" and EntryNo = "&iEntNo&" and InstrumentEntryNo = "&iInsEntNo&"  "
				Response.Write "<BR><BR>" & sQuery
				con.execute(sQuery)

				'Response.Write "<B>Insertion in Acc_T_CreatedVoucherInstrumentDet</B>"&"<BR>"
				sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
						 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,InstrumentAmount,ClearedOn)"&_
						 " Values ("&iCreatedTransNo&","&iCtr&",'"&sBkInsType&"','"&sBkInsNo&"',convert(datetime,'"&sBkInsDate&"',103),"&_
						 "'"&sBkPayableat&"','"&sBkDrawnOn&"',"&iBkInsAmt&"," & dtClearedOn & " ) "
				Response.Write "<BR><BR>" & sQuery
				con.execute(sQuery)
			End IF
			iCtr = iCtr + 1
		Next
	End IF 'IF sVouCode = "02" Then
'****************************************************************************************************

	sQuery = "Select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close

	IF sVouCode="02" THEN

		iFrmTransNo = iTransNo
		sFrmTransTy = sTransType

'		Response.Clear

		sQuery = "insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				 "PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
				 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus,ClearedOn) values "&_
				 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				 "NULL,"&sAccHeadCode&",'"&iVouNo&"','"&sCrVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				 ",'"&sPayTo&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"'," & dtClearedOn & ") "


	ELSEIF sVouType="" or sVouCode="08" THEN
		sQuery = "insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				 "PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
				 "CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"&_
				 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				 "NULL,NULL,'"&iVouNo&"','"&sCrVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				 ",NULL,"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"
	ELSE
		sQuery = "insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
				 "PartyType,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
				 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"&_
				 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
				 "NULL,"&sAccHeadCode&",'"&iVouNo&"','"&sCrVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				 ",'"&sPayTo&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"
	END IF

	Response.Write "<BR><BR>" & sQuery	
	con.execute(sQuery)
	'Response.End
 
 
	FOR EACH EntryNode IN Root.childNodes
	 
		IF  EntryNode.nodeName="Entry" THEN
			sEntryno=EntryNode.Attributes.getNamedItem("No").value
			sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
			sAmount=EntryNode.Attributes.getNamedItem("Amount").value
			sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value
			
			iF CStr(sVouCode) <> "08" Then
				sGrpID = EntryNode.Attributes.getNamedItem("GroupId").value
			END IF
			Response.Write i &"#####"&"<BR>"
			
			IF CStr(sVouCode) = "01" or CStr(sVouCode) = "02" or CStr(sVouCode) = "08" Then
				iTdsAmount = EntryNode.Attributes.getNamedItem("TdsAmount").value
				iTdsPer = EntryNode.Attributes.getNamedItem("TdsPercentage").value
				
				if trim(iTdsAmount) = "" then iTdsAmount = 0
				if trim(iTdsPer) = "" then iTdsPer = 0
			Else
				iTdsAmount = 0
				iTdsPer = 0
			End IF
			sAccCode = ""
			sTDSNarr = "TDS Entry"
			TDSFlag = "Y"
			sVouEntryno  = sEntryno + 1
			sExp = "//Entry [@No = "& sEntryno & "]/AccHead"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.length <> 0 Then
				sAccCode=TempNode.Item(0).Attributes.Item(0).nodeValue
				sAccType=TempNode.Item(0).Attributes.Item(4).nodeValue
				Response.Write TempNode.Item(0).Attributes.Item(3).nodeValue
			End IF
			
			FOR EACH HeaderNode IN EntryNode.childNodes
				IF HeaderNode.nodeName="TDS" THEN
					sTdsCheck = "Y"
					iAccHeadNo	 = HeaderNode.getAttribute("AccHeadCode")
					iTDSEntryAmt = HeaderNode.getAttribute("TDSAmount")
					iTDSEntryPer = HeaderNode.getAttribute("TdsPercentage")
					iPayRecAmt	 = HeaderNode.getAttribute("PayRecAmount")
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
					sQuery =" Insert into Acc_T_CreatedVoucherDetails(CreatedTransNo,AccountingUnit,VoucherEntryNumber,"&_
							" AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"&_
							" VoucherNarration,Amount,TransCrDrIndication,TDSonAmount,TDSPercentage,PackCode,TDSFlag) "&_
							" values("&iCreatedTransNo&",'"&sAccUnit&"',"&sVouEntryno&","&iAccHeadNo&",NULL,NULL,NULL, "&_
							" '"&sTDSNarr&"',"&iPayRecAmt&",'"&sEntryType&"',"&iTDSEntryAmt&","&iTDSEntryPer&","&sGrpId&",'"&TDSFlag&"')"
							Response.Write "<BR>TDS="& sQuery & "<BR><BR>"
							con.execute(sQuery)
					sVouEntryno = sVouEntryno + 1
				END IF 'IF HeaderNode.nodeName="TDS" THEN


				IF 	HeaderNode.nodeName="Narration" THEN
						sNarration=Replace(HeaderNode.text,"'","''")
				END IF 'End of Check for Narration Node
			NEXT
	'-------------END OF PROCESSING CHILD NODES OF ENTRY NODE---------------------
			Response.Write "<p>sEntryno======="&sEntryno&"--"&sAccCode&"<Br><Br>"

			IF StrComp(sAccType,"G")=0 THEN
				sQuery = "insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
				sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
				sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
				sQuery = sQuery& iCreatedTransNo&",'"&sAccUnit&"'"
				sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
				sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&iTdsAmount&","&iTdsPer&")"
			ELSE
				sTemp=Split(sAccCode,"?")
				sQuery = "insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
				sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
				sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication, TDSonAmount, TDSPercentage) values ("
				sQuery = sQuery& iCreatedTransNo&",'"&sAccUnit&"'"
				sQuery = sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
				sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&iTdsAmount&", "&iTdsPer&")"
			END IF
		Response.Write "T1="&sQuery& "<BR><BR>"
			con.execute(sQuery)

			IF sEntryType="C" THEN
				dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			ELSE
				dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			END IF

			'sExp = "//AccHead"
			sExp = "//Entry [@No = "& sEntryno & "]/AccHead"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.length <> 0 Then
				sAccCode=TempNode.Item(0).Attributes.Item(0).nodeValue
				sAccType=TempNode.Item(0).Attributes.Item(4).nodeValue  
			End IF

			FOR EACH HeaderNode IN EntryNode.childNodes
				IF HeaderNode.nodeName="TDS" THEN
					iAccHeadNo	 = HeaderNode.getAttribute("AccHeadCode")
					iTDSEntryAmt = HeaderNode.getAttribute("TDSAmount")
					iTDSEntryPer = HeaderNode.getAttribute("TdsPercentage")
					iPayRecAmt	 = HeaderNode.getAttribute("PayRecAmount")
					sFormula	 = HeaderNode. getAttribute("Formula")

					IF CStr(sEntryType) = "C" Then
						sTdsType = "D"
					Else
						sTdsType = "C"
					End IF

					sQuery = "insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
					sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
					sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,PackCode) values ("
					sQuery = sQuery& iTransNo&",'"&sAccUnit&"'"
					sQuery = sQuery& ","&sVouEntryno&","&iAccHeadNo&",NULL,NULL,NULL,"
					sQuery = sQuery&" '"&sTDSNarr&"',"&iPayRecAmt&",'"&sTdsType&"', "&iTDSEntryAmt&","&sGrpId&")"

					Response.Write "<BR>TDS="& sQuery & "<BR><BR>"
					con.execute(sQuery)

					sVouEntryno = sVouEntryno + 1
'+++++++++++++++++++++++++++++++++++TDS Enteries Copied ++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++++++++++++++++++++++TDS Enteries Copied ++++++++++++++++++++++++++++++++++++++++++++++
					'sVouEntryno = sVouEntryno + 1
				END IF 'IF HeaderNode.nodeName="TDS" THEN
				IF 	HeaderNode.nodeName="Narration" THEN
						sNarration=Replace(HeaderNode.text,"'","''")
				END IF 'End of Check for Narration Node
			NEXT

			sTDSNarr = sTDSNarr &" For the Voucher " & iVouNo &" Dt: " & sVoucDate
			'IF CDbl(dTotalTdsAmt) > 0 and UBound(sTemp) > 2 Then
			'	sVouEntryno = sVouEntryno + 1
			'	sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
			'	sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			'	sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage,PackCode) values ("
			'	sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
			'	sQuery=sQuery& ","&sVouEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(3)&","&sTemp(1)&", "
			''	sQuery=sQuery&" '"&sTDSNarr&"',"&dTotalTdsAmt&",'"&sEntryType&"',"&iTDSEntryAmt&","&iTDSEntryPer&","&sGrpId&")"

			'	Response.Write "<br>" & sQuery & "<BR>"
			'	con.execute(sQuery)
			'End IF


			sVouEntryno = sVouEntryno + 1
			Response.Write "<p>Chk="& sAccCode &"<br><br>"
			sTemp=Split(sAccCode,"?")
			IF StrComp(sAccType,"G")=0  THEN
				sQuery = "insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
				sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage) values ("
				sQuery = sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
				sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&", "&iTdsPer&")"
			ELSE
				IF sEntryType="C" THEN
					bPartyReceiptFlag=true
				END IF

				IF CStr(sTdsCheck) = "Y" Then
					sAmount = dTotal
				End IF
		
				
				sQuery = "insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
				sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage) values ("
				sQuery = sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery = sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
				sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&", "&iTdsPer&")"
			END IF
			
			Response.Write sQuery& "<BR><BR><BR>"
			con.execute(sQuery)
			
			sVouEntryno = sVouEntryno + 1
			IF CDbl(dTotalTdsAmt) > 0 and UBound(sTemp) > 2 Then

				sInternalNo = 0
				IF CStr(sVouType) = CStr(sEntryType) Then
					Response.Write "Calling From 1 <br><br><br>"
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

				Response.Write "<br>" & sQuery & "<BR>"
				con.execute(sQuery)
			End IF

			iEntNo = sEntryno
			dAddTotal=0


			FOR EACH HeaderNode IN EntryNode.childNodes 'Additional Details Updation Starts Here
				IF 	HeaderNode.nodeName="PayRec" THEN
					FOR EACH  nodANL IN HeaderNode.childNodes
						sAddCode=nodANL.Attributes.getNamedItem("No").value
						sAddAmount=nodANL.Attributes.getNamedItem("AmtToAdjust").value
						sDocType=nodANL.Attributes.getNamedItem("DocType").value
						sAdjTy = nodANL.Attributes.getNamedItem("AdjType").value
						iAdvNo = nodANL.Attributes.getNamedItem("PayableNo").value
						sSelInvNo = nodANL.Attributes.getNamedItem("InvNo").value
						sSelInvDate = nodANL.Attributes.getNamedItem("InvDate").value
						sCodeTy = ""

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


						IF CStr(Trim(sAdjTy)) = "D" Then
							sQuery = "Select DrCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&sAddCode&" "
							objRs.Open sQuery,Con
							IF Not objRs.EOF Then
								sAddCode = objRs(0)
								sCodeTy = "R"
							End IF
							objRs.Close
						Elseif CStr(Trim(sAdjTy)) = "C" Then
							sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&sAddCode&" "
							objRs.Open sQuery,Con
							IF Not objRs.EOF Then
								sAddCode = objRs(0)
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

						sAdjType = sAdjTy
	'**************************** Created Voucher Adjustment Details Starts Here *************************************************************************************************************************************
						IF CStr(sAdjTy) <> "P" and CStr(sAdjTy) <> "R" Then
							IF CDbl(sAddAmount) > 0 THEN
								IF sDocType="C" and CStr(sCodeTy) = "" THEN
									sQuery = "Select PayablesNumber From Acc_T_CreatedPayables Where PayablesNumber = "&sAddCode&" "&_
											 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "


									objRs.Open sQuery,con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									End IF

									'Response.Write sQuery &"<br><br>"
									objRs.Close
									con.execute(sQuery)
								ELSEIF sDocType="D"  and CStr(sCodeTy) = "" THEN
								
								'-----------------------UPDATE RECEIVABLE TABLE---------------------------
									sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where ReceivableNumber = "&sAddCode&" "&_
											 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "

									 Response.Write sQuery &"<br>"
									 'Response.End 
									objRs.Open sQuery,con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									End IF
									objRs.Close
								'	Response.Write sQuery &"<br>"
									con.execute(sQuery)
								Elseif CStr(sCodeTy) = "R" Then
									sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where ReceivableNumber = "&sAddCode&" "&_
											 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "
									objRs.Open sQuery,con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									End IF
									objRs.Close
									'Response.Write sQuery &"<br><br>"
									con.execute(sQuery)
								Elseif CStr(sCodeTy) = "P" Then
									sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where ReceivableNumber = "&sAddCode&" "&_
											 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "
									objRs.Open sQuery,con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iCreatedTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									End IF
									objRs.Close

								'	Response.Write sQuery &"<br><br>"
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

							'Response.Write sQuery &"<br><br>"
							Con.Execute sQuery

							sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, "&_
									 "PaidOn, AmountPaid, AdjustType,VoucherEntryNumber) "&_
									 "VALUES ("&iCrAdvNo&", "&iCreatedTransNo&", Convert(datetime,getDate(),103), "&sAddAmount&", 'A',"&iEntNo&") "

							'Response.Write sQuery &"<br><br>"
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

							'Response.Write sQuery &"<br><br>"
							Con.Execute sQuery

							sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, "&_
									 "ReceivedOn, AmountReceived, AdjustType,VoucherEntryNumber) "&_
									 "VALUES ("&iCrAdvNo&", "&iCreatedTransNo&", Convert(datetime,getDate(),103), "&sAddAmount&", 'A',"&iEntNo&") "

							'Response.Write sQuery &"<br><br>"
							Con.Execute sQuery

						End IF 'Created Voucher Adjustements Completed.

	'**************************** Created Voucher Adjustment Details Ends Here *************************************************************************************************************************************
	'**************************** Accounted Voucher Adjustment Details Starts Here *************************************************************************************************************************************
						'Response.Clear

						sAddCode=nodANL.Attributes.getNamedItem("PayableNo").value
						sAddAmount=nodANL.Attributes.getNamedItem("AmtToAdjust").value
						sDocType=nodANL.Attributes.getNamedItem("DocType").value
						sAdjType = nodANL.Attributes.getNamedItem("AdjType").value
						sSelInvNo = nodANL.Attributes.getNamedItem("InvNo").value
						sSelInvDate = nodANL.Attributes.getNamedItem("InvDate").value

						IF CStr(sAdjType) <> "P" and CStr(sAdjType) <> "R" Then
							IF CDbl(sAddAmount) >0 THEN
								IF CStr(sAdjType) = "I" Then
									sQuery = "Select ReceivableNumber From Acc_T_Receivables Where OUDefinitionID = '"&sOrgId&"'  "&_
											 "and ReceivableNumber = "&sAddCode&" and PartyType = '"&sTemp(0)&"' and  "&_
											 "PartySubType = "&sTemp(1)&" and PartyCode = "&sTemp(3)&"  "

									'Response.Write sQuery &"<br><br>"
									objRs.Open sQuery,Con
									IF Not objRs.EOF Then

										sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									End IF
									objRs.Close
									'Response.Write sQuery &"<br><br>"
									con.execute(sQuery)
								ElseIF CStr(sAdjType) = "PI" Then
									sQuery = "Select PayablesNumber From Acc_T_Payables Where OUDefinitionID = '"&sOrgId&"'  "&_
											 "and PayablesNumber = "&sAddCode&" and PartyType = '"&sTemp(0)&"' and  "&_
											 "PartySubType = "&sTemp(1)&" and PartyCode = "&sTemp(3)&"  "
									objRs.Open sQuery,Con
									IF Not objRs.EOF Then
										sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
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
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
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
												 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									Else
										sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
												 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
												 "getdate(),"&sAddAmount&","&iEntNo&")"
									End IF
									objRs.Close
									con.execute(sQuery)
								End IF
							END IF
						Else
							sQuery = "UPDATE Acc_T_AdvancePayments SET AdvanceAdjusted = isNull(AdvanceAdjusted,0) + "&sAddAmount&"  "&_
									 "WHERE AdvanceNumber = "&sAddCode&" "
							con.execute(sQuery)
						End IF 'Accounted Voucher Adjustments Completed.
	'**************************** Accounted Voucher Adjustment Details Ends Here *************************************************************************************************************************************
					NEXT
				END IF
			NEXT '-END OF PROCESSING CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION-

			'iToBeCrAdvNo = 0

			IF sAccType="P" AND CDbl(sAmount)> CDbl(dAddTotal) THEN
				Response.Write sVouType &" " & CDbl(sAmount) &"---> " &  CDbl(dAddTotal) &"<br><br>"
				IF sVouType="C" THEN
					sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
					'Response.Write sQuery
					objRs.Open sQuery,Con
					If Not objRs.EOF Then
						iCrAdvNo = objRs(0)
						iToBeCrAdvNo = objRs(0)
					End IF
					objRs.Close

					sQuery = "Select isNull(Max(AdvanceNumber),0)+1 from Acc_T_AdvancePayments "
					objRs.Open sQuery,Con
					If Not objRs.EOF Then
						iAdvNo = objRs(0)
					End IF
					objRs.Close

					sExp = "//Entry[@No="&sEntryno&"]/AccHead"
					Set AdvNode = Root.selectNodes(sExp)
					IF AdvNode.length <> 0 Then
						Set newElem = oDOM.createAttribute("AdvNo")
						newElem.value = iCrAdvNo
						AdvNode.Item(0).setAttributeNode(newElem)
					End IF

					IF Cdbl(dAddAmt) <> 0 Then
						sQuery = "INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
								 " VALUES("&iCrAdvNo&", "& iCreatedTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&","&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,NULL)"
						'Response.Write sQuery
						con.execute(sQuery)

						sQuery = "INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
								 " VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&","&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
						con.execute(sQuery)
					Else
						sQuery = "INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
								 " VALUES("&iCrAdvNo&", "& iCreatedTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&","&CDbl(sAmount)&",NULL,NULL)"
						'Response.Write sQuery
						con.execute(sQuery)

						sQuery = "INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
								 " VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&","&CDbl(sAmount)&",NULL,NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
						'	Response.Write sQuery & "<br><br><br>"
						con.execute(sQuery)

					End IF
				ELSEIF sVouType="D" THEN
					sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
					objRs.Open sQuery,Con
					If Not objRs.EOF Then
						iCrAdvNo = objRs(0)
						iToBeCrAdvNo = objRs(0)
					End IF
					objRs.Close

					sQuery = "Select isNull(Max(AdvanceNumber),0)+1 from Acc_T_AdvancePayments "
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
						sQuery = "INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
								 " VALUES("&iCrAdvNo&", "& iCreatedTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)- CDbl(dAddTotal)&",NULL)"
						'Response.Write sQuery
						con.execute(sQuery)

						sQuery = "INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
								 " VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
						'		Response.Write sQuery & "<br><br><br>"
						Con.execute(sQuery)
					Else
						sQuery = "INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
								 " VALUES("&iCrAdvNo&", "& iCreatedTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)&",NULL)"
						'Response.Write sQuery
						con.execute(sQuery)

						sQuery = "INSERT INTO Acc_T_AdvancePayments(AdvanceNumber,TransactionNumber, OUDefinitionID, PartyType, PartySubType, "&_
								 "PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted, CreatedAdvanceNo, CreatedTransNo)"&_
								 " VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
								 ""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)&",NULL,"&iCrAdvNo&","&iCreatedTransNo&")"
						'	Response.Write sQuery & "<br><br><br>"
						Con.execute(sQuery)
					End IF
				END IF
			END IF
		'-------------END OF PROCESSING FOR ADVANCE TABLE UPDATION ---------------
		END IF
		'-------------END OF ENTRY NODE CHECK ---------------
	NEXT

'Response.Clear

	IF CStr(sVouCode) = "02" Then
		sExp = "//BankInstrumentDet"
		Set tempNode = Root.selectNodes(sExp)
		IF tempNode.length <> 0 Then
			sInsType = tempNode.Item(0).Attributes.getNamedItem("InsType").value
			sInsNo = tempNode.Item(0).Attributes.getNamedItem("InsNo").value
			sInsDate = tempNode.Item(0).Attributes.getNamedItem("InsDate").value
			sInsPayat = tempNode.Item(0).Attributes.getNamedItem("PayAt").value
			sInsDrawat = tempNode.Item(0).Attributes.getNamedItem("DrawnOn").value

			sQuery	= "update Acc_T_CreatedVoucherHeader set CreatedVouchStatus='"&sVouStatus&"'," &_
					  "BankInstrumentType='"&sInsType&"',BankInstrumentNo='"&sInsNo&"',PayableAt='"&sInsPayat&"',"&_
					  "BankInstrumentDate=convert(datetime,'"&sInsDate&"',103),DrawnOnBank='"&sInsDrawat&"',"&_
					  "OtherApplnTableName=" & sOtherApplnTableName & ",OtherApplnTransNo=" & nOtherApplnTransNo & "," & _
					  "FromApplication=" & nFromApplication & ",ClearedOn=convert(datetime," & dtClearedOn & ",103)" & _
					  " where CreatedTransNo="&iCreatedTransNo
			Response.Write "<p>" & sQuery
			con.execute(sQuery)
		End IF
	End IF




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
		sNewNarr = sNewNarr&" - "&sNarration

		sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "
		Con.Execute sQuery

		sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "
		Con.Execute sQuery

		sQuery = "Update Acc_T_VoucherDetails Set VoucherNarration = '"&sNewNarr&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "
		Con.Execute sQuery
	End IF

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
			 "CreatedTransNo = "&iCreatedTransNo&" and AdjustType is NULL "
	objRs.Open sQuery,Con
	If Not objRs.EOF Then
		iRecTot = objRs(0)
	End if
	objRs.Close

	sQuery = "Select Count(1) From Acc_T_CreatedPybleAdjDet WHere  "&_
			 "CreatedTransNo = "&iCreatedTransNo&" and AdjustType is NOT NULL"
	objRs.Open sQuery,Con
	If Not objRs.EOF Then
		iAdvTot = objRs(0)
	End if
	objRs.Close

	sQuery = "Select Count(1) From Acc_T_CreatedRcvbleAdjDet WHere  "&_
			 "CreatedTransNo = "&iCreatedTransNo&" and AdjustType is NOT NULL "
	objRs.Open sQuery,Con
	If Not objRs.EOF Then
		iAdvTot = iAdvTot + objRs(0)
	End if
	objRs.Close

	IF CDbl(iAdvTot + iPayTot + iRecTot) <> 0 Then 'atleast one adjustment is made.
		IF CStr(iPayTot) <> "0" Then
			sQuery = "Select Count(1) From Acc_T_CreatedPybleAdjDet C,Acc_T_CreatedPayables P "&_
					 "Where P.PayablesNumber = C.PayablesNumber and C.CreatedTransno = "&iCreatedTransNo&" and  "&_
					 "P.AmountPayable >= P.AmountPaid and C.AdjustType is NULL "
			objRs.Open sQuery,con
			IF Not objRs.EOF and CStr(iPayTot) <> CStr(objRs(0)) Then
				sErrCheck = "Y" 'Error Found
			End if
			objRs.Close
		End IF

		IF CStr(iRecTot) <> "0" Then
			sQuery = "Select Count(1) From Acc_T_CreatedRcvbleAdjDet C,Acc_T_CreatedReceivables P "&_
					 "Where P.ReceivableNumber = C.ReceivableNumber and C.CreatedTransno = "&iCreatedTransNo&" and  "&_
					 "P.AmountReceivable >= P.AmountReceived and C.AdjustType is NULL "
			'Response.write sQuery
			objRs.Open sQuery,con
			IF Not objRs.EOF and CStr(iRecTot) <> CStr(objRs(0)) Then
				sErrCheck = "Y" 'Error Found
			End if
			objRs.Close
		End IF

		IF CStr(iAdvTot) <> "0" Then
			sQuery = "Select Count(1) From Acc_T_CreatedRcvbleAdjDet C,Acc_T_CreatedAdvances P "&_
					 "Where C.ReceivableNumber = P.CreatedAdvanceNo and C.CreatedTransno = "&iCreatedTransNo&" and  "&_
					 "P.ActualVoucherAmount >= P.AdvanceAdjusted and C.AdjustType is NOT NULL "
			objRs.Open sQuery,con
			IF Not objRs.EOF Then
				iAdjAdvTot = objRs(0)
			End if
			objRs.Close
		End IF

		IF CStr(iAdvTot) <> "0" Then
			sQuery = "Select Count(1) From Acc_T_CreatedPybleAdjDet C,Acc_T_CreatedAdvances P "&_
					 "Where C.PayablesNumber = P.CreatedAdvanceNo and C.CreatedTransno = "&iCreatedTransNo&" and  "&_
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

		'IF CStr(sErrCheck) = "Y" Then 'Error is Found
		'	Con.RollbackTrans
		'	Response.Redirect "AccErrDisp.asp?BookCode="&sVouCode
		'	Response.End
		'End IF

	End IF
	'**************** End of To Checking adjusted Bill is Not Been Adjusted By Some Other **********************
	
	
	'Response.Write "<p> Testing Completed.<br><BR>"
	'Response.End 
	
	'-----------------------PROCESS ENTRY NODES FOR CC/ANAL NODES-------------
	FOR EACH EntryNode IN Root.childNodes
	    Response.Write "<p>Welcome"
		IF  EntryNode.nodeName="Entry" THEN
			sEntryno=EntryNode.Attributes.getNamedItem("No").value
			sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
			sAmount=EntryNode.Attributes.getNamedItem("Amount").value
			sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value



			sExp="//Entry[@No='"&sEntryno&"']/AccHead"
			set tempNode=Root.selectNodes(sExp)
			sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

			set nodANL=oDOM.createElement("Root")
			nodANL.setAttribute "TransNo",iCreatedTransNo
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
			
			sQuery = "Delete from Acc_T_CreatedVoucherCCDet where CreatedtransNo = "& iOldCrTransNo  &" and VoucherEntryNumber = "& sEntryno
			Response.Write "<p>"& sQuery
			con.execute sQuery
			
            sQuery = "Delete from Acc_T_CretedVoucherAHDet where CreatedtransNo = "& iOldCrTransNo   &" and VoucherEntryNumber = "& sEntryno
            Response.Write "<p>"& sQuery
			con.execute sQuery
			
			Response.Write "XML = "&  nodANL.xml

			if bAddFlag then
			   Dim adoConn
			  ' Set adoConn = Server.CreateObject("ADODB.Connection")
			  ' adoConn.ConnectionString = con
			  ' adoConn.CursorLocation = 3
			  ' adoConn.Open
			   sQuery = "Proc_VouCCANALUpdate"
			   Dim adoCmd
			   Set adoCmd = Server.CreateObject("ADODB.Command")
			   Set adoCmd.ActiveConnection =Con
			   adoCmd.CommandText = sQuery
			   adoCmd.CommandType = 4 'adCmdStoredProc
			   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)
			   'Dim adoRS
			   'Set adoRS = 
			   adoCmd.Execute()
			end if
		end if
		bAddFlag = false
	NEXT

	'================== Contra Entry Details Starts Here =====================================================


	

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
						sContraCheck ="T"
							iSno = iSno + 1
							sQuery = "Select BookCode,BookNumber,BookName From Acc_R_ApplicableAccountHeads Where "&_
									 "BookAccountHead = "&iEntAccHead&" and Useable = '0' And OUDefinitionID = '"&iHeadUnit&"' "
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
									Response.Write "<p> NoChFlag =  "&  sNoChFlag  & "<p>"
									IF sNoChFlag = True Then 'Copied from Sangeeth
									    If trim(sContraCheck) = "T"  Then
		                                    sNewVouchNo  = sCrVouNo
			                                sNewAccVouNo = iVouNo
		                                Else
                                		    
		                                    IF strcomp(sVouType,"D")=0 THEN
			                                    sQuery = "select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
					                                     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sToBookCode&"' and BookNumber= "&sToBookNo
		                                    ELSE
			                                    sQuery = "select CrSeriesNo,CrSeriesCode,CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
					                                     "OUDefinitionID='"&sOrgId&"' and BookCode='"&sToBookCode &"' and BookNumber= "&sToBookNo
		                                    END IF

		                                    Response.Write sQuery &"<br><br>"
		                                    objRs.open sQuery,con
			                                    iSeriesNo=objRs(0)
			                                    iSeriesCode=objRs(1)
			                                    iCrSeriesNo=objRs(2)
			                                    iCrSeriesCode=objRs(3)
		                                    objRs.close()

		                                    sNewAccVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
		                                    sNewVouchNo  = GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sVoucDate)
		                                End If
									
									Else
										sNewVouchNo = sCrConVouNo
										sNewAccVouNo = sConVouNo
									End IF
									sQuery = "Select isNull(Max(CreatedTransNo),0) + 1 From Acc_T_CreatedVoucherHeader "
									objRs.open sQuery,con
										iCrTransNo = objRs(0)
									objRs.close()

									sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
									objRs.open sQuery,con
										iNewTransNo = objRs(0)
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

	'Con.CommitTrans ' blocked on sep 29,2011
	'Response.End

	Response.Write "Contra Check " & sContraCheck &"<br>"

	IF CStr(sContraCheck) = "T" Then
		Set objDom = Server.CreateObject("Microsoft.XMLDOM")
		objDom.Load server.MapPath("../Temp/Transaction/"&Session.SessionID&"-Voucher_Bank.xml")
		Set Root=objDom.documentElement

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
				 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,TDSGroupID) values"&_
				 "("&iCrTransNo&",'"&iHeadUnit&"','"&sToBookCode&"',"&sToBookNo&",'"&sNewTransType&"',"&_
				 "NULL,"&sAccHeadCode&",'"&sNewVouchNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
				 ",'"&sPayTo&"','"&sNewVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"',"&iTdsGroupID&")"

		Response.Write sQuery &"<BR><BR>"
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

		Response.Write sQuery &"<BR><BR>"
		Con.Execute sQuery

	'-----------------------------------------------------------------------------------------------------------------------------
		sQuery = "Update Acc_T_VoucherHeader set ContraTransactionNumber  = "&iNewTransNo&" where TransactionNumber = "& iTransNo&" "
		Response.Write "Update="&sQuery &"<BR><BR>"
		Con.Execute sQuery
		'-----------------------------------------------------------------------------------------------------------------------------

		'Response.Write "<b>"&sVouCode&"-"&sToBookCode&"</b>"
		If trim(sToBookCode) = "02" then

			if objfs.FileExists(Server.MapPath("../temp/transaction/"&Session.SessionID&"-BankInsDet.xml")) then
				oDOMNew.Load server.MapPath("../temp/transaction/"&Session.SessionID&"-BankInsDet.xml")
			
				Set RootNew = oDOMNew.documentElement
				sExp = "//BankInstrumentDet"
				Set tempNode = RootNew.selectNodes(sExp)

				sQuery = "Delete from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo  = "&iCreatedTransNo&" "
				Response.Write "<BR><BR>" & sQuery
				con.execute(sQuery)
			
				iCtr = 1
	

				For i = 0 to tempNode.length -1
					IF tempNode.length <> 0 Then

						sBKInsType	= tempNode.Item(0).Attributes.getNamedItem("InsType").Value
						sBkInsNo		= tempNode.Item(0).Attributes.getNamedItem("InsNo").Value
						sBkInsDate	= tempNode.Item(0).Attributes.getNamedItem("InsDate").Value
						iBkInsAmt	= tempNode.Item(0).Attributes.getNamedItem("InsAmt").Value
						sBkPayableat	= tempNode.Item(0).Attributes.getNamedItem("PayAt").Value

						sBkDrawnOn	= tempNode.Item(0).Attributes.getNamedItem("DrawnOn").Value
						sOption     = tempNode.Item(0).Attributes.getNamedItem("Option").value
						sAction		= tempNode.Item(0).Attributes.getNamedItem("Action").value
						iBankInsEntNo	= tempNode.Item(0).Attributes.getNamedItem("BankInsEntNo").value
					End IF


				'**************************Newly Added by Maheswari on 17th June 2008********************************

					IF trim(iBankInsEntNo) <> "0" then
						sTemp	  = split(sBkInsNo,"-")
						iEntNo	  = sTemp(0)
						iInsEntNo = sTemp(1)
						sBkInsNo  = sTemp(2)

					End If
					'Response.Write "sInsNo="&sInsNo  &"<BR>"
					'Response.Write "Option="&sOption&"<BR>"

					If trim(sOption) = "Y"  then
						IF trim(sAction) = "C" then
							sQuery = "Update Acc_R_BankInstrumentUsage set Status = 'C' where  CreatedTransNo  = "&iCreatedTransNo &" "&_
								 " and EntryNo = "&iEntNo&"  and InstrumentEntryNo = "&iInsEntNo&" and InstrumentNo = '"&sBkInsNo &"' "
							Response.Write "<BR><BR>" & sQuery
							con.execute(sQuery)
						ElseIF trim(sAction) = "R" then
							sQuery = "Update Acc_R_BankInstrumentUsage set CreatedTransNo  = 0,Status = 'N' where  CreatedTransNo  = "&iCreatedTransNo&" "&_
									"and EntryNo = "&iEntNo&" and InstrumentEntryNo = "&iInsEntNo&" and InstrumentNo = '"&sBkInsNo&"' "
							Response.Write "<BR><BR>" & sQuery
							con.execute(sQuery)
						End IF
						'Response.Write "<B>Insertion in Acc_T_CreatedVoucherInstrumentDet</B>"&"<BR>"
						sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
								 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,BankInstrumentEntryNo,InstrumentEntryNo1,InstrumentAmount,ClearedOn)"&_
								 " Values ("&iCrTransNo&","&iCtr&",'"&sBkInsType&"','"&sBkInsNo&"',convert(datetime,'"&sBkInsDate&"',103),"&_
								 "'"&sBkPayableat&"','"&sBkDrawnOn&"',"&iEntNo&","&iInsEntNo&","&iBkInsAmt&"," & dtClearedOn & " ) "
						Response.Write "1="&sQuery &"<BR><BR>"
						con.execute(sQuery)

					Else
						sQuery = "Update Acc_R_BankInstrumentUsage set CreatedTransNo  = "&iCreatedTransNo&",Status = 'U' where  InstrumentNo = '"&sBkInsNo&"' "
								'" and EntryNo = "&iEntNo&" and InstrumentEntryNo = "&iInsEntNo&"  "
						Response.Write "<BR><BR>" & sQuery
						con.execute(sQuery)

						'Response.Write "<B>Insertion in Acc_T_CreatedVoucherInstrumentDet</B>"&"<BR>"
						sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
								 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,InstrumentAmount,ClearedOn)"&_
								 " Values ("&iCrTransNo&","&iCtr&",'"&sBkInsType&"','"&sBkInsNo&"',convert(datetime,'"&sBkInsDate&"',103),"&_
								 "'"&sBkPayableat&"','"&sBkDrawnOn&"',"&iBkInsAmt&"," & dtClearedOn & " ) "
						Response.Write "2="&sQuery &"<BR><BR>"
						con.execute(sQuery)
					End IF
				Next
			End IF
		End if 'if objfs.FileExists(Server.MapPath("../temp/transaction/"&Session.SessionID&"-BankInsDet.xml")) then
'===============================
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
				'Response.Write   "<BR>****"& sAmount &"-"& iTdsAmount & "***<BR>"
				sAmount = Cdbl(sAmount) - Cdbl(iTdsAmount)
				'Response.Write " sAmount = "& sAmount &"<BR><BR>"
				sTDSNarr = "TDS Entry"
				TDSFlag = "Y"

	'---------PROCESS THE CHILD NODES OF ENTRY NODE FOR DETAIL TABLE UPDATION----
				FOR EACH HeaderNode IN EntryNode.childNodes
					IF HeaderNode.nodeName="AccHead" THEN
							sAccCode=HeaderNode.Attributes.getNamedItem("No").value
							sAccType=HeaderNode.Attributes.getNamedItem("Type").value
					END IF 'End of Check for Account head Node
					IF HeaderNode.nodeName="TDS" THEN
						'Response.Write HeaderNode.nodeName &"<BR>"
						iAccHeadNo	 = HeaderNode.getAttribute("AccHeadCode")
						iTDSEntryAmt = HeaderNode.getAttribute("TDSAmount")
						iTDSEntryPer = HeaderNode.getAttribute("TDSPercentage")
						iPayRecAmt	 = HeaderNode.getAttribute("PayRecAmount")
						sFormula	 = HeaderNode. getAttribute("Formula")

						sQuery =" Insert into Acc_T_CreatedVoucherDetails(CreatedTransNo,AccountingUnit,VoucherEntryNumber,"&_
								" AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"&_
								" VoucherNarration,Amount,TransCrDrIndication,TDSonAmount,PackCode,TDSFlag) "&_
								" values("&iCrTransNo&",'"&sAccUnit&"',"&sVouEntryno&","&iAccHeadNo&",NULL,NULL,NULL, "&_
								" '"&sTDSNarr&"',"&iPayRecAmt&",'"&sEntryType&"',"&iTDSEntryAmt&","&sGrpId&",'"&TDSFlag&"')"
								Response.Write "<BR>TDS="& sQuery & "<BR><BR>"
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
					sQuery = "insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
					sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
					sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
					sQuery = sQuery& iCrTransNo&",'"&sAccUnit&"'"
					sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
					sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&iTdsAmount&","&iTdsPer&")"

					Response.Write "T2="&sQuery& "<BR><BR>"
					Con.Execute sQuery


					sQuery = "insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
					sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
					sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSOnAmount,TDSPercentage) values ("
					sQuery = sQuery& iNewTransNo&",'"&sAccUnit&"'"
					sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
					sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&", "&iTdsPer&")"
					Response.Write "T3="&sQuery& "<BR><BR>"
					con.execute(sQuery)

					'This will Remove the Extra Entry Entered in GL_Transactions for the Bank/Cash Receipt/Payment entry

					'sQuery = "Delete  Acc_t_gltransactions where oudefinitionid = '"&sOrgId&"' "&_
					'		 "and bookcode = '"&sToBookCode&"' and booknumber = "&sToBookNo&" and transactiontype = '"&sNewTransType&"' "&_
					'		 "and accounthead = "&sAccCode&"  and TransactionNumber = "&iNewTransNo&" "
					'Response.Write sQuery &"<BR><BR>"
	'===========================================================================================================
					'Con.Execute sQuery
				END IF
			End IF
		Next 'Entry Node For Details Tabel Insertion

	'If Both From and To Accounthead is marked as summary then the Debit/Credit value
	'will chaged to from account head

		Dim iRecAff
		sQuery = "Select SummaryPosting From VwOrgGLHeads Where Accounthead IN ("&iFrmAccCode&","&iToAccCode&") "&_
				 "and OUDefinitionID = '"&sOrgId&"' and SummaryPosting = '1' "

		'Response.Clear
		Response.Write sQuery

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
			sQuery = "Update Acc_T_GLTransactions Set Amount = Amount + "&dTotal&" Where BookCode = '"&sToBookCode&"' and  "&_
					 "BookNumber = "&sToBookNo&" and TransactionNumber = 0 and TransCrDrIndication = '"&sFrmCrDrTy&"' and  "&_
					 "AccountHead = "&iFrmAccCode&" and Amount - "&dTotal&" >= 0 and Month(VoucherDate) = "&iMon&" and "&_
					 "Year(VoucherDate) = "&iYear&" "

			Response.Write sQuery &"<br>"
			Con.Execute sQuery,iRecAff
			Response.Write iRecAff &"<br>"

			IF CStr(iRecAff) <> "0" Then
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount + "&dTotal&" Where BookCode = '"&sVouCode&"' and  "&_
						 "BookNumber = "&sBookNo&" and TransactionNumber = 0 and TransCrDrIndication = '"&sFrmCrDrTy&"' and "&_
						 "AccountHead = "&iFrmAccCode&" and Month(VoucherDate) = "&iMon&" and Year(VoucherDate) = "&iYear&" "
				Response.Write sQuery &"<br>"
				Con.Execute sQuery,iRecAff
				Response.Write iRecAff &"<br>"
			End IF

			sQuery = "Select TransactionNumber From Acc_T_GLTransSummaryBreakup Where TransactionNumber = "&iTransNo&" "
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iConChkCnt = 1
			Else
				iConChkCnt = 0
			End IF
			objRs.Close

			Response.Write "<br><br><br>"
			Response.Write iFrmTransNo &"<br><br><br>"
			'''blocked by S.Maheswari on 20th feb 2009
			'IF Cstr(iConChkCnt) = "0" Then
			'	sQuery = "INSERT INTO Acc_T_GLTransSummaryBreakup (OUDefinitionID, AccountHead, BookCode, BookNumber,  "&_
			'			 "TransactionType, TransactionNumber, VoucherDate, VoucherEntryNumber, VoucherNo,  "&_
			'			 "VoucherAmount, TransCrDrIndi) "&_
			'			 "VALUES ('"&sOrgId&"', "&iFrmAccCode&", '"&sVouCode&"', "&sBookNo&",  "&_
			'			 "'"&sFrmTransTy&"', "&iFrmTransNo&", Convert(datetime,'"&sVoucDate&"',103), 1, "&iFrmTransNo&", "&dTotal&", '"&sFrmCrDrTy&"') "

			'	Response.Write sQuery &"<br><br>"
			'	Con.Execute sQuery

			'	sQuery = "Delete Acc_T_GLTransSummaryBreakup Where TransactionNumber = "&iNewTransNo&" and VoucherEntryNumber = 1 "
			'	Con.Execute sQuery

			'End IF


		End IF
		''''start sopy from sangeeth
		Response.Write "*************"&sNorToContra &"<BR>"
		IF CStr(sNorToContra) = "Y" Then 'If the Voucher Type is Been Amendmed Fron a Normal To Contra then Execute the following
			sQuery = "Update Acc_T_GLTransactions Set Amount = Amount + "&dTotal&" Where BookCode = '"&sVouCode&"' and  "&_
			"BookNumber = "&iFrmBookNo&" and TransactionNumber = 0 and TransCrDrIndication = '"&sFrmCrDrTy&"' and  "&_
			"AccountHead = "&iFrmAccCode&" and Month(VoucherDate) = "&iMon&" and "&_
			"Year(VoucherDate) = "&iYear&" and OUDefinitionID = '"&sOrgId&"'"

			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery

			sQuery = "Update Acc_T_GLTransactions Set Amount = Amount + "&dTotal&" Where BookCode = '"&sToBookCode&"' and  "&_
			"BookNumber = "&sToBookNo&" and TransactionNumber = 0 and TransCrDrIndication = '"&sVouType&"' and  "&_
			"AccountHead = "&iToAccCode&" and Month(VoucherDate) = "&iMon&" and "&_
			"Year(VoucherDate) = "&iYear&" and OUDefinitionID = '"&sOrgId&"'"

			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery


			IF Cstr(sFrmCrDrTy) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount + "&dTotal&"  "&_
						 "Where AccountHead = "&iFrmAccCode&" And OUDefinitionID = '"&sOrgId&"' "&_
						 "And MonthYear = '"&sMonYear&"' "
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount + "&dTotal&"  "&_
						 "Where AccountHead = "&iFrmAccCode&" And OUDefinitionID = '"&sOrgId&"' "&_
						 "And MonthYear = '"&sMonYear&"' "
			End IF

			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery

			IF CStr(sVouType) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount + "&dTotal&"  "&_
						 "Where AccountHead = "&iToAccCode&" And OUDefinitionID = '"&sOrgId&"' "&_
						 "And MonthYear = '"&sMonYear&"' "
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount + "&dTotal&"  "&_
						 "Where AccountHead = "&iToAccCode&" And OUDefinitionID = '"&sOrgId&"' "&_
						 "And MonthYear = '"&sMonYear&"' "
			End IF

			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
		End IF
		'''end copy from Sangeeth


		objDom.Load server.MapPath("../Temp/Transaction/"&Session.SessionID&"-Voucher_Bank.xml")

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

		'Response.Clear


		sRetValChk = sRevConChk&":"&iTransNo&":"&iNewTransNo

		'objDom.save server.MapPath("../xmldata/voucher/"&iCrTransNo&".xml")

	End IF 'Contra Check
	'================== Contra Entry Details Ends Here =====================================================
	'------------------END OF PROCESSING ENTRY NODES--------------------------
	sErrCheck = "T"
	'Con.CommitTrans
	'Con.RollbackTrans



	Set newElem  = oDOM.createAttribute("TransNo")
	newElem.value = iTransNo
	Root.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("VoucherNo")
	newElem.value = sVouNo
	Root.setAttributeNode(newElem)
	
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")
	

	IF CStr(sRetValChk) = "" Then
		sRetValChk = "F:0:0"
	End IF
	Dim dTotAdjAmt
	IF Cstr(iToBeCrAdvNo) = "" Then
		iToBeCrAdvNo = 0
	End IF
'**************Updating the Old Advance Numbers With the New Numbers ****************************\
	IF Cstr(sCheckAdjRootVal) = "Y"  and Cstr(iTobeCrAdvNo) <> "0" Then
		sExp = "//AdjDet"
		Set tempNode = AdjRoot.selectNodes(sExp)
		For iCtr = 0 To tempNode.length - 1
			IF CStr(tempNode.Item(iCtr).Attributes.getNamedItem("ValFrm").Value) = "CRADJADV" Then
				sQuery = "Update ACC_T_CreatedAdvanceAdj Set CreatedAdvanceNo = "&iToBeCrAdvNo&" Where "&_
						 "CreatedTransNo = "&tempNode.Item(iCtr).Attributes.getNamedItem("AdjCrTrNo").Value
				Response.Write sQuery
				Con.Execute sQuery
			ElseIF CStr(tempNode.Item(iCtr).Attributes.getNamedItem("ValFrm").Value) = "Pay" Then
				sQuery = "Update Acc_T_CreatedPybleAdjDet Set PayablesNumber = "&iToBeCrAdvNo&" "&_
						 "Where CreatedTransNo = "&tempNode.Item(iCtr).Attributes.getNamedItem("AdjCrTrNo").Value
				Con.Execute sQuery
			ElseIF CStr(tempNode.Item(iCtr).Attributes.getNamedItem("ValFrm").Value) = "Rec" Then
				sQuery = "Update Acc_T_CreatedRcvbleAdjDet Set ReceivableNumber = "&iToBeCrAdvNo&" "&_
						 "Where CreatedTransNo = "&tempNode.Item(iCtr).Attributes.getNamedItem("AdjCrTrNo").Value
				Con.Execute sQuery
			End IF

			sQuery = "Select isNull(SUM(AmountAdjusted),0) From ACC_T_CreatedAdvanceAdj Where  "&_
					 "CreatedAdvanceNo = "&iToBeCrAdvNo
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				dTotAdjAmt = objRs(0)
			Else
				dTotAdjAmt = 0
			End IF
			objRs.Close

			sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = "&dTotAdjAmt&" Where CreatedAdvanceNo = "&iToBeCrAdvNo
			Con.Execute sQuery
			sQuery = "Update Acc_T_AdvancePayments Set AdvanceAdjusted = "&dTotAdjAmt&" Where CreatedAdvanceNo = "&iToBeCrAdvNo
			Con.Execute sQuery

		Next
	End IF

'**************Updating the Old Advance Numbers With the New Numbers ****************************

	CreateVou = sRetValChk

End Function

Function GetBookReceiptNoCrVou(sBookCode,sUnit,iBookNo,sVouType)
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

		GetBookReceiptNoCrVou  = iReciptNo
		'Response.Clear
		Response.Write "Receipt Number is " & iReciptNo &"<br><br><br>"
		'Response.End
	End Function
	
    Function GetBookReceiptNoVou(sBookCode,sUnit,iBookNo,sVouType)
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

		GetBookReceiptNoVou = iReciptNo
		'Response.Clear
		Response.Write "Receipt Number is " & iReciptNo &"<br><br><br>"
		'Response.End
	End Function
	

%>




