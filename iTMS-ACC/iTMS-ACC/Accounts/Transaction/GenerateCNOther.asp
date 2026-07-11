<%
Function CreateCNOther(iCrVouNo,iVouNo,sOldBookNo,sOldVouDate,sChkGj)
	Response.Clear
	Dim oDOM,nodHeader,Root,objRs,sQuery,EntryNode,HeaderNode,nodANL,newElem
	dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp,sVouName,sExp
	dim sOrgId,sBookNo,sVouType,sVouCode,sApprove,sVoucDate,sAccUnit,sAccHeadCode
	dim dTotal,sTransType,dCRAmt,dDRAmt,sAccType,sAccCode,sEntryType,sEntryno,iTransNo
	dim sDocType,sVouStatus,sParType,sParSubType,sParCode,iPayableNo,sApprover,sSelInvNo
	dim iSeriesNo,iSeriesCode,sPayTo,bAddFlag,dTdsAmt,sTdsPer,Checknode,iCrTransNo
	Dim dTdsPer,iCrPayNo,dTBDiffAmt,sTempNoSer,sNoSerTy,adoCmd
	dim sParName,sPayAt,iVouAmt,sCrDr,iEntryNo,sTrCrDr,objRs1,iAccHead,sNarr,sEntAmt
	
	set objRs1 = Server.CreateObject("ADODB.Recordset")
	
	sVouCode = "07"
	sVouName = "CN"
	sApprove = GetUserID()
	dTdsPer = 0
	sChkGj = "Y"
	sApprover = Request.Form("selUserID")
	sVouStatus="010104" 'Crearted For Accounting
	
	set objRs  = server.CreateObject("adodb.recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml")
	set Root=oDOM.documentElement

	sOrgId=Root.Attributes.getNamedItem("UnitNo").value
	sBookNo =Root.Attributes.getNamedItem("BookNo").value
	sVouType=Root.Attributes.getNamedItem("CRDR").value
	sVoucDate=Root.Attributes.getNamedItem("VouDate").value
	
	sTemp=Split(Root.Attributes.getNamedItem("PartyCode").value,"?")

	sParType=trim(sTemp(0))
	sParSubType=trim(sTemp(1))
	sParName = trim(sTemp(2))
	sParCode=trim(sTemp(3))

	Root.Attributes.getNamedItem("Approver").value=sApprove
	sPayTo=Replace(Root.childNodes(0).Attributes.getNamedItem("Payto").value,"'","''")

	FOR EACH EntryNode IN Root.childNodes
		dTotal=dTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
	NEXT

	sTransType=sVouName&"R"

	sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCrTransNo=objRs(0)
	objRs.Close
	
	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			 "PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			 "CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"&_
			 "("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			 "'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
			 ",'"&sVouType&"','OT',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

	Response.Write sQuery& "<BR>"
	con.execute(sQuery)
	
	FOR EACH EntryNode IN Root.childNodes
		sEntryno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(3).nodeValue
		sEntryType=EntryNode.Attributes.Item(1).nodeValue
		sAccUnit=EntryNode.Attributes.Item(4).nodeValue
	
		FOR EACH HeaderNode IN EntryNode.childNodes
			IF HeaderNode.nodeName="AccHead" THEN
					sAccCode=HeaderNode.Attributes.Item(0).nodeValue
					sAccType=HeaderNode.Attributes.Item(4).nodeValue
			END IF 'End of Check for Account head Node
			IF 	HeaderNode.nodeName="Narration" THEN
					sNarration=HeaderNode.text
			END IF 'End of Check for Narration Node
		NEXT
		sNarration = Replace(sNarration,"'"," ")
		
		sQuery = "Insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication) values ("
		sQuery = sQuery& iCrTransNo&",'"&sAccUnit&"'"
		sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"')"
			
		Response.Write sQuery& "<BR><BR>"
		con.execute(sQuery)
		
	Next
	
	sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
	objRs.open sQuery,con
		iCrPayNo = objRs(0)
	objRs.Close

	

	sQuery = "INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
			 "VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
			 " PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iCrPayNo&","&iCrTransNo&",'"&sOrgId&"',"&_
			 "convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",NULL,"&_
			 "NULL,"&dTotal&",0,'"&sPayTo&"')"
	con.execute(sQuery)
	
	
	IF CStr(sChkGj) = "Y" Then
		sQuery = "Select DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
				 "OUDefinitionID='"&sOrgId&"' and BookCode='08' and BookNumber= "&sBookNo
 
		objRs.open sQuery,con
		if not objRs.EOF then
			iSeriesNo=objRs(0)
			iSeriesCode=objRs(1)
		end if	
		objRs.close() 
		iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

		sQuery = "AccountDnCNGJVou"
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection =con
		adoCmd.CommandText = sQuery
		adoCmd.CommandType = 4 'adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@iCrTransNo",3,1,3,iCrTransNo)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@AccVouNo",201,1,30,iVouNo)
		adoCmd.Execute()
	End IF

	
End Function

%>
