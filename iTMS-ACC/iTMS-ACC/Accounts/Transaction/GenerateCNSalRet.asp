<%
Function CreateCNSalRet(iCrVouNo,iVouNo,sOldBkNo,sOldVouDate)
	'XML DOM Variables
	Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot
	dim sParCode,sParSubType,sParType,oNodTemp,sQuery,nodANL,newElem,dEachItmVal
	dim EntryNode,HeaderNode,sAddtional,iSno,sAmount,sTemp,sNarration,sAccount
	dim sOrgId,sBookNo,sVouType,sVouName,sVoucDate,sAccUnit,sAccHeadCode,iTemp
	dim sVouCode,sApprove,dTotal,sTransType,sDocType,sVouStatus,iCounter,iCtr
	dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo,sSalType,ActualTransNo
	dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
	dim dAdjAmtTotal,sAppTy,sAppBy,iCrTransNo,objfs,sCallTy,dTotalVouchVal
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,sTaxMode
	Dim sPayIns,iRecNo,sExp,TempNode,iAccPayNo,iCrRecNo,dAdvAmt,dTBDiffAmt,TaxCalNode
	Dim dHdVal,dDetVal,dTaxDebVal,dTaxCreVal,dAccVal,sItemDetails,iSalTrNo
	Dim sTaxTrans,sTempNoSer

	sVouStatus="010104" 'Accounted
	sVouCode="07"
	sVouType="C"


	set objRs  = server.CreateObject("adodb.recordset")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.load server.MapPath("../temp/transaction/Voucher Entry_CNAMD_"&Session.SessionID&".xml")
	Set oNodRoot=oDOM.documentElement

	For each oNodTemp in oNodRoot.childNodes
		IF oNodTemp.nodeName="Header" then
			For Each oNodEntry in  oNodTemp.childNodes
				if oNodEntry.nodeName="Organization" then
					sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				end if
				if oNodEntry.nodeName="Book" then
					sBookNo=oNodEntry.Attributes.Item(0).nodeValue
				end if
				if oNodEntry.nodeName="Party" then
					sParType=oNodEntry.Attributes.Item(0).nodeValue
					sParSubType=oNodEntry.Attributes.Item(1).nodeValue
					sParCode=oNodEntry.Attributes.Item(3).nodeValue
				end if
				if oNodEntry.nodeName="SaleInvoice" then
					sPurInvNo=oNodEntry.Attributes.Item(0).nodeValue
					sPurInvDt=oNodEntry.Attributes.Item(1).nodeValue
					sAppTy = oNodEntry.Attributes.Item(3).nodeValue
					sAppBy = oNodEntry.Attributes.Item(4).nodeValue
					ActualTransNo = oNodEntry.Attributes.getNamedItem("CrTransNo").nodeValue
					iSalTrNo = oNodEntry.Attributes.getNamedItem("CrTransNo").nodeValue
				end if
				if oNodEntry.nodeName="SalesType" then
					sSalType=oNodEntry.Attributes.Item(0).nodeValue
				end if
			next
		end if
		
		if oNodTemp.nodeName="Details" then
			set oNodDeatils=oNodTemp
		end if
		
		if oNodTemp.nodeName="TaxDetails" then
			set oNodTaxRoot=oNodTemp
			'dTotal=oNodTemp.Attributes.Item(0).nodeValue
			dTotalVouchVal=Trim(oNodTemp.Attributes.Item(0).nodeValue)
			dTotal=oNodTemp.Attributes.Item(0).nodeValue
		end if
	next

	sExp = "//Voucher"
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		IF CStr(iCrVouNo) <> "" Then
			TempNode.Item(0).Attributes.getNamedItem("CreatedVouNo").Value = Cstr(iCrVouNo)
		End IF
		TempNode.Item(0).Attributes.getNamedItem("VouNo").Value = iVouNo
	End IF

	sTransType="CNR"
	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue
	

	sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCrTransNo=objRs(0)
	objRs.Close

	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close
	
	sTempNoSer = GetNewNoSer(sBookNo,sOldBkNo,sVoucDate,sOldVouDate,sVouCode,sOrgId,sVouType)
	sTemp = Split(sTempNoSer,"||")
	IF CStr(sTemp(0)) <> "NA" Then
		iVouNo = sTemp(0)
	End IF
	
	IF CStr(sTemp(1)) <> "NA" Then
		iCrVouNo = sTemp(0)
	End IF
	

	sQuery = "Insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,BankInstrumentNo) values"
	sQuery = sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotalVouchVal
	sQuery = sQuery&",'Sales Return ','SR','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"',"&iSalTrNo&" )"

	Response.Write sQuery& "<BR>"
	con.execute(sQuery)

	sQuery = "Insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus,BankInstrumentNo) values"
	sQuery = sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"','"&iCrVouNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery = sQuery&",'Sales Return ','"&sVouType&"','SR',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"',"&iSalTrNo&")"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)

	For each EntryNode in oNodDeatils.childNodes
		sEntryno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
		dQty=EntryNode.Attributes.Item(3).nodeValue
		sUOM=EntryNode.Attributes.Item(4).nodeValue
		dBasicAmount="0"
		dRate=EntryNode.Attributes.Item(6).nodeValue
		dDisPer="0"
		dDisAmount="0"

		sEntryType="D"
		sNarration=""
		if	CDbl(sAmount) > 0 then
			for each HeaderNode in EntryNode.childNodes
				if HeaderNode.nodeName="AccHead" then
						sAccCode=HeaderNode.Attributes.Item(0).nodeValue
						sAccType=HeaderNode.Attributes.Item(4).nodeValue
				end if 'End of Check for Account head Node
			next 'End of Entry Node Loop

			sQuery = "Insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
			sQuery = sQuery& iCrTransNo&",'"&sOrgId&"'"
			sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"
			Response.Write sQuery& "<BR><BR>"
			con.execute(sQuery)
			
			sAmount = CDbl(sAmount) + CDbl(dEachItmVal)
			IF CDbl(iCtr) = 1 Then
				Response.Write dEachItmVal &"==========<br>"
				sAmount = CDbl(sAmount) + CDbl(dDiffVal)
				iCtr = 2
			End IF

			sQuery = "Insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
			sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
			sQuery = sQuery& iTransNo&",'"&sOrgId&"'"
			sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"
			
			Response.Write sQuery& "<BR><BR>"
			con.execute(sQuery)
		End IF		
	Next'End of Voucher Node Loop

	iSno=1
	For each EntryNode in oNodTaxRoot.childNodes
		iCatCode=EntryNode.Attributes.Item(0).nodeValue
		iTaxCode=EntryNode.Attributes.Item(1).nodeValue
		sTaxMode=EntryNode.Attributes.Item(2).nodeValue
		dTaxPer=EntryNode.Attributes.Item(4).nodeValue
		dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
		sAccCode=EntryNode.Attributes.Item(6).nodeValue
		
		'IF CStr(iCatCode) = "0" and CStr(iTaxCode) = "0" Then
		'	dTaxAmount = EntryNode.Attributes.Item(4).nodeValue
		'End IF
		 
		if	CDbl(dTaxAmount) > 0 then
			sEntryType = "D"
		else
			sEntryType = "C"
		End if
		
		dTaxAmount = Abs(dTaxAmount)
		
		if	CDbl(dTaxAmount) > 0 then
			if sTaxMode="P" then
				sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
				Response.Write sQuery& "<BR><BR>"		 
				con.execute(sQuery)
				
				sQuery = "INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
				Response.Write sQuery& "<BR><BR>"
				con.execute(sQuery)	
			else
				sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
				Response.Write sQuery& "<BR><BR>"
				con.execute(sQuery)
				
				sQuery = "INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
				Response.Write sQuery& "<BR><BR>"
				con.execute(sQuery)	
			end if
		end if
		iSno=cint(iSno)+1
	Next

	Response.Clear
	sQuery = "select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "&ActualTransNo&" "&_
			 "And AmountReceivable > AmountReceived and AmountReceivable - AmountReceived >= "&dTotal&" "
				 
	Response.Write sQuery &"<br><br>"
	objRs.open sQuery,con
		
	If Not objRs.EOF Then
		iPayableNo = objRs(0)
	Else
		iPayableNo = 0
	End If
	objRs.Close

	if CDbl(iPayableNo) <> 0 THEN
		sPayIns = "N"
		iRecNo = iPayableNo
		sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, ReceivedOn, "&_
				 "AmountReceived) VALUES ("&iPayableNo&", "&iCrTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&") "
					 
		Response.Write sQuery &"<br><br>"
		con.execute(sQuery)
	else
	'----------------------------- CREATE PAYABLE IF RECEIVABLE DETAILS DOESNOT EXIST ------------------
		sPayIns = "Y"
		sQuery = "Select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close
		sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber,CreatedTransNo, OUDefinitionID, "&_
				"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
				"PartyBillDate, AmountPayable, AmountPaid) values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
					
		Response.Write sQuery &"<br><br>"
			
		con.execute(sQuery)
	end if

	IF CStr(sPayIns) = "Y"  Then
		dAdjAmtTotal = 0
		sExp = "//AdvanceDetails/Advance[@AmountToAdj!=0]"
		Set TempNode = oNodRoot.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			For iCounter = 0 To TempNode.length - 1
				iAccPayNo = TempNode.ITem(iCounter).Attributes.getNamedItem("CreatedTransNo").value
				dAdvAmt = Trim(TempNode.ITem(iCounter).Attributes.getNamedItem("AmountToAdj").value)
				dAdvAmt = CDbl(dAdvAmt)
				dAdjAmtTotal = Cdbl(dAdjAmtTotal + dAdvAmt)
				sQuery = "Select DRCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&iAccPayNo&" "
				objRs.Open sQuery,Con
				IF Not objRs.EOF Then
					iCrRecNo = objRs(0)
				Else
					iCrRecNo = 0
				End IF
				objRs.Close
				sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
						 "ReceivedOn,AmountReceived) Values ("&iCrRecNo&","&iCrTransNo&","&_
						 "getdate(),"&dAdvAmt&")"
							 
				Response.Write sQuery &"<br><br>"		 
				Con.Execute sQuery
			Next
			sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn, "&_
					 "AmountPaid) VALUES ("&iPayableNo&", "&iCrTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dAdjAmtTotal&") "
					 
			Con.Execute sQuery
			Response.Write sQuery &"<br><br>"
		End IF
	End IF

	sQuery = "Select A.ReceivableNumber From Acc_T_Receivables A,Acc_T_CreatedReceivables C "&_
			 "Where A.CreatedReceivable = C.ReceivableNumber and C.CreatedTransNo = "&iSalTrNo&" "
	objRs.open sQuery,con
		
	If Not objRs.EOF Then
		iPayableNo = objRs(0)
	Else
		iPayableNo = 0
	End If
	objRs.Close

	if CDbl(iPayableNo) <> 0 THEN
		sQuery = "insert into Acc_T_RcvblAdjustmentDetails (ReceivableNumber, RecdByTransactionNo, ReceivedOn, AmountReceived) "&_
				 "values("&iPayableNo&","&iTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&")"
				 
		Response.Write sQuery &"<br><br>"
		con.execute(sQuery)
			
	else
	'----------------------------- CREATE PAYABLE IF RECEIVABLE DETAILS DOESNOT EXIST ------------------
		sQuery = "select isnull(max(PayablesNumber),0)+1 from Acc_T_Payables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close

		sQuery = "Select PayablesNumber From Acc_T_CreatedPayables Where CreatedTransNo = "&iCrTransNo
		objRs.Open sQuery,Con
		
		IF Not objRs.EOF Then
			iCrRecNo = objRs(0)
		Else
			iCrRecNo = 0
		End IF
		objRs.Close

		sQuery="INSERT INTO Acc_T_Payables(PayablesNumber,TransactionNumber, OUDefinitionID, "&_
				"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
				"PartyBillDate, AmountPayable, AmountPaid, CRCreatedPayable) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0, "&iCrRecNo&" )"
		con.execute(sQuery)
	end if

	
	'Con.RollbackTrans
	'Response.End
	
	sExp = "//Entry"
	Set TaxCalNode = oNodRoot.selectNodes(sExp)
	For iCtr = 0 To TaxCalNode.length - 1
		sItemDetails = sItemDetails&", "&TaxCalNode.Item(iCtr).Attributes.getNamedItem("PayTo").Value
	Next
		
	sItemDetails = Trim(Mid(sItemDetails,2,200))
		
	sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sItemDetails&"' Where  "&_
			 "TransactionNumber = "&iTransNo
	Con.Execute sQuery
		
	sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sItemDetails&"' Where  "&_
			 "TransactionNumber = "&iTransNo
	Con.Execute sQuery
	
	'oNodRoot.Attributes.Item(0).nodeValue= iCrTransNo
	'oNodRoot.Attributes.Item(1).nodeValue= iVouNo
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

	IF objfs.FileExists(Server.MapPath("../temp/transaction/Voucher Entry_CNAMD_"&Session.SessionID&".xml")) then
		'objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher Entry_CNAMD_"&Session.SessionID&".xml"))
	End IF
End Function
%>
