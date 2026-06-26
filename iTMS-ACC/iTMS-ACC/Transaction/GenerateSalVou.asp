<%
Function CreateSalInv(sCrVouNo,iVouNo)
	Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot
	Dim oNodTemp,sQuery,sParCode,sParSubType,sParType,EntryNode,HeaderNode,nodANL,newElem,sExp
	dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp,sOrgId,sBookNo,sVouType
	dim sVouName,sVouCode,sApprove,sVoucDate,sAccUnit,dTotal,sTransType
	dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo,sDocType,sVouStatus
	dim sSalType,sRefernceNo,iCounter,dAdjAmtTotal,TempNode,dRndOff,CheckNode
	dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sSalInvNo,sSalInvDt
	Dim dRoundOff,sRoundoffHead,dNoofPack,dPackTy,dRatePer,iCrTransNo
	Dim sAppTy,sAppBy,sGroupCode,bAddFlag,sAdvNarr,dEntRndVal,objfs
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,sTaxMode
	dim iAdvTranNo,dAdvAmount,iCrAdvno,sAdjType,iCrPayNo,iCrPayable,sPayNarration,iCrPayableNo
	Dim iAgentCode,sCommType,dCommPer,dCommAmount,iCurrencyCode,sAgType,sAgParType
	Dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,iAdvNo,sNewNarr,dTBDiffAmt,sStr,sItemFlag
	Dim dInvWtRnd

    sItemFlag = false
	sVouCode="05"
	sVouType="D"
	sTransType="SJR"

	Set objRs  = server.CreateObject("adodb.recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")

	If objfs.FileExists(server.MapPath("../../Sales/xmldata/General.xml")) Then
	
		oDOM.Load server.MapPath("../../Sales/xmldata/General.xml")
		Set oNodRoot = oDOM.documentElement
		sExp = "//ROUNDOFF"
		Set oNodEntry = oNodRoot.Selectnodes(sExp)
		If oNodEntry.Length > 0 then
			sRoundoffHead = oNodEntry.Item(0).Attributes.Item(0).nodevalue
		else
			sRoundoffHead = "0"
		end if
	Else
		sRoundoffHead = "0"
	End IF

	oDOM.load server.MapPath("../temp/transaction/Voucher Amd_SAL_"&Session.SessionID&".xml")

	set oNodRoot=oDOM.documentElement
	sExp = "//Organization"
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		sOrgId =  TempNode.Item(0).Attributes.getNamedItem("OrgId").Value
	End IF

	for each oNodTemp in oNodRoot.childNodes
		if oNodTemp.nodeName="Header" then
			for Each oNodEntry in  oNodTemp.childNodes
				if oNodEntry.nodeName="Book" then
					sBookNo=oNodEntry.Attributes.Item(0).nodeValue
				end if
				if oNodEntry.nodeName="SalesType" then
					sSalType=oNodEntry.Attributes.Item(0).nodeValue
				end if
				if oNodEntry.nodeName="Party" then
					sParType=oNodEntry.Attributes.Item(0).nodeValue
					sParSubType=oNodEntry.Attributes.Item(1).nodeValue
					sParCode=oNodEntry.Attributes.Item(3).nodeValue
				end if
				if oNodEntry.nodeName="SaleInvoice" then
					sSalInvNo=oNodEntry.Attributes.Item(0).nodeValue
					sSalInvDt=oNodEntry.Attributes.Item(1).nodeValue
					sRefernceNo=oNodEntry.Attributes.Item(2).nodeValue
					sAppTy = Trim(oNodEntry.Attributes.Item(3).nodeValue)
					sAppBy = Trim(oNodEntry.Attributes.Item(4).nodeValue)
				end if
			next
		end if
		if oNodTemp.nodeName="AgentDetails" then
			set oNodAgent=oNodTemp
		end if

		if oNodTemp.nodeName="Details" then
			set oNodDeatils=oNodTemp
		end if
		if oNodTemp.nodeName="TaxDetails" then
			set oNodTaxRoot=oNodTemp
			dTotal=oNodTemp.Attributes.Item(2).nodeValue
			dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
		end if
		if oNodTemp.nodeName="AdvanceDetails" then
			set oNodAdvRoot=oNodTemp
		end if
	next

	sExp = "//Voucher"
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.getNamedItem("CreatedVouNo").Value = sCrVouNo 
		TempNode.Item(0).Attributes.getNamedItem("VouNo").Value = iVouNo
	End IF

	sExp = "//Details"
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		set oNodDeatils = TempNode.Item(0)
	End IF

	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

	Set newElem = oDOM.CreateElement("Tax")
	newElem.SetAttribute "CatCode","0"
	newElem.SetAttribute "TaxCode","0"
	newElem.SetAttribute "TaxMode","0"
	newElem.SetAttribute "TaxFormula","0"
	newElem.SetAttribute "TaxValue","0"
	newElem.SetAttribute "TaxAmount",Cdbl(dRoundOff)
	newElem.SetAttribute "AccHead",sRoundoffHead
	newElem.Text = "ROUND OFF"
	oNodTaxRoot.Appendchild newElem

	sVouStatus = "010104" 'Voucher created and Accounted
	'con.BeginTrans

	sQuery="Select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close

	sQuery="Select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCrTransNo=objRs(0)
	objRs.Close

	sQuery = " insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"
	sQuery = sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sSalInvNo&"',convert(datetime,'"&sSalInvDt&"',103),"&dTotal
	sQuery = sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)

	sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
	sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus,AccountedOn) values"
	sQuery=sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sSalInvNo&"','"&sSalInvNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"',getdate())"


	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)

	For each EntryNode in oNodDeatils.childNodes
	'	Response.end
		sEntryno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		sItemDesc=Replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
		dQty=EntryNode.Attributes.Item(3).nodeValue
		sUOM=EntryNode.Attributes.Item(4).nodeValue
		dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
		dRate=EntryNode.Attributes.Item(6).nodeValue
		dDisPer=EntryNode.Attributes.Item(8).nodeValue
		dDisAmount=EntryNode.Attributes.Item(9).nodeValue
		dEntRndVal = EntryNode.Attributes.Item(10).nodeValue
		dNoofPack = EntryNode.Attributes.Item(11).nodeValue
		dPackTy = EntryNode.Attributes.Item(12).nodeValue
		dRatePer = EntryNode.Attributes.Item(13).nodeValue
		
		IF Cstr(dEntRndVal) = "" Then
			dEntRndVal = 0
		End IF
		
		IF Cstr(dNoofPack) = "" Then
			dNoofPack = 0
		End IF
		
		sEntryType="C"
		sNarration=""
		
		IF CStr(sAccUnit) = "S" Then
				sAccUnit = sOrgId
		End IF

		IF CStr(sAccUnit) = "" Then
			sAccUnit = sOrgId
		End IF

		For each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					sAccCode=HeaderNode.Attributes.Item(0).nodeValue
					sAccType=HeaderNode.Attributes.Item(4).nodeValue
			end if 'End of Check for Account head Node
		next 'End of Entry Node Loop

		sQuery = "insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,RoundOffvalue,NoofPack,PackCode,RatePer ) values ("
		sQuery = sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&", "&dEntRndVal&","&dNoofPack&","&dPackTy&","&dRatePer&")"
		
		Response.Write sQuery& "<BR><BR>"
		con.execute(sQuery)
		
		if sItemFlag = false then
		    IF CStr(sRoundoffHead) = "0" Then
		   	        dInvWtRnd = EntryNode.getAttribute("Amount")
			        dInvWtRnd = CDbl(dInvWtRnd) + CDbl(dRoundOff)
			        EntryNode.setAttribute "Amount",dInvWtRnd
			        sAmount=EntryNode.Attributes.Item(2).nodeValue
			        sItemFlag = true
	        End IF
	    end if 'if sItemFlag = false then
		
		
		sQuery = "insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
		sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,RoundOffvalue,NoofPack,PackCode ) values ("
		sQuery = sQuery& iTransNo&",'"&sAccUnit&"'"
		sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","&dEntRndVal&","&dNoofPack&","&dPackTy&")"
				
		Response.Write sQuery& "<BR><BR>"
		con.execute(sQuery)

	next'End of Voucher Node Loop


	iSno=1
	For each EntryNode in oNodTaxRoot.childNodes
		iCatCode=EntryNode.Attributes.Item(0).nodeValue
		iTaxCode=EntryNode.Attributes.Item(1).nodeValue
		sTaxMode=EntryNode.Attributes.Item(2).nodeValue
		dTaxPer=EntryNode.Attributes.Item(4).nodeValue
		dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
		sAccCode=EntryNode.Attributes.Item(6).nodeValue
		
		if dTaxAmount >=0 then
			sEntryType = "C"
		else
			sEntryType = "D"
			dTaxAmount=dTaxAmount*-1
		End if
		
		if sTaxMode="P" then
			sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					 ""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
		else
			sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					 ""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
		end if
		Response.Write sQuery& "<BR><BR>"
		con.execute(sQuery)
		
		IF CStr(sAccCode) <> "0" Then
		
			if sTaxMode="P" then
				sQuery = "INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
			else

				sQuery = "INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
			end if
			
			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
		
		End IF
		iSno=cint(iSno)+1
		
		
	Next

	Response.Clear
	dAdjAmtTotal=0
	if  IsObject(oNodAdvRoot) then
		for each EntryNode in oNodAdvRoot.childNodes
			'iAdvTranNo=EntryNode.Attributes.Item(0).nodeValue
			'iCrAdvno=EntryNode.Attributes.Item(0).nodeValue
			dAdvAmount=EntryNode.Attributes.Item(5).nodeValue
			iCrAdvno = EntryNode.Attributes.Item(7).nodeValue
			sAdjType = EntryNode.Attributes.Item(8).nodeValue
			IF CStr(sAdjType) = "B" Then
				dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

				sQuery = "Select AdvanceNumber From ACC_T_ADVANCEPAYMENTS Where CreatedAdvanceNo = "&iCrAdvno&" "
				'Response.Write sQuery &"<br><br>"
				
				objRs.Open sQuery,Con
				IF Not objRs.EOF Then
					'iCrAdvno = objRs(0)
					iAdvNo = objRs(0)
				Else
					iAdvNo = 0
				End IF
				objRs.Close
				
				'sExp = "//AdvanceDetails[@TransNo="&iAdvTranNo&" and @AdvNo]"
				'Set CheckNode = oNodAdvRoot.selectNodes(sExp)

				'IF CheckNode.length <> 0 Then
				'	iAdvNo = EntryNode.Attributes.Item(5).nodeValue
				'Else
				'	iAdvNo = 0
				'End IF
				
				Response.Write iAdvNo &"<br>"

				'IF CStr(iAdvNo) = "0" Then
				'	sQuery = "Select AdvanceNumber From Acc_T_AdvancePayments Where TransactionNumber = "&iAdvTranNo&" "
				'	
				'	objRs.Open sQuery,con
				'	IF Not objRs.EOF Then
				'		iAdvNo = objRs(0)
				'	End IF
				'	objRs.Close
				'End IF

				sQuery = "update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
						 " where CreatedAdvanceNo = "&iCrAdvno
						 
				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
				
				sQuery = "update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
						 " where AdvanceNumber = "&iAdvNo
				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
				
				sAdvNarr = "Sales Invoice No: "&sSalInvNo&" DT: " & sVoucDate
				sQuery = "INSERT INTO ACC_T_CreatedAdvanceAdj (CreatedAdvanceNo, CreatedTransNo,  "&_
						 "AdjustedOn, AmountAdjusted, Narration)  "&_
						 "VALUES ("&iCrAdvno&", "&iCrTransNo&", Convert(Datetime,"&sVoucDate&",103), "&dAdvAmount&", '"&sAdvNarr&"') "
						 
				'Response.Write sQuery &"<br>"
				con.execute(sQuery)
			Elseif CStr(sAdjType) = "SR" and CDbl(dAdvAmount) <> 0 Then
				dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
				sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&iAdvTranNo&" "
				objRs.Open sQuery,Con
				IF Not objRs.EOF Then
					iCrPayNo = objRs(0)
				End IF
				objRs.Close
				sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
						 "PaidOn,AmountPaid) Values ("&iCrPayNo&","&iCrTransNo&","&_
						 "getdate(),"&dAdvAmount&")"
				con.execute(sQuery)
				
				sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
					     "PaidOn,AmountPaid) Values ("&iAccPayNo&","&iTransNo&","&_
						 "getdate(),"&dAdvAmount&")"
				con.execute(sQuery)
				
			Elseif CStr(sAdjType) = "SC" and CDbl(dAdvAmount) <> 0 Then
				dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
				sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&iCrAdvno&" "
				'Response.Write sQuery &"<br><br>"
				objRs.Open sQuery,Con
				IF Not objRs.EOF Then
					iCrPayNo = objRs(0)
				Else
					iCrPayNo = 0
				End IF
				objRs.Close
				
				IF CStr(iCrPayable) <> "0" Then
					sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
							 "PaidOn,AmountPaid) Values ("&iCrPayNo&","&iCrTransNo&","&_
							 "getdate(),"&dAdvAmount&")"
					con.execute(sQuery)
					
					sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
							 "PaidOn,AmountPaid) Values ("&iAccPayNo&","&iTransNo&","&_
							 "getdate(),"&dAdvAmount&")"
					con.execute(sQuery)
				
				End IF
			Elseif CStr(sAdjType) = "OT" and CDbl(dAdvAmount) <> 0 Then
				dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
				sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&iCrAdvno&" "
				objRs.Open sQuery,Con
				IF Not objRs.EOF Then
					iCrPayNo = objRs(0)
				End IF
				objRs.Close
				sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
						 "PaidOn,AmountPaid) Values ("&iCrPayNo&","&iCrTransNo&","&_
						 "getdate(),"&dAdvAmount&")"
				con.execute(sQuery)
				
				sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
						 "PaidOn,AmountPaid) Values ("&iAccPayNo&","&iTransNo&","&_
						 "getdate(),"&dAdvAmount&")"
				con.execute(sQuery)
			End IF
		Next
	end if

	sPayNarration="SALE INV NO:"&sSalInvNo&" Dt:"&sSalInvDt
	sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_CreatedReceivables"
	objRs.open sQuery,con
		iCrPayable=objRs(0)
	objRs.Close

	sQuery="INSERT INTO Acc_T_CreatedReceivables(ReceivableNumber, CreatedTransNo, OUDefinitionID,"&_
			"VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber,"&_
			" PartyInvoiceDate, AmountReceivable, AmountReceived,Narration)values("&iCrPayable&","&iCrTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sSalInvNo&"',"&_
			"convert(datetime,'"&sSalInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"
			
	'Response.Write sQuery &"<br><br>"
	con.execute(sQuery)

	sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_Receivables"
	objRs.open sQuery,con
		iPayableNo=objRs(0)
	objRs.Close

	sQuery="INSERT INTO Acc_T_Receivables(ReceivableNumber,TransactionNumber, OUDefinitionID, "&_
			"VoucherDate, PartyType, PartySubType, PartyCode,PartyInvoiceNumber, "&_
			"CreatedReceivable,PartyInvoiceDate, AmountReceivable, AmountReceived,Narration) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sSalInvNo&"',"&_
			""&iCrPayable&",convert(datetime,'"&sSalInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"

	'Response.Write sQuery& "<BR>"
	con.execute(sQuery)

	if CDbl(dAdjAmtTotal) > 0 then
		sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
				 "ReceivedOn,AmountReceived) Values ("&iCrPayable&","&iCrTransNo&","&_
				 "convert(datetime,'"&sSalInvDt&"',103),"&dAdjAmtTotal&")"

		'Response.write squery
		con.execute(sQuery)
		
		sQuery = "insert into Acc_T_RcvblAdjustmentDetails (ReceivableNumber,RecdByTransactionNo,"&_
				 "ReceivedOn,AmountReceived) Values ("&iPayableNo&","&iTransNo&","&_
				 "convert(datetime,'"&sSalInvDt&"',103),"&dAdjAmtTotal&")"
		'Response.Write sQuery & "<br><br>"
		con.execute(sQuery)
	end if

	'Con.RollbackTrans
	'Response.End
	'IF  IsObject(oNodAgent) then
	'	For each EntryNode in oNodAgent.childNodes
	'		iAgentCode=EntryNode.Attributes.Item(0).nodeValue
	'		sCommType=EntryNode.Attributes.Item(2).nodeValue
	'		dCommPer=EntryNode.Attributes.Item(3).nodeValue
	'		dCommAmount=EntryNode.Attributes.Item(4).nodeValue
	'		sAgType = EntryNode.Attributes.Item(5).nodeValue
	'		sAgParType = EntryNode.Attributes.Item(6).nodeValue
	'			
	'		sQuery="INSERT INTO Sal_T_AdditionalAgents(AgentCode, CommissionType,CurrencyCode, "&_
	'				"AccTransactionNo, AgentCommission,CommissionToPay,AgentType,AgentSubType) "&_
	'				"values ("&iAgentCode&",'"&sCommType&"',1,"&_
	'				""&iTransNo&","&dCommAmount&",'1','"&sAgType&"',"&sAgParType&")"
	'
	'	'Response.Write sQuery& "<BR>"
	'	con.execute(sQuery)
	'	Next
	'End IF

	sNewNarr = Trim(sNewNarr)
	sNewNarr = Mid(sNewNarr,2)
	sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
			 "TransactionNumber = "&iTransNo&" "
	con.execute(sQuery)
		
	sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
			 "TransactionNumber = "&iTransNo&" "
	con.execute(sQuery)
	
	'******************************** Cost and ANal Procedure Call ************************
	sExp = "//Entry"
	Set HeaderNode = oNodRoot.selectNodes(sExp)
	IF HeaderNode.length <> 0 Then
		For iCtr = 0 To HeaderNode.length - 1
			sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value
			sExp="//Entry[@No='"&sEntryno&"']/AccHead"
			set tempNode=oNodRoot.selectNodes(sExp)
			sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value
			set nodANL=oDOM.createElement("Root")
			nodANL.setAttribute "TransNo",iTransNo
			nodANL.setAttribute "EntryNo",sEntryno
			nodANL.setAttribute "UnitCode", sorgid
			nodANL.setAttribute "GlHead",sAccCode
			nodANL.setAttribute "ACTFlag","C"

			sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
			set tempNode=oNodRoot.selectNodes(sExp)
			if tempNode.length >0 then
				set EntryNode=tempNode.item(0).cloneNode(true)
				nodANL.appendChild(EntryNode)
				bAddFlag=true
			end if

			sExp="//Entry[@No='"&sEntryno&"']/Analytical"
			set tempNode=oNodRoot.selectNodes(sExp)
			if tempNode.length >0 then
				set EntryNode=tempNode.item(0).cloneNode(true)
				nodANL.appendChild(EntryNode)
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
		next 'End of Entry Node Loop
	End IF
	
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

	'******************************** Coast and ANal Procedure Call ************************
End Function

%>

