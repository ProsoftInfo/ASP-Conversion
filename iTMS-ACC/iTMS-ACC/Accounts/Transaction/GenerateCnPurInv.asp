<%
Function CreateCNPurInv(iCrVouNo,sVouNo)
	Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot
	dim sParCode,sParSubType,sParType,oNodTemp,sQuery,sDocType,sVouStatus
	dim EntryNode,HeaderNode,nodANL,newElem,sAddtional,iSno,sAmount,sTemp
	dim sNarration,sAccount,sOrgId,sBookNo,sVouType,sVouName
	dim sVouCode,sApprove,sVoucDate,sAccUnit,dTotal,sTransType,dCRAmt,dDRAmt
	dim sAccType,sAccCode,sEntryType,sEntryno,iCrTransNo,sAppBy,sAppTy,sExp,NarrNode
	dim sPurchaseType,ActualTransNo,sDispType,TaxCalNode,dTaxValNoAcc,iCtr
	Dim dEachItmVal,dEachItm,dNoAccTotVal,dAdjAmtTotal,dDiffVal,sAccHeadCode,iTemp
	dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,dTBDiffAmt
	Dim dHdVal,dDetVal,dTaxDebVal,dTaxCreVal,dAccVal,iTransNo,iCrPayNo

	sVouStatus="010104" 'Crearted For Accounting to be Approved
	sVouCode="07"
	sVouType = "C"

	Set objRs  = server.CreateObject("adodb.recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.load server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml")
	Set oNodRoot=oDOM.documentElement

	For each oNodTemp in oNodRoot.childNodes
		IF oNodTemp.nodeName="Header" then
			for Each oNodEntry in  oNodTemp.childNodes
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
				if oNodEntry.nodeName="PurInvoice" then
					sPurInvNo=oNodEntry.Attributes.Item(0).nodeValue
					sPurInvDt=oNodEntry.Attributes.Item(1).nodeValue
					sAppTy = oNodEntry.Attributes.getNamedItem("Approval").value
					sAppBy = oNodEntry.Attributes.getNamedItem("Approver").value
					ActualTransNo = oNodEntry.Attributes.getNamedItem("PurTransNo").value
					
				end if
				if oNodEntry.nodeName="PurchaseType" then
					sPurchaseType=oNodEntry.Attributes.Item(0).nodeValue
				end if
			next
		end if
		if oNodTemp.nodeName="Details" then
			set oNodDeatils=oNodTemp
		end if
		if oNodTemp.nodeName="TaxDetails" then
			set oNodTaxRoot=oNodTemp
			dTotal=oNodTemp.Attributes.Item(0).nodeValue
		end if
	next

	sExp = "//Narration"
	Set NarrNode = oNodRoot.selectNodes(sExp)
	IF NarrNode.length <> 0 Then
		sNarration = NarrNode.Item(0).Text
	Else
		sNarration = ""
	End IF

	sExp = "//TaxDetails/Tax[@AccHead=0]"
	Set TaxCalNode = oNodRoot.selectNodes(sExp)
	IF TaxCalNode.length <> 0 Then
		For iCtr = 0 To TaxCalNode.length - 1
			dTaxValNoAcc = CDbl(dTaxValNoAcc) + CDbl(TaxCalNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value)
		Next
	End IF
	sExp = "//Entry"
	Set TaxCalNode = oNodRoot.selectNodes(sExp)
	dEachItm = TaxCalNode.length

	IF CDbl(dEachItm) = 1 Then
		dEachItmVal = dTaxValNoAcc
	Else
		dEachItmVal =  Round(CDbl(dTaxValNoAcc) / CDbl(dEachItm),0)
		dNoAccTotVal = CDbl(dEachItmVal) * CDbl(dEachItm)
		dDiffVal = CDbl(dTaxValNoAcc) - CDbl(dNoAccTotVal)
	End IF

	sExp = "//Voucher"
	Set NarrNode = oNodRoot.selectNodes(sExp)
	IF NarrNode.length <> 0 Then
		NarrNode.Item(0).Attributes.getNamedItem("CreatedVouNo").Value = iCrVouNo
		NarrNode.Item(0).Attributes.getNamedItem("VouNo").Value = sVouNo
	End IF

	sTransType="CNR"
	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

	'con.BeginTrans

	sQuery = "Select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCrTransNo=objRs(0)
	objRs.Close

	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close


	sQuery = "Insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus, BankInstrumentNo) values"
	sQuery = sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery = sQuery&",'Purchase ADD ','PA','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"', "&ActualTransNo&")"

	Response.Write sQuery &"<br><br>"
	con.execute(sQuery)

	sQuery = "Insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus, BankInstrumentNo) values"
	sQuery = sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sVouNo&"','"&iCrVouNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery = sQuery&",'Purchase Add ','"&sVouType&"','PA',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"', "&ActualTransNo&")"

	Response.Write sQuery& "<BR>"
	con.execute(sQuery)

	iCtr = 1
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

		sEntryType = "D"
		IF	CDbl(sAmount) > 0 then
			For each HeaderNode in EntryNode.childNodes
				IF HeaderNode.nodeName="AccHead" then
						sAccCode=HeaderNode.Attributes.Item(0).nodeValue
						sAccType=HeaderNode.Attributes.Item(4).nodeValue
				end if 'End of Check for Account head Node
			next 'End of Entry Node Loop
			
			sAmount = CDbl(sAmount) + CDbl(dEachItmVal)
			IF CDbl(iCtr) = 1 Then
				sAmount = CDbl(sAmount) + CDbl(dDiffVal)
				iCtr = 2
			End IF
			
			'EntryNode.Attributes.Item(2).nodeValue = sAmount

			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
			sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"

			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
			
			sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
			sQuery=sQuery& iTransNo&",'"&sOrgId&"'"
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"
			
			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
		
		end if
	next'End of Voucher Node Loop

	sQuery = "Select Max(isNull(PayablesNumber,0))+ 1 From Acc_T_CreatedPayables "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		iCrPayNo = objRs(0)
	Else
		iCrPayNo =  "0"
	End IF
	objRs.Close
		
	sQuery = "INSERT INTO Acc_T_CreatedPayables (PayablesNumber, CreatedTransNo, OUDefinitionID, PartyType, "&_
			 "PartySubType, PartyCode, VoucherDate, PartyBillNumber, PartyBillDate,AmountPayable, Narration) "&_
			 "VALUES ("&iCrPayNo&", "&iCrTransNo&", '"&sOrgId&"', "&_
			 "'"&sParType&"', "&sParSubType&", "&sParCode&", Convert(datetime,'"&sVoucDate&"',103), "&_
			 "'"&sPurInvNo&"', Convert(datetime,'"&sPurInvDt&"',103), "&dTotal&", 'CR Note For Purchase Invoice') "
				 
	Response.Write sQuery &"<br><br>"
	con.execute(sQuery)

	sQuery = "Select Max(isNull(PayablesNumber,0))+ 1 From Acc_T_Payables "
	objRs.open sQuery,con
	If Not objRs.EOF Then
		iPayableNo = objRs(0)
	Else
		iPayableNo = 0
	End If
	objRs.Close
			
	sNarration = "Cr Note For Purchase Inv "&sPurInvNo&" "&sPurInvDt
			
			
	sQuery = "INSERT INTO Acc_T_Payables (PayablesNumber, CreatedPayablesNumber, TransactionNumber, "&_
			 "OUDefinitionID, PartyType, PartySubType, PartyCode, VoucherDate, DocumentType, "&_
			 "PartyBillNumber, PartyBillDate, AmountPayable, Narration) "&_
			 "VALUES("&iPayableNo&", "&iCrPayNo&", "&iTransNo&", '"&sOrgId&"', "&_
			 " '"&sParType&"', "&sParSubType&", "&sParCode&", Convert(datetime,'"&sVoucDate&"',103), "&_
			 " 'B', '"&sPurInvNo&"', Convert(datetime,'"&sPurInvDt&"',103), "&dTotal&", '"&sNarration&"') "
					 

	Con.Execute sQuery		
	'------------------------------------------------------------------------------------------------------------------------

	'---------------------------------- FOR TAX NODE DETAILS ----------------
	iSno=1
	dim sTaxMode
	for each EntryNode in oNodTaxRoot.childNodes
		iCatCode=EntryNode.Attributes.Item(0).nodeValue
		iTaxCode=EntryNode.Attributes.Item(1).nodeValue
		sTaxMode=EntryNode.Attributes.Item(2).nodeValue
		dTaxPer=EntryNode.Attributes.Item(4).nodeValue
		dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
		sAccCode=EntryNode.Attributes.Item(6).nodeValue

		If dTaxAmount >=0 then
			sEntryType = "C"
		else
			sEntryType = "D"
			dTaxAmount=dTaxAmount*-1
		End if

		if	CDbl(dTaxAmount) > 0 and cint(sAccCode)>0  then
			if sTaxMode="P" then
				sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sPurchaseType&",NULL,"&dTaxPer&","&dTaxAmount&")"
				con.execute(sQuery)
				
				sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iTransNo&","&sAccCode&",'"&sTaxTrType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sPurchaseType&",NULL,"&dTaxPer&","&Abs(dTaxAmount)&")"
				con.execute(sQuery)	
				
			else
				sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sPurchaseType&",NULL,NULL,"&dTaxAmount&")"
				con.execute(sQuery)
				
				sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iTransNo&","&sAccCode&",'"&sTaxTrType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sPurchaseType&",NULL,NULL,"&Abs(dTaxAmount)&")"
				con.execute(sQuery)	
			end if
		end if
		iSno=cint(iSno)+1
	Next

	'============== Checking For Voucher Entry ===========================
		
	sQuery = "Select VoucherAmount,CRDRIndication From Acc_T_VoucherHeader Where TransactionNumber = "&iTransNo&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		dHdVal = objRs(0) 'Credit Value
	End IF
	objRs.Close
		
	sQuery = "Select SUM(Amount) From Acc_T_VoucherDetails Where TransactionNumber = "&iTransNo&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		dDetVal = objRs(0) ' Debit Value
	End IF
	objRs.Close
		
	sQuery = "Select SUm(TaxAMount),TransCrDrIndication From Acc_T_VoucherTaxDetails Where TransactionNumber = "&iTransNo&" "&_
			 "Group By TransCrDrIndication "
				 
	objRs.Open sQuery,Con
	Do While Not objRs.EOF
		IF CStr(objRs(1)) = "D" Then
			dTaxDebVal = objRs(0)
		Else
			dTaxCreVal = objRs(0)
		End IF
		objRs.MoveNext
	Loop
	objRs.Close
		
	dDetVal = FormatNumber(dDetVal,2,,,0)
	dTaxDebVal = FormatNumber(dTaxDebVal,2,,,0)
	dTaxCreVal = FormatNumber(dTaxCreVal,2,,,0)
	dHdVal = FormatNumber(dHdVal,2,,,0)
		
	dAccVal = CDbl(dDetVal) + CDbl(dTaxDebVal) - CDbl(dTaxCreVal)
	dAccVal = FormatNumber(dAccVal,2,,,0)
		
	IF CDbl(dAccVal) <> CDbl(dHdVal) Then
		Response.Write "Debit and Credit Does'nt Match Inserion Rolled Back "
		Con.RollbackTrans
		Response.End
	End IF
		
	sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNarration&"' Where  "&_
			 "TransactionNumber = "&iTransNo
	Con.Execute sQuery
		
	sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNarration&"' Where  "&_
			 "TransactionNumber = "&iTransNo
	Con.Execute sQuery

	oNodRoot.Attributes.Item(0).nodeValue= iCrTransNo
	oNodRoot.Attributes.Item(1).nodeValue= sVouNo

	'dTBDiffAmt = GetTBDiff(sOrgId)
	'IF CStr(dTBDiffAmt) <> "0" Then	
	'	Response.Write "<br><br><b>Differences : "& dTBDiffAmt
	'	Response.Write "<br></br></b>"
	'	Con.RollbackTrans
	'	Response.End
	'Else
	'	Con.CommitTrans
	'	'Con.RollbackTrans
	'	'Response.End
	'	
	'
	'	oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	'	
	'End IF

End Function
%>
