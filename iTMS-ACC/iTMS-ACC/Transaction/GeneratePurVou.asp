<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

Function CreatePurInv(sCrVouNo,iVouNo,iOldBookNo,sOldVouDate)
	'XML DOM Variables
	Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot
	dim sParCode,sParSubType,sParType,oNodTemp,sQuery,EntryNode,HeaderNode,nodANL,newElem
	dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp,sOrgId,sBookNo,sVouType
	dim sVouCode,sApprove,sVoucDate,sAccUnit,sVouName,sDocType,sVouStatus
	dim dTotal,sTransType,sAccType,sAccCode,sEntryType,sEntryno,iTransNo
	dim sSalType,sRefernceNo,iCounter,iCrTransNo
	dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
	Dim sAppTy,sAppBy,sExp,dRoundOff,sRoundoffHead,sGroupCode,bAddFlag,Objfs
	Dim CheckNode,CheckVal,dInvWtRnd,dAdjAmtTotal,sAccHeadCode,iTemp,iAdvNo
	Dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,sTaxMode
	Dim iAdvTranNo,dAdvAmount,iCrAdvNo,sAdjType,iCrRecNo,sAdjNarr,sNewNarr,dTbDiff
	Dim sPayNarration,iCrPayableNo,sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode
	Dim iSeriesNo,iSeriesCode,iCrSeriesNo,iCrSeriesCode,sNoSerTy,sPurCategory,sItemFlag
	sItemFlag = false
	
	sVouCode="04"
	sVouType="C"

	Set objRs  = server.CreateObject("adodb.recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")

	oDOM.Load server.MapPath("../../Purchase/xmldata/General.xml")
	Set oNodRoot = oDOM.documentElement
	sExp = "//ROUNDOFF"
	Set oNodEntry = oNodRoot.Selectnodes(sExp)
	If oNodEntry.Length > 0 then
		sRoundoffHead = oNodEntry.Item(0).Attributes.Item(0).nodevalue
	else
		sRoundoffHead = "0"
	end if

	oDOM.load server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")

	set oNodRoot=oDOM.documentElement
	sExp = "//TaxDetails[@RoundOffValue]"
	Set CheckNode = oNodRoot.selectNodes(sExp)
	IF CheckNode.length <> 0 Then
		CheckVal = "T"
	Else
		CheckVal = "F"
	End IF

	sExp = "//Voucher"
	Set CheckNode = oNodRoot.selectNodes(sExp)
	Response.Write CheckNode.Length
	IF CheckNode.length <> 0 Then
		 CheckNode.Item(0).Attributes.Item(1).nodeValue = sCrVouNo
		 CheckNode.Item(0).Attributes.Item(3).nodeValue = iVouNo
	End IF

	For Each oNodTemp in oNodRoot.childNodes
		if oNodTemp.nodeName="Header" then
			for Each oNodEntry in  oNodTemp.childNodes
				if oNodEntry.nodeName="Organization" then
					sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				end if
				if oNodEntry.nodeName="Book" then
					sBookNo=oNodEntry.Attributes.Item(0).nodeValue
				end if
				if oNodEntry.nodeName="PurchaseType" then
					sSalType=oNodEntry.Attributes.Item(0).nodeValue
				end if
				if oNodEntry.nodeName="Party" then
					sParType=oNodEntry.Attributes.Item(0).nodeValue
					sParSubType=oNodEntry.Attributes.Item(1).nodeValue
					sParCode=oNodEntry.Attributes.Item(3).nodeValue
				end if
				if oNodEntry.nodeName="PurInvoice" then
					sPurInvNo=oNodEntry.Attributes.Item(0).nodeValue
					sPurInvDt=oNodEntry.Attributes.Item(1).nodeValue
					sAppTy = oNodEntry.Attributes.Item(2).nodeValue
					sAppBy = oNodEntry.Attributes.Item(3).nodeValue
					sRefernceNo=sPurInvNo&"-"&sPurInvDt
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
			Set oNodTaxRoot=oNodTemp
			dTotal=oNodTemp.Attributes.Item(2).nodeValue
			IF CStr(CheckVal) = "T" Then
				dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
			Else
				dRoundOff = 0
			End IF
		end if
		if oNodTemp.nodeName="AdvanceDetails" then
			set oNodAdvRoot=oNodTemp
		end if
	next

	sExp = "//PurCategory"
	Set TempNode = oNodRoot.selectnodes(sExp)
	IF TempNode.length <> 0 Then
		sPurCategory = TempNode.Item(0).Attributes.getNamedItem("Code").Value
	End IF
	
	sExp = "//AdvanceDetails"
	Set TempNode = oNodRoot.selectnodes(sExp)
	sTransType="PJR"
	
	

	sVouStatus = "010104" 'Voucher To be Approved"
	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

	'Addition of Roundoff Node
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
	
	sQuery = "Select M.CounterType From Acc_M_BookNumberSeries B,Ms_NumberSeries M "&_
			 "Where M.SeriesNo = B.DrSeriesNo And B.BookCode = '04' And "&_
			 "B.BookNumber = "&sBookNo&" And B.OUDefinitionID = '"& sOrgId &"' " 
	
	objRs.open sQuery,con
		sNoSerTy = Trim(objRs(0))
	objRs.close()		 
'&&&&&&&&&&&&&&&&&&&&&&&&&& No Series Generation Check Starts Here &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
	IF CStr(sBookNo) <> CStr(iOldBookNo) Then
		sQuery = "Select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
				 "OUDefinitionID='"&sOrgId&"' and BookCode='04' and BookNumber= "&sBookNo

		objRs.open sQuery,con
		if not objRs.EOF then
			iSeriesNo=objRs(0)
			iSeriesCode=objRs(1)
			iCrSeriesNo=objRs(2)
			iCrSeriesCode=objRs(3)
		end if
		objRs.close()

		sCrVouNo = GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sVoucDate)
		iVouNo = GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
		
	ElseIF Mid(sVoucDate,4,2) <> Mid(sOldVouDate,4,2) and CStr(sNoSerTy) <> "Y"    Then 'IF The Voucher Date is been changed from month to month
		sQuery = "Select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
				 "OUDefinitionID='"&sOrgId&"' and BookCode='04' and BookNumber= "&sBookNo

		objRs.open sQuery,con
		if not objRs.EOF then
			iSeriesNo=objRs(0)
			iSeriesCode=objRs(1)
			iCrSeriesNo=objRs(2)
			iCrSeriesCode=objRs(3)
		end if
		objRs.close()

		sCrVouNo = GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sVoucDate)
		iVouNo = GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
	End IF
	
	'Response.Write "<br>"
	'Response.Write sNoSerTy &"   " & iVouNo &"<br>"
	'Response.End
	

'&&&&&&&&&&&&&&&&&&&&&&&&&& No Series Generation Check Ends Here &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

	sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCrTransNo=objRs(0)
	objRs.Close

	sQuery = "Select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close


	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,InvCategoryCode) values"
	sQuery = sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery = sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"','"&sPurCategory&"')"

	'Response.Write sQuery& "<BR>"
	con.execute(sQuery)

	sQuery = "insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus,InvCategoryCode) values"
	sQuery = sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"','"&sCrVouNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery = sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"','"&sPurCategory&"')"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)

	For each EntryNode in oNodDeatils.childNodes
		sEntryno=EntryNode.Attributes.Item(0).nodeValue
		dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
		sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
		dQty=EntryNode.Attributes.Item(3).nodeValue
		sUOM=EntryNode.Attributes.Item(4).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		dRate=EntryNode.Attributes.Item(6).nodeValue
		dDisPer=EntryNode.Attributes.Item(8).nodeValue
		dDisAmount=EntryNode.Attributes.Item(9).nodeValue

		sEntryType="D"
		sNarration=""
		
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
		sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
		sQuery = sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"

		'Response.Write sQuery& "<BR>"
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
		sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
		sQuery = sQuery& iTransNo&",'"&sAccUnit&"'"
		sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery = sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"

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
			sEntryType = "D"
		else
			sEntryType = "C"
			dTaxAmount = Abs(dTaxAmount)
			dTaxPer = abs(dTaxPer)
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
		con.execute(sQuery)
		
		IF CInt(sAccCode)>0 then
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
		end if
		iSno=cint(iSno)+1
	Next

	dAdjAmtTotal = 0
	'Response.Clear
	
	IF IsObject(oNodAdvRoot) Then
		for each EntryNode in oNodAdvRoot.childNodes
			
			iAdvTranNo=EntryNode.Attributes.Item(0).nodeValue
			dAdvAmount=EntryNode.Attributes.Item(5).nodeValue
			'iCrAdvNo = EntryNode.Attributes.Item(6).nodeValue
			iCrAdvNo = EntryNode.Attributes.Item(7).nodeValue
			sAdjType = EntryNode.Attributes.Item(8).nodeValue
			
			'Response.Write iAdvTranNo &"<br><br>"
			
			sExp = "//AdvanceDetails[@TransNo="&iAdvTranNo&" and @AdvNo]"
			Set CheckNode = oNodAdvRoot.selectNodes(sExp)
			IF CheckNode.length <> 0 Then
				iAdvNo = EntryNode.Attributes.Item(5).nodeValue
			Else
				iAdvNo = 0
			End IF

			IF CStr(iAdvNo) = "0" and CStr(sAdjType) = "I" Then
				sQuery = "Select AdvanceNumber From Acc_T_AdvancePayments Where TransactionNumber = "&iAdvTranNo&" "
				objRs.Open sQuery,con
				IF Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close
			End IF

			'=====================================================================================
			'Added On	:	29/11/2004
			'Reason		:	To update the adjustment based on the selection Bank or Cash
			'				Debit Notes.
			'=====================================================================================
			
			sAdjNarr = "Adjusted to Purchase Invoice No "&sCrVouNo&" DT: "& sVoucDate &" Amt: "& dTotal
			if CDbl(dAdvAmount)>0 then
				IF CStr(sAdjType) = "I" and CDbl(dAdvAmount) <> 0 Then ' Voucher Like Cask/Bank Payment
					dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
					sQuery = "Select CreatedAdvanceNo From ACC_T_ADVANCEPAYMENTS Where AdvanceNumber = "&iCrAdvno&" "
					objRs.Open sQuery,Con
					IF Not objRs.EOF Then
						iCrAdvno = objRs(0)
					End IF
					objRs.Close

					sQuery = "update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
							 " where CreatedAdvanceNo = "&iCrAdvNo

					Response.Write sQuery &"<br><br>"
					con.execute(sQuery)
					
					sQuery = "INSERT INTO ACC_T_CREATEDADVANCEADJ (CREATEDADVANCENO, CREATEDTRANSNO, ADJUSTEDON,  "&_
							 "AMOUNTADJUSTED, NARRATION) VALUES "&_
							 "("&iCrAdvNo&", "&iCrTransNo&", Convert(Datetime,'"&sVoucDate&"',103), "&dAdvAmount&", '"&sAdjNarr&"') "

					Response.Write sQuery &"<br><br>"
					con.execute(sQuery)
					
					sQuery = "update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
							 " where AdvanceNumber = "&iAdvNo

					Response.Write sQuery &"<br><br>"
					con.execute(sQuery)
					
					Response.Write dAdjAmtTotal &"<br><br>"

				Elseif CStr(sAdjType) = "D" and CDbl(dAdvAmount) <> 0 Then ' If the adjustment is of Debit Notes.
					dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
					sQuery = "Select DRCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&iAdvTranNo&" "

					Response.Write sQuery &"<br><br>"
					objRs.Open sQuery,Con
					IF Not objRs.EOF Then
						iCrRecNo = objRs(0)
					Else
						iCrRecNo = 0
					End IF
					objRs.Close

					sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
							 "ReceivedOn,AmountReceived) Values ("&iCrRecNo&","&iCrTransNo&","&_
							 "getdate(),"&dAdvAmount&")"

					Response.Write sQuery &"<br><br>"
					Con.Execute sQuery
					
					sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
							 "ReceivedOn,AmountReceived) Values ("&iAdjRecNo&","&iTransNo&","&_
							 "getdate(),"&dAdvAmount&")"
					con.execute(sQuery)

				End IF
			end if
		Next
	End IF

	IF CDbl(dTotal)>CDbl(dAdjAmtTotal) then
		sQuery = "Select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
		objRs.open sQuery,con
			iCrPayableNo=objRs(0)
		objRs.Close
		
		sQuery = "Select isnull(max(PayablesNumber),0)+1 from Acc_T_Payables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close
		
		sPayNarration="PUR INV No:"&sPurInvNo &" Dt:"&sPurInvDt

		sQuery = "INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
				 "VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
				 "PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iCrPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
				 "convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				 "convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"
		con.execute(sQuery)
		
		sQuery="INSERT INTO Acc_T_Payables(PayablesNumber,TransactionNumber, OUDefinitionID, "&_
				"CreatedPayablesNumber,VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
				"PartyBillDate, AmountPayable, AmountPaid,Narration) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				""&iCrPayableNo&",convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"
		'Response.Write sQuery& "<BR>"
		con.execute(sQuery)

		sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
				 "PaidOn,AmountPaid) Values ("&iCrPayableNo&","&iCrTransNo&","&_
				 "convert(datetime,'"&sVoucDate&"',103),"&dAdjAmtTotal&")"
					
		Response.Write sQuery
		con.execute(sQuery)
		
		sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
				 "PaidOn,AmountPaid) Values ("&iPayableNo&","&iTransNo&","&_
				 "getdate(),"&dAdjAmtTotal&")"
		con.execute(sQuery)
	end if

	sNewNarr = Trim(sNewNarr)
	sNewNarr = Mid(sNewNarr,2)
	sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
			 "TransactionNumber = "&iTransNo&" "
	con.execute(sQuery)
		
	sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
			 "TransactionNumber = "&iTransNo&" "
	con.execute(sQuery)
	
	sExp = "//Entry"
	Set HeaderNode = oNodRoot.selectNodes(sExp)
	IF HeaderNode.length <> 0 Then
		For iCtr = 0 To HeaderNode.length - 1
			sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value
			sExp="//Entry[@No='"&sEntryno&"']/AccHead"
			Set tempNode=oNodRoot.selectNodes(sExp)
			sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value
			Set nodANL=oDOM.createElement("Root")
			nodANL.setAttribute "TransNo",iCrTransNo
			nodANL.setAttribute "EntryNo",sEntryno
			nodANL.setAttribute "UnitCode", sorgid
			nodANL.setAttribute "GlHead",sAccCode
			nodANL.setAttribute "ACTFlag","C"
			sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
			Set tempNode=oNodRoot.selectNodes(sExp)
			IF tempNode.length >0 then
				set EntryNode=tempNode.item(0).cloneNode(true)
				nodANL.appendChild(EntryNode)
				bAddFlag=true
			End IF

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
	
	'Response.End

'******************************** Coast and ANal Procedure Call ************************
	'Set newElem  = oDOM.createAttribute("CreatedTransNo")
	'newElem.value = iCrTransNo
	'oNodRoot.setAttributeNode(newElem)

	'Set newElem  = oDOM.createAttribute("CreatedVouNo")
	'newElem.value = sCrVouNo
	'oNodRoot.setAttributeNode(newElem)
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	'objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"))
End Function

%>

	
