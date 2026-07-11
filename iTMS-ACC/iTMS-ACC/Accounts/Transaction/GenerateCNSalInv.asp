<%
Function CreateCNSalInv(sCrVouNo,iVouNo,sDispTy,sOldBkNo,sOldVouDate,sChkGj)
	Response.Clear
	Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot
	dim sParCode,sParSubType,sParType,oNodTemp,sQuery,sNarration,sAccount
	dim EntryNode,HeaderNode,nodANL,newElem,sAddtional,iSno,sAmount,sTemp
	dim sOrgId,sBookNo,sVouType,sVouName,sVouCode,sApprove,sVoucDate,sAccUnit
	dim dTotal,sTransType,sAccType,sAccCode,sEntryType,sEntryno,iTransNo
	dim sDocType,sVouStatus,iCounter,sSalType, ActualTransNo,sExp,NarrNode,sVouNarr
	dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt,iCrRecNo
	dim dAdjAmtTotal,sAppTy,sAppBy,iCrTransNo,objfs,sCallTy,sFromSal,sAccHeadCode,iTemp
	Dim TaxCalNode,dTaxValNoAcc,iCtr,dEachItmVal,dEachItm,dNoAccTotVal,dDiffVal
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,sTaxMode,dTBDiffAmt
	Dim dHdVal,dDetVal,dTaxDebVal,dTaxCreVal,dAccVal,TempNode,iSalTrNo,sTaxTrans
	dim sParName,sPayAt,iVouAmt,sCrDr,iEntryNo,sTrCrDr,objRs1,iAccHead,sNarr,sEntAmt
	Dim iSeriesNo,iSeriesCode,adoCmd
	
	set objRs1 = Server.CreateObject("ADODB.Recordset")

	sAppTy = Request.Form("optApprove")
	sFromSal = Request.Form("hFromSal")
	
	
	sFromSal = "N"
	sVouStatus="010104" 'Waiting For Accounting 
	sVouCode="07"
	sVouType="C"
	sChkGj = "Y"

	set objRs  = server.CreateObject("adodb.recordset")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.load server.MapPath("../temp/transaction/Voucher Amd_DN_"&Session.SessionID&".xml")
	set oNodRoot=oDOM.documentElement

	For each oNodTemp in oNodRoot.childNodes
		if oNodTemp.nodeName="Header" then
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
				if oNodEntry.nodeName="SaleInvoice" then
					sPurInvNo=oNodEntry.Attributes.Item(0).nodeValue
					sPurInvDt=oNodEntry.Attributes.Item(1).nodeValue
					sAppTy = oNodEntry.Attributes.Item(3).nodeValue
					sAppBy = oNodEntry.Attributes.Item(4).nodeValue
					ActualTransNo = oNodEntry.Attributes.Item(5).nodeValue
					iSalTrNo = oNodEntry.Attributes.Item(5).nodeValue
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
			dTotal=oNodTemp.Attributes.Item(0).nodeValue
		end if
	next

	sExp = "//Narration"
	Set NarrNode = oNodRoot.selectNodes(sExp)
	IF NarrNode.length <> 0 Then
		sVouNarr = Trim(NarrNode.item(0).text)
	Else
		sVouNarr = ""
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
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.getNamedItem("CreatedVouNo").Value = sCrVouNo
		TempNode.Item(0).Attributes.getNamedItem("VouNo").Value = iVouNo
	End IF

	sTransType="CNR"
	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

	

	sQuery = "Select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCrTransNo=objRs(0)
	objRs.Close

	
	'Response.Clear
	
	sQuery = "Insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
	sQuery = sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,BankInstrumentNo,PayableAt,PurchaseBillType) values"
	sQuery = sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery = sQuery&",'Sales Return ','OS','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"',"&iSalTrNo&",'"&sDispTy&"','"&sDispTy&"')"

	con.execute(sQuery)
	'Response.Write sChkGj
	'=================================================================================================
	
	For each EntryNode in oNodDeatils.childNodes
		sEntryno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
		dQty=EntryNode.Attributes.Item(3).nodeValue
		sUOM=EntryNode.Attributes.Item(4).nodeValue
		dBasicAmount="0"
		dRate=EntryNode.Attributes.Item(6).nodeValue
		dDisPer=EntryNode.Attributes.Item(8).nodeValue
		dDisAmount=EntryNode.Attributes.Item(9).nodeValue


		sEntryType="D"
		sNarration=""
		
		if	CDbl(sAmount) > 0 then
			for each HeaderNode in EntryNode.childNodes
				if HeaderNode.nodeName="AccHead" then
						sAccCode=HeaderNode.Attributes.Item(0).nodeValue
						sAccType=HeaderNode.Attributes.Item(4).nodeValue
				end if 'End of Check for Account head Node
			next 'End of Entry Node Loop
		
			sAmount = CDbl(sAmount) + CDbl(dEachItmVal)
			
			IF CDbl(iCtr) = 1 Then
				sAmount = CDbl(sAmount) + CDbl(dDiffVal)
				iCtr = 2
			End IF 
			
			sQuery = "Insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
			sQuery = sQuery& iCrTransNo&",'"&sOrgId&"'"
			sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery = sQuery&" '"&sVouNarr&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"

			con.execute(sQuery)
		End IF
	next'End of Voucher Node Loop
	'Response.Write "Test Sal Inv....."
	'Response.End 
	'=============================================================================================
	iSno=1
	For each EntryNode in oNodTaxRoot.childNodes
		IF EntryNode.nodeName = "Tax" Then
			iCatCode=EntryNode.Attributes.Item(0).nodeValue
			iTaxCode=EntryNode.Attributes.Item(1).nodeValue
			sTaxMode=EntryNode.Attributes.Item(2).nodeValue
			dTaxPer=EntryNode.Attributes.Item(4).nodeValue
			dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
			sAccCode=EntryNode.Attributes.Item(6).nodeValue
		
			if	CDbl(dTaxAmount) > 0 then
				sEntryType = "D"
			else
				sEntryType = "C"
				dTaxAmount=dTaxAmount*-1
			End if
			
			IF CStr(sSalType) = "" Then
				sSalType = 0
			End IF
			
			if sTaxMode="P" then
				sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
				con.execute(sQuery)
			Else
				sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
				con.execute(sQuery)
			End IF
			iSno=cint(iSno)+1
		End IF
	Next

	
	sQuery = "select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
	objRs.open sQuery,con
		iPayableNo=objRs(0)
		iCrRecNo=objRs(0)
	objRs.Close

	sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber,CreatedTransNo, OUDefinitionID, "&_
			"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
			"PartyBillDate, AmountPayable, AmountPaid) values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
			"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
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
	
	
	'if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher Amd_DN_"&Session.SessionID&".xml")) then
	'	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher Amd_DN_"&Session.SessionID&".xml"))
	'End IF	
	
End Function
%>
