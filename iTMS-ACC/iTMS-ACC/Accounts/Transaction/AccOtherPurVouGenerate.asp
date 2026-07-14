<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccOtherPurVouGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	March  09, 2006
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

'XML DOM Variables

Dim oDOM,nodHeader,Root,objRs,sQuery,sExp,objFSO
dim sParCode,sParSubType,sParType,sErrChk
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
DIM sOrgId,sBookNo,sVouType,iVouNo,sVouName
DIM sVouCode,sApprove,sVoucDate,sAccUnit
DIM dTotal,sTransType,dCRAmt,dDRAmt
DIM sAccType,sAccCode,sEntryType,sEntryno,iTransNo
DIM sDocType,sVouStatus,sPayTo,iPartyReceiptNo,bPartyReceiptFlag
dim iTemp,oNodRoot,oNodTemp,oNodEntry,sSalType,sPurInvNo
DIM iSeriesNo,iSeriesCode,iCreatedVouNo,iCreatedTransNo,sAccHeadCode
Dim sPurInvDt,sRefernceNo,sCrVouNo,iCrTransNo
Dim oNodAgent,oNodDeatils,oNodTaxRoot,oNodAdvRoot
Dim iCrSeriesNo,iCrSeriesCode,iCatCode,iTaxCode,dTaxPer,dTaxAmount
Dim sRoundoffHead,dRoundOff,sPara,dAccDate
Dim bAddFlag,sVouEntTy,dInvWtRnd,TaxEntryNode,sTaxAccCode

sVouName=Request("hVouName")
sPara = Request("hPara")
sVouCode="04"
sVouType="C"
iCreatedTransNo = Request("hTransNo")
dAccDate = Request("hAccDate")
'Response.Write dAccDate
bPartyReceiptFlag=false
sAccUnit = Request.Form("selUnitId")
IF trim(sPara) = "App" then
	sVouStatus="010103" 'Voucher Approved
ElseIF trim(sPara) = "Acc" then
	sVouStatus="010104" 'Voucher Accounted
End IF
Response.Write "<p>"&sPara
'Response.End

SET objRs  = server.CreateObject("adodb.recordset")
set objFSO = Server.CreateObject("Scripting.FileSystemObject")
sQuery = "Select CreatedTransNo From Acc_T_VoucherHeader Where CreatedTransNo = "&iCreatedTransNo&" "

'Response.write sQuery
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouEntTy = "Y" 'Already Accounted
Else
	sVouEntTy = "N" 'Can Be Accounted
End IF
objRs.Close

IF Cstr(sVouEntTy) = "N" Then

	' Create our DOM Document Objects
	SET oDOM = Server.CreateObject("Microsoft.XMLDOM")
if objFSO.FileExists(server.MapPath("../../Purchase/xmldata/General.xml"))  then
	oDOM.Load server.MapPath("../../Purchase/xmldata/General.xml")
	Set oNodRoot = oDOM.documentElement
	sExp = "//ROUNDOFF"
	Set oNodEntry = oNodRoot.Selectnodes(sExp)
	If oNodEntry.Length > 0 then
		sRoundoffHead = oNodEntry.Item(0).Attributes.Item(0).nodevalue
	else
		sRoundoffHead = "0"
	end if
else
    sRoundoffHead = "0"
end if 

	oDOM.Load server.MapPath("../temp/transaction/Voucher Entry_OthPUR_"&Session.SessionID&".xml")

	set oNodRoot=oDOM.documentElement
	sCrVouNo = oNodRoot.Attributes.Item(1).nodeValue
	iCrTransNo = oNodRoot.Attributes.Item(0).nodeValue

	sExp = "//AccHead"
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		sAccCode = TempNode.Item(0).Attributes.getNamedItem("No").Value
	End IF



	for each oNodTemp in oNodRoot.childNodes
		if oNodTemp.nodeName="Header" then
			for Each oNodEntry in  oNodTemp.childNodes
				if oNodEntry.nodeName="Organization" then
					sOrgId=oNodEntry.Attributes.Item(0).nodeValue
					oNodEntry.setAttribute "AccUnit",sAccUnit
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
			set oNodTaxRoot=oNodTemp
			'dTotal=oNodTemp.Attributes.Item(0).nodeValue
			dTotal=oNodTemp.Attributes.Item(0).nodeValue
			dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
		end if
		if oNodTemp.nodeName="AdvanceDetails" then
			set oNodAdvRoot=oNodTemp
		end if


	next

	IF CStr(sAccUnit) = "" Then
		sAccUnit = sOrgId
	End IF

	sTransType="PJR"

	iCreatedTransNo = iCrTransNo

	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

	sQuery="select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo

	objRs.open sQuery,con
	if not objRs.EOF then
		iSeriesNo=objRs(0)
		iSeriesCode=objRs(1)

	end if
	objRs.close()
	
	
	sExp = "//Tax[@CatCode=0 and @TaxCode=0]"
	Set CheckNode = oNodRoot.selectNodes(sExp)
	IF CheckNode.length <> 0 Then
		sRoundoffHead = CheckNode.Item(0).Attributes.getNamedItem("AccHead").Value
		dRoundOff = CheckNode.Item(0).Attributes.getNamedItem("TaxAmount").Value
	End IF
	
'	IF CStr(sRoundoffHead) = "0" Then
'	    if Trim(dRoundOff)="" or IsNull(dRoundOff) then  dRoundOff = "0"
'		sExp = "//Entry"
'		Set CheckNode = oNodRoot.selectNodes(sExp)
'		IF CheckNode.length <> 0 Then
'			dInvWtRnd = CheckNode.Item(0).Attributes.Item(2).NodeValue 
'			dInvWtRnd = CDbl(dInvWtRnd) + CDbl(dRoundOff)
'			CheckNode.Item(0).Attributes.Item(2).NodeValue  = dInvWtRnd
'		End IF
'	End IF
	
	con.BeginTrans

		'Response.Write sOrgId&" "&iSeriesNo&" "&iSeriesCode&" "&sVoucDate

	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close
	'============================================================================================
	IF trim(sPara) = "Acc" then 	'For Accounting
		'============================================================================================
		iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
		'============================================================================================
		sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
		sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
		sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"
		sQuery=sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
		sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"','"&sCrVouNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
		sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"

		Response.Write sQuery& "<BR><BR>"
		con.execute(sQuery)
		'============================================================================================

		sQuery = "Update Acc_T_CreatedVoucherHeader Set VoucherDate = Convert(Datetime,'"&sVoucDate&"',103) Where CreatedTransNo = "&iCrTransNo
		con.execute(sQuery)

		for each EntryNode in oNodDeatils.childNodes
		dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,dItemCode,dClassCode

			sEntryno=EntryNode.Attributes.Item(0).nodeValue
			dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
			sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
			dQty=EntryNode.Attributes.Item(3).nodeValue
			sUOM=EntryNode.Attributes.Item(4).nodeValue
			sAmount=EntryNode.Attributes.Item(2).nodeValue
			dRate=EntryNode.Attributes.Item(6).nodeValue
			dDisPer=EntryNode.Attributes.Item(8).nodeValue
			dDisAmount=EntryNode.Attributes.Item(9).nodeValue
			dItemCode = EntryNode.Attributes.Item(10).nodeValue
			dClassCode = EntryNode.Attributes.Item(11).nodeValue
			sEntryType = EntryNode.getAttribute("CRDR")


			'sEntryType="D"
			sNarration=""

			for each HeaderNode in EntryNode.childNodes
				if HeaderNode.nodeName="AccHead" then
						sAccCode=HeaderNode.Attributes.Item(0).nodeValue
						sAccType=HeaderNode.Attributes.Item(4).nodeValue
				end if 'End of Check for Account head Node
			next 'End of Entry Node Loop
							
			if CDbl(sEntryno)=1 then
	            if oNodTaxRoot.hasChildNodes() then
                    for each TaxEntryNode in oNodTaxRoot.childNodes

		                iCatCode=TaxEntryNode.Attributes.Item(0).nodeValue
		                iTaxCode=TaxEntryNode.Attributes.Item(1).nodeValue
		                sTaxMode=TaxEntryNode.Attributes.Item(2).nodeValue
		                dTaxPer=TaxEntryNode.Attributes.Item(4).nodeValue
		                dTaxAmount=TaxEntryNode.Attributes.Item(5).nodeValue
		                sTaxAccCode=TaxEntryNode.Attributes.Item(6).nodeValue

		                IF Cstr(dTaxAmount) = "" Then
			                dTaxAmount = 0
		                End IF
		                if CDbl(sTaxAccCode)=0 then
	                        sAmount = cdbl(sAmount) + cdbl(dTaxAmount)
	                    end if
	                Next
			    end if 'if oNodTaxRoot.hasChildNodes() then
			end if 'if CDbl(sEntryno)=1 then
			
		'============================================================================================
				sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
				sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
				sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemCode,ClassificationCode ) values ("
				sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
				sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
				sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","&dItemCode&","&dClassCode&")"

		Response.Write sQuery& "<BR><BR>"
		con.execute(sQuery)

		Next

		iSno=1
		dim sTaxMode,iPurType
		sExp = "//TaxDetails"
		Set oNodTaxRoot = oNodRoot.selectNodes(sExp)

		For iCtr = 0 To oNodTaxRoot.length - 1
			iPurType = oNodTaxRoot.Item(iCtr).Attributes.getNamedItem("PurchaseType").Value

			for each EntryNode in oNodTaxRoot.Item(iCtr).childNodes
				iCatCode=EntryNode.Attributes.Item(0).nodeValue
				iTaxCode=EntryNode.Attributes.Item(1).nodeValue
				sTaxMode=EntryNode.Attributes.Item(2).nodeValue
				dTaxPer=EntryNode.Attributes.Item(4).nodeValue
				dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
				sAccCode=EntryNode.Attributes.Item(6).nodeValue
				IF CStr(iTaxCode) <> "0" Then
					iPurType = EntryNode.Attributes.getNamedItem("PurchaseType").Value
				End IF

				IF CStr(dTaxAmount) = "" Then
					dTaxAmount = 0
				End IF

				if dTaxAmount >=0 then
					'sEntryType = "C"
					sEntryType = "D"
				else
					'sEntryType = "D"
					sEntryType = "C"
					dTaxAmount=dTaxAmount*-1
				End if
				IF CStr(iTaxCode) <> "0" and CStr(iCatCode) <> "0" Then
					if CInt(sAccCode) > 0 then
						if sTaxMode="P" then
							sQuery = "INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
									 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
									 ""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&iPurType&",NULL,"&dTaxPer&","&dTaxAmount&")"
						else

							sQuery = "INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
									 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
									 ""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&iPurType&",NULL,NULL,"&dTaxAmount&")"
						end if


						Response.Write sQuery &"<br><br>"
						con.execute(sQuery)
						iSno=cint(iSno)+1
					end if
				End IF
			Next
		Next

		'======== Inserting the Value For the RoundedOff Value ====================================
		Dim sTaxEntTy
		sTaxEntTy = sEntryType

		sExp = "//TaxDetails/Tax[@CatCode=0 and @TaxCode=0]"
		Set TempNode = oNodRoot.selectNodes(sExp)
		IF TempNode.Length <> 0 Then
			iCatCode=TempNode.Item(0).Attributes.Item(0).nodeValue
			iTaxCode=TempNode.Item(0).Attributes.Item(1).nodeValue
			sTaxMode=TempNode.Item(0).Attributes.Item(2).nodeValue
			dTaxPer=TempNode.Item(0).Attributes.Item(4).nodeValue
			dTaxAmount=TempNode.Item(0).Attributes.Item(4).nodeValue
			sAccCode=TempNode.Item(0).Attributes.Item(6).nodeValue

			IF CDbl(dTaxAmount) < 0 Then 'IF is an Negative Value
				sTaxEntTy = "C"
			End IF
			dTaxAmount = Abs(dTaxAmount)

			if CInt(sAccCode)>0 Then
				sQuery = "INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
						 "TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
						 ""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"

				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery
			End IF

		End IF

		dim iAdvTranNo,dAdvAmount,dAdjAmtTotal,iPayableNo,iAdvNo,CheckNode,sAdjType
		Dim iAdjRecNo
		dAdjAmtTotal=0


		Response.Write "<br><br>"
		'Response.Clear

		for each EntryNode in oNodAdvRoot.childNodes
			iAdvTranNo = EntryNode.Attributes.Item(0).nodeValue
			dAdvAmount = EntryNode.Attributes.Item(5).nodeValue
			iAdjRecNo = EntryNode.Attributes.Item(6).nodeValue
			sAdjType = EntryNode.Attributes.Item(8).nodeValue

			'===================================================================================
			'Added On	:	17 Nov 2004
			'Added By	:	Manohar Prabhu.R
			'Reason		:	To Take the Advance Number if the Adv No is Not Availabe taking the Value form ythe tabel

			sExp = "//AdvanceDetails[@TransNo="&iAdvTranNo&" and @AdvNo]"
			Set CheckNode = oNodAdvRoot.selectNodes(sExp)

			IF CheckNode.length <> 0 Then
				iAdvNo = EntryNode.Attributes.Item(5).nodeValue
			Else
				iAdvNo = 0
			End IF

			IF CStr(iAdvNo) = "0" and CStr(sAdjType) = "I" Then
				sQuery = "Select AdvanceNumber From Acc_T_AdvancePayments Where TransactionNumber = "&iAdvTranNo&" "

				Response.Write sQuery &"<br><br>"
				objRs.Open sQuery,con
				IF Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close
			End IF
		'=====================================================================================

		if CDbl(dAdvAmount)>0 then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
		end if
		'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
		'----------- INCLUDED ON 04/05/2004 ----------------------

				'sQuery="update Acc_T_AdvancePayments set AdvanceAdjusted="&dAdvAmount &_
				'" where AdvanceNumber = "&iAdvTranNo

				'con.execute(sQuery)


		'------------ END OF UPDATION ---------------------

		'===================================================================================
		'Added On	:	17 Nov 2004
		'Added By	:	Manohar Prabhu.R
		'Reason		:	Updating the Adjustment value with with the Advance Number.
			Dim iCrAdvNo

			IF CStr(sAdjType) = "I" and CDbl(dAdvAmount) <> 0 Then
				sQuery = "update Acc_T_AdvancePayments set AdvanceAdjusted="&dAdvAmount &_
						 " where AdvanceNumber = "&iAdvNo

				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)

				sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo
				Objrs.Open sQuery,Con
				IF Not Objrs.Eof Then
					iCrADvNo = Objrs(0)
				End IF
				Objrs.close

				sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = isNull(AdvanceAdjusted,0) + "&dAdvAmount&" Where CreatedAdvanceNo = "&iCrAdvNo
				con.execute(sQuery)

			Elseif CStr(sAdjType) = "D" and CDbl(dAdvAmount) <> 0 Then
			'===================================================================================
			'Added On	:	29 Nov 2004
			'Added By	:	Manohar Prabhu.R
			'Reason		:	IF the adjusted is a Debot Note

				sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
						 "ReceivedOn,AmountReceived) Values ("&iAdjRecNo&","&iTransNo&","&_
						 "getdate(),"&dAdvAmount&")"

				Response.Write sQuery& "<BR>"

				con.execute(sQuery)
			End IF
		Next

		dim sPayNarration,iCrPayableNo
		'if CDbl(dTotal)>CDbl(dAdjAmtTotal) then
		sQuery="select isnull(PayablesNumber,0) from Acc_T_CreatedPayables where CreatedTransNo = "&iCreatedTransNo

		Response.Write sQuery &"<br><br>"
		objRs.open sQuery,con

		If Not objRs.EOF Then
			iCrPayableNo = objRs(0)
		Else
			iCrPayableNo = 0
		End If
		objRs.Close

		sPayNarration="PUR INV No:"&sPurInvNo &" Dt:"&sPurInvDt

		sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_Payables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close

		'============================================================================================
		sQuery="INSERT INTO Acc_T_Payables(PayablesNumber,TransactionNumber, OUDefinitionID, "&_
				"CreatedPayablesNumber,VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
				"PartyBillDate, AmountPayable, AmountPaid,Narration) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				""&iCrPayableNo&",convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"

		Response.Write sQuery &"<br><br>"
		con.execute(sQuery)
		'============================================================================================

		If cdbl(dAdjAmtTotal) > 0 then
			sQuery="insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
						"PaidOn,AmountPaid) Values ("&iPayableNo&","&iTransNo&","&_
						"getdate(),"&dAdjAmtTotal&")"

			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
		end if

		sQuery = "Update Acc_T_CreatedVoucherHeader set  CreatedVouchStatus='"&sVouStatus&"',   "&_
				 "BookNumber = "&sBookNo&" Where CreatedTransNo="&iCreatedTransNo

		Response.Write sQuery &"<br><br>"
		con.execute(sQuery)

		'end if
	End IF 'IF trim(sPara) = "Acc" then
	IF trim(sPara) = "App" then  'For Approval
	
	    for each EntryNode in oNodAdvRoot.childNodes
			iAdvTranNo = EntryNode.Attributes.Item(0).nodeValue
			dAdvAmount = EntryNode.Attributes.Item(5).nodeValue
			iAdjRecNo = EntryNode.Attributes.Item(6).nodeValue
			sAdjType = EntryNode.Attributes.Item(8).nodeValue

			sExp = "//AdvanceDetails[@TransNo="&iAdvTranNo&" and @AdvNo]"
			Set CheckNode = oNodAdvRoot.selectNodes(sExp)

			IF CheckNode.length <> 0 Then
				iAdvNo = EntryNode.Attributes.Item(5).nodeValue
			Else
				iAdvNo = 0
			End IF

			IF CStr(iAdvNo) = "0" and CStr(sAdjType) = "I" Then
				sQuery = "Select AdvanceNumber From Acc_T_AdvancePayments Where TransactionNumber = "&iAdvTranNo&" "

				Response.Write sQuery &"<br><br>"
				objRs.Open sQuery,con
				IF Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close
			End IF
		'=====================================================================================

		if CDbl(dAdvAmount)>0 then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
		end if
		    Dim sAdjNarr
		    sAdjNarr = "Adjusted to Purchase Invoice No "&sCrVouNo&" DT: "& sVoucDate &" Amt: "& dTotal
		
			IF CStr(sAdjType) = "I" and CDbl(dAdvAmount) <> 0 Then
				sQuery = "update Acc_T_AdvancePayments set AdvanceAdjusted="&dAdvAmount &" where AdvanceNumber = "&iAdvNo

				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)

				sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo
				Objrs.Open sQuery,Con
				IF Not Objrs.Eof Then
					iCrADvNo = Objrs(0)
				End IF
				Objrs.close

				sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = "&dAdvAmount&" Where CreatedAdvanceNo = "&iCrAdvNo
				    Response.Write "<p>"& sQuery
				con.execute(sQuery)
				
				sQuery = "Update Acc_T_AdvancePayments set AdvanceAdjusted = "& dAdvAmount &" where CreatedAdvanceNo = "& iCrAdvNo &" and CreatedTransNo = "& iCrTransNo
				    Response.Write "<p>"& sQuery
				con.execute(sQuery)
				
				
				sQuery = " Select CreatedAdvanceNo from ACC_T_CREATEDADVANCEADJ where CreatedAdvanceNo = "& iCrAdvNo &" and CreatedTransNo = "& iCrTransNo
				objrs.Open sQuery,con,3
				if not objRs.EOF then
				    sQuery =" Update ACC_T_CREATEDADVANCEADJ set AMOUNTADJUSTED = "& dAdvAmount &" where CreatedAdvanceNo = "& iCrAdvNo &" and CreatedTransNo = "& iCrTransNo
				    Response.Write "<p>"& sQuery
				    con.execute sQuery
				else
				    sQuery = "INSERT INTO ACC_T_CREATEDADVANCEADJ (CREATEDADVANCENO, CREATEDTRANSNO, ADJUSTEDON,  "&_
					 "AMOUNTADJUSTED, NARRATION) VALUES "&_
					 "("&iCrAdvNo&", "&iCrTransNo&", Convert(Datetime,'"&sVoucDate&"',103), "&dAdvAmount&", '"&sAdjNarr&"') "

			    Response.Write sQuery &"<br><br>"
			    con.execute(sQuery)
				end if
				objrs.Close 

			Elseif CStr(sAdjType) = "D" and CDbl(dAdvAmount) <> 0 Then

				sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
						 "ReceivedOn,AmountReceived) Values ("&iAdjRecNo&","&iTransNo&","&_
						 "getdate(),"&dAdvAmount&")"

				Response.Write sQuery& "<BR>"

				con.execute(sQuery)
			End IF
		Next

		sQuery="select isnull(PayablesNumber,0) from Acc_T_CreatedPayables where CreatedTransNo = "&iCreatedTransNo

		Response.Write sQuery &"<br><br>"
		objRs.open sQuery,con

		If Not objRs.EOF Then
			iCrPayableNo = objRs(0)
		Else
			iCrPayableNo = 0
		End If
		objRs.Close

		sPayNarration="PUR INV No:"&sPurInvNo &" Dt:"&sPurInvDt
		If cdbl(dAdjAmtTotal) > 0 then
			sQuery="insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
				"PaidOn,AmountPaid) Values ("&iCrPayableNo&","&iCrTransNo&","&_
				"convert(datetime,'"&sVoucDate&"',103),"&dAdjAmtTotal&")"
	        Response.Write "<p>"&sQuery
	        con.execute(sQuery)
		end if
	

		sQuery="update Acc_T_CreatedVoucherHeader set  BookNumber="&sBookNo&",CreatedVouchStatus='"&sVouStatus&"',"&_
			   "ApprovedBy = '"& getUserid() &"',ApprovedOn = Convert(Datetime,'"&dAccDate&"',103),VoucherDate = Convert(Datetime,'"&dAccDate&"',103) "&_
			   "where CreatedTransNo="&iCreatedTransNo &" "

		Response.Write "<p>"&sQuery&"<BR><BR>"
		con.execute(sQuery)

	ElseIF trim(sPara) = "Edt" then 'For Edit
		sQuery = "Update Acc_T_CreatedVoucherHeader set BookNumber = "&sBookNo&",VoucherDate = Convert(Datetime,'"&dAccDate&"',103) Where CreatedTransNo="&iCreatedTransNo

		Response.Write "<p>"&sQuery &"<br><br>"
		con.execute(sQuery)

	End IF
	
	
	If Trim(sPara) = "Edt" or Trim(sPara)="App" then
	    Dim ndMiscRoot,ndMiscAdv,iMiscNo,iMiscAmt,iMiscAccHead,iEntNo,sMiscCrDr
	    
	    sQuery = "Select isNull(Max(VoucherEntryNumber),0) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iCreatedTransNo 
	    objrs.Open sQuery,con
	    if not objRs.EOF then
	        iEntNo = objRs(0)
	    end if
	    objRs.Close 
	    
	    sQuery = "Select CrDrIndication from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iCreatedTransNo 
	    Response.Write "<p>"& sQuery
	    objrs.Open sQuery,con
	    if not objRs.EOF then
	        sMiscCrDr = objRs(0)
	    end if
	    objRs.Close 
	    
	    for each ndMiscRoot in oNodRoot.childNodes
	        if ndMiscRoot.nodeName="MiscAdvanceDetails" then
	            for each ndMiscAdv in ndMiscRoot.childNodes
	                if ndMiscAdv.nodeName="Advance" then
	                    iMiscNo = ndMiscAdv.getAttribute("MiscNo")
	                    iMiscAmt = ndMiscAdv.getAttribute("TobeAdjAmount")
	                    iMiscAccHead = ndMiscAdv.getAttribute("AccHead")
	                    iEntNo = iEntNo + 1
	                    sQuery = "Select * from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iCreatedTransNo &" and MiscTransNo = "& iMiscNo
	                    Response.Write "<p>"& sQuery
	                    objRs.Open sQuery,con,3
	                    if objrs.EOF then
	                        sQuery = "Insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,VoucherEntryNumber,AccountingUnit, "&_
	                                 " AccUnitAccountHead,VoucherNarration,Amount,TransCrDrIndication,MiscTransNo,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode) "&_
	                                 " values("& iCreatedTransNo &","& iEntNo&",'"& sOrgId &"',"& iMiscAccHead &",'Misc Adjustment for Invoice',"& iMiscAmt &",'"& sMiscCrDr &"',"& iMiscNo &",NULL,NULL,NULL)"
	                        Response.Write "<p>"& sQuery
	                        con.execute sQuery
	                        
	                        sQuery  = "Update Acc_T_CreatedVoucherHeader set VoucherAmount = VoucherAmount - "& iMiscAmt &" where CreatedTransNo = "& iCreatedTransNo 
	                        Response.Write "<p>"& sQuery
	                        con.execute sQuery
	                        
	                        sQuery = "Update Acc_T_MiscPymtRequestHeader set AdjustmentStatus ='Y' where MiscTransNo="& iMiscNo 
	                        Response.Write "<p>"& sQuery
	                        con.execute sQuery
	                        
	                    end if
	                                
	                    objRs.Close 
	                end if
	            next
	        end if
	    next
	End if 'If Trim(sPara) = "Edt" or Trim(sPara)="App" then

'Response.End
	'================== Checking for Debit and Credit Tally ===================================
	'Dim dDebTotal,dCreTotal
	'sQuery = "Select VoucherAmount,CrDrIndication From Acc_T_VoucherHeader  "&_
	'		 "Where TransactionNumber = "&iTransNo&" "
	'Objrs.Open sQuery,Con
	'IF Not Objrs.Eof Then
	'	IF Cstr(Objrs(1)) = "C" Then
	'		dCreTotal = Cdbl(Objrs(0))
	'	Else
	'		dDebTotal = Cdbl(Objrs(0))
	'	End IF
	'End IF
	'objrs.Close

	'sQuery = "Select Sum(Amount),TransCrDrIndication From Acc_T_VoucherDetails Where "&_
	'		 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "

	'Objrs.Open sQuery,Con
	'Do While Not Objrs.Eof
	'	IF Cstr(Objrs(1)) = "C" Then
	'		dCreTotal = Cdbl(dCreTotal) + Cdbl(Objrs(0))
	'	Else
	'		dDebTotal = Cdbl(dDebTotal) + Cdbl(Objrs(0))
	'	End IF
	'	Objrs.MoveNext
	'loop
	'objrs.Close

	'sQuery = "Select Sum(TaxAmount),TransCrDrIndication From Acc_T_VoucherTaxDetails Where "&_
	'		 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "

	'Objrs.Open sQuery,Con
	'Do While Not Objrs.Eof
	'	IF Cstr(Objrs(1)) = "C" Then
	'		dCreTotal = Cdbl(dCreTotal) + Cdbl(Objrs(0))
	'	Else
	'		dDebTotal = Cdbl(dDebTotal) + Cdbl(Objrs(0))
	''	End IF
	'	Objrs.MoveNext
	'loop
	'objrs.Close



	'IF Cdbl(dCreTotal) <> Cdbl(dDebTotal) Then
	'	Response.Clear
	'	Response.Write "Total Credit and Debit Does'nt Match Transaction is been Deleted "
	'	Response.Write "<br>"
	'	Response.Write "Debit Total " & FormatNumber(dDebTotal,2,,,0) &" Credit Total " & FormatNumber(dCreTotal,2,,,0)
	'	Con.RollbackTrans
	'	Response.End
	'End IF

		Dim sTbGLVal,sTempVal,sDiffVal,dTBDRAmt,dTBCRAmt
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
		
		


		IF CDbl(dTBDRAmt) <> CDbl(dTBCRAmt) Then
			con.RollbackTrans

			sErrChk = "F"
			'Response.Clear
			Response.Write "<h3>Debit and Credit is Not Matching </h3><br>"
			Response.Write "<b>Debit Amount  :--> </b>" & dTBDRAmt &"<br>"
			Response.Write "<b>Credit Amount :--> </b>" & dTBCRAmt &"<br>"
			Response.Write "<b>Differences   :--> </b>" & sDiffVal &"<br>"
			Response.Write "<h3> Voucher Not Created " &"<br>"
			Response.End

		End IF

	'================== Checking for Debit and Credit Tally ===================================

	
'	Con.RollbackTrans
'	Response.End
	
	Response.clear
	con.CommitTrans
	IF trim(sPara) <> "App" then
		'******************************** Cost and ANal Procedure Call ************************

		dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode

		sExp = "//Entry"
		Set HeaderNode = oNodRoot.selectNodes(sExp)
		IF HeaderNode.length <> 0 Then

			for iCtr = 0 To HeaderNode.length - 1

				sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value
				'sEntryType=HeaderNode.Item(iCtr).Attributes.getNamedItem("CRDR").value
				'sAmount=HeaderNode.Item(iCtr).Attributes.getNamedItem("Amount").value
				'sAccUnit=HeaderNode.Item(iCtr).Attributes.getNamedItem("AccUnit").value

				sExp="//Entry[@No='"&sEntryno&"']/AccHead"
				set tempNode=oNodRoot.selectNodes(sExp)
				sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

				set nodANL=oDOM.createElement("Root")
				nodANL.setAttribute "TransNo",iTransNo
				nodANL.setAttribute "EntryNo",sEntryno
				nodANL.setAttribute "UnitCode", sorgid
				nodANL.setAttribute "GlHead",sAccCode
				nodANL.setAttribute "ACTFlag","V"


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

		'******************************** Coast and ANal Procedure Call ************************

	End IF 'IF trim(sPara) <> "App" then
End IF
'Response.clear
	'----To code for Tax Form deatils----------------
	'Response.Write  "con.Errors.count="&con.Errors.count
	if con.Errors.count <>0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count-1
			Response.Write con.Errors(iCounter).Description &"<br>"
			Response.Write con.Errors(iCounter).Source &"<br>"
		next
		'Redirect to Error Handling System
	else

		'con.CommitTrans

		Set newElem  = oDOM.createAttribute("TransNo")
		newElem.value = iTransNo
		oNodRoot.setAttributeNode(newElem)

		Set newElem  = oDOM.createAttribute("VouNo")
		newElem.value = iVouNo
		oNodRoot.setAttributeNode(newElem)

		'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

		'	Response.Redirect ("AppOtherVoucherList.asp?optCriteria=Exist&selUnitId="&sOrgId&"&selVoucher="&sVouCode&"&selBook="&sBookNo&"&selApplication=2")
			Response.Redirect ("PURCHASEVOUCHERS.ASP")


	end if


%>

