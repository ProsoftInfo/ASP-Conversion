<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccSalVouGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	April 14, 2004
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
<%
'XML DOM Variables
Dim oDOM,objFSO,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType
dim EntryNode,HeaderNode,nodANL,newElem,sExp
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus
dim iSeriesNo,iSeriesCode
dim sSalType,sRefernceNo,iCounter
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sSalInvNo,sSalInvDt,dRoundOff,sRoundoffHead
dim iCrTransNo,sCrVouNo,iCrSeriesNo,iCrSeriesCode
Dim sAppTy,sAppBy,sStr,Tempnode,sPayNarration,bAddFlag,sVouEntTy
Dim sNewNarr,dRndOFf,dInvWtRnd,sFormVal,sSelVouTy,sCallTy,sRetVal

sFormVal = Request("hFormVal")
sSelVouTy = Request("voutype")
sCallTy = Request("CallType")



dim dAdjAmtTotal
sVouStatus="010104" 'Voucher Accounted
sVouCode="05"
sVouType="D"
iCrTransNo=Request("hTransNo")
sTransType="SJR"
sAccUnit = Request("selUnitId")

set objRs  = server.CreateObject("adodb.recordset")
set objFSO = CreateObject("Scripting.FileSystemObject")
'sQuery = "Select OUDefinitionID From Acc_T_CreatedVoucherHeader Where TransactionNumber = "&iCrTransNo&" "
'Objrs.Open sQuery,Con
'IF Not Objrs.eof Then
'	sAccUnit = Trim(Objrs(0))
'End IF
'Objrs.close

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
IF objFSO.FileExists(server.MapPath("../../Sales/xmldata/General.xml")) THEN
    oDOM.Load server.MapPath("../../Sales/xmldata/General.xml")
    Set oNodRoot = oDOM.documentElement
    sExp = "//ROUNDOFF"
    Set oNodEntry = oNodRoot.Selectnodes(sExp)
    If oNodEntry.Length > 0 then
	    sRoundoffHead = oNodEntry.Item(0).Attributes.Item(0).nodevalue
    else
	    sRoundoffHead = "0"
    end if
ELSE
    sRoundoffHead = "0"
END IF ' IF objFSO.FileExists(server.MapPath("../../Sales/xmldata/General.xml")) THEN

'oDOM.load server.MapPath("../temp/transaction/Voucher Entry_SAL_"&Session.SessionID&".xml")
'oDOM.Load server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
sRetVal = GetVouchXML(iCrTransNo)
oDOM.Load server.MapPath(sRetVal)

'oDOM.Save server.MapPath("../temp/transaction/Temp_"&Session.SessionID&".xml")
sQuery = "Select CreatedTransNo From Acc_T_VoucherHeader Where CreatedTransNo = "&iCrTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouEntTy = "Y" 'Already Accounted
Else
	sVouEntTy = "N" 'Can Be Accounted
End IF
objRs.Close

IF CStr(sVouEntTy) = "N" Then

	set oNodRoot=oDOM.documentElement
	sExp = "//Organization"
	Set Tempnode = oNodRoot.selectNodes(sExp)
	IF Tempnode.length <> 0 Then
		Tempnode.Item(0).setAttribute "AccUnit",sAccUnit
	End IF

	for each oNodTemp in oNodRoot.childNodes
		if oNodTemp.nodeName="Header" then
			for Each oNodEntry in  oNodTemp.childNodes
				if oNodEntry.nodeName="Organization" then
					sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				end if
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
					sAppTy = oNodEntry.Attributes.Item(2).nodeValue
					sAppBy = oNodEntry.Attributes.Item(4).nodeValue

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
	
	
	sExp = "//Tax[@CatCode=0 and @TaxCode=0]"
	Set CheckNode = oNodRoot.selectNodes(sExp)
	IF CheckNode.length <> 0 Then
		sRoundoffHead = CheckNode.Item(0).Attributes.getNamedItem("AccHead").Value
		dRoundOff = CheckNode.Item(0).Attributes.getNamedItem("TaxAmount").Value
	End IF
	
	IF CStr(sRoundoffHead) = "0" Then
		sExp = "//Entry"
		Set CheckNode = oNodRoot.selectNodes(sExp)
		IF CheckNode.length <> 0 Then
			dInvWtRnd = CheckNode.Item(0).Attributes.Item(2).NodeValue 
			dInvWtRnd = CDbl(dInvWtRnd) + CDbl(dRoundOff)
			CheckNode.Item(0).Attributes.Item(2).NodeValue  = dInvWtRnd
		End IF
	End IF

	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

	con.BeginTrans

	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close

'=================== Blocked by Manohar on 20/12/2006 For Sangeeth Since Interunit is Not Defined ate this time
'	IF CStr(sAccUnit) = "S" Then
'			sAccUnit = sOrgId
'	End IF
'
'	IF CStr(sAccUnit) = "" Then
'		sAccUnit = sOrgId
'	End IF
'=================== Blocked by Manohar on 20/12/2006 For Sangeeth Since Interunit is Not Defined ate this time

	sAccUnit = sOrgId
	'=====================================================================================================================================================
	sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
	sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus,AccountedOn) values"
	sQuery=sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sSalInvNo&"','"&sSalInvNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"',getdate())"


	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)
'	Response.End 
	'=====================================================================================================================================================


	for each EntryNode in oNodDeatils.childNodes
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc

		sEntryno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		sItemDesc=Replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
		dQty=EntryNode.Attributes.Item(3).nodeValue
		sUOM=EntryNode.Attributes.Item(4).nodeValue
		dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
		dRate=EntryNode.Attributes.Item(6).nodeValue
		dDisPer=EntryNode.Attributes.Item(8).nodeValue
		dDisAmount=EntryNode.Attributes.Item(9).nodeValue
		dRndOFf = EntryNode.Attributes.Item(10).nodeValue
		
		Response.Write "<p>sAmount = "& sAmount
		
		IF CStr(dRndOFf) = "" Then
			dRndOFf = 0
		End IF
		
		sNewNarr = sNewNarr&", "&sItemDesc
		
		sEntryType="C"
		sNarration=""

		for each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					sAccCode=HeaderNode.Attributes.Item(0).nodeValue
					sAccType=HeaderNode.Attributes.Item(4).nodeValue
			end if 'End of Check for Account head Node
		next 'End of Entry Node Loop

		sEntryno = CDbl(sEntryno) + 1
		
	'=====================================================================================================================================================
			sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,RoundOffvalue ) values ("
			sQuery=sQuery& iTransNo&",'"&sOrgId&"'"
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","&dRndOFf&")"
			
			Response.Write "<p>"&sQuery& "<BR><BR>"
			con.execute(sQuery)
	'=====================================================================================================================================================
	
	'=================== Updation of Narration in GL  ================================
	    sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sItemDesc&"' Where  "&_
			     "TransactionNumber = "&iTransNo&" AND VoucherEntryNumber="& sEntryno 
			     Response.Write "<p>"& sQuery
	    con.execute(sQuery)
    	
    	
	    '=================== Updation of Narration in GL  ================================
	
	
	iSno=cint(iSno)+1
	Next

	iSno=1
	dim sTaxMode
	for each EntryNode in oNodTaxRoot.childNodes
		iCatCode=EntryNode.Attributes.Item(0).nodeValue
		iTaxCode=EntryNode.Attributes.Item(1).nodeValue
		sTaxMode=EntryNode.Attributes.Item(2).nodeValue
		dTaxPer=EntryNode.Attributes.Item(4).nodeValue
		dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
		sAccCode=EntryNode.Attributes.Item(6).nodeValue
		if dTaxAmount >=0 then
			sEntryType = "C"
			'sEntryType = "D"
		else
			sEntryType = "D"
			'sEntryType = "C"
			dTaxAmount=dTaxAmount*-1
		End if
		if CInt(sAccCode)>0 then
			if sTaxMode="P" then
				sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
			else

				sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
			end if


			Response.Write sQuery& "<BR><BR>"

			con.execute(sQuery)
			iSno=cint(iSno)+1
		end if
	Next


	'Response.Clear

	dim iAdvTranNo,dAdvAmount,iAdvNo,CheckNode,iPayNo,sAdjType,iAccPayNo

	dAdjAmtTotal=0
	if  IsObject(oNodAdvRoot) then
		for each EntryNode in oNodAdvRoot.childNodes
			iAdvTranNo = Trim(EntryNode.Attributes.Item(0).nodeValue)
			dAdvAmount = Trim(EntryNode.Attributes.Item(5).nodeValue)
			iAccPayNo = EntryNode.Attributes.Item(6).nodeValue
			sAdjType = EntryNode.Attributes.Item(8).nodeValue

	'===================================================================================
	'Added On	:	17 Nov 2004
	'Added By	:	Manohar Prabhu.R
	'Reason		:	To Take the Advance Number if the Adv No is Not Availabe taking the Value form ythe tabel

		IF CStr(sAdjType) = "B" and CDbl(dAdvAmount) <> 0 Then
			sExp = "//AdvanceDetails[@TransNo="&iAdvTranNo&" and @AdvNo]"
			Set CheckNode = oNodAdvRoot.selectNodes(sExp)

			IF CheckNode.length <> 0 Then
				iAdvNo = EntryNode.Attributes.Item(5).nodeValue
			Else
				iAdvNo = 0
			End IF

			IF CStr(iAdvNo) = "0" Then
				sQuery = "Select AdvanceNumber From Acc_T_AdvancePayments Where TransactionNumber = "&iAdvTranNo&" "
				objRs.Open sQuery,con
				IF Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close
			End IF
	'=====================================================================================


			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

	'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
	'----------- INCLUDED ON 05/05/2004 ----------------------

			sQuery = "update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
					 " where AdvanceNumber = "&iAdvNo

			'Response.Write sQuery &"<br><br>"

			con.execute(sQuery)
		'Elseif CStr(sAdjType) = "SR" or Cstr(sAdjType) = "SC" and CDbl(dAdvAmount) <> 0 Then
		Elseif CStr(sAdjType) = "SR" or Cstr(sAdjType) = "SC" or Cstr(sAdjType) = "OT" Then
			IF CDbl(dAdvAmount) <> 0  Then
				dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
				sQuery = "insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
					     "PaidOn,AmountPaid) Values ("&iAccPayNo&","&iTransNo&","&_
						 "getdate(),"&dAdvAmount&")"
				Response.Write sQuery &"<br><br> =="
				con.execute(sQuery)
			End IF
		End IF
	'------------ END OF UPDATION ---------------------

			'sQuery="update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
			'" where TransactionNumber="&iAdvTranNo
			'con.execute(sQuery)
		Next
	end if
	dim iCrPayableNo
	sPayNarration="SALE INV NO:"&sSalInvNo&" Dt:"&sSalInvDt

	sQuery="select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "&iCrTransNo
	objRs.open sQuery,con

	If Not objRs.EOF Then
		iCrPayableNo = objRs(0)
	Else
		iCrPayableNo = 0
	End If
	objRs.Close

	'==========================================================================================================================
	sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_Receivables"
	objRs.open sQuery,con
		iPayableNo=objRs(0)
	objRs.Close

	sQuery="INSERT INTO Acc_T_Receivables(ReceivableNumber,TransactionNumber, OUDefinitionID, "&_
			"VoucherDate, PartyType, PartySubType, PartyCode,PartyInvoiceNumber, "&_
			"CreatedReceivable,PartyInvoiceDate, AmountReceivable, AmountReceived,Narration) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sSalInvNo&"',"&_
			""&iCrPayableNo&",convert(datetime,'"&sSalInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"

	'Response.Write sQuery& "<BR>"
	con.execute(sQuery)
	'==========================================================================================================================

	Response.Write dAdjAmtTotal &"<br><br>"

	if CDbl(dAdjAmtTotal) > 0  then
		sQuery="insert into Acc_T_RcvblAdjustmentDetails (ReceivableNumber,RecdByTransactionNo,"&_
					"ReceivedOn,AmountReceived) Values ("&iPayableNo&","&iTransNo&","&_
					"convert(datetime,'"&sSalInvDt&"',103),"&dAdjAmtTotal&")"
		Response.Write sQuery & "<br><br>"
		con.execute(sQuery)

	end if
	
	
'***************** Blocked By Manohar on 03/May/08 For Since the Agent Details is Available In the Created Voucher
'***************** stage itself only updating the Account Trans Number is Needed.
	'=======================================================================================================================
	'dim iAgentCode,sCommType,dCommPer,dCommAmount,iCurrencyCode,sAgType,sAgParType
	'if  IsObject(oNodAgent) then
	'	for each EntryNode in oNodAgent.childNodes
			'iAgentCode=EntryNode.Attributes.Item(0).nodeValue
			'sCommType=EntryNode.Attributes.Item(2).nodeValue
			'dCommPer=EntryNode.Attributes.Item(3).nodeValue
			'dCommAmount=EntryNode.Attributes.Item(4).nodeValue
			'sAgType = EntryNode.Attributes.Item(5).nodeValue
			'sAgParType = EntryNode.Attributes.Item(6).nodeValue
			
			'sQuery = "Update Sal_T_AdditionalAgents Set AccTransactionNo = "&iTransNo&" Where ACCCrTransNo = "&iCrTransNo
			
			
			'sQuery="INSERT INTO Sal_T_AdditionalAgents(AgentCode, CommissionType,CurrencyCode, "&_
			'		"AccTransactionNo, AgentCommission,CommissionToPay,AgentType,AgentSubType) "&_
			'		"values ("&iAgentCode&",'"&sCommType&"',1,"&_
			'		""&iTransNo&","&dCommAmount&",'1','"&sAgType&"',"&sAgParType&")"

		'Response.Write sQuery& "<BR>"
	'	con.execute(sQuery)
	'	Next
	'end if
	'=======================================================================================================================
	'Response.Clear
	sExp = "//AgentDetails/Agent"
	Set Tempnode = oNodRoot.selectNodes(sExp)
	Response.Write Tempnode.length
	IF Tempnode.length > 0 Then
		sQuery = "Update Sal_T_AdditionalAgents Set AccTransactionNo = "&iTransNo&" Where ACCCrTransNo = "&iCrTransNo
		Response.Write sQuery &"<br><br><br>"
		con.execute(sQuery)
	End IF
	
	sQuery="update Acc_T_CreatedVoucherHeader set  CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iCrTransNo
	con.execute(sQuery)
	
	'Con.RollbackTrans
	'Response.End
	
	
	'=================== Updation of Narration in Party Tables ================================
	sNewNarr = Trim(sNewNarr)
	sNewNarr = Mid(sNewNarr,2)
	sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
			 "TransactionNumber = "&iTransNo&" "
	con.execute(sQuery)
	
	'=================== Updation of Narration in Party Tables ================================
	
	'================== Checking for Debit and Credit Tally ===================================
	Dim dDebTotal,dCreTotal
	sQuery = "Select VoucherAmount,CrDrIndication From Acc_T_VoucherHeader  "&_
			 "Where TransactionNumber = "&iTransNo&" "
	Objrs.Open sQuery,Con
	IF Not Objrs.Eof Then
		IF Cstr(objRs(1)) = "C" Then
			dCreTotal = FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		Else
			dDebTotal = FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		End IF
	End IF
	objrs.Close

	sQuery = "Select Sum(Amount),TransCrDrIndication From Acc_T_VoucherDetails Where "&_
			 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "

	Objrs.Open sQuery,Con
	Do While Not Objrs.Eof
		IF Cstr(Objrs(1)) = "C" Then
			dCreTotal = Cdbl(dCreTotal) +  FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		Else
			dDebTotal = Cdbl(dDebTotal) +  FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		End IF
		Objrs.MoveNext
	loop
	objrs.Close

	sQuery = "Select Sum(TaxAmount),TransCrDrIndication From Acc_T_VoucherTaxDetails Where "&_
			 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "

	Objrs.Open sQuery,Con
	Do While Not Objrs.Eof
		IF Cstr(Objrs(1)) = "C" Then
			dCreTotal = Cdbl(dCreTotal) +  FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		Else
			dDebTotal = Cdbl(dDebTotal) +  FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		End IF
		Objrs.MoveNext
	loop
	objrs.Close

    dCreTotal = FormatNumber(Cdbl(dCreTotal),2,0,0,0)
    dDebTotal = FormatNumber(Cdbl(dDebTotal),2,0,0,0)

	IF Cdbl(dCreTotal) <> Cdbl(dDebTotal) Then
		'Response.Clear
		Response.Write "Total Credit and Debit Does'nt Match Transaction is been Deleted "
		Response.Write "<br>"
		Response.Write "Debit Total " & FormatNumber(dDebTotal,2,,,0) &" Credit Total " & FormatNumber(dCreTotal,2,,,0)
		Con.RollbackTrans
		Response.End
	End IF
		
'================== Checking for Debit and Credit Tally ===================================

   ' Con.RollbackTrans
   ' Response.End
	
	con.CommitTrans
	
	'******************************** Coast and ANal Procedure Call ************************


	'----To code for Tax Form deatils----------------

	if con.Errors.count <>0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) &"<br>"
		next
		'Redirect to Error Handling System
	else

		'con.CommitTrans
		'con.RollBackTrans

		'sStr = "//Voucher"
		'Set Tempnode = oNodRoot.selectNodes(sStr)
		'IF Tempnode.length <> 0 Then
		'	Tempnode.Item(0).Attributes.Item(2).value = iTransNo
		'End IF

		'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
		IF CStr(sCallTy) <> "A" Then
			Response.Redirect ("SalesVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
		Else
			Response.Redirect ("ACCSALVOUCHERS.ASP")
		End IF


	end if
Else
	Response.Redirect ("SalesVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
End IF
%>