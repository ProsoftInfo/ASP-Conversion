<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccPurVouGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January  31, 2003
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
dim sParCode,sParSubType,sParType
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
Dim sRoundoffHead,dRoundOff
Dim bAddFlag,sVouEntTy,sNewNarr,dInvWtRnd,sFormVal,sSelVouTy,sRetVal

sVouName=Request("hVouName")
sVouCode="04"
sVouType="C"
iCreatedTransNo = Request("hTransNo")
bPartyReceiptFlag=false
sAccUnit = Request.Form("selUnitId")
sFormVal = Request("hFormVal")
sSelVouTy = Request("voutype")
sVouStatus="010104" 'Voucher Accounted

SET objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
SET oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objFSO = CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(server.MapPath("../../Purchase/xmldata/General.xml")) then
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
end if 'if objFSO.FileExists(server.MapPath("../../Purchase/xmldata/General.xml")) then

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")
sRetVal = GetVouchXML(iCreatedTransNo)
oDOM.Load server.MapPath(sRetVal)
oDOM.Save server.MapPath("../temp/transaction/Temp.xml")

sQuery = "Select CreatedTransNo From Acc_T_VoucherHeader Where CreatedTransNo = "&iCreatedTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouEntTy = "Y" 'Already Accounted
Else
	sVouEntTy = "N" 'Can Be Accounted
End IF
objRs.Close

'Response.write sVouEntTy
'Response.end

IF CStr(sVouEntTy) = "N" Then

	set oNodRoot=oDOM.documentElement
	sCrVouNo = oNodRoot.Attributes.Item(1).nodeValue
	iCrTransNo = oNodRoot.Attributes.Item(0).nodeValue

	sExp = "//AccHead"
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		sAccCode = TempNode.Item(0).Attributes.getNamedItem("No").Value
	End IF

	sExp = "//Tax[@CatCode=0 and @TaxCode=0]"
	Set CheckNode = oNodRoot.selectNodes(sExp)
	IF CheckNode.length <> 0 Then
		sRoundoffHead = CheckNode.Item(0).Attributes.getNamedItem("AccHead").Value
		dRoundOff = CheckNode.Item(0).Attributes.getNamedItem("TaxAmount").Value
	End IF

	Response.Write "<p> RoundoffHead="& sRoundoffHead&"<P>"
	'Response.End

	IF cstr(dRoundOff) = "" Then
		dRoundOff = 0
	End IF

	IF CStr(sRoundoffHead) = "0" Then
		sExp = "//Entry"
		Set CheckNode = oNodRoot.selectNodes(sExp)
		IF CheckNode.length <> 0 Then
			dInvWtRnd = CheckNode.Item(0).Attributes.Item(2).NodeValue
			Response.Write dInvWtRnd &"================="
			Response.write dRoundOff &"============="
			dInvWtRnd = CDbl(dInvWtRnd) + CDbl(dRoundOff)
			CheckNode.Item(0).Attributes.Item(2).NodeValue  = dInvWtRnd
		End IF
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
			IF oNodTemp.Attributes.Item(0).nodeValue <> "" Then
				set oNodDeatils=oNodTemp
			End IF
		end if

		if oNodTemp.nodeName="TaxDetails" then
			set oNodTaxRoot=oNodTemp
			'dTotal=oNodTemp.Attributes.Item(0).nodeValue
			dTotal=oNodTemp.Attributes.Item(2).nodeValue
			dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
		end if
		if oNodTemp.nodeName="AdvanceDetails" then
			set oNodAdvRoot=oNodTemp
		end if


	next

	IF Cstr(sAccUnit) = "S" Then
		sAccUnit = sOrgId
	End IF

	IF CStr(sAccUnit) = "" Then
		sAccUnit = sOrgId
	End IF

	sTransType="PJR"

	sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

	sQuery="select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo

	objRs.open sQuery,con
	if not objRs.EOF then
		iSeriesNo=objRs(0)
		iSeriesCode=objRs(1)

	end if
	objRs.close()

	con.BeginTrans

	'Response.Write sOrgId&" "&iSeriesNo&" "&iSeriesCode&" "&sVoucDate

	iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

	'============================================================================================
	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close
	'============================================================================================

	'============================================================================================
	
	sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
	sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"
	sQuery=sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"','"&sCrVouNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)
'	Response.End 
	'============================================================================================

	for each EntryNode in oNodDeatils.childNodes
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc

		sEntryno=EntryNode.Attributes.Item(0).nodeValue
		dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
		sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
		dQty=EntryNode.Attributes.Item(3).nodeValue
		sUOM=EntryNode.Attributes.Item(4).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		dRate=EntryNode.Attributes.Item(6).nodeValue
		dDisPer=EntryNode.Attributes.Item(8).nodeValue
		dDisAmount=EntryNode.Attributes.Item(9).nodeValue

		sNewNarr = sNewNarr&", "&sItemDesc

		sEntryType="D"
		sNarration=""

		for each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					sAccCode=HeaderNode.Attributes.Item(0).nodeValue
					sAccType=HeaderNode.Attributes.Item(4).nodeValue
			end if 'End of Check for Account head Node
		next 'End of Entry Node Loop


		'IF CStr(sEntryno) = "1" Then
		'	Response.Write sAmount &"<br><br>"
		'	sAmount = CDbl(sAmount) + CDbl(dRoundOff)
		'End IF



	'============================================================================================
			sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
			sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)

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

		IF Cstr(dTaxAmount) = "" Then
			dTaxAmount = 0
		End IF

		if Cdbl(dTaxAmount) >=0 then
			'sEntryType = "C"
			sEntryType = "D"
		else
			'sEntryType = "D"
			sEntryType = "C"
			dTaxAmount=dTaxAmount*-1
		End if

		IF CStr(sAccCode) = "" Then
			sAccCode = 0
		ENd IF

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


			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
			iSno=cint(iSno)+1
		end if
	Next

	dim iAdvTranNo,dAdvAmount,dAdjAmtTotal,iPayableNo,iAdvNo,CheckNode,sAdjType
	Dim iAdjRecNo
	dAdjAmtTotal=0


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

			'sQuery="update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
			'" where AdvanceNumber = "&iAdvTranNo

			'con.execute(sQuery)


	'------------ END OF UPDATION ---------------------

	'===================================================================================
	'Added On	:	17 Nov 2004
	'Added By	:	Manohar Prabhu.R
	'Reason		:	Updating the Adjustment value with with the Advance Number.

		IF CStr(sAdjType) = "I" and CDbl(dAdvAmount) <> 0 Then
			sQuery = "update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
					 " where AdvanceNumber = "&iAdvNo

			con.execute(sQuery)
		Elseif CStr(sAdjType) = "D" and CDbl(dAdvAmount) <> 0 Then
	'===================================================================================
	'Added On	:	29 Nov 2004
	'Added By	:	Manohar Prabhu.R
	'Reason		:	IF the adjusted is a Debot Note

			sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
					 "ReceivedOn,AmountReceived) Values ("&iAdjRecNo&","&iTransNo&","&_
					 "getdate(),"&dAdvAmount&")"

			'Response.Write sQuery& "<BR>"

			con.execute(sQuery)
		End IF

	'===============================================================================================

	Next

	dim sPayNarration,iCrPayableNo
	'if CDbl(dTotal)>CDbl(dAdjAmtTotal) then
		sQuery="select isnull(PayablesNumber,0) from Acc_T_CreatedPayables where CreatedTransNo = "&iCreatedTransNo
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
		'Response.Write sQuery& "<BR>"
		con.execute(sQuery)
		'============================================================================================

	If cdbl(dAdjAmtTotal) > 0 then
		sQuery="insert into Acc_T_PybleAdjustmentDetails(PayablesNumber,PaidByTransactionNo,"&_
					"PaidOn,AmountPaid) Values ("&iPayableNo&","&iTransNo&","&_
					"getdate(),"&dAdjAmtTotal&")"
		con.execute(sQuery)
	end if

	'end if

	sQuery="update Acc_T_CreatedVoucherHeader set  CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iCreatedTransNo
	con.execute(sQuery)

	sNewNarr = Trim(sNewNarr)
	sNewNarr = Mid(sNewNarr,2)
	sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
			 "TransactionNumber = "&iTransNo&" "
	con.execute(sQuery)

	sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNewNarr&"' Where  "&_
			 "TransactionNumber = "&iTransNo&" "
	con.execute(sQuery)

'================== Checking for Debit and Credit Tally ===================================
	Dim dDebTotal,dCreTotal
	sQuery = "Select VoucherAmount,CrDrIndication From Acc_T_VoucherHeader  "&_
			 "Where TransactionNumber = "&iTransNo&" "
    Response.Write "<p>"&sQuery
    Objrs.Open sQuery,Con
	IF Not Objrs.Eof Then
		IF Cstr(Objrs(1)) = "C" Then
			dCreTotal = FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		Else
			dDebTotal = FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		End IF
	End IF
	objrs.Close

	sQuery = "Select Sum(Amount),TransCrDrIndication From Acc_T_VoucherDetails Where "&_
			 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "
    Response.Write "<p>"&sQuery
	Objrs.Open sQuery,Con
	Do While Not Objrs.Eof
	    Response.Write "<p>"& FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		IF Cstr(Objrs(1)) = "C" Then
			dCreTotal = Cdbl(dCreTotal) + FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		Else
			dDebTotal = Cdbl(dDebTotal) + FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		End IF
		Objrs.MoveNext
	loop
	objrs.Close

	sQuery = "Select Sum(TaxAmount),TransCrDrIndication From Acc_T_VoucherTaxDetails Where "&_
			 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "
Response.Write "<p>"&sQuery
	Objrs.Open sQuery,Con
	Do While Not Objrs.Eof
	    Response.Write "<p>"& FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		IF Cstr(Objrs(1)) = "C" Then
			dCreTotal = Cdbl(dCreTotal) + FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		Else
			dDebTotal = Cdbl(dDebTotal) + FormatNumber(Cdbl(Objrs(0)),2,0,0,0)
		End IF
		Objrs.MoveNext
	loop
	objrs.Close

Response.Write "<p>"
Response.Write "<p>"& FormatNumber(Cdbl(dCreTotal),2,0,0,0)
Response.Write "<p>"& FormatNumber(Cdbl(dDebTotal),2,0,0,0)
Response.Write "<p>"& FormatNumber(Cdbl(dCreTotal),2,0,0,0) - FormatNumber(Cdbl(dDebTotal),2,0,0,0)

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
'Con.RollbackTrans
'Response.End

	con.CommitTrans

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






	'----To code for Tax Form deatils----------------

	if con.Errors.count <>0 then
		con.RollbackTrans

		for iCounter=0 to con.Errors.count-1
			Response.Write con.Errors(iCounter).Description &"<br>"
			Response.Write con.Errors(iCounter).Source &"<br>"
		next
		'Redirect to Error Handling System
	else

		'con.CommitTrans

		'Set newElem  = oDOM.createAttribute("TransNo")
		'newElem.value = iTransNo
		'oNodRoot.setAttributeNode(newElem)

		'Set newElem  = oDOM.createAttribute("VouNo")
		'newElem.value = iVouNo
		'oNodRoot.setAttributeNode(newElem)

		'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
		Response.Redirect ("PurchaseVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)

	end if
Else
	Response.Redirect ("PurchaseVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
End IF

%>

