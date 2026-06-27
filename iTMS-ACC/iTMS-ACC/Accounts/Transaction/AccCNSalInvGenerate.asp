<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccCNSalInvGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 18, 2003
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
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,iCounter
dim iSeriesNo,iSeriesCode,iSalTrNo
dim sSalType, ActualTransNo,iCrTransNo
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
Dim TaxCalNode,dTaxValNoAcc,iCtr,dEachItmVal,dEachItm,dNoAccTotVal,dDiffVal
dim dAdjAmtTotal,sAppTy,sAppBy,sExp,NarrNode,sCrVouNo,sVouEntTy
Dim sSelVouTy,sFormVal,sRetVal

sSelVouTy = Request("voutype")
sFormVal = Request("hFormVal")


sVouStatus = "010104"

sVouCode="07"
sVouType="C"



iCrTransNo = Request("hTransNo")

set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
'oDOM.Load server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
sRetVal = GetVouchXML(iCrTransNo)
oDOM.Load server.MapPath(sRetVal)

set oNodRoot=oDOM.documentElement
sCrVouNo = oNodRoot.Attributes.Item(1).nodeValue

sQuery = "Select CreatedTransNo From Acc_T_VoucherHeader Where CreatedTransNo = "&iCrTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouEntTy = "Y" 'Already Accounted
Else
	sVouEntTy = "N" 'Can Be Accounted
End IF
objRs.Close

IF CStr(sVouEntTy) = "N" Then

	for each oNodTemp in oNodRoot.childNodes
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
					iSalTrNo = oNodEntry.Attributes.Item(5).nodeValue
				end if
				if oNodEntry.nodeName="SalesType" then
					sSalType=oNodEntry.Attributes.Item(0).nodeValue
				end if
			next
		end if
		if oNodTemp.nodeName="Details" then
			set oNodDeatils=oNodTemp
			sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue
		end if
		if oNodTemp.nodeName="TaxDetails" then
			set oNodTaxRoot=oNodTemp
			dTotal=oNodTemp.Attributes.Item(0).nodeValue
		end if
	next

	sExp = "//Narration"
	Set NarrNode = oNodRoot.selectNodes(sExp)
	IF NarrNode.Length <> 0 Then
		sNarration = NarrNode.Item(0).text
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

	sTransType="CNR"

	

	dim sAccHeadCode,iTemp

	sQuery="select DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
		
	Response.Write sQuery

	objRs.open sQuery,con
	if not objRs.EOF then
		iSeriesNo=objRs(0)
		iSeriesCode=objRs(1)
	end if
	objRs.close()

	con.BeginTrans

	iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close
	
	sSalType = 0

	sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
	sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
	sQuery=sQuery&"PayToRecdFrom,CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"
	sQuery=sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
	sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"','"&sCrVouNo&"',"&iCrTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal
	sQuery=sQuery&",'Sales Invoices ','"&sVouType&"','OS',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"

	Response.Write sQuery& "<BR><BR><BR>"
	con.execute(sQuery)

	iCtr = 1
	for each EntryNode in oNodDeatils.childNodes
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc

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
		'sNarration=""
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

			sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
			sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
			sQuery=sQuery& iTransNo&",'"&sOrgId&"'"
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
			sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"
		Response.Write sQuery& "<BR><BR><BR>"
		con.execute(sQuery)
	End IF

	next'End of Voucher Node Loop
	'Response.Clear
	'-----------------------------TAX TABLE UPDATE------------------
	iSno=1
	dim sTaxMode,sTaxTrans
	for each EntryNode in oNodTaxRoot.childNodes
		iCatCode=EntryNode.Attributes.Item(0).nodeValue
		iTaxCode=EntryNode.Attributes.Item(1).nodeValue
		sTaxMode=EntryNode.Attributes.Item(2).nodeValue
		dTaxPer=EntryNode.Attributes.Item(4).nodeValue
		dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
		sAccCode=EntryNode.Attributes.Item(6).nodeValue
		
		Response.Write dTaxAmount &"<br><br><br>"
		
		IF CDbl(dTaxAmount) >= 0 Then
			sTaxTrans = "D"
		Else
			sTaxTrans = "C"
			sTaxMode = "F"
		End IF
		
		dTaxAmount = Abs(dTaxAmount)
		
		IF CStr(sAccCode) = "" Then
			sAccCode = 0
		End IF
		
	if	CDbl(sAccCode) > 0 then
		if sTaxMode="P" then
			sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iTransNo&","&sAccCode&",'"&sTaxTrans&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
		else
			sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iTransNo&","&sAccCode&",'"&sTaxTrans&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
		end if
	end if
	iSno=cint(iSno)+1
	Next

	'----TO CODE FOR TAX FORM DEATILS----------------

	'----------------------------- RECEIVABLE TABLE UPDATE------------------
		'sQuery="select isnull(ReceivableNumber,0) from Acc_T_CreatedRcvbleAdjDet where CreatedTransNo = "&iCrTransNo

		'sQuery = "Select isNull(ReceivableNumber,0) From  Acc_T_CreatedReceivables Where CreatedTransNo = "&iSalTrNo&" "&_
		'		 "And AmountReceivable > AmountReceived and AmountReceivable - AmountReceived >= "&dTotal&" "

    if Trim(iSalTrNo)<>"" then
		sQuery = "Select A.ReceivableNumber From Acc_T_Receivables A,Acc_T_CreatedReceivables C "&_
				 "Where A.CreatedReceivable = C.ReceivableNumber and  "&_
				 "C.CreatedTransNo = "&iSalTrNo&" and A.AmountReceivable > A.AmountReceived and  "&_
				 "A.AmountReceivable - A.AmountReceived >= "&dTotal&" "

		Response.Write sQuery &"<br><br>"

		objRs.open sQuery,con

		If Not objRs.EOF Then
			iPayableNo = objRs(0)
		Else
			iPayableNo = 0
		End If
		objRs.Close
	end if 'if Trim(iSalTrNo)<>"" then
	'
	''blocked by ragavendran on Oct 17,2011
'		if CDbl(iPayableNo) <> 0 THEN
'
'			'sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, ReceivedOn, "&_
'					 '"AmountReceived) VALUES ("&iPayableNo&", "&iTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&") "
'			'con.execute(sQuery)
'
'			sQuery = "insert into Acc_T_RcvblAdjustmentDetails (ReceivableNumber, RecdByTransactionNo, ReceivedOn, AmountReceived) "&_
'					 "values("&iPayableNo&","&iTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&")"
'
'
'			'Response.Write sQuery &"<br><br>"
'			' Date		:	18/11/2004
'			' Reason	:	Modified for Party settelment change Requested (To treate CR Note as Seprate Adj entry)
'			con.execute(sQuery)
'
'		else--ragav
	'----------------------------- CREATE PAYABLE IF RECEIVABLE DETAILS DOESNOT EXIST ------------------
			sQuery = "select isnull(max(PayablesNumber),0)+1 from Acc_T_Payables"
			objRs.open sQuery,con
				iPayableNo=objRs(0)
			objRs.Close

			Dim iCrRecNo
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
					
			'Response.Write sQuery
			con.execute(sQuery)
'		end if--ragav

	sQuery="update Acc_T_CreatedVoucherHeader set  CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iCrTransNo
	con.execute(sQuery)
	
	'============== Checking For Voucher Entry ===========================
	Dim dHdVal,dDetVal,dTaxDebVal,dTaxCreVal,dAccVal
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
	
	'Response.Write dAccVal &" === " & dHdVal &"<br>"
	
	IF CDbl(dAccVal) <> CDbl(dHdVal) Then
	    'Response.Clear 
	    Response.Write "<p><b>Debit Total = "& dAccVal 
	    Response.Write "<p><b>Credit Total = "& dHdVal
		Response.Write "<p><b>Debit and Credit Does'nt Match Inserion Rolled Back "
		Con.RollbackTrans
		Response.End
	End IF
	
	sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNarration&"' Where  "&_
			 "TransactionNumber = "&iTransNo
	Con.Execute sQuery
	
	sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNarration&"' Where  "&_
			 "TransactionNumber = "&iTransNo
	Con.Execute sQuery
	
	
	

'	Con.RollbackTrans
'	Response.End
	
	con.CommitTrans

	'******************************** Cost and ANal Procedure Call ************************

	'dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode,bAddFlag,sExp

	'sExp = "//Entry"
	'Set HeaderNode = oNodRoot.selectNodes(sExp)
	'IF HeaderNode.length <> 0 Then

		'for iCtr = 0 To HeaderNode.length - 1

			'sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value

			'sExp="//Entry[@No='"&sEntryno&"']/AccHead"
			'set tempNode=oNodRoot.selectNodes(sExp)
			'sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

			'set nodANL=oDOM.createElement("Root")
			'nodANL.setAttribute "TransNo",iTransNo
			'nodANL.setAttribute "EntryNo",sEntryno
			'nodANL.setAttribute "UnitCode", sorgid
			'nodANL.setAttribute "GlHead",sAccCode
			'nodANL.setAttribute "ACTFlag","V"

			'sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
			'set tempNode=oNodRoot.selectNodes(sExp)
			'if tempNode.length >0 then
				'set EntryNode=tempNode.item(0).cloneNode(true)
				'nodANL.appendChild(EntryNode)
				'bAddFlag=true
			'end if

			'sExp="//Entry[@No='"&sEntryno&"']/Analytical"
			'set tempNode=oNodRoot.selectNodes(sExp)
			'if tempNode.length >0 then
				'set EntryNode=tempNode.item(0).cloneNode(true)
				'nodANL.appendChild(EntryNode)
				'bAddFlag=true
			'end if
			'if bAddFlag then
			  'Dim adoConn

			  ' Set adoConn = Server.CreateObject("ADODB.Connection")
			   'adoConn.ConnectionString = con
			   'adoConn.CursorLocation = 3
			   'adoConn.Open
			   'sQuery = "Proc_VouCCANALUpdate"

			   'Dim adoCmd
			   'Set adoCmd = Server.CreateObject("ADODB.Command")
			   'Set adoCmd.ActiveConnection =adoConn
			   'adoCmd.CommandText = sQuery
			   'adoCmd.CommandType = 4 'adCmdStoredProc
			   'adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)

			   'Dim adoRS
			   'Set adoRS = adoCmd.Execute()
			'end if
		'next 'End of Entry Node Loop
	'End IF
	'******************************** Coast and ANal Procedure Call ************************


	if con.Errors.count <>0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) &"<br>"
		next
		'Redirect to Error Handling System
	else
		'con.CommitTrans
		'Con.RollbackTrans

		'oNodRoot.Attributes.Item(2).nodeValue= iTransNo
		'oNodRoot.Attributes.Item(3).nodeValue= iVouNo

		'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
		Response.Redirect ("CreditVouchers.asp?hFormVal="&sSelVouTy&"&voutype="&sSelVouTy)
	end if
Else
	Response.Redirect ("CreditVouchers.asp?hFormVal="&sSelVouTy&"&voutype="&sSelVouTy)
End IF

%>
