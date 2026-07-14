<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNOthInvGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	Nov 05, 2004
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
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,iCounter
dim iSeriesNo,iSeriesCode
dim sSalType, ActualTransNo
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
dim dAdjAmtTotal,sAppTy,sAppBy,iCrTransNo,objfs,sCallTy,sFromSal
Dim sExp,NarrNode,sVouNarr,iInvNo,sCrtVouType
Dim TaxCalNode,dTaxValNoAcc,iCtr,dEachItmVal,dEachItm,dNoAccTotVal,dDiffVal,sDispType
dim sChkVal,sPayAt,sDrwOn,iBookNo,iCRDRNoteEntryNo,iSalInvoiceNo
Dim rsTemp,iItemCode,iClassCode


iBookNo = Request("hBookcode")
sChkVal = Request("hChkVal")

sAppTy = Request.Form("optApprove")
sCallTy = Request.Form("hCallType")
sFromSal = Request.Form("hFromSal")
sDispType = Request.Form("SelCrAgain")
sCrtVouType =Request("hCrtVouType")
'Response.Write sCrtVouType

IF CStr(sAppTy) = "Y" Then
	sVouStatus="010101" 'Waiting For Approval
Else
	sVouStatus="010105" 'Waiting For Accounting
End IF

sVouCode="07"
sVouType="C"

ActualTransNo = Request("hdTransNo")
sCallTy = Request.Form("hCallType")

set objRs  = server.CreateObject("adodb.recordset")
set rsTemp  = server.CreateObject("adodb.recordset")
Set objfs = CreateObject("Scripting.FileSystemObject")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.load server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")
set oNodRoot=oDOM.documentElement

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
				iInvNo = oNodEntry.Attributes.Item(5).nodeValue
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
	if oNodTemp.nodeName="SalesInvoiceEntry" then
	    iSalInvoiceNo = oNodTemp.getAttribute("InvoiceNo")
	    Response.Write "<p> Invoice No = "& iSalInvoiceNo
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

 '   IF CDbl(dEachItm) = 1 Then
'	    dEachItmVal = dTaxValNoAcc
 '   Else
'	    dEachItmVal =  Round(CDbl(dTaxValNoAcc) / CDbl(dEachItm),0)
'	    dNoAccTotVal = CDbl(dEachItmVal) * CDbl(dEachItm)
'	    dDiffVal = CDbl(dTaxValNoAcc) - CDbl(dNoAccTotVal)
'
 '   End IF


sTransType="CNR"
iInvNo = Request("hdTransNo")

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

dim sAccHeadCode,iTemp

sQuery="select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo

objRs.open sQuery,con
if not objRs.EOF then
	iSeriesNo=objRs(0)
	iSeriesCode=objRs(1)
end if
objRs.close()

con.BeginTrans

iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iCrTransNo=objRs(0)
objRs.Close

If trim(sChkVal) = "Y" then 
	sCrtVouType = "G"
	sDrwOn = iBookNo
Else 
	sDrwOn = "NULL"
End If
if Trim(sDrwOn)<>"NULL" then sDrwOn = Pack(sDrwOn)
sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,PurchaseBillType,BankInstrumentNo,PayableAt,DrawnOnBank) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
sQuery=sQuery&",'Sales Return ','OS','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"','"&sDispType&"',"&iInvNo&",'"&sCrtVouType&"',"&sDrwOn&")"

Response.Write sQuery& "<BR>"

con.execute(sQuery)

sQuery = "Select isNull(Max(CRDRNoteEntryNo),0)+1 from ACC_T_CreatedVoucherCRDRNotes"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    iCRDRNoteEntryNo = rsTemp(0)    
end if
rsTemp.Close 

sQuery = "Insert into ACC_T_CreatedVoucherCRDRNotes (CRDRNoteEntryNo,CROrDRNote,CreatedTransNo,RefType,RefNumber)"&_
         " values("& iCRDRNoteEntryNo &",'CNR',"& iCrTransNo &",'OS',"& iSalInvoiceNo &")"
Response.Write "<p>"& sQuery
con.execute sQuery

for each EntryNode in oNodDeatils.childNodes
dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc

	sEntryno=EntryNode.Attributes.Item(0).nodeValue
	sAmount=EntryNode.Attributes.Item(2).nodeValue
	
	sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
	dQty=EntryNode.Attributes.Item(3).nodeValue
	sUOM=EntryNode.Attributes.Item(4).nodeValue
	dBasicAmount="0"
	dRate=EntryNode.Attributes.Item(6).nodeValue
	dDisPer=EntryNode.Attributes.Item(8).nodeValue
	dDisAmount=EntryNode.Attributes.Item(9).nodeValue
iItemCode = EntryNode.getAttribute("ItemCode")
iClassCode = EntryNode.getAttribute("ClassCode")
	sEntryType="D"
	sNarration=""
'if	CDbl(sAmount) > 0 then
	for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue
				sAccType=HeaderNode.Attributes.Item(4).nodeValue
		end if 'End of Check for Account head Node
	next 'End of Entry Node Loop
Response.Write "<p>"& sAmount 

'sAmount = CDbl(sAmount) + CDbl(dEachItmVal)

'	IF CDbl(iCtr) = 1 Then
'		sAmount = CDbl(sAmount) + CDbl(dDiffVal)
'		iCtr = 2
'	End IF
Response.Write "<BR>"&dDisAmount &"<BR><BR>"
if Trim(dDisAmount)="" or IsNull(dDisAmount) then dDisAmount = "0"

	sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
	sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
	sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
	sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemCode,ClassificationCode) values ("
	sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
	sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
	sQuery=sQuery&" '"&sVouNarr&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
	sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","& iItemCode &","& iClassCode &")"
Response.Write sQuery
	con.execute(sQuery)
'End IF

next'End of Voucher Node Loop

'-----------------------------TAX TABLE UPDATE------------------
iSno=1
dim sTaxMode
for each EntryNode in oNodTaxRoot.childNodes
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

		'if	CDbl(dTaxAmount) > 0 then
			if sTaxMode="P" then
				sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
				con.execute(sQuery)
			else
				sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
				con.execute(sQuery)
			end if
		'end if
		Response.Write "<p>"&sQuery


		iSno=cint(iSno)+1
	End IF
Next

'----TO CODE FOR TAX FORM DEATILS----------------

'----------------------------- RECEIVABLE TABLE UPDATE------------------
	'sQuery="select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "&ActualTransNo
	sQuery = "select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "&ActualTransNo&" "&_
			 "And AmountReceivable > AmountReceived and AmountReceivable - AmountReceived >= "&dTotal&" "


	Response.Write sQuery
	objRs.open sQuery,con

	If Not objRs.EOF Then
		iPayableNo = objRs(0)
	Else
		iPayableNo = 0
	End If
	objRs.Close
'
''blocked by ragav on Oct 17,2011
'	if CDbl(iPayableNo) <> 0 THEN
'
'		sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, ReceivedOn, "&_
'				 "AmountReceived) VALUES ("&iPayableNo&", "&iCrTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&") "
'		con.execute(sQuery)
'
'	else--ragav
'----------------------------- CREATE PAYABLE IF RECEIVABLE DETAILS DOESNOT EXIST ------------------
		sQuery = "select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close
		IF trim(sCrtVouType) = "G" then 

			sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber,CreatedTransNo, OUDefinitionID, "&_
					"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
					"PartyBillDate, AmountPayable, AmountPaid) values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
					"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
					"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&","&dTotal&")"
		Else
			sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber,CreatedTransNo, OUDefinitionID, "&_
					"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
					"PartyBillDate, AmountPayable, AmountPaid) values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
					"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
					"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
		End IF
		con.execute(sQuery)
'	end if--ragav

'////////////////////////////////////////////////////////////////////////////////////////////

IF Cstr(sAppBy) <> "I" Then
	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
		 	"Values("&iCrTransNo&",1,"&sAppBy&" ) "

	'Con.Execute sQuery
End IF
'/////////////////////////////////////// Approval Entry /////////////////////////////////////

'con.rollbacktrans
'Response.End 
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
		'nodANL.setAttribute "TransNo",iCrTransNo
		'nodANL.setAttribute "EntryNo",sEntryno
		'nodANL.setAttribute "UnitCode", sorgid
		'nodANL.setAttribute "GlHead",sAccCode
		'nodANL.setAttribute "ACTFlag","C"

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

		  '	Set adoConn = Server.CreateObject("ADODB.Connection")
		  ' adoConn.ConnectionString = con
		  ' adoConn.CursorLocation = 3
		  ' adoConn.Open

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
Response.Clear 
'Response.End 

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	'con.CommitTrans

	oNodRoot.Attributes.Item(0).nodeValue= iCrTransNo
	oNodRoot.Attributes.Item(1).nodeValue= iVouNo

	'IF the Selected Sales Voucher is From the Sales Invoice then Update the Changes in Sales
	'Related tabel also.
	IF CStr(sFromSal) = "Y" Then

		Dim adoConn
		Set adoConn = Server.CreateObject("ADODB.Connection")
		adoConn.ConnectionString = con
		adoConn.CursorLocation = 3
		adoConn.Open

		sQuery = "Acc_CrDrForOthInv"

	'	oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

		Dim adoCmd
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection =adoConn
		adoCmd.CommandText = sQuery
		adoCmd.CommandType = 4 'adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(oDOM.xml),oDOM.xml)

		Dim adoRS
		'Set adoRS = adoCmd.Execute()
	End IF


	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

	if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")) then
		objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"))
	End IF

	IF CStr(sCallTy) = "OINV" Then
		Response.Redirect ("VouCNOtherInvDisplay.asp?TransNo="&iCrTransNo&"&CrtType="&sCrtVouType)
	Else
		Response.Redirect ("VouCNSalReturnDisplay.asp?TransNo="&iCrTransNo&"&CrtType="&sCrtVouType)
	End IF


end if
%>
