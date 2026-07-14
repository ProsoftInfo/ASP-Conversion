<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BrsGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 31, 2003
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
dim oDOM,Root,nodHeader,objRs,sQuery,iBrsNo,sReason,sOrgId
dim iBookNo,dPassBalance,sPassCrDr,dBookBalance,sBookCrDr,sFromDt,sToDt
dim iTransNo,dAmont,sTransType,sClearedOn,sRecVal
dim dPayTotal,dRecTotal,dTotal,sExp,tempNode,iCounter,sPayTo,sEntryNo,sNarration
dim dtVouDate,iAccHeadNo,iAmount,sVouType,sChkVal,iBookAccHead,sVouCode,sVouStatus
Dim iCreatedVouNo,iVouNo,iCreatedTransNo,iCrSeriesNo,iCrSeriesCode,iSeriesNo,iSeriesCode
Dim iBkInsNo,dBkInsDate,sPayAt,sDrwOnBank,sBkInsType,iInsNo
sReason=Request("txtReason")
sChkVal = Request("hChkVal")
dtVouDate = Request("hVouDate")
iAccHeadNo = Request("hAccNo")
iAmount = Request("hAmt")
sVouType = Request("hCrDr")
sNarration = Request("hNarr")
sRecVal = Request("hRecVal")
iTransNo = Request("hTransNo")
 Response.Write "iTransNo ="&iTransNo
'Response.Write "Date="&dtVouDate&"<BR>iAccHeadNo="&iAccHeadNo&"<BR>iAmount="&iAmount &"<BR>sVouType="&sVouType&"<BR>"
Response.Write "sNarration="&sNarration &"<BR>"
set objRs  = server.CreateObject("adodb.recordset")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml")

set Root=oDOM.documentElement
sOrgId = root.Attributes.Item(0).nodeValue
iBookNo= root.Attributes.Item(1).nodeValue
sFromDt= root.Attributes.Item(3).nodeValue
sToDt= root.Attributes.Item(4).nodeValue
dPassBalance= root.Attributes.Item(5).nodeValue
sPassCrDr= root.Attributes.Item(6).nodeValue
dBookBalance= root.Attributes.Item(7).nodeValue
sBookCrDr= root.Attributes.Item(8).nodeValue
dPayTotal=0
dRecTotal=0

IF Cstr(sPassCrDr) = "DR" Then
	sPassCrDr = "D"
Else
	sPassCrDr = "C"
End IF

sExp="//Voucher"
set tempNode=Root.selectNodes(sExp)

for iCounter=0 to tempNode.length-1
	dAmont=tempNode.item(iCounter).Attributes.Item(6).nodeValue
	sTransType=tempNode.item(iCounter).Attributes.Item(7).nodeValue
	if tempNode.item(iCounter).Attributes.Item(2).nodeValue="N" then
		if sTransType="P" then
			dPayTotal=CDbl(dPayTotal)+ CDbl(dAmont)
		else
			dRecTotal=CDbl(dRecTotal)+ CDbl(dAmont)
		end if

	end if
next
con.BeginTrans
'**************************************************************************************
IF trim(sChkVal) = "True" then
	sQuery="select AccountDescription from Acc_M_GLAccountHead where AccountHead = 	"&iAccHeadNo
	objRs.open sQuery,con
	IF Not objRs.EOF then
		sPayTo = objRs(0)
	End IF
	objRs.close
	sQuery="select BookAccountHead from Acc_R_ApplicableAccountHeads where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='02' and BookNumber= "&iBookNo

	objRs.open sQuery,con
		iBookAccHead=objRs(0)
	objRs.close()

	sQuery="select CreatedDrSeriesNo,CreatedDrSeriesCode,DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='02' and BookNumber= "&iBookNo

	objRs.open sQuery,con
		iCrSeriesNo=objRs(0)
		iCrSeriesCode=objRs(1)
		iSeriesNo=objRs(0)
		iSeriesCode=objRs(1)
	objRs.close()

	iCreatedVouNo=GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,dtVouDate)
	iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,dtVouDate)

	sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
	objRs.open sQuery,con
		iCreatedTransNo=objRs(0)
	objRs.Close
	sVouCode="02"
	sVouStatus="010101"
	sEntryno = 1
	'Response.Write "sVouType="&sVouType
	if sVouType="D" then
		sTransType="BAP"
	else
		sTransType="BAP"
	end if
	sQuery = "Select isNull(BankInstrumentNo,'NULL'),convert(VarChar,BankInstrumentDate,103),isNull(PayableAt,'NULL'),isNull(DrawnOnBank,'NULL'),BankInstrumentType from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iTransNo
	objRs.Open sQuery,con

	IF Not objRs.EOF then
		iBkInsNo	= objRs(0)
		dBkInsDate  = objRs(1)
		sPayAt		= objRs(2)
		sDrwOnBank  = objRs(3)
		sBkInsType  = objRs(4)
	End If
	objRs.Close
	'Acc_T_CreatedVoucherHeader

	IF dBkInsDate = "" then dBkInsDate = "NULL"
'	sQuery="insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
'			"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
'			"PayToRecdFrom,CrDrIndication,BankInstrumentType,BankInstrumentDate,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"&_
'			"("&iCreatedTransNo&",'"&sOrgId&"','"&sVouCode&"',"&iBookNo&",'"&sTransType&"',"&_
'			"NULL,NULL,NULL,"&iBookAccHead&",'"&iCreatedVouNo&"',convert(datetime,'"&dtVouDate &"',103),"&iAmount &_
'			",'"&sPayTo&"','"&sVouType&"','Bank Charges',convert(datetime,'"&dtVouDate&"',103),"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

	sQuery="insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			"PayToRecdFrom,CrDrIndication,BankInstrumentType,BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"&_
			"("&iCreatedTransNo&",'"&sOrgId&"','"&sVouCode&"',"&iBookNo&",'"&sTransType&"',"&_
			"NULL,NULL,NULL,"&iBookAccHead&",'"&iCreatedVouNo&"',convert(datetime,'"&dtVouDate &"',103),"&iAmount &_
			",'"&sPayTo&"','"&sVouType&"','"&sBkInsType&"',"&iBkInsNo&",convert(datetime,'"&dBkInsDate&"',103),'"&sPayAt&"','"&sDrwOnBank&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)

	'Acc_T_CreatedVoucherDetails
		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication) values ("
		sQuery=sQuery& iCreatedTransNo&",'"&sOrgId&"'"
		sQuery=sQuery& ","&sEntryno&","&iAccHeadNo&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&iAmount&",'"&sVouType&"')"
	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)

End If
'******************************End********************************************************
IF sRecVal <> "BR" then 'Reconciled
	sQuery="select isnull(max(BRSTransactionNo),0)+1 from Acc_T_BankReconciliation"
	objRs.open sQuery,con
		iBrsNo=objRs(0)
	objRs.Close


	sQuery="INSERT INTO Acc_T_BankReconciliation(BRSTransactionNo, BRSDoneOn, BRSPeriodFrom, BRSPeriodTo,"&_
			" BankBalance, ApplicationBalance, DepositedNotCleared,IssuedNotPresented,ReasonForDifference,BookNo,CDIndication) VALUES("&_
			""&iBrsNo&",getdate(),convert(datetime,'"&sFromDt&"',103),convert(datetime,'"&sToDt&"',103),"&_
			""&dPassBalance&","&dBookBalance&","&dRecTotal&","&dPayTotal&",'"&sReason&"',"&iBookNo&",'"&sPassCrDr&"' )"
	con.execute(sQuery)
	Response.Write sQuery&"<BR><BR>"

	for iCounter=0 to tempNode.length-1
		if tempNode.item(iCounter).Attributes.Item(2).nodeValue="Y" then
			iTransNo=tempNode.item(iCounter).Attributes.Item(1).nodeValue
			dAmont=tempNode.item(iCounter).Attributes.Item(6).nodeValue
			sClearedOn=tempNode.item(iCounter).Attributes.Item(8).nodeValue
			iInsNo=tempNode.item(iCounter).Attributes.Item(9).nodeValue
			'sQuery="Update Acc_T_VoucherHeader set BRSTransactionNo="&iBrsNo&", ClearedOn=convert(datetime,'"&sClearedOn&"',103)"&_
			'" where TransactionNumber="&iTransNo

			sQuery="Update Acc_T_CreatedVoucherHeader set BRSTransactionNo="&iBrsNo&", ClearedOn=convert(datetime,'"&sClearedOn&"',103)"&_
			" where CreatedTransNo="&iTransNo

			Response.Write sQuery&"<BR><BR>"
			con.execute(sQuery)


			sQuery="Update Acc_T_CreatedVoucherInstrumentDet set BRSTransactionNo="&iBrsNo&", ClearedOn=convert(datetime,'"&sClearedOn&"',103)"&_
			" where CreatedTransNo="&iTransNo&" and BankInstrumentNo = "&iInsNo
			Response.Write sQuery&"<BR><BR>"
			con.execute(sQuery)
		end if
	next
End IF
	'con.RollbackTrans
 'Response.End
if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	con.CommitTrans
	Response.Redirect ("BANKRECONCILIATION.asp")
end if
%>

