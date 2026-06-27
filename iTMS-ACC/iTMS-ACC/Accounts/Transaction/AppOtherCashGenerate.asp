<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppOtherCashGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	September 23, 2002
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
<%
dim oDOM,Root,sQuery,objRs,sSuppInvDate

dim iTransNo,iSeriesNo,iSeriesCode,sVouNo,sPara,sFormVal
dim sVouCode,sBookNo,sBookName,sOrgId,sVoucDate,sVouStatus,sMessage,sMessage1
dim iBkAccHead,iCrSeriesNo,iCrSeriesCode,sCrVouNo,iAccTrNo,dAccDate,iVouEntryNo

'sFormVal = Session("ReDirVal")
sFormVal = Session("ACTN")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")

sOrgId=Request("hOrgId")
sVouCode=Request("hBookCode")
sBookName=Request("hBookName")
sBookNo=Request("selBook")
iTransNo=Request("hTransNo")
dAccDate = Request("hAccDate")

sVouStatus="010103"
sPara = Request("hPara")
'Response.Write "111"
'Response.Write "AAAAAA="& getUserid()
'Response.Write "iTransNo="&iTransNo  &"***"&sPara
'Response.Write "<p>sFormVal="&sFormVal

'Response.End
'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
'set Root=oDOM.documentElement
'sVoucDate=Root.attributes.getNamedItem("VouDate").value

	sQuery = "Select Convert(Varchar,VoucherDate,103) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo
	'Response.Write sQuery &"<br>"
	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		sVoucDate = objRs(0)
	End IF
	objRs.Close

	sQuery="select isnull(BookAccountHead,0) from "&_
		"vwOrgBookNames where OUDefinitionID = '" & sOrgId & "' and BookNumber="&sBookNo&" and BookCode='"&sVouCode&"'"

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	while not objRs.EOF
		iBkAccHead=objRs(0)
		objRs.MoveNext
	wend
	objRs.Close

con.BeginTrans

If CStr(sPara) = "Acc" then 	''For Accounting
	sQuery = "Select R.SuppInvoiceNo+'--'+Convert(Varchar(10),R.SuppInvoiceDate,103) SuppInvDate  "&_
			 "From RCV_T_InvoiceHeader R,Acc_T_CreatedVoucherHeader A Where "&_
			 "R.InvoiceNumber = A.OtherApplnTransNo And A.CreatedTransNo = "&iTransNo
	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		sSuppInvDate = objRs("SuppInvDate")
	Else
		sSuppInvDate = ""
	End IF
	objRs.Close



	sQuery="select CreatedCrSeriesNo,CreatedCrSeriesCode,CrSeriesNo,CrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo

	objRs.open sQuery,con
		iCrSeriesNo=objRs(0)
		iCrSeriesCode=objRs(1)
		iSeriesNo=objRs(2)
		iSeriesCode=objRs(3)
	objRs.close()

	sCrVouNo=GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,dAccDate)
	sVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,dAccDate)

	if sVouCode="08" then
		sQuery="update Acc_T_CreatedVoucherHeader set BookNumber="&sBookNo&", "&_
		" CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iTransNo
	else
		sQuery="update Acc_T_CreatedVoucherHeader set AccountHead="&iBkAccHead&", BookNumber="&sBookNo&", "&_
		" CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iTransNo

	end if
	Response.Write sQuery&"<BR><BR>"
	con.execute(sQuery)


	sQuery = "Select Max(isNull(TransactionNumber,0))+1 From Acc_T_VoucherHeader "
	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		iAccTrNo = objRs(0)
	End IF
	objRs.Close


	sQuery = "Select OUDefinitionID,TransactionType,CreatedVoucherNo,PayToRecdFrom,Convert(Varchar,VoucherDate,103),VoucherAmount, "&_
			 "CrDrIndication,BankInstrumentType,PayableAt,CreatedBy From Acc_T_CreatedVoucherHeader "&_
			 "Where CreatedTransNo = "&iTransNo
	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		sQuery = "INSERT INTO Acc_T_VoucherHeader(TransactionNumber, OUDefinitionID, BookCode, BookNumber, "&_
				 "TransactionType, AccountHead, PartyType, PartySubType, PartyCode, VoucherNumber, "&_
				 "CreatedVoucherNo, CreatedTransNo, PayToRecdFrom, VoucherDate, VoucherAmount, CrDrIndication, "&_
				 "CreatedBy, CreatedOn, ApprovedBy, ApprovedOn, AccountedBy, AccountedOn, VoucherStatus,BRSTransactionNo) "&_
				 "VALUES ("&iAccTrNo&", '"&objRs(0)&"', '01', "&sBookNo&", '"&objRs(1)&"', "&iBkAccHead&", NULL, 0, 0, "&_
				 "'"&sVouNo&"', '"&objRs(2)&"', "&iTransNo&", '"&objRs(3)&"', Convert(Datetime,'"&dAccDate&"',103), "&objRs(5)&", "&_
				 "'"&objRs(6)&"', "&objRs("CreatedBy")&", Convert(Datetime,getDate(),103), "&objRs("CreatedBy")&", Convert(Datetime,getDate(),103), "&objRs("CreatedBy")&", Convert(Datetime,getDate(),103), '010104',NULL)"


		Response.Write sQuery &"<br><br>"
		con.execute sQuery
	End IF
	objRs.Close

	sQuery = "Update Acc_T_CreatedVoucherHeader Set VoucherDate = Convert(Datetime,'"&dAccDate&"',103) Where CreatedTransNo = "&iTransNo
	con.execute(sQuery)

	sQuery = "Select VoucherEntryNumber, AccountingUnit, AccUnitAccountHead, "&_
			 "VoucherNarration, Amount,TransCrDrIndication, "&_
			 "ItemCode,ClassificationCode,ItemDescription,InvoicedQuantity,InvoicedUoM, " & _
			 "InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount " & _
			 "FROM Acc_T_CreatedVoucherDetails Where CreatedTransNo = "&iTransNo

	Response.Write sQuery  &"<br><br>"
	objRs.Open sQuery,con
	Do While Not objRs.EOF
		iVouEntryNo = Cdbl(iVouEntryNo) + 1
		sQuery = "INSERT INTO Acc_T_VoucherDetails(TransactionNumber, VoucherEntryNumber, AccountingUnit, "&_
				 "AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, VoucherNarration, Amount,  "&_
				 "TransCrDrIndication,AccUnitAccountHead,ItemCode,ClassificationCode,ItemDescription, "&_
				 "InvoicedQuantity,InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount) "&_
				 "VALUES ("&iAccTrNo&", "&objRs("VoucherEntryNumber")&", '"&objRs("AccountingUnit")&"', NULL, 0, 0,  "&_
				 "'"&sSuppInvDate&"', "&objRs("Amount")&", '"&objRs("TransCrDrIndication")&"',"&objRs("AccUnitAccountHead")&", "&_
				 " "&objRs("ItemCode")&", "&objRs("ClassificationCode")&", '"&Replace(objRs("ItemDescription"),"'","''")&"', "&_
				 " "&objRs("InvoicedQuantity")&",'"&objRs("InvoicedUoM")&"',"&objRs("InvoicedRate")&", "&_
				 " "&objRs("BasicAmount")&","&objRs("DiscountPercent")&","&objRs("DiscountAmount")&") "

		Response.Write sQuery &"<br><br>"
		con.execute sQuery
		objRs.MoveNext
	Loop
	objRs.Close

	sQuery = "Select AccountHead,TaxAMount,TransCrDrIndication From Acc_T_CreatedVoucherTaxDet Where  "&_
			 "CreatedTransNo = "&iTransNo&" and TaxAmount <> 0 "
	objRs.Open sQuery,con
	Do While Not objRs.EOF
		iVouEntryNo = Cdbl(iVouEntryNo) + 1
		sQuery = "INSERT INTO Acc_T_VoucherDetails(TransactionNumber, VoucherEntryNumber, AccountingUnit, "&_
				 "AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, VoucherNarration, Amount,  "&_
				 "TransCrDrIndication,AccUnitAccountHead) "&_
				 "VALUES ("&iAccTrNo&", "&iVouEntryNo&", '"&sOrgID&"', NULL, 0, 0,  "&_
				 "'"&sSuppInvDate&"', "&objRs("TaxAMount")&", '"&objRs("TransCrDrIndication")&"',"&objRs("AccountHead")&") "

		Response.Write sQuery &"<br><br>"
		con.execute sQuery
		objRs.MoveNext
	Loop
	objRs.Close



	sQuery = "Update Acc_T_CreatedVoucherHeader Set  CreatedVouchStatus = '010104' Where CreatedTransNo = "&iTransNo
	Response.Write sQuery&"<BR><BR>"
	con.execute sQuery

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
		Response.Clear
		con.RollbackTrans

		'sErrChk = "F"
		Response.Clear
		Response.Write "<h3>Debit and Credit is Not Matching </h3><br>"
		Response.Write "<b>Debit Amount  :--> </b>" & dTBDRAmt &"<br>"
		Response.Write "<b>Credit Amount :--> </b>" & dTBCRAmt &"<br>"
		Response.Write "<b>Differences   :--> </b>" & sDiffVal &"<br>"
		Response.Write "<h3> Voucher Not Created " &"<br>"
		Response.End

	End IF


ElseIf CStr(sPara) = "App" then  'For Approval Case

	sQuery="update Acc_T_CreatedVoucherHeader set AccountHead="&iBkAccHead&", BookNumber="&sBookNo&",CreatedVouchStatus='"&sVouStatus&"',"&_
		   "ApprovedBy = '"& getUserid() &"',ApprovedOn = Convert(Datetime,'"&dAccDate&"',103),VoucherDate = Convert(Datetime,'"&dAccDate&"',103) "&_
		   "where CreatedTransNo="&iTransNo &" "
		   '
	Response.Write sQuery&"<BR><BR>"
	con.execute(sQuery)

ElseIf CStr(sPara) = "Edt" then  'For Edit Case

	sQuery="update Acc_T_CreatedVoucherHeader set AccountHead="&iBkAccHead&", BookNumber="&sBookNo&","&_
			"  VoucherDate = Convert(Datetime,'"&dAccDate&"',103) where CreatedTransNo="&iTransNo
	Response.Write sQuery&"<BR><BR>"
	con.execute(sQuery)

End IF
'Con.RollbackTrans
  ' 	Response.End

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	con.CommitTrans
	Response.Clear
	If CStr(sPara) = "Acc" then
		sMessage="Created Voucher No : "&sCrVouNo &", Accounted Voucher No : " & sVouNo
		sMessage1="Voucher Accounted Sucessfully"

	End IF
end if

%>
<HTML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script language="javascript">
function Message(strr, strr1, sFormVal) {
	alert(String(strr || "") + "\n" + String(strr1 || ""));
	window.location.href = "CashVouchers.asp?ACTN=" + encodeURIComponent(sFormVal || "");
}
function Redire(strr, sFormVal) {
	window.location.href = "CashVouchers.asp?ACTN=" + encodeURIComponent(sFormVal || "");
}
</script>
<%If CStr(sPara) = "Acc" then %>

	<BODY onLoad = "Message('<%=sMessage%>','<%=sMessage1%>','<%=sFormVal%>')"/>
<%Else %>
	<BODY onLoad = "Redire('<%=sMessage%>','<%=sFormVal%>')"/>
<% End IF%>
</html>

