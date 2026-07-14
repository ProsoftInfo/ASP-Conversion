<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MiscInvoiceInsert.asp
	'Module Name				:	Purchase 
	'Author Name				:	Ragavendran R
	'Created On					:	April 08,2011
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<%
	Dim oDOM,nodHeader,Root,objRs,sQuery
dim sNarration,sAccount,sAddtional,iSno
dim dTotal,sOrgId,iCtr,iEntryno
dim EntryNode,HeaderNode,dAmount,sVouStatus,sCrtBy,sTotAmount,objfs
dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim iTransNo,iBkHeadCode,bOtherUnit,sExp,Tempnode,sAmount,sParType,sParCode,sParSubType
Dim sRefNo,sVouTy,sPayThr,sCheckPayToSelection,sMiscPartyName
Dim sPassPayTo,sPassPayFor,sPassRefNo,sCrDrTrans,iApp,sPayFor,sPayForName,sCurrDate
Dim sCurrMon,sCurrDay,sAppRefNo,sAppRefDate,sAppRefType,sMiscPartyCode,sAdjustAgainstInvoice
Dim sChequeNo,sChequeDate

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objfs = CreateObject("Scripting.FileSystemObject")

sCurrMon = Month(Date)
sCurrDay = Day(Date)

IF Len(sCurrMon) = 1 then
	sCurrMon = "0" & sCurrMon
End IF

IF Len(sCurrDay) = 1 then
	sCurrDay = "0" & sCurrDay
End IF

sCurrDate = sCurrDay&"/"&sCurrMon&"/" & Year(Date)



sVouStatus = "010101" 'Created for Approval
sCrtBy = getUserid()
iApp = Request("hAppCode")

oDOM.Load server.MapPath("../temp/Transaction/MISCPayment"&Session.SessionID&".xml")

Set Root = oDOM.documentElement
sPayThr = Root.Attributes.Item(17).nodeValue
IF CStr(sPayThr) = "" Then
	sPayThr = "C" 
End IF

sQuery = "Select isNull(Max(MiscTransNo),0) + 1 From Acc_T_MiscPymtRequestHeader "
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sQuery
	.Open
End With
IF Not objRs.EOF Then
	iTransNo = objRs(0)
End IF
objRs.Close

Set Root=oDOM.documentElement
sExp = "//voucher"
Set Tempnode = Root.selectNodes(sExp)
IF Tempnode.length <> 0 Then
	sOrgId = Tempnode.Item(0).Attributes.getNamedItem("UnitNo").value
	sParType = Tempnode.Item(0).Attributes.getNamedItem("PartyType").value
	sParSubType = Tempnode.Item(0).Attributes.getNamedItem("PartySubType").value
	sParCode = Tempnode.Item(0).Attributes.getNamedItem("PartyCode").value
	sRefNo = Tempnode.Item(0).Attributes.getNamedItem("ReferenceNo").value
	
	sPassPayTo  = Tempnode.Item(0).Attributes.getNamedItem("hPayTo").value
	sPassPayFor = Tempnode.Item(0).Attributes.getNamedItem("hPayFor").value
	sPassRefNo  = Tempnode.Item(0).Attributes.getNamedItem("hRefNo").value
	
	sPayFor = Tempnode.Item(0).Attributes.getNamedItem("PayFor").value
	sPayForName = Tempnode.Item(0).Attributes.getNamedItem("PayForName").value
	sAppRefNo = Tempnode.Item(0).Attributes.getNamedItem("AppRefNo").value
	sAppRefType = Tempnode.Item(0).Attributes.getNamedItem("AppRefType").value
	sAppRefDate = Tempnode.Item(0).Attributes.getNamedItem("AppRefDate").value
	sParCode = Tempnode.Item(0).Attributes.getNamedItem("Code").value
	sChequeNo = Tempnode.Item(0).Attributes.getNamedItem("CheNo").value
	sChequeDate = Tempnode.Item(0).Attributes.getNamedItem("CheDate").value
	
End IF
if trim(sAppRefType) = "" or trim(sAppRefType)="N" then sAppRefType = "NULL"
If Trim(sPassRefNo) = "" Then sPassRefNo = 0

sExp = "//Narration"
Set Tempnode = Root.selectNodes(sExp)
IF Tempnode.length <> 0 Then
	sNarration = Tempnode.Item(0).Text
End IF

sExp = "//voucher"
Set Tempnode = Root.selectNodes(sExp)

For Each HeaderNode in Tempnode.item(0).childNodes
	IF CStr(HeaderNode.nodeName) = "Entry" Then
		sPayTo = HeaderNode.Attributes.Item(2).nodeValue
		sAmount = HeaderNode.Attributes.Item(3).nodeValue
		sCrDrTrans = HeaderNode.getAttribute("CRDR")
		sAmount = CDbl(sAmount)
		sTotAmount = CDbl(sTotAmount) + sAmount
		sAdjustAgainstInvoice = HeaderNode.getAttribute("CheckVal")
		sCheckPayToSelection = HeaderNode.getAttribute("PayToSelCheck")
		sMiscPartyName = HeaderNode.getAttribute("MiscPartyName")
		sMiscPartyCode = HeaderNode.getAttribute("MiscPartyCode")
	End IF
Next


IF CStr(sPayThr) = "C" Then 'Cash Voucher
	IF CStr(sCrDrTrans) = "C" Then
		sVouTy = "CAP"
	Else
		sVouTy = "CAR"
	End IF
Else 'Bank Voucher
	IF CStr(sCrDrTrans) = "C" Then
		sVouTy = "BAP"
	Else
		sVouTy = "BAR"
	End IF
End IF
if Trim(sChequeNo)="" or IsNull(sChequeNo) then sChequeNo ="NULL"
if Trim(sChequeNo)<>"NULL" then sChequeNo=pack(sChequeNo)

if Trim(sChequeDate)="" or IsNull(sChequeDate) then sChequeDate ="NULL"
if Trim(sChequeDate)<>"NULL" then sChequeDate=pack(sChequeDate)

con.BeginTrans

sQuery = "INSERT INTO Acc_T_MiscPymtRequestHeader (MiscTransNo, OUDefinitionID, TransactionType,  "&_
		 "VoucherDate, VoucherAmount, CrDrIndication, "&_
		 "CreatedBy, CreatedOn, CreatedVouchStatus, PayToRecdFrom, CreatedMiscPymtNo,PaymentFor,ReferenceNo,"&_
		 " BankInstrumentType,AppRefNo,AppRefDate,AppRefType,ApplicationCode,MiscPartyCode,AdjustAgainstInvoice,ChequeNo,ChequeDate) "&_
		 "VALUES ("&iTransNo&", '"&sOrgId&"', '"&sVouTy&"',  "&_
		 "Convert(Datetime,'"&sCurrDate&"',103), "&sAmount&",  "&_
		 "'"&sCrDrTrans&"', "&sCrtBy&", getDate(), '"&sVouStatus&"', '"&sMiscPartyName&"', '"&sRefNo&"','"& sPassPayFor&"',"&_
		 "" & sPassRefNo & ", '"&sPayThr&"','"& sAppRefNo &"','"& sAppRefDate &"',"&sAppRefType&","& iApp &","&sMiscPartyCode&",'"& sAdjustAgainstInvoice&"',"& sChequeNo &","& sChequeDate &") "
		Response.Write "<p>"&sQuery
		Con.Execute sQuery
				 
IF CStr(sParType) <> "0" Then
	sQuery = "UPDATE Acc_T_MiscPymtRequestHeader SET PartyType = '"&sParType&"',  "&_
			 "PartySubType = "&sParSubType&" , PartyCode = "&sParCode&" WHERE MiscTransNo = "&iTransNo&" "
			 Response.Write "<p>"&sQuery
	Con.Execute sQuery
End IF

IF CStr(sPayFor) = "O" Then
	sQuery = "UPDATE Acc_T_MiscPymtRequestHeader SET OtherPaymentFor = '"&sPayForName&"' WHERE MiscTransNo = "&iTransNo&" "
	con.execute sQuery
	Response.Write "<p>"&sQuery
End IF

		
IF CStr(sCrDrTrans) = "D" Then
	sCrDrTrans = "C"
Else
	sCrDrTrans = "D"
End IF

For Each HeaderNode in Tempnode.item(0).childNodes
	IF CStr(HeaderNode.nodeName) = "Entry" Then
		sAmount = HeaderNode.Attributes.Item(3).nodeValue
		iEntryno = HeaderNode.Attributes.Item(0).nodeValue
		
		sQuery = "INSERT INTO Acc_T_MiscPaymentReqDetails (MiscTransNo, AccountingUnit, VoucherNarration, "&_
				 "Amount, TransCrDrIndication, VoucherEntryNumber) VALUES "&_
				 "("&iTransNo&", '"&sOrgId&"', '"&PackQuote(sNarration)&"', "&sAmount&", '"&sCrDrTrans&"', "&iEntryno&") "
		Response.Write "<p>"&sQuery
		Con.Execute sQuery
		
	End IF
Next

'Con.RollBackTrans
'Response.End 
Response.clear
Con.commitTrans
if objfs.FileExists(Server.MapPath("../temp/Transaction/MISCPayment"&Session.SessionID&".xml")) then 
    objfs.DeleteFile(Server.MapPath("../temp/Transaction/MISCPayment"&Session.SessionID&".xml"))
end if
Response.Redirect "MISCINVOICES.ASP?APPCODE="& iApp

%>