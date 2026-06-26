<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MiscInvoiceUpdate.asp
	'Module Name				:	Common 
	'Author Name				:	Ragavendran R
	'Created On					:	July 17,2013
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
<!--#include File="../../include/purpopulate.asp" -->
<!--#include file="../../include/Accpopulate.asp"-->
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
Dim sChequeDate,sChequeNo

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

iTransNo = Request.QueryString("InvNo")

oDOM.Load server.MapPath("../temp/Transaction/MISCPaymentEdit"&Session.SessionID&".xml")

Set Root = oDOM.documentElement
sPayThr = Root.Attributes.Item(17).nodeValue
IF CStr(sPayThr) = "" Then
	sPayThr = "C" 
End IF

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

sQuery = "Update Acc_T_MiscPymtRequestHeader set OUDefinitionID='"& sOrgId &"', TransactionType='"&sVouTy&"',VoucherDate=Convert(Datetime,'"&sCurrDate&"',103), VoucherAmount="&sAmount&", CrDrIndication='"&sCrDrTrans&"', "&_
		 "CreatedBy="&sCrtBy&", CreatedOn= getDate(), CreatedVouchStatus='"&sVouStatus&"', PayToRecdFrom='"&sMiscPartyName&"', CreatedMiscPymtNo='"&sRefNo&"',PaymentFor='"& sPassPayFor&"',ReferenceNo=" & sPassRefNo & ","&_
		 " BankInstrumentType='"&sPayThr&"',AppRefNo='"& sAppRefNo &"',AppRefDate='"& sAppRefDate &"',AppRefType="&sAppRefType&",ApplicationCode="& iApp &",MiscPartyCode="&sMiscPartyCode&",AdjustAgainstInvoice='"& sAdjustAgainstInvoice&"',"&_
		 " ChequeNo = "&sChequeNo&",ChequeDate="& sChequeDate &""&_
		 " where MiscTransNo = "&iTransNo
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
sQuery = "Delete from Acc_T_MiscPaymentReqDetails where MiscTransNo="& iTransNo
con.execute sQuery

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
if objfs.FileExists(Server.MapPath("../temp/Transaction/MISCPaymentEdit"&Session.SessionID&".xml")) then 
    objfs.DeleteFile(Server.MapPath("../temp/Transaction/MISCPaymentEdit"&Session.SessionID&".xml"))
end if
Response.Redirect "MISCINVOICES.ASP?APPCODE="& iApp

%>