
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ManageVouchers_MoveUpdate_forBANK.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	Septemper 26,2011
	'Modified By                :   
	'Modified On                :   
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
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->


<%
dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,sQuery
dim iVouNo,objRs,objRs1,sVouDate,bActionFlag
dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sAccHead,sParType,sPartSubType
dim iEnNo,Entrynode,HeaderNode,dOpeningBal
dim sParCode,sNarration,sAccHeadname,sAccUnitName,bOtherUnits,iBookAccHead,dTransLimit
Dim sVouCkTy,sLastVouDt,sCallVouTy,sPopVouTy,sFinPeriod,sFinTemp,sMaxDate,sMinDate
Dim iPreBookVal,objDOM
Dim sFinFrm,sFinTo,sValTemp2,sFormVal,sSelArg
Dim sConStatus

'Response.Write Request.ServerVariables("SCRIPTNAME")

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sFinTemp = Split(sFinPeriod,":")
	sMaxDate = "31/03/"&sFinTemp(1)
	sMinDate = "01/04/"&sFinTemp(0)
End IF

dim sAccount,sAddtional,iSno
dim dTotal
dim sVoucDate,iBookCode,sPayTo,sUserId,sVal,sValTemp
sUserId = getUserID()
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2
sCallVouTy = Request("VouTy")
sSelArg = Request("voutype")
sFormVal = Request("hFormVal")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("horgName")
sOrgId   = session("organizationcode")
sOrgName = session("orgshortname")

sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("selVouTypepe")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")

iPreBookVal = sBookCode
sConStatus = GetContraStatus(sTransNo)
'Response.Write "sTransNo="&sTransNo
sVal=Request("Val")
sValTemp=Split(sVal,"~")

IF Cstr(sVal) <> "" Then
	sQuery = "Select H.OUDEFINITIONID,D.OrgUnitDescription From DCS_OrganizationUnitDefinitions D, "&_
		 "Acc_T_CreatedVoucherHeader H Where H.OUDEFINITIONID = D.OUDEFINITIONID "&_
		 "and H.CreatedTransNo = "&sValTemp(0)&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sOrgId = objRs(0)
		sOrgName = objRs(1)
	End IF
	objRs.Close

Else

	sQuery = "Select Top 1 OUDefinitionID,OrgUnitDescription From DCS_OrganizationUnitDefinitions "&_
			 "Where Len(OUDefinitionID) > 4 Order By OUDefinitionID "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sOrgId = objRs(0)
		sOrgName = objRs(1)
	End IF
	objRs.Close
End IF

IF CStr(iBookAccHead) = "" Then
	sQuery = "Select Top 1 BookNumber,BookName,isNull(BookAccountHead,0),OtherUnitTransaction From vwOrgBookNames Where  "&_
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '02' Order By BookName "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sBookCode = objRs(0)
		sBookName = objRs(1)
		iBookAccHead = objRs(2)
		bOtherUnits = objRs(3)
	Else
		sBookCode = "02"
		sBookName = ""
		iBookAccHead = 0
		bOtherUnits = 1
	End IF
	objRs.Close
End IF

'Response.Write sTransNo


oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

'Response.Write sTransNo

oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_BA_"&Session.SessionID&".xml")

set Root = oDOM.DocumentElement

Root.setAttribute "BookNo",		Request.Form("hNewBookNo")
Root.setAttribute "BookAcchead",Request.Form("hNewBookAccHead")

oDOM.Save server.MapPath("../temp/transaction/Voucher Entry_BA_"&Session.SessionID&".xml")


IF CStr(sTransNo) = "" Then
	sTransNo = 0
End IF

IF CStr(sCallVouTy) = "R" Then
	sVouType = "D"
	sVouCkTy = "BAP"
	sPopVouTy = "D"
Else
	sVouType = "C"
	sVouCkTy = "BAR"
	sPopVouTy = "C"
End IF

'IF CStr(sVouType) = "C" Then
'	sVouCkTy = "BAP"
'Else
'	sVouCkTy = "BAR"
'End IF

sQuery = "Select Convert(Char,VoucherDate,103) From Acc_T_CreatedVoucherHeader Where CreatedTransNo =  "&_
		 "(Select Max(CreatedTransNo) From Acc_T_CreatedVoucherHeader Where BookCode = '02'  "&_
		 "and OUDefinitionID = '"&sOrgId&"' and TransactionType ='"&sVouCkTy&"' ) "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sLastVouDt = Trim(objRs(0))
End IF
objRs.Close

'sQuery = "Select CreatedBy From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sTransNo&" "
'objRs.Open sQuery,Con
'IF Not objRs.EOF Then
'	sUserId = Trim(objRs(0))
'End IF
'objRs.Close


'IF CStr(sUserId) = "" Then
sUserId = session("userid")
'End IF

sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp2(0))
sFinTo = Trim(sValTemp2(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"

Dim TempNode
Set Root = objDOM.CreateElement("Root")
objDOM.appendChild Root
Set Entrynode = objDOM.CreateElement("Details")
Root.appendChild Entrynode
'****************** This Voucher that is been Adjusted in Any Purchase/Sales Vouchers Will be
'****************** taken here for consideration *********************************************

sQuery = "Select H.CreatedVoucherNo,Convert(Varchar,H.VoucherDate,103) VoucherDate,H.VoucherAmount,C.AmountAdjusted, "&_
		 "A.CreatedAdvanceNo,C.CreatedTransNo,H.CrDrIndication From Acc_T_CreatedVoucherHeader H, "&_
		 "ACC_T_CreatedAdvanceAdj C,Acc_T_CreatedAdvances A Where A.CreatedTransNo = "&sTransNo&" "&_
		 "And A.CreatedAdvanceNo = C.CreatedAdvanceNo and C.CreatedTransNo = H.CreatedTransNo  "

'Response.Write sQuery
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountAdjusted")
	TempNode.setAttribute "AdjCrTrNo", objRs("CreatedTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CreatedAdvanceNo")
	TempNode.setAttribute "ValFrm", "CRADJADV"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close

sQuery = "Select CreatedVoucherNo,Convert(Varchar,VoucherDate,103) VoucherDate,VoucherAmount, "&_
		 "AmountPaid,AdjCrTransNo,CradvNo From CrPayAdvAdjDet Where CreatedTransNo = "&sTransNo
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountPaid")
	TempNode.setAttribute "AdjCrTrNo", objRs("AdjCrTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CradvNo")
	TempNode.setAttribute "ValFrm", "Pay"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close

sQuery = "Select CreatedVoucherNo,Convert(Varchar,VoucherDate,103) VoucherDate,VoucherAmount, "&_
		 "AmountReceived,AdjCrTransNo,CradvNo From CrRecAdvAdjDet Where CreatedTransNo = "&sTransNo
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountReceived")
	TempNode.setAttribute "AdjCrTrNo", objRs("AdjCrTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CradvNo")
	TempNode.setAttribute "ValFrm", "Rec"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close


objDOM.Save server.MapPath("../temp/transaction/Voucher ADJAMD_BA_"&Session.SessionID&".xml")

Response.Clear 

Response.Redirect "AmdAccGenerate.asp?hTransNo=" & sTransNo

%>