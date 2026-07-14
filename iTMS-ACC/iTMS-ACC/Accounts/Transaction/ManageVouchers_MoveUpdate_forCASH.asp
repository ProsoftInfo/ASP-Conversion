
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ManageVouchers_MoveUpdate_forCASH.asp
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
<!--#include virtual="/include/CheckACCPrevFinYear.asp"-->

<%
Dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,sQuery
Dim iVouNo,objRs,objRs1,sVouDate,bActionFlag,sVal,sValTemp
Dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sAccHead,sParType,sPartSubType
Dim iEnNo,Entrynode,HeaderNode,dOpeningBal
Dim sParCode,sNarration,sAccHeadname,sAccUnitName,bOtherUnits,iBookAccHead,dTransLimit
Dim sVouCkTy,sLastVouDt,sSelVouTy
Dim sFinPeriod,sFinFrm,sFinTo,sValTemp2,sFormVal,sSelArg
Dim sAccount,sAddtional,iSno
Dim dTotal,sConStatus
Dim sVoucDate,iBookCode,sPayTo,sUserId,iPreBookVal,sRetVal
Dim sExp,TempNode


sUserId = session("userid")
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,sLogUID

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")


sOrgId = Session("organizationcode")
sOrgName=Request.Form("orgshortname")
sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("selVouTypepe")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")
sSelVouTy = Request("VOUTY")

sSelArg = Request("voutype")
sFormVal = Request("hFormVal")

iPreBookVal = sBookCode

 'Response.Write sTransNo

sVal=Request("Val")
sValTemp=Split(sVal,"~")
IF CStr(sVouType) = "" Then
	IF CStr(sSelVouTy) = "P" Then
		sVouType = "C"
	Else
		sVouType = "D"
	End IF
End IF

'Response.Write sVouType

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
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '01' Order By BookName "


	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sBookCode = objRs(0)
		sBookName = objRs(1)
		iBookAccHead = objRs(2)
		bOtherUnits = objRs(3)
	Else
		iBookAccHead = 0
	End IF
	objRs.Close
End IF

'Response.Write sTransNo
oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

IF CStr(sTransNo) <> "" Then

 	sRetVal = GetVouchXML(sTransNo)
	oDOM.Load server.MapPath(sRetVal)
	oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")
End IF


set Root = oDOM.documentElement

'Response.write sTransNo
'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")

sConStatus = GetContraStatus(sTransNo)
IF CStr(sVouType) = "C" Then
	sVouCkTy = "CAP"
Else
	sVouCkTy = "CAR"
End IF

sQuery = "Select Convert(Char,VoucherDate,103),CreatedBy From Acc_T_CreatedVoucherHeader Where CreatedTransNo =  "&_
		 "(Select Max(CreatedTransNo) From Acc_T_CreatedVoucherHeader Where BookCode = '01'  "&_
		 "and OUDefinitionID = '"&sOrgId&"' and TransactionType ='"&sVouCkTy&"' ) "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sLastVouDt = Trim(objRs(0))
	'sUserId = Trim(objRs(1))
End IF
objRs.Close

'IF CStr(sUserId) = "" Then
sUserId = session("userid")
'End IF

sLogUID = session("userid")

sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp2(0))
sFinTo = Trim(sValTemp2(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"
bOtherUnits = 1

Root.setAttribute "BookNo",		Request.Form("hNewBookNo")
Root.setAttribute "BookAcchead",Request.Form("hNewBookAccHead")

oDOM.Save server.MapPath("../temp/transaction/Voucher Entry_CA_"&Session.SessionID&".xml")

Response.Clear 

Response.Redirect "AmdAccGenerate.asp?hTransNo=" & sTransNo

%>