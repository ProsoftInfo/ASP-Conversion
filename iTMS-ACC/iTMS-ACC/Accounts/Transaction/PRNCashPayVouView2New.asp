<%@ Language=VBScript %>
<%	option explicit%>
<%
Response.Expires=-1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name					:	PRNCashPayVouView2New.asp
	'Module Name					:
	'Author Name					:	S.MAHESWARI
	'Created On						:	23 JUL 2008
	'Modified On					:
	'Tables Used					:
	'Temporary Tables				:
	'Temporary Files				:
	'Input Parameter				:
	'Connects To					:
	'Procedures/Functions Used		:
	'Internal Variables				:
	'Database						:
	'Queries Used					:
	'Counters						:
	'String							:
	'Boolean						:
	'Object Holders					:
	'Description					:

%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/PrintFunctions.asp"-->
<!--#include File="../../include/GetOrganization.asp"-->
<%
'------------------------Declaration Constants -----------------------------
dim sPgPitch,sPrFooter,sPgMargin,sPgBreak,iPgLineNo,iRecCount,sDisplayHead
dim aiHeaderColWidth(5,11),PageSpace,sPagetitle2,sTstr
dim RootNode,VoucherNode,PartyNode,sLineBreakStr
dim sOrgID
dim sPartyCode,sExp
dim iTraNo,sBlankLine
dim iPageLen,iActualpgLen,iNoOfLinesCtr
Dim oDom,objTxt,objRs,sQuery,objFSO,Root


dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim iTransNo,iBkHeadCode,bOtherUnit,iTdsAmount,i
Dim sAddress1, sAddress2,sCity,sState,sPostcode,sTranIndication,sTranEntryIndication
Dim iCreatedBy,sCreatedOn,sVouStatus,sEmpName,iPartyCtrlAcc,sAdjType,hCrDr
Dim iEntryNo,iHeadOfAcc,sHeadOfAccName,iHeadOfAccAmt,iPartyCode,iEntryAmt
Dim iNetAmtPaid,iTotRecovered,sNarrFlag,sAddnFlag,iTotAddnlAmt,sAdjFlag,sAdjOn,iTotAdj
Dim objRsTemp,sNarr,sPartyName,sRetVal,sFlagTotRec,sFlagNetPaid,sAmtFlag,sNetFlag,sDetFlag
Dim sNarr1,sNarr2,sBillType,iVouAmt,sItemDesc,sum,iCount,iCtrUpto,iChkLine
Dim sRecovFlag,sAddPayFlag,sAddNewPayFlag

iNetAmtPaid = 0
iTotRecovered = 0
iNoOfLinesCtr = 0
iCtrUpto = 27
sNarrFlag = False
sAddnFlag = False
sAdjFlag = False

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request("Value")

sLineBreakStr = "\n"

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)

set Root = oDOM.documentElement
'Response.Write sRetVal

sOrgId = Root.Attributes.Item(0).nodeValue
sOrgName = Root.Attributes.Item(1).nodeValue
iBookCode = Root.Attributes.Item(2).nodeValue
sBookName = Root.Attributes.Item(3).nodeValue
sVouType = Root.Attributes.Item(4).nodeValue
sVoucDate = Root.Attributes.Item(5).nodeValue
iBkHeadCode = Root.Attributes.Item(6).nodeValue
'Response.Write "sVouType ="& sVouType 
'Response.End 
sApprove=Root.Attributes.Item(7).nodeValue
iVouNo=Root.Attributes.Item(9).nodeValue

set objRs = Server.CreateObject("ADODB.Recordset")
set objRsTemp = Server.CreateObject("ADODB.Recordset")

sQuery="select OtherUnitTransaction from vwOrgBookNames where OUDefinitionID = '" & sOrgId &"' and BookNumber="&iBookCode&" and BookCode='01' "

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

if not 	objRs.EOF then
	bOtherUnit=objRs(0)
else
	bOtherUnit=0
end if
objRs.Close

'==============================Voucher View header queries starts here ============================

sQuery = "select Distinct CreatedVoucherNo,TransactionType,OUDefinitionID,isNull(PayToRecdFrom,'')," &_
		"PartyCode,CreatedBy,Convert(varchar,CreatedOn,103),CreatedVouchStatus" &_
		" from VW_Created_CashVoucherView where CreatedTransNo="& iTransNo
'Response.Write sQuery

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iVouNo = objRs(0)
	sVouType = objRs(1)
	sOrgId = objRs(2)
	sPayTo = objRs(3)
	iPartyCode = objRs(4)
	iCreatedBy = objRs(5)
	sCreatedOn = objRs(6)
	sVouStatus = objRs(7)
end if
objRs.Close

sQuery = "Select isnull(PurchaseBillType,''),VoucherAmount,CrDrIndication from Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
'Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
if not 	objRs.EOF then
	sBillType = objRs(0)
	iVouAmt	  = objRs(1)
	hCrDr	  = objRs(2)
End If
objRs.Close

If trim(sVouStatus) = "010104" Then
	sQuery = "select VoucherNumber from ACC_T_VoucherHeader where CreatedTransNo="&iTransNo
	with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
	end with
	if not 	objRs.EOF then
		iVouNo = objRs(0)
	End If
	objRs.Close
End If
set objRs.ActiveConnection = nothing
If trim(sVouType) = "CAP" Then
	sTranIndication = "C"
	sTranEntryIndication = "D"
Else
	sTranIndication = "D"
	sTranEntryIndication = "C"
End If
sQuery = "Select VoucherEntryNumber,AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"'"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
'Response.write sQuery
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iEntryNo = objRs(0)
	iHeadOfAcc = objRs(1)
	iHeadOfAccAmt = FormatNumber(objRs(2),2,,,-2)
	sNarr = objRs(3)
	iPartyCtrlAcc = objRs(4)
end if
objRs.Close

'Response.Write iPartyCode
IF iHeadOfAcc <> "" Then
	sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
Else
	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
End If
'Response.write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sHeadOfAccName = objRs(0)
end if
objRs.Close
'Response.Write "sHeadOfAccName="&sHeadOfAccName


sQuery = "select OrgUnitDescription,Address1,Address2,City,State,PostCode from DCS_OrganizationUnitDefinitions where OUDefinitionID='"&sOrgId&"'"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sOrgName = objRs(0)
	sAddress1 = objRs(1)
	sAddress2 = objRs(2)
	sCity = objRs(3)
	sState = objRs(4)
	sPostcode = objRs(5)
else
	sOrgName = ""
	sAddress1 = ""
	sAddress2 = ""
	sCity = ""
	sState = ""
	sPostcode = ""
end if
objRs.Close

If iPartyCode <> "0" Then
	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCode

	With Objrs
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.Open
	End With
	set objRs.ActiveConnection = nothing
	IF not objRs.EOF  Then
		sPartyName = Trim(Objrs(0))
	End IF
	objRs.Close
End If

	sQuery = "SELECT LoginId FROM DCS_User WHERE InternalUserID ="&iCreatedBy
With Objrs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
set objRs.ActiveConnection = nothing
IF not objRs.EOF  Then
	sEmpName = Trim(Objrs(0))
End IF
objRs.Close
	
dim sAccHeadName, sTempAddlPayRec, sTemp,sTempNarrRec,sTempRecevAdj,sTempPayAdj
sAccHeadName = sHeadOfAccName
'==============================Voucher print header queries ends here ============================

'No and Date
aiHeaderColWidth(0,0)=67
aiHeaderColWidth(0,1)=16
aiHeaderColWidth(0,2)=84

'Head of Account, Amount
aiHeaderColWidth(1,0)=14
aiHeaderColWidth(1,1)=70
aiHeaderColWidth(1,2)=23
'Paid To, Rupees
aiHeaderColWidth(2,0)=14
aiHeaderColWidth(2,1)=46
aiHeaderColWidth(2,2)=7
aiHeaderColWidth(2,3)=16

'Details
aiHeaderColWidth(3,0)=3
aiHeaderColWidth(3,1)=57
aiHeaderColWidth(3,2)=3
aiHeaderColWidth(3,3)=21
aiHeaderColWidth(3,4)=9
aiHeaderColWidth(3,5)=50
aiHeaderColWidth(3,6)=13
aiHeaderColWidth(3,7)=90


'Prepared by, Checked by, Passed by
aiHeaderColWidth(4,0)=3
aiHeaderColWidth(4,1)=16
aiHeaderColWidth(4,2)=3
aiHeaderColWidth(4,3)=18
aiHeaderColWidth(4,4)=2
aiHeaderColWidth(4,5)=18

'Sub Details
aiHeaderColWidth(5,0)=4
aiHeaderColWidth(5,1)=39
aiHeaderColWidth(5,2)=15
aiHeaderColWidth(5,3)=4
aiHeaderColWidth(5,4)=11
aiHeaderColWidth(5,5)=6

%>

<%

'------------------------End of Declaration Constants ----------------------
%>
<%


set objFSO = Server.CreateObject("Scripting.FileSystemObject")
set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Reports/"& Session.SessionID &"_CashVoucher.txt"))
Dim sText,sTempRecoveries


	dim nCashVoucherNo,dCashVoucherDate,sHeadOfAccount,sAmount,sPaidTo,iAmount
	dim sDetails, sReceivedPayment,sPreparedBy,sCheckedBy,sPassedBy, iCtr
	Dim sDesc,sDesc2,iwidth
	'Assigning Hardcoded values for Variables
	nCashVoucherNo		= iVouNo
	dCashVoucherDate	= sVoucDate

	IF trim(sBillType) = "C" then
		sDesc	= AmountWords(replace(iVouAmt,",",""))
	Else
		sDesc	= AmountWords(replace(iHeadOfAccAmt,",",""))
	End IF

	sDesc2 = ""
	If Len(sDesc) > 50 Then
		For i = 1 to 50
			If Mid(sDesc,50-i,1) = " " Then
				iWidth = 50-i
			Exit For
			End if
		Next
		sDesc2 = Mid(sDesc,iWidth+1,Len(sDesc))
		sDesc =  Mid(sDesc,1,50-i)
	End If

	sPaidTo = sPayTo
	iAmount	= iHeadOfAccAmt
	sDetails = UCase(sNarr)
	sReceivedPayment = "ReceivedPayment"



	IF Cstr(sVouStatus) <> "010104" Then
		sPreparedBy	= sEmpName &"/" & nCashVoucherNo &" - " & sCreatedOn

	Else
		sPreparedBy	= sEmpName & "-" & sCreatedOn
	End IF



	sCheckedBy	= "CheckedBy"
	sPassedBy	= "PassedBy"
	sTemp = ""
	'Blank Lines
	'stext =stext & formattprint("10Pitch","")
	stext = stext & chr(10) 'line feed character
	stext = stext & "" & vbcrlf
	


	'Number and date
	sTemp = sTemp & myAlign("",aiHeaderColWidth(0,0)+6,"L")
	IF Cstr(sVouStatus) = "010104" Then
		sTemp = sTemp & myAlign(nCashVoucherNo,7,"L")
		if trim(sVouType)="CAR" then
		    sTemp = sTemp & myAlign("Receipt",7,"R")
		end if'if trim(sVouType)="CAR" then
	Else
		sTemp = sTemp & myAlign(" ",7,"L")
	End IF
	sText = sText & sTemp
	sText = sText & vbCrLf
	sTemp = ""
	stext = stext & " " & vbcrlf
	sTemp = sTemp & myAlign("",aiHeaderColWidth(0,0)+6,"L")
	'sTemp = sTemp & myAlign("",aiHeaderColWidth(0,2),"L")'aiHeaderColWidth(0,2)=84
	sTemp = sTemp & myAlign(dCashVoucherDate,aiHeaderColWidth(0,1),"L")
	sText = sText & sTemp
	sText = sText & " " & vbCrLf &" " & vbCrLf &" "
	sTemp = ""
	'Blank Lines
	stext = stext & " " & vbcrlf
	'Head of Account
	'Response.Write "sHeadOfAccount="&sAccHeadName
	'Response.End
	sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+8,"L")
	'sTemp = sTemp & myAlign("",aiHeaderColWidth(1,2),"L")'aiHeaderColWidth(1,2)=23
	sTemp = sTemp & myAlign(sAccHeadName,aiHeaderColWidth(1,1),"L")
	sText = sText & sTemp
	sText = sText & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf
	sTemp = ""
	sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+8,"L")
	'sTemp = sTemp & myAlign("",aiHeaderColWidth(1,2),"L")'aiHeaderColWidth(1,2)=23
	sTemp = sTemp & myAlign(sDesc,aiHeaderColWidth(1,1),"L")
	sTemp = sTemp & vbCrLf
	if trim(sDesc2) <> "" then
		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0)+8,"L")
		'sTemp = sTemp & myAlign("",aiHeaderColWidth(1,2),"L")'aiHeaderColWidth(1,2)=23
		sTemp = sTemp & myAlign(sDesc2,aiHeaderColWidth(1,1),"L")& vbCrLf
	Else
		sTemp = sTemp & vbCrLf
	end if 'if trim(sDesc) <> "" then
	sText = sText & sTemp

	sTemp = ""
	sText = sText & " " & vbCrLf
	
	'sPaidTo = "To Check "
	if trim(sVouType)="CAR" then
	    sTemp = sTemp & myAlign("",7,"R")
    	sTemp = sTemp & myAlign("Received From",15,"L")
	else
	    sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)+8,"R")
	end  if
	sTemp = sTemp & myAlign(sPaidTo,aiHeaderColWidth(2,1)+1,"L")
	sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0)-10,"L")
	sTemp = sTemp & myAlign("",1,"L")

'Response.Write iAmount
'2,0 = 14 2,1 = 46
	
	IF trim(sBillType) = "C" then
		sTemp = sTemp  & myAlign(FormatNumber(iVouAmt,2,,,-2),aiHeaderColWidth(2,1),"L")
	Else
		sTemp = sTemp &  myAlign(iAmount,aiHeaderColWidth(2,1),"L")
	End IF

	sText = sText & sTemp  &" " & vbCrLf
	
	sTemp = "" & formattprint("CONDENSESTART","") 'Condensed formatt
	sText = sText & sTemp  &" " & vbCrLf & vbCrLf
	

	iNoOfLinesCtr = 15
	
	sTemp = ""
	'sTemp = sTemp & myAlign("",20,"L")
	'sTemp = sTemp & myAlign("Details start : " + trim(iNoOfLinesCtr),20,"L") ' kk remove
	
	'note : this line will be added after functions call ref: AddLineHere
	sText = sText & sTemp  &" " & vbCrLf ' kk add 
	
Dim TCrDr, sNextLine,sHeadFlag,rs,sDivFlag
Dim sTotFlag,sRecFlag,sNarrCtr,j,CrDr
Dim iFinTotAmt,iFinTotRec,iFinNetPaid
Set rs = Server.CreateObject("ADODB.RecordSet")

 
sQuery = "select CrDrIndication from Acc_T_CreatedVoucherHeader where createdtransno = "& iTransNo &" "
	rs.Open sQuery,con
	if Not rs.EOF then 
		CrDr = rs(0)
	end if
	rs.Close 
	iFinTotAmt = 0
	iFinTotRec = 0
	iFinNetPaid = 0
	'Total Amount
	sQuery = "Select sum(isnull(Amount,0)) from Acc_T_CreatedVoucherdetails where TransCrDrIndication <> '"&CrDr&"' and  createdtransno = "& iTransNo &" "
	rs.Open sQuery,con
	if Not rs.EOF then 
		iFinTotAmt = rs(0)
	else
		iFinTotAmt = 0
	end if 
	rs.Close
	'Total Recovered
	sQuery = "Select sum(isnull(Amount,0)) from Acc_T_CreatedVoucherdetails where TransCrDrIndication = '"&CrDr&"' and  createdtransno = "& iTransNo &" "
	'Response.Write sQuery
	rs.Open sQuery,con
	if Not rs.EOF then 
		iFinTotRec  = rs(0)
	else
		iFinTotRec = 0
	end if 
	rs.Close
	IF trim(iFinTotAmt) = "" then iFinTotAmt = 0
	IF trim(iFinTotRec) = "" then iFinTotRec = 0
	If trim(iFinTotAmt) <> "0" and trim(iFinTotRec) <>  "0" then
		iFinNetPaid = cdbl(iFinTotAmt) - cdbl(iFinTotRec)
	Else
		iFinNetPaid = iFinTotAmt
	End IF
	
	GotoRecoveries()
	GotoAddPaymentsNew(sBillType)
	GotoAddPayments(sBillType)
	
	'Response.Write  sRecovFlag &"==="& sAddPayFlag&"==="& sAddNewPayFlag
	'
	' for testing
	Dim sFinalData,nLen1,nFindLines
	
	nFindLines = 0
	nLen1 = 1
	sFinalData = sText
	
	do while trim(sFinalData) <> ""
	
		nLen1 = instr(1,sFinalData,sLineBreakStr )
		'Response.Write "<p>" & nLen1
		'Response.Write "<p>" &  mid(sFinalData,1,nLen1+1)
		if nLen1 > 0 then
			nFindLines = nFindLines + 1
		else
			exit do	
		end if 
		
		sFinalData = mid(sFinalData,nLen1+len(sLineBreakStr))
	loop
	''
	nFindLines = nFindLines + 1
	'Response.Write "<p>nFindLines =  "  & nFindLines
	'Response.End 
	
	'ref: AddLineHere
	'iNoOfLinesCtr  = 15 + nFindLines
	iNoOfLinesCtr  = 16 + nFindLines
	
	sText = Replace(sText,sLineBreakStr,vbCrLf)
	
	sTemp = ""
	
	 	
	IF sDetFlag = True then 
		sTemp = sTemp & myAlign(" ",9,"L")
		sTemp = sTemp & myAlign(" ",6,"L")
		sTemp = sTemp & myAlign("Total Recovered  ",aiHeaderColWidth(5,1)-4,"R")
		sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
		sTemp = sTemp & myAlign(FormatNumber(iTotRecovered,2,,,-2),aiHeaderColWidth(5,2),"R")
		sTemp = sTemp & vbCrLf
		iNoOfLinesCtr  = iNoOfLinesCtr + 1 
	End IF 'IF sDetFlag = True then
	'Response.Write sAddPayFlag
	'Response.end
	
		
	
	IF sRecFlag = True and  (sAddPayFlag = True or sAddNewPayFlag = true ) then
		
		for i = iNoOfLinesCtr to 30 ' 32
			sTemp = sTemp & vbCrLf 
			'sTemp = sTemp & "xxx = " & trim(iNoOfLinesCtr) & vbCrLf  ' kk remove
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		Next
		
	Else
		'Response.Write "<p> " & iNoOfLinesCtr
	'Response.End 	
		for i = iNoOfLinesCtr to 28 ' 30
			sTemp = sTemp & vbCrLf
			'sTemp = sTemp & trim(iNoOfLinesCtr) & vbCrLf  ' kk remove
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		Next
	End If 
	
	
 
	sTemp = sTemp & myAlign("",aiHeaderColWidth(4,0)+14,"L")
	sTemp = sTemp & myAlign(sPreparedBy,aiHeaderColWidth(4,1)+11,"L")
	'sTemp = sTemp & myAlign(trim(iNoOfLinesCtr) & " " & sPreparedBy,aiHeaderColWidth(4,1)+11,"L")' kk remove
	sTemp = sTemp & myAlign("",aiHeaderColWidth(4,0),"L")
	
	sText = sText & sTemp
	sText = sText & vbCrLf
	sText = sText & vbCrLf
	sTemp = ""
	sText = sText & formattprint("CONDENSEEND","") 'Condensed formatt
	sText = sText & vbCrLf
	sText = sText & vbCrLf
	sText = sText & vbCrLf

	sText = sText & vbCrLf
	sText = sText & vbCrLf
	
	

	'sText = sText & chr(12)

'		Response.End


	objTxt.write sText
	Response.Redirect "../../Components/FormattPrintNew.asp?server=server&filepath=/Accounts/temp/Reports/"& Session.SessionID &"_CashVoucher.txt&exitpath=/Accounts/reports/PRNYarnInvoiceCumDDView.asp&frame=_parent"
%> 


<% 

'================================= USER DEFINED FUNCTIONS ================================='
'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function centerAlign(str1,width)
		dim diff,strlen,val, i, str, newstr, blank
			str = str1
			strlen = len(str)
			diff = width - strlen
			for i=0 to (diff-1)/2
				blank = blank & " "
			next
			newstr = blank & str & blank
		centerAlign = newstr
	end function
	'------------------------End OF myAlign Function----------------------------
%>

 
<%


	' FUNCTIONS START
	Function printDetails()
	Dim sPreparedBy,sCheckedBy,sPassedBy,sEmpName
		sPreparedBy	= xPreparedBy
		sCheckedBy		= xCheckedBy
		sPassedBy		= xPassedBy
		sEmpName		= xEmpName
			'Prepared by, Checked by, Passed by
		sTemp = ""
			sTemp = sTemp & myAlign("",aiHeaderColWidth(4,0),"L")
			sTemp = sTemp & myAlign(sEmpName,aiHeaderColWidth(4,1),"L")
			sTemp = sTemp & myAlign("Chek",aiHeaderColWidth(4,0),"L")
			sTemp = sTemp & myAlign(sCheckedBy,aiHeaderColWidth(4,1),"L")
			sTemp = sTemp & myAlign("Pass",aiHeaderColWidth(4,0),"L")
			sTemp = sTemp & myAlign(sPassedBy,aiHeaderColWidth(4,1),"L")
		objtxt.write sTemp & vbCrLf & vbCrLf
	End Function

	Function printHeadAmtPaid(temp, xPayTo)
	dim iAmount, sTemp,sPaidTo
		iAmount = temp
		sPaidTo = xPayTo
		'Head of Account
		sTemp = ""
		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0),"L")
		sTemp = sTemp & myAlign(sPaidTo,aiHeaderColWidth(1,1),"L")
		objtxt.write sTemp & " " & vbCrLf & " " & vbCrLf
		'Amount
		sTemp = ""
		sTemp = sTemp & myAlign("",aiHeaderColWidth(1,0),"L")
		IF trim(sBillType) = "C" then
			sTemp = sTemp & myAlign(AmountWords(iVouAmt),aiHeaderColWidth(1,1),"L")
		Else
			sTemp = sTemp & myAlign(AmountWords(iAmount),aiHeaderColWidth(1,1),"L")
		End IF

		objtxt.write sTemp & " " &  vbCrLf & " " & vbCrLf
		'PaidTo, Rs
		sTemp = ""
			sTemp = sTemp & myAlign("",aiHeaderColWidth(2,0),"L")
			sTemp = sTemp & myAlign(sPaidTo,aiHeaderColWidth(2,1),"L")
			sTemp = sTemp & myAlign("",aiHeaderColWidth(2,2),"L")
			sTemp = sTemp & myAlign(iAmount,aiHeaderColWidth(2,3),"L") 
			objtxt.write sTemp & vbCrLf
	End Function
	'FUNCTIONS END
%> 

<% 'Fetch from Recoveries
Function GotoRecoveries()
	sQuery = "Select Count(*) From Acc_T_CreatedVoucherDetails Where CreatedTransNo = "& iTransNo &" and TransCrDrIndication='"& hCrDr &"' and VoucherEntryNumber <> "& iEntryNo&" "
	'esponse.Write sQuery
	'Response.End


	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	if not 	objRs.EOF then
		TCrDr = objRs(0)
	end if
	objRs.Close
	

	IF  TCrDr >= 1 then
		iNoOfLinesCtr = iNoOfLinesCtr + 3
		sRecovFlag = True
		sQuery = "Select count(1) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iTransNo&" and Amount <> 0 and TransCrDrIndication='"&sTranIndication&"'"
		'Response.Write sQuery

		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		if not 	objRs.EOF then
			iCount = objRs(0)
		end if
		objRs.Close

		sQuery = "Select VoucherNarration from  Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' "
		objRs.Open sQuery,con
		if not objRs.EOF then
			sNarr = objRs(0)
		end if
		objRs.Close

		IF sNarrFlag <> True then 
			IF sNarr <> "" then
				sNarrFlag = True

				j = len(sNarr) / 90
				'Response.Write  cint(j)
				
				IF Len(UCase(sNarr)) > 90 then

					For i = 1 to Len(sNarr) step 90
						sum = i + 90
						IF sum > len(sNarr) then
							if iCount <= 0 then sTemp = sTemp & sLineBreakStr

							sTemp = sTemp & myAlign(" ",aiHeaderColWidth(4,3)-2,"L")
							sTemp = sTemp & myAlign(" " & UCase(mid(sNarr,i,90)),aiHeaderColWidth(3,7)+2,"L")'90 bytes
							sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L") '3 bytes
							'Total Rec Rs.

							IF iCount > 0 then
								IF iNoOfLinesCtr = 18 then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sTotFlag = True
								ElseIF iNoOfLinesCtr = 19 then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sRecFlag = True
								ElseIF iNoOfLinesCtr = 20 then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sNetFlag = True
								End IF
							End IF 'IF iCount > 0 then
						Else
							
							sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
							sTemp = sTemp & myAlign(" " & UCase(mid(sNarr,i,90)),aiHeaderColWidth(3,7)+2,"L")
							sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L") '3 bytes
							'Total Rec Rs.
							IF iCount > 0 then
								IF iNoOfLinesCtr = 18 and sTotFlag <> True then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sTotFlag = True
								ElseIF iNoOfLinesCtr = 19 and sRecFlag <> True then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sRecFlag = True
								ElseIF iNoOfLinesCtr = 20 and sNetFlag <> True then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sNetFlag = True
								Else
									sTemp = sTemp & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
								End IF
							End IF 'IF iCount > 0 then
						End If
						sNarrCtr = sNarrCtr +1
						 
					Next
					sNextLine = True
				Else
	   				sNarrCtr = sNarrCtr +1
	   				
					'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
					sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
					sTemp = sTemp & myAlign(" " &UCase(sNarr),aiHeaderColWidth(3,7)+1,"L")
					sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
					 
					IF iCount > 0 then
						'Total Amount Rs.
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
						sTemp = sTemp & myAlign("Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
						'sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
						'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
						sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr

						'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
						sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-3,"L") '18 bytes
						'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L") '50 bytes
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L") '90 bytes
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes
						'Total Rec Rs.
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
						sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
						'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
						sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr

						'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
						sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-3,"L") '18 bytes
						'Response.Write "sNarrCtr="&sNarrCtr&"<BR>"
						IF sNarrCtr = 1 then
							'sTemp = sTemp & myAlign("HHHHHHHHHHHHH",aiHeaderColWidth(3,5),"L") '50 bytes
							sTemp = sTemp & myAlign("HHHHHHHHHHHHH",aiHeaderColWidth(3,7),"L") '90 bytes
							sHeadFlag = True
							iNoOfLinesCtr = iNoOfLinesCtr + 1
						Else
							'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L") '50 bytes
							sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L") '90 bytes
						End If
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes
						'Net Paid Rs.
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+8,"L")
						sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
						'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
						sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr

						sTotFlag = True
						sRecFlag = True
						sNetFlag = True
						sNextLine = False

						iNoOfLinesCtr = iNoOfLinesCtr + 3 
						'Response.Write sTemp &"<br>"
					End IF 'IF iCount > 0 then
				End IF
				'sNarrFlag = True
			Else
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L")
			 	sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L")
			 	sNarrCtr = sNarrCtr +1
			End IF 'IF sNarr <> "" then
		 
		End IF 'IF sNarrFlag <> True then
 
		'End If 'IF sNarrFlag  <> True then
		'sText = sText & sTemp
		'IF sNarrFlag = True then
		IF iCount > 0 then
			IF iNoOfLinesCtr = 18 and sTotFlag <> True then
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
				'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L") '50 bytes
				'sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes

				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
				sTemp = sTemp & myAlign("Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				'sTemp = sTemp & myAlign(FormatNumber(iVouAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
				sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sTotFlag = True
			End IF
			IF iNoOfLinesCtr = 19 and sRecFlag <> True then
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L") '50 bytes
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes

				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
				sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
				sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sRecFlag = True
			End IF
			IF iNoOfLinesCtr = 20 and sNetFlag <> True then
				'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
				sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-3,"L")
				IF sNarrCtr = 1 then
					sTemp = sTemp & myAlign("HHHHHHHHHHHHH",aiHeaderColWidth(3,5)-5,"L") '50 bytes
					sHeadFlag = True
					iNoOfLinesCtr = iNoOfLinesCtr + 1
				Else
					sTemp = sTemp & myAlign("",aiHeaderColWidth(3,5),"L") '50 bytes
					sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes
				End If
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+51,"L")
				sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
				sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sNetFlag = True
			End IF
		End IF 'IF iCount > 0 then
		'sText = sText & sLineBreakStr
		'iNoOfLinesCtr = iNoOfLinesCtr + 1 'aaaaa
		'Fetch From Recoveries
		sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> "&iEntryNo&" and TransCrDrIndication='"&sTranIndication&"'  and Amount <> 0"
			' Response.Write sQuery	& sLineBreakStr
			'Response.End
		iCtr = 1
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		if not 	objRs.EOF then
			iChkLine = iCtrUpto - iNoOfLinesCtr
			'Response.Write iCount &"----"& iChkLine &"---"&iNoOfLinesCtr
			'Response.End

			'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
			'sTemp = sTemp & myAlign(" RECOVERIES :",aiHeaderColWidth(3,4)+5,"L")

			sTemp = Replace(sTemp,"HHHHHHHHHHHHH"," RECOVERIES :  ")
			
			IF sHeadFlag <> True then
				sTemp = sTemp &" " &sLineBreakStr
				sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
				sTemp = sTemp & myAlign(" RECOVERIES :",aiHeaderColWidth(3,5)+2,"L")&sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 2
			End IF
			Do While not objRs.EOF
				iHeadOfAcc = objRs(0)
				iEntryAmt = cdbl(objRs(1))
				iTotRecovered = iTotRecovered + iEntryAmt
				sDetFlag = True
				'Response.Write iTotRecovered
				If iHeadOfAcc <> "" Then
					sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
				Else
					sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
				End If
				with objRsTemp
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				Set objRsTemp.ActiveConnection = Nothing
				sHeadOfAccName = ""
				if not 	objRsTemp.EOF then
					sHeadOfAccName = objRsTemp(0)
				end if
				objRsTemp.Close
			'	Response.Write "<P>"&iCount&"="&iChkLine &"<BR>"
			'	Response.End
				'sTemp = sTemp & sLineBreakStr
				IF iCount <= iChkLine then
				'	Response.Write "<P>sHeadOfAccName="&sHeadOfAccName &"<BR>"
					sTemp = sTemp & myAlign(" ",aiHeaderColWidth(4,3)-2,"L")
					sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
					sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
					sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R") & sLineBreakStr
					iNoOfLinesCtr = iNoOfLinesCtr + 1
					
					sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
					
					 sDivFlag = True
				Else
		 			sTemp = sTemp & myAlign(" ",aiHeaderColWidth(4,3)-2,"L")
					sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
					sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R")
					sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
					'iNoOfLinesCtr = iNoOfLinesCtr + 1
				End IF
					
				dim sRepflag
				If iCtr < objrs.RecordCount Then
					objRs.MoveNext
					If Not objRs.EOF Then
						iHeadOfAcc = objRs(0)
						iEntryAmt = cdbl(objRs(1))
						iTotRecovered = iTotRecovered + iEntryAmt
						sDetFlag = True
						If iHeadOfAcc <> "" Then
							sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
						Else
							sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
						End If
						with objRsTemp
							.CursorLocation = 3
							.CursorType = 3
							.Source = sQuery
							.ActiveConnection = con
							.Open
						end with
						Set objRsTemp.ActiveConnection = Nothing
						sHeadOfAccName = ""

						if not 	objRsTemp.EOF then
							sHeadOfAccName = objRsTemp(0)
						end if
						objRsTemp.Close
						sRepflag = True

						'Response.Write sHeadOfAccName &sLineBreakStr
						IF iCount <= iChkLine then

							sTemp = sTemp & myAlign(" ",aiHeaderColWidth(3,4)-2,"L")
							sTemp = sTemp & myAlign(" ",7,"L")
							sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
							sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
							'iNoOfLinesCtr = iNoOfLinesCtr + 1
							sDivFlag = true
						Else
							'sTemp = sTemp & myAlign("E ",aiHeaderColWidth(3,4)-2,"L")
							sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
							sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R")
							'iNoOfLinesCtr = iNoOfLinesCtr + 1
						End IF
						
						sTemp = sTemp & sLineBreakStr
						iNoOfLinesCtr = iNoOfLinesCtr + 1
						iCtr = iCtr + 1
					End If

				End If
				If trim(sBillType) = "C" then
					iNetAmtPaid =  iVouAmt
				Else
					iNetAmtPaid = iHeadOfAccAmt + iTotAddnlAmt - iTotRecovered
				End IF
					
			objRs.MoveNext
			iCtr = iCtr + 1
			Loop
		end if
		objRs.Close
		sText = sText & sTemp
'	iNoOfLinesCtr = iNoOfLinesCtr + 1
	End IF 'IF  TCrDr >= 1 then
	sTemp = ""
	
End Function
%>

<%	'Additional Payments

Function GotoAddPayments(sBillType)
 	 
	If trim(sBillType) <> "C" then
		'sTemp = sTemp &  formattprint("CONDENSESTART","") 'Condensed formatt
		
		 
		
		'IF sRecovFlag <> True then iNoOfLinesCtr = iNoOfLinesCtr + 1
		sQuery = "Select count(1) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iTransNo&" and Amount <> 0 and TransCrDrIndication='"&sTranEntryIndication&"'"
		'Response.Write sQuery
		'Response.End
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		if not 	objRs.EOF then
			iCount = objRs(0)
		end if
		objRs.Close 

		sQuery = "Select VoucherNarration from  Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' "
		objRs.Open sQuery,con
		if not objRs.EOF then
			sNarr = objRs(0)

		end if
		objRs.Close
  
		If Trim(sBillType) = "C" then
			sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode,isNull(ItemDescription,'') from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' and Amount <> 0"
		Else
			sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode,isNull(ItemDescription,'') from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' and VoucherEntryNumber <> "&iEntryNo&"  and Amount <> 0"
		End If

		 
		iCtr = 1
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		if 	objRs.EOF then		
			iCount =0
		else
			sAddPayFlag = True
		end if
		 '9 bytes

		sNarrCtr = 0
		IF sNarrFlag <> True then 
			IF sNarr <> "" then
				sNarrFlag = True

				j = len(sNarr) / 90
				'Response.Write  cint(j)
				'
				IF Len(UCase(sNarr)) > 90 then
			
					For i = 1 to Len(sNarr) step 90
						sum = i + 90

						IF sum > len(sNarr) then
							'Response.Write "1="& mid(sNarr,i,90) & "<br>"
							if iCount <= 0 then sTemp = sTemp & sLineBreakStr
							sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
							sTemp = sTemp & myAlign(" " & UCase(mid(sNarr,i,50)),aiHeaderColWidth(3,7)+2,"L")'90 bytes
							sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L") '3 bytes
							'Total Rec Rs.

							IF iCount > 0 then
								IF iNoOfLinesCtr = 18 then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Total Amountt Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sTotFlag = True
								ElseIF iNoOfLinesCtr = 19 then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sRecFlag = True
								ElseIF iNoOfLinesCtr = 20 then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sNetFlag = True
								End IF
							End IF 'IF iCount > 0 then
						Else
							sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
							sTemp = sTemp & myAlign(" " & UCase(mid(sNarr,i,90)),aiHeaderColWidth(3,7)+2,"L")
							sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L") '3 bytes
							'Total Rec Rs.
							IF iCount > 0 then
								IF iNoOfLinesCtr = 18 and sTotFlag <> True then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sTotFlag = True
								ElseIF iNoOfLinesCtr = 19 and sRecFlag <> True then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sRecFlag = True
								ElseIF iNoOfLinesCtr = 20 and sNetFlag <> True then
									sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
									sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
									sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
									'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
									sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
									iNoOfLinesCtr = iNoOfLinesCtr + 1
									sNetFlag = True
								End IF
							End IF 'IF iCount > 0 then
						End If
						sNarrCtr = sNarrCtr +1
					Next
					sNextLine = True
				Else
					IF sRecovFlag <> true then 	
						sTemp = sTemp & sLineBreakStr& sLineBreakStr  
						iNoOfLinesCtr = iNoOfLinesCtr + 2
					End IF
		   			sNarrCtr = sNarrCtr +1
					'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
					sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
					sTemp = sTemp & myAlign(" " & UCase(sNarr),aiHeaderColWidth(3,7)+2,"L")
					sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				
					IF iCount > 0 then
						'Total Amount Rs.
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
						sTemp = sTemp & myAlign("Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
						'sTemp = sTemp & myAlign(FormatNumber(iVouAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
						'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
						sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr

						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L") '90 bytes
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes
						'Total Rec Rs.
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+17,"L")
						sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
						'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
						IF trim(iFinTotRec) <> 0 and trim(iFinTotRec) <> "" then 
							sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
						Else
							sTemp = sTemp & myAlign("0.00",aiHeaderColWidth(3,6),"R") & sLineBreakStr
						End IF

						'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
						sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-3,"L")
						'Response.Write "sNarrCtr="&sNarrCtr&"<BR>"
						IF sNarrCtr = 1 then
							sTemp = sTemp & myAlign("HHHHHHHHHHHHH",aiHeaderColWidth(3,7),"L") '90 bytes
							sHeadFlag = True 
						Else
							sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L") '90 bytes
							sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes
						End If
						'Net Paid Rs.
						sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+5,"L")
						sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
						'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
						sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr

						sTotFlag = True
						sRecFlag = True
						sNetFlag = True
						sNextLine = False

						iNoOfLinesCtr = iNoOfLinesCtr + 3
						'Response.Write sTemp &"<br>"
					End IF 'IF iCount > 0 then
				End IF
				'sNarrFlag = True
			Else
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L")
			 	sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L")
			 	sNarrCtr = sNarrCtr +1
			End IF 'IF sNarr <> "" then
		Else
			IF trim(iCount) <> "0"  then  
				sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L") '18 bytes
				sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",aiHeaderColWidth(3,7),"L") &sLineBreakStr'90 bytes
				'sHeadFlag = True
				 iNoOfLinesCtr = iNoOfLinesCtr + 1
			End IF
				
		End IF 'IF sNarrFlag <> True then 
		 

		IF iCount > 0 then
			IF iNoOfLinesCtr = 18 and sTotFlag <> True then
			 
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+26,"L")
				sTemp = sTemp & myAlign("Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				'sTemp = sTemp & myAlign(FormatNumber(iVouAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				'sTemp = sTemp & myAlign("AAA",aiHeaderColWidth(3,6),"R") & sLineBreakStr
				sTemp = sTemp & myAlign(FormatNumber(iFinTotAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sTotFlag = True
			End IF
			IF iNoOfLinesCtr = 19 and sRecFlag <> True then
				'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
				sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L") '18 bytes
				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L") '90 bytes
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes

				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
				sTemp = sTemp & myAlign("(-)Recovered Rs.",aiHeaderColWidth(3,6)+3,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				'sTemp = sTemp & myAlign("DDDDDDDDDDDDD",aiHeaderColWidth(3,6),"R") & sLineBreakStr
				sTemp = sTemp & myAlign(FormatNumber(iFinTotRec,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sRecFlag = True
			End IF
			IF iNoOfLinesCtr = 20 and sNetFlag <> True then
			'	sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
				sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-3,"L")
				IF sNarrCtr = 1 then
					sTemp = sTemp & myAlign("HHHHHHHHHHHHH",aiHeaderColWidth(3,7),"L") '90 bytes
					sHeadFlag = True
					 sTemp = sTemp & sLineBreakStr 
				Else
					sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L") '90 bytes
					sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes
				End If

				sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+6,"L")
				sTemp = sTemp & myAlign("Net Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				'sTemp = sTemp & myAlign("ZZZZZZZZZZZZZ",aiHeaderColWidth(3,6),"R") & sLineBreakStr
				sTemp = sTemp & myAlign(FormatNumber(iFinNetPaid,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sNetFlag = True
			End IF
		End IF 'IF iCount > 0 then
		
		sText = sText & sLineBreakStr
		iNoOfLinesCtr =iNoOfLinesCtr + 1
		
					if not 	objRs.EOF then
						'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
						'sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",aiHeaderColWidth(3,3),"L")

						sTemp = Replace(sTemp,"HHHHHHHHHHHHH",myAlign(" ADDITIONAL PAYMENTS :",22,"L"))
						Do While not objRs.EOF
							iHeadOfAcc = objRs(0)
							iEntryAmt = cdbl(objRs(1))
							'iTotRecovered = iTotRecovered + iEntryAmt
							iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
							sItemDesc = objRs(4)


							sDetFlag = True
							If iHeadOfAcc <> "" Then
								sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
							Else
								sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
							End If
							'Response.Write sQuery
							with objRsTemp
								.CursorLocation = 3
								.CursorType = 3
								.Source = sQuery
								.ActiveConnection = con
								.Open
							end with
							Set objRsTemp.ActiveConnection = Nothing
							sHeadOfAccName = ""
							if not 	objRsTemp.EOF then
								sHeadOfAccName = objRsTemp(0)
							end if
							objRsTemp.Close
							
							If trim(sItemDesc) <> "" then sHeadOfAccName = sItemDesc &"-"&sHeadOfAccName
							 
																		
							'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
							sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
							sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
							sTemp = sTemp & sLineBreakStr
							
							iNoOfLinesCtr = iNoOfLinesCtr + 1
							
							objRs.MoveNext
							iCtr = iCtr + 1
						Loop
					End IF
					objRs.Close
					If Trim(sBillType) = "C" then
					'To display Tax Entries from Acc_T_CreatedVoucherTaxDet (Tax table)
						sQuery = "select AccountHead,TaxEntryNo from Acc_T_CreatedVoucherTaxDet where createdTransno = "&iTransNo&"  and TaxAmount <> 0  order by 2 "
						'esponse.Write "<P>Tax="& sQuery &"<BR>"
		'				Response.end
						objRs.Open sQuery,con
						Do while not objRs.EOF
							iHeadOfAcc = objRs(0)
							sDetFlag = True
							with objRsTemp
								.CursorLocation = 3
								.CursorType = 3
								If iHeadOfAcc <> "" Then
									.Source = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
								Else
									.Source = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
								End If
								.ActiveConnection = con
								.Open
							end with
							Set objRsTemp.ActiveConnection = Nothing
							sHeadOfAccName = ""
							if not 	objRsTemp.EOF then
								sHeadOfAccName = objRsTemp(0)
							end if
							objRsTemp.Close

							sQuery = "Select Sum(TaxAmount) from Acc_T_CreatedVoucherTaxDet where CreatedTransNo="&iTransNo&" and  AccountHead = "&iHeadOfAcc&" "
							objRsTemp.Open sQuery,con
							If not objRsTemp.EOF then
								iEntryAmt = cdbl(objRsTemp(0))
								'iTotRecovered = iTotRecovered + iEntryAmt
								iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
							End IF
							objRsTemp.Close

							'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
							sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
							'sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
							sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R")
							sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
							sTemp = sTemp & sLineBreakStr
							iNoOfLinesCtr = iNoOfLinesCtr + 1
							
							objRs.MoveNext
						iCtr = iCtr + 1
						Loop
					objRs.Close
					End IF 'If Trim(sBillType) = "C" then
				'	sTemp = sTemp & sLineBreakStr
 
				sText = sText & sTemp
				'sText = sText & sLineBreakStr
		End IF 'If trim(sBillType) <> "C" then
	  
		sTemp = ""
		
		
End Function
%>
<%	'Additional Payments New (Prchase Bill type = "C")
'Newly added on Sep 23 2008 by S.Maheswari

Function GotoAddPaymentsNew(sBillType)
	If trim(sBillType) = "C" then
	'Response.Clear
	
		sAddNewPayFlag = True
		sQuery = "Select count(1) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iTransNo&" and Amount <> 0 and TransCrDrIndication='"&sTranEntryIndication&"'"
		'Response.Write sQuery
		' Response.End
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		if not 	objRs.EOF then
			iCount = objRs(0)
		end if
		objRs.Close
		'Response.Write iCount


		sQuery = "Select VoucherNarration from  Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' "
		objRs.Open sQuery,con
		if not objRs.EOF then
			sNarr = objRs(0)

		end if
		objRs.Close

		If Trim(sBillType) = "C" then
			sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode,isNull(ItemDescription,'') from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' and Amount <> 0"
		Else
			sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode,isNull(ItemDescription,'') from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' and VoucherEntryNumber <> "&iEntryNo&"  and Amount <> 0"
		End If
		'Response.Write "<B>Additional Payments :</b>"&sQuery
		'Response.End
		iCtr = 1
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		if 	objRs.EOF then
			iCount =0
		end if
		 '9 bytes

		sNarrCtr = 0
		'Response.Clear
		'Response.Write sNarr
		'Response.End
	IF sNarrFlag <> True then 
		IF sNarr <> "" then
			sNarrFlag = True

			j = len(sNarr) / 90
			'Response.Write  cint(j)
			'
			IF Len(UCase(sNarr)) > 90 then

				For i = 1 to Len(sNarr) step 90
					sum = i + 90

					IF sum > len(sNarr) then
						'Response.Write "1="& mid(sNarr,i,90) & "<br>"
						if iCount <= 0 then sTemp = sTemp & sLineBreakStr
						sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")

						sTemp = sTemp & myAlign(" " & UCase(mid(sNarr,i,50)),aiHeaderColWidth(3,7)+2,"L")'90 bytes
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L") '3 bytes
						'Total Rec Rs.

						IF iCount > 0 then
							IF iNoOfLinesCtr = 18 then
								'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
								sTemp = sTemp & myAlign(" ",3,"L")
								sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",21,"L")
								sTemp = sTemp & myAlign(" Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
								sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
								sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
								iNoOfLinesCtr = iNoOfLinesCtr + 1
								sTotFlag = True
							End IF
						End IF 'IF iCount > 0 then
					Else
						sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
						sTemp = sTemp & myAlign(" " & UCase(mid(sNarr,i,90)),aiHeaderColWidth(3,7)+2,"L")
						sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L") '3 bytes
						'Total Rec Rs.
						IF iCount > 0 then
							IF iNoOfLinesCtr = 18 and sTotFlag <> True then
								'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
								sTemp = sTemp & myAlign(" ",3,"L")
								sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",21,"L")
								sTemp = sTemp & myAlign(" Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
								sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
								sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
								iNoOfLinesCtr = iNoOfLinesCtr + 1
								sTotFlag = True
							End IF
						End IF 'IF iCount > 0 then
					End If
					sNarrCtr = sNarrCtr +1
				Next
				sNextLine = True
			Else
			 
		   	    sNarrCtr = sNarrCtr +1
				'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
				sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-2,"L")
				sTemp = sTemp & myAlign(" " & UCase(sNarr),aiHeaderColWidth(3,7)+2,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				IF iCount > 0 then
					'Total Amount Rs.
					'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0)+10,"L")
					sTemp = sTemp & myAlign(" ",3,"L")
					sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",90,"L")
					sTemp = sTemp & myAlign(" Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
					sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
					sTemp = sTemp & myAlign(FormatNumber(iHeadOfAccAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr


					sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
					sTemp = sTemp & myAlign("",aiHeaderColWidth(3,7),"L") '90 bytes
					sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0),"L") '3 bytes

					'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L") '9 bytes
					sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3)-3,"L")
					'Response.Write "sNarrCtr="&sNarrCtr&"<BR>"
					IF sNarrCtr = 1 then
						sTemp = sTemp & myAlign("",90,"L")
						sHeadFlag = True
					Else
						sTemp = sTemp & myAlign("",15,"L")
						sTemp = sTemp & myAlign(" ",3,"L")
						sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",90,"L")&sLineBreakStr 
						iNoOfLinesCtr = iNoOfLinesCtr + 1
					End If
					'Net Paid Rs.
					sTotFlag = True
					sRecFlag = True
					sNetFlag = True
					sNextLine = False
					iNoOfLinesCtr = iNoOfLinesCtr + 3
					'Response.Write sTemp &"<br>"
				End IF 'IF iCount > 0 then
			End IF
			'sNarrFlag = True
		Else
			sTemp = sTemp & myAlign("",15,"L")
			sTemp = sTemp & myAlign(" ",3,"L")
			sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",90,"L")&sLineBreakStr
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		 	'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,0),"L")
		 	sNarrCtr = sNarrCtr +1
		End IF 'IF sNarr <> "" then
	End	IF 'IF sNarrFlag <> True then 
		'End If 'IF sNarrFlag  <> True then
		'sText = sText & sTemp
		'IF sNarrFlag = True then

		IF iCount > 0 then
			IF iNoOfLinesCtr = 18 and sTotFlag <> True then
				sTemp = sTemp & myAlign("",10,"L")
				sTemp = sTemp & myAlign(" Total Amount Rs.",aiHeaderColWidth(3,6)+3,"L")
				sTemp = sTemp &	myAlign("",aiHeaderColWidth(3,0)-2,"L")
				sTemp = sTemp & myAlign(FormatNumber(iVouAmt,2,,,-2),aiHeaderColWidth(3,6),"R") & sLineBreakStr
				iNoOfLinesCtr = iNoOfLinesCtr + 1
				sTotFlag = True
			End IF
		End IF 'IF iCount > 0 then
		
		sText = sText & sLineBreakStr
		iNoOfLinesCtr = iNoOfLinesCtr + 1
		
				if not 	objRs.EOF then
					'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
					'sTemp = sTemp & myAlign("ADDITIONAL PAYMENTS :",aiHeaderColWidth(3,3),"L")

					'sTemp = Replace(sTemp,"HHHHHHHHHHHHH",myAlign("ADDITIONAL PAYMENTS :",21,"L"))
					Do While not objRs.EOF
						iHeadOfAcc = objRs(0)
						iEntryAmt = cdbl(objRs(1))
						'iTotRecovered = iTotRecovered + iEntryAmt
						iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
						sItemDesc = objRs(4)


						sDetFlag = True
						If iHeadOfAcc <> "" Then
							sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
						Else
							sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
						End If
						'Response.Write sQuery
						with objRsTemp
							.CursorLocation = 3
							.CursorType = 3
							.Source = sQuery
							.ActiveConnection = con
							.Open
						end with
						Set objRsTemp.ActiveConnection = Nothing
						sHeadOfAccName = ""
						if not 	objRsTemp.EOF then
							sHeadOfAccName = objRsTemp(0)
						end if
						objRsTemp.Close

						If trim(sItemDesc) <> "" then sHeadOfAccName = sItemDesc &"-"&sHeadOfAccName

						'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
						sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3),"L")
						sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
						'sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-3,"L")
						sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R")
						sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
						sTemp = sTemp & sLineBreakStr
						iNoOfLinesCtr = iNoOfLinesCtr + 1
						
						objRs.MoveNext
						iCtr = iCtr + 1
					Loop
				End IF
				objRs.Close
				If Trim(sBillType) = "C" then
				'To display Tax Entries from Acc_T_CreatedVoucherTaxDet (Tax table)
					sQuery = "select AccountHead,TaxEntryNo from Acc_T_CreatedVoucherTaxDet where createdTransno = "&iTransNo&"  and TaxAmount <> 0  order by 2 "
					'Response.Write "<P>Tax="& sQuery &"<BR>"
	'				Response.end
					objRs.Open sQuery,con
					Do while not objRs.EOF
						iHeadOfAcc = objRs(0)
						sDetFlag = True
						with objRsTemp
							.CursorLocation = 3
							.CursorType = 3
							If iHeadOfAcc <> "" Then
								.Source = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
							Else
								.Source = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
							End If
							.ActiveConnection = con
							.Open
						end with
						Set objRsTemp.ActiveConnection = Nothing
						sHeadOfAccName = ""
						if not 	objRsTemp.EOF then
							sHeadOfAccName = objRsTemp(0)
						end if
						objRsTemp.Close

						sQuery = "Select Sum(TaxAmount) from Acc_T_CreatedVoucherTaxDet where CreatedTransNo="&iTransNo&" and  AccountHead = "&iHeadOfAcc&" "
						objRsTemp.Open sQuery,con
						If not objRsTemp.EOF then
							iEntryAmt = cdbl(objRsTemp(0))
							'iTotRecovered = iTotRecovered + iEntryAmt
							iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
						End IF
						objRsTemp.Close

						'sTemp = sTemp & myAlign("",aiHeaderColWidth(3,4),"L")
						sTemp = sTemp & myAlign("",aiHeaderColWidth(4,3),"L")
						sTemp = sTemp & myAlign(sHeadOfAccName,aiHeaderColWidth(5,1)-1,"L")
						sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
						sTemp = sTemp & myAlign(FormatNumber(iEntryAmt,2,,,-2),aiHeaderColWidth(5,4),"R")
						sTemp = sTemp & myAlign("",aiHeaderColWidth(5,0)-2,"L")
						
						sTemp = sTemp & sLineBreakStr
						iNoOfLinesCtr = iNoOfLinesCtr + 1
						
					objRs.MoveNext
					iCtr = iCtr + 1
					Loop
				objRs.Close
				End IF 'If Trim(sBillType) = "C" then
				'sTemp = sTemp & sLineBreakStr

				IF sDetFlag = True then

				End IF 'IF sDetFlag = True then

				'sTemp = Replace(sTemp,"DDDDDDDDDDDDD",myAlign(FormatNumber(iTotRecovered,2,,,-2),aiHeaderColWidth(3,6),"R"))
				'sTemp = Replace(sTemp,"ZZZZZZZZZZZZZ",myAlign(FormatNumber(iTotAddnlAmt,2,,,-2),aiHeaderColWidth(3,6),"R"))


				sText = sText & sTemp
				'sText = sText & sLineBreakStr
	End IF 'If trim(sBillType) = "C" then
	sTemp = ""
End Function
%>
