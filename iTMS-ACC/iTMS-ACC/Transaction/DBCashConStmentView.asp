<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	DBCashConStmentView.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	S.Maheswari
	'Created On					:	21 Jan 2009
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<%
dim objRs,objRs1,objRs2,objRs3,sQuery,iPageNo,isNo,sUnitDesc
dim sOrgId,sBookCode,iBookNo,sFromDate,sToDate,sBookName
dim iTransNo,sVouNo,sVouDate,sAccDescription,sOrgName,sAccHeadDesc
dim iVocEntryNo,sAccUnitHead,sAccUnitPartyCode,iAccHeadParam
dim sVovNarration,dAmount,sTransCrDrId,sCurrentDate,dOpeningAmt,iAccHead
dim sOpenCrDr,sCloseCrDr,dCloseBal,dPayTotal,dRecTotal,sDisplayHead
dim bFlag,saTemp,sOptSel,iVouNoFrom,iVouNoTo,dGAmount,dLAmount,sVouType
dim sGetVal,sFinMonYear,sMonthDay,sVouFromDate,sVouToDate,iCurrMonth,sUnitDet

sBookCode="01"
iPageNo=1
isNo = 0


set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
set objRs3 =server.CreateObject("adodb.Recordset")

'----------- To Get The Values From the Selection Page ----------------

sGetVal=Request.QueryString("Value")
saTemp=split(sGetVal,"|")
sOrgId= saTemp(0)
sOrgName=saTemp(1)
sBookName=saTemp(2)
iBookNo= cstr(saTemp(3))
sFromDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
sToDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)

sOptSel=saTemp(4)
sVouType=saTemp(7)

'To Display Organizations Full Description

sQuery="Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID='" & sOrgId & "'"

with objRs
	.CursorLocation =3
	.CursorType=3
	.ActiveConnection =con
	.Source =sQuery
	.Open
End with
set objrs.ActiveConnection =nothing
If not objRs.EOF then
	sOrgName =objRs(0)
End if
objRs.Close

'--------------- Coding For The Option Voucher Number -----------------

If sOptSel ="VouNo" Then
	iVouNoFrom=saTemp(5)
	iVouNoTo=saTemp(6)
	sFromDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	sToDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	sDisplayHead="For Voucher Number " & iVouNoFrom & " To " & iVouNoTo

'--------------- Coding For The Option Voucher Date -------------------

Elseif sOptSel ="VouDate" Then
	sVouFromDate=saTemp(5)
	sVouToDate=saTemp(6)
	sFromDate=saTemp(5)
	sToDate=saTemp(6)
	'sDisplayHead="For Voucher Date " & sVouFromDate & " To " & sVouToDate
	sDisplayHead="For the Period " & sVouFromDate & " To " & sVouToDate

'--------------- Coding For The Option Voucher Amount -----------------

Elseif sOptSel="Amount" Then
	dGAmount=saTemp(5)
	dLAmount=saTemp(6)

	sFromDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	sToDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)

	sDisplayHead="For The Amount Greater Than " & dGAmount & " And Less Than " & dLAmount

'--------------- Coding For The Option Account Head -------------------

Elseif sOptSel="AccHead" Then
	iAccHeadParam =saTemp(5)
	sAccHeadDesc=saTemp(8)
	sFromDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	sToDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	sDisplayHead="For The Account Head " & sAccHeadDesc
End if


'--------- Query for to Select the Account Head -----------------------

sQuery="SELECT BOOKACCOUNTHEAD FROM ACC_R_APPLICABLEACCOUNTHEADS WHERE BOOKCODE='" & sBookCode & "' AND BOOKNUMBER='"& iBookNo & "' AND OUDEFINITIONID='" & sOrgId & "'"
With objRs3
	.CursorLocation =3
	.CursorType=3
	.ActiveConnection=con
	.Source=sQuery
	.Open
End with

set objRs3.ActiveConnection =nothing
If not objRs3.EOF then
 iAccHead =objRs3(0)
End if
objRs3.Close


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Cash - Day Book</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
</HEAD>
<Script language="vbs">
Function CloseWindow()
	window.close()
End Function
Function ShowVouch(iTransNo)
	'MsgBox itransno
	showModalDialog "CashVouchDisp.asp?TransNo="&iTransNo,"","dialogHeight:390px;dialogWidth:640px;center:Yes;help:No;resizable:No;status:No"
	Exit Function
End Function
Function PrintWindow()
	sPassStr = document.formname.hValue.value
	showModalDialog "PRNDBCashStmentView.asp?Value="&sPassStr,"A","dialogHeight:150px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
End Function
</script>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type="hidden" name="hValue" value="<%=sGetVal%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
<!--#include file="../../include/ReportHeader_Accounts.asp"-->
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td>
<TABLE BORDER="0" CELLSPACING=0 CELLPADDING=0>
<TR><TD class="FieldCell" valign="bottom" width="60%"><b>
<%=sOrgName%> </b></TD>
</TR>
<Tr>
<TD class="FieldCell" valign="bottom" width="60%">
Cash Statement For<%="  "&sBookName%> </td></Tr>
<TR><TD class="FieldCell" valign="bottom" width="60%">
<%=sDisplayHead%> </TD>
</TR>
</TABLE>

</td>
<td width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td align="right" valign="top">
<TABLE BORDER="0" CELLSPACING=0 CELLPADDING=0>
<TR>
<TD class="FieldCell" valign="bottom" width="0">Date
</TD>
<TD class="FieldCellSub" valign="bottom" width="0">
<span class="DataOnly"><%=FormatDate(Date)%> </span></TD>
</TR>
<TR>
<TD class="FieldCell" width="0">Page No
</TD>
<TD class="FieldCellSub" width="0">
<span class="DataOnly"><%=iPageNo%> </span></TD>
</TR>
</TABLE>

</td>
</tr>
</table>

								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                <div class="frmBody" id="frm2" style="width: 755; height:415;">
<TABLE BORDER="0" CELLSPACING=1 CELLPADDING=0 width=100% class="ExcelTable">
<tr>
<TD ROWSPAN=2 class="ExcelHeaderCell" align="right" ><P align="center">Particulars</TD>
<TD COLSPAN=3 class="ExcelHeaderCell" align="right" ><P align="center">Amount</TD>
</tr>
<tr>
<TD class="ExcelHeaderCell" align="center" width=80>Receipts</TD>
<TD class="ExcelHeaderCell" align="center" width=80>Payments</TD>
<TD class="ExcelHeaderCell" align="center" width=100>Opening /<br>Closing</TD>
</tr>
<%
dPayTotal=0
dRecTotal=0
bFlag=true
Dim iCrTransNo

'--- Selection of Transaction Entries Based On The Option Selected ----
'--------------- Recordset For Transaction Entries --------------------
sQuery = "Select TransactionNumber,VoucherNumber,convert(char,VoucherDate,103),TransactionType,CreatedTransNo from  Acc_T_VoucherHeader"&_
	" where BookCode='"&sBookCode & "'and BookNumber="& iBookNo&" and OUDefinitionID='"&sOrgId&"'"

If  sOptSel="VouDate" Then
	sQuery=sQuery + " and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
	"VoucherDate <= convert(datetime,'"& sVouToDate &"',103)"

Elseif sOptSel="VouNo" Then
	sQuery =sQuery + " and VoucherNumber>= '" & iVouNoFrom & "' and VoucherNumber<= '" & iVouNoTo &"'"

'Elseif sOptSel="AccHead" Then
'		sQuery =sQuery &" and AccountHead=" & iAccHeadParam

End if

'sQuery =sQuery+ " order by VoucherDate"
sQuery =sQuery+ " Order By VoucherDate,TransactionType Desc "

 ' Response.Write sQuery
	with objRs
		.CursorLocation =3
		.CursorType =3
		.ActiveConnection =con
		.Source =sQuery
		.Open
	End With
	objRs.ActiveConnection =nothing
	Set iTransNo=objRs(0)
	Set sVouNo=objRs(1)
	Set sVouDate=objRs(2)
	Set iCrTransNo = objRs(4)

'---- Setting Flag For Checking Current And Previous voucher Dates ----
'------------- Displaying All the Transaction Details -----------------
If objRs.EOF then
	sCurrentDate=sFromDate
End if
'------------------ Function Call For Opening Amount ------------------
	dOpeningAmt=GetDayOpening(sOrgId,iAccHead,sFromDate)
	dOpeningAmt=FormatNumber(dOpeningAmt,2,,,0)
	dCloseBal=dOpeningAmt
	'	Response.Write dCloseBal & "<BR>"
	If objRs.RecordCount>0 Then
'----------Credit/Debit Indication For Opening and Closing ------------
		if dOpeningAmt <0 then
			sOpenCrDr="CR"
			sCloseCrDr="CR"
			dOpeningAmt=dOpeningAmt*-1
		else
			sOpenCrDr="DR"
			sCloseCrDr="DR"
		end if

'		Response.Write dCloseBal &"<BR><BR>"
%>
<tr>
<TD  class="ExcelDisplayCell" align="Right"><P><b>Opening as on <%=sFromDate%>&nbsp;</b></TD>
<TD class="ExcelDisplayCell" align="right">&nbsp;</td>
<TD class="ExcelDisplayCell" align="Right">&nbsp;</TD>
<P><TD class="ExcelDisplayCell" align="Right"><b><%=Abs(dCloseBal)&"&nbsp;"&sOpenCrDr%></b></TD>
</tr>
<% dim iAllTransNo

	while not objRs.EOF
		iAllTransNo = iAllTransNo &","& objRs(0)
		'Response.Write  objRs(0) &"<BR><BR>"
		if bFlag then
			iCurrMonth=Mid(sVouDate,4,2)
			sCurrentDate=sVouDate
			bFlag=false
		end if

	%>
		</TD>
		</tr>
	<% Dim	sRecFlag,sExpFlag
	'-------------- Query For Total Receipt Display -------------------
	sQuery = "select Amount,TransCrDrIndication,AccUnitAccountHead,AccUnitPartyCode from Acc_T_VoucherDetails where TransactionNumber ="&iTransNo

	If sOptSel ="Amount" Then
		sQuery=sQuery+ " and Amount between "& dGAmount & " and " & dLAmount
	Elseif sOptSel="AccHead" Then
		sQuery =sQuery &" and AccUnitAccountHead=" & iAccHeadParam
	End if
	If sVouType ="R" Then
		sQuery=sQuery& " and TransCrDrIndication='C'"
	Elseif sVouType="P" Then
		sQuery=sQuery&" and TransCrDrIndication='D'"
	End if
	sQuery = sQuery &" Order By VoucherEntryNumber "
	 ' Response.Write sQuery
	objRs1.Open sQuery,con

	while not objRs1.EOF
		dAmount					=objRs1(0)
		sTransCrDrId			=objRs1(1)

		if sTransCrDrId="C" then
			'dRecTotal=CDbl(dRecTotal)+CDbl(dAmount)
		end if
		objRs1.MoveNext
	wend
	objRs1.Close
	objRs.MoveNext
wend
IF sRecFlag <> True then
		%>
		<tr>
			<TD class="ExcelDisplayCell" align="Right"><P align="center"><B>Receipts</B></TD>
			<TD class="ExcelDisplayCell" align="right"><P><%'=FormatNumber(dRecTotal,2,,,0)%></TD>
			<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
			<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
		</tr>

		<%
			sRecFlag = True
		End IF
		iAllTransNo = mid(iAllTransNo,2)
		'only for Receipts
		'dim iNewTransno
		dAmount = 0

		'sQuery = "select CreatedTransNo from Acc_T_CreatedVoucherHeader where TransactionType = 'CAR' and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and VoucherDate <= convert(datetime,'"& sVouToDate & "',103)"
		'objRs1.Open sQuery,con
		'while not objRs1.EOF
		'	iNewTransno = iNewTransno &","&objRs1(0)
		'	objRs1.MoveNext
		'wend
		'objRs1.Close
		'iNewTransno = mid(iNewTransno,2)
		'dAmount = 0
		'sQuery = "Select AccUnitAccountHead,isnull(AccUnitPartyCode,0) from Acc_T_VoucherDetails where "&_
		'		 "TransactionNumber in ("&iNewTransno&") and TransCrDrIndication <> 'C' Group by AccUnitAccountHead,isnull(AccUnitPartyCode,0) "
 		 'Response.Write sQuery & vbCrLf
		'Response.End


		sQuery ="Select AccUnitAccountHead,isnull(AccUnitPartyCode,0) from Acc_T_VoucherDetails where transactionnumber in "&_
				"(Select Transactionnumber from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader "&_
				" where TransactionType = 'CAR' and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
				" VoucherDate <= convert(datetime,'"& sVouToDate & "',103)) ) and TransCrDrIndication = 'C' "&_
				" Group by AccUnitAccountHead,isnull(AccUnitPartyCode,0)"
		objRs1.Open sQuery,con
		while not objRs1.EOF

			sAccUnitHead			=objRs1(0)
			sAccUnitPartyCode		=objRs1(1)
			if 	IsNull(sAccUnitHead) or IsEmpty(sAccUnitHead) then
				sQuery = "select PartyName from APP_M_PartyMaster where " &_
						" PartyCode ="&sAccUnitPartyCode&" order by PartyName"
				objRs2.Open sQuery,con
				if not objRs2.EOF then
					sAccDescription = objRs2(0)
				end if
				objRs2.Close
				sQuery = "Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber from "&_
						 "Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader where "&_
						 "TransactionType = 'CAR' and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
						 "VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) and TransCrDrIndication = 'C'"  'and AccUnitPartyCode = "&sAccUnitPartyCode&"
				'Response.Write sQuery
				objRs2.Open sQuery,con
				if not objRs2.EOF then
					dAmount = objRs2(0)

				end if
				objRs2.close
			else
				sQuery = "select AccountDescription from Acc_M_GLAccountHead where " &_
					   " AccountHead ="&sAccUnitHead&" order by AccountDescription "
				objRs2.Open sQuery,con
				if not objRs2.EOF then
					sAccDescription = objRs2(0)
				end if
				objRs2.Close
				sQuery = "Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber from "&_
						 "Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader where "&_
						 "TransactionType = 'CAR' and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
						 "VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) and AccUnitAccountHead ="&sAccUnitHead&" and TransCrDrIndication = 'C'" ' and AccUnitAccountHead = "&sAccUnitHead&"
				  'Response.Write sQuery
				objRs2.Open sQuery,con
				if not objRs2.EOF then
					dAmount = objRs2(0)

				end if
				objRs2.close
			end if
			dRecTotal=CDbl(dRecTotal)+CDbl(dAmount)
			'Response.Write sQuery
			'Response.Write dAmount
			IF dAmount <> "" then%>
			<tr>
				<TD class="ExcelDisplayCell" align="Left"><P><%=sAccDescription%></TD>
				<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(dAmount,2,,,0)%></TD>
				<TD class="ExcelDisplayCell" align="right"><P></TD>
				<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
			</tr>
		<%End IF

		objRs1.MoveNext
		wend
		objRs1.close
	'Response.End
	'Response.Write iAllTransNo
	 '-------------- Query For Account Head Wise Display -------------------
	'Expenses
	'only for payments
	'sQuery = "select CreatedTransNo from Acc_T_CreatedVoucherHeader where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and VoucherDate <= convert(datetime,'"& sVouToDate & "',103) "
	'	objRs1.Open sQuery,con
	'	while not objRs1.EOF
	'		iNewTransno = iNewTransno &","&objRs1(0)
	'		objRs1.MoveNext
	'	wend
	'	objRs1.Close
	'	iNewTransno = mid(iNewTransno,2)

	'sQuery = "Select AccUnitAccountHead,isnull(AccUnitPartyCode,0) from Acc_T_VoucherDetails where "&_
	'		 "TransactionNumber in ("&iAllTransNo&") and TransCrDrIndication = 'D' Group by AccUnitAccountHead,isnull(AccUnitPartyCode,0) "
	sQuery ="Select AccUnitAccountHead,isnull(AccUnitPartyCode,0) from Acc_T_VoucherDetails where transactionnumber in "&_
				"(Select Transactionnumber from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from "&_
				" Acc_T_CreatedVoucherHeader where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) "&_
				" and VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) and TransCrDrIndication = 'D' "&_
				" Group by AccUnitAccountHead,isnull(AccUnitPartyCode,0)"

	objRs1.Open sQuery,con
	'Response.Write sQuery
	IF sExpFlag <> true then %>
		<tr>
			<TD class="ExcelDisplayCell" align="Right"><P align="center"><B>Payments</b></TD>
			<TD class="ExcelDisplayCell" align="right"><P></TD>
			<TD class="ExcelDisplayCell" align="right"><P></TD>
			<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
		</tr>

	<%sExpFlag = true
	End If

	while not objRs1.EOF

		sAccUnitHead			=objRs1(0)
		sAccUnitPartyCode		=objRs1(1)

		if 	IsNull(sAccUnitHead) or IsEmpty(sAccUnitHead) then
			sQuery = "select PartyName from APP_M_PartyMaster where " &_
					" PartyCode ="&sAccUnitPartyCode&" "
			objRs2.Open sQuery,con


			if not objRs2.EOF then
				sAccDescription = objRs2(0)
			end if
			objRs2.Close
			sQuery = " Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber "&_
					 " from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader "&_
					 " where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
					 " VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) "&_
					 " and AccUnitPartyCode = "&sAccUnitPartyCode&"  and TransCrDrIndication = 'D'"
			'Response.Write sQuery
			objRs2.Open sQuery,con
			if not objRs2.EOF then
				dAmount = objRs2(0)
			end if
			objRs2.close
		else
			sQuery = "select AccountDescription from Acc_M_GLAccountHead where " &_
				   " AccountHead ="&sAccUnitHead&" "
			objRs2.Open sQuery,con
			if not objRs2.EOF then
				sAccDescription = objRs2(0)
			end if
			objRs2.Close
			sQuery = " Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber "&_
					 " from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader "&_
					 " where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
					 " VoucherDate <= convert(datetime,'"& sVouToDate & "',103))) "&_
					 " and AccUnitAccountHead = "&sAccUnitHead&"  and TransCrDrIndication = 'D'"
				' Response.Write sQuery
			objRs2.Open sQuery,con
			if not objRs2.EOF then
				dAmount = objRs2(0)
			end if
			objRs2.close
		end if

		'if sTransCrDrId="D" then
			dPayTotal=CDbl(dPayTotal)+CDbl(dAmount)
		'end if
		'Response.Write dAmount
		%>
		<tr>
			<TD class="ExcelDisplayCell" align="Left"><P><%=sAccDescription%></TD>
			<TD class="ExcelDisplayCell" align="right"><P></TD>
			<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(dAmount,2,,,0)%></TD>
			<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
		</tr>
	<%
		objRs1.MoveNext
	wend
	objRs1.Close


'-------------- Updating Closing Balance For Each Entries -------------
'If not isnull(dRecTotal)then' or CDbl(dPayTotal)>0 then

 'Response.Write dCloseBal &"+"& dRecTotal &"-"&dPayTotal
	dCloseBal=CDbl(dCloseBal)+ CDbl(dRecTotal)-CDbl(dPayTotal) 'checking
%>
	<tr>
	<TD  class="ExcelDisplayCell" align="Right"><P><b>Closing for <%=sToDate%>&nbsp;<b></TD>
	<TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dRecTotal,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dPayTotal,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><P><b>
	<%
		if dCloseBal<0 then
			sCloseCrDr="CR"
		    Response.Write FormatNumber(dCloseBal*-1,2,,,0) &"&nbsp;"&sCloseCrDr
		else
			sCloseCrDr="DR"
			Response.Write FormatNumber(dCloseBal,2,,,0) &"&nbsp;"&sCloseCrDr
		end if

	%>
	</b>
	</TD>
	</tr>
<%
Else
	dOpeningAmt=FormatNumber(dOpeningAmt,2,,,0)
	dCloseBal=dOpeningAmt

%>
<tr>
<TD  class="ExcelDisplayCell" align="Right"><P><b>Opening as on <%=sFromDate%>&nbsp;</b></TD>
<TD class="ExcelDisplayCell" align="right">&nbsp;</td>
<TD class="ExcelDisplayCell" align="Right">&nbsp;</TD>
<P><TD class="ExcelDisplayCell" align="Right"><b><%=dCloseBal&"&nbsp;"&sOpenCrDr%></b></TD>
</tr>
<tr>
<TD  class="ExcelDisplayCell" align="Right"><P><b>Closing for <%=sToDate%>&nbsp;</b></TD>
<TD class="ExcelDisplayCell" align="right">&nbsp;</td>
<TD class="ExcelDisplayCell" align="Right">&nbsp;</TD>
<P><TD class="ExcelDisplayCell" align="Right"><b><%=dCloseBal&"&nbsp;"&sOpenCrDr%></b></TD>
</tr>
<%
End if
%>
</TABLE>
                                </div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center"> <input type="button" value="Ok" OnClick="CloseWindow()" class="ActionButton"  >
                                                 <input type="button" value="Print" OnClick="PrintWindow()" class="ActionButton" >
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>