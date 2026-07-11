<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CashBalanceStmt.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Sre Hari M
	'Created On					:	Feb 15, 2006
	'Modified By                :   Ragavendran R
	'Modified On                :   Jan 18,2011
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim sFinPeriod,Objrs,Objrs1,Objrs2,iCnt,sSql,iCrTransNo
	Dim sFormVal,sTemparr,sUnitID,sBookNo,sFrmDate,sToDate
	Dim iBookIndx,iAccIndx,sFlag,iAccHead,sAccHeadName
	Dim sCurrDate,iVouAmt,sFinTemp,sBookName

	dim iTransNo,sVouNo,sVouDate,sAccDescription
	dim iVocEntryNo,sAccUnitHead,sAccUnitPartyCode,iAccHeadParam
	dim sVovNarration,dAmount,sTransCrDrId,sCurrentDate,dOpeningAmt
	dim sOpenCrDr,sCloseCrDr,dCloseBal,dPayTotal,dRecTotal,sDisplayHead
	dim bFlag,sOptSel,iVouNoFrom,iVouNoTo,dGAmount,dLAmount,sVouType
	dim sGetVal,sFinMonYear,sMonthDay,sVouFromDate,sVouToDate,iCurrMonth
	Dim sBookCode,iBookNo
	sBookCode="01"


	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")
	set Objrs2=server.CreateObject ("ADODB.recordset")

	sFinPeriod=session("finperiod")
	sUnitID = Session("organizationcode")
	Response.Write "<font color=red>"

	'sVouFromDate ="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	'sVouToDate ="31/03/2012"

	sFinTemp = Split(sFinPeriod,":")

	''modified by ragav on 16/08/12
	sVouFromDate = "31/03/"&sFinTemp(0)
	sVouToDate = "01/04/"&sFinTemp(1)

	'''''''


	iBookNo = Request("selBook")
	sFrmDate = Request("hFromDate")
	sToDate =  Request("hToDate")
	iAccHead = Request("hAccHead")

	sFormVal = Request("hFormVal")
	sTemparr = Split(sFormVal,"|")

	'Response.Write "<p>sTemparr="&sFormVal
	'sTemparr=010101|1|1|01/04/2010|31/03/2011|1|14|Andhra Bank C/A
	'Response.Write "UBOUND = "&  UBound(sTemparr)

	If UBound(sTemparr) > 2 Then
		iBookNo = sTemparr(1)
		iBookIndx = sTemparr(2)
		sFrmDate = sTemparr(3)
		sToDate = sTemparr(4)
		iAccIndx = sTemparr(5)
		iAccHeadParam = sTemparr(6)
		sAccHeadName = sTemparr(7)
	End IF

	If sFrmDate = "" Then
		sFrmDate = sVouFromDate
		sToDate = sVouToDate
	End IF
	'Response.Write "<p><font color=red>iBookNo"&iBookNo
	If iBookNo <> "" Then
		sSql="select BookNumber,Upper(BookName),isnull(BookAccountHead,0) from "&_
				"vwOrgBookNames where OUDefinitionID = '" & sUnitID & "' and BookCode='01' and BookAccountHead is not null "

				if Trim(iBookNo)<>"S" then
					sSql = sSql & " and BookNumber = "& iBookNo
				end if
			sSql = sSql & " Order By BookName "

		'Response.Write sSql
		objRs.Open sSql,con

		IF not objRs.EOF then
			iBookNo = objRs(0)
			sBookName = objRs(1)
			'iBookAccHead = objRs(2)
		End If
		objRs.close
	Else
		'Now single book only ther.so only i assaighned value directly.
		iBookNo = "1"
		sBookName = "CASH BOOK"
		'iBookNo = "ALL"
	End IF

	IF Cstr(sUnitID) = "" Then
		sSql = "Select Top 1 OUDefinitionID From VwUserUnitList Where ApplicationCode = 1 and InternalUserID = "&getUserID()&" Order By OUDefinitionID "
		Objrs.Open sSql,Con
		IF Not Objrs.Eof Then
			sUnitID = Objrs(0)
		Else
			sUnitID = ""
		End IF
		objrs.close
	End IF

		sSql="SELECT Distinct BOOKACCOUNTHEAD FROM ACC_R_APPLICABLEACCOUNTHEADS WHERE BOOKCODE='" & sBookCode & "'  AND OUDEFINITIONID='" & sUnitID & "'"
		If iBookNo <> "ALL" THEN
			ssql = ssql & "AND BOOKNUMBER='"& iBookNo & "'"
		End IF
		'Response.Write "<p>"& sSql
		With objRs
			.CursorLocation =3
			.CursorType=3
			.ActiveConnection=con
			.Source=sSql
			.Open
		End with

		set objRs.ActiveConnection =nothing
		If not objRs.EOF then
			 iAccHead =objRs(0)
		End if
		objRs.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="OutData"><PartyType/></XML>
<XML ID="PartyData"><Party/></XML>
<XML id="AccHeadData">
<account/>
</XML>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/CashBalanceStmtCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate()">
<%
	Const iPageSize=16
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	'iCnt=Request.Form("hCnt")
%>
	<form method="POST" name="formname" action="CashBalanceStmt.asp" >

	<input type=hidden name="hUnitNo" value="<% =sUnitID%>">
	<input type=hidden name="hUnitName" value="<% =session("orgShortName")%>">
	<input type=hidden name="hBookNo" value="<%=iBookIndx%>">
	<input type=hidden name="hBookVal" value="<%=iBookNo%>">
	<input type=hidden name="hAccHead" value="<%=iAccHead%>">
	<input type=hidden name="hAccIndex" value="<%=iAccIndx%>">
	<input type=hidden name="hAccTxt" value="<%=sAccHeadName%>">
	<input type=hidden name="hFromDate" value="<%=sFrmDate%>">
	<input type=hidden name="hToDate" value="<%=sToDate%>">
	<input type=hidden name="hFlag" value="<%=sFlag%>">
	<input type=hidden name="hAccHeadP" value="<%=iAccHeadParam%>">

	<input type=hidden name="hVouName" value="CA">
	<input type=hidden name="hFinPeriod" value="<%=sFinPeriod%>">
	<input type=hidden name="hFormVal" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Balances
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%">
<div id="idUnprocessed" style="display: none">
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>


<tr>
<td class="FieldCellSub">
</td>
<td class="FieldCellSub">Cash Book
</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selBook" class="FormElem" onchange="GetBookNo()">
	<option value="S">Select Book</option>
</select>
</td>
</tr>


<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Voucher Date</td>

	<%'Response.Write InsertDatePicker("ctlVouFromDate") %>

    <td class="FieldCellSub" valign="middle">
		<input type="text" id="ctlVouFromDate" name="ctlVouFromDate" class="formelem itms-date-picker" data-itms-datepicker="1" size="10">
	</td>


	 <td class="FieldCell"></td>
	<td class="FieldCellSub">To
	</td>

	<%'Response.Write InsertDatePicker("ctlVouToDate") %>

        <td class="FieldCellSub" valign="middle">
			<input type="text" id="ctlVouToDate" name="ctlVouToDate" class="formelem itms-date-picker" data-itms-datepicker="1" size="10">
		</td>

</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCellSub">Account Head</td>
	<td class="FieldCellSub" colspan="4">
		<select class="formelem" OnChange="SelectAccHead()" size="1" name="selAccHead">
			<option value="0">Select Option</option>
			<option value="G">General Ledger</option>
		</select>

		<a href="Javascript:ResetAccHead()"><img border="0" width="11" height="11" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Remove Account Head" ></a>
	</td>
</tr>
<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub"></td>
	<td colspan="4" class="FieldCellSub">
		<input type="text" name="txtAccHead" size="70" Readonly class="FormElemRead">
	</td>
</tr>

<tr>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell" >
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
</td>
<td class="FieldCell" >
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
<table cellspacing="1" class="ExcelTable" width="100%" >
<tr>
<TD ROWSPAN=2 class="ExcelHeaderCell">Particulars</TD>
<TD COLSPAN=3 class="ExcelHeaderCell">Amount</TD>
</tr>
<tr>
<TD class="ExcelHeaderCell" width=80>Receipts</TD>
<TD class="ExcelHeaderCell" width=80>Payments</TD>
<TD class="ExcelHeaderCell" width=100>Opening /<br>Closing</TD>
</tr>

<%
	Response.Write "<font color=red>"
	iCnt=0

	sSql = " Select TransactionNumber,VoucherNumber,convert(char,VoucherDate,103),TransactionType,CreatedTransNo from  Acc_T_VoucherHeader"&_
		   " where BookCode='"&sBookCode & "'  and OUDefinitionID='"&sUnitID &"'"
	If iBookNo <> "ALL" Then
		sSql = sSql & " and BookNumber="& iBookNo &" "
	End IF
	If sFrmDate <> "" and sToDate <> "" Then
		sSql = sSql & " and VoucherDate >= convert(datetime,'"& sFrmDate & "',103) and "&_
			"VoucherDate <= convert(datetime,'"& sToDate &"',103)"
	End IF

'	If iAccHeadParam <> "" Then
'		sSql = sSql & " and AccountHead IN ("& iAccHeadParam &")"
'	End IF

	if iAccHead<>"" then
		sSql = sSql & " and AccountHead in ("& iAccHead &")"
	end if

	sSql = sSql & " Order By VoucherDate,TransactionType Desc "

	'Response.Write "<p> "& sSql

	with Objrs
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source=sSql
		.Open
	end with

	set Objrs.ActiveConnection=Nothing

	Set iTransNo=objRs(0)
	Set sVouNo=objRs(1)
	Set sVouDate=objRs(2)
	Set iCrTransNo = objRs(4)

	'---- Setting Flag For Checking Current And Previous voucher Dates ----
	'------------- Displaying All the Transaction Details -----------------
	If objRs.EOF then
		sCurrentDate=sFrmDate
	End if
	'Response.Write "<p>data"&sUnitID & "=="& iAccHead &"=="& sVouFromDate
	if trim(sUnitID)<>"" and trim(iAccHead)<>"" and trim(sVouFromDate)<>"" then
	    dOpeningAmt=GetDayOpening(sUnitID,iAccHead,sVouFromDate)
	end if 'if trim(sUnitID)<>"" and trim(iAccHead)<>"" and trim(sVouFromDate)<>"" then
	dOpeningAmt=FormatNumber(dOpeningAmt,2,,,0)
	dCloseBal=dOpeningAmt
	'Response.Write dCloseBal & "<BR>"
	If objRs.RecordCount>0 Then

		'******* Start of Paging
		Objrs.PageSize=iPageSize
		if iCurrentPage=0 then iCurrentPage=1
		Objrs.AbsolutePage=iCurrentPage
		iTotalPage=objrs.PageCount

	'----------Credit/Debit Indication For Opening and Closing ------------
		if dOpeningAmt <0 then
			sOpenCrDr="CR"
			sCloseCrDr="CR"
			dOpeningAmt=dOpeningAmt*-1
		else
			sOpenCrDr="DR"
			sCloseCrDr="DR"
		end if

		%>
		<tr>
		  <td colspan="4" class="ExcelDisplaycell" align="center"><B><%=sBookName%></B></td>
        </tr>

		<tr>
			<TD  class="ExcelDisplayCell" align="Right"><b>Opening as on <%=sFrmDate%>&nbsp;</b></TD>
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
		<!--</TD>
		</tr>-->
	<% Dim	sRecFlag,sExpFlag
	'-------------- Query For Total Receipt Display -------------------
	ssql = "select Amount,TransCrDrIndication,AccUnitAccountHead,AccUnitPartyCode from Acc_T_VoucherDetails where TransactionNumber ="&iTransNo

	If iAccHeadParam <> "" Then
		ssql =ssql &" and AccUnitAccountHead In (" & iAccHeadParam &" ) "
	End IF

	If sVouType ="R" Then
		ssql=ssql& " and TransCrDrIndication='C'"
	Elseif sVouType="P" Then
		ssql=ssql&" and TransCrDrIndication='D'"
	End if
	ssql = ssql &" Order By VoucherEntryNumber "

	'Response.Write "<p>sql="&sSql
	objRs1.Open ssql,con

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
			<TD class="ExcelDisplayCell" align="Right"><B>Receipts</B></TD>
			<TD class="ExcelDisplayCell" align="right"><%'=FormatNumber(dRecTotal,2,,,0)%></TD>
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


		ssql ="Select AccUnitAccountHead,isnull(AccUnitPartyCode,0) from Acc_T_VoucherDetails where transactionnumber in "&_
				"(Select Transactionnumber from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader "&_
				" where TransactionType = 'CAR' and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
				" VoucherDate <= convert(datetime,'"& sVouToDate & "',103)) ) and TransCrDrIndication = 'C' "

			if Trim(iAccHeadParam)<>"" then
				sSql = sSql & "  and AccUnitAccountHead in ("& iAccHeadParam &")"
			end if 'if Trim(iAccHeadParam)<>"" then

				sSql = sSql & "  Group by AccUnitAccountHead,isnull(AccUnitPartyCode,0)"

				'Response.Write "<p>"& sSql
						objRs1.Open ssql,con

		while not objRs1.EOF

			sAccUnitHead			=objRs1(0)
			sAccUnitPartyCode		=objRs1(1)

			if 	IsNull(sAccUnitHead) or IsEmpty(sAccUnitHead) then
				ssql = "select PartyName from APP_M_PartyMaster where " &_
						" PartyCode ="&sAccUnitPartyCode&" order by PartyName"
				objRs2.Open ssql,con
				if not objRs2.EOF then
					sAccDescription = objRs2(0)
				end if
				objRs2.Close
				ssql = "Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber from "&_
						 "Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader where "&_
						 "TransactionType = 'CAR' and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
						 "VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) and TransCrDrIndication = 'C'"  'and AccUnitPartyCode = "&sAccUnitPartyCode&"
				'Response.Write ssql
				objRs2.Open ssql,con
				if not objRs2.EOF then
					dAmount = objRs2(0)

				end if
				objRs2.close
			else
				ssql = "select AccountDescription from Acc_M_GLAccountHead where " &_
					   " AccountHead ="&sAccUnitHead&" order by AccountDescription "
				objRs2.Open ssql,con
				if not objRs2.EOF then
					sAccDescription = objRs2(0)
				end if
				objRs2.Close
				ssql = "Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber from "&_
						 "Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader where "&_
						 "TransactionType = 'CAR' and VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
						 "VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) and AccUnitAccountHead ="&sAccUnitHead&" and TransCrDrIndication = 'C'" ' and AccUnitAccountHead = "&sAccUnitHead&"
				  'Response.Write ssql
				objRs2.Open ssql,con
				if not objRs2.EOF then
					dAmount = objRs2(0)

				end if
				objRs2.close
			end if
			dRecTotal=CDbl(dRecTotal)+CDbl(dAmount)
			'Response.Write ssql
			'Response.Write dAmount
			IF dAmount <> "" then%>
			<tr>
				<TD class="ExcelDisplayCell" align="Left"><%=sAccDescription%></TD>
				<TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmount,2,,,0)%></TD>
				<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
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
	'ssql = "select CreatedTransNo from Acc_T_CreatedVoucherHeader where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and VoucherDate <= convert(datetime,'"& sVouToDate & "',103) "
	'	objRs1.Open ssql,con
	'	while not objRs1.EOF
	'		iNewTransno = iNewTransno &","&objRs1(0)
	'		objRs1.MoveNext
	'	wend
	'	objRs1.Close
	'	iNewTransno = mid(iNewTransno,2)

	'ssql = "Select AccUnitAccountHead,isnull(AccUnitPartyCode,0) from Acc_T_VoucherDetails where "&_
	'		 "TransactionNumber in ("&iAllTransNo&") and TransCrDrIndication = 'D' Group by AccUnitAccountHead,isnull(AccUnitPartyCode,0) "
	ssql ="Select AccUnitAccountHead,isnull(AccUnitPartyCode,0) from Acc_T_VoucherDetails where transactionnumber in "&_
				"(Select Transactionnumber from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from "&_
				" Acc_T_CreatedVoucherHeader where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) "&_
				" and VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) and TransCrDrIndication = 'D' "

		if Trim(iAccHeadParam)<>"" then
			sSql = sSql & "  and AccUnitAccountHead in ("& iAccHeadParam &")"
		end if 'if Trim(iAccHeadParam)<>"" then

		sSql = sSql &" Group by AccUnitAccountHead,isnull(AccUnitPartyCode,0)"

		'Response.Write "<textarea>"& sSql & "</textarea>"

	objRs1.Open ssql,con
	'Response.Write ssql
	IF sExpFlag <> true then %>
		<tr>
			<TD class="ExcelDisplayCell" align="right"><B>Payments</b></TD>
			<TD class="ExcelDisplayCell" align="right"></TD>
			<TD class="ExcelDisplayCell" align="right"></TD>
			<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
		</tr>

	<%sExpFlag = true
	End If

	while not objRs1.EOF

		sAccUnitHead			=objRs1(0)
		sAccUnitPartyCode		=objRs1(1)

		if 	IsNull(sAccUnitHead) or IsEmpty(sAccUnitHead) then
			ssql = "select PartyName from APP_M_PartyMaster where " &_
					" PartyCode ="&sAccUnitPartyCode&" "
			objRs2.Open ssql,con


			if not objRs2.EOF then
				sAccDescription = objRs2(0)
			end if
			objRs2.Close
			ssql = " Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber "&_
					 " from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader "&_
					 " where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
					 " VoucherDate <= convert(datetime,'"& sVouToDate & "',103) )) "&_
					 " and AccUnitPartyCode = "&sAccUnitPartyCode&"  and TransCrDrIndication = 'D'"
			'Response.Write ssql
			objRs2.Open ssql,con
			if not objRs2.EOF then
				dAmount = objRs2(0)
			end if
			objRs2.close
		else
			ssql = "select AccountDescription from Acc_M_GLAccountHead where " &_
				   " AccountHead ="&sAccUnitHead&" "
			objRs2.Open ssql,con
			if not objRs2.EOF then
				sAccDescription = objRs2(0)
			end if
			objRs2.Close
			ssql = " Select sum(Amount) from Acc_T_VoucherDetails where TransactionNumber in (Select Transactionnumber "&_
					 " from Acc_T_VoucherHeader where CreatedTransNo in (select CreatedTransNo from Acc_T_CreatedVoucherHeader "&_
					 " where TransactionType = 'CAP' and  VoucherDate >= convert(datetime,'"& sVouFromDate & "',103) and "&_
					 " VoucherDate <= convert(datetime,'"& sVouToDate & "',103))) "&_
					 " and AccUnitAccountHead = "&sAccUnitHead&"  and TransCrDrIndication = 'D'"
				' Response.Write ssql
			objRs2.Open ssql,con
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
			<TD class="ExcelDisplayCell" align="Left"><%=sAccDescription%></TD>
			<TD class="ExcelDisplayCell" align="right"></TD>
			<TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmount,2,,,0)%></TD>
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
	<TD  class="ExcelDisplayCell" align="Right"><b>Closing for <%=sToDate%>&nbsp;<b></TD>
	<TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dRecTotal,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dPayTotal,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><b>
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
<TD class="ExcelDisplayCell" align="Right"><b>Opening as on <%=sFrmDate%>&nbsp;</b></TD>
<TD class="ExcelDisplayCell" align="right">&nbsp;</td>
<TD class="ExcelDisplayCell" align="Right">&nbsp;</TD>
<P><TD class="ExcelDisplayCell" align="Right"><b><%=dCloseBal&"&nbsp;"&sOpenCrDr%></b></TD>
</tr>
<tr>
<TD  class="ExcelDisplayCell" align="Right"><b>Closing for <%=sToDate%>&nbsp;</b></TD>
<TD class="ExcelDisplayCell" align="right">&nbsp;</td>
<TD class="ExcelDisplayCell" align="Right">&nbsp;</TD>
<P><TD class="ExcelDisplayCell" align="Right"><b><%=dCloseBal&"&nbsp;"&sOpenCrDr%></b></TD>
</tr>
<%
End if
%>
</table>
<!--/div-->
</td>
<td align="center" class="ClearPixel" width="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>

<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hCnt" value=<%=iCnt  %>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="PaginateAcc('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage - 1%>')" id=button4 name=button4>
<%		end if	%>
<SELECT class="FormElem" onChange="PaginateAcc(this(this.selectedIndex).value)" id=select1 name=select1>
<%
For lnPage = 1 To iTotalPage
If lnPage = iCurrentPage Then
%>
<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
<%		else	%>
<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
<%		end if
next
%>
</SELECT>
<%
if iCurrentPage = iTotalPage then
%>
<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

<%		else	%>
<input type="button" value=" >> " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage + 1%>')" id=button7 name=button7>
<input type="button" value=" >| " class="ActionButtonX" onclick="PaginateAcc('<%=iTotalPage%>')" id=button8 name=button8>
<%		end if
End If
%>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td>

	<tr>

		<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td class="ActionCell">
                         <input type="button" value="Print" OnClick="PrintWindow()" class="ActionButton"  id=button10 name=button10>
					</td>
				</tr>
			</table>
		</td>

    </tr>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>
<tr>
<td align="center" class="BottomPack" colspan="3">
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
</body>
</html>
