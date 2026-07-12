<%@ Language="VBScript" %>
<% option explicit %>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ACCCASHVOUCHERS.ASP
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Sre Hari M
	'Created On					:	Feb 15, 2006
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
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim sFinPeriod,Objrs,Objrs1,Objrs2,iCnt,sSql,iCrTransNo,sOptType
	Dim dcrs,sUnitLID,sUnitLName,sUnitSName,AccVoucherNo,sorgID,sBookID
	Dim sFormVal,sTemparr,sUnitID,sBookNo,sFrmDate,sToDate,sFrmAmt,sToAmt
	Dim sFrmNo,sToNo,iBookIndx,iAccIndx,sFlag,iAccHead,sAccHeadName,sSelVouTy
	Dim sCurrDate,sCurrDay,sCurrMon,sCurrYear,dChqFrmNo,dChqToNo,iAccTrNo


	sCurrDay = Day(Date)
	sCurrMon = Month(Date)
	sCurrYear = Year(Date)

	IF Trim(Len(sCurrDay)) = 1 Then
		sCurrDay = "0"&sCurrDay
	End IF

	IF Trim(Len(sCurrMon)) = 1 Then
		sCurrMon = "0"&sCurrMon
	End IF

	sCurrDate = sCurrDay&"/"&sCurrMon&"/"&sCurrYear


	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")



	sFinPeriod=session("finperiod")
	sOptType = Request("OptCriteria")
	sFormVal = Request("hFormVal")
	'Response.Write sOptType &"<br><br>"
	sTemparr = Split(sFormVal,"|")

	sUnitID = Request("hUnitID")
	sBookNo = Request("selBookNo")
	sFrmDate = Cstr(sFrmDate)
	sToDate =  Cstr(sToDate)
	sFrmAmt = Request("hAmtFrom")
	sToAmt = Request("hAmtTo")
	iBookIndx = Request("hBookNo")
	iAccIndx = Request("hAccIndex")
	sFlag = Cstr(sFlag)
	iAccHead = Request("hAccHead")
	sAccHeadName = Request("hAccTxt")
	sSelVouTy = Request("voutype")
	dChqFrmNo = Request("hChqFrom")
	dChqToNo = Request("hChqTo")


	'Response.write Request("voutype")

	IF CStr(sUnitID) = ""  and UBound(sTemparr)>2 Then
		sUnitID = sTemparr(0)
	End IF

	IF CStr(sBookNo) = "" and UBound(sTemparr)>2 Then
		sBookNo = sTemparr(2)
	End IF

	IF CStr(iBookIndx) = "" and UBound(sTemparr)>2 Then
		iBookIndx = sTemparr(3)
	End IF

	IF CStr(sFrmDate) = "" and UBound(sTemparr)>2 Then
		sFrmDate = sTemparr(4)
	End IF

	IF CStr(sToDate) = "" and UBound(sTemparr)>2 Then
		sToDate = sTemparr(5)
	End IF

	IF CStr(sFrmAmt) = "" and UBound(sTemparr)>2 Then
		sFrmAmt = sTemparr(6)
	End IF

	IF CStr(sToAmt) = "" and UBound(sTemparr)>2 Then
		sToAmt = sTemparr(7)
	End IF

	IF CStr(sFrmNo) = "" and UBound(sTemparr)>2 Then
		sFrmNo = sTemparr(8)
	End IF

	IF CStr(sToNo) = "" and UBound(sTemparr)>2 Then
		sToNo = sTemparr(9)
	End IF

	IF CStr(iAccIndx) = "" and UBound(sTemparr)>2 Then
		iAccIndx = sTemparr(10)
	End IF

	IF CStr(sFlag) = "" and UBound(sTemparr)>2 Then
		sFlag = sTemparr(11)
	End IF

	IF CStr(iAccHead) = "" and UBound(sTemparr)>2 Then
		iAccHead = sTemparr(12)
	End IF

	IF CStr(sAccHeadName) = "" and UBound(sTemparr)>2 Then
		sAccHeadName = sTemparr(13)
	End IF


	'Response.Write dChqFrmNo &" " & dChqToNo


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData"><Book/></script>
<script type="application/xml" data-itms-xml-island="1" ID="OutData"><PartyType/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT SRC="../../scripts/VouTransactions.js"></SCRIPT>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/ACCCashVouchersCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate()">
<%
	Const iPageSize=15
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

	iCurrentPage=CInt(Request.Form("hPageSelection"))

	'iCnt=Request.Form("hCnt")
%>
	<form method="POST" name="formname" action="ACCCASHVOUCHERS.ASP">
	<input type=hidden name="hTransNo" value="">
	<input type=hidden name="hUnitNo" value="<%=sUnitID%>">
	<input type=hidden name="hBookNo" value="<%=iBookIndx%>">
	<input type=hidden name="hBookVal" value="<%=sBookNo%>">
	<input type=hidden name="hAccHead" value="<%=iAccHead%>">
	<input type=hidden name="hAccIndex" value="<%=iAccIndx%>">
	<input type=hidden name="hAccTxt" value="<%=sAccHeadName%>">
	<input type=hidden name="hFromDate" value="<%=sFrmDate%>">
	<input type=hidden name="hToDate" value="<%=sToDate%>">
	<input type=hidden name="hFlag" value="<%=sFlag%>">
	<input type=hidden name="hAmtFrom" value="<%=sFrmAmt%>">
	<input type=hidden name="hAmtTo" value="<%=sToAmt%>">
	<input type=hidden name="hVouFrom" value="<%=sFrmNo%>">
	<input type=hidden name="hVouTo" value="<%=sToNo%>">
	<input type=hidden name="hVouName" value="CA">
	<input type=hidden name="hFinPeriod" value="<%=sFinPeriod%>">
	<input type=hidden name="hFormVal" value="">
	<input type=hidden name="hChqFrom" value="<%=dChqFrmNo%>">
	<input type=hidden name="hChqTo" value="<%=dChqToNo%>">
	<input type=hidden name="hUpdate" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Cash Vouchers
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

<!--tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="ToolBarTable">
<tr>
<td width="40" align="center" valign="middle" class="ToolBarCell" onclick="toolClick(this)" onmouseover="toolrollover(this)" onmouseout="toolrollout(this)">
<span style="cursor: pointer" title="New">
<p align="center"><font face="Wingdings" size="5">2</font></p>
</span>
</td>
<td align="center" class="ToolBarCell">&nbsp;
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" colspan="3" class="MiddlePack">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr-->

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
<td valign="center" class="SubTitle">&nbsp;&nbsp;
<%
	Dim aFlag,saTemp
	aFlag=false
		'Response.Write Cstr(sFlag)
	IF CStr(sSelVouTy) = "A" or CStr(sSelVouTy) = "" or InStr(1,sSelVouTy,CStr("C, P, T"))>0 Then
		if Trim(sUnitID)="" then
				sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
				& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD WHERE A.BOOKCODE='01' AND CREATEDVOUCHSTATUS = '010104' AND ISNULL(A.PURCHASEBILLTYPE,'') <> 'C' "
		else
				sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
				& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD WHERE A.BOOKCODE='01' AND A.OUDEFINITIONID='"&sUnitID &"' AND CREATEDVOUCHSTATUS = '010104' AND ISNULL(A.PURCHASEBILLTYPE,'') <> 'C' "
		end if

		IF Trim(sBookNo) <> "" Then
			sSql = sSql&" AND A.BOOKNUMBER="& Cstr(sBookNo) &"  "
		End IF




		'IF Cstr(sFlag) = "" and Cstr(sCurrDate) <> "" Then
		'	sSql = sSql &" AND CONVERT(DATETIME,A.VOUCHERDATE,103) = CONVERT(DATETIME,'"&sCurrDate&"',103) "
		'End IF

		aFlag=true
		'Response.Write ("<Input type=checkbox name=voutype value=A checked onclick=ChkVouType()>All&nbsp;")
		'Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType() >Created&nbsp;")
		'Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType() >Approved&nbsp;")
		'Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType() >Accounted&nbsp;")
		select case Cstr(sFlag)
		case "VouNo"
				sSql=sSql+" AND A.CREATEDVOUCHERNO BETWEEN '"&Cstr(sFrmNo) &"' AND '"& Cstr(sToNo)&"'"
		case "VouAmount"
				sSql=sSql+" AND A.VOUCHERAMOUNT BETWEEN "&sFrmAmt &" AND "&sToAmt &""
		case "AccHead"
				if Request("selacchead")="G" then
					sSql=sSql+" and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead="&iAccHead&") "
				else
					saTemp=Split(iAccHead,"?")
					sSql=sSql+" and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
					" AccUnitPartyType ='"&Trim(saTemp(0))&"' and AccUnitPartySubType="&Trim(saTemp(1))&" and AccUnitPartyCode="&Trim(saTemp(3))&") "
				end if
		case "Cheque"
				sSql=sSql+" and isNumeric(BankInstrumentNo) = 1 and Cast(BankInstrumentNo AS Numeric) Between "&dChqFrmNo&" and "&dChqToNo&" "
		end select

		if not Cstr(sFlag)= "VouDate" then
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103)"
		else
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& Cstr(sFrmDate) &"',103) " _
			& "AND CONVERT(DATETIME,'"& Cstr(sToDate) & "',103)"
		end if

		sSql  = sSql + " order by A.VoucherDate Desc,A.TransactionType,A.CreatedVoucherNo desc,A.CreatedTransNo desc"
	End IF

if not aFlag then
	if Trim(sUnitID)="" then
			sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
			& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD WHERE A.BOOKCODE='02' AND ISNULL(A.PURCHASEBILLTYPE,'') <> 'C' AND (A.CREATEDVOUCHSTATUS='0'"
	else
			sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
			& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD WHERE A.BOOKCODE='02' AND ISNULL(A.PURCHASEBILLTYPE,'') <> 'C' AND A.OUDEFINITIONID='"&sUnitID &"' AND A.BOOKNUMBER='"& Cstr(sBookNo) &"' AND(A.CREATEDVOUCHSTATUS='0'"
	end if

		Response.Write ("<Input type=checkbox name=voutype value=A onclick=ChkVouType()>All&nbsp;")
		if Instr(1,sSelVouTy,"C") > 0 then
			Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType() checked>Created&nbsp;")
			sSql=sSql+"OR A.CREATEDVOUCHSTATUS='010101' OR A.CREATEDVOUCHSTATUS='010102'"
		Else
			Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType()>Created&nbsp;")
		End IF

		if Instr(1,sSelVouTy,"P") > 0 then
			Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType() checked >Approved&nbsp;")
			sSql=sSql+"OR A.CREATEDVOUCHSTATUS='010103' OR A.CREATEDVOUCHSTATUS='010105'"
		Else
			Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType()>Approved&nbsp;")
		End IF

		if Instr(1,sSelVouTy,"T") > 0 Then
			Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType() checked >Accounted&nbsp;")
			sSql=sSql+"OR A.CREATEDVOUCHSTATUS='010104'"
		Else
			Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType()>Accounted&nbsp;")
		end if

		select case Cstr(sFlag)
		case "VouNo"
				sSql=sSql+")"+"AND A.CREATEDVOUCHERNO BETWEEN '"&Cstr(sFrmNo) &"' AND '"& Cstr(sToNo)&"'"
		case "VouAmount"
				sSql=sSql+")"+"AND A.VOUCHERAMOUNT BETWEEN "&sFrmAmt &" AND "&sToAmt &""
		case "AccHead"
				saTemp=Split(iAccHead,"?")
				'if Request("selacchead")="G" then
				IF UBound(saTemp) < 2 Then
					sSql=sSql+")"+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead="&iAccHead&") "
				else
					saTemp=Split(iAccHead,"?")
					sSql=sSql+")"+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
					" AccUnitPartyType ='"&Trim(saTemp(0))&"' and AccUnitPartySubType="&Trim(saTemp(1))&" and AccUnitPartyCode="&Trim(saTemp(3))&") "
				end if
		case ""
				sSql=sSql+")"
		case "VouDate"
				sSql=sSql+")"
		case "0"
				sSql=sSql+")"
		case "Cheque"
				sSql=sSql+")"+" AND ISNUMERIC(BANKINSTRUMENTNO) = 1 and Cast(BankInstrumentNo AS Numeric) BETWEEN "&dChqFrmNo &" AND "&dChqToNo &""
		end select

		if not Cstr(sFlag)= "VouDate" then
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103)"
		else
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& Cstr(sFrmDate) &"',103) " _
			& "AND CONVERT(DATETIME,'"& Cstr(sToDate) & "',103)"
		end if

		sSql  = sSql + " order by A.VoucherDate Desc,A.TransactionType,A.CreatedVoucherNo desc,A.CreatedTransNo desc"
end if
'	 Response.Write sSql
%>
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
<td width="100%">
<div id="idUnprocessed" style="width: 575; display: none">
<table cellpadding="0" cellspacing="0">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<tr>
	<td class="FieldCellSub">&nbsp;&nbsp;</td>
	<td class="FieldCellSub">Unit Name</td>
	<td class="FieldCellSub" colspan="4">
	<select size="1" name="selUnitId" class="FormElem" onchange="DisplayBook()">
		<option value="">Select</option>
		<%
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) ORDER BY ORGANIZATIONUNITID"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUnitLID = dcrs(0)
		set sUnitLName = dcrs(1)
		set sUnitSName = dcrs(3)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				if CStr(sUnitID)=sUnitLID then
		%>
		          <OPTION VALUE="<%=sUnitLID%>" selected><%=sUnitSName%></Option>
		          <%else%>
		          <OPTION VALUE="<%=sUnitLID%>" ><%=sUnitSName%></Option>
		<%
				end if
				dcrs.MoveNext
			Loop
		end if

		dcrs.Close
	%>
	</select>

	</td>
</tr>

<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Cash Book</td>
	<td class="FieldCellSub" colspan="4">
	<select size="1" name="selBook" class="FormElem" onchange="GetBookNo()">
		<option>Select</option>
	</select>
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%IF CStr(sOptType) = "VouNo" Then %>
		<input type="radio" value="VouNo" name="OptCriteria" onclick="Optselection()" Checked>Voucher No. From
	<%Else%>
		<input type="radio" value="VouNo" name="OptCriteria" onclick="Optselection()">Voucher No. From
	<%End IF %>
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtVouNoFrom" Readonly size="20" class="FormElem">
	</td>

	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">To
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtVouNoTo" Readonly size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%IF CStr(sOptType) = "VouDate" Then %>
		<input type="radio" value="VouDate" name="OptCriteria" onclick="OptSelection()" Checked>Voucher Date
	<%Else%>
		<input type="radio" value="VouDate" name="OptCriteria" onclick="OptSelection()" >Voucher Date
	<%End IF %>
	</td>
	<td class="FieldCellSub">
	<%Response.Write InsertDatePicker("ctlVouFromDate") %>
	</td>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">To	</td>
	<td class="FieldCellSub">
	<%Response.Write InsertDatePicker("ctlVouToDate") %>
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%If CStr(sOptType)="VouAmount" then%>
		<input type="radio" value="VouAmount" name="OptCriteria" onclick="OptSelection()" checked>Amount From
	<%else%>
		<input type="radio" value="VouAmount" name="OptCriteria" onclick="OptSelection()">Amount From
	<% end if%>
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtFromAmount" Readonly size="20" class="FormElem"></td>
	<td class="FieldCellSub"></td>

	<td class="FieldCellSub">To	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtToAmount" Readonly size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<% if CStr(sOptType)="AccHead" then %>
		<input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()" checked>Account Head
	<%else%>
		<input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()">Account Head
	<%end if%>
	</td>
	<td class="FieldCellSub" colspan="4">
	<% if CStr(sOptType)="AccHead" then %>
		<select class="formelem" OnChange="SelectAccHead()" size="1" name="selAccHead">
	<%Else%>
		<select class="formelem" disabled OnChange="SelectAccHead()" size="1" name="selAccHead">
	<%ENd IF %>
			<option value="0">Select Option</option>
			<option value="G">General Ledger</option>
		</select>
	</td>
</tr>
<tr>
	<td class="FieldCellSub" ></td>
	<td class="FieldCellSub" ></td>
	<td colspan="4" class="FieldCellSub">
		<input type="text" name="txtAccHead" Readonly size="70" class="FormElemRead">
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
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()" >
</td>
</table>
</div>
</td>
</tr>
<tr>
<td align="center" class="MiddlePack">
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
<table border="0" cellspacing="1" class="ExcelTable" width="100%">
<tr>
<td class="ExcelHeaderCell" align="center" width="10" >S.No.
</td>
<td class="ExcelHeaderCell" align="center" width="10" >
</td>
<td class="ExcelHeaderCell" align="center" >Number
</td>
<td class="ExcelHeaderCell" align="center" >Date
</td>
<td class="ExcelHeaderCell" align="center" >Type
</td>
<td class="ExcelHeaderCell" align="center" >A/c. Head / Party
</td>
<td class="ExcelHeaderCell" align="center" >Amount
</td>
<td class="ExcelHeaderCell" align="center" >Status
</td>
</tr>

<%

	'Response.Write sSql
	iCnt=0
	with Objrs
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source=sSql
		.Open
	end with
	set Objrs.ActiveConnection=Nothing
	IF  not Objrs.EOF then

	'******* Start of Paging
	Objrs.PageSize=iPageSize
	if iCurrentPage=0 then iCurrentPage=1
	Objrs.AbsolutePage=iCurrentPage
	iTotalPage=objrs.PageCount
	For iPageCtr=1 to objrs.PageSize
		iCnt=iCnt+1
		iCrTransNo=Objrs("createdtransno")
		sSql = "Select VoucherNumber,TransactionNumber From Acc_T_VoucherHeader Where CreatedTransno = "&iCrTransNo
		Objrs1.Open sSql,con
		IF Not Objrs1.EOF Then
			AccVoucherNo = Objrs1(0)
			iAccTrNo = Objrs1(1)
		Else
			AccVoucherNo = ""
			iAccTrNo = ""
		End IF
		Objrs1.Close

		Dim iParCode,AccParName

		sSql =  "Select Top 1 D.AccUnitPartyCode From  Acc_T_CreatedVoucherDetails D Where  "&_
				"D.AccUnitPartyCode <> 0 And CreatedTransNo = "&iCrTransNo
		Objrs1.Open sSql,Con
		IF Not Objrs1.EOF Then
			iParCode = Objrs1(0)
		Else
			iParCode = 0
		End IF
		Objrs1.Close

		IF CStr(iParCode) = "0" Then
			sSql =  "Select Top 1 H.AccountDescription From  Acc_T_CreatedVoucherDetails D, "&_
					"Acc_M_GLAccountHead H Where D.AccUnitAccountHead <> 0 And  "&_
					"D.CreatedTransNo = "&iCrTransNo&" And D.AccUnitAccountHead = H.AccountHead "

			Objrs1.Open sSql,Con
			IF Not Objrs1.EOF Then
				AccParName = Objrs1(0)
			Else
				AccParName = ""
			End IF
			Objrs1.Close
		Else
			sSql = "Select PartyName From App_M_PartyMaster Where PartyCode = "&iParCode
			Objrs1.Open sSql,Con
			IF Not Objrs1.EOF Then
				AccParName = Objrs1(0)
			Else
				AccParName = ""
			End IF
			Objrs1.Close
		End IF
%>
<tr>
	<td class="ExcelSerial" align="center" ><%=iCnt%></td>
	<td class="ExcelDisplayCell" align="center" width="10" >
	<input type="radio" name="Chkbox" text="<%=Objrs("createdvoucherno")&"&"& Right(Objrs("transactiontype"),1)&"@"& Right(CStr(Objrs("createdvouchstatus")),2)%>" value="<%=iCrTransNo %>" >
	<td class="ExcelDisplayCell" align="left" >
	<a href="#" onclick="return ShowVouch(<%=iCrTransNo %>)" class="ExcelDisplayLink" title="<%=Objrs("createdvoucherno")%>-<%=iCrTransNo %>"><%=Objrs("createdvoucherno") %> </a></td>
	<td class="ExcelDisplayCell" align="left" ><%=FormatDate(Objrs("voucherdate"))%></td>
	<td class="ExcelDisplayCell" align="left" ><%=Objrs("TRANSACTIONTYPE")%></td>
	<!--td class="ExcelDisplayCell" align="left" ><%=Objrs("accountdescription")%></td-->
	<td class="ExcelDisplayCell" align="left" ><%=AccParName%></td>
	<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(Objrs("Voucheramount")) %></td>
	<td class="ExcelDisplayCell" align="left" ><%=AccVoucherNo %></td>
</tr>
<%
	Objrs.MoveNext
	if Objrs.EOF then exit for
	next
	end if
	Objrs.Close
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
<td valign="top" align="right">
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hCnt" value=<%=iCnt%>>
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
<td valign="middle" class="ActionCell">
<p align="center">
<input type="button" value="Edit" name="B9" class="ActionButton" tabindex="3" onclick="ChkforEdit()">
<!--input type="button" value="Approve" name="B10" class="ActionButton" tabindex="3" onclick="ChkforApprove()">
<input type="button" value="Account" name="B11" class="ActionButton" tabindex="3" onclick="ChkforAccount()">
<input type="button" value="Delete" name="B12" class="ActionButton" tabindex="3" onclick="ChkforDelete()" -->
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
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
