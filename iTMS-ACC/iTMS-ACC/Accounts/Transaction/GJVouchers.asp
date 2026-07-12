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
	'Program Name				:	GJVouchers.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 29,2011
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
	Dim sFinPeriod,Objrs,Objrs1,Objrs2,iCnt,sSql,sQuery,iCrTransNo,sHead
	Dim dcrs,sUnitLID,sUnitLName,sUnitSName,sOptType,sAccHead,sAmount,AccVoucherNo
	Dim sFormVal,sTemparr,sUnitID,sBookNo,sFrmDate,sToDate,sFrmAmt,sToAmt
	Dim sFrmNo,sToNo,iBookIndx,iAccIndx,sFlag,iAccHead,sAccHeadName,sSelVouTy,sUserID
	Dim sField1,sField2,sField3,sField4,sField5,sSortBy,Arr1,nFieldSelected
	Dim sValTemp,sFinFromDate,sFinToDate

	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")

sUnitID = Session("organizationcode")
	sFinPeriod=session("finperiod")
	sOptType = Request("OptCriteria")
	sFormVal = Request("hFormVal")
'	Response.Write sFormVal
	sTemparr = Split(sFormVal,"|")
	'sUnitID = Request("hUnitID")
	sUserID = getUserId()
	sBookNo = Request("selBookNo")
	sFrmDate = Request("hFromDate")
	sToDate =  Request("hToDate")
	sFrmAmt = Request("hAmtFrom")
	sToAmt = Request("hAmtTo")
	iBookIndx = Request("hBookNo")
	iAccIndx = Request("hAccIndex")
	sFlag = Cstr(sFlag)
	iAccHead = Request("hAccHead")
	sAccHeadName = Request("hAccTxt")
	sSelVouTy = sSelVouTy


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

	IF UBound(sTemparr)>13 Then
		sUserID = sTemparr(14)
	End IF

	IF CStr(sUserID) = "" Then
		sUserID = getUserID()
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


sSelVouTy = Request.Form("voutype")
'Response.Write sSelVouTy

'sField1  = Request("hField1")
	'sField2  = Request("hField2")
	'sField3  = Request("hField3")
	'sField4  = Request("hField4")
	'sField5  = Request("hField5")

	'if trim(sField1) = "" then sField1 = "N:A"
	'if trim(sField2) = "" then sField2 = "D:A"
	'if trim(sField3) = "" then sField3 = "T:A"
	'if trim(sField4) = "" then sField4 = "A:A"
	'if trim(sField5) = "" then sField5 = "M:A"


	sField1  = ""
	sField2  = ""
	sField3  = ""
	sField4  = ""
	sField5  = ""

	nFieldSelected = trim(Request.Form("hFieldSelected"))
	if trim(nFieldSelected) = "" then nFieldSelected = 0

	if nFieldSelected = "1"  then
		sField1  = Request("hField1")
	end if

	if nFieldSelected = "2" then
		sField2  = Request("hField2")
	end if

	if nFieldSelected = "3" then
		sField3  = Request("hField3")
	end if

	if nFieldSelected = "4" then
		sField4  = Request("hField4")
	end if

	if nFieldSelected = "5" then
		sField5  = Request("hField5")
	end if

	'--------------------
	if nFieldSelected = "0" then
		sField1  = "N:A"
	end if


	sSortBy = ""

	if trim(sField1) <> ""  then
		if instr(1,sField1,":") > 0 then
			Arr1 = Split(sField1,":")

			if Arr1(1) = "A" then
				sSortBy = "H.CREATEDVOUCHERNO"
			else
				sSortBy = "H.CREATEDVOUCHERNO desc"
			end if
		end if
	end if


	if trim(sField2) <> ""  then
		if instr(1,sField2,":") > 0 then
			Arr1 = Split(sField2,":")

			if Arr1(1) = "A" then
				sSortBy = "H.VOUCHERDATE"
			else
				sSortBy = "H.VOUCHERDATE desc "
			end if
		end if
	end if

	'if trim(sField3) <> ""  then
	'	if instr(1,sField3,":") > 0 then
	'		Arr1 = Split(sField3,":")
	'
	'		if Arr1(1) = "A" then
	'			sSortBy = "H.TRANSACTIONTYPE"
	'		else
	'			sSortBy = "H.TRANSACTIONTYPE desc "
	'		end if
	'	end if
	'end if

	'blocked
	'if trim(sField4) <> "" and 1 = 2  then
	'	if instr(1,sField4,":") > 0 then
	'		Arr1 = Split(sField4,":")
	'
	'		if Arr1(1) = "A" then
	'			sSortBy = sSortBy  & ","
	'		else
	'			sSortBy = sSortBy  & ", desc "
	'		end if
	'	end if
	'end if

	if trim(sField5) <> ""  then
		if instr(1,sField5,":") > 0 then
			Arr1 = Split(sField5,":")

			if Arr1(1) = "A" then
				sSortBy = "Amount"
			else
				sSortBy = "Amount desc "
			end if
		end if
	end if

sFinPeriod = Session("FinPeriod")
sValTemp = Split(sFinPeriod,":")
sFinFromDate = "01/04/"&Trim(sValTemp(0))
sFinToDate= "31/03/"&Trim(sValTemp(1))

if Trim(sFrmDate)="" then
    if DateDiff("d",sToDate,date)>0 then
        sFrmDate = sFinFromDate
        sToDate = sFinToDate
    else
        sFrmDate = sFinFromDate
        sToDate = date
    end if
end if


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
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempXMLData"><Root></Root></script>
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<script SRC="../../scripts/printwindow.js"></SCRIPT>
<script SRC="../../scripts/VouTransactions.js"></SCRIPT>
<script SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script SRC="../../scripts/GJVouchersCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate();">
<%
	Const iPageSize=20
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	'iCnt=Request.Form("hCnt")
%>
	<form method="POST" name="formname" action="GJVouchers.asp">
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
	<input type=hidden name="hVouName" value="GJ">
	<input type=hidden name="hFinPeriod" value="<%=sFinPeriod%>">
	<input type=hidden name="hFormVal" value="">
	<input type=hidden name="hUserID" value="">

	<input type="hidden" name="hField1" value="<%=sField1%>">
	<input type="hidden" name="hField2" value="<%=sField2%>">
	<input type="hidden" name="hField3" value="<%=sField3%>">
	<input type="hidden" name="hField4" value="<%=sField4%>">
	<input type="hidden" name="hField5" value="<%=sField5%>">

	<input type="hidden" name="hFieldSelected" value="<%=nFieldSelected%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				General Journal
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
									<td align="center" colspan="3" class="MiddlePack" height="7px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
<%
	Dim aFlag,saTemp,sVal
	aFlag=false

	IF CStr(sSelVouTy) = "A" or CStr(sSelVouTy) = "" or InStr(1,sSelVouTy,CStr("C, P, T"))>0 Then
		if Trim(sUnitID)="" then
			sSql="Select H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,Sum(D.Amount)Amount,H.CreatedVoucherNo,H.CreatedVouchStatus ,Convert(Char,H.VoucherDate,103)VoucherDate From Acc_T_CreatedVoucherHeader " _
			& "H,Acc_T_CreatedVoucherDetails D Where D.CreatedTransNo = H.CreatedTransNo and H.BookCode = '08' "
			''& "H,Acc_T_CreatedVoucherDetails D Where D.CreatedTransNo = H.CreatedTransNo and H.BookCode = '08' and D.TransCrDrIndication='D' "
		else
			sSql="Select H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,Sum(D.Amount)Amount,H.CreatedVoucherNo,H.CreatedVouchStatus ,Convert(Char,H.VoucherDate,103)VoucherDate From Acc_T_CreatedVoucherHeader " _
			& "H,Acc_T_CreatedVoucherDetails D Where D.CreatedTransNo = H.CreatedTransNo and H.BookCode = '08' AND H.OUDEFINITIONID='"&sUnitID &"' "
			''& "H,Acc_T_CreatedVoucherDetails D Where D.CreatedTransNo = H.CreatedTransNo and H.BookCode = '08' and D.TransCrDrIndication='D'AND H.OUDEFINITIONID='"&sUnitID &"' "
		end if

		IF Cstr(sBookNo) <> "" and Cstr(sBookNo) <> "S" Then
			sSql = sSql & "AND BOOKNUMBER="& sBookNo &" "
		End IF
		sVal = request("hUserId")

		If Cstr(sVal) = "" then
			IF Cstr(sUserID) <> "A" Then
				sSql = sSql & "AND CREATEDBY = "&sUserID &" "
			End IF
		Else
			IF Cstr(sVal) <> "A" Then
				sSql = sSql & "AND CREATEDBY = "&sVal &" "
			End IF
		End IF


		if not Cstr(sFlag)= "VouDate" then
			sSql=sSql+"and CONVERT(DATETIME,H.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103) "
		end if

		sSql = sSql & " and TransCrDrIndication = 'D'"

		aFlag=true
		Response.Write ("<Input type=checkbox name=voutype value=A checked onclick=ChkVouType()>All&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType() >Created&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType() >Approved&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType() >Accounted&nbsp;")
		select case Cstr(sFlag)
		case "VouNo"
				sSql=sSql+"Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate,H.TransactionType Having H.CREATEDVOUCHERNO BETWEEN '"&sFrmNo &"' AND '"& sToNo&"'"
		case "VouAmount"
				sSql=sSql+"Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate,H.TransactionType Having Sum(D.Amount) BETWEEN "&sFrmAmt &" AND "& sToAmt&" "
		case "AccHead"
				if Request("selAccHead")="G" then
					sSql=sSql+"and H.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead="&iAccHead&") Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate,H.TransactionType "
				else
					saTemp=Split(iAccHead,"?")
					sSql=sSql+"and H.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
					" AccUnitPartyType ='"&Trim(saTemp(0))&"' and AccUnitPartySubType="&Trim(saTemp(1))&" and AccUnitPartyCode="&Trim(saTemp(3))&") Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate,H.TransactionType "
				end if
		case ""
				sSql=sSql+"Group By H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate "
		case "0"
				sSql=sSql+"Group By H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate "
		case "VouDate"
				sSql=sSql+"AND CONVERT(DATETIME,H.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& sFrmDate &"',103) " _
				& "AND CONVERT(DATETIME,'"& sToDate & "',103)Group By H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate "
		end select
		'sSql=sSql+"ORDER BY H.CREATEDTRANSNO DESC "
		sSql=sSql & " Order By " &  sSortBy  &  ",H.CREATEDTRANSNO "
	End IF



if not aFlag then
	if Trim(sUnitID)="" then
			sSql="Select H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,Sum(D.Amount)Amount,H.CreatedVoucherNo,H.CreatedVouchStatus ,Convert(Char,H.VoucherDate,103)VoucherDate From Acc_T_CreatedVoucherHeader " _
			& "H,Acc_T_CreatedVoucherDetails D Where D.CreatedTransNo = H.CreatedTransNo and H.BookCode = '08' and D.TransCrDrIndication='D' AND(H.CREATEDVOUCHSTATUS='0'"
		else
			sSql="Select H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,Sum(D.Amount)Amount,H.CreatedVoucherNo,H.CreatedVouchStatus ,Convert(Char,H.VoucherDate,103)VoucherDate From Acc_T_CreatedVoucherHeader " _
			& "H,Acc_T_CreatedVoucherDetails D Where D.CreatedTransNo = H.CreatedTransNo and H.BookCode = '08' and D.TransCrDrIndication='D'AND H.OUDEFINITIONID='"&sUnitID &"' AND H.BOOKNUMBER='"& sBookNo &"' AND(H.CREATEDVOUCHSTATUS='0'"
		end if

		Response.Write ("<Input type=checkbox name=voutype value=A onclick=ChkVouType()>All&nbsp;")
		if Instr(1,sSelVouTy,"C") > 0 then
			Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType() checked>Created&nbsp;")
			sSql=sSql+"OR H.CREATEDVOUCHSTATUS='010101' OR H.CREATEDVOUCHSTATUS='010102'"
		Else
			Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType()>Created&nbsp;")
		End IF

		if Instr(1,sSelVouTy,"P") > 0 then
			Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType() checked >Approved&nbsp;")
			sSql=sSql+"OR H.CREATEDVOUCHSTATUS='010103' OR H.CREATEDVOUCHSTATUS='010105'"
		Else
			Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType()>Approved&nbsp;")
		End IF

		if Instr(1,sSelVouTy,"T") > 0 Then
			Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType() checked >Accounted&nbsp;")
			sSql=sSql+"OR H.CREATEDVOUCHSTATUS='010104'"
		Else
			Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType()>Accounted&nbsp;")
		end if

		if not Cstr(sFlag)= "VouDate" then
			sSql=sSql+"and CONVERT(DATETIME,H.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103) "
		end if

		sSql = sSql & " and TransCrDrIndication = 'D'"

		'Response.Write " ============================== "

		select case Cstr(sFlag)
		case "VouNo"
				sSql=sSql+")"+"Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate Having H.CREATEDVOUCHERNO BETWEEN '"&sFrmNo &"' AND '"& sToNo&"'"
		case "VouAmount"
				sSql=sSql+")"+"Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate Having Sum(D.Amount) BETWEEN '"&sFrmAmt &"' AND '"& sToAmt&"'"
		case "AccHead"
				if Request("selAccHead")="G" then
					sSql=sSql+")"+"and H.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead="&iAccHead&") Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate,H.TransactionType "
				else
					saTemp=Split(iAccHead,"?")
					sSql=sSql+")"+"and H.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
					" AccUnitPartyType ='"&Trim(saTemp(0))&"' and AccUnitPartySubType="&Trim(saTemp(1))&" and AccUnitPartyCode="&Trim(saTemp(3))&") Group By H.CreatedTransNo,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate,H.TransactionType "
				end if
		case ""
				sSql=sSql+")"+"Group By H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate "
		case "0"
				sSql=sSql+")"+"Group By H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate "
		case "VouDate"
				sSql =sSql+")"+"AND CONVERT(DATETIME,H.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& sFrmDate &"',103) " _
				& "AND CONVERT(DATETIME,'"& sToDate & "',103) Group By H.CreatedTransNo,H.TransactionType,H.CreatedVoucherNo,H.CreatedVoucherNo,H.CreatedVouchStatus,H.VoucherDate "
		end select

		'sSql=sSql+"ORDER BY H.CREATEDTRANSNO DESC "

		sSql=sSql & " Order By " &  sSortBy  &  ",H.CREATEDTRANSNO "
end if

'	Response.Write "<p>"& sSql
%>
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
<td width="100%">
<div id="idUnprocessed" style="width: 575px; display: none">
<table cellpadding="0" cellspacing="0">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<!--<tr>
<td class="FieldCellSub">&nbsp;&nbsp;</td>
<td class="FieldCellSub">Unit Name</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selUnitId" class="FormElem" onchange="DisplayBook()">
	<option value="">Select Unit</option>
	 <%
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "Select Distinct OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION From VwUserUnitList WHere ApplicationCode = 1 and InternalUserID = "&getUserID()&" Order By OUDEFINITIONID "
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUnitLID = dcrs(0)
		set sUnitLName = dcrs(1)
		set sUnitSName = dcrs(3)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				if sUnitID=sUnitLID then
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
</tr>-->

<tr>
<td class="FieldCellSub">
</td>
<td class="FieldCellSub">GJ Book
</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selBook" class="FormElem" onchange="GetBookNo()">
	<option value="">Select Book</option>
</select>
</td>
</tr>

<tr>
<td class="FieldCellSub">
</td>
<td class="FieldCellSub">User ID
</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selUser" class="FormElem" onchange="GetUser()">
	<option value="0">Select User</option>
<%IF CStr(sUserID) = "A" Then %>
	<option value="A" Selected>All</option>
<%Else%>
	<option value="A">All</option>
	<%
		End IF
		Dim rsTemp,sqry
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		sqry = "SELECT DISTINCT INTERNALUSERID,LOGINID FROM VwUserUnitList WHERE APPLICATIONCODE = 1 "
		Response.Write "qry="& sqry
		rsTemp.Open sqry,con
		Do while not rsTemp.EOF
	%>
			<option value="<%=rsTemp(0)%>" <% If trim(sUserID) = trim(rsTemp(0)) then Response.Write "Selected" %> > <%=rsTemp(1)%></option>

	<%
			rsTemp.MoveNext
		loop
	rsTemp.Close
	%>
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
	<td class="FieldCellSub">Voucher No. To
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
	<td class="FieldCellSub">	</td>
	<td class="FieldCellSub">To
	</td>
	<td class="FieldCellSub">
		<%Response.Write InsertDatePicker("ctlVouToDate") %>
	</td>
</tr>

<tr>
	<td class="FieldCell">	</td>
	<td class="FieldCell">
		<%If CStr(sOptType)="VouAmount" then%>
		<input type="radio" value="VouAmount" name="OptCriteria" onclick="OptSelection()" checked>	Amount From
	<%else%>
		<input type="radio" value="VouAmount" name="OptCriteria" onclick="OptSelection()">	Amount From
	<% end if%>
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtFromAmount" Readonly size="20" class="FormElem">
	</td>
	<td class="FieldCellSub">	</td>
	<td class="FieldCellSub">To
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtToAmount" Readonly size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<% if CStr(sOptType)="AccHead" then %>
		<input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()" checked>	Account Head
	<%else%>
		<input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()">	Account Head
	<%end if%>
	</td>
	<td class="FieldCellSub" colspan="4">
		<select class="formelem" disabled OnChange="SelectAccHead()" size="1" name="selAccHead">
			<option value="0">Select Option</option>
			<option value="G">General Ledger</option>
		</select>
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
<table border="0" cellspacing="1px" class="ExcelTable" width="100%" >
<tr>
<td class="ExcelHeaderCell" width="10px" >S.No.
</td>
<td class="ExcelHeaderCell" width="10px">
</td>
<td class="ExcelHeaderCell">
<%
if trim(sField1) <> ""  then
	if instr(1,sField1,":") > 0 then
		Arr1 = Split(sField1,":")

		if Arr1(1) = "A" then ' if ascending order is exist, give option to descending order
			%>
			<span style="cursor:pointer" onclick="Sort(1,'N','D')">Number</span>
			<%
		else
			%>
			<span style="cursor:pointer" onclick="Sort(1,'N','A')">Number</span>
			<%
		end if
	end if
else
	%>
	<span style="cursor:pointer" onclick="Sort(1,'N','A')">Number</span>
	<%
end if
%>
</td>
<td class="ExcelHeaderCell">
<%
if trim(sField2) <> ""  then
	if instr(1,sField2,":") > 0 then
		Arr1 = Split(sField2,":")

		if Arr1(1) = "A" then ' if ascending order is exist, give option to descending order
			%>
			<span style="cursor:pointer" onclick="Sort(2,'D','D')">Date</span>
			<%
		else
			%>
			<span style="cursor:pointer" onclick="Sort(2,'D','A')">Date</span>
			<%
		end if
	end if
else
	%>
	<span style="cursor:pointer" onclick="Sort(2,'D','A')">Date</span>
	<%
end if
%>

</td>

<td class="ExcelHeaderCell">A/c. Head / Party
</td>
<td class="ExcelHeaderCell">
<%
if trim(sField5) <> ""  then
	if instr(1,sField5,":") > 0 then
		Arr1 = Split(sField5,":")

		if Arr1(1) = "A" then ' if ascending order is exist, give option to descending order
			%>
			<span style="cursor:pointer" onclick="Sort(5,'M','D')">Amount</span>
			<%
		else
			%>
			<span style="cursor:pointer" onclick="Sort(5,'M','A')">Amount</span>
			<%
		end if
	end if
else
	%>
	<span style="cursor:pointer" onclick="Sort(5,'M','A')">Amount</span>
	<%
end if
%>
</td>
<td class="ExcelHeaderCell">Status
</td>
</tr>

<%
	Dim sPartyCode

	iCnt=0
	'Response.Write "<P style='color:red' >" & sSql
	'Response.Write "<P style='color:red' >" & sSortBy
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
		sQuery="Select IsNull(AccUnitAccountHead,0),IsNull(AccUnitPartyCode,0) from Acc_T_CreatedVoucherDetails where CreatedTransNo="&Objrs("CreatedTransNo")
		with Objrs1
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source=sQuery
			.Open
		end with
		set Objrs1.ActiveConnection=nothing
			sHead =Objrs1(0)
			sPartyCode=Objrs1(1)
		Objrs1.Close

		if CStr(sHead)<>"0"  then
			sQuery ="SELECT ACCOUNTDESCRIPTION FROM ACC_M_GLACCOUNTHEAD WHERE ACCOUNTHEAD="&sHead
		else
			sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE="&sPartyCode
		end if
			with Objrs1
				.ActiveConnection=con
				.CursorLocation=3
				.CursorType=3
				.Source=sQuery
				.Open
			end with
			set Objrs1.ActiveConnection=nothing
				sAccHead=Objrs1(0)
			Objrs1.Close

		iCrTransNo=Objrs("createdtransno")
	%>
<tr>
<td class="ExcelSerial" align="center" ><%=iCnt%></td>
<td class="ExcelDisplayCell" align="center" width="10" >
<%If Right(CStr(Objrs("CREATEDVOUCHSTATUS")),2)="04" then %>
	<input type="checkbox" name="Chkbox" value="<%=iCrTransNo %>" disabled >
<%else%>
	<input type="checkbox" name="Chkbox" text="<%=Objrs("createdvoucherno")&"&"& Right(Objrs("transactiontype"),1)&"@"&Right(CStr(Objrs("CREATEDVOUCHSTATUS")),2) %>" value="<%=iCrTransNo %>" >
<%end if%>
<td class="ExcelDisplayCell" align="left" >
<a href="#" onclick="ShowVouch(<%=iCrTransNo%>); return false;" class="ExcelDisplayLink"><%=Objrs("createdvoucherno") %></a></td>
<td class="ExcelDisplayCell" align="left" ><%=Trim(Objrs("voucherdate"))%></td>
<td class="ExcelDisplayCell" align="left" ><%=sAccHead%></td>
<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(Objrs("amount")) %></td>
	<%

	sSql ="Select CreatedVouchStatus,VoucherNumber from Acc_T_CreatedVoucherHeader H , Acc_T_VoucherHeader v where H.CreatedTransNo=v.CreatedTransNo and " _
	& "right(H.CreatedVouchStatus,2)=04  and H.CreatedTransNo="&iCrTransNo
	With ObjRs1
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.Open
	End With
	Set ObjRs1.ActiveConnection = nothing
	if not ObjRs1.EOF then AccVoucherNo=ObjRs1(1)
	ObjRs1.Close

	if Right(CStr(Objrs("CREATEDVOUCHSTATUS")),2)="01" then
	Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Created" &"</td>")
	elseif Right(CStr(Objrs("CREATEDVOUCHSTATUS")),2)="04" then
	Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& AccVoucherNo &"</td>")
	else
	Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Approved" &"</td>")
	end if
	%>
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
<td align="center" class="ClearPixel" width="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top" align="right">
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
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td class="ActionCell">

<input type="button" value="Edit" name="B9" class="ActionButton" tabindex="3" onclick="ChkforEdit()">
<input type="button" value="Approve" name="B10" class="ActionButton" tabindex="3" onclick="ChkforApprove()">
<input type="button" value="Account" name="btnAcc" class="ActionButton" tabindex="3" onclick="ChkforAccount()">
<input type="button" value="Delete" name="B12" class="ActionButton" tabindex="3" onclick="ChkforDelete()">
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
