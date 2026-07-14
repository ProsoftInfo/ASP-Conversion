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
	'Program Name				:	ChequePrinting.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 01, 2011
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
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<%
	Dim sFinPeriod,Objrs,Objrs1,Objrs2,iCnt,sSql,iCrTransNo,sOptType
	Dim dcrs,sUnitLID,sUnitLName,sUnitSName,AccVoucherNo,sorgID,sBookID
	Dim sFormVal,sTemparr,sUnitID,sBookNo,sFrmDate,sToDate,sFrmAmt,sToAmt
	Dim sFrmNo,sToNo,iBookIndx,iAccIndx,sFlag,iAccHead,sAccHeadName,sSelVouTy
	Dim sCurrDate,sCurrDay,sCurrMon,sCurrYear,dChqFrmNo,dChqToNo,sOptVouType
	Dim AccFlag,sTemp
	Dim iVouNo,	dtVouDate ,	iVouAmt,iChqAmt
	Dim sValTemp2,sFinFromDate,sFinToDate
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
	'sOptType = Request("OptCriteria")
	sTemp = split(sFinPeriod,":")
	sFrmDate = Request("hFromDate")
	'IF trim(sFrmDate) = "" then sFrmDate = "01/04/"&sTemp(0)
	'sToDate = formatDate(date)
	'''''''
	sFormVal = Request("hFormVal")
	'Response.Write sFormVal &"<br><br>"
	sTemparr = Split(sFormVal,"|")

'	sUnitID = Request("hUnitID")
'	IF sUnitID = "" then sUnitID = "010101"
    sUnitID = Session("organizationcode")
    sUnitLID = Session("organizationcode")
    sUnitSName = Session("OrgShortName")


	sBookNo = Request("selBookNo")


	'sSelVouTy = Request("voutype")

	'sOptVouType = Request("OptVouTy")
	'IF sOptVouType = "" then sOptVouType = "C,D"
	'Response.Write "sOptVouType="&sOptVouType
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

	'IF CStr(sFrmDate) = "" and UBound(sTemparr)>2 Then
	'	sFrmDate = sTemparr(4)
	'End IF

	'IF CStr(sToDate) = "" and UBound(sTemparr)>2 Then
	'	sToDate = sTemparr(5)
	'End IF

	'IF CStr(sFrmAmt) = "" and UBound(sTemparr)>2 Then
	'	sFrmAmt = sTemparr(6)
	'End IF

	'IF CStr(sToAmt) = "" and UBound(sTemparr)>2 Then
	'	sToAmt = sTemparr(7)
	'End IF

	'IF CStr(sFrmNo) = "" and UBound(sTemparr)>2 Then
	'	sFrmNo = sTemparr(8)
	'End IF

	'IF CStr(sToNo) = "" and UBound(sTemparr)>2 Then
	'	sToNo = sTemparr(9)
	'End IF

	'IF CStr(iAccIndx) = "" and UBound(sTemparr)>2 Then
	'	iAccIndx = sTemparr(10)
	'End IF

	IF CStr(sFlag) = "" and UBound(sTemparr)>2 Then
		sFlag = sTemparr(11)
	End IF

	'IF CStr(iAccHead) = "" and UBound(sTemparr)>2 Then
	'	iAccHead = sTemparr(12)
	'End IF

	'IF CStr(sAccHeadName) = "" and UBound(sTemparr)>2 Then
	'	sAccHeadName = sTemparr(13)
	'End IF

	'IF CStr(dChqFrmNo) = "" and UBound(sTemparr)>2 Then
	'	dChqFrmNo = sTemparr(14)
	'End IF

	'IF CStr(dChqToNo) = "" and UBound(sTemparr)>2 Then
	'	dChqToNo = sTemparr(15)
	'End IF

	'Response.Write dChqFrmNo &" " & dChqToNo

sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFromDate = "01/04/"& sValTemp2(0)
sFinToDate = "31/03/"&sValTemp2(1)
if Trim(sFrmDate)="" then
    sFrmDate = sFinFromDate
    sToDate = sFinToDate
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
<script type="application/xml" data-itms-xml-island="1" ID="PartyData"><Party/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData"><account/></script>
<script type="application/xml" data-itms-xml-island="1" ID="SearchData" ><Root/></script>

<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT SRC="../../scripts/VouTransactions.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT SRC="../../scripts/BankVouchersCompat.js"></SCRIPT>
<SCRIPT SRC="../../scripts/ChequePrintingCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="init();SetDate();DisplayBook();">
<%
	Const iPageSize=20
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt
	Dim oDOM,objfs,Root,node,sXMLFlag
	Set oDOM = CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	sSelVouTy = Request("voutype")
	sOptVouType = Request("OptVouTy")
	'dChqFrmNo = Request("hChqFrom")
	'dChqToNo = Request("hChqTo")
	'IF sOptVouType = "" then sOptVouType = "C,D"
	'Response.Write  "sSelVouTy ="&sSelVouTy
	sXMLFlag = False
	IF  objfs.FileExists(Server.MapPath("../temp/transaction/SearchCriteria_ACC_"&session.SessionID&".xml")) then
		oDOM.load  server.MapPath("../temp/transaction/SearchCriteria_ACC_"&session.SessionID&".xml")

		Set Root = oDOM.documentElement

		if Root.haschildnodes then
			IF trim(Root.getAttribute("Src")) = "BankGrid" then
				For each node in Root.childnodes
					if trim(node.NodeName) = "VoucherType" then
						sOptVouType = node.getAttribute("Value")
					end if
					if trim(node.NodeName) = "VoucherNo" then
						sFrmNo = node.getAttribute("From")
						sToNo = node.getAttribute("To")
						If trim(sFrmNo) <> "" and trim(sToNo) <> "" then iVouNo = "VouNo"
					end if
					if trim(node.NodeName) = "VoucherDate" then

						sFrmDate = node.getAttribute("From")
						sToDate  = node.getAttribute("To")
						If trim(sFrmDate) <> "" and trim(sToDate) <> "" then dtVouDate =	"VouDate"
					end if
					if trim(node.NodeName) = "VoucherAmount" then
						sFrmAmt = node.getAttribute("From")
						sToAmt  = node.getAttribute("To")
						If trim(sFrmAmt) <> "" and trim(sToAmt) <> "" then iVouAmt =	"VouAmount"
					end if
					if trim(node.NodeName) = "ChequeNo" then
						dChqFrmNo = node.getAttribute("From")
						dChqToNo  = node.getAttribute("To")
						'Response.Write "CHQ NO ="& dChqFrmNo
						If trim(dChqFrmNo) <> "" and trim(dChqToNo) <> "" then iChqAmt =	"Cheque"
					end if

					if trim(node.NodeName) = "AccHead" then
						iAccHead = node.getattribute("No")
						sAccHeadName  = node.getattribute("Name")
						iAccIndx = node.getattribute("AccIndex")
						IF trim(iAccHead) <> "0" then AccFlag = True

					end if
				Next
			Else
				sXMLFlag = True
			End IF 'IF trim(Root.getAttribute("Src")) = "BankGrid" then
		end if
	Else
		sXMLFlag = True

		'Response.Write "test-"&iAccHead
	End IF
	IF trim(sXMLFlag) = "True" then
			sOptVouType = "C,D"
			sAccHeadName = Request("hAccTxt")
			iVouNo    = Request("hVouNoFlag")
			'Response.Write "iVouNo="& iVouNo &"<BR>"
			'dtVouDate = Request("hVouDtFlag")
			iVouNo    = Request("hVouNoFlag")
			iVouAmt   = Request("hVouAmtFlag")
			iChqAmt   = Request("hChqFlag")
			dtVouDate =	"VouDate"
			iVouAmt   = Request("hVouAmtFlag")
			sFrmAmt = Request("hAmtFrom")
			sToAmt = Request("hAmtTo")
			iBookIndx = Request("hBookNo")
			iAccIndx = Request("hAccIndex")
			sFlag = Cstr(sFlag)
			sOptType = "VouDate"
			iAccHead = Request("hAccHead")
			If trim(iAccHead) <> "" then
				AccFlag = True
				sAccHeadName = Request("hAccTxt")
			Else
				AccFlag = False
				iAccHead = 0
				sAccHeadName = ""
			End IF
		End IF
'	 oDOM.save server.MapPath("../temp/transaction/SearchCriteria_ACC_"&session.SessionID&".xml")

	'iCnt=Request.Form("hCnt")
%>
	<form method="POST" name="formname" action="ChequePrinting.asp">
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

	<input type=hidden name="hVouNoFlag" value="<%=iVouNo%>">
	<input type=hidden name="hVouDtFlag" value="<%=dtVouDate%>">
	<input type=hidden name="hVouAmtFlag" value="<%=iVouAmt%>">
	<input type=hidden name="hChqFlag" value="<%=iChqAmt%>">

	<input type=hidden name="hAmtFrom" value="<%=sFrmAmt%>">
	<input type=hidden name="hAmtTo" value="<%=sToAmt%>">
	<input type=hidden name="hVouFrom" value="<%=sFrmNo%>">
	<input type=hidden name="hVouTo" value="<%=sToNo%>">
	<input type=hidden name="hVouName" value="BA">
	<input type=hidden name="hFinPeriod" value="<%=sFinPeriod%>">
	<input type=hidden name="hFormVal" value="">
	<input type=hidden name="hChqFrom" value="<%=dChqFrmNo%>">
	<input type=hidden name="hChqTo" value="<%=dChqToNo%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Cheque Printing
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
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="center" class="SubTitle">&nbsp;&nbsp;
<%
	Dim aFlag,saTemp
	aFlag=false

	'Response.Write sFrmAmt
	'Response.Write "<p>sSelVouTy="&sSelVouTy
	IF CStr(sSelVouTy) = "A" or CStr(sSelVouTy) = "" or InStr(1,sSelVouTy,CStr("C, P, T"))>0 Then
		'if Trim(sUnitID)="" then
		'		'sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION,isNull(A.BRSTransactionNo,0) BRSTransactionNo FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
		'		sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
		'		& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD INNER JOIN ACC_T_CREATEDVOUCHERINSTRUMENTDET AS I ON A.CREATEDTRANSNO = I.CREATEDTRANSNO WHERE A.BOOKCODE='02'"
		'else
		'		sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION  FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
		'		& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD INNER JOIN ACC_T_CREATEDVOUCHERINSTRUMENTDET AS I ON A.CREATEDTRANSNO = I.CREATEDTRANSNO WHERE A.BOOKCODE='02'AND A.OUDEFINITIONID='"&sUnitID &"' "
		'end if

		sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION  FROM ACC_T_CREATEDVOUCHERHEADER AS A " & _
			 "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD WHERE A.BOOKCODE='02'"
		if Trim(sUnitID)<>"" then
			 sSql = sSql & " AND A.OUDEFINITIONID='"&sUnitID &"' "
		end if

		sSql=" SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION,I.BANKINSTRUMENTNO,"&_
			 " CONVERT(DATETIME,I.BANKINSTRUMENTDATE,103),I.INSTRUMENTAMOUNT,A.PAYTORECDFROM,A.BOOKNUMBER  "&_
			 " FROM ACC_T_CREATEDVOUCHERHEADER A , ACC_M_GLACCOUNTHEAD V,Acc_T_CreatedVoucherInstrumentDet I" & _
			 " WHERE A.ACCOUNTHEAD=V.ACCOUNTHEAD AND I.CREATEDTRANSNO = A.CREATEDTRANSNO  AND A.BOOKCODE='02'"

		if Trim(sUnitID)<>"" then
			 sSql = sSql & " AND A.OUDEFINITIONID='"&sUnitID &"' "
		end if

		'Response.Write "<BR> chk="& sOptVouType &"<BR>"
		IF CStr(sOptVouType) = "C,D" then
			sSql = sSql & "AND A.TRANSACTIONTYPE IN('BAR','BAP') "
		ElseIF trim(sOptVouType) = "C" then
			sSql = sSql & "AND A.TRANSACTIONTYPE = 'BAR' "
		ElseIF trim(sOptVouType) = "D" then
			sSql = sSql & "AND A.TRANSACTIONTYPE = 'BAP' "
		End IF
		IF Cstr(sBookNo) <> "" and Cstr(sBookNo) <> "S" Then
			sSql = sSql&" AND A.BOOKNUMBER="& Cstr(sBookNo) &"  "
		End IF


'************* This Will Display The Current Date Vouchers Only on First Display *******************************
		'IF Cstr(sFlag) = "" and Cstr(sCurrDate) <> "" Then
		'	sSql = sSql &" AND CONVERT(DATETIME,A.VOUCHERDATE,103) = CONVERT(DATETIME,'"&sCurrDate&"',103) "
		'End IF

		aFlag=true
		Response.Write ("<Input type=checkbox name=voutype value=A checked onclick=ChkVouType()>All&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType() >Created&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType() >Approved&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType() >Accounted&nbsp;")
		if trim(iVouNo) = "VouNo" then
				sSql=sSql+" AND A.CREATEDVOUCHERNO BETWEEN '"&Cstr(sFrmNo) &"' AND '"& Cstr(sToNo)&"'"
		end if
		if trim(iVouAmt) = "VouAmount" then
				sSql=sSql+" AND A.VOUCHERAMOUNT BETWEEN "&sFrmAmt &" AND "&sToAmt &""

		end if
		if trim(iChqAmt) = "Cheque" then
				sSql=sSql+" and A.CreatedTransNo in ( Select I.CreatedTransNo from Acc_T_CreatedVoucherInstrumentDet I" &_
						" where isNumeric(I.BankInstrumentNo) = 1 and Cast(I.BankInstrumentNo AS Numeric) Between "&dChqFrmNo&" and "&dChqToNo&" )"
		end if
		If Accflag = True then
			if Request("selacchead")="G" then
				sSql=sSql+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead in ("&Request("hAccHead")&")) "
			else
				'Response.Write "aaa="& Request("hAccHead")
				IF trim(Request("hAccHead")) <> "" and trim(Request("hAccHead")) <> "0" then
					saTemp=Split(Request("hAccHead"),"?")
					sSql=sSql+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
					" AccUnitPartyType in ("&Trim(saTemp(0))&") and AccUnitPartySubType in ("&Trim(saTemp(1))&") and AccUnitPartyCode  in ("&Trim(saTemp(3))&")) "
				end if
			end if
		End IF
		if not Cstr(dtVouDate)= "VouDate" then
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103) ORDER BY A.CREATEDTRANSNO DESC "
		else
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& Cstr(sFrmDate) &"',103) " _
			& "AND CONVERT(DATETIME,'"& Cstr(sToDate) & "',103) ORDER BY A.CREATEDTRANSNO DESC "
		end if

'		Response.Write "1="& sSql
	End IF

	'Response.Write aFlag

	if not aFlag then
		'if Trim(sUnitID)="" then
		'		sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
		'		& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD INNER JOIN ACC_T_CREATEDVOUCHERINSTRUMENTDET AS I ON A.CREATEDTRANSNO = I.CREATEDTRANSNO WHERE A.BOOKCODE='02'"
		'else
		'		sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
		'		& "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD INNER JOIN ACC_T_CREATEDVOUCHERINSTRUMENTDET AS I ON A.CREATEDTRANSNO = I.CREATEDTRANSNO WHERE A.BOOKCODE='02' AND " _
		'		& "A.OUDEFINITIONID='"&sUnitID &"' "
		'		IF Cstr(sBookNo) <> "" and Cstr(sBookNo) <> "S" Then
		'			sSql=sSql & " AND A.BOOKNUMBER='"& Cstr(sBookNo) &"' "
		'		End IF
		'end if

		sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION  FROM ACC_T_CREATEDVOUCHERHEADER AS A " & _
			 "INNER JOIN ACC_M_GLACCOUNTHEAD AS V ON A.ACCOUNTHEAD=V.ACCOUNTHEAD WHERE A.BOOKCODE='02'"
		if Trim(sUnitID)<>"" then
			 sSql = sSql & " AND A.OUDEFINITIONID='"&sUnitID &"' "
		end if
		IF Cstr(sBookNo) <> "" and Cstr(sBookNo) <> "S" Then
			sSql=sSql & " AND A.BOOKNUMBER='"& Cstr(sBookNo) &"' "
		End IF

		'ADDED BY UMAMAHESWARI S ON 01st APRIL 2011
		sSql=" SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,V.ACCOUNTDESCRIPTION,"&_
			 " I.BANKINSTRUMENTNO, CONVERT(DATETIME,I.BANKINSTRUMENTDATE,103),I.INSTRUMENTAMOUNT,A.PAYTORECDFROM,A.BOOKNUMBER "&_
			 " FROM ACC_T_CREATEDVOUCHERHEADER A ,ACC_M_GLACCOUNTHEAD V,Acc_T_CreatedVoucherInstrumentDet I" & _
			 " WHERE A.ACCOUNTHEAD=V.ACCOUNTHEAD AND I.CREATEDTRANSNO = A.CREATEDTRANSNO AND A.BOOKCODE='02'"

		if Trim(sUnitID)<>"" then
			 sSql = sSql & " AND A.OUDEFINITIONID='"&sUnitID &"' "
		end if
		IF Cstr(sBookNo) <> "" and Cstr(sBookNo) <> "S" Then
			sSql=sSql & " AND A.BOOKNUMBER='"& Cstr(sBookNo) &"' "
		End IF


		 'Response.Write sOptVouType
		IF CStr(sOptVouType) = "C,D" then
			sSql = sSql & "AND A.TRANSACTIONTYPE IN('BAR','BAP') "
		ElseIF trim(sOptVouType) = "C" then
			sSql = sSql & "AND A.TRANSACTIONTYPE = 'BAR' "
		ElseIF trim(sOptVouType) = "D" then
			sSql = sSql & "AND A.TRANSACTIONTYPE = 'BAP' "
		End IF
		sSql = sSql & "AND(A.CREATEDVOUCHSTATUS='0' "
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
		sSql=sSql+")"
		if trim(iVouNo) = "VouNo" then
				sSql=sSql+" AND A.CREATEDVOUCHERNO BETWEEN '"&Cstr(sFrmNo) &"' AND '"& Cstr(sToNo)&"'"
		end if
		if trim(iVouAmt) = "VouAmount" then
			sSql=sSql+" AND A.VOUCHERAMOUNT BETWEEN "&sFrmAmt &" AND "&sToAmt &""
		end if


		if trim(iChqAmt) =  "Cheque" then
			sSql=sSql+" and A.CreatedTransNo in ( Select I.CreatedTransNo from Acc_T_CreatedVoucherInstrumentDet I" &_
						" where isNumeric(I.BankInstrumentNo) = 1 and Cast(I.BankInstrumentNo AS Numeric) Between "&dChqFrmNo&" and "&dChqToNo&" )"
		end if

		If Accflag = True then
			if Request("selacchead")="G" then
				sSql=sSql+" and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead in ("&Request("hAccHead")&")) "
			else
				IF trim(iAccHead) <> "0" and  trim(iAccHead) <> "" then
					saTemp=Split(iAccHead,"?")
					sSql=sSql+" and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
					" AccUnitPartyType in ("&Trim(saTemp(0))&") and AccUnitPartySubType in ("&Trim(saTemp(1))&")  and AccUnitPartyCode in ("&Trim(saTemp(3))&")) "
				End if
			end if
		End IF

		if not Cstr(dtVouDate)= "VouDate" then
			sSql=sSql+"  AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103)  ORDER BY A.CREATEDTRANSNO DESC "
		else
			sSql=sSql+" AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& Cstr(sFrmDate) &"',103) " _
			& "AND CONVERT(DATETIME,'"& Cstr(sToDate) & "',103)  ORDER BY A.CREATEDTRANSNO DESC "
		end if
'		  Response.Write "2="& sSql
	end if
'	Response.Write sSql
%>
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
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Bank Book</td>
	<td class="FieldCellSub" colspan="4">
	<select size="1" name="selBook" class="FormElem" onchange="GetBookNo()">
		<option value ="S">Select</option>
	</select>
	</td>
</tr>
<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Voucher Type</td>
	<td class="FieldCellSub" colspan="4">
	<input type="radio" name="OptVouTy" class="FormElem" value="C,D" <%If sOptVouType = "C,D" then Response.Write "Checked"%>>Both
	<input type="radio" name="OptVouTy" class="FormElem" value="C"  <%If sOptVouType = "C" then Response.Write "Checked"%>>Receipts
	<input type="radio" name="OptVouTy" class="FormElem" value="D"  <%If sOptVouType = "D" then Response.Write "Checked"%>>Payments
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%IF CStr(sOptType) = "VouNo" Then %>
		<input type="Checkbox" value="VouNo" name="ChkVouNo" onclick="Optselection()" >Voucher No. From
	<%Else%>
		<input type="checkbox" value="VouNo" name="ChkVouNo" onclick="Optselection()">Voucher No. From
	<%End IF %>
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtVouNoFrom"  size="20" class="FormElem">
	</td>

	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">To
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtVouNoTo"  size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%IF CStr(sOptType) = "VouDate" Then %>
		<input type="checkbox" value="VouDate" name="ChkVouDt" onclick="OptSelection()" Checked>Voucher Date
	<%Else%>
		<input type="checkbox" value="VouDate" name="ChkVouDt" onclick="OptSelection()" >Voucher Date
	<%End IF %>
	</td>
	<%'Response.Write InsertDatePicker("ctlVouFromDate") %>

    <%'Response.Write InsertDatePicker("ctlVouFromDate") %>

    <td class="FieldCellSub" valign="middle">
		<input type="date" id="ctlVouFromDate" name="ctlVouFromDate" class="FormElem itms-date-picker" style="width:89px">
	</td>


	<td class="FieldCell"></td>
	<td class="FieldCellSub">To
	</td>

	<%'Response.Write InsertDatePicker("ctlVouToDate") %>

        <td class="FieldCellSub" valign="middle">
			<input type="date" id="ctlVouToDate" name="ctlVouToDate" class="FormElem itms-date-picker" style="width:89px">
		</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%If CStr(sOptType)="VouAmount" then%>
		<input type="checkbox" value="VouAmount" name="ChkVouAmt" onclick="OptSelection()" >Amount From
	<%else%>
		<input type="checkbox" value="VouAmount" name="ChkVouAmt" onclick="OptSelection()">Amount From
	<% end if%>
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtFromAmount"  size="20" class="FormElem"></td>
	<td class="FieldCellSub"></td>

	<td class="FieldCellSub">To	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtToAmount"  size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%If CStr(sOptType)="Cheque" then%>
		<input type="checkbox" value="Cheque" name="ChkChq" onclick="OptSelection()" >Cheque No From
	<%else%>
		<input type="checkbox" value="Cheque" name="ChkChq" onclick="OptSelection()">Cheque No From
	<% end if%>
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtFromChqNo"   size="20" class="FormElem"></td>
	<td class="FieldCellSub"></td>

	<td class="FieldCellSub">To	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtToChqNo"   size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCellSub">
	<%' if CStr(sOptType)="AccHead" then %>
		<!--input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()" checked-->Account Head
	<%'else%>
		<!--input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()"-->
	<%'end if%>
	</td>
	<td class="FieldCellSub" colspan="4">
	<% 'if CStr(sOptType)="AccHead" then %>
		<select class="FormElem" OnChange="SelectAccHead()" size="1" name="selAccHead">
	<%'Else%>
		<!--select class="formelem" disabled OnChange="SelectAccHead()" size="1" name="selAccHead"-->
	<%'ENd IF %>
			<option value="0">Select Option</option>
			<option value="G">General Ledger</option>
		</select>
				<a href="#" onclick="ResetAccHead(); return false;"><img border="0" width="11" height="11" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Remove Account Head" ></a>
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

</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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
<table border="0" cellspacing="1px" class="ExcelTable" width="100%">
<tr>
<td class="ExcelHeaderCell" align="center" width="10px" >S.No.
</td>
<td class="ExcelHeaderCell" width="10px" >
</td>
<td class="ExcelHeaderCell">Instrument<Br>Number
</td>
<td class="ExcelHeaderCell" >Instrument<Br>Date
</td>
<td class="ExcelHeaderCell" >Instrument<BR> Amount
</td>
<td class="ExcelHeaderCell">PayTo / <BR>Received From
</td>
<td class="ExcelHeaderCell">Voucher No
</td>
<td class="ExcelHeaderCell">Voucher Date
</td>
<td class="ExcelHeaderCell">Amount
</td>
</tr>
<%

	Dim iParCode,AccParName,iBrsCount,iInsCount,iInsBrsount
	'Response.Write "Query="&sSql
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
		sSql = "select Count(*) from Acc_T_CreatedVoucherInstrumentDet where createdTransNo = "&iCrTransNo&" "
		Objrs1.Open sSql,con
		IF Not Objrs1.EOF Then
			iInsCount = Objrs1(0)
		End If
		Objrs1.Close
		sSql = "select Count(*) from Acc_T_CreatedVoucherInstrumentDet where createdTransNo = "&iCrTransNo&" "&_
				"and BRSTransactionNo <> 0"
		Objrs1.Open sSql,con
		IF Not Objrs1.EOF Then
			iInsBrsount = Objrs1(0)
		End If
		Objrs1.Close

		iBrsCount = cint(iInsCount)- cint(iInsBrsount)
	%>
<tr>
<td class="ExcelSerial" align="center" ><%=iCnt%></td>
<td class="ExcelDisplayCell" align="center" width="10" >
<%'If Right(CStr(Objrs("createdvouchstatus")),2)="04" then %>
	<!--<input type="checkbox" name="Chkbox" value="<%=iCrTransNo %>" disabled >-->
<%'else%>
	<input type="checkbox" name="Chkbox" text="<%=Objrs("createdvoucherno")&"&"& Right(Objrs("transactiontype"),1)&"@"& Right(CStr(Objrs("createdvouchstatus")),2) &"@" & Trim(iBrsCount)&"@" & trim(AccParName)&"@"&trim(Objrs(12))%>" value="<%=iCrTransNo %>" >
<%'end if%>
<td class="ExcelDisplayCell" align="left">
	<a href="#" onclick="ShowInsDet(<%=iCrTransNo%>); return false;" class="ExcelDisplayLink"><%=Objrs(8)%></a></td>
<td class="ExcelDisplayCell" align="right"><%=Objrs(9) %></td>
<td class="ExcelDisplayCell" align="right"><%=FormatNumber(Objrs(10))%></td>
<td class="ExcelDisplayCell" align="lEFT"><%=Objrs(11)%></td>

<td class="ExcelDisplayCell" align="left" >
	<a href="#" onclick="ShowVouch(<%=iCrTransNo %>); return false;" class="ExcelDisplayLink"><%=Objrs("createdvoucherno") %></a>
</td>
<td class="ExcelDisplayCell" align="left" ><%=FormatDate(Objrs("voucherdate"))%></td>
<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(Objrs("Voucheramount")) %></td>

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

	if Right(CStr(Objrs("createdvouchstatus")),2)="01" then
		'Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Created" &"</td>")
	elseif Right(CStr(Objrs("createdvouchstatus")),2)="04" then
		'Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& AccVoucherNo &"</td>")
	else
		'Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Approved" &"</td>")
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
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td class="ActionCell">

<input type="button" value="Print Cheque" name="B9" class="ActionButtonX" tabindex="3" onclick="PrintCheque()">
<!--<input type="button" value="Edit" name="B9" class="ActionButton" tabindex="3" onclick="ChkforEdit()">
<input type="button" value="Approve" name="B10" class="ActionButton" tabindex="4" onclick="ChkforApprove()">
<input type="button" value="Account" name="B11" class="ActionButton" tabindex="5" onclick="ChkforAccount()">
<input type="button" value="Delete" name="B12" class="ActionButton" tabindex="6" onclick="ChkforDelete()">-->

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
