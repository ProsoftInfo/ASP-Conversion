<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CashVouchers.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Sre Hari M
	'Created On					:	Feb 15, 2006
	'Modified By                :   Ragavendran R
	'Modified On                :   Jan 18,2011
	'Modified By				:	UmaMaheswari S
	'Tables Used				:	April 29, 2011
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
	Session("ACTN")=trim(Request("ACTN"))
%>
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim sFinPeriod,Objrs,Objrs1,Objrs2,iCnt,sSql,iCrTransNo
	Dim dcrs,sUnitLID,sUnitLName,sUnitSName,sOptType,AccVoucherNo
	Dim sFormVal,sTemparr,sUnitID,sBookNo,sFrmDate,sToDate,sFrmAmt,sToAmt
	Dim sFrmNo,sToNo,iBookIndx,iAccIndx,sFlag,iAccHead,sAccHeadName
	Dim sSelVouTy,sCurrDate,sCurrDay,sCurrMon,sCurrYear,sUserID,sVouTy,sOptVouType
	Dim AccFlag, sACTN
	Dim iVouNo,	dtVouDate ,	iVouAmt
	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")

	sFinPeriod=session("finperiod")
	'sOptType = Request("OptCriteria")

	'''''''
	iVouNo    = Request("hVouNoFlag")
	dtVouDate =	Request("hVouDtFlag")
	iVouAmt   = Request("hVouAmtFlag")

	sFormVal = Request("hFormVal")
'	sUnitID = Request("hUnitID")
	sUserID = getUserId()
    sUnitID = Session("organizationcode")
    sUnitSName = Session("OrgShortName")
	sTemparr = Split(sFormVal,"|")
	sBookNo = Request("selBookNo")
	sFrmDate = Request("hFromDate")
	sToDate =  Request("hToDate")
	sFrmAmt = Request("hAmtFrom")
	sToAmt = Request("hAmtTo")
	iBookIndx = Request("hBookNo")
	iAccIndx = Request("hAccIndex")
	sFlag = Cstr(sFlag)

	iAccHead = Request("hAccHead")
	
	If trim(iAccHead) <> "" then AccFlag = True
	sAccHeadName = Request("hAccTxt")
	sSelVouTy = Request("voutype")
	'sACTN = trim(Request("ACTN"))
	sACTN = Session("ACTN")
	'Response.Write "<p><font color=red>sACTN="&sACTN
	'Response.Write "<p><font color=red>sACTN="&Request("hAction")

	sOptVouType = Request("OptVouTy")
	IF sOptVouType = "" then sOptVouType = "C,D"
	'Response.Write "iAccHead="&iAccHead
	'Response.Write "sOptVouType="& sOptVouType
	'sUserID = Request("selUser")
	'Response.Write "sUserID="&sUserID
	'Response.Write sFormVal & "<br><br>"

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

	IF UBound(sTemparr)>14 Then
		sVouTy = sTemparr(15)
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
'Response.Write sVouTy
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
<SCRIPT SRC="../../scripts/VouTransactions.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT SRC="../../scripts/CashVouchersCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate()">
<%
	Const iPageSize=16
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	'iCnt=Request.Form("hCnt")
%>
	<form method="POST" name="formname" action="CashVouchers.asp?ACTN=<%=sACTN%>" >
	<input type=hidden name="hTransNo" value="">
	<input type=hidden name="hAppNo" value="">
	<input type=hidden name="hUnitNo" value="<% =sUnitID%>">
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

	<input type=hidden name="hAmtFrom" value="<%=sFrmAmt%>">
	<input type=hidden name="hAmtTo" value="<%=sToAmt%>">
	<input type=hidden name="hVouFrom" value="<%=sFrmNo%>">
	<input type=hidden name="hVouTo" value="<%=sToNo%>">
	<input type=hidden name="hVouName" value="CA">
	<input type=hidden name="hFinPeriod" value="<%=sFinPeriod%>">
	<input type=hidden name="hFormVal" value="">
	<input type=hidden name="hUserID" value="">
	
	<input type=hidden name="hAction" value="<%=sACTN%>">
	<input type=hidden name="hVocType" value="<%=sOptVouType%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle">
				<p align="center">Cash Vouchers
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
<td valign="right" class="SubTitle">&nbsp;&nbsp;
<%
	Dim aFlag,saTemp,iCrBy,sCrName,sVal
	aFlag=false
	'Response.Write Trim(sSelVouTy) &"====== "
	IF CStr(sSelVouTy) = "A" or CStr(sSelVouTy) = "" or InStr(1,sSelVouTy,CStr("C, P, T"))>0 Then
		if Trim(sUnitID)="" then
				sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,A.CREATEDBY,ISNULL(A.FROMAPPLICATION,0),ISNULL(PURCHASEBILLTYPE,'') FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
				& "WHERE A.BOOKCODE='01'"
		else
				sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,A.CREATEDBY,ISNULL(A.FROMAPPLICATION,0),ISNULL(PURCHASEBILLTYPE,'') FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
				& "WHERE A.BOOKCODE='01' AND A.OUDEFINITIONID='"&sUnitID &"' "
		end if

		IF trim(sOptVouType) = "C" then
			sSql = sSql & "AND A.TRANSACTIONTYPE = 'CAR' "
		ElseIF trim(sOptVouType) = "D" then
			sSql = sSql & "AND A.TRANSACTIONTYPE = 'CAP' "
		End IF
		IF Cstr(sBookNo) <> "" and Cstr(sBookNo) <> "S" Then
			sSql = sSql & "AND A.BOOKNUMBER="& sBookNo &" "
		End IF
		sVal = request("hUserId")
		'Response.Write "<BR><BR>sUserID="& sVal &"<BR><BR>"
		If Cstr(sVal) = "" then
			IF Cstr(sUserID) <> "A" Then
				sSql = sSql & "AND A.CREATEDBY = "&sUserID &" "
			End IF
		Else
			IF Cstr(sVal) <> "A" Then
				sSql = sSql & "AND A.CREATEDBY = "&sVal &" "
			End IF
		End IF

		IF CStr(sVouTy) <> "" and Trim(CStr(sVouTy)) <> "C,D" Then
			sSql = sSql & "AND A.CrDrIndication IN('"&sVouTy&"') "
		End IF

		aFlag=true
		Response.Write ("<Input type=checkbox name=voutype value=A checked onclick=ChkVouType()>All&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType() >Created&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType() >Approved&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType() >Accounted&nbsp;")

		If trim(iVouNo) = "VouNo" then
				sSql=sSql+"AND A.CREATEDVOUCHERNO BETWEEN '"&sFrmNo &"' AND '"& sToNo&"'"
		End If
		If trim(iVouAmt) = "VouAmount" then
				sSql=sSql+"AND A.VOUCHERAMOUNT BETWEEN "&Cstr(sFrmAmt)&" AND "& Cstr(sToAmt)&" "
		End IF

		If Accflag = True then

			if Request("selacchead")="G" then
				sSql=sSql+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead in ("&Request("hAccHead")&")) "
			else
				saTemp=Split(Request("hAccHead"),"?")
				sSql=sSql+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
				" AccUnitPartyType in ("&Trim(saTemp(0))&") and AccUnitPartySubType in ("&Trim(saTemp(1))&") and AccUnitPartyCode  in ("&Trim(saTemp(3))&")) "
			end if
		End IF
	'Response.Write dtVouDate
		if not Cstr(dtVouDate)= "VouDate" then
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103)  ORDER BY CONVERT(DATETIME,A.VOUCHERDATE,103) DESC,A.CREATEDTRANSNO   "
		else
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& sFrmDate &"',103) " _
			& "AND CONVERT(DATETIME,'"& sToDate & "',103)  ORDER BY CONVERT(DATETIME,A.VOUCHERDATE,103) DESC,A.CREATEDTRANSNO   "
		end if
		'Response.Write "1="& sSql
	End IF

if not aFlag then
	if Trim(sUnitID)="" then
			sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,A.CREATEDBY,ISNULL(A.FROMAPPLICATION,0),ISNULL(PURCHASEBILLTYPE,'') FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
			& "WHERE A.BOOKCODE='01'AND(A.CREATEDVOUCHSTATUS='0'"
	else
			sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,A.CREATEDBY,ISNULL(A.FROMAPPLICATION,0),ISNULL(PURCHASEBILLTYPE,'') FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
			& "WHERE A.BOOKCODE='01' AND A.OUDEFINITIONID='"&sUnitID &"' "
	end if

	IF CStr(sBookNo) <> "" and CStr(sBookNo) <> "S" Then
		sSql = sSql & " AND A.BOOKNUMBER="& sBookNo &"  "
	End IF

	'IF CStr(sVouTy) <> "" and Trim(CStr(sVouTy)) <> "C,D" Then
	'	sSql = sSql & "AND A.CrDrIndication IN('"&sVouTy&"') "
	'End IF
	IF trim(sOptVouType) = "C" then
		sSql = sSql & "AND A.TRANSACTIONTYPE = 'CAR' "
	ElseIF trim(sOptVouType) = "D" then
		sSql = sSql & "AND A.TRANSACTIONTYPE = 'CAP' "
	End IF

	if Trim(sUnitID)<> "" then
		sSql = sSql & "AND(A.CREATEDVOUCHSTATUS='0' "
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
		sSql=sSql+")"
		 'Response.Write Trim(sFlag) &"====== "
		if trim(iVouNo) = "VouNo" then
				sSql=sSql+"AND A.CREATEDVOUCHERNO BETWEEN '"&sFrmNo&"' AND '"& sToNo&"'"
		end if
		if trim(iVouAmt) = "VouAmount" then
				sSql=sSql+"AND A.VOUCHERAMOUNT BETWEEN '"&sFrmAmt&"' AND '"& sToAmt&"'"

		end if


		If Accflag = True then
			if Request("selacchead")="G" then
				sSql=sSql+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead in ("&Request("hAccHead")&")) "
			else
				saTemp=Split(iAccHead,"?")
				sSql=sSql+"and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
				" AccUnitPartyType in ("&Trim(saTemp(0))&") and AccUnitPartySubType in ("&Trim(saTemp(1))&")  and AccUnitPartyCode in ("&Trim(saTemp(3))&")) "
			end if
		End IF


		if not Cstr(dtVouDate)= "VouDate" then
			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
			& "AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103) ORDER BY CONVERT(DATETIME,A.VOUCHERDATE,103) DESC,A.CREATEDTRANSNO  "
		else

			sSql=sSql+"AND CONVERT(DATETIME,A.VOUCHERDATE,103)BETWEEN CONVERT(DATETIME,'"& sFrmDate &"',103) " _
			& "AND CONVERT(DATETIME,'"& sToDate & "',103) ORDER BY CONVERT(DATETIME,A.VOUCHERDATE,103) DESC,A.CREATEDTRANSNO   "
		end if
		'Response.Write "2="& sSql
end if

	 ' Response.Write sSql
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
				if CStr(sUnitID) = sUnitLID then
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
<td class="FieldCellSub">Cash Book
</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selBook" class="FormElem" onchange="GetBookNo()">
	<option value="S">Select Book</option>
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
	<%IF trim(sUserID) = "A" Then %>
			<option value="A" Selected>All</option>
		<%Else%>
			<option value="A">All</option>

		<%
	   End IF
		Dim rsTemp,sqry
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		sqry = "SELECT DISTINCT INTERNALUSERID,LOGINID FROM VwUserUnitList WHERE APPLICATIONCODE = 1  Order By LOGINID "
		'Response.Write "qry="& sqry
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
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Voucher Type</td>
	<td class="FieldCellSub" colspan="4">
	<input type="radio" name="OptVouTy" class="FormElem" value="C,D" <%If sOptVouType = "C,D" then Response.Write "Checked"%>>Both
	<input type="radio" name="OptVouTy" class="FormElem" value="C"  <%If sOptVouType = "C" then Response.Write "Checked"%>>Receipts
	<input type="radio" name="OptVouTy" class="FormElem" value="D"  <%If sOptVouType = "D" then Response.Write "Checked"%>>Payments
	<!--select size="1" name="selVouTy" class="FormElem">
	<%'IF CStr(Trim(sVouTy)) = "C" Then %>
		<option Value="C,D">Both</option>
		<option Value="D">Receipts</option>
		<option Value="C"  Selected>Payments</option>
	<%'Elseif CStr(Trim(sVouTy)) = "D" Then %>
		<option Value="C,D">Both</option>
		<option Value="D"  Selected>Receipts</option>
		<option Value="C">Payments</option>
	<%'Else %>
		<option Value="C,D"  Selected>Both</option>
		<option Value="D">Receipts</option>
		<option Value="C">Payments</option>
	<%'End IF%>
	</select-->
	</td>
</tr>


<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<%IF CStr(sOptType) = "VouNo" Then %>
		<input type="checkbox" value="VouNo" name="ChkVouNo" onclick="Optselection()" >Voucher No. From
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
		<input type="checkbox" value="VouDate" name="ChkVouDt" onclick="OptSelection()" >Voucher Date
	<%Else%>
		<input type="checkbox" value="VouDate" name="ChkVouDt" onclick="OptSelection()" >Voucher Date
	<%End IF %>
	</td>

	<%'Response.Write InsertDatePicker("ctlVouFromDate") %>

    <td class="FieldCellSub" valign="middle">
		<input type="date" id="ctlVouFromDate" name="ctlVouFromDate" class="formelem itms-date-picker" style="width:89px">
	</td>


	 <td class="FieldCell"></td>
	<td class="FieldCellSub">To
	</td>

	<%'Response.Write InsertDatePicker("ctlVouToDate") %>

        <td class="FieldCellSub" valign="middle">
			<input type="date" id="ctlVouToDate" name="ctlVouToDate" class="formelem itms-date-picker" style="width:89px">
		</td>

</tr>

<tr>
	<td class="FieldCell">	</td>
	<td class="FieldCell">
		<%If CStr(sOptType)="VouAmount" then%>
		<input type="checkbox" value="VouAmount" name="ChkVouAmt" onclick="OptSelection()" >	Amount From
	<%else%>
		<input type="checkbox" value="VouAmount" name="ChkVouAmt" onclick="OptSelection()">	Amount From
	<% end if%>
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtFromAmount"  size="20" class="FormElem">
	</td>
	<td class="FieldCellSub">	</td>
	<td class="FieldCellSub">To
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtToAmount"  size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCellSub">
	<%' if CStr(sOptType)="AccHead" then %>
		<!--input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()" checked-->	Account Head
	<%'else%>
		<!--input type="radio" value="AccHead" name="OptCriteria" onclick="OptSelection()"-->
	<%'end if%>
	</td>
	<td class="FieldCellSub" colspan="4">
	<%' if CStr(sOptType)="AccHead" then %>
		<select class="formelem" OnChange="SelectAccHead()" size="1" name="selAccHead">
	<%'Else%>
		<!--select class="formelem" disabled OnChange="SelectAccHead()" size="1" name="selAccHead"-->
	<%'End IF %>
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

<SCRIPT LANGUAGE=vbscript RUNAT=Server>

</SCRIPT>
<%
	Dim iParCode,AccParName,iFrmApplNo,sPurBillType
	iCnt=0
	 'Response.Write "Qry="&sSql
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
		iCrTransNo   = Objrs("createdtransno")
		iFrmApplNo	 = Objrs(8)
		sPurBillType = Objrs(9)
		'Response.Write "AA="& iFrmApplNo
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
		'Response.Write "<BR><BR>"
		'Response.Write "Chk="&Objrs("createdvouchstatus")
	%>
<tr>
<td class="ExcelSerial" align="center" ><%=iCnt%></td>
<td class="ExcelDisplayCell" align="center" width="10" >

<%If sACTN = "P" or sACTN = "U" Then%>
	<input type="checkbox" name="Chkbox" text="<%=Objrs("createdvoucherno")&"&"& Right(Objrs("transactiontype"),1)&"@"& Right(CStr(Objrs("createdvouchstatus")),2) &"@" & trim(AccParName)%>" value="<%=iCrTransNo%>" >
<%Else%>
	<%If Right(CStr(Objrs("createdvouchstatus")),2)="04" then %>
		<input type="checkbox" name="Chkbox" value="<%=iCrTransNo%>" disabled >
	<%else%>
		<input type="checkbox" name="Chkbox" text="<%=Objrs("createdvoucherno")&"&"& Right(Objrs("transactiontype"),1)&"@"& Right(CStr(Objrs("createdvouchstatus")),2) &"@" & trim(AccParName)%>" value="<%=iCrTransNo %>" >
	<%end if%>
<%End IF%>
<input type="hidden" name="hFrmAppNo" value="<%=iFrmApplNo%>">
<input type="hidden" name="hPurBillType" value="<%=sPurBillType%>">
<td class="ExcelDisplayCell" align="left" >
<a href="#" onclick="ShowVouch(<%=iCrTransNo%>); return false;" class="ExcelDisplayLink"><%=Objrs("createdvoucherno") %></a></td>
<td class="ExcelDisplayCell" align="left" ><%=FormatDate(Objrs("voucherdate"))%></td>
<td class="ExcelDisplayCell" align="left" ><%=Objrs("transactiontype")%> </td>

<td class="ExcelDisplayCell" align="left" ><%=AccParName%></td>
<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(Objrs("Voucheramount")) %></td>
	<%

	sSql ="Select CreatedVouchStatus,VoucherNumber from Acc_T_CreatedVoucherHeader H , Acc_T_VoucherHeader v where H.CreatedTransNo=v.CreatedTransNo and " _
	& "right(H.CreatedVouchStatus,2)=04  and H.CreatedTransNo="&iCrTransNo

	'sSql = sSql & "AND A.CREATEDBY = "&sUserID &" "
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

	IF trim(iFrmApplNo) = "2" then
		if Right(CStr(Objrs("createdvouchstatus")),2)="01" then
			Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Created" &"<Font color=Red>*</Font>"&"</td>")
		elseif Right(CStr(Objrs("createdvouchstatus")),2)="04" then
			Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& AccVoucherNo &"<Font color=Red>*</Font>"&"</td>")
		else
			Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Approved" &"<Font color=Red>*</Font>"&"</td>")
		end if
	Else
		if Right(CStr(Objrs("createdvouchstatus")),2)="01" then
			Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Created" &"</td>")
		elseif Right(CStr(Objrs("createdvouchstatus")),2)="04" then
			Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& AccVoucherNo &"</td>")
		else
			Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Approved" &"</td>")
		end if
	End If
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
<td valign="top" align="right" class="FieldCell"><Font color=Red><b>*</b></Font> Vouchers posted from purchase module
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
<td valign="middle" class="ActionCell">
<p align="center">
<%If sACTN = "L" Then%>
<input type="button" value="Edit" name="B9" class="ActionButton" tabindex="3" onclick="ChkforEdit()">
<input type="button" value="Approve" name="B10" class="ActionButton" tabindex="3" onclick="ChkforApprove()">
<input type="button" value="Account" name="btnAcc" class="ActionButton" tabindex="3" onclick="ChkforAccount()">
<input type="button" value="Delete" name="B12" class="ActionButton" tabindex="3" onclick="ChkforDelete()">
<%ElseIf sACTN = "U" Then%>
<input type="button" value="Update" name="B13" class="ActionButton" tabindex="1" onclick="ChkforUpdate()">
<input type="button" value="Regenerate Voucher No" name="B13" class="ActionButtonX" tabindex="1" onclick="RegVoucherNo()">
<%ElseIf sACTN = "P" Then%>
<input type="button" value="Print" name="B14" class="ActionButton" tabindex="1" onclick="ChkforPrint()">
<input type="button" value="Print All" name="B15" class="ActionButton" tabindex="1" onclick="ChkforPrintAll()">
<%ElseIf sACTN = "M" Then%>
<input type="button" value="Edit" name="B15" class="ActionButton" tabindex="1" onclick="ChkforEdit()">
<input type="button" value="Delete" name="B15" class="ActionButton" tabindex="1" onclick="ChkforDelete()">
<input type="button" value="Cancel" name="B15" class="ActionButton" tabindex="1" onclick="ChkforCancel()">
<input type="button" value="Move" name="B15" class="ActionButton" tabindex="1" onclick="ChkforMove()">
<input type="button" value="Reverse" name="B15" class="ActionButton" tabindex="1" onclick="ChkforReverse()">
<%End If%>
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
