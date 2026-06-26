<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BrsBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 27,2003
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
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%

dim sFinPeriod,sFinTemp,sMaxDate,sMinDate,Da,Mo,Yr,sToDt,sFromDt

sFinPeriod = Session("FinPeriod")
'Response.Write sFinPeriod
IF CStr(sFinPeriod) <> "" Then
	sFinTemp = Split(sFinPeriod,":")
	sToDt = "31/03/"&sFinTemp(1)
	sFromDt = "01/04/"&sFinTemp(0)
End IF

IF year(sFromDt) = Year(Date()) then
	Da = Day(Date())
	IF len(Da) = 1 then Da = 0&Da
	Mo = Month(Date())
	IF len(Mo) = 1 then Mo = 0&Mo
	 sToDt = Da&"/"&Mo&"/"&Year(Date())
End IF

dim sOrgId,sBookId,iBookNo,sBookName,sFromDate,sTodate,dPassBalance
dim sBookClosing,sPassCrDr,iBookAccHead,dClosingBal
dim objRs,objRs1,objRs2
dim sQuery,iSno,iTransNo,sVouNo,dAmount,sTransType,sInstrumentDet
Dim iAccHead,iParCode,sPayRecName,sOrgName,BookBalCRDR,sMonthYear
dim oDOM,Root,objhttp,sChkVal,sTemp,sArr,sVouText,sPartyCode,sParSubType,sPartyName
dim dBkClosing,dDifference,sBkClosingCrDr,dPrnTotal,dOrgBookBal,sClearedOn,iCtr,iInsNo,iPartyCode

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sOrgId = Session("organizationcode")
sOrgName = Session("OrgShortName")

sPartyCode =Request("hPartyCode")
sParSubType=Request("hParSubType")
sPartyName=Request("hPartyName")

IF trim(sOrgId) = "" then sOrgId = "010101"

sFromDate = Request("hSelFromDt")
sTodate =Request("hSelToDt")

IF sFromDate = "" then sFromDate = sFromDt
IF sTodate = "" then sTodate = sToDt

'Month year selection
Dim sMonth,sYear,sLastDt
sMonthYear  = request("SelToMonth")

If sMonthYear <> "" Then
	sMonth = Right(sMonthYear,2)
	sYear = left(sMonthYear,4)

	sFromDate  = "01/"&CInt(sMonth)&"/"&sYear

	IF CInt(sMonth) = 4 or CInt(sMonth) = 6 or CInt(sMonth) = 9 or CInt(sMonth) = 11 Then
		sLastDt = "30/"&sMonth&"/"&sYear
	Elseif CInt(sMonth) = 2 Then
		IF CInt(sYear) Mod 4 = 0 Then
			sLastDt = "29/"&sMonth&"/"&sYear
		Else
			sLastDt = "28/"&sMonth&"/"&sYear
		End IF
	Else
		sLastDt = "31/"&sMonth&"/"&sYear
	End IF
	sTodate = sLastDt
End IF
'Response.Write "<p>sMonthYear="&sMonthYear & "=" & sMonth & "=" & sYear
'dPassBalance = 0
dPassBalance = cdbl(Request("txtPassBalance"))
sPassCrDr = Request("optPassBalCrDR")

sTemp = Request("selBook")
sBookName = Request("hBookName")
IF trim(sTemp) <> "" then
	sArr = Split(sTemp,"~")
	iBookAccHead = sArr(0)
	iBookNo=sArr(1)
Else
	iBookAccHead = ""
End IF

sChkVal = Request("hChkVal")
IF sChkVal = "" then sChkVal = "BNR"
'Response.Write sChkVal
'Response.Write sFromDate &"--"&sTodate
IF trim(iBookAccHead) = "" then
	sQuery="select BookNumber,Upper(BookName),isnull(BookAccountHead,0),OtherUnitTransaction from "&_
			"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='02' and BookAccountHead "&_
			"is not null Order By BookName "
		'Response.Write sQuery
		objRs.Open sQuery,con

		IF not objRs.EOF then
			iBookNo = objRs(0)
			sBookName = objRs(1)
			iBookAccHead = objRs(2)
		End If
		objRs.close
End IF
'Response.Write "iBookAccHead="&iBookAccHead&"<BR>"

IF trim(sPassCrDr) = "" then sPassCrDr = "CR"
If trim(sPassCrDr) = "CR" then
	BookBalCRDR = "DR"
Else
	BookBalCRDR = "CR"
End If
Set Root = oDOM.createElement("voucher")
Root.setAttribute "orgId",sOrgId
Root.setAttribute "BookNo",iBookNo
Root.setAttribute "BookName",sBookName
Root.setAttribute "FormDt",sFromDate
Root.setAttribute "ToDt",sTodate
Root.setAttribute "PassBal",dPassBalance
Root.setAttribute "PassBalCRDR",sPassCrDr
Root.setAttribute "BookBal","0"
Root.setAttribute "BookBalCRDR",BookBalCRDR
Root.setAttribute "orgName",sOrgName
Root.setAttribute "AccHead",iBookAccHead
oDOM.AppendChild(Root)


oDom.save Server.MapPath("../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml")
'oDOM.Load server.MapPath("../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml")

set Root=oDOM.documentElement


sOrgId=Root.Attributes.GetNamedItem("orgId").value
iBookNo=Root.Attributes.GetNamedItem("BookNo").value
sBookName= Root.Attributes.GetNamedItem("BookName").value
sFromDate= Root.Attributes.GetNamedItem("FormDt").value
sTodate= Root.Attributes.GetNamedItem("ToDt").value
dPassBalance= Root.Attributes.GetNamedItem("PassBal").value
sPassCrDr= Root.Attributes.GetNamedItem("PassBalCRDR").value
iBookAccHead=Root.Attributes.GetNamedItem("AccHead").value

'Response.Write "sPassCrDr="&sPassCrDr
'Response.write FormatDate(CDate(sTodate)+1)

dBkClosing = GetDayOpening(sOrgId,iBookAccHead,FormatDate(CDate(sTodate)+1))
dOrgBookBal = dBkClosing
dPrnTotal = dBkClosing


if CDbl(dBkClosing)<0 then
	Root.Attributes.GetNamedItem("BookBal").value=CDbl(dBkClosing)*-1
	Root.Attributes.GetNamedItem("BookBalCRDR").value="CR"
	sBkClosingCrDr="CR"
else
	Root.Attributes.GetNamedItem("BookBal").value=CDbl(dBkClosing)
	Root.Attributes.GetNamedItem("BookBalCRDR").value="DR"
	sBkClosingCrDr="DR"
end if

if sPassCrDr="CR" then	dPassBalance=dPassBalance*-1
dDifference=dBkClosing - CDbl(dPassBalance)

if dBkClosing < 0 then dBkClosing=CDbl(dBkClosing)*-1
if sPassCrDr="CR" then	dPassBalance=dPassBalance*-1


oDOM.save server.MapPath("../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml")

'sQuery="select TransactionNumber,VoucherNumber,VoucherAmount,right(TransactionType,1),"&_
'		"rtrim(isNull(BankInstrumentType,''))+' No:'+rtrim(isNull(BankInstrumentNo,''))+' Dt:'+rtrim(convert(char,isNull(BankInstrumentDate,''),103)),"&_
'		"PayToRecdFrom,Convert(Char,isNull(ClearedOn,''),103) from Acc_T_VoucherHeader"&_
'		" where OUDefinitionID='"&sOrgId&"'and BookCode='02' and BookNumber="&iBookNo&" and "&_
'		"VoucherDate between convert(datetime,'"&sFromDate&"',103) and convert(datetime,'"&sTodate&"',103) "&_
'		"and Convert(datetime,isNull(ClearedOn,getDate()),103) > Convert(datetime,'"&sTodate&"',103) "&_
'		"order by VoucherDate "

'If Reconciled -- clearedon field should not be null else null
'Response.Write sPartyCode &"==="

IF trim(sPartyCode) <> "" then

	'sQuery = "Select Distinct CreatedTransNo,CreatedVoucherNo,rtrim(isNull(BankInstrumentType,''))+' No:'+rtrim(isNull(BankInstrumentNo,''))+' Dt:'+rtrim(convert(char,isNull(BankInstrumentDate,''),103)),"&_
	'	 "BankInstrumentType,Right(TransactionType,1),InstrumentAmount,isNull(ClearedOn,0) ClearedOn,BankInstrumentNo,VoucherDate,isNull(AccUnitPartyCode,0),isNull(AccUnitAccountHead,0) from VwInstrumentDetails "&_
	'	 "where AccUnitPartyCode = "&sPartyCode&" and VoucherDate  >= convert(datetime,'"&sFromDate&"',103) and VoucherDate  <= convert(datetime,'"&sTodate&"',103)"

	sQuery = "Select Distinct CreatedTransNo,CreatedVoucherNo,rtrim(isNull(BankInstrumentType,''))+' No:'+rtrim(isNull(BankInstrumentNo,''))+' Dt:'+rtrim(convert(char,isNull(BankInstrumentDate,''),103)),"&_
		 "BankInstrumentType,Right(TransactionType,1),InstrumentAmount,isNull(ClearedOn,0) ClearedOn,BankInstrumentNo,VoucherDate from VwInstrumentDetails "&_
		 "where AccUnitPartyCode = "&sPartyCode&" and VoucherDate  >= convert(datetime,'"&sFromDate&"',103) and VoucherDate  <= convert(datetime,'"&sTodate&"',103)"

	If trim(sChkVal) = "BR" then 'Reconciled
		sVouText = "Vouchers Reconciled for "
		sQuery = sQuery &" and isNull(ClearedOn,0) <> 0 "
	ElseIf trim(sChkVal) = "BNR" then 'To Reconciled
		sVouText = "Vouchers yet to be Reconciled for "
		'sQuery = sQuery &" and isNull(ClearedOn,0) = 0 and TransactionType = 'BAP' and CreatedVouchStatus = '010104' "
		sQuery = sQuery &" and isNull(ClearedOn,0) = 0  and (TransactionType+CreatedVouchStatus = 'BAP010104' Or TransactionType+CreatedVouchStatus = 'BAR010103') "
	End IF
	sQuery = sQuery &" Order by VoucherDate "

'sQuery = sQuery &" and D.AccUnitPartyCode="&sPartyCode&" and D.AccUnitPartySubType= "&sParSubType&" and D.CreatedTransNo = H.CreatedTransNo "
Else

	'sQuery = "Select Distinct CreatedTransNo,CreatedVoucherNo,rtrim(isNull(BankInstrumentType,''))+' No:'+rtrim(isNull(BankInstrumentNo,''))+' Dt:'+rtrim(convert(char,isNull(BankInstrumentDate,''),103)),"&_
	'		 "BankInstrumentType,Right(TransactionType,1),InstrumentAmount,isNull(ClearedOn,0) ClearedOn,BankInstrumentNo,VoucherDate,isNull(AccUnitPartyCode,0),isNull(AccUnitAccountHead,0) from VwInstrumentDetails "&_
	'		 "where VoucherDate >= convert(datetime,'"&sFromDate&"',103) and VoucherDate <= convert(datetime,'"&sTodate&"',103)"

	sQuery = "Select Distinct CreatedTransNo,CreatedVoucherNo,rtrim(isNull(BankInstrumentType,''))+' No:'+rtrim(isNull(BankInstrumentNo,''))+' Dt:'+rtrim(convert(char,isNull(BankInstrumentDate,''),103)),"&_
			 "BankInstrumentType,Right(TransactionType,1),InstrumentAmount,isNull(ClearedOn,0) ClearedOn,BankInstrumentNo,VoucherDate from VwInstrumentDetails "&_
			 "where VoucherDate >= convert(datetime,'"&sFromDate&"',103) and VoucherDate <= convert(datetime,'"&sTodate&"',103)"

	If trim(sChkVal) = "BR" then 'Reconciled
		sVouText = "Vouchers Reconciled for "
		sQuery = sQuery &" and isNull(ClearedOn,0) <> 0 "
	ElseIf trim(sChkVal) = "BNR" then 'To Reconciled
		sVouText = "Vouchers yet to be Reconciled for "
		'sQuery = sQuery &" and isNull(ClearedOn,0) = 0 and TransactionType = 'BAP' and CreatedVouchStatus = '010104'  "
		sQuery = sQuery &" and isNull(ClearedOn,0) = 0  and (TransactionType+CreatedVouchStatus = 'BAP010104' Or TransactionType+CreatedVouchStatus = 'BAR010103') "
	End IF
	sQuery = sQuery &" Order by VoucherDate "
 'Response.Write sQuery &"<br>"
End IF
' Response.Write "<br>" & sQuery &"<br>"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
'Response.Write sQuery

'IF  objRs.EOF then
'	objRs.Close
'	IF trim(sPartyCode) <> "" then
'
'	sQuery = "Select Distinct CreatedTransNo,CreatedVoucherNo,rtrim(isNull(BankInstrumentType,''))+' No:'+rtrim(isNull(BankInstrumentNo,''))+' Dt:'+rtrim(convert(char,isNull(BankInstrumentDate,''),103)),"&_
'			 "BankInstrumentType,Right(TransactionType,1),InstrumentAmount,isNull(ClearedOn,0),BankInstrumentNo,VoucherDate,isNull(AccUnitPartyCode,0),isNull(AccUnitAccountHead,0) from VwInstrumentDetails "&_
'			 "where AccUnitPartyCode = "&sPartyCode&" and VoucherDate  between convert(datetime,'"&sFromDate&"',103) and convert(datetime,'"&sTodate&"',103)"
'
'		If trim(sChkVal) = "BR" then 'Reconciled
'			sVouText = "Vouchers Reconciled for "
'			sQuery = sQuery &" and isNull(ClearedOn,0) <> 0 "
'		ElseIf trim(sChkVal) = "BNR" then 'To Reconciled
'			sVouText = "Vouchers yet to be Reconciled for "
'			sQuery = sQuery &" and isNull(ClearedOn,0) = 0 and TransactionType = 'BAR' and CreatedVouchStatus = '010103' "
'		End IF
'		sQuery = sQuery &" Order by VoucherDate "
'
'	'sQuery = sQuery &" and D.AccUnitPartyCode="&sPartyCode&" and D.AccUnitPartySubType= "&sParSubType&" and D.CreatedTransNo = H.CreatedTransNo "
'	Else
'
'		sQuery = "Select Distinct CreatedTransNo,CreatedVoucherNo,rtrim(isNull(BankInstrumentType,''))+' No:'+rtrim(isNull(BankInstrumentNo,''))+' Dt:'+rtrim(convert(char,isNull(BankInstrumentDate,''),103)),"&_
'				 "BankInstrumentType,Right(TransactionType,1),InstrumentAmount,isNull(ClearedOn,0),BankInstrumentNo,VoucherDate,isNull(AccUnitPartyCode,0),isNull(AccUnitAccountHead,0)  from VwInstrumentDetails "&_
'				 "where VoucherDate between convert(datetime,'"&sFromDate&"',103) and convert(datetime,'"&sTodate&"',103)"
'
'		If trim(sChkVal) = "BR" then 'Reconciled
'			sVouText = "Vouchers Reconciled for "
'			sQuery = sQuery &" and isNull(ClearedOn,0) <> 0 "
'		ElseIf trim(sChkVal) = "BNR" then 'To Reconciled
'			sVouText = "Vouchers yet to be Reconciled for "
'			sQuery = sQuery &" and isNull(ClearedOn,0) = 0 and TransactionType = 'BAR' and CreatedVouchStatus = '010103'"
'		End IF
'		sQuery = sQuery &" Order by VoucherDate "
'
'	End IF
'	with objRs
'		.CursorLocation = 3
'		.CursorType = 3
'		.Source = sQuery
'		.ActiveConnection = con
'		.Open
'	end with
'	set objRs.ActiveConnection = nothing
'End IF
'Response.Write sQuery

Set	iTransNo=objRs(0)
Set	sVouNo=objRs(1)
Set	sInstrumentDet= objRs(2)
'Set sPayRecName=objRs(3)
Set	sTransType=objRs(4)
Set	dAmount=objRs(5)
Set iInsNo=objRs(7)


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML ID="UnitBookData"></XML>
<XML ID="OutData">
<BankRecon orgId="" BookNo="" BookName="" FormDt="" ToDt="" PassBal="" PassBalCRDR="" BookBal="" BookBalCRDR="" orgName="" AccHead=""/>
</XML>
<xml id="VoucherData">
<TransDet>
<%
dim dtClrdon
iCtr = 1
If not objRs.EOF then

	Do While Not objRs.EOF
	'Response.Write "Check="& Objrs("ClearedOn") &"<BR>"
		IF Cstr(Objrs("ClearedOn")) <> "01/01/1900" or  Cstr(Objrs("ClearedOn")) <> "1/1/1900" Then

%>
			<Voucher SlNo="<%=iCtr%>" TransNo="<%=iTransNo%>" Flag="N" VouNo="<%=sVouNo%>" InstruDet="<%=Replace(sInstrumentDet,"&"," ")%>" PayRec="<%=Replace(sPayRecName,"&"," ")%>" TransAmount="<%=dAmount%>" TransType="<%=sTransType%>" ClearedOn="" InsNo="<%=iInsNo%>"/>
<%
		End IF
		iCtr = iCtr  + 1
		objRs.MoveNext
	loop
objRs.MoveFirst
end if
objrs.Close
%>
</TransDet>
</xml>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT language="javascript">
function vd(val)
{

	//init
	var month;
	var day;
	var year;
	var delim = new Array("/","");
	var monthArray = new Array(0,31,29,31,30,31,30,31,31,30,31,30,31);

	dtString = val
	while ((dtString.charAt(0) == " ") && (dtString.length != 0))
	dtString = dtString.substring(1,dtString.length - 1)
	while ((dtString.charAt(dtString.length - 1) == " ") && (dtString.length != 0))
	dtString = dtString.substring(0,dtString.length - 1)
	//get date components
	i = 0; startPos = 0; pos = 0;
	do
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1)return false;

	//get day
	day  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;
	i = 0;
	//get month
	do
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1) return false;
	month  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;

	//get year
	year = parseInt(dtString.substring(startPos,dtString.length),10);
	// valid dateformat check
	if (isNaN(day) || isNaN(month) || isNaN(year)){
		return false;
	}

	// valid month check
	if ((month < 1) || (month > 12)) {
		return false;
	}

	// valid date check
	if ((day < 1) || (day > monthArray[month])) {
		return false;
	}

	// valid year check
	if(year < 1900) {
		return false;
	}
	//check for leap year
	if ((month == 2) && (day == 29))
	if ((((year % 4) == 0) && ((year % 100) != 0)) == false)
	{
		return false;
	}
	return true;

//if we've gotten this far, return true
return true;
} // end function vd

function checkValidDate(val,todaysdate,flag)
{
	//init
	var month;
	var day;
	var year;
	var delim = new Array("/","");
	var monthArray = new Array(0,31,29,31,30,31,30,31,31,30,31,30,31);

	dtString = val
	while ((dtString.charAt(0) == " ") && (dtString.length != 0))
	dtString = dtString.substring(1,dtString.length - 1)
	while ((dtString.charAt(dtString.length - 1) == " ") && (dtString.length != 0))
	dtString = dtString.substring(0,dtString.length - 1)
	//get date components
	i = 0; startPos = 0; pos = 0;
	do
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1)return false;

	//get day
	day  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;
	i = 0;
	//get month
	do
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1) return false;
	month  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;

	//get year
	year = parseInt(dtString.substring(startPos,dtString.length),10);
	// valid dateformat check
	if (isNaN(day) || isNaN(month) || isNaN(year)){
		return false;
	}

	// valid month check
	if ((month < 1) || (month > 12)) {
		return false;
	}

	// valid date check
	if ((day < 1) || (day > monthArray[month])) {
		return false;
	}

	// valid year check
	if(year < 1900) {
		return false;
	}
	//check for leap year
	if ((month == 2) && (day == 29))
	if ((((year % 4) == 0) && ((year % 100) != 0)) == false)
	{
		return false;
	}
	ValidDate = val
	var D1 = day;
	var M1 = month - 1;
	var Y1 = year;

	var D2 = parseInt(todaysdate.substring(0,2),10);
	var M2 = parseInt(todaysdate.substring(3,5),10) - 1;
	var Y2 = parseInt(todaysdate.substring(6,10),10);

	date1 = new Date(Y1,M1,D1);
	date2 = new Date(Y2,M2,D2);

	// Flag - 0 : to check for the given date to be less than todays date
	// Flag - 1 : to check for the given date to be greater than todays date
	if (flag == 0) {
		if (date1 > date2) {
			alert("Date entered should be less than or equal to today's date");
			return false;
		}
		else
			return true;
	}
	if (flag == 1) {
		if (date1 < date2) {
			//alert("Date entered should be greater than or equal to today's date");
			alert("Date entered should be greater than or equal to " + todaysdate);
			return false;
		}
		else
			return true;
	}

} // end function checkValidDate

</SCRIPT>
<SCRIPT language="vbscript">
Function DisplayBook()
dim iUnitNo,arrTemp
dim Root
	'document.formname.selBook.options.length = 1

	'if objUnit.selectedIndex <> "0" then
		'iUnitNo= objUnit(objUnit.selectedIndex).value
		iUnitNo = document.formname.hOrgId.value

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=02&orgID=" & iUnitNo , false
		objhttp.send
		'alert( objhttp.responseXML.xml)
		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml

			Set Root = UnitBookData.documentElement

			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.GetNamedItem("BookName").value
				document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.GetNamedItem("AccHead").value & "~" & HeaderNode.Attributes.GetNamedItem("BookNumber").value
				document.formname.hBookAccHead.value = HeaderNode.Attributes.GetNamedItem("AccHead").value
				document.formname.hBookNo.value =  HeaderNode.Attributes.GetNamedItem("BookNumber").value

			next
		end if
	'end if
end Function

function FnInit()
	Dim sMonYr,sMon,sYr,sToDate,sLastDt

	sMonYr = document.formname.SelToMonth.value
	document.formname.hMonthYr.value = Trim(sMonYr)

	'sMon = Right(sMonYr,2)
	'sYr = Left(sMonYr,4)
	'IF CInt(sMon) = 4 or CInt(sMon) = 6 or CInt(sMon) = 9 or CInt(sMon) = 11 Then
	'	sLastDt = "30/"&sMon&"/"&sYr
	'Elseif CInt(sMon) = 2 Then
	'	IF CInt(sYr) Mod 2 = 0 Then
	'		sLastDt = "29/"&sMon&"/"&sYr
	'	Else
	'		sLastDt = "28/"&sMon&"/"&sYr
	'	End IF
	'Else
	'	sLastDt = "31/"&sMon&"/"&sYr
	'End IF

	'document.formname.hSelFromDt.value = document.formname.ctlFrmDate.getDate()
	'document.formname.hSelToDt.value = document.formname.ctlToDate.getDate()


	'if document.formname.selUnitId.selectedIndex<1 then
	'	MsgBox ("Select Unit")
	'	document.formname.selUnitId.focus
	'	exit function
	'end if
	'if document.formname.selBook.selectedIndex<1 then
	'	MsgBox ("Select Bank ")
	'	document.formname.selBook.focus
	'	exit function
	'end if
	if trim(document.formname.txtPassBalance.value)="" then
		MsgBox ("Enter Passbook Balance")
		document.formname.txtPassBalance.select
		exit function
	ELSEIF	IsNumeric(document.formname.txtPassBalance.value)=false then
		MsgBox ("Enter numeric value in Passbook Balance")
		document.formname.txtPassBalance.select
		exit function
	ELSEIF CDbl(document.formname.txtPassBalance.value)<0 or CDbl(document.formname.txtPassBalance.value)> 9999999999.99 then
		Msgbox("Passbook Balance should be between 0 and 9999999999.99")
		document.formname.txtPassBalance.select
		exit function
	end if

	Set Root = OutData.documentElement
	dim saTemp
	document.formname.hOrgName.value = document.formname.hOrgName.value
	'IF trim(document.formname.selBook.value) <> "" and trim(document.formname.selBook.value) <> "0" then
		saTemp=Split(document.formname.selBook.value,"~")
		Root.Attributes.GetNamedItem("orgId").value=document.formname.hOrgId.value
		Root.Attributes.GetNamedItem("orgName").value=document.formname.hOrgName.value

		Root.Attributes.GetNamedItem("BookNo").value=saTemp(1)
		Root.Attributes.GetNamedItem("AccHead").value=saTemp(0)

		Root.Attributes.GetNamedItem("BookName").value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text
		Root.Attributes.GetNamedItem("FormDt").value=document.formname.hSelFromDt.value
		Root.Attributes.GetNamedItem("ToDt").value=document.formname.hSelToDt.value
		Root.Attributes.GetNamedItem("PassBal").value=document.formname.txtPassBalance.value

		document.formname.hBookNo.value = saTemp(1)
		document.formname.hBookName.value = document.formname.selBook.options(document.formname.selBook.selectedIndex).text
		document.formname.hBookAccHead.value = saTemp(0)

	'End If
	if document.formname.optPassBalCrDR(0).checked then
		Root.Attributes.GetNamedItem("PassBalCRDR").value="CR"
	else
		Root.Attributes.GetNamedItem("PassBalCRDR").value="DR"
	end if


	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=Bank Recon&Mod=BA", false
	objhttp.send OutData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		document.formname.submit()
	end if

End function

'Function GetPassBookBal(sMonYr,sObj)
Function GetPassBookBal(sType,sObj)
	Dim Objhttp,dRetVal,sRetVal,sBookNo,sTemp
	If sType = "S" Then
		sTemp = Split(sObj.Value,"~")
		sBookNo = Trim(sTemp(1))
	Else
		sBookNo = sObj
	End IF

	If sType <> "M" Then
		'sMonYr = Trim(document.formname.hToDt.value)
		If len(Month(document.formname.hToDt.value)) = 1 Then
			sMonYrV=Year(document.formname.hToDt.value) & "0"&Month(document.formname.hToDt.value)
		Else
			sMonYrV=Year(document.formname.hToDt.value) & Month(document.formname.hToDt.value)
		End IF
	Else
		sMonYrV= document.formname.SelToMonth.value
	End IF

	sMonYr = sMonYrV&":"&sBookNo&":"&document.formname.horgId.value

	document.formname.hBookNo.value = sBookNo
	Set Objhttp = CreateObject("Microsoft.XMLHTTP")
	Objhttp.Open "POST","XMLGETPASSBkBAL.asp?MonYr="&sMonYr,False
	Objhttp.send
	sRetVal = Objhttp.responsetext
	'msgbox sRetVal

	IF IsNumeric(sRetVal) Then
		IF CDbl(sRetVal) < 0 Then
			document.formname.optPassBalCrDR(1).checked = True
		Else
			document.formname.optPassBalCrDR(0).checked = True
		End IF
		document.formname.txtPassBalance.value = FormatNumber(Abs(sRetVal),2,,,0)
	End IF

End Function

FUNCTION setClearOn(objTran,iTranNo,sDate,sTransTy,dAmt,iSlno,iInsNo)
Dim dDiffAmt,sTempAmt,dNewDiff,dPass

'alert(iSlno)
IF trim(document.formname.hChkVal.value) <> "BR" then
	dPass = Trim(document.all.dPassBook.innerTEXT)

	sTempAmt = Trim(StrReverse(dPass))
	sTempAmt = Trim(Mid(sTempAmt,3))
	dPass = Trim(StrReverse(sTempAmt))
	dDiffAmt = CDbl(dPass)

	dDiffAmt = document.all.dDiffAmt.innerTEXT

	sTempAmt = Trim(StrReverse(dDiffAmt))
	'MsgBox sTempAmt
	sTempAmt = Trim(Mid(sTempAmt,3))
	sTempAmt = Trim(StrReverse(sTempAmt))
	dDiffAmt = CDbl(sTempAmt)

	'MsgBox dDiffAmt



	IF objTran.checked = True Then
		IF CStr(sTransTy) = "R" Then
			dDiffAmt = CDbl(dDiffAmt) + CDbl(dAmt)
		Else
			dDiffAmt = CDbl(dDiffAmt) - CDbl(dAmt)
		End IF
	Else
		IF CStr(sTransTy) = "R" Then
			dDiffAmt = CDbl(dDiffAmt) - CDbl(dAmt)
		Else
			dDiffAmt = CDbl(dDiffAmt) + CDbl(dAmt)
		End IF

	End IF
	dNewDiff = CDbl(dDiffAmt) - CDbl(dPass)
End IF
Set Root = VoucherData.documentElement
'IF document.formname.hchkVal.value <> "BR" then sDate = document.formname.ctlDate.GetDate()
	if objTran.checked then

		For Each HeaderNode In Root.childNodes
			'if HeaderNode.Attributes.Item(1).nodeValue=iTranNo then
		'	alert HeaderNode.Attributes.Item(9).nodeValue&"="&iInsNo
			'if HeaderNode.Attributes.Item(9).nodeValue=iInsNo then
			if HeaderNode.Attributes.Item(0).nodeValue=iSlno then
				HeaderNode.Attributes.Item(2).nodeValue="Y"
				HeaderNode.Attributes.Item(8).nodeValue=sDate
			end if
		next

		'alert(document.formname.hchkVal.value)
		IF document.formname.hchkVal.value <> "BR" then
			Eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).readOnly=false
			Eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value=""
			Eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).focus()
			IF CDbl(dDiffAmt) < 0 Then
				document.all.dDiffAmt.innerHTML = FormatNumber(dDiffAmt,2,,,0)&" DR"
			Else
				document.all.dDiffAmt.innerHTML = FormatNumber(dDiffAmt,2,,,0)&" CR"
			End IF

			IF CDbl(dNewDiff) < 0 Then
				document.all.dNewDiff.innerHTML = FormatNumber(Abs(dNewDiff),2,,,0)&" DR"
			Else
				document.all.dNewDiff.innerHTML = FormatNumber(Abs(dNewDiff),2,,,0)&" CR"
			End IF
		End IF
	else
		For Each HeaderNode In Root.childNodes
			'if HeaderNode.Attributes.Item(1).nodeValue=iTranNo then
			'if HeaderNode.Attributes.Item(9).nodeValue=iInsNo then
			if HeaderNode.Attributes.Item(0).nodeValue=iSlno then
				HeaderNode.Attributes.Item(2).nodeValue="N"
				HeaderNode.Attributes.Item(8).nodeValue=""
			end if
		next
		IF document.formname.hchkVal.value <> "BR" then
			eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).readOnly=true
			eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value=""

			IF CDbl(dDiffAmt) < 0 Then
				document.all.dDiffAmt.innerHTML = FormatNumber(dDiffAmt,2,,,0)&" DR"
			Else
				document.all.dDiffAmt.innerHTML = FormatNumber(dDiffAmt,2,,,0)&" CR"
			End IF

			IF CDbl(dNewDiff) < 0 Then
				document.all.dNewDiff.innerHTML = FormatNumber(Abs(dNewDiff),2,,,0)&" DR"
			Else
				document.all.dNewDiff.innerHTML = FormatNumber(Abs(dNewDiff),2,,,0)&" CR"
			End IF
		end if 'IF document.formname.hchkVal.value <> "BR" then

	end if 'if objTran.checked then
END FUNCTION

FUNCTION finalDone()
dim iTranNo
iCnt = document.formname.hSno.value - 1

sFlag = 0


For nkk = 1 to iCnt
	IF eval("document.formname.chkTransNo"&nkk).checked = True then
		sFlag = 0
		Exit for
	ElseIF eval("document.formname.chkTransNo"&nkk).checked <> True  then
		sFlag = sFlag + 1
	Else
		sFlag = 0
	End IF
Next
 IF cint(sFlag) <> 0 then
	Alert("Select Voucher No")
	exit function
End IF

	Set Root = VoucherData.documentElement
	'alert Root.xml
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(2).nodeValue="Y" then
			iSlno =HeaderNode.Attributes.Item(0).nodeValue
			iTranNo=HeaderNode.Attributes.Item(1).nodeValue

			'alert(iSlno)
			'alert(iTranNo)
			'alert  trim(eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value)
			if trim(eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value)="" then
				MsgBox("Enter Cleared On Date")
				eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).focus
				exit function
			else

				if not vd(eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value) then
					alert("Enter Valid Date - dd/mm/yyyy format")
					eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).focus
					exit function
				end if

				sTodaysDate = right("0" & trim(day(date())),2) & "/" & right("0" & trim(month(date())),2) & "/" & right(trim(Year(date())),4)
				'alert(sTodaysDate)
				if not checkValidDate(eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value,sTodaysDate,0) then
					eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).focus
					exit function
				end if

				if not checkValidDate(eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value,document.formname.hFromDt.value,1) then
					eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).focus
					exit function
				end if
			end if

			HeaderNode.Attributes.Item(8).nodeValue=eval("document.formname.txtClearedOn"+iSlno+"Z"+iTranNo).value

		end if

	next

	document.formname.B6.disabled = true
	'	alert Root.xml
	set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLUpdate.asp?Name=Bank Recon&Mod=BA", false
		objhttp.send VoucherData.XMLDocument
		if objhttp.responseText <> "" then
			Msgbox(objhttp.responseText)
		else
			document.formname.action="BrsDisplay.asp"
			document.formname.submit()
		end if
END FUNCTION
FUNCTION showCharges()
	document.formname.action="BrsCommEntry.asp"
	document.formname.submit()
END FUNCTION
Function SetDate()
'alert(document.formname.hSelFromDt.value)

'document.formname.hSelFromDt.value
'IF document.formname.hSelFromDt.value <> "" then 	document.formname.ctlFrmDate.setdate = document.formname.hSelFromDt.value
'IF document.formname.hSelToDt.value <> "" then 	document.formname.ctlToDate.setdate = document.formname.hSelToDt.value

document.formname.selBook.value = document.formname.hBookAccHead.value &"~" & document.formname.hBookNo.value
'PayToId.innerHTML = document.formname.hPartyName.value

End Function
Function MinDate()
	sFromDate = document.formname.ctlFrmDate.GetDate
	sToDate = document.formname.ctlToDate.GetDate
	sMinDate = document.formname.hMinDt.value
	sMaxDate = document.formname.hMaxDt.value
	If dateDiff("d",sFromDate,sMinDate) > 0  or  dateDiff("d",sFromDate,sMaxDate) < 0 then
		alert("Date Should be within the Financial Year  "& sMinDate&" to " & sMaxDate )
		document.formname.ctlFrmDate.SetDate =	document.formname.hMinDt.value
		exit Function
	end if
	If dateDiff("d",sToDate,sMinDate) > 0 or datediff("d",sToDate,sMaxDate) < 0  then
		alert("Date Should be within the Financial Year  "& sMinDate &" to " & sMaxDate )
		document.formname.ctlToDate.SetDate = document.formname.hMaxDt.value
		exit Function
	end if
End Function


Function ShowData()
	document.formname.action = "PendingForClearance.asp"
	document.formname.submit
End Function
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="SetDate();GetPassBookBal('L',<%=iBookNo%>)">
<form method="POST" name="formname" >
<input type="hidden" name="horgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookAccHead" value="<%=iBookAccHead%>">
<input type="hidden" name="hFromDt" value="<%=sFromDt%>">
<input type="hidden" name="hToDt" value="<%=sToDt%>">
<input type="hidden" name="hChkVal" value="<%=sChkVal%>">
<input type="hidden" name="hSelFromDt" value="<%=sFromDate%>">
<input type="hidden" name="hSelToDt" value="<%=sTodate%>">

<input type="hidden" name="hMinDt" value="<%=sFromDt%>">
<input type="hidden" name="hMaxDt" value="<%=sToDt%>">
<input type="hidden" name="hMonthYr" value="<%=sMonthYear%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hBookNo" value="<%=iBookNo%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Bank Reconciliation</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>

	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<!--<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<!--tr>
								<td class="TabCurrentCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="100">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher List
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="100">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Bank Charges
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
                                    <tr>
								  		<td align="center">Reconciled List</td>
								  	</tr-->
								<!--</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>-->
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="7px">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>
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
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;

</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%">
<div id="idUnprocessed" style=" display: none">
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="125px">Bank</td>
                            <td class="FieldCell">
                            <!--<select size="1" name="selBook" class="FormElem" onChange="GetPassBookBal(document.formname.ctlFrmDate.GetDate(),this)">-->
                            <select size="1" name="selBook" class="FormElem" onChange="GetPassBookBal('S',this)">
									<!--OPTION value="0">Select a Bank Book</option-->
									<%
									sQuery="select BookNumber,Upper(BookName),isnull(BookAccountHead,0),OtherUnitTransaction from "&_
											"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='02' and BookAccountHead "&_
											"is not null Order By BookName "
										'Response.Write sQuery
										objRs2.Open sQuery,con
									Do while not objRs2.EOF
										%>
										<option value="<%=objRs2(2)%>~<%=objRs2(0)%>"><%=objRs2(1)%></option>
										<%
									objRs2.MoveNext
									loop
									objRs2.Close
									%>
							</select></td>
                                </tr>


								<tr>
									<TD class="FieldCell" valign="top">
									 Upto Month</TD>
									<TD class="FieldCell" valign="top">
									<select size="1" name="SelToMonth" class="FormElem" onChange="GetPassBookBal('M',<%=iBookNo%>)">
										<%

											Dim sFromYear,sToYear,iCounter
											sFromYear=CDbl(Right(getFromFinYear(),4)&Left(getFromFinYear(),2))
											sToYear=CDbl(Right(getToFinYear(),4)&Left(getToFinYear(),2))
											iCounter=sFromYear
											Do While iCounter<=sToYear
												If cint(Month(sTodate)) = cint(Right(iCounter,2)) Then
												%>
													<option value="<%=iCounter%>" selected><%=MonthName (Right(iCounter,2))&"-"&Left(iCounter,4)%></option>
												<%
												Else
												%>
													<option value="<%=iCounter%>"><%=MonthName (Right(iCounter,2))&"-"&Left(iCounter,4)%></option>
												<%
												End IF
												'Response.Write "<option Value="""&iCounter&""">"&MonthName (Right(iCounter,2))&"-"&Left(iCounter,4)&"</option>"

												iCounter=CDbl(iCounter)+1
												IF CDbl(Right(iCounter,2))>12 Then
										 			iCounter=CDbl(CDbl(Left(iCounter,4))+1&"01")
												End IF
											Loop
										%>
									</select>
									</td>


                              <!---<tr>
								<td class="FieldCell" width="125"> Reconciliation From&nbsp;</td>
								<td class="FieldCell" vAlign="top">
									<%' Response.Write InsertDatePicker("ctlFrmDate")%>  <% 'Response.Write InsertDatePicker("ctlToDate")%>
									<object id="ctlFrmDate"  onblur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"         codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
										<param name="_ExtentX" value="2355">
										<param name="_ExtentY" value="529">
									</object>&nbsp;&nbsp; To &nbsp;
									<object id="ctlToDate"  onblur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"         codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
										<param name="_ExtentX" value="2355">
										<param name="_ExtentY" value="529">
									</object>
									</td>
                                <tr>-->

                              </tr>

                            <td class="FieldCell" width="125px"> Passbook Balance&nbsp;</td>
                            <td class="FieldCell">
                            <input type="text" name="txtPassBalance" size="15" maxlength="13" style="text-align:right" class="FormElem" value="" Readonly>
                            <input type="radio" name="optPassBalCrDR" checked value="Cr"> Cr
                            <input type="radio" name="optPassBalCrDR" value="Dr"> Dr
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="ActionCell">
                                                                <input type="button" value="Go" name="B2" class="ActionButton" onClick="FnInit()" >
                                                                <input type="reset" value="Reset" name="B5" class="ActionButton" >
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
                            <tr>
                            <td align="center" class="BottomPack" colspan="3">
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
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>


							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top" width="100%">

                                   <table border="0" cellspacing="1" class="ExcelTable" width="570px">

                                      <tr>
                                        <td class="ExcelHeaderCell" width="270px">Pay To / Received From</td>
                                        <td class="ExcelHeaderCell" width="100px">Debit</td>
										<td class="ExcelHeaderCell" width="100px">Credit</td>
                                      </tr>
                                      <%
										Dim dOpeningBal,sOpenCrDr,sCloseCrDr,dPrnTot,sAccType,sPrintCrDr, sTransTy,sTransAlTy

										'dOpeningBal = GetDayOpening(sOrgId,iBookAccHead,sTodate)
										dOpeningBal = dOrgBookBal
										'Response.Write "<p>dOpeningBal="&dOpeningBal
										'Response.Write "<p><font color=red>sBookName="&sBookName

										IF CDbl(dOpeningBal) < 0 Then
											sPrintCrDr = "CR"
										Else
											sPrintCrDr = "DR"
										End IF

										dPrnTot = FormatNumber(dOpeningBal,2,,,0)

										sQuery = "Select AccountType From Acc_M_BankDetails Where OUDefinitionID = '"&sOrgId&"' and "&_
												 "BookCode = '02' and BookNumber = "&iBookNo&" "

										objRs.Open sQuery,Con
										IF Not objRs.EOF Then
											sAccType = objRs(0)
										Else
											sAccType = "NA"
										End IF
										objRs.Close

										IF CStr(sAccType) = "CU" Then
											sTransTy = "BAR"
											sTransAlTy = "BAP"
										ElseIF CStr(sAccType) = "CC" Then
											sTransTy = "BAP"
											sTransAlTy = "BAR"
										End IF

                                      %>
                                      <tr>
										<td colspan="3" class="ExcelDisplayCell" align="center"><B><%=sBookName%></B></td>
                                      </tr>
                                      <tr>
										<td align="Left" class="ExcelDisplayCell" ><b>Balance as Per Our Books</b></td>
										<%IF CStr(sPrintCrDr) = "DR" Then %>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(Abs(dOpeningBal),2,,,0)%></td>
											<td align="Right" class="ExcelDisplayCell"></td>

										<%Else%>
											<td align="Right" class="ExcelDisplayCell"></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(Abs(dOpeningBal),2,,,0)%></td>
										<%End IF %>

										<!--<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(Abs(dOpeningBal),2,,,0)%></td>-->
									</tr>

									<%
                                    Dim sTransCrDrId,sAccHead,sAccDescription,dRecTotal,dPayTotal,dAddTotal
                                    Dim iBrsTrNo,dLessTotal

									sQuery = "Select PayToRecdFrom,Convert(char,CreatedOn,103),isNull(BankInstrumentType,''),isNull(BankInstrumentNo,''), "&_
												"convert(char,isNull(BankInstrumentDate,''),103),isNull(DrawnOnBank,''),isNull(PayableAt,''),VoucherAmount,CrDrIndication, "&_
												"AccountHead, OUDefinitionID, isNull(BrsTransactionNo,0),Convert(Char,isNull(ClearedOn,''),103) from Acc_T_VoucherHeader Where  "&_
												"BookCode='02' and BookNumber="&iBookNo&" and TransactionType = '"&sTransTy&"' "&_
												"and Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sFromDate&"',103)  "&_
												"and Convert(datetime,VoucherDate,103) <= Convert(datetime,'"&sTodate &"',103)  "&_
												"and Convert(datetime,isNull(ClearedOn,getDate()),103) >  "&_
												"Convert(datetime,'"&sTodate&"',103)   "

									'Response.Write "<p>sQuery="&squery


									with objRs
										.CursorLocation =3
										.CursorType=3
										.ActiveConnection =con
										.Source =sQuery
										.Open
									End With

									set objRs.ActiveConnection =nothing

									If not objRs.EOF then
										set sTransCrDrId	=objRs(8)
										set sAccHead		=objRs(9)
									End IF

									Dim dTotAmount
									Do While Not objrs.EOF

										dAmount	= objRs(7)
										dAddTotal = CDbl(dAmount) + CDbl(dAddTotal)
										dTotAmount= cdbl(dTotAmount)+cdbl(dAmount)
										'Response.Write "<p>sTransCrDrId="&sTransCrDrId
										'Add C Db D Cr
										'Less C Cr D Db
										If sTransCrDrId="D" then
											dRecTotal=cdbl(dRecTotal)+cdbl(dAmount)
											dPrnTot = Cdbl(dPrnTot) - Cdbl(dAmount)
										Else
											dPayTotal=cdbl(dPayTotal)+cdbl(dAmount)
											dPrnTot = Cdbl(dPrnTot) + Cdbl(dAmount)
										End IF

										objrs.MoveNext
									Loop
									objrs.Close
                                      %>
                                    <tr>
										<%IF CStr(sAccType) = "CC" Then %>
											<td align="Left" class="ExcelDisplayCell"><b>Add:Cheque Issued Not Presented For Payment  </b></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dRecTotal,2,,,0)%></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dPayTotal,2,,,0)%></td>
										<%Else%>
											<td align="Left" class="ExcelDisplayCell"><b>Less:Cheque issued Not Presented for Payment  </b></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dRecTotal,2,,,0)%></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dPayTotal,2,,,0)%></td>
										<%End IF%>

									</tr>
									<%
									sQuery = "Select PayToRecdFrom,Convert(char,CreatedOn,103),BankInstrumentType,BankInstrumentNo, "&_
											"convert(char,BankInstrumentDate,103),DrawnOnBank,PayableAt,VoucherAmount,CrDrIndication, "&_
											"isNull(AccountHead,0), OUDefinitionID, isNull(BrsTransactionNo,0),Convert(Char,isNull(ClearedOn,''),103) from Acc_T_VoucherHeader Where  "&_
											"BookCode='02' and BookNumber="&iBookNo&" and TransactionType = '"&sTransAlTy&"' and BrsTransactionNo is NULL "&_
											"and Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sFromDate &"',103)  "&_
											"and Convert(datetime,VoucherDate,103) <= Convert(datetime,'"&sTodate&"',103)  "&_
											"and Convert(datetime,isNull(ClearedOn,getDate()),103) >  "&_
											"Convert(datetime,'"&sTodate&"',103)   "

									with objRs
										.CursorLocation =3
										.CursorType=3
										.ActiveConnection =con
										.Source =sQuery
										.Open
									End With
									set objRs.ActiveConnection =nothing
									If not objRs.EOF then
										set sTransCrDrId	=objRs(8)
									End IF
									dRecTotal ="0"
									dPayTotal = "0"
									While not objRs.EOF
										dAmount			=objRs(7)
										iBrsTrNo		=Trim(objRs(11))
										sAccHead		=objRs(9)

										dLessTotal = CDbl(dLessTotal) + CDbl(dAmount)
										dTotAmount=cdbl(dTotAmount)+cdbl(dAmount)

										If sTransCrDrId="D" then
											dRecTotal=cdbl(dRecTotal)+cdbl(dAmount)
											dPrnTot = Cdbl(dPrnTot) - Cdbl(dAmount)
										Else
											dPayTotal=cdbl(dPayTotal)+cdbl(dAmount)
											dPrnTot = Cdbl(dPrnTot) + Cdbl(dAmount)
										End IF

										objrs.MoveNext
									Wend
									objrs.Close
									%>
									<tr>
										<%IF CStr(sAccType) = "CC" Then %>
											<td align="Left" class="ExcelDisplayCell" ><b>Less:Cheque Presented Not Cleared On Date  </b></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dRecTotal,2,,,0)%></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dPayTotal,2,,,0)%></td>
										<%Else%>
											<td align="Left" class="ExcelDisplayCell" ><b>Add:Cheque Presented Not Cleared On Date  </b></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dRecTotal,2,,,0)%></td>
											<td align="Right" class="ExcelDisplayCell"><%=FormatNumber(dPayTotal,2,,,0)%></td>
										<%End IF%>

									</tr>
									<tr>
										<td class=ExcelDisplayCell align="right" colspan=2>&nbsp;</td>
										<td class=ExcelDisplayCell align="right"></td>
									</tr>
									<tr>
										<td class="ExcelDisplayCell" align="Left">Closing Balance (Our Bank Book)</td>
										<td class="ExcelDisplayCell" align="Right"></td>
										<td class="ExcelDisplayCell" align="Right">
										<%
										if dPrnTotal <0 then
											dPrnTot = Abs(dPrnTot)
											Response.Write Trim(FormatNumber(dPrnTot,2,,,0)) &"&nbsp;CR"
										else
											Response.Write Trim(FormatNumber(dPrnTot,2,,,0)) &"&nbsp;DR"
										end if
										%>
										</td>
									</tr>
									<tr>
										<td class="ExcelDisplayCell" align="Left">Balance as per Pass Book</td>
										<td class="ExcelDisplayCell" align="Right"></td>
										<td class="ExcelDisplayCell" align="Right">
										<%
										Response.Write Trim(FormatNumber(dPassBalance,2,,,0))

										%>
										</td>
									</tr>
									<%
									dDifference = CDbl(dPrnTot) - CDbl(dPassBalance)
									%>
									<tr>
										<td class="ExcelDisplayCell" align="Left">Difference</td>
										<td class="ExcelDisplayCell" align="Right"></td>
										<td class="ExcelDisplayCell" align="Right">
										<%
										if dDifference <0 then
											dDifference=dDifference*-1
											Response.Write FormatNumber(cstr(dDifference),2,,,0) &"&nbsp;CR"
										else
											Response.Write FormatNumber(cstr(dDifference),2,,,0) &"&nbsp;DR"
										end if%>
									</td>
									</tr>
                                </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top" width="100%" align="center">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="MiddlePack" colspan="6"> </td>
                                </tr>
                                <!--<% IF trim(sChkVal) <> "BR" then %>
                                <tr>
                                <td class="FieldCellSub" colspan="2">Balance as Per Our Bank Book </td>
                            <td>
                            <%'=FormatNumber(Abs(dBkClosing),2,,,0)%><%'=sBkClosingCrDr%>&nbsp;
                            <span class="DataOnly">
							<%

										Response.Write Trim(FormatNumber(Abs(dOrgBookBal),2,,,0))


								%>
                            </span>
                            </td>
                            </tr>
                            <tr>

                            <td class="FieldCellSub" width="95">Closing Balance </td>
                            <td>
                            <%'=FormatNumber(dBkClosing,2,,,0)%><%'=sBkClosingCrDr%>&nbsp;
                            <span class="DataOnly" id="dDiffAmt">
							<%
									if dPrnTotal <0 then
										dPrnTotal = Abs(dPrnTotal)
										Response.Write Trim(FormatNumber(dPrnTotal,2,,,0)) &"&nbsp;CR"
									else
										Response.Write Trim(FormatNumber(dPrnTotal,2,,,0)) &"&nbsp;DR"
									end if

								%>
                            </span>
                            </td>
                            <td class="FieldCellSub" width="110">Passbook Balance                            </td>
                            <td> <span class="DataOnly" id="dPassBook"><%=FormatNumber(dPassBalance,2,,,0)%><%=sPassCrDr%>
                            &nbsp;</span>
                                                        </td>
                            <td class="FieldCellSub" width="65">
                            Difference
                            </td>
                            <td>
                              <span class="DataOnly" id="dNewDiff" >
                            <%
									dDifference = CDbl(dPrnTotal) - CDbl(dPassBalance)
									if dDifference <0 then
										dDifference=dDifference*-1
										Response.Write FormatNumber(cstr(dDifference),2,,,0) &"&nbsp;CR"
									else
										Response.Write FormatNumber(cstr(dDifference),2,,,0) &"&nbsp;DR"
									end if
									'if dPrnTotal <0 then
									'	dPrnTotal = Abs(dPrnTotal)
									'	Response.Write FormatNumber(dPrnTotal,2,,,0) &"&nbsp;CR"
									'else
									'	Response.Write FormatNumber(dPrnTotal,2,,,0) &"&nbsp;DR"
									'end if

                            %>
                            &nbsp;
                            </span>
                            </td>
                                </tr>
                            <%End IF ' IF trim(sChkVal) <> "BR" then %>-->
                                <tr>
                            <td class="MiddlePack" colspan="6"> </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5px">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>

							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="ActionCell">
															<% IF trim(sChkVal) = "BR" then %>
																<input type="button" value="Create Bank Charges" name="B6" class="ActionButtonX" onClick="finalDone()" >
															<% Else %>

															<% End IF %>
																<input type="button" value="Pending For Clearence" name="B2" class="ActionButtonX" onclick="ShowData()">
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
							<tr>
<td align="center" class="BottomPack" colspan="3">
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
</BODY>
</HTML>
