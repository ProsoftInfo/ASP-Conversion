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
Dim iAccHead,iParCode,sPayRecName,sOrgName,BookBalCRDR
dim oDOM,Root,objhttp,sChkVal,sTemp,sArr,sVouText,sPartyCode,sParSubType,sPartyName
dim dBkClosing,dDifference,sBkClosingCrDr,dPrnTotal,dOrgBookBal,sClearedOn,iCtr,iInsNo,iPartyCode

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sOrgId = Request("selUnitId")
sPartyCode =Request("hPartyCode")
sParSubType=Request("hParSubType")
sPartyName=Request("hPartyName")
IF trim(sOrgId) = "" then sOrgId = "010101"
sOrgName = Request("hOrgName")
sFromDate = Request("hSelFromDt")
sTodate =Request("hSelToDt")
IF sFromDate = "" then sFromDate = sFromDt
IF sTodate = "" then sTodate = sToDt
dPassBalance = 0
sPassCrDr = Request("optPassBalCrDR")
'Response.Write "sPassCrDr="&sPassCrDr
sTemp = Request("selBook")
sBookName = Request("hBookName")
'Response.Write "BookAccHead="&sBookName
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
oDOM.Load server.MapPath("../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml")

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
%>
</TransDet>
</xml>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
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
		iUnitNo = document.formname.selUnitId.value

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

	'sMonYr = document.formname.SelToMonth.value
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

	document.formname.hSelFromDt.value = document.formname.ctlFrmDate.getDate()
	document.formname.hSelToDt.value = document.formname.ctlToDate.getDate()


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
	document.formname.hOrgName.value = document.formname.selUnitId.options(document.formname.selUnitid.selectedIndex).text
	'IF trim(document.formname.selBook.value) <> "" and trim(document.formname.selBook.value) <> "0" then
		saTemp=Split(document.formname.selBook.value,"~")
		Root.Attributes.GetNamedItem("orgId").value=document.formname.selUnitId.value
		Root.Attributes.GetNamedItem("orgName").value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text

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

Function GetPassBookBal(sMonYr,sObj)
	Dim Objhttp,dRetVal,sRetVal,sBookNo,sTemp

	sTemp = Split(sObj.Value,"~")
	sBookNo = Trim(sTemp(1))
	sMonYr = sMonYr&":"&sBookNo&":"&document.formname.horgId.value
	document.formname.hBookNo.value = sBookNo
	Set Objhttp = CreateObject("Microsoft.XMLHTTP")
	Objhttp.Open "POST","XMLGETPASSBkBAL.asp?MonYr="&sMonYr,False
	Objhttp.send
	sRetVal = Objhttp.responsetext
'	msgbox sRetVal

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
IF document.formname.hSelFromDt.value <> "" then 	document.formname.ctlFrmDate.setdate = document.formname.hSelFromDt.value
IF document.formname.hSelToDt.value <> "" then 	document.formname.ctlToDate.setdate = document.formname.hSelToDt.value

document.formname.selBook.value = document.formname.hBookAccHead.value &"~" & document.formname.hBookNo.value
PayToId.innerHTML = document.formname.hPartyName.value

	IF 	trim(document.formname.hChkVal.value) = "BNR" then
		document.formname.BRType(0).checked = True
		document.formname.BRType(1).checked = False
	Else
		document.formname.BRType(1).checked = True
		document.formname.BRType(0).checked = False
	End IF
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
Function Recon(obj)

	If trim(obj.value) = "BR" then
		document.formname.BRType(0).checked = True
		document.formname.BRType(1).checked = False
		document.formname.hChkVal.value = obj.value

	ElseIf trim(obj.value) = "BNR" then
		document.formname.BRType(0).checked = False
		document.formname.BRType(1).checked = True
		document.formname.hChkVal.value =obj.value

	End IF
document.formname.B6.disabled = false
document.formname.hPartyCode.value = ""
document.formname.hPartyName.value = ""
PayToId.innerHTML = ""
document.formname.submit
End Function
Function PartySelect()
sOrgId = document.formname.horgId.value

   ' OutValue = showModalDialog("PartySelectionAll.asp?orgId="+sOrgId,"","dialogHeight:500px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
   ' arrTemp = split(OutValue,":")

   ' while UBound(arrTemp) = 0
	'    OutValue = showModalDialog("PartySelectionAll.asp?"&OutValue,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'    arrTemp = split(OutValue,":")
   ' wend

   ' sRetValue = OutValue
    'alert(sRetValue)
   ' IF sRetValue = "" Or sRetValue="-1:0" then
'	     exit function
 '   Else
 '
'	    sTemp = split(sRetValue,":")
'
'	    PayToId.innerHTML = sTemp(0)
'	    document.formname.hPartyName.value =  sTemp(0)
'	    document.formname.hPartyCode.value =  sTemp(1)
''	    document.formname.hParSubType.value =  sTemp(2)
  '  End IF
  
        sTempValWindowSize = GetWindowSizeForPopup("2")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
	    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sOrgId,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if
	    
	    if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    PayToId.innerHTML = ndEntry.getAttribute("RetField0")
            	    document.formname.hPartyName.value =  ndEntry.getAttribute("RetField0")
            	    document.formname.hPartyCode.value =  ndEntry.getAttribute("RetField1")
            	    document.formname.hParSubType.value =  ndEntry.getAttribute("RetField4")
		        exit for
                end if
            next
        end if
	    
End Function
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="SetDate()">
<form method="POST" name="formname" >
<input type="hidden" name="horgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="">
<input type="hidden" name="hPartyName" value="<%=sPartyName%>">
<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
<input type="hidden" name="hParSubType" value="<%=sParSubType%>">
<input type="hidden" name="hBookAccHead" value="<%=iBookAccHead%>">
<input type="hidden" name="hFromDt" value="<%=sFromDt%>">
<input type="hidden" name="hToDt" value="<%=sToDt%>">
<input type="hidden" name="hChkVal" value="<%=sChkVal%>">
<input type="hidden" name="hSelFromDt" value="<%=sFromDate%>">
<input type="hidden" name="hSelToDt" value="<%=sTodate%>">

<input type="hidden" name="hMinDt" value="<%=sFromDt%>">
<input type="hidden" name="hMaxDt" value="<%=sToDt%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hBookNo" value="<%=iBookNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Bank Reconciliation</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>

	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
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
								</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
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
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0" width="100%">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;

<input type="CheckBox"  name="BRType" Value="BNR" Onclick="Recon(this)" Checked  > To Reconcile
&nbsp;<input type="CheckBox"  name="BRType" Value="BR" Onclick="Recon(this)" > Reconciled
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
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="125">Organization </td>
                            <td class="FieldCell">
                          <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook()">

							<%populateOrganizationListDB%>
                            </select>
                                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="125">Bank</td>
                            <td class="FieldCell">
                            <select size="1" name="selBook" class="FormElem" onChange="GetPassBookBal(document.formname.ctlFrmDate.GetDate(),this)">
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
                                <tr>

                              </tr>

                            <td class="FieldCell" width="125"> Passbook Balance&nbsp;</td>
                            <td class="FieldCell">
                            <input type="text" name="txtPassBalance" size="15" maxlength="13" style="text-align:right" class="FormElem" value="0.00">
                            <input type="radio" name="optPassBalCrDR" checked value="Cr"> Cr
                            <input type="radio" name="optPassBalCrDR" value="Dr"> Dr
                            </td>
                                </tr>
                                <tr>
                                <td class="FieldCell" width="125">
                                 Received From
                                <td class="FieldCell"><span id="PayToId" class="DataOnly"> </span>
                                &nbsp; <a><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" onclick="PartySelect()"></a>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" >
															<p align="center">
                                                                <input type="button" value="Go" name="B2" class="ActionButton" onClick="FnInit()" >
                                                                <input type="reset" value="Reset" name="B5" class="ActionButton" >
														</td>

													</tr>
													<tr>
														<td class="MiddlePack" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
									<td class="SubTitle"><p align="center"> <%=sVouText%> <%=sBookName%> </td>
                                </tr>
                                    </table>
								</td>
								<%IF trim(sChkVal) <> "BR" then %>
								<!--
								<tr>
									<td align="left" width=5 class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"></td>
									<td align="left" width=250 class="FieldCell">
										<table cellpadding="0" cellspacing="0" width="100%">

											<tr>
												<td class="FieldCell">Cleared On </td>
												<td class="FieldCell" width="160">
												<% '' Function Call to Insert Date Picker
													'Response.Write InsertDatePicker("ctlDate")%>
												</td>
											</tr>
										</table>
									</td>

								</tr>
								-->
								<%End IF 'IF trim(sChkVal) <> "BR" then %>

								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">

                                                <table border="0" cellspacing="1" class="ExcelTable" width="570">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"></td>
                                        <td class="ExcelHeaderCell" align="center">Voucher</td>
                                        <td class="ExcelHeaderCell" align="center" >Instrument</td>
                                        <td class="ExcelHeaderCell" align="center" width="100">Received From</td>
                                        <td class="ExcelHeaderCell" align="center" width="50">Receipt</td>
                                        <td class="ExcelHeaderCell" align="center" width="50" >Payment</td>
                                        <td class="ExcelHeaderCell" align="center" >Cleared On</td>
                                            </tr>

<%
Dim iAccUnitAcctHead
iSno=1
If not objRs.EOF then
	Do While Not objRs.EOF
		sClearedOn = Trim(Objrs(6))
		'Response.Write sClearedOn
		iInsNo = Trim(Objrs(7))

		iPartyCode = 0
		iAccUnitAcctHead  = 0

		sQuery = "Select isNull(AccUnitPartyCode,0),isNull(AccUnitAccountHead,0) from Acc_T_CreatedVoucherDetails where CreatedTransNo= " & iTransNo & ""
		objRs1.Open sQuery,con
		If not objRs1.EOF then
			iPartyCode = objRs1(0)
			iAccUnitAcctHead  = objRs1(1)
		End If
		objRs1.close




		'Response.Write iPartyCode
		IF trim(iPartyCode) <> "0" then
			sQuery = "Select PartyName from App_M_PartyMaster where PartyCode = "&iPartyCode&""
			objRs1.Open sQuery,con
			If not objRs1.EOF then
				sPayRecName=objRs1(0)
			End If
			objRs1.close
		Else
			sQuery = "Select AccountDescription from Acc_M_GLAccountHead where AccountHead = "&iAccUnitAcctHead&""
			objRs1.Open sQuery,con
			If not objRs1.EOF then
				sPayRecName=objRs1(0)
			End If
			objRs1.close
		End IF
		IF Cstr(Objrs("ClearedOn")) = "01/01/1900" then
			dtClrdon = "1/1/1900"
		ElSE
			dtClrdon = Objrs(6)
		End IF

	'Response.Write "sClearedOn="&sClearedOn
	'IF Cstr(sClearedOn) = "1/1/1900" Then
	'IF  Cstr(sClearedOn) <> "01/01/1900" or   Cstr(sClearedOn) <> "1/1/1900" Then
	IF trim(Request.form("hChkVal")) = "BR" then
		IF Cstr(dtClrdon) <> "1/1/1900" then
%>

                                        <tr>
                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
                                        <td class="ExcelFieldCell"><p align="center">
                                        <% 'IF Cstr(sClearedOn) = "01/01/1900" Then %>
                                        	<input type="checkbox" name="chkTransNo<%=iSno%>" value="<%=iTransNo%>"  onclick="setClearOn(this,'<%=iTransNo%>','<%=FormatDate(date())%>','<%=sTransType%>','<%=dAmount%>','<%=iSno%>','<%=iInsNo%>')" class="FormElem"></td>
                                        <%'Else%>
                                        	<!--input type="checkbox" name="chkTransNo" value="<%=iTransNo%>"  onclick="setClearOn(this,'<%=iTransNo%>','<%=FormatDate(date())%>','<%=sTransType%>','<%=dAmount%>')" Checked disabled class="FormElem"></td-->
                                        <%'End IF %>
                                        <td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
                                        <td class="ExcelDisplayCell" align="left"><%=sInstrumentDet%></td>
                                        <td class="ExcelDisplayCell" align="left"><%=sPayRecName%></td>
                                        <td class="ExcelDisplayCell" align="right">
										<%
											if sTransType="R"  then
												Response.Write FormatNumber(dAmount,2,,,0)
												dPrnTotal = CDbl(dPrnTotal) - CDbl(dAmount)
											End IF

										%>
                                        </td>
                                        <td class="ExcelDisplayCell" align="right">
                                        <%if sTransType="P"  then
											Response.Write FormatNumber(dAmount,2,,,0)
											dPrnTotal = CDbl(dPrnTotal) + CDbl(dAmount)
										  End IF
                                        %>
                                        </td>
                                        <td class="ExcelInputCell" align="right" width="10">
                                        <%'   IF Cstr(sClearedOn) = "1/1/1900" or trim(sClearedOn) = "01/01/1900" Then
											IF Cstr(dtClrdon) = "1/1/1900" then%>
                                           		<input type="text" name="txtClearedOn<%=iSno%>Z<%=iTransNo%>" size="11" maxlength="10" readonly class="FormElem">
                                        <%   Else %>
                                           		<input type="text" name="txtClearedOn<%=iSno%>Z<%=iTransNo%>" size="11" maxlength="10" readonly class="FormElem" Value="<%=formatDate(sClearedOn)%>" disabled>
                                        <%   End IF %>
										</td>
	                                           </tr>

<%
			iSno=CInt(iSno)+1
		End IF

	Else 'IF trim(Request.form("hChkVal")) = "BNR" then
		IF Cstr(dtClrdon) = "1/1/1900" then
	%>

	                                        <tr>
	                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
	                                        <td class="ExcelFieldCell"><p align="center">
	                                        <% 'IF Cstr(sClearedOn) = "01/01/1900" Then %>
	                                        	<input type="checkbox" name="chkTransNo<%=iSno%>" value="<%=iTransNo%>"  onclick="setClearOn(this,'<%=iTransNo%>','<%=FormatDate(date())%>','<%=sTransType%>','<%=dAmount%>','<%=iSno%>','<%=iInsNo%>')" class="FormElem"></td>
	                                        <%'Else%>
	                                        	<!--input type="checkbox" name="chkTransNo" value="<%=iTransNo%>"  onclick="setClearOn(this,'<%=iTransNo%>','<%=FormatDate(date())%>','<%=sTransType%>','<%=dAmount%>')" Checked disabled class="FormElem"></td-->
	                                        <%'End IF %>
	                                        <td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
	                                        <td class="ExcelDisplayCell" align="left"><%=sInstrumentDet%></td>
	                                        <td class="ExcelDisplayCell" align="left"><%=sPayRecName%></td>
	                                        <td class="ExcelDisplayCell" align="right">
											<%
												if sTransType="R"  then
													Response.Write FormatNumber(dAmount,2,,,0)
													dPrnTotal = CDbl(dPrnTotal) - CDbl(dAmount)
												End IF

											%>
	                                        </td>
	                                        <td class="ExcelDisplayCell" align="right">
	                                        <%if sTransType="P"  then
												Response.Write FormatNumber(dAmount,2,,,0)
												dPrnTotal = CDbl(dPrnTotal) + CDbl(dAmount)
											  End IF
	                                        %>
	                                        </td>
	                                        <td class="ExcelInputCell" align="right" width="10">
	                                        <%'   IF Cstr(sClearedOn) = "1/1/1900" or trim(sClearedOn) = "01/01/1900" Then
												IF Cstr(dtClrdon) = "1/1/1900" then%>
	                                           		<input type="text" name="txtClearedOn<%=iSno%>Z<%=iTransNo%>" size="11" maxlength="10" readonly class="FormElem">
	                                        <%   Else %>
	                                           		<input type="text" name="txtClearedOn<%=iSno%>Z<%=iTransNo%>" size="11" maxlength="10" readonly class="FormElem" Value="<%=formatDate(sClearedOn)%>" disabled>
	                                        <%   End IF %>
											</td>
		                                           </tr>

	<%
				iSno=CInt(iSno)+1
			End IF
		End IF 'IF trim(Request.form("hChkVal")) = "BR" then
		objRs.MoveNext

	LOOP
end if
objRs.Close

'Response.Write dPrnTotal &" ============== "
%>
<input type="hidden" name="hSno" value="<%=iSno%>">
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
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" align="center">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="MiddlePack" colspan="6"> </td>
                                </tr>
                                <% IF trim(sChkVal) <> "BR" then %>
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
                            <%End IF ' IF trim(sChkVal) <> "BR" then %>
                                <tr>
                            <td class="MiddlePack" colspan="6"> </td>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															<% IF trim(sChkVal) = "BR" then %>
																<input type="button" value="Create Bank Charges" name="B6" class="ActionButtonX" onClick="finalDone()" >
															<% Else %>
																<input type="button" value="Next" name="B6" class="ActionButton" onClick="finalDone()" >
															<% End IF %>
                                                                <input type="button" value="Cancel" name="B2" class="ActionButton" >
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
