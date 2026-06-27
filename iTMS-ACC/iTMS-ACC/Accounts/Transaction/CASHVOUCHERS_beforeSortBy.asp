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
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouTransactions.js"></SCRIPT>
<SCRIPT ID="oButtonScript" FOR="ctlVouFromDate" EVENT="onBlur()" LANGUAGE="VBSCRIPT">
	CheckFromDate()
</SCRIPT>
<SCRIPT ID="oButtonScript" FOR="ctlVouToDate" EVENT="onBlur()" LANGUAGE="VBSCRIPT">
	CheckToDate()
</SCRIPT>
<Script Language=vbscript>
Dim sFlag,SecT
SecT=false
Dim sPeriod,sMonth,sYear,tMonth,tYear
'****************************
Function ResetAccHead()
	document.formname.selAccHead.value = "0"
	document.formname.hAccHead.value = ""
	document.formname.txtAccHead.value = ""
End Function
'***************************
Function CheckFromDate()
	sPeriod=document.formname.hFinperiod.value
	sMonth=Month(document.formname.ctlVouFromDate.GetDate)
	if Len(sMonth)=1 then
		sMonth="0"&sMonth
	end if
	sYear=Year(document.formname.ctlVouFromDate.GetDate)
	sTemp = split(document.formname.hFinPeriod.value,":")
	sFrmYr = "01/04/"&sTemp(0)
	sToYr = "31/03/"&sTemp(1)

	if sYear&sMonth < Left(sPeriod,4)&"03" or sYear&sMonth >Right(sPeriod,4)&"04" then
		MsgBox "From Date must be Between "& sFrmYr  &" and "&sToYr,64,"Cash Vouchers"
		document.formname.ctlVouFromDate.setDate=date
		document.formname.ctlVouFromDate.focus()
	end if
End Function

Function CheckToDate()
	sPeriod=document.formname.hFinperiod.value
	sMonth=Month(document.formname.ctlVouToDate.GetDate)
	if Len(sMonth)=1 then
		sMonth="0"&sMonth
	end if
	sYear=Year(document.formname.ctlVouToDate.GetDate)
	sTemp = split(document.formname.hFinPeriod.value,":")
	sFrmYr = "01/04/"&sTemp(0)
	sToYr = "31/03/"&sTemp(1)
	if sYear&sMonth < Left(sPeriod,4)&"03" or sYear&sMonth >Right(sPeriod,4)&"04" then
		'MsgBox "To Date must be Between Financial Year",64,"Cash Vouchers"
		MsgBox "To Date must be Between "& sFrmYr  &" and "&sToYr,64,"Cash Vouchers"
		document.formname.ctlVouToDate.setDate=date
		document.formname.ctlVouToDate.focus()
	end if
End Function


Function DisplayBook()
dim iUnitNo,arrTemp,BkCode,iUnitName,iBookVal,iBookNo
dim Root
'-----------Beginning of populate partytype
set objhttp = CreateObject("MSXML2.XMLHTTP")

'if 	document.formname.hUnitNo.selectedIndex<>"0" then
	iUnitNo=document.formname.hUnitNo.value
	'MsgBox iUnitNo
	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentElement
			iCounter=document.formname.SelAccHead.length
			For Each HeaderNode In Root.childNodes
				set oText1 = document.createElement("<Option>" )
					oText1.Text = HeaderNode.text

					oText1.Value = HeaderNode.Attributes.getNamedItem("ParType").Value

				document.formname.selAccHead.add oText1,iCounter
				iCounter=CDbl(iCounter)+1
			next
	end if
'else
'		document.formname.selAccHead.length=2
'end if
'-------------End of populate party type
	document.formname.selBook.options.length = 1
	'if document.formname.selUnitId.selectedIndex <> "0" then
		BkCode= "01"

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode="&BkCode&"&orgID=" & iUnitNo , false
		objhttp.send

		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement

			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value  = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if
	'end if
	'document.formname.hUnitNo.value=iUnitNo
end Function

Function GetBookNo()
	document.formname.hBookNo.value=document.formname.selBook.selectedIndex
	document.formname.hBookVal.value= document.formname.selBook.value
End Function

Function ChkVouType()
dim sVouType,i
	for i=1 to 3
		if document.formname.voutype(i).checked then
			document.formname.voutype(0).checked = False
			sVouType= document.formname.voutype(i).value&document.formname.voutype(i).value
		end if
	next
	GetFormDet
	document.formname.submit()
End Function

Function ChkforApprove()
Dim j,sTrans,sMsgNo,sFrmAppNo
	if not document.formname.hCnt.value="1" then
		for j=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(j).checked then
				sTrans=sTrans&":"&document.formname.chkbox(j).value
				sFrmAppNo= document.formname.hFrmAppNo(j).value
				sPurBillType= document.formname.hPurBillType(k).value
				sChkForApp=Split(document.formname.chkbox(j).text,"@")
			end if
		next
		IF trim(sChkForApp(2)) = "" then
			MsgBox("Approve is not Possible.Select A/c. Head or Party.")
			Exit function
		End IF
		sTrans =Mid(sTrans,2)
	else
		sTrans=document.formname.chkbox.value
	end if

	GetFormDet
	IF trim(sFrmAppNo) = "2" and trim(sPurBillType) <> "M" then
		document.formname.hAppNo.value = sFrmAppNo
		document.formname.action="AppOtherCAView.asp?TransNo="&sTrans&"&sPara=App"
		document.formname.submit
		exit function
	End IF

	'IF trim(sFrmAppNo) = "2" then
	'	MsgBox "Approval Is Not Possible",64,"Cash Vouchers"
	'	exit function
	'End IF



	if not Trim(sTrans)="" then
		document.formname.hTransNo.value=sTrans
		sMsgNo=MsgBox("Do you want to Approve", vbQuestion + vbOKCancel )
		if sMsgNo=1 then
			document.formname.action="AppVouStatusUpdateAll.asp"
			document.formname.submit
		end if
	end if

End Function

Function ChkforDelete()
	
	Dim dTrans,k,sMsgNo,sFrmAppNo
	
	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"|"&document.formname.chkbox(k).value
				sFrmAppNo= document.formname.hFrmAppNo(k).value
				sPurBillType= document.formname.hPurBillType(k).value
			end if
		next
		dTrans=Mid(dTrans,2)
	else
		dTrans=document.formname.chkbox.value
	end if
	if trim(dTrans) = "" then
		alert("Select Voucher")
		exit function
	end if 
	'IF trim(sFrmAppNo) = "2" then
	IF trim(sFrmAppNo) = "2" and trim(sPurBillType) <> "M" then
		MsgBox "Deletion Is Not Possible",64,"Cash Vouchers"
		exit function
	End IF
	if not Trim(dTrans)="" then
		document.formname.hTransNo.value="0"&"|"&dTrans
		sMsgNo=MsgBox("This will Permanently Delete the Voucher(s)" & vbCrLf &"Click OK to Delete", vbQuestion + vbOKCancel,"Cash Vouchers" )
		if sMsgNo=1 then
			document.formname.action="VouDeletionAll.asp"
			GetFormDet
			document.formname.submit
		end if
	end if
End Function

Function ChkforEdit()
Dim dTrans,k,sTVal,sVal,VouTy,arrTemp,sFrmAppNo,sPurBillType
	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"~"&document.formname.chkbox(k).value
				sVal= Split(document.formname.chkbox(k).text,"&")
				sFrmAppNo= document.formname.hFrmAppNo(k).value
				sPurBillType= document.formname.hPurBillType(k).value
			end if
		next
		'alert sFrmAppNo


		dTrans=Mid(dTrans,2)
		if InStr(1,dTrans,"~")> 0 then
			MsgBox "Edit One by One",64,"Cash Vouchers"
			exit function
		end if
	else
		dTrans=document.formname.chkbox.value
		sVal=Split( document.formname.chkbox.text,"&")
	end if
	
	IF trim(sFrmAppNo) = "2" and trim(sPurBillType) <> "M" then
		'MsgBox "Edit Is Not Possible",64,"Cash Vouchers"
		document.formname.hAppNo.value = sFrmAppNo
		document.formname.action="AppOtherCAView.asp?TransNo="&dTrans&"&sPara=Edt"
		document.formname.submit
		exit function
	End IF
	if not Trim(dTrans)="" then
		arrTemp= Split(sVal(1),"@")
		document.formname.hTransNo.value=dTrans
		sTVal=dTrans&"~"&sVal(0)&"~"&"A"&"&"&"VouTy="&arrTemp (0)
		document.formname.action="CashVoucher.asp?Val="&sTVal
		GetFormDet
		document.formname.submit
	end if
		
End Function

Function ChkforAccount()
Dim dTrans,k,sTVal,sVal,VouTy,sStatus,sFlag
	document.formname.btnAcc.disabled = True
	sFlag=false

	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"~"&document.formname.chkbox(k).value
				sVal= Split(document.formname.chkbox(k).text,"&")
				sStatus= Split(document.formname.chkbox(k).text,"@")
				sFrmAppNo= document.formname.hFrmAppNo(k).value
				sPurBillType= document.formname.hPurBillType(k).value
				sFlag=true
			end if
		next

		if sFlag then
			if CStr(sStatus(1))="01"  then
				MsgBox "Only Approved Vouchers can be Accounted",64,"Cash Vouchers"
				document.formname.btnAcc.disabled = False
				exit function
			end if
		end if

		dTrans=Mid(dTrans,2)
		if InStr(1,dTrans,"~")> 0 then
			MsgBox "Account One by One",64,"Cash Vouchers"
			document.formname.btnAcc.disabled = False
			exit function
		end if
	else
		dTrans=document.formname.chkbox.value
		sVal= Split(document.formname.chkbox.text,"&")
		sFrmAppNo= document.formname.hFrmAppNo.value
		sPurBillType= document.formname.hPurBillType.value
	end if
	 'alert dTrans
	'alert sFrmAppNo &"--"&sPurBillType
	'exit function
	GetFormDet

	'IF trim(sFrmAppNo) = "2" then
	IF trim(sFrmAppNo) = "2" and trim(sPurBillType) <> "M" then
		document.formname.hAppNo.value = sFrmAppNo
		document.formname.action="AppOtherCAView.asp?TransNo="&dTrans&"&sPara=Acc"
		document.formname.submit
		exit function
	End IF

	'IF Not CheckCashAcc(dTrans) Then
	'	Exit Function
	'End IF

	if not Trim(dTrans)="" then
		'sTVal=dTrans&"~"&sVal(0)&"~"&"A"&"&"&"VouTy="&sVal(1)
		document.formname.hTransNo.value=dTrans
		document.formname.action="AccVouGenerate.asp"
		GetFormDet
		document.formname.btnAcc.disabled = True
		document.formname.submit
	end if

End Function

Function CheckCashAcc(iTransNo)
	Dim objHttp,sRetVal,sRetTemp,sDispVal,sTemp
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	'Msgbox iTransNo

	objhttp.Open "GET","XMLGetCashStatus.asp?TransNo="&iTransNo , false
	objhttp.send
	sRetVal = objHttp.responseText
	sRetVal = Cstr(Trim(sRetVal))

	'Msgbox StrComp(Cstr(sRetVal),"A:A")

	IF StrComp(Cstr(sRetVal),"A:A") = 0 Then
		CheckCashAcc = True
		Exit Function
	Else
		sRetTemp = Split(Cstr(sRetVal),":")
	End IF

	IF sRetTemp(0) = "R" Then
		MsgBox("Negative Cash Balance Transaction not allowed ")
		CheckCashAcc = False
		Exit Function
	End IF

	IF sRetTemp(1) = "R" Then
		MsgBox("Amount greater than 20,000 Transaction not allowed ")
		CheckCashAcc = False
		Exit Function
	Elseif sRetTemp(1) = "W" Then
		sTemp = MsgBox("Amount greater than 20,000 Transaction, Do you wnat to Continue",4,"Cash Voucher ")
		IF sTemp = 7 Then
			CheckCashAcc = False
			Exit Function
		End IF
	End IF

	CheckCashAcc = True
	Exit Function



End Function

Function ShowVouch(iCrTransNo)
	showModalDialog "CashVouchView_San.asp?TransNo="&iCrTransNo,"","dialogHeight:440px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
	Exit Function
End Function

Function OptSelection()
	if document.formname.ChkVouNo.checked then
		document.formname.hVouNoFlag.value = document.formname.ChkVouNo.value
	else
		document.formname.hVouNoFlag.value = ""
	end if
	if document.formname.ChkVouDt.checked then
		document.formname.hVouDtFlag.value = document.formname.ChkVouDt.value
	else
		document.formname.hVouDtFlag.value = ""
	end if

	if document.formname.ChkVouAmt.checked then
		document.formname.hVouAmtFlag.value = document.formname.ChkVouAmt.value
	else
		document.formname.hVouAmtFlag.value = ""
	end if
	IF trim(document.formname.selAccHead.value) <> "0" then
		sFlag="AccHead"
	End IF
	document.formname.hFlag.value=sFlag
End Function

Function SelectAccHead()
dim iGlHead,sOrgId,sAccHead,arrTemp,sRetVal,nodParty
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
set objhttp = CreateObject("Microsoft.XMLHTTP")
Set nodAccHead = AccHeadData.documentElement
Set nodParty = PartyData.documentElement
'If document.formname.selUnitId.selectedIndex="0" then
'	Msgbox "Select Organaisation Id",64,"Cash Vouchers"
'	document.formname.SelAccHead.selectedIndex=0
'	document.formname.selUnitId.focus
'Else
if document.formname.selBook.value="S" Then
		Msgbox "Select Book",64,"Cash Vouchers"
		document.formname.selBook.focus
		document.formname.SelAccHead.selectedIndex=0
		Exit Function
Else
	sOrgId=document.formname.hUnitNo.value
	sBookNo=document.formname.selBook.value
	if trim(document.formname.SelAccHead.value)="0" then exit function
	if 	document.formname.SelAccHead.value="G" then
		document.formname.hAccHead.value = ""
		document.formname.txtAccHead.value = ""
		Set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?orgid="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo&"&hSelectMode=M",AccHeadData,"dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		'alert OutValue.xml
		sAct = UCase(trim(OutValue.getAttribute("Action")))
		sQuery = trim(OutValue.getAttribute("PassQuery"))
		'   alert  sAct
		if ucase(trim(sAct)) <> "CLOSE" then

			do while sAct <> "DONE"

				Set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,AccHeadData,"dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
				sAct = UCase(trim(OutValue.getAttribute("Action")))
				if ucase(Trim(sAct)) = "CLOSE" then exit do
				sQuery = trim(OutValue.getAttribute("PassQuery"))
			loop
		end if
		' ALERT OutValue.xml
		If  OutValue.hasChildNodes Then
			sExp ="//account/Entry"
			Set AccHeadNode = nodAccHead.Selectnodes(sExp)
			if AccHeadNode.Length > 0 then
				for itr = 0 to AccHeadNode.Length - 1
					document.formname.hAccHead.value = document.formname.hAccHead.value & "," &AccHeadNode.Item(itr).Attributes.getNamedItem("RetField0").value
					document.formname.txtAccHead.value= document.formname.txtAccHead.value & "," &AccHeadNode.Item(itr).Attributes.getNamedItem("RetField5").Value
				next
			end if
			document.formname.hAccHead.value = mid(document.formname.hAccHead.value,2)
			document.formname.txtAccHead.value = mid(document.formname.txtAccHead.value,2)
		else
			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
		End if
	else
		document.formname.hAccHead.value = ""
		document.formname.txtAccHead.value = ""
		sPartyType=document.formname.SelAccHead.value& "?" & document.formname.SelAccHead.options(document.formname.SelAccHead.selectedIndex).text


		Set OutValue = showModalDialog("../../Common/PartySelection.asp?orgid="&sOrgId&"&Party="&sPartyType&"&hSelectMode=M",PartyData,"dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		'alert OutValue.xml
		sAct = UCase(trim(OutValue.getAttribute("Action")))
		sQuery = trim(OutValue.getAttribute("PassQuery"))
		'  alert  sAct
		if ucase(trim(sAct)) <> "CLOSE" then

			do while sAct <> "DONE"

				Set OutValue = showModalDialog("../../Common/PartySelection.asp?"&sQuery,PartyData,"dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
				sAct = UCase(trim(OutValue.getAttribute("Action")))
				if ucase(Trim(sAct)) = "CLOSE" then exit do
				sQuery = trim(OutValue.getAttribute("PassQuery"))
			loop
		end if

		'OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		'arrTemp = split(OutValue,":")

		'while UBound(arrTemp) = 0
		'	OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		'	arrTemp = split(OutValue,":")
		'wend

		'If trim(OutValue) = "-1:0" then exit function
		'sRetValue = OutValue
		'sTemp = Split(sRetValue,":")
		'alert(OutValue.xml)
		If  OutValue.hasChildNodes Then
			sExp ="//Party/Entry"
			Set PartyNode = nodParty.Selectnodes(sExp)
			if PartyNode.Length > 0 then
				for itr = 0 to PartyNode.Length - 1
					sParVal = ""
					sParTy = sParTy &",'"& PartyNode.Item(itr).Attributes.getNamedItem("RetField3").value&"'"
					sParSubType = sParSubType &","&  PartyNode.Item(itr).Attributes.getNamedItem("RetField4").value
					sParCode =sParCode  &","& PartyNode.Item(itr).Attributes.getNamedItem("RetField1").value
					sPartyName = sPartyName  &","& PartyNode.Item(itr).Attributes.getNamedItem("RetField0").value

				Next
				sParVal = mid(sParTy,2) &"?"&mid(sParSubType,2)&"?"&mid(sPartyName,2)&"?"&mid(sParCode,2)
			End IF

			
			document.formname.hAccHead.value = sParVal
			document.formname.txtAccHead.value =mid(sPartyName,2)
			'document.formname.hAccHead.value = mid(document.formname.hAccHead.value,2)
			'document.formname.txtAccHead.value = mid(document.formname.txtAccHead.value,2)
		else

			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
			document.formname.txtAccHead.value=""

		End IF
	'alert "Test="& document.formname.hAccHead.value
	'	objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
	'	objhttp.send

	'	IF objhttp.responseText <> "" Then
	'		sRetVal2 = objhttp.responseText
	'		GetPartyHeadXml sParCode,sPartyName,sRetVal2
	'	End IF
	'	Set nodAccHead = AccHeadData.documentElement

	'	if nodAccHead.hasChildNodes then
	'		For Each HeaderNode In nodAccHead.childNodes
	'			document.formname.hAccHead.value=sPartyType&"?"& HeaderNode.Attributes.getNamedItem("No").Value
	'			document.formname.txtAccHead.value=HeaderNode.Attributes.getNamedItem("Name").Value
	'		next
	'	else
	'		document.formname.SelAccHead.selectedIndex=0
	'		document.formname.hAccHead.value="0"
	'		document.formname.txtAccHead.value=""
	'	End if
	end if
End if
	document.formname.hAccIndex.value=document.formname.selAccHead.selectedIndex
	document.formname.hAccTxt.value=document.formname.txtAccHead.value
End Function

Function Validate()
Dim sFromDate,sToDate

		sFromDate=document.formname.ctlVouFromDate.GetDate
		sToDate=document.formname.ctlVouToDate.GetDate
		sUserID = document.formname.selUser.value
		document.formname.hUserID.value = sUserID
		'alert(sUserID)

	'if document.formname.selUnitId.selectedIndex >0 and document.formname.selBook.selectedIndex<1 then
	'	MsgBox "Select a book",64,"Cash Vouchers"
	'	document.formname.selBook.focus
	'	exit function
	'end if

	IF document.formname.hVouDtFlag.value  = "VouDate" then
		if dateDiff("d",sFromDate,sToDate)<0 then
			MsgBox "To Date Should be Greater than From Date"
			exit function
		end if
	end if
	IF document.formname.hVouNoFlag.value  = "VouNo" then
		if document.formname.txtVouNoFrom.value="" or document.formname.txtVouNoTo.value="" then
			MsgBox "Voucher No Empty",64
			exit function

		else
			document.formname.hVouFrom.value=document.formname.txtVouNoFrom.value
			document.formname.hVouTo.value=document.formname.txtVouNoTo.value
		end if
	end if
	IF document.formname.hVouAmtFlag.value  = "VouAmount" then
		if document.formname.txtFromAmount.value="" or document.formname.txtToAmount.value="" then
			MsgBox "Voucher Amount Empty",64
			exit function
		else
			document.formname.hAmtFrom.value=document.formname.txtFromAmount.value
			document.formname.hAmtTo.value=document.formname.txtToAmount.value
		end if
	end if

	document.formname.hFromDate.value=sFromDate
	document.formname.hToDate.value=sToDate
	GetFormDet
	document.formname.submit
End Function

Function ChkReset()
	document.formname.ctlVouFromDate.setDate=date
	document.formname.ctlVouToDate.setDate=date
	'document.formname.selUnitId.selectedIndex=0
	document.formname.selAccHead.selectedIndex=0
	document.formname.selBook.selectedIndex=0
	document.formname.selAccHead.disabled=true
	document.formname.txtFromAmount.value=""
	document.formname.txtToAmount.value=""
	document.formname.txtVouNoFrom.value=""
	document.formname.txtVouNoTo.value=""
	document.formname.txtAccHead.value=""
	document.formname.hFlag.value=""
	sFlag=""
	DisplayBook()
End Function

Function SetDate()
	Dim sFDate,sTDate
	sFlag=document.formname.hFlag.value

	IF document.formname.hVouNoFlag.value  = "VouNo" then
		document.formname.txtVouNoFrom.value=document.formname.hVouFrom.value
		document.formname.txtVouNoTo.value=document.formname.hVouTo.value
		document.formname.ChkVouNo.checked = True
	End IF
	IF document.formname.hVouAmtFlag.value  = "VouAmount" then
		document.formname.txtFromAmount.value=document.formname.hAmtFrom.value
		document.formname.txtToAmount.value=document.formname.hAmtTo.value
		document.formname.ChkVouAmt.checked = True
	End IF
	IF document.formname.hVouDtFlag.value  = "VouDate" then
		document.formname.ChkVouDt.checked = True
	End IF

	sFDate=document.formname.hFromDate.value
	sTDate=document.formname.hToDate.value

	if Trim(sFDate)<>"" and Trim(sTDate)<>"" then
		document.formname.ctlVouFromDate.setDate=sFDate
		document.formname.ctlVouToDate.setDate=sTDate
	end if
	Call DisplayBook()

	IF document.formname.selBook.length > 1 Then
		document.formname.selBook.selectedIndex = document.formname.hBookNo.value
		document.formname.selAccHead.selectedIndex=document.formname.hAccIndex.value
		document.formname.txtAccHead.value=document.formname.hAccTxt.value
	Else
		document.formname.selBook.selectedIndex = 0
	End IF
End Function

Function GetFormDet()
	Dim sFormVal
	sFormVal = document.formname.hUnitNo.Value
	sFormVal = sFormVal&"|"&document.formname.hUnitNo.value
	sFormVal = sFormVal&"|"&document.formname.selBook.value
	sFormVal = sFormVal&"|"&document.formname.selBook.selectedIndex
	sFormVal = sFormVal&"|"&document.formname.ctlVouFromDate.getDate()
	sFormVal = sFormVal&"|"&document.formname.ctlVouToDate.getDate()
	sFormVal = sFormVal&"|"&document.formname.txtFromAmount.value
	sFormVal = sFormVal&"|"&document.formname.txtToAmount.value
	sFormVal = sFormVal&"|"&document.formname.txtVouNoFrom.value
	sFormVal = sFormVal&"|"&document.formname.txtVouNoTo.value
	sFormVal = sFormVal&"|"&document.formname.hAccIndex.value


	IF document.formname.ChkVouNo.checked = True Then
		sFormVal = sFormVal&"|"&document.formname.ChkVouNo.Value
	Elseif document.formname.ChkVoudt.checked = True Then
		sFormVal = sFormVal&"|"&document.formname.ChkVoudt.Value
	Elseif document.formname.ChkVouAmt.checked = True Then
		sFormVal = sFormVal&"|"&document.formname.ChkVouAmt.Value
	Else
		sFormVal = sFormVal&"|0"
	End IF
	sFormVal = sFormVal&"|"&document.formname.hAccHead.value
	sFormVal = sFormVal&"|"&document.formname.txtAccHead.value
	sFormVal = sFormVal&"|"&document.formname.selUser.value

	IF trim(document.formname.OptVouTy(0).checked) = True then sVouType = document.formname.OptVouTy(0).value
	IF trim(document.formname.OptVouTy(1).checked) = True then sVouType = document.formname.OptVouTy(1).value
	IF trim(document.formname.OptVouTy(2).checked) = True then sVouType = document.formname.OptVouTy(2).value

	'sFormVal = sFormVal&"|"&document.formname.selVouTy.value
	sFormVal = sFormVal&"|"&sVouType
	document.formname.hFormVal.Value = sFormVal
	'alert(sFormVal)
	'MsgBox document.formname.hFormVal.Value
End Function
'Added by Maheshwari on 25th Apr 2007 to get UserID
Function GetUser()
	Dim sUserID
	document.formname.hUserID.value = document.formname.selUser.value
End Function

Function ChkforPrint()
	Dim sTrans
	if not document.formname.hCnt.value="1" then
		for j=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(j).checked then
				sTrans=sTrans&":"&document.formname.chkbox(j).value
			end if
		next
		sTrans =Mid(sTrans,2)
	else
		sTrans=document.formname.chkbox.value
	end if
	'alert(sTrans)
	If sTrans <> "" Then
		sStatus= showModalDialog("PRNCashRecVouView2.asp?Value="&sTrans,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
	Else
		alert("select any one voucher for printing")
		Exit Function
	End IF
End Function

Function RegVoucherNo()
	if document.formname.selBook.value = "S" then
		MsgBox "Select a book"
		'document.formname.selBook.focus
		exit function
	end if
	if document.formname.OptVouTy(1).checked  = false and document.formname.OptVouTy(2).checked  = false then
		MsgBox "Select Voucher Type - Receipts / Payments"
		exit function
	end if
	if document.formname.OptVouTy(1).checked = true then
		document.formname.hVocType.value = "C"
	elseif document.formname.OptVouTy(2).checked = true then
		document.formname.hVocType.value = "D"	
	end if 	
	document.formname.hFromDate.value=document.formname.ctlVouFromDate.GetDate	
	
	document.formname.action = "AccVoucherNo_Generate.asp?BookCode=01"
	document.formname.submit
End Function

</script>
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
<span style="cursor: hand" title="New">
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
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
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
		<object id="ctlVouFromDate"  classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"      codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
			<param name="_ExtentX" value="2355">
			<param name="_ExtentY" value="529">
		</object>
	</td>


	 <td class="FieldCell"></td>
	<td class="FieldCellSub">To
	</td>

	<%'Response.Write InsertDatePicker("ctlVouToDate") %>

        <td class="FieldCellSub" valign="middle">
			<object id="ctlVouToDate"  classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"      codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
				<param name="_ExtentX" value="2355">
				<param name="_ExtentY" value="529">
			</object>
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
<a href="#" LANGUAGE="VBSCRIPT" onclick="ShowVouch(<%=iCrTransNo%>)" class="ExcelDisplayLink"><%=Objrs("createdvoucherno") %></a></td>
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
