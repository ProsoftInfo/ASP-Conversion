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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
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
<XML ID="UnitBookData"><Book/></XML>
<XML ID="OutData"><PartyType/></XML>
<XML ID="PartyData"><Party/></XML>
<XML id="AccHeadData"><account/></XML>
<XML ID="SearchData" ><Root/></XML>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouTransactions.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
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
		MsgBox "From Date must be Between "& sFrmYr  &" and "&sToYr,64,"Bank Vouchers"
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
		MsgBox "To Date must be Between "& sFrmYr  &" and "&sToYr,64,"Bank Vouchers"
		document.formname.ctlVouToDate.setDate=date
		document.formname.ctlVouToDate.focus()
	end if
End Function

Function DisplayBook()
dim iUnitNo,arrTemp,BkCode,iUnitName,iBookVal,iBookNo
dim Root


'-----------Beginning of populate partytype
set objhttp = CreateObject("MSXML2.XMLHTTP")

	iUnitNo=document.formname.hUnitNo.value
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
'-------------End of populate party type
	document.formname.selBook.options.length = 1
		BkCode= "02"

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode="&BkCode&"&orgID=" & iUnitNo , false
		objhttp.send
'		alert(objhttp.responseText)
		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement

			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value  = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if


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

	'Msgbox document.formname.hFDate.Value
	GetFormDet
	document.formname.submit()
End Function

Function ChkforApprove()
Dim j,sTrans,sMsgNo,sChkForApp
	if not document.formname.hCnt.value="1" then
		for j=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(j).checked then
				sTrans=sTrans&":"&document.formname.chkbox(j).value
				sChkForApp=Split(document.formname.chkbox(j).text,"@")
				'alert sChkForApp(j)
			end if
		next
		sTrans =Mid(sTrans,2)
		IF trim(sChkForApp(3)) = "" then
		MsgBox("Approve is not Possible.Select A/c. Head or Party.")
		Exit function
	End IF
	else
		sTrans=document.formname.chkbox.value
	end if



	if not Trim(sTrans)="" then
		document.formname.hTransNo.value=sTrans
		sMsgNo=MsgBox("Do you want to Approve", vbQuestion + vbOKCancel )
		if sMsgNo=1 then
			document.formname.action="AppVouStatusUpdateAll.asp"
			GetFormDet
			document.formname.submit
		end if
	end if

End Function

Function ChkforDelete()
Dim dTrans,k,sMsgNo
	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"|"&document.formname.chkbox(k).value
			end if
		next
		dTrans=Mid(dTrans,2)
	else
		dTrans=document.formname.chkbox.value
	end if
	if not Trim(dTrans)="" then
		document.formname.hTransNo.value="0|"&dTrans
		sMsgNo=MsgBox("This will Permanently Delete the Voucher(s)" & vbCrLf &"Click OK to Delete", vbQuestion + vbOKCancel )
		if sMsgNo=1 then
			document.formname.action="VouDeletionAll.asp"
			GetFormDet
			document.formname.submit
		end if
	end if
End Function
Function ChkforEdit()
Dim dTrans,k,sTVal,sVal,VouTy,arrTemp
	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"~"&document.formname.chkbox(k).value
				sVal= Split(document.formname.chkbox(k).text,"&")
			end if
		next
		dTrans=Mid(dTrans,2)
		if InStr(1,dTrans,"~")> 0 then
			MsgBox "Edit One by One",64
			exit function
		end if
	else
		dTrans=document.formname.chkbox.value
		sVal= Split(document.formname.chkbox.text,"&")
	end if
	if not Trim(dTrans)="" then
		arrTemp=Split(sVal(1),"@")
		document.formname.hTransNo.value=dTrans
		sTVal=dTrans&"~"&sVal(0)&"~"&"A"&"&"&"VouTy="&arrTemp(0)
		document.formname.action="BankVoucher.asp?Val="&sTVal
		GetFormDet
		document.formname.submit
	end if
End Function

Function ChkforAccount()
Dim dTrans,k,sTVal,sVal,VouTy,sStatus,sFlag
	sFlag=false
	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"~"&document.formname.chkbox(k).value
				sVal= Split(document.formname.chkbox(k).text,"&")
				'alert document.formname.chkbox(k).text
				sStatus= Split(document.formname.chkbox(k).text,"@")
				sFlag=true
			end if
		next
'alert dTrans
		sVouType = split(sStatus(0),"&")
		if sFlag then
			if CStr(sStatus(1))="01"  then
				MsgBox "Only Approved Vouchers can be Accounted",64,"Cash Vouchers"
				exit function
			end if
			If sVouType(1) = "R" then
				if CStr(sStatus(2)) > "0"  then
					MsgBox "Only Bank Reconciliation Vouchers can be Accounted",64,"Cash Vouchers"
					exit function
				end if
			End If
		end if

		dTrans=Mid(dTrans,2)

		if InStr(1,dTrans,"~")> 0 then
			MsgBox "Account One by One",64,"Cash Vouchers"
			exit function
		end if


	else
		dTrans=document.formname.chkbox.value
		sVal= Split(document.formname.chkbox.text,"&")
	end if
	if not Trim(dTrans)="" then
		'sTVal=dTrans&"~"&sVal(0)&"~"&"A"&"&"&"VouTy="&sVal(1)
		document.formname.hTransNo.value=dTrans
		document.formname.action="AccVouGenerate.asp"
		GetFormDet
		document.formname.submit
	end if
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
dim iGlHead,sOrgId,sAccHead,arrTemp,sRetVal
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

set objhttp = CreateObject("Microsoft.XMLHTTP")
Set nodAccHead = AccHeadData.documentElement
Set nodParty = PartyData.documentElement
'If document.formname.selUnitId.selectedIndex="0" then
'	Msgbox "Select Organaisation Id"
'	document.formname.SelAccHead.selectedIndex=0
'	document.formname.selUnitId.focus
'Else

if document.formname.selBook.value="S" Then
		Msgbox "Select Book"
		document.formname.selBook.focus
		document.formname.SelAccHead.selectedIndex=0
		Exit Function
Elseif trim(document.formname.SelAccHead.value) = "0" then
	document.formname.hAccHead.value = ""
	document.formname.txtAccHead.value = ""
	document.formname.hAccIndex.value = ""
	document.formname.hAccTxt.value = ""
	Exit Function

Else
	sOrgId=document.formname.hUnitNo.value
	sBookNo=document.formname.selBook.value

	if 	document.formname.SelAccHead.value="G" then
		document.formname.hAccHead.value = ""
		document.formname.txtAccHead.value = ""

		sTempValWindowSize = GetWindowSizeForPopup("5")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)

        Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo&"&hSelectMode=M",AccHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		sAct = UCase(trim(OutValue.getAttribute("Action")))
		sQuery = trim(OutValue.getAttribute("PassQuery"))
		if ucase(trim(sAct)) <> "CLOSE" then
			do while sAct <> "DONE"
				set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,AccHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
				sAct = UCase(trim(OutValue.getAttribute("Action")))
				if ucase(Trim(sAct)) = "CLOSE" then exit do
				sQuery = trim(OutValue.getAttribute("PassQuery"))
			loop
		end if

		If  OutValue.hasChildNodes Then
			'sExp ="//AccHead"
		    sExp = "//Entry"
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

		sTempValWindowSize = GetWindowSizeForPopup("12")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)

     OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&Party="&sPartyType,"","dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
            arrTemp = Split(Outvalue,":")
	        while UBound(arrTemp)=0
		        OutValue = showModalDialog("../../Common/"&sProgramName&"?"&OutValue,"","dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		        arrTemp = Split(Outvalue,":")
	        wend

            if UBound(arrTemp) <= 1 then
	            document.formname.selAccHead.selectedIndex = 0
	            document.formname.selAccHead.focus()
	            exit function
            End IF

       	If  OutValue<>"" Then
			sRetValue = OutValue
            sTemp = Split(sRetValue,":")
            sParTy = sTemp(4)
            sParSubType = sTemp(3)
            sParCode = sTemp(1)
            sPartyName = sTemp(0)

			sParVal = sParTy &"?"&sParSubType&"?"&sPartyName&"?"&sParCode

			document.formname.hAccHead.value = sParVal
			document.formname.txtAccHead.value =sPartyName
		else

			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
			document.formname.txtAccHead.value=""

		End IF
	end if
End if
		document.formname.hAccIndex.value=document.formname.selAccHead.selectedIndex
		document.formname.hAccTxt.value=document.formname.txtAccHead.value
End Function

Function Validate()
Dim sFromDate,sToDate,objhttp,Root
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	sFromDate=document.formname.ctlVouFromDate.GetDate
	sToDate=document.formname.ctlVouToDate.GetDate
	if document.formname.selBook.value="S" Then
		Msgbox "Select Book",64,"Bank Vouchers"
		document.formname.selBook.focus
		document.formname.SelAccHead.selectedIndex=0
		Exit Function
	end if
	'IF document.formname.hVouDtFlag.value  = "VouDate" then
		if dateDiff("d",sFromDate,sToDate)<0 then
			MsgBox "To Date Should be Greater than From Date"
			exit function
		end if
	'end if
	IF trim(document.formname.ChkVouNo.checked) = "True" then
		'IF document.formname.hVouNoFlag.value  = "VouNo" then
		document.formname.hVouNoFlag.value  = "VouNo"
		if document.formname.txtVouNoFrom.value="" or document.formname.txtVouNoTo.value="" then
			MsgBox "Voucher No Empty",64
			exit function
		else
			document.formname.hVouFrom.value=document.formname.txtVouNoFrom.value
			document.formname.hVouTo.value=document.formname.txtVouNoTo.value
			iVouNoFrom = document.formname.txtVouNoFrom.value
			iVouNoTo   = document.formname.txtVouNoTo.value

		end if
	end if
	IF trim(document.formname.ChkVouAmt.checked) = "True" then
		document.formname.hVouAmtFlag.value  = "VouAmount"
		if document.formname.txtFromAmount.value="" or document.formname.txtToAmount.value="" then
			MsgBox "Voucher Amount Empty",64
			exit function
		else
		  	document.formname.hAmtFrom.value=document.formname.txtFromAmount.value
			document.formname.hAmtTo.value=document.formname.txtToAmount.value
			iVouAmtFrom = document.formname.txtFromAmount.value
			iVouAmtTo = document.formname.txtToAmount.value
		end if
	end if

	IF trim(document.formname.ChkChq.checked) = "True" then
		document.formname.hChqFlag.value  =  "Cheque"

		if document.formname.txtFromChqNo.value = "" or document.formname.txtToChqNo.value = "" Then
			MsgBox "Enter Cheque No From To "
			Exit Function
		elseif Not IsNumeric(document.formname.txtFromChqNo.value) Then
			MsgBox "Enter Only Numeric Value in Cheque No"
			Exit Function
		elseif Not IsNumeric(document.formname.txtToChqNo.value) Then
			MsgBox "Enter Only Numeric Value in Cheque No"
			Exit Function
		Else
			document.formname.hChqFrom.Value = document.formname.txtFromChqNo.value
			document.formname.hChqTo.Value = document.formname.txtToChqNo.value
			iChqFrmNo = document.formname.txtFromChqNo.value
			iChqToNo = document.formname.txtToChqNo.value

		End IF

	end if


		IF document.formname.OptVouTy(0).checked then
			sVouType = "C,D"
		ElseIF document.formname.OptVouTy(1).checked then
			sVouType = "C"
		ElseIF document.formname.OptVouTy(2).checked then
			sVouType = "D"
		End IF
		dtVouForom = sFromDate
		dtVouTo = sToDate
		iAccIndex   =  document.formname.selAccHead.selectedIndex
		'	alert  document.formname.hAccHead.value
		IF document.formname.hAccHead.value <> "0" and document.formname.hAccHead.value <> "" then
			iAccHead = document.formname.hAccHead.value
			sAccHeadName = document.formname.txtAccHead.value
		Else
			iAccHead = "0"
			document.formname.hAccHead.value = "0"
			document.formname.txtAccHead.value =""
		End IF
		Set Root = SearchData.documentElement
		'alert Root.xml
		IF Root.haschildnodes then
			For each node in Root.childnodes
				Set RemNode = node
				Root.removechild RemNode
			Next
		End IF
		sPara = sVouType &":"& dtVouForom &":"& dtVouTo &":"& iVouNoFrom &":"& iVouNoTo &":"& iVouAmtFrom &":"& iVouAmtTo &":"& iAccHeadNo &":"& sAccHeadName  &":"&  iAccIndex &":"& iChqFrmNo &":"& iChqToNo

		objhttp.Open "GET","GetXMLSearchCriteria.asp?ConPara="&sPara&"&Src=BankGrid", false
		objhttp.send
'		alert objhttp.responsetext
		'  alert objhttp.responseXML.xml
		'exit function
		if objhttp.responseXML.xml <> "" then
			SearchData.loadxml objhttp.responseXML.xml
			'For Each HeaderNode In Root.childNodes
		end if

		objhttp.Open "POST","XMLSave.asp?Name=SearchCriteria&Mod=ACC", false
		objhttp.send SearchData.XMLDocument
	document.formname.hFromDate.value=sFromDate
	document.formname.hToDate.value=sToDate
	GetFormDet
	document.formname.submit
End Function
'**************
Function ResetAccHead()
	document.formname.selAccHead.value = "0"
	document.formname.hAccHead.value = ""
	document.formname.txtAccHead.value = ""
End Function
'*********
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

	Function ShowVouch(iCrTransNo)

		showModalDialog "BankVouchView_San.asp?TransNo="&iCrTransNo,"","dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No"
		Exit Function
	End Function

	Function ShowInsDet(iCrTransNo)
		showModalDialog "InstrumentDetView.asp?TransNo="&iCrTransNo,"","dialogHeight:250px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No"
		Exit Function
	End Function

Function SetDate()
	Dim sFDate,sTDate
	sFlag=document.formname.hFlag.value

	Set Root = SearchData.documentElement
	   'alert "Root="&Root.xml
		If Root.haschildnodes then
			IF Root.getAttribute("Src") = "BankGrid" then
				For each node in Root.childnodes

					if trim(node.NodeName) = "VoucherType" then
						document.formname.hVouType.value = node.getAttribute("Value")

					end if
					if trim(node.NodeName) = "VoucherNo" then
						If trim(node.getAttribute("From")) <> "" and trim(node.getAttribute("To")) <> "" then
							document.formname.hVouNoFlag.value  = "VouNo"
							document.formname.hVouFrom.value=node.getAttribute("From")
							document.formname.hVouTo.value=node.getAttribute("To")
						end if
					end if
					if trim(node.NodeName) = "VoucherDate" then

						document.formname.hFromDate.value=node.getAttribute("From")
						document.formname.hToDate.value=node.getAttribute("To")
						document.formname.hVouDtFlag.value  = "VouDate"
					end if
					if trim(node.NodeName) = "VoucherAmount" then
						If trim(node.getAttribute("From")) <> "" and trim(node.getAttribute("To")) <> "" then
							document.formname.hVouAmtFlag.value  = "VouAmount"
							document.formname.hAmtFrom.value=node.getAttribute("From")
							document.formname.hAmtTo.value=node.getAttribute("To")
						end if
					end if
					if trim(node.NodeName) = "ChequeNo" then
						If trim(node.getAttribute("From")) <> "" and trim(node.getAttribute("To")) <> "" then
							document.formname.hChqFlag.value  = "Cheque"
							document.formname.hChqFrom.value=node.getAttribute("From")
							document.formname.hChqTo.value=node.getAttribute("To")
						end if
					end if
					if trim(node.NodeName) = "AccHead" then
						IF node.getAttribute("No") <> "" then
							document.formname.hAccHead.value  = "AccHead"
							document.formname.hAccHead.value = node.getAttribute("No")
							document.formname.hAccIndex.value = node.getattribute("AccIndex")
							document.formname.txtAccHead.value =node.getAttribute("Name")
						End if
					end if
				Next
			End If  'IF Root.getAttribute("Src") = "BankGrid" then
		End IF
	IF trim(document.formname.hVouNoFlag.value)  = "VouNo" then
		document.formname.txtVouNoFrom.value=document.formname.hVouFrom.value
		document.formname.txtVouNoTo.value=document.formname.hVouTo.value
		document.formname.ChkVouNo.checked = True
	End IF
	IF trim(document.formname.hVouAmtFlag.value)  = "VouAmount" then
		document.formname.txtFromAmount.value=document.formname.hAmtFrom.value
		document.formname.txtToAmount.value=document.formname.hAmtTo.value
		document.formname.ChkVouAmt.checked = True
	End IF
	IF trim(document.formname.hVouDtFlag.value) = "VouDate" then
		document.formname.ChkVouDt.checked = True
	End IF

	IF trim(document.formname.hChqFlag.value) = "Cheque" then
		document.formname.ChkChq.checked = True
		document.formname.txtFromChqNo.value = document.formname.hChqFrom.value
		document.formname.txtToChqNo.value = document.formname.hChqTo.value
	End IF
	'if CStr(document.formname.hFromDate.value)=""  then
	'	document.formname.hFromDate.value=(Date)
	'	document.formname.hToDate.value=(date)
	'end if

	Call DisplayBook()
	OptSelection()

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
	Elseif document.formname.ChkChq.checked = True Then
		sFormVal = sFormVal&"|"&document.formname.ChkChq.Value
	Else
		sFormVal = sFormVal&"|0"
	End IF

	sFormVal = sFormVal&"|"&document.formname.hAccHead.value
	sFormVal = sFormVal&"|"&document.formname.txtAccHead.value
	sFormVal = sFormVal&"|"&document.formname.txtFromChqNo.value
	sFormVal = sFormVal&"|"&document.formname.txtToChqNo.value

	document.formname.hFormVal.Value = sFormVal
	'MsgBox document.formname.hFormVal.Value
End Function

Function PrintCheque()
	Dim sTransNo,nBookNo

	if not document.formname.hCnt.value="1" then
		for j=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(j).checked then
				sTransNo=sTransNo & "," & document.formname.chkbox(j).value
				nBookNo = nBookNo & "," & Split(document.formname.chkbox(j).text,"@")(4)
			End IF
		Next
	Else
		sTransNo=document.formname.chkbox.value
	End IF
	If sTransNo = "" Then
		alert("Select any one Instrument No For Printing")
		Exit Function
	Else
		sTransNo = mid(sTransNo,2)
		nBookNo = mid(nBookNo,2)
	End IF
	'alert(sTransNo)

	document.formname.action = "ChequeVoucherView.asp?TransNo="&sTransNo&"&BookNo="&nBookNo
	document.formname.submit
End Function
'************************
Function init()
sFromDate = document.formname.hFromDate.value
   sTodate = document.formname.hToDate.value

   if DateDiff("d",sTodate,date)>0 then
        document.formname.ctlVouFromDate.setdate=sFromDate
        document.formname.ctlVouToDate.setDate=sTodate
   else
        document.formname.ctlVouFromDate.setdate=sFromDate
        document.formname.ctlVouToDate.setDate=date
   end if
End Function
</script>
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
		<object id="ctlVouFromDate"  classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"       codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89px" height="20px" class="FormElem" viewastext>
			<param name="_ExtentX" value="2355">
			<param name="_ExtentY" value="529">
		</object>
	</td>


	<td class="FieldCell"></td>
	<td class="FieldCellSub">To
	</td>

	<%'Response.Write InsertDatePicker("ctlVouToDate") %>

        <td class="FieldCellSub" valign="middle">
			<object id="ctlVouToDate"  classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"       codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89px" height="20px" class="FormElem" viewastext>
				<param name="_ExtentX" value="2355">
				<param name="_ExtentY" value="529">
			</object>
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
				<a href="Javascript:ResetAccHead()"><img border="0" width="11" height="11" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Remove Account Head" ></a>
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

<SCRIPT LANGUAGE=vbscript RUNAT=Server>

</SCRIPT>
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
	<a href="#" Language="VbScript" onclick="ShowInsDet(<%=iCrTransNo%>)" class="ExcelDisplayLink"><%=Objrs(8)%></a></td>
<td class="ExcelDisplayCell" align="right"><%=Objrs(9) %></td>
<td class="ExcelDisplayCell" align="right"><%=FormatNumber(Objrs(10))%></td>
<td class="ExcelDisplayCell" align="lEFT"><%=Objrs(11)%></td>

<td class="ExcelDisplayCell" align="left" >
	<a href="#" LANGUAGE="VBSCRIPT" onclick="ShowVouch(<%=iCrTransNo %>)" class="ExcelDisplayLink"><%=Objrs("createdvoucherno") %></a>
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
