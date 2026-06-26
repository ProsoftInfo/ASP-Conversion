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
<XML ID="UnitBookData"><Book/></XML>
<XML ID="OutData"><PartyType/></XML>
<XML id="AccHeadData">
<account/>
</XML>
<XML id="PartyData"><Root></Root></XML>
<XML id="TempXMLData"><Root></Root></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
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

Function Sort(nFieldNo,sOrderByField,sOrder)

	eval("document.formname.hField" +  trim(nFieldNo)).value = trim(sOrderByField) & ":" & trim(sOrder)

	document.formname.hFieldSelected.value = nFieldNo

	document.formname.submit
End Function

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
		MsgBox "From Date must be Between "& sFrmYr  &" and "&sToYr,64,"GJ Vouchers"
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
		MsgBox "To Date must be Between "& sFrmYr  &" and "&sToYr,64,"GJ Vouchers"
		document.formname.ctlVouToDate.setDate=date
		document.formname.ctlVouToDate.focus()
	end if
End Function

Function DisplayBook()
dim iUnitNo,arrTemp,BkCode,iUnitName,iBookVal,iBookNo
dim Root
'-----------Beging of populate partytype
set objhttp = CreateObject("MSXML2.XMLHTTP")


	iUnitNo=document.formname.hUnitNo.value
	'MsgBox iUnitNo
	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentElement
			'Msgbox Root.xml
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
'	if document.formname.selUnitId.selectedIndex <> "0" then
		BkCode= "08"

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
'	end if

'	document.formname.hUnitNo.value=iUnitNo
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
Dim j,sTrans,sMsgNo
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
		document.formname.hTransNo.value="0"&"|"&dTrans
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
		document.formname.action="JournalVoucher.asp?Val="&sTVal
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
	end if
	if not Trim(dTrans)="" then
		'sTVal=dTrans&"~"&sVal(0)&"~"&"A"&"&"&"VouTy="&sVal(1)
		document.formname.hTransNo.value=dTrans
		document.formname.action="AccVouGenerate.asp"
		GetFormDet
		document.formname.btnAcc.disabled = True
		document.formname.submit
	end if
End Function

Function ShowVouch(iCrTransNo)
	showModalDialog "GJVouchView_San.asp?TransNo="&iCrTransNo,"","dialogHeight:390px;dialogWidth:640px;center:Yes;help:No;resizable:No;status:No"
	Exit Function
End Function

Function OptSelection()
if document.formname.optCriteria(0).checked then
	sFlag=document.formname.optCriteria(0).value
	document.formname.txtVouNoFrom.readOnly =false
	document.formname.txtVouNoTo.readOnly =false
	document.formname.txtFromAmount.value =""
	document.formname.txtToAmount.value =""
	document.formname.txtFromAmount.readOnly=True
	document.formname.txtToAmount.readOnly=True
	document.formname.selAccHead.disabled=true
	document.formname.selAccHead.selectedIndex=0
	document.formname.txtAccHead.value=""
Elseif document.formname.optCriteria(1).checked then
	sFlag=document.formname.optCriteria(1).value
	document.formname.txtVouNoFrom.value =""
	document.formname.txtVouNoTo.value =""
	document.formname.txtFromAmount.value =""
	document.formname.txtToAmount.value =""
	document.formname.txtVouNoFrom.readOnly =true
	document.formname.txtVouNoTo.readOnly =true
	document.formname.txtFromAmount.readOnly=true
	document.formname.txtToAmount.readOnly =true
	document.formname.selAccHead.disabled=true
	document.formname.selAccHead.selectedIndex=0
	document.formname.txtAccHead.value=""
Elseif document.formname.optCriteria(2).checked then
	sFlag=document.formname.optCriteria(2).value
	document.formname.txtVouNoFrom.value =""
	document.formname.txtVouNoTo.value =""
	document.formname.txtVouNoFrom.readOnly =true
	document.formname.txtVouNoTo.readOnly =true
	document.formname.txtFromAmount.readOnly=false
	document.formname.txtToAmount.readOnly =false
	document.formname.selAccHead.disabled=true
	document.formname.selAccHead.selectedIndex=0
	document.formname.txtAccHead.value=""
Elseif document.formname.optCriteria(3).checked then
	sFlag=document.formname.optCriteria(3).value
	document.formname.txtVouNoFrom.value =""
	document.formname.txtVouNoTo.value =""
	document.formname.txtFromAmount.value =""
	document.formname.txtToAmount.value =""
	document.formname.txtVouNoFrom.readOnly =true
	document.formname.txtVouNoTo.readOnly =true
	document.formname.txtFromAmount.readOnly=True
	document.formname.txtToAmount.readOnly =True
	document.formname.selAccHead.disabled=false
End if
document.formname.hFlag.value=sFlag
End Function

Function SelectAccHead()
dim iGlHead,sOrgId,sAccHead,arrTemp,sRetVal
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
set objhttp = CreateObject("Microsoft.XMLHTTP")

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
Else
	sOrgId=document.formname.hUnitNo.value
	sBookNo=document.formname.selBook.value
	if 	document.formname.SelAccHead.value="G" then

	    sTempValWindowSize = GetWindowSizeForPopup("5")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)

	    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

	    if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    sRetVal = ndEntry.getAttribute("RetField0")&":"&ndEntry.getAttribute("RetField1")&":"&ndEntry.getAttribute("RetField2")&":"&ndEntry.getAttribute("RetField3")&":"&ndEntry.getAttribute("RetField4")&":"&ndEntry.getAttribute("RetField5")&":"&ndEntry.getAttribute("RetField6")
                end if
            next
        end if


		GetGlHeadXml(sRetVal)

		Set nodAccHead = AccHeadData.documentElement
		if nodAccHead.hasChildNodes then
			For Each HeaderNode In nodAccHead.childNodes
				document.formname.hAccHead.value=HeaderNode.Attributes.getNamedItem("No").Value
				document.formname.txtAccHead.value=HeaderNode.Attributes.getNamedItem("Name").Value
			next
		else
			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
		End if
	else
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
        end if

		objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
		objhttp.send

		IF objhttp.responseText <> "" Then
			sRetVal2 = objhttp.responseText
			GetPartyHeadXml sParCode,sPartyName,sRetVal2
		End IF
		Set nodAccHead = AccHeadData.documentElement

		if nodAccHead.hasChildNodes then
			For Each HeaderNode In nodAccHead.childNodes
				document.formname.hAccHead.value=sPartyType&"?"& HeaderNode.Attributes.getNamedItem("No").Value
				document.formname.txtAccHead.value=HeaderNode.Attributes.getNamedItem("Name").Value
			next
		else
			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
			document.formname.txtAccHead.value=""
		End if
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
'	if document.formname.selUnitId.selectedIndex >0 and document.formname.selBook.selectedIndex<1 then
'		MsgBox "Select a book"
'		document.formname.selBook.focus
'		exit function
'	end if

	if sFlag="VouDate" then
		if dateDiff("d",sFromDate,sToDate)<0 then
			MsgBox "To Date Should be Greater than From Date"
			exit function
		end if
	elseif sFlag ="VouNo" then
		if document.formname.txtVouNoFrom.value="" or document.formname.txtVouNoTo.value="" then
			MsgBox "Voucher No Empty",64
			exit function
		else
			document.formname.hVouFrom.value=document.formname.txtVouNoFrom.value
			document.formname.hVouTo.value=document.formname.txtVouNoTo.value
		end if
	elseif sFlag ="VouAmount" then
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
	document.formname.optCriteria(0).checked=false
	document.formname.optCriteria(1).checked=false
	document.formname.optCriteria(2).checked=false
	document.formname.optCriteria(3).checked=false
	'document.formname.selUnitId.selectedIndex=0
	document.formname.selBook.selectedIndex=0
	document.formname.selAccHead.selectedIndex=0
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

		Call DisplayBook()

	Select case sFlag
		case "VouAmount"
			document.formname.txtFromAmount.value=document.formname.hAmtFrom.value
			document.formname.txtToAmount.value=document.formname.hAmtTo.value
			document.formname.OptCriteria(2).checked = True
		case "VouNo"
			document.formname.txtVouNoFrom.value=document.formname.hVouFrom.value
			document.formname.txtVouNoTo.value=document.formname.hVouTo.value
			document.formname.OptCriteria(0).checked = True
		case "VouDate"
			document.formname.OptCriteria(1).checked = True
		case "AccHead"
			document.formname.OptCriteria(3).checked = True
	end select

	sFDate=document.formname.hFromDate.value
	sTDate=document.formname.hToDate.value


	if Trim(sFDate)<>"" and Trim(sTDate)<>"" then
		document.formname.ctlVouFromDate.setDate=sFDate
		document.formname.ctlVouToDate.setDate=sTDate
	end if

	IF document.formname.selBook.length > 1 Then
		document.formname.selBook.selectedIndex = document.formname.hBookNo.value
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


	IF document.formname.OptCriteria(0).checked = True Then
		sFormVal = sFormVal&"|"&document.formname.OptCriteria(0).Value
	Elseif document.formname.OptCriteria(1).checked = True Then
		sFormVal = sFormVal&"|"&document.formname.OptCriteria(1).Value
	Elseif document.formname.OptCriteria(2).checked = True Then
		sFormVal = sFormVal&"|"&document.formname.OptCriteria(2).Value
	Elseif document.formname.OptCriteria(3).checked = True Then
		sFormVal = sFormVal&"|"&document.formname.OptCriteria(3).Value
	Else
		sFormVal = sFormVal&"|0"
	End IF
	sFormVal = sFormVal&"|"&document.formname.hAccHead.value
	sFormVal = sFormVal&"|"&document.formname.txtAccHead.value
	sFormVal = sFormVal&"|"&document.formname.selUser.value

	document.formname.hFormVal.Value = sFormVal
	'MsgBox document.formname.hFormVal.Value
End Function
Function GetUser()
Dim sUserID
	document.formname.hUserID.value = document.formname.selUser.value

End Function

</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate();DisplayBook();">
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
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
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
			<span style="cursor:hand" onclick="Sort(1,'N','D')">Number</span>
			<%
		else
			%>
			<span style="cursor:hand" onclick="Sort(1,'N','A')">Number</span>
			<%
		end if
	end if
else
	%>
	<span style="cursor:hand" onclick="Sort(1,'N','A')">Number</span>
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
			<span style="cursor:hand" onclick="Sort(2,'D','D')">Date</span>
			<%
		else
			%>
			<span style="cursor:hand" onclick="Sort(2,'D','A')">Date</span>
			<%
		end if
	end if
else
	%>
	<span style="cursor:hand" onclick="Sort(2,'D','A')">Date</span>
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
			<span style="cursor:hand" onclick="Sort(5,'M','D')">Amount</span>
			<%
		else
			%>
			<span style="cursor:hand" onclick="Sort(5,'M','A')">Amount</span>
			<%
		end if
	end if
else
	%>
	<span style="cursor:hand" onclick="Sort(5,'M','A')">Amount</span>
	<%
end if
%>
</td>
<td class="ExcelHeaderCell">Status
</td>
</tr>

<SCRIPT LANGUAGE=vbscript RUNAT=Server>

</SCRIPT>
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
<a href="#" LANGUAGE="VBSCRIPT" onclick="ShowVouch(<%=iCrTransNo%>)" class="ExcelDisplayLink"><%=Objrs("createdvoucherno") %></a></td>
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
