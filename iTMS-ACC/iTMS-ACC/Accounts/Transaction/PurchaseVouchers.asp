<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PurchaseVouchers.asp
	'Module Name				:	ACCOUNTS (Transcation)
    'Author Name				:	Ragavendran R
	'Created On					:	Jan 28,2011
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

	Dim sFinPeriod,Objrs,Objrs1,iCnt,sSql,iCrTransNo
	Dim dcrs,AccVoucherNo
	Dim sUnitID,sBookNo,sVouFrm,sAction
	Dim sFrmDate,sToDate,sFrmAmt,sToAmt
	Dim sFrmNo,sToNo,iAccHead,sSelectTransType,sUserID
	Dim sParsubType,iParCode,sXmlAccFlag,sXmlParFlag,sVouAccFlag
	Dim sVouNoFlag,sVouDtFlag,sVouAmtFlag
	Dim sPurTypeValue,sOrgName
	Dim sField1,sField2,sField3,sField4,sField5,sSortBy,Arr1,nFieldSelected

	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

    Dim sValTemp2,sFinFromDate,sFinTodate

	Const sGridName ="PurchaseGrid"
	Const iPageSize=20


	Dim oDOM,objfs,Root,node,sDataExistInXML

	Set oDOM = CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set Objrs=server.CreateObject ("ADODB.recordset")
	Set Objrs1=server.CreateObject ("ADODB.recordset")

	iCurrentPage=CInt(Request.Form("hPageSelection"))
    sUnitID = Session("organizationcode")
    sOrgName = Session("OrgShortName")
	sFinPeriod=session("finperiod")

    sAction = Request.QueryString("ACTN")

	sXmlAccFlag = Request("hXmlAccFlag")
	sXmlParFlag = Request("hXmlParFlag")

	sDataExistInXML = False

	IF  objfs.FileExists(Server.MapPath("../temp/transaction/SearchCriteria_ACC_"&session.SessionID&".xml")) then
		oDOM.load  server.MapPath("../temp/transaction/SearchCriteria_ACC_"&session.SessionID&".xml")

		Set Root = oDOM.documentElement

		IF trim(Root.getAttribute("Src")) = sGridName then

			sDataExistInXML = true
			sSelectTransType = Root.getAttribute("TransType")

			if Root.haschildnodes then



				For each node in Root.childnodes

					if trim(node.NodeName) = "SearchOption" then

						sUnitID = node.getAttribute("SelUnitId")
						sBookNo = node.getAttribute("SelBook")
						sUserID = node.getAttribute("SelUser")
						sPurTypeValue = node.getAttribute("SelPurType")
						sVouFrm = node.getAttribute("SelVouFrom")
					end if

					if trim(node.NodeName) = "VoucherNo" then
						sFrmNo = node.getAttribute("From")
						sToNo = node.getAttribute("To")
						If trim(sFrmNo) <> "" and trim(sToNo) <> "" then sVouNoFlag = "Y"
					end if
					if trim(node.NodeName) = "VoucherDate" then
						sFrmDate = node.getAttribute("From")
						sToDate  = node.getAttribute("To")
						If trim(sFrmDate) <> "" and trim(sToDate) <> "" then sVouDtFlag =	"Y"
					end if
					if trim(node.NodeName) = "VoucherAmount" then
						sFrmAmt = node.getAttribute("From")
						sToAmt  = node.getAttribute("To")
						If trim(sFrmAmt) <> "" and trim(sToAmt) <> "" then sVouAmtFlag =	"Y"
					end if
					if trim(node.NodeName) = "AccHead" then
						iAccHead = node.getattribute("No")
						sParsubType = node.getattribute("ParSubTypeValue")

						IF  ( trim(iAccHead) <> "" and trim(iAccHead) <> "0" ) then
							sVouAccFlag = True
							iParCode  = iAccHead
						end if

						if ( trim(sParsubType) <> "" and trim(sParsubType) <> "0" ) then
							sVouAccFlag = True
						end if

					end if
				Next
			end if 'if Root.haschildnodes then
		End IF 'IF trim(Root.getAttribute("Src")) = sGridName then
	End IF

	 if trim(sAction)="A" then
	    if trim(sSelectTransType) ="" then sSelectTransType = "P"
	 else
	    if trim(sSelectTransType) ="" then sSelectTransType = "A"
	 end if

	 if trim(sVouFrm) = "" then sVouFrm = "P"

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
				sSortBy = "A.CREATEDVOUCHERNO"
			else
				sSortBy = "A.CREATEDVOUCHERNO desc"
			end if
		end if
	end if


	if trim(sField2) <> ""  then
		if instr(1,sField2,":") > 0 then
			Arr1 = Split(sField2,":")

			if Arr1(1) = "A" then
				sSortBy = "A.VOUCHERDATE"
			else
				sSortBy = "A.VOUCHERDATE desc "
			end if
		end if
	end if

	'if trim(sField3) <> ""  then
	'	if instr(1,sField3,":") > 0 then
	'		Arr1 = Split(sField3,":")
	'
	'		if Arr1(1) = "A" then
	'			sSortBy = "A.CreatedVouchStatus"
	'		else
	'			sSortBy = "A.CreatedVouchStatus desc "
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
				sSortBy = "A.VOUCHERAMOUNT"
			else
				sSortBy = "A.VOUCHERAMOUNT desc "
			end if
		end if
	end if



sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFromDate = "01/04/"& sValTemp2(0)
sFinToDate = "31/03/"&sValTemp2(1)
if Trim(sFrmDate)="" then
    sFrmDate = sFinFromDate
    sToDate = sFinToDate
end if
if DateDiff("d",sToDate,date)<0 then
    sToDate = date
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
<% IF trim(sXmlAccFlag) <> "True" then %>
	<XML id="AccHeadData"><AccHead/></XML>
<% Else %>
	<XML id="AccHeadData" src="<%="../temp/transaction/PartySubType_PUR_"&Session.SessionID&".xml"%>"></XML>
<% End IF %>
<% IF trim(sXmlParFlag) <> "True" then %>
	<XML id="PartyData"><PARTY/></XML>
<% Else %>
	<XML id="PartyData" src="<%="../temp/transaction/PartyType_PUR_"&Session.SessionID&".xml"%>"></XML>
<% End IF %>


<% IF trim(sDataExistInXML)  then %>
	<XML id="SearchData" src="<%="../temp/transaction/SearchCriteria_ACC_"&Session.SessionID&".xml"%>"></XML>
<% Else %>
	<XML ID="SearchData" ><Root/></XML>
<% End IF %>

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

<Script Language="vbscript">
Dim SecT
SecT=false
Dim sPeriod,sMonth,sYear,tMonth,tYear

Function Sort(nFieldNo,sOrderByField,sOrder)

	eval("document.formname.hField" +  trim(nFieldNo)).value = trim(sOrderByField) & ":" & trim(sOrder)

	document.formname.hFieldSelected.value = nFieldNo

	document.formname.submit
End Function


Function ChangeStatusOfInputFields(obj)
	Dim sDisabled
	if obj.checked then
		sDisabled = false
	else
		sDisabled = true
	end if

	if obj.name = "ChkVouNo" then
		document.formname.txtVouNoFrom.disabled = sDisabled
		document.formname.txtVouNoTo.disabled  = sDisabled
	end if

	if obj.name = "ChkVouAmt" then
		document.formname.txtFromAmount.disabled = sDisabled
		document.formname.txtToAmount.disabled  = sDisabled
	end if

End Function

Function PaginateAcc(iPageNo)
	document.formname.hPageSelection.value = iPageNo
	document.formname.action="PurchaseVouchers.asp?ACTN="& document.formname.hAction.value
	document.formname.submit
end Function

Function CheckFromDate()
	sPeriod=document.formname.hFinperiod.value
	sMonth=Month(document.formname.ctlVouFromDate.GetDate)
	if Len(sMonth)= 1 then
		sMonth="0"&sMonth
	end if
	sYear=Year(document.formname.ctlVouFromDate.GetDate)
	sTemp = split(document.formname.hFinPeriod.value,":")
	sFrmYr = "01/04/"&sTemp(0)
	sToYr = "31/03/"&sTemp(1)
	if sYear&sMonth < Left(sPeriod,4)&"03" or sYear&sMonth >Right(sPeriod,4)&"04" then
		MsgBox "From Date must be Between "& sFrmYr  &" and "&sToYr,64,"Purchase Vouchers"
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
		MsgBox "To Date must be Between "& sFrmYr  &" and "&sToYr,64,"Purchase Vouchers"
		document.formname.ctlVouToDate.setDate=date
		document.formname.ctlVouToDate.focus()
	end if
End Function

Function DisplayBook()
dim iUnitNo,arrTemp,BkCode,iUnitName,iBookVal,iBookNo
dim Root
	iUnitNo=document.formname.hOrgID.value
'-----------Beginning of populate partytype
			set objhttp = CreateObject("MSXML2.XMLHTTP")

		'	if 	document.formname.selUnitId.selectedIndex<>"0" then
		'
		'		objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
		'		objhttp.send

		'		if objhttp.responseXML.xml <> "" then
		'				OutData.loadXML objhttp.responseXML.xml
		'				Set Root = OutData.documentElement
		'				iCounter=document.formname.SelAccHead.length
		'				For Each HeaderNode In Root.childNodes
		'					set oText1 = document.createElement("<Option>" )
		'						oText1.Text = HeaderNode.text
		'						oText1.Value = HeaderNode.Attributes.getNamedItem("ParType").Value
		'					document.formname.selAccHead.add oText1,iCounter
		'					iCounter=CDbl(iCounter)+1
		'				next
		'		end if
		'	else
		'			document.formname.selAccHead.length=2
		'	end if
'-------------End of populate party type
	document.formname.selBook.options.length = 1
	'if document.formname.selUnitId.selectedIndex <> "0" then
		BkCode= "04"

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode="&BkCode&"&orgID=" & iUnitNo , false
		objhttp.send
		'alert objhttp.responsetext
		'alert objhttp.responsexml.xml
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

end Function



Function ChkVouType()
dim sTransType,i

    if trim(document.formname.hAction.value)="L" then
	    for i=1 to 3
		    if document.formname.voutype(i).checked then
			    sTransType= sTransType & "," & document.formname.voutype(i).value
		    end if
	    next
	else
	    for i=0 to 1
		    if document.formname.voutype(i).checked then
			    sTransType= sTransType & "," & document.formname.voutype(i).value
		    end if
	    next
	end if 'if trim(document.formname.hAction.value)="L" then

	if trim(sTransType) <> "" then
		sTransType = mid(sTransType,2)
	end if

	Set Root = SearchData.documentElement
	Root.setAttribute "Src",document.formname.hGridName.value
	Root.setAttribute "TransType",sTransType


	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=SearchCriteria&Mod=ACC", false
	objhttp.send SearchData.XMLDocument
	document.formname.action="PurchaseVouchers.asp?ACTN="& document.formname.hAction.value
	document.formname.submit()
End Function

Function ChkforApprove()
Dim j,dTrans,sMsgNo
	if not document.formname.hCnt.value="1" then
		for j=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(j).checked then
				dTrans=dTrans&":"&document.formname.chkbox(j).value
				sFrmAppNo= document.formname.hFrmAppNo(j).value
			end if
		next
		dTrans =Mid(dTrans,2)
	else
		dTrans=document.formname.chkbox.value
	end if

	if Trim(dTrans)="" then
		alert("Select Voucher")
		Exit Function
	end if

	IF trim(sFrmAppNo) = "2" then
		document.formname.hAppNo.value = sFrmAppNo
		document.formname.action="AppOtherPURViewWithMulTax.asp?TransNo="&dTrans&"&sPara=App"
		document.formname.submit
		exit function
	End IF
	if not Trim(dTrans)="" then
		document.formname.hTransNo.value=dTrans
		sMsgNo=MsgBox("Do you want to Approve", vbQuestion + vbOKCancel )
		if sMsgNo=1 then
			document.formname.action="AppVouStatusUpdateAll.asp"

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
				sFrmAppNo= document.formname.hFrmAppNo(k).value
			end if
		next
		dTrans=Mid(dTrans,2)
	else
		dTrans=document.formname.chkbox.value
	end if
	IF trim(sFrmAppNo) = "2" then
		MsgBox "Deletion Is Not Possible",64,"Purchase Vouchers"
		exit function
	End IF

	'alert(dTrans)
	if trim(dTrans) = "" then
		alert("Select Voucher")
		document.formname.btnAcc.disabled = False
		exit function
	end if
	If confirm("Do you want to delete the selected voucher") Then
	Else
		Exit Function
	End IF

	dTrans = dTrans&"|0"
	if not Trim(dTrans)="" then
		document.formname.hTransNo.value=dTrans
		sMsgNo=MsgBox("This will Permanently Delete the Voucher(s)" & vbCrLf &"Click OK to Delete", vbQuestion + vbOKCancel )
		if sMsgNo=1 then
			document.formname.action="PurVouDeletion.asp"

			document.formname.submit
		end if
	end if
End Function

Function ChkforEdit()
Dim dTrans,k,sTVal,sVal,VouTy
	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"~"&document.formname.chkbox(k).value
				sVal= Split(document.formname.chkbox(k).text,"&")
				sFrmAppNo= document.formname.hFrmAppNo(k).value
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

	if Trim(dTrans)="" then
		alert("Select Voucher")
		Exit Function
	end if

	IF trim(sFrmAppNo) = "2" then
		document.formname.hAppNo.value = sFrmAppNo
		document.formname.action="AppOtherPURViewWithMulTax.asp?TransNo="&dTrans&"&sPara=Edt"
		document.formname.submit
		exit function
	End IF
	if not Trim(dTrans)="" then
		document.formname.hTransNo.value=dTrans
		'document.formname.action="VouPURAmdDetails.asp"
		'document.formname.action="VouAmdPURBkSel.asp"

		document.formname.action = "PURCHASEVOUCHERENTRYEDIT.ASP?ACTN="&"E"&":"&dTrans
		document.formname.submit
	end if
End Function

Function ChkforACC()
Dim dTrans,k,sTVal,sVal,VouTy,sVouSt

	document.formname.btnAcc.disabled = True

	if not document.formname.hcnt.value ="1" then
		for k=0 to document.formname.hcnt.value-1
			if document.formname.chkbox(k).checked then
				dTrans=dTrans&"~"&document.formname.chkbox(k).value
				sVal= Split(document.formname.chkbox(k).text,"&")
				sFrmAppNo= document.formname.hFrmAppNo(k).value
			end if
		next
		dTrans=Mid(dTrans,2)
		if InStr(1,dTrans,"~")> 0 then
			MsgBox "Account One by One",64
			document.formname.btnAcc.disabled = False
			exit function
		end if
	else
		dTrans=document.formname.chkbox.value
		sVal= Split(document.formname.chkbox.text,"&")
	end if

	'alert(dTrans)
	if trim(dTrans) = "" then
		alert("Select Voucher")
		document.formname.btnAcc.disabled = False
		exit function
	end if

	Set sVouSt = Eval("document.formname.hVouSts"&dTrans)

	IF sVouSt.Value = "01" Then
		Msgbox "Only Approved Vouchers can be Accounted "
		document.formname.btnAcc.disabled = False
		Exit Function
	Elseif sVouSt.Value = "04" Then
		Msgbox "Voucher already Accounted "
		Exit Function
	End IF
	IF trim(sFrmAppNo) = "2" then
		document.formname.hAppNo.value = sFrmAppNo
		document.formname.action="AppOtherPURViewWithMulTax.asp?TransNo="&dTrans&"&sPara=Acc"
		document.formname.submit
		exit function
	End IF
	if not Trim(dTrans)="" then
		document.formname.hTransNo.value=dTrans
		document.formname.action="AccPurVouGenerate.asp"

		document.formname.btnAcc.disabled = True
		document.formname.submit
	end if
End Function


Function SelectAccHead()
dim iGlHead,sOrgId,sAccHead,arrTemp,sRetVal
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
set objhttp = CreateObject("Microsoft.XMLHTTP")


if document.formname.selBook.value="S" Then
		Msgbox "Select Book"
		document.formname.selBook.focus
		'document.formname.SelAccHead.selectedIndex=0
		Exit Function
Else
	sOrgId=document.formname.hOrgID.value
	sBookNo=document.formname.selBook.value
End if

End Function

Function Validate()

Dim sFromDate,sToDate

	sSelUnitId = document.formname.hOrgID.value
	sSelBook   = document.formname.SelBook.value
	sSelUser = document.formname.selUser.value
	sSelPurType = document.formname.SelPurType.value
	sSelVouFrm = document.formname.SelVouFrm.value

	sFromDate=document.formname.ctlVouFromDate.GetDate
	sToDate=document.formname.ctlVouToDate.GetDate


	IF document.formname.ChkVouDt.checked  then
		if dateDiff("d",sFromDate,sToDate)<0 then
			MsgBox "To Date Should be Greater than From Date"
			exit function
		end if
	end if
	IF document.formname.ChkVouNo.checked then
		if document.formname.txtVouNoFrom.value="" or document.formname.txtVouNoTo.value="" then
			MsgBox "Voucher No Empty",64
			exit function
		else
			iVouNoFrom = document.formname.txtVouNoFrom.value
			iVouNoTo   = document.formname.txtVouNoTo.value

		end if
	end if

	IF document.formname.ChkVouDt.checked then
		dtVouForom = sFromDate
		dtVouTo = sToDate
	end if

	IF document.formname.ChkVouAmt.checked then
		if document.formname.txtFromAmount.value="" or document.formname.txtToAmount.value="" then
			MsgBox "Voucher Amount Empty",64
			exit function
		else
			iVouAmtFrom = document.formname.txtFromAmount.value
			iVouAmtTo = document.formname.txtToAmount.value
		end if
	end if

	document.formname.hFromDate.value=sFromDate
	document.formname.hToDate.value=sToDate


	sParSubTypeValue = ""
	sParSubTypeName = ""
	IF spParSubType.innerHTML <> "" then
		document.formname.hXmlAccFlag.value = True

		sParSubTypeValue = document.formname.hParSubTypeVal.value
		sParSubTypeName = document.formname.hParSubTypeName.value
	end if

	iAccHeadNo = "0"
	sAccHeadName = ""
	IF document.formname.txtAccHead.value <> "" then
		document.formname.hXmlParFlag.value = True
		iAccHeadNo = document.formname.hAccHead.value
		sAccHeadName = document.formname.txtAccHead.value
	End IF

	'********

	Set Root = SearchData.documentElement
	'alert Root.xml
	IF Root.haschildnodes then
		For each node in Root.childnodes
			Set RemNode = node
			Root.removechild RemNode
		Next
	End IF

	sTransType = ""
	Dim i

	if trim(document.formname.hAction.value)="L" then
	    for i=1 to 3
		    if document.formname.voutype(i).checked then
			    sTransType= sTransType & "," & document.formname.voutype(i).value
		    end if
	    next
	else
	    for i=0 to 1
		    if document.formname.voutype(i).checked then
			    sTransType= sTransType & "," & document.formname.voutype(i).value
		    end if
	    next
	end if 'if trim(document.formname.hAction.value)="L" then


	if trim(sTransType) <> "" then
		sTransType = mid(sTransType,2)
	end if

	Root.setAttribute "Src",document.formname.hGridName.value
	Root.setAttribute "TransType",sTransType

	Set Elem = OutData.createElement("SearchOption")
	Elem.setAttribute "SelUnitId",sSelUnitId
	Elem.setAttribute "SelBook",sSelBook
	Elem.setAttribute "SelUser",sSelUser
	Elem.setAttribute "SelPurType",sSelPurType
	Elem.setAttribute "SelVouFrom",sSelVouFrm
	Elem.setAttribute "VoucherType", ""
	Root.appendchild Elem

	Set Elem1 = OutData.createElement("VoucherNo")
	Elem1.setAttribute "From", iVouNoFrom
	Elem1.setAttribute "To", iVouNoTo
	Root.appendchild Elem1


	Set Elem2 = OutData.createElement("VoucherDate")
	Elem2.setAttribute "From", dtVouForom
	Elem2.setAttribute "To", dtVouTo
	Root.appendchild Elem2


	Set Elem3 = OutData.createElement("VoucherAmount")
	Elem3.setAttribute "From", iVouAmtFrom
	Elem3.setAttribute "To", iVouAmtTo
	Root.appendchild Elem3

	Set Elem4 = OutData.createElement("AccHead")
	Elem4.setAttribute "No", iAccHeadNo
	Elem4.setAttribute "Name",sAccHeadName
	Elem4.setAttribute "ParSubTypeValue",sParSubTypeValue
	Elem4.setAttribute "ParSubTypeName",sParSubTypeName
	Root.appendchild Elem4


	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=SearchCriteria&Mod=ACC", false
	objhttp.send SearchData.XMLDocument


	'********
	document.formname.action="PurchaseVouchers.asp?ACTN="& document.formname.hAction.value
	document.formname.submit
End Function

Function ChkReset()
	document.formname.ctlVouFromDate.setDate=date
	document.formname.ctlVouToDate.setDate=date
	'document.formname.selUnitId.selectedIndex=0
	document.formname.selPurType.selectedIndex=0
	document.formname.selBook.selectedIndex=0
	document.formname.txtFromAmount.value=""
	document.formname.txtToAmount.value=""
	document.formname.txtVouNoFrom.value=""
	document.formname.txtVouNoTo.value=""
	document.formname.txtAccHead.value=""

	DisplayBook()
End Function

Function ClearSearch()

	Set Root = SearchData.documentElement
	'alert Root.xml
	IF Root.haschildnodes then
		For each node in Root.childnodes
			Set RemNode = node
			Root.removechild RemNode
		Next
	End IF

	Root.setAttribute "Src",document.formname.hGridName.value
	Root.setAttribute "TransType",""

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=SearchCriteria&Mod=ACC", false
	objhttp.send SearchData.XMLDocument

	document.formname.action="PurchaseVouchers.asp?ACTN="& document.formname.hAction.value
	document.formname.submit()
End Function

Function ShowVouch(iTransNo)
	showModalDialog "PurchaseVouchView_San.asp?TransNo="&iTransNo,"","dialogHeight:410px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No"
End Function


Function SetDate()
	Dim sFDate,sTDate


	Set Root = SearchData.documentElement


	sFDate=document.formname.hFromDate.value
	sTDate=document.formname.hToDate.value

	if Trim(sFDate)<>"" and Trim(sTDate)<>"" then
		document.formname.ctlVouFromDate.setDate=sFDate
		document.formname.ctlVouToDate.setDate=sTDate
	end if

	'*****

	'alert "Root="&Root.xml
	 If Root.haschildnodes then
	 	IF Root.getAttribute("Src") = document.formname.hGridName.value then




	 		For each node in Root.childnodes
				'alert(	node.xml)

				if trim(node.NodeName) = "SearchOption" then

					'document.formname.selUnitId.value	= node.getAttribute("SelUnitId")
					Call DisplayBook()
					document.formname.SelBook.value		= node.getAttribute("SelBook")
					document.formname.selUser.value		= node.getAttribute("SelUser")
					document.formname.SelPurType.value	= node.getAttribute("SelPurType")

					sVouFromData = node.getAttribute("SelVouFrom")


		            if trim(document.formname.hAction.value)="L" then
					    for i=1 to 3
    						if instr(1,trim(Root.getAttribute("TransType")),trim(document.formname.voutype(i).value) ) > 0 then
							    document.formname.voutype(i).checked = True
						    end if
					    next
					else
					    for i=0 to 1
    						if instr(1,trim(Root.getAttribute("TransType")),trim(document.formname.voutype(i).value) ) > 0 then
							    document.formname.voutype(i).checked = True
						    end if
					    next
					end if' if trim(document.formname.hAction.value)="L" then


				end if

	 			if trim(node.NodeName) = "VoucherType" then
	 				'document.formname.hVouType.value = node.getAttribute("Value")
	 			end if
	 			if trim(node.NodeName) = "VoucherNo" then
	 				If trim(node.getAttribute("From")) <> "" and trim(node.getAttribute("To")) <> "" then
	 					document.formname.txtVouNoFrom.value=node.getAttribute("From")
	 					document.formname.txtVouNoTo.value=node.getAttribute("To")
	 					document.formname.ChkVouNo.checked = True
	 				end if
	 			end if
	 			if trim(node.NodeName) = "VoucherDate" then
	 				If trim(node.getAttribute("From")) <> "" and trim(node.getAttribute("To")) <> "" then
	 					document.formname.hFromDate.value=node.getAttribute("From")
	 					document.formname.hToDate.value=node.getAttribute("To")
	 					document.formname.ChkVouDt.checked = True
	 				end if
	 			end if

	 			if trim(node.NodeName) = "VoucherAmount" then
	 				If trim(node.getAttribute("From")) <> "" and trim(node.getAttribute("To")) <> "" then
	 					document.formname.txtFromAmount.value=node.getAttribute("From")
	 					document.formname.txtToAmount.value=node.getAttribute("To")
	 					document.formname.ChkVouAmt.checked = True
	 				end if
	 			end if

	 			if trim(node.NodeName) = "AccHead" then
	 				IF node.getAttribute("No") <> "" then
	 					document.formname.hAccHead.value = node.getAttribute("No")
	 					document.formname.txtAccHead.value =node.getAttribute("Name")
	 					document.formname.hParSubTypeVal.value = node.getAttribute("ParSubTypeValue")
						document.formname.hParSubTypeName.value = node.getAttribute("ParSubTypeName")
						spParSubType.innerText = node.getAttribute("ParSubTypeName")

	 				End if
	 			end if
	 		Next
	 	End If  'IF Root.getAttribute("Src") = document.formname.hGridName.value then
	 End IF

	'*****

if sVouFromData = "" then sVouFromData = "P"
document.formname.SelVouFrm.value	= sVouFromData

ChangeStatusOfInputFields(document.formname.ChkVouNo)
ChangeStatusOfInputFields(document.formname.ChkVouAmt)

End Function
'========================================================================================================
Function ResetAccHead()
	spParSubType.innerText = ""
	document.formname.hParSubTypeVal.value = ""
	document.formname.hParSubTypeName.value = ""
	document.formname.hAccHead.value = ""
	document.formname.txtAccHead.value  = ""
End Function

'********************Start******Added Nelwy by S.Maheswari on Mar 10 th 2009 for multiple selection of party sub type and party ***************
'========================================================================================================
Function SelAccHeadPopup()
	dim Root,objhttp

	sUnit = document.formname.hOrgID.value

'	if trim(sUnit) = "" then
'		alert("Select Unit")
'		document.formname.selUnitId.focus
'		Exit Function
'	end if

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	Set Root = AccHeadData.documentElement


	showModalDialog "PartySubTypeMultipleSel.asp?Unit="&sUnit,AccHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No"
	'window.open "AccountHeadMultipleSel.asp?Unit="&sUnit,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No"
	Set Root = AccHeadData.documentElement
	 'alert Root.xml
	IF Not Root.haschildnodes then
		spParSubType.innerHTML = ""
		exit function
	End IF

	IF Root.haschildnodes then

		For each node in Root.childnodes

			IF trim(node.NodeName) = "Details" then
				'alert node.getAttribute("PartySubTypeName")
				sParSubTypeName = sParSubTypeName&","&node.getAttribute("PartySubTypeName")
				sParSubTypeValue = sParSubTypeValue &",'"& node.getAttribute("PartyType") & node.getAttribute("PartySubType")&"'"

			End If
			' Root.appendChild node
		Next

	End IF
	objhttp.Open "POST","XMLSave.asp?Name=PartySubType&Mod=PUR", false
	objhttp.send AccHeadData.XMLDocument

	document.formname.hParSubTypeVal.value = mid(sParSubTypeValue,2)
	document.formname.hParSubTypeName.value = mid(sParSubTypeName,2)
	spParSubType.innerHTML = mid(sParSubTypeName,2)
End Function
'========================================================================================================
Function SelPartyPopup()
	Dim Root,Root1,objhttp
	set Root = AccHeadData.documentElement
	set Root1 = PartyData.documentElement

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	sOrgId = document.formname.hOrgID.value
'	if trim(sOrgId) = "" then
'		alert("Select Unit")
'		document.formname.selUnitId.focus
'		Exit Function
'	end if

	document.formname.hPartyName.value = ""
	document.formname.txtAccHead.value = ""
	IF Not Root.haschildnodes then
		sPartyType = "0"
	Else
		IF Root.haschildnodes then
			for each node in Root.childnodes
				IF trim(node.NodeName) = "Details" then
					sParType = 	sParType & "," & node.getAttribute("PartyType")
					sParSubType = sParSubType & "," & node.getAttribute("PartySubType")
					sParSubTypeName = sParSubTypeName &","& node.getAttribute("PartySubTypeName")
				End IF
			Next
		End IF
		sParType = mid(sParType,2)
		sParSubType = mid(sParSubType,2)
		sParSubTypeName = mid(sParSubTypeName,2)
		sPartyType  = sParType & "?" & sParSubType & "?" & sParSubTypeName
	End IF


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


		    objhttp.Open "POST","XMLSave.asp?Name=PartyType&Mod=PUR", false
		    objhttp.send PartyData.XMLDocument

		    document.formname.hPartyName.value = mid(sPartyName,2)
		    document.formname.hParCode.value = mid(sParCode,2)
		    'sParVal = mid(sParTy,2) &"?"&mid(sParSubType,2)&"?"&mid(sPartyName,2)&"?"&mid(sParCode,2)
		    document.formname.hAccHead.value = mid(sParCode,2)
		    document.formname.txtAccHead.value = mid(sPartyName,2)

	    End If
End Function
' ********************************************** End ************************************************************
</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate();DisplayBook()">

	<form method="POST" name="formname">

	<input type=hidden name="hTransNo" value="">
	<input type=hidden name="hAppNo" value="">
	<input type=hidden name="hAction" value="<%=sAction%>">


	<input type=hidden name="hFromDate" value="<%=sFrmDate%>">
	<input type=hidden name="hToDate" value="<%=sToDate%>">

	<input type=hidden name="hGridName" Value="<%=sGridName%>">
	<input type=hidden name="hFinPeriod" value="<%=sFinPeriod%>">

	<input type=hidden name="hXmlAccFlag" value="<%=sXmlAccFlag%>">
	<input type=hidden name="hXmlParFlag" value="<%=sXmlParFlag%>">


	<input type=hidden name="hParSubTypeVal" value="">
	<input type=hidden name="hParSubTypeName" value="">
	<input type=hidden name="hAccHead" value="">
	<input type=hidden name="hParCode" Value="">
	<input type=hidden name="hPartyName" Value="">
	<input type=hidden name="hOrgID" value="<%=sUnitID %>">
	<input type=hidden name="hOrgName" value="<%=sOrgName %>">

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
				Purchase Vouchers
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
<td valign="center" class="SubTitle">&nbsp;&nbsp;
<%

	sSql="SELECT A.CREATEDVOUCHERNO,A.VOUCHERDATE,A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.VOUCHERAMOUNT,A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,A.BANKINSTRUMENTTYPE,V.PARTYNAME,isNull(A.OtherApplnTransNo,0) OtherApp ,isNull(A.FromApplication,0) FromApp FROM ACC_T_CREATEDVOUCHERHEADER AS A " _
		& "INNER JOIN APP_M_PARTYMASTER AS V ON A.PARTYCODE=V.PARTYCODE WHERE A.BOOKCODE='04' "

	if Trim(Cstr(sUnitID))<> "" then
		sSql =sSql & " AND A.OUDEFINITIONID='"&Cstr(sUnitID) &"'"
	end if
	if Trim(sPurTypeValue)<> "" and trim(sPurTypeValue)<>"0" then
		sSql= sSql & " AND A.BANKINSTRUMENTTYPE='"& sPurTypeValue &"'"
	end if

	If trim(sBookNo) <> "" Then
		sSql = sSql & " AND A.BOOKNUMBER='"& Cstr(sBookNo) &"' "
	End if

	IF CStr(sVouFrm) = "A" Then 'Only Vouchers Created in Accounts Module Only
		sSql = sSql & " AND isNull(A.OtherApplnTransNo,0) = 0 "
	Elseif CStr(sVouFrm) = "P" Then 'Only Vouchers Created in Purchase Module Only
		sSql = sSql & " AND isNull(A.OtherApplnTransNo,0) <> 0 "
	End IF



	If trim(sUserID) <> "" then
		IF trim(sUserID) <> "A"  and trim(sUserID) <> "0" Then
			sSql = sSql & " AND A.CREATEDBY = "& sUserID &" "
		End IF
	End IF

	IF CStr(sSelectTransType) = "" or CStr(sSelectTransType) = "A" then
	else
		sSql = sSql & " AND ( A.CREATEDVOUCHSTATUS='0' "
	end if

	if trim(sAction)="L" then
	    IF CStr(sSelectTransType) = "" or CStr(sSelectTransType) = "A" then
		    Response.Write ("<Input type=checkbox name=voutype value=A checked onclick=ChkVouType()>All&nbsp;")
	    else
		    Response.Write ("<Input type=checkbox name=voutype value=A onclick=ChkVouType()>All&nbsp;")
	    end if


	    if Instr(1,sSelectTransType,"C") > 0 then
		    Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType() checked>Created&nbsp;")
		    sSql=sSql+" OR A.CREATEDVOUCHSTATUS='010101' OR A.CREATEDVOUCHSTATUS='010102'"
	    Else
		    Response.Write ("<Input type=checkbox name=voutype value=C onclick=ChkVouType()>Created&nbsp;")
	    End IF
	end if

	if Instr(1,sSelectTransType,"P") > 0 then
		Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType() checked >Approved&nbsp;")
		sSql=sSql+" OR A.CREATEDVOUCHSTATUS='010103' OR A.CREATEDVOUCHSTATUS='010105'"
	Else
		Response.Write ("<Input type=checkbox name=voutype value=P onclick=ChkVouType()>Approved&nbsp;")
	End IF

	if Instr(1,sSelectTransType,"T") > 0 Then
		Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType() checked >Accounted&nbsp;")
		sSql=sSql+" OR A.CREATEDVOUCHSTATUS='010104'"
	Else
		Response.Write ("<Input type=checkbox name=voutype value=T onclick=ChkVouType()>Accounted&nbsp;")
	end if

	IF CStr(sSelectTransType) = "" or CStr(sSelectTransType) = "A" then
	else
		sSql=sSql+")"
	end if



	If Cstr(sVouNoFlag) = "Y" then
			sSql=sSql+" AND A.CREATEDVOUCHERNO BETWEEN '"& sFrmNo &"' AND '"& sToNo&"'"
	End If
	If Cstr(sVouAmtFlag ) = "Y" then
			sSql=sSql+" AND A.VOUCHERAMOUNT BETWEEN '"& sFrmAmt &"' AND '"& sToAmt&"'"
	End If


	If sVouAccFlag = True then

		if trim(sParsubType) <> "" then

			sSql=sSql+" and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherHeader where "&_
				" ltrim(rtrim(PartyType))+cast(ltrim(rtrim(PartySubType)) as Char ) in("&sParsubType&")  "
		end if

		If trim(iParCode) <> "" and trim(iParCode) <> "0" then
				sSql=sSql+"and V.PartyCode in ("& iParCode &") "
		End If
		if trim(sParsubType) <> "" then
			sSql=sSql+" )"
		end if
	End IF





	if not Cstr(sVouDtFlag)= "Y" then
		sSql=sSql+" AND CONVERT(DATETIME,A.VOUCHERDATE,103) BETWEEN CONVERT(DATETIME,'"& "01/04/" & LEFT(sFinPeriod,4) &"',103) " _
		& " AND CONVERT(DATETIME,'"& "31/03/" & RIGHT(sFinPeriod,4) & "',103)"
	else
		sSql=sSql+" AND CONVERT(DATETIME,A.VOUCHERDATE,103) BETWEEN CONVERT(DATETIME,'"& sFrmDate &"',103) " _
		& " AND CONVERT(DATETIME,'"& sToDate & "',103)"
	end if

	'sSql=sSql & "  ORDER BY A.CREATEDTRANSNO DESC "
	sSql=sSql & " Order By " &  sSortBy  &  ",A.CREATEDTRANSNO "
	'Response.Write "<br> sSql = " & sSql
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
		<option value="">Select</option>
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

		If not dcrs.EOF then
			Do While Not dcrs.EOF
			%>
		          <OPTION VALUE="<%=dcrs(0)%>" ><%=dcrs(3)%></Option>
			<%
				dcrs.MoveNext
			Loop
		end if

		dcrs.Close
	%>
	</select>
	</td>
</tr>-->

<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Purchase Book</td>
	<td class="FieldCellSub" colspan="4">
	<select size="1" name="selBook" class="FormElem" >
		<option value="" >Select Book</option>	</select>
	</td>
</tr>
<tr>
<td class="FieldCellSub">
</td>
<td class="FieldCellSub">User ID
</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selUser" class="FormElem">
	<option value="0">Select User</option>

	<option value="A">All</option>

	<%

		Dim rsTemp,sqry
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		sqry = "SELECT DISTINCT INTERNALUSERID,LOGINID FROM VwUserUnitList WHERE APPLICATIONCODE = 1  Order By LOGINID "
		Response.Write "qry="& sqry
		rsTemp.Open sqry,con
		Do while not rsTemp.EOF
	%>
			<option value="<%=rsTemp(0)%>" > <%=rsTemp(1)%></option>

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
	<td class="FieldCellSub">Purchase Type</td>
	<td class="FieldCellSub" colspan="4">
	<select size="1" name="selPurType" class="FormElem" >
		<option value=0>Select Purchase Type</option>
		<%
			dim sCode,sValue,sName,sshortCode
			with dcrs
				.ActiveConnection=con
				.CursorLocation=3
				.CursorType=3
				.Source="Select PurchaseType,PurchaseTypeName,PurTypeShortName from APP_M_PurchaseTypes Where Active = 'Y'"
				.Open
			end with
			set dcrs.ActiveConnection=nothing
			Set sCode = dcrs(0)
		  	Set sValue = dcrs(1)
		  	Set sshortCode=dcrs(2)
		  	Do while not dcrs.EOF

		%>
				<option value="<%=sCode %>"><%=sshortCode&"---"&sValue%></option>
		<%

				dcrs.MoveNext
			Loop
			dcrs.Close


		%>
		</select>
	</td>

</tr>
<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Vouchers From</td>
	<td class="FieldCellSub" colspan="4">
		<select size="1" name="selVouFrm" class="FormElem">
		<option value=0 >All</option>
		<option value=A>From Accounts Only</option>
		<option value=P>From Other Applications Only</option>
		</Select>
	</td>
<tr>





<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">

	<input type="checkbox" value="Y" name="ChkVouNo"  OnClick="ChangeStatusOfInputFields(this)"  >Vou No. From

	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtVouNoFrom"  size="20" class="FormElem">
	</td>

	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Vou No. To
	</td>
	<td class="FieldCellSub">
		<input type="text" name="txtVouNoTo"  size="20" class="FormElem">
	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">

	<input type="checkbox" value="Y" name="ChkVouDt" OnClick="ChangeStatusOfInputFields(this)"  >Voucher Date

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

	<input type="checkbox" value="Y" name="ChkVouAmt"  OnClick="ChangeStatusOfInputFields(this)" >	Amount From

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
	<td class="FieldCellSub" align="right">
			Account Head
	</td>
	<td class="FieldCellSub" colspan="4">
		<span id="spParSubType"  class="DataOnly" ></span>&nbsp;
		<a href="Javascript:SelAccHeadPopup()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Account Head" ></a>
		<!--select class="formelem" disabled OnChange="SelectAccHead()" size="1" name="selAccHead">
			<option value="0">Select Option</option>
			<option value="G">Select Account Head</option>
		</select-->
		&nbsp;
		<a href="Javascript:ResetAccHead()"><img border="0" width="11" height="11" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Remove Account Head" ></a>

	</td>
</tr>
<tr>
	<td class="FieldCellSub" ></td>
	<td class="FieldCellSub" ></td>
	<td colspan="4" class="FieldCellSub">
		<input type="text" name="txtAccHead" Readonly size="70" class="FormElemRead">
		<a href="Javascript:SelPartyPopup()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" ></a>
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
<td class="FieldCell" >
	<input type="button" value="Clear Search" name="CmdClear" class="ActionButtonX" onclick="ClearSearch()" >
</td>
</table>
</div>
</td>
</tr>
<tr>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5px">
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
<td class="ExcelHeaderCell" width="10px" rowspan="2" >S.No.
</td>
<td class="ExcelHeaderCell" width="10px" rowspan="2" >
</td>
<td class="ExcelHeaderCell" rowspan="2" >
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
<td class="ExcelHeaderCell" rowspan="2" >
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
<td class="ExcelHeaderCell" rowspan="2" >Party Name
</td>
<td class="ExcelHeaderCell" colspan="3">Amount
</td>

<td class="ExcelHeaderCell" rowspan="2" >Status
</td>
</tr>
<tr>
<td class="ExcelHeaderCell">
<%
if trim(sField5) <> ""  then
	if instr(1,sField5,":") > 0 then
		Arr1 = Split(sField5,":")

		if Arr1(1) = "A" then ' if ascending order is exist, give option to descending order
			%>
			<span style="cursor:hand" onclick="Sort(5,'M','D')">Invoice</span>
			<%
		else
			%>
			<span style="cursor:hand" onclick="Sort(5,'M','A')">Invoice</span>
			<%
		end if
	end if
else
	%>
	<span style="cursor:hand" onclick="Sort(5,'M','A')">Invoice</span>
	<%
end if
%>
</td>
<td class="ExcelHeaderCell">Paid
</td>
<td class="ExcelHeaderCell">Over Due
</td>
</tr>
<SCRIPT LANGUAGE=vbscript RUNAT=Server>

</SCRIPT>
<% dim iFromApp,dTotAmtPaid,iNoOfDays,sInvoiceDate,iDueDays,sTransType,nTransactionNo
	'Response.write "<p><font color=red>"&sSql &"<br><br>"
	'Response.Write "<P style='color:red' >" & sSortBy

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

	'dTotAmtPaid = 0
	'******* Start of Paging
	Objrs.PageSize=iPageSize
	if iCurrentPage=0 then iCurrentPage=1
	Objrs.AbsolutePage=iCurrentPage
	iTotalPage=objrs.PageCount
	For iPageCtr=1 to objrs.PageSize

		iCnt=iCnt+1
		iCrTransNo=Objrs("createdtransno")
		iFromApp = Objrs("FromApp")
		'Response.Write "iFromApp="&iFromApp

		'************************** Newly Added by S.Maheswari on 4th Mar 2009 to display total amount paid *********************************************************************
		sqry = "Select PayablesNumber from Acc_T_Payables where TransactionNumber = (Select TransactionNumber from Acc_T_VoucherHeader where CreatedTransno = "& iCrTransNo &" ) "
        'Response.Write sQuery &"<BR>"
        Objrs1.Open sqry,con

        dTotAmtPaid = 0
        do while not objRs1.EOF
			'Acc_T_AdvancePayments
			sqry = "Select AmountPaid from Acc_T_PybleAdjustmentDetails where  Payablesnumber = "& objRs1(0) &" "
			'Response.Write sqry
			rsTemp.Open sqry,con
			do while not rsTemp.EOF
				dTotAmtPaid = dTotAmtPaid + cdbl(rsTemp(0))
				rsTemp.MoveNext
			loop
			rsTemp.Close
			Objrs1.MoveNext
		loop
		objRs1.Close
	   '***************************************************************************************************************************************************************************

	   'Find Over due
	   sqry = " select datediff(day,convert(datetime,PartyBillDate,103),convert(datetime,getdate(),103)) From Acc_T_Payables where  TransactionNumber = (Select TransactionNumber from Acc_T_VoucherHeader where CreatedTransno = "& iCrTransNo &" )"
	   rsTemp.Open sqry,con
	   If Not rsTemp.EOF Then
		iNoOfDays = rsTemp(0)
	   End If
	   rsTemp.Close
	   'Response.Write "<p><font color=red>iNoOfDays="&iNoOfDays

	   sSql = " Select Distinct TransactionType,Convert(Char,VoucherDate,103),TransactionNumber "&_
			  " From Acc_T_VoucherHeader Where CreatedTransno = "& iCrTransNo &" "

		with objRs1
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sSql
			.Open
		End with
		set objRs1.ActiveConnection =nothing
		IF Not objRs1.EOF Then
			sTransType = objRs1(0)
			sInvoiceDate = objRs1(1)
			nTransactionNo = objRs1(2)
		End IF
		objRs1.Close
		'Response.Write "<p><font color=red>iDueDays="&iCrTransNo & "="& iNoOfDays & "="& nTransactionNo & "="& sInvoiceDate
		IF CStr(sTransType) = "SJR" Then
			iDueDays = GetDueDays(iCrTransNo,iNoOfDays,nTransactionNo,sInvoiceDate)
		Else
			iDueDays = 0
		End IF

	%>
<tr>
<td class="ExcelSerial"><%=iCnt%></td>
<td class="ExcelDisplayCell" width="10px" >
<%If Right(CStr(Objrs("createdvouchstatus")),2)="04" then %>
	<input type="checkbox" name="Chkbox" value="<%=iCrTransNo%>" disabled >
<%else%>
	<input type="checkbox" name="Chkbox" text="<%=Objrs("createdvoucherno")&"&"& Right(Objrs("transactiontype"),1)%>" value="<%=iCrTransNo %>" >
<%end if%>

<Input type="hidden" name="hVouSts<%=iCrTransNo%>" Value="<%=Right(CStr(Objrs("createdvouchstatus")),2)%>">
<Input type="hidden" name="hFrmAppNo" value="<%=iFromApp%>">
<td class="ExcelDisplayCell" align="left" ><a href="#" LANGUAGE="VBSCRIPT" onclick="ShowVouch(<%=iCrTransNo %>)" class="ExcelDisplayLink"><%=Objrs("createdvoucherno") %></a></td>
<td class="ExcelDisplayCell" align="left" ><%=FormatDate(Objrs("voucherdate"))%></td>
<td class="ExcelDisplayCell" align="left" ><%=Objrs("partyname")%></td>
<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(Objrs("Voucheramount")) %></td>
<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(dTotAmtPaid) %></td>
<td class="ExcelDisplayCell" align="right" ><%=iDueDays%></td>
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
		Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Created")
	elseif Right(CStr(Objrs("createdvouchstatus")),2)="04" then
		Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& AccVoucherNo)
	else
		Response.Write ("<td class=ExcelDisplayCell align=left height=16>"& "Approved")
	end if


	IF Trim(Objrs("FromApp")) <> "0" Then
		Response.Write "<font color=Red>*</Font></td>"
		'Response.Write(Trim(Objrs("FromApp")))
	Else
		Response.Write "</td>"
	End IF

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
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top" align="right" class="FieldCell"><Font color=Red><b>*</b></Font> Vouchers posted from purchase module
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
<td align="center" class="ClearPixel" width="5px">
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

<%if trim(sAction)="L" then %>
    <input type="button" value="Edit" name="B9" class="ActionButton" tabindex="3" onclick="ChkforEdit()">
    <input type="button" value="Approve" name="B10" class="ActionButton" tabindex="3" onclick="ChkforApprove()">
<%end if'if trim(sAction)="L" then   %>
<input type="button" value="Account" name="btnAcc" class="ActionButton" tabindex="3" onclick="ChkforACC()">
<%if trim(sAction)="L" then %>
<input type="button" value="Delete" name="B12" class="ActionButton" tabindex="3" onclick="ChkforDelete()">
<%end if'if trim(sAction)="L" then   %>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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
<%
Function GetDueDays(iCrTrNo,iOutStdDays,iAccTrNo,sBillDate)
		Dim sQuery,sObjDueRs,iOthAppNo,iPayTerms,iPayNoDays,iPayCount,iTotalDueDay
		Dim iDurPer,iDurDays,sPayTillDate,sObjPayRs,iPayRecNo,iParPayAmt,iPayInvAmt
		Dim iPayAmtToCome
		Set sObjDueRs = Server.CreateObject("ADODB.RecordSet")
		Set sObjPayRs = Server.CreateObject("ADODB.RecordSet")

		sQuery = "Select isNull(OtherApplnTransNo,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iCrTrNo
		'Response.Write sQuery
		sObjDueRs.Open sQuery,Con
		IF Not sObjDueRs.EOF Then
			iOthAppNo = sObjDueRs(0)
		End IF
		sObjDueRs.Close

		IF CStr(iOthAppNo) = "0" Then
			GetDueDays = 0
		Else
			sQuery = "Select InvPymtTerms From Sal_T_InvoiceHeader Where SaleTransactionNo = " & iOthAppNo
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayTerms = sObjDueRs(0)
			End IF
			sObjDueRs.Close

			'iPayTerms = 2

			sQuery = "Select Count(1) From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
			Response.Write squery
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayCount = sObjDueRs(0)
			End IF
			sObjDueRs.Close



			IF iPayCount = 1 Then
				sQuery = "Select DueDay From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayNoDays = sObjDueRs(0)
				End IF
				sObjDueRs.Close

				'iPayNoDays = 300

				'Response.Write iAccTrNo &"   " & iPayNoDays &"<br>"
				iTotalDueDay = CDbl(iOutStdDays) - CDbl(iPayNoDays)
			Else
				iPayAmtToCome = 0
				'sQuery = "Select ReceivableNumber,AmountReceivable From Acc_T_Receivables Where TransactionNumber = "&iAccTrNo
				sQuery = "Select PayablesNumber,AmountPayable From Acc_T_Payables Where TransactionNumber = "&iAccTrNo
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayRecNo = sObjDueRs(0)
					iPayInvAmt = sObjDueRs(1)
				End IF
				sObjDueRs.Close

				sQuery = "Select DueDay,DuePercent From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				Do While Not sObjDueRs.EOF
					iDueDays = sObjDueRs(0)
					iDurPer = sObjDueRs(1)

					'Response.Write iPayInvAmt &"  " & iDurPer &"<br>"

					iPayAmtToCome = Cdbl((CDbl(iPayInvAmt) * CDbl(iDurPer))/100) + CDbl(iPayAmtToCome)


					'sQuery = "Select Top 1 Convert(Varchar,DateAdd(day,"&iDueDays&",Convert(Datetime,'"&sBillDate&"',103)),103) From Acc_T_RcvblAdjustmentDetails  "
					sQuery = "Select Top 1 Convert(Varchar,DateAdd(day,"&iDueDays&",Convert(Datetime,'"&sBillDate&"',103)),103) From Acc_T_PybleAdjustmentDetails  "
					'Response.Write sQuery
					sObjPayRs.Open sQuery,Con

					IF Not sObjPayRs.EOF Then
						sPayTillDate = sObjPayRs(0)
					End IF
					sObjPayRs.Close

					'sQuery = "Select Sum(AmountReceived) From Acc_T_RcvblAdjustmentDetails Where ReceivableNumber = "&iPayRecNo&" and "&_
					'		 "Convert(Datetime,ReceivedOn,103) <=  Convert(Datetime,'"&sPayTillDate&"',103)  "
					sQuery = "Select Sum(AmountPaid) From Acc_T_PybleAdjustmentDetails Where PayablesNumber = "&iPayRecNo&" and "&_
							 "Convert(Datetime,PaidOn,103) <=  Convert(Datetime,'"&sPayTillDate&"',103)  "

					'Response.Write sQuery &"<br>"
					sObjPayRs.Open sQuery,Con
					IF sObjPayRs.EOF Then
						iParPayAmt = sObjPayRs(0)
					End IF
					sObjPayRs.Close

					'iParPayAmt = 119000
					'Response.Write iParPayAmt &"  " & iPayAmtToCome &"<br>"

					IF CDbl(iParPayAmt) >= CDbl(iPayAmtToCome) Then
						iTotalDueDay = 0
					Else
						iTotalDueDay = CDbl(iOutStdDays) - CDbl(iDueDays)
					End IF

					'Response.Write iTotalDueDay &"<br>"



					sObjDueRs.MoveNext
				Loop
				sObjDueRs.Close
			End IF
		End IF

		IF iTotalDueDay <0 Then
			iTotalDueDay = 0
		End IF

		IF CStr(iTotalDueDay) = "" Then
			iTotalDueDay = 0
		End IF

		GetDueDays = iTotalDueDay
	End Function
%>