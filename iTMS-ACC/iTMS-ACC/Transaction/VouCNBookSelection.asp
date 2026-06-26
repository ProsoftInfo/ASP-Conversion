<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	Feb 27,2003
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
<%
dim sFinPeriod,sFinFun

sFinPeriod = Session("FinPeriod")
sFinFun = GetfinYear(sFinPeriod)
If trim(sFinFun) = "True" Then
	Response.Redirect ("../../welcome_Welcome.asp?sFinFun="&sFinFun&"")
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="CommData"><Book/></XML>
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>

<SCRIPT language="vbscript">
FUNCTION popPartyType()
	dim iHeadCount

	iUnitNo=document.formname.selUnitId.value
	iBkNo=document.formname.selBook.value

	for iCounter=1 to document.formname.selPartyType.length
		document.formname.selPartyType.remove(1)
	next

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		iCounter=1
		For Each HeaderNode In Root.childNodes
			set oText1 = document.createElement("<Option>" )
				oText1.Text = HeaderNode.text
				oText1.Value = HeaderNode.Attributes.Item(0).nodeValue

			document.formname.selPartyType.add oText1,iCounter
			iCounter=CDbl(iCounter)+1
		next
	end if

END FUNCTION

'Function DisplayBook(objUnit)
Function DisplayBook()
dim iUnitNo,arrTemp
dim Root
	document.formname.selBook.options.length = 1
	iUnitNo = document.formname.selUnitId.value
	'alert iUnitNo
	'if objUnit.selectedIndex <> "0" then
		'iUnitNo= objUnit(objUnit.selectedIndex).value
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'alert document.formname.ChkGJ.checked
		if document.formname.ChkGJ.checked = True then
			objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=08&orgID=" & iUnitNo , false
		else
			objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=07&orgID=" & iUnitNo , false
		end if
		objhttp.send

		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement

			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if
		popPartyType
	'end if
end Function

Function NewPage()
	document.formname.action = "CreditNoteNewPage.asp"
	document.formname.submit
End Function

Function VouCreate
	Dim sType,sTemp

	If document.formname.ChkGJ.checked = True then
		document.formname.hChkVal.value = "Y"
	else
		document.formname.hChkVal.value = "N"
	End If
	if validate then
		if document.formname.selVoucherType.value="SC" then
			document.formname.action="VouCNCommisionEntry.asp"
		elseif document.formname.selVoucherType.value="SR" then
			document.formname.action="VouCNSalReturnEntry.asp"
		elseif document.formname.selVoucherType.value="OT" then
			document.formname.action="VouCNOthersEntry.asp"
		End IF

		IF CStr(document.formname.selInvoiceNo.value) = "S" and Cstr(document.formname.selVoucherType.value) <> "OT" Then
			MsgBox "Select Invoice "
			document.formname.selInvoiceNo.focus()
			Exit Function
		End IF

		sType = Split(document.formname.selInvoiceNo.value,":")

		'if document.formname.selVoucherType.value="OIS" or document.formname.selVoucherType.value="OIP" then
		'	IF sType(1) = "S" Then
		'		document.formname.action="VouCNOtherInvEntry.asp"
		'	Else
		'		document.formname.action="VouCNPurInvEntry.asp"
		'	End IF
		'End IF

		if document.formname.selVoucherType.value="OIS" Then
			document.formname.action="VouCNOtherInvEntry.asp"
		End IF

		if document.formname.selVoucherType.value="OIP" Then
			document.formname.action="VouCNPurInvEntry.asp"
		End IF

		sTemp = Split(document.formname.hVouDetails.value,":")

		if document.formname.selVoucherType.value="OSC" or document.formname.selVoucherType.value="OPC" then
			IF UBound(sTemp) <> 0 Then
				document.formname.action="VouCNOthersEntry.asp"
			Else
				document.formname.action="VouCNOthCommEntry.asp"
			End IF
		end if
		DisButt()
		document.formname.submit
	end if
End function

function validate()
Dim iCounter,sInvTemp
	if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")
		validate= false
		exit function
	end if
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select Credit Note")
		validate= false
		exit function
	end if
	if document.formname.selVoucherType.selectedIndex<1 then
		MsgBox ("Select Voucher type")
		validate=false
		exit function
	end if
	if document.formname.selVoucherType.selectedIndex=1 then
		if document.formname.selInvoiceNo.selectedIndex<1 then
			'MsgBox ("Select Invoice No")
			'validate=false
			'exit function
		end if
	end if
	document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
	document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text
	'For Sales Commission Multiple Entries checking
	IF document.formname.selInvoiceNo.size > 1 Then
		For iCounter = 0 To document.formname.selInvoiceNo.length - 1
			IF (document.formname.selInvoiceNo.options(iCounter).selected) Then
				sInvTemp = sInvTemp&":"&document.formname.selInvoiceNo.options(iCounter).text
			End IF
		Next

		sInvTemp = Mid(sInvTemp,2)
		document.formname.hVouDetails.value = sInvTemp
	Else
		document.formname.hVouDetails.value =document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.selectedIndex).text
	End IF

	validate=true
End function

function selParty(objAcc)
dim sOrgId,sPartyType,sBookCode
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim sSelVal

sRetVal2 = "0:0"
sOrgId=document.formname.selUnitId.value
sBookCode=document.formname.selBook.value

document.formname.selVoucherType.length=1
IF (document.formname.selInvoiceNo.multiple) Then
	document.formname.selInvoiceNo.length = 0
Else
	document.formname.selInvoiceNo.length = 1
End IF
document.formname.txtPartyName.value=""
document.formname.hPartyCode.value=""

sPartyType=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
if objAcc.selectedIndex >0 then
	OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")
	while UBound(arrTemp) = 0
		OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend

	IF UBound(arrTemp) <= 1 Then Exit Function
	sRetValue = OutValue
	sTemp = Split(sRetValue,":")
	sParTy = sTemp(4)
	sParSubType = sTemp(3)
	sParCode = sTemp(1)
	sPartyName = sTemp(0)

	'MsgBox sRetValue

	sRetVal2 = sRetVal2&":0"
	GetPartyHeadXml sParCode,sPartyName,sRetVal2
	Set nodAccHead = AccHeadData.documentElement

	if nodAccHead.hasChildNodes then
		'User Has Selected a Party
		For Each HeaderNode In nodAccHead.childNodes
			document.formname.hPartyCode.value=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
			document.formname.txtPartyName.value=HeaderNode.Attributes.Item(3).nodeValue
		next
		if Cstr(objAcc.value) <> "CR?1"   Then
			'Process For Sales Return
			document.formname.selVoucherType.length=1
			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Sales Return"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="SR"

			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Sales Invoices"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="OIS"

			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Others"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="OT"


			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Purchase Invoices"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="OIP"


			set objhttp = CreateObject("MSXML2.XMLHTTP")
			'MsgBox document.formname.hPartyCode.value

			objhttp.Open "GET","XMLSalInvDetails.asp?BookCode=05&OrgId="&sOrgId&"&PartyCode=" & document.formname.hPartyCode.value , false
			objhttp.send
			'MsgBox objhttp.responseText

			if objhttp.responseXML.xml <> "" then
				UnitBookData.loadXML objhttp.responseXML.xml
				Set Root = UnitBookData.documentElement
				For Each HeaderNode In Root.childNodes
					sSelVal = HeaderNode.Attributes.getNamedItem("TransNo").value
					sSelVal = sSelVal & ":"& HeaderNode.Attributes.getNamedItem("TotalCrDrValue").value

					document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
					document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text = HeaderNode.Attributes.getNamedItem("InvDetails").value
					document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = sSelVal
				next
				'document.formname.selVoucherType.selectedIndex=2
			end if	'End of Party Has Any Invoice Check
			document.formname.selInvoiceNo.multiple = false

		else
			'Process For Sales Commision
			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Sales Commision"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="SC"
		'	document.formname.btnAmend.disabled = True
			document.formname.selInvoiceNo.multiple = true
			document.formname.selInvoiceNo.size = 4
			document.formname.selInvoiceNo.length = 0

			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Others"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="OT"

			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Sales Invoices"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="OSC"

			document.formname.selVoucherType.length = document.formname.selVoucherType.length+1
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).text = "Purchase Invoices"
			document.formname.selVoucherType.options(document.formname.selVoucherType.length-1).Value ="OPC"


			'MsgBox document.formname.hPartyCode.value

			set objhttp = CreateObject("MSXML2.XMLHTTP")
			objhttp.Open "GET","XMLCommisionDetails.asp?OrgId="&sOrgId&"&AgentCode=" & document.formname.hPartyCode.value , false
			objhttp.send
			if objhttp.responseXML.xml <> "" then
				UnitBookData.loadXML objhttp.responseXML.xml
				Set Root = UnitBookData.documentElement
				For Each HeaderNode In Root.childNodes
					document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
					document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
					document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
				next
				'document.formname.selVoucherType.selectedIndex=1
			end if	'End of Agent Has Any Commision Check



		end if 'End of Party Type is Agent or party Check
	else
		objAcc.selectedIndex=0
	end if 'End of Party Head Processing
End if 'End of If any Part Type Selected Check
End function

function setInvoiceNo()
	if document.formname.selVoucherType.value = "SR" Then
		document.formname.selInvoiceNo.length = 1
		document.formname.selInvoiceNo.size = 1
		PopOthInv("S")
	End IF

	if document.formname.selVoucherType.value="OT" then
		document.formname.selInvoiceNo.disabled=true
'		document.formname.btnAmend.disabled = false
	Else
		document.formname.selInvoiceNo.disabled=false
	end if

	IF document.formname.selVoucherType.value = "SC" Then
		popCommVouchers()
'		document.formname.btnAmend.disabled = true
	Else
'		document.formname.btnAmend.disabled = false
	End IF

	IF document.formname.selVoucherType.value="OIS" then
		document.formname.selInvoiceNo.length = 1
		document.formname.selInvoiceNo.size = 1
		'document.formname.btnAmend.disabled = true
		PopOthInv("S")
	End IF

	IF document.formname.selVoucherType.value="OIP" then
		'document.formname.selInvoiceNo.length = 1
		'document.formname.selInvoiceNo.size = 1
		'document.formname.btnAmend.disabled = true
		PopOthInv("P")
	End IF

	'Sales Commission Not Marked For Commission Payment
	IF document.formname.selVoucherType.value="OSC" then
		'document.formname.selInvoiceNo.length = 1
		'document.formname.selInvoiceNo.size = 1
		'document.formname.btnAmend.disabled = true
		PopOthComm("S")
	End IF

	'Purchase Commission Not Marked For Commission Payment
	IF document.formname.selVoucherType.value="OPC" then
		document.formname.selInvoiceNo.length = 1
		document.formname.selInvoiceNo.size = 1
		PopOthComm("P")
	End IF


End function

Function PopOthInv(sType)
	Dim Objhttp,sOrgId,sSelVal

	sOrgId = document.formname.selUnitId.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	If trim(document.formname.selVoucherType.options( document.formname.selVoucherType.selectedIndex).value) = "SR" then
		sVouType = "SR"
	Else
		sVouType = "O"
	End IF
	objhttp.Open "GET","XMLSalInvDetails.asp?BookCode=45&OrgId="&sOrgId&"&Type="&sVouType&"&PartyCode=" & document.formname.hPartyCode.value&"&sCallTy="&sVouType , false
	objhttp.send
	'alert objhttp.responsetext

	'alert objhttp.responseXML.xml
	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			IF CStr(HeaderNode.Attributes.Item(5).nodeValue) = CStr(sType) Then
				sSelVal = HeaderNode.Attributes.Item(0).nodeValue & ":"& HeaderNode.Attributes.getNamedItem("FromValue").value
				document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
				document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
				document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = sSelVal
			End IF
		next

	end if	'End of Agent Has Any Commision Check
End Function

Function PopOthComm(sType)
	Dim Objhttp,sOrgId,sSelVal

	sOrgId = document.formname.selUnitId.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	IF CStr(sType) = "S" Then
		objhttp.Open "GET","XMLInvDetails.asp?BookCode=05&OrgId="&sOrgId&"&PartyCode=" & document.formname.hPartyCode.value&"&sCallTy=S" , false
	Else
		objhttp.Open "GET","XMLInvDetails.asp?BookCode=04&OrgId="&sOrgId&"&PartyCode=" & document.formname.hPartyCode.value&"&sCallTy=S" , false
	End IF
	objhttp.send


	if objhttp.responseXML.xml <> "" then
		CommData.loadXML objhttp.responseXML.xml
		Set Root = CommData.documentElement
		document.formname.selInvoiceNo.length = 0
		document.formname.selInvoiceNo.multiple = True

		IF CDbl(Root.childNodes.length) > 8 Then
			document.formname.selInvoiceNo.size = 8
		Else
			document.formname.selInvoiceNo.size = Root.childNodes.length
		End IF

		For Each HeaderNode In Root.childNodes
			sSelVal = HeaderNode.Attributes.Item(0).nodeValue
			document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = sSelVal
		next

	Else
		document.formname.selInvoiceNo.length = 0
		document.formname.selInvoiceNo.size = 4
		document.formname.selInvoiceNo.multiple = True

	end if	'End of Agent Has Any Commision Check
End Function

Function popVoucherNo(sVal)
Dim iPurTy,iParTy,iParCode,sTemparr,sVouchTy

	IF document.formname.selUnitId.selectedIndex = 0 Then
		MsgBox "Select Unit "
		document.formname.selUnitId.focus()
		Exit Function
	End IF

	IF document.formname.selBook.selectedIndex = 0 Then
		MsgBox "Select Book "
		document.formname.selBook.focus()
		Exit Function
	End IF

	sCallTy = sVal
	sOrgId = document.formname.selUnitId.value
	iBookId = "07"
	iBookNo = document.formname.selBook.value
	iPurTy = "0"
	iParTy = document.formname.selPartyType.value
	iParCode = document.formname.hPartyCode.value
	sTrans = "CNR"
	IF document.formname.selVoucherType.selectedIndex = 0 Then
		MsgBox "Select Voucher Type "
		document.formname.selVoucherType.focus()
		Exit Function
	End IF
	sVouchTy = document.formname.selVoucherType.value


	IF CStr(sVouchTy) = "SR" and CStr(sVal) = "C" Then
		'Exit Function
	End IF

	IF CStr(sVouchTy) = "SR" Then
		sTemp = showModalDialog("VouchSelForCNDN.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans&"&sPurTy="&iPurTy&"&sParTy="&iParTy&"&iParCode="&iParCode&"&VouchTy="&sVouchTy,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
		'document.formname.btnAmend.disabled = True
	Else
		sTemp = showModalDialog("VouchSelForSalPur.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans&"&sPurTy="&iPurTy&"&sParTy="&iParTy&"&iParCode="&iParCode&"&VouchTy="&sVouchTy,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	End IF
	IF CStr(sTemp) = "0" Then
		Exit Function
	End IF

	'MsgBox sTemp
	IF CStr(sVouchTy) = "SR" Then
		sTemparr = Split(sTemp,"~")
		document.formname.TransNo.value = sTemparr(0)
		document.formname.hTransNo.value = sTemparr(0)
		document.formname.txtVouchNo.value = sTemparr(1)
	Else
		sTemparr = Split(sTemp,"~")
		document.formname.TransNo.value = sTemparr(0)
		document.formname.hTransNo.value = sTemparr(0)
		document.formname.txtVouchNo.value = sTemparr(2)
	End IF

	'IF CStr(sVal) = "A" Then
	'	document.formname.btnAmend.disabled = True
	'	document.formname.btnDel.disabled = True
	'Else
	'	document.formname.btnAmend.disabled = False
	'	document.formname.btnDel.disabled = False
	'End IF

	'IF CStr(sVouchTy) = "SC" Then
	'	document.formname.btnAmend.disabled = True
	'End IF

End Function

Function popCommVouchers() 'Display of Commission Vouchers
	Dim sOrgId,objhttp
	sOrgId = document.formname.selUnitId.value

	document.formname.selInvoiceNo.length = 0
	document.formname.selInvoiceNo.multiple = True

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLCommisionDetails.asp?OrgId="&sOrgId&"&AgentCode=" & document.formname.hPartyCode.value , false
	objhttp.send
	'alert objhttp.responseText
	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		IF CDbl(Root.childNodes.length) > 8 Then
			document.formname.selInvoiceNo.size = 8
		Else
			document.formname.selInvoiceNo.size = Root.childNodes.length
		End IF

		For Each HeaderNode In Root.childNodes
			document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
		next
		'document.formname.selVoucherType.selectedIndex=1
	end if	'End of Agent Has Any Commision Check
End Function

Function VouView()
	'MsgBox document.formname.selVoucherType.value = "OIS"
	IF validate() Then
		IF document.formname.selVoucherType.value = "SR" Then
			document.formname.action = "VouCNSalReturnDisplay.asp"
			document.formname.submit()
		Elseif document.formname.selVoucherType.value = "SC" Then
			document.formname.action = "VouCNCommisionView.asp"
			DisButt()
			document.formname.submit()
		Elseif document.formname.selVoucherType.value = "OIS" Then
			document.formname.action = "VouCNOtherInvDisplay.asp"
			DisButt()
			document.formname.submit()
		Elseif document.formname.selVoucherType.value = "OIP" Then
			document.formname.action = "VouDNPurReturnDisplay.asp"
			DisButt()
			document.formname.submit()
		Else
			document.formname.action = "VouCNOtherView.asp"
			DisButt()
			document.formname.submit()
		End IF
	End IF
End Function

Function VouDel()
	IF validate() Then
		document.formname.action = "VouCNCommDelView.asp"
		DisButt()
		document.formname.submit()
	End IF
End Function

Function VouAmend()

	if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")

		exit function
	end if
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select Credit Note")

		exit function
	end if
	IF document.formname.txtVouchNo.value = "" Then
		Msgbox "Select Voucher No "
		exit function
	End IF

	document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
	document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text

	'MsgBox document.formname.selVoucherType.value

	IF document.formname.selVoucherType.value = "SC" Then
		document.formname.action = "VouCNCommissionAmd.asp"
	Elseif document.formname.selVoucherType.value = "SR" Then
		document.formname.action = "VouCNSalRetEntryAmd.asp"
		'document.formname.action = "VouCNSalReturnEntry2.asp"
	Elseif document.formname.selVoucherType.value = "OIP" Then
		document.formname.action = "VouCNPurInvAmd.asp"
	Elseif document.formname.selVoucherType.value = "OIS" Then
		document.formname.action = "VouCNSalInvAmd.asp"
	Else
		document.formname.action = "VouCNCommAmend.asp"
	End IF
	DisButt()
	document.formname.submit()

End Function

Function DisButt()
	'document.formname.btnAmend.disabled = True
	document.formname.btnCreate.disabled = True
	'document.formname.btnDel.disabled = True
	'document.formname.btnView.disabled = True
End Function

</script>
<script language="javascript">
function openItmsVoucherDialog(url, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, "", features, callback);
		return;
	}
	alert("The compatibility script is still loading. Please try again.");
}

function popVoucherNo(sVal) {
	var form = document.formname;
	var sCallTy;
	var sOrgId;
	var iBookId;
	var iBookNo;
	var iPurTy;
	var iParTy;
	var iParCode;
	var sTrans;
	var sVouchTy;
	var page;
	var url;
	if (form.selUnitId.selectedIndex === 0) {
		alert("Select Unit ");
		form.selUnitId.focus();
		return;
	}
	if (form.selBook.selectedIndex === 0) {
		alert("Select Book ");
		form.selBook.focus();
		return;
	}
	if (form.selVoucherType.selectedIndex === 0) {
		alert("Select Voucher Type ");
		form.selVoucherType.focus();
		return;
	}
	sCallTy = sVal;
	sOrgId = form.selUnitId.value;
	iBookId = "07";
	iBookNo = form.selBook.value;
	iPurTy = "0";
	iParTy = form.selPartyType.value;
	iParCode = form.hPartyCode.value;
	sTrans = "CNR";
	sVouchTy = form.selVoucherType.value;
	page = String(sVouchTy) === "SR" ? "VouchSelForCNDN.asp" : "VouchSelForSalPur.asp";
	url = page + "?flag=" + encodeURIComponent(sCallTy) +
		"&orgId=" + encodeURIComponent(sOrgId) +
		"&BookCode=" + encodeURIComponent(iBookId) +
		"&BookNo=" + encodeURIComponent(iBookNo) +
		"&TransType=" + encodeURIComponent(sTrans) +
		"&sPurTy=" + encodeURIComponent(iPurTy) +
		"&sParTy=" + encodeURIComponent(iParTy) +
		"&iParCode=" + encodeURIComponent(iParCode) +
		"&VouchTy=" + encodeURIComponent(sVouchTy);
	openItmsVoucherDialog(url, "dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (sTemp) {
		var parts;
		if (String(sTemp || "0") === "0") {
			return;
		}
		parts = String(sTemp).split("~");
		form.TransNo.value = parts[0] || "";
		form.hTransNo.value = parts[0] || "";
		form.txtVouchNo.value = String(sVouchTy) === "SR" ? (parts[1] || "") : (parts[2] || "");
	});
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="horgName" value="">
<input type="hidden" name="TransNo" value="">
<input type="hidden" name="hTransNo" value="">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hVouDetails" value="">
<input type="hidden" name="hFromApp" value="C">
<input type="hidden" name="hChkVal" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Credit Note Voucher Entry
		</td>
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
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<!--td class="TabCell" valign="bottom" align="center" width="90">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Adjustments
											</td>
										</tr>
									</table>
								</td-->
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<tr><td align="center">Voucher</td>
								  	</tr>
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
						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="168">Organization </td>
                            <td class="FieldCell" >
                             <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook()">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationListDB%>
                              </select> &nbsp;

								<input type="Checkbox" name="ChkGJ" class="FormElem" checked onchange="DisplayBook()">Create as GJ Voucher
                              </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Book</td>
                            <td class="FieldCell">
                            <select size="1" name="selBook" class="FormElem">
                        <option value="S">Select Book</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Select Party Type</td>
                            <td class="FieldCell">
                            <select size="1" name="selPartyType" class="FormElem" onChange="selParty(this)">
								<option value="S">Select Party Type</option>
								</select>
								</td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Party Name</td>
                            <td class="FieldCell" colspan="3"> <input type="text" name="txtPartyName" size="40" class="FormElem"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Voucher Type</td>
                            <td class="FieldCell">
                            <select size="1" name="selVoucherType" class="FormElem" onChange="setInvoiceNo()">
							<option value="S">Select Voucher Type</option>
							<option value="SR">Sales Return</option>
							<option value="SC">Sales Commision</option>
							<option value="OT">Others</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Select Invoice Number</td>
                            <td class="FieldCell">
                            <select size="1" name="selInvoiceNo" class="FormElem">
							<option value="S">Select Invoice Number</option>
                            </select></td>
                                </tr>
                                <!--tr>
									 <td class="FieldCell" width="168">Voucher Number</td>
									 <td class="FieldCell">
									 <input type="text" name="txtVouchNo" size="20" class="FormElem" readonly>
									 <
									  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<a href="javascript:popVoucherNo('C')">
									<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<a href="javascript:popVoucherNo('A')">
									<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers"></a> >
									 </td>
                            </tr-->
                                    </table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>

							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">

                                                                <input type="button" value="Create" onClick="VouCreate()" name="btnCreate" class="ActionButton" >
                                                                <!--input type="button" value="View" name="btnView" onClick="VouView()" class="ActionButton">
                                                                <input type="button" value="Amendment" name="btnAmend" onClick="VouAmend()"  class="ActionButtonX">
                                                                <input type="button" value="Delete" name="btnDel" onClick="VouDel()" class="ActionButton" -->
                                                                <input type="reset" value="Reset" name="B10" class="ActionButton" >
                                                                <input type="button" value="New" onClick="NewPage()" name="btnNew" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" width="5">
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
