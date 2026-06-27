<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 16, 2003
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML id="OutData"><Root/></xml>
<XML ID="CommData"><Book/></XML>
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

Function DisplayBook(objUnit)
dim iUnitNo,arrTemp
dim Root
	document.formname.selBook.options.length = 1

	if objUnit.selectedIndex <> "0" then
		iUnitNo= objUnit(objUnit.selectedIndex).value

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=06&orgID=" & iUnitNo , false
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
	end if
end Function
Function NewPage()
	document.formname.action = "DebitNoteNewPage.asp"
	document.formname.submit
End Function

Function VouCreate
	Dim sTemp
	'MsgBox document.formname.selVoucherType.value
	'Exit Function

	if validate then
		'if document.formname.selVoucherType.value="PR" then
		'	document.formname.action="VouDNPurReturnEntry.asp"
		'elseif document.formname.selVoucherType.value="OT" then
		'	document.formname.action="VouDNOthersEntry.asp"
		'elseif document.formname.selVoucherType.value="SC" then
		'	document.formname.action="VouDNCommisionEntry.asp"
		'End IF

		'sTemp = Split(document.formname.hVouDetails.value,":")


		'if document.formname.selVoucherType.value = "OIP" Then
		'
		'	IF document.formname.selPartyType.value <> "CR?1" and document.formname.selPartyType.value <> "DR?2" Then
		'		document.formname.action="VouDNOtherInvEntry.asp"
		'	Else
		'		IF UBound(sTemp) <> 0 Then
		'			document.formname.action="VouDNOthersEntry.asp"
		'		Else
		'			document.formname.action="VouDNCommisionEntry.asp"
		'		End IF
		'	End IF
'
'		Elseif document.formname.selVoucherType.value = "OIS" Then
'			IF document.formname.selPartyType.value <> "CR?1" and document.formname.selPartyType.value <> "DR?2" Then
'				document.formname.action="VouDNSalInvEntry.asp"
'			Else
'				IF UBound(sTemp) > 1 Then
'					document.formname.action="VouDNOthersEntry.asp"
'				Else
'					document.formname.action="VouDNCommisionEntry.asp"
'				End IF
'			End IF
'		end if

'		DisButt()

		if document.formname.selVoucherType.value="OT" then
			document.formname.action="VouDNOthersEntry.asp"
		elseif document.formname.selVoucherType.value="OP" then
			document.formname.action="VouDNOtherInvEntry.asp"
		elseif document.formname.selVoucherType.value="SI" then
			document.formname.action="VouDNSalInvEntry.asp"
		End IF

		document.formname.submit
	end if
End function

function validate()
	if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")
		validate= false
		exit function
	end if
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select Debit Note")
		validate= false
		exit function
	end if
	if document.formname.selVoucherType.selectedIndex<1 then
		MsgBox ("Select Voucher type")
		validate=false
		exit function
	end if

	if document.formname.selVoucherType.value <> "OT" Then
		if document.formname.selInvoiceNo.selectedIndex<1 and document.formname.selRefNo.selectedIndex<1 then
			MsgBox ("Select Invoice No")
			validate=false
			exit function
		end if
	end if

	document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
	document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text

	IF document.formname.selInvoiceNo.size > 1 Then
		For iCounter = 0 To document.formname.selInvoiceNo.length - 1
			IF (document.formname.selInvoiceNo.options(iCounter).selected) Then
				sInvTemp = sInvTemp&":"&document.formname.selInvoiceNo.options(iCounter).text
			End IF
		Next

		sInvTemp = Mid(sInvTemp,2)
		document.formname.hVouDetails.value = sInvTemp
	Else
		document.formname.hVouDetails.value = document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.selectedIndex).text
	End IF

	'document.formname.hVouDetails.value=document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.selectedIndex).text

	validate=true
End function

function selParty(objAcc)
dim sOrgId,sPartyType,sBookCode,sSelVal
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp

sRetVal2 = "0:0"
sOrgId=document.formname.selUnitId.value
sBookCode=document.formname.selBook.value

'document.formname.selVoucherType.length=1
document.formname.selInvoiceNo.length=1
document.formname.selRefNo.length=1
document.formname.txtPartyName.value=""
document.formname.hPartyCode.value=""

sPartyType=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
if objAcc.selectedIndex >0 then
	'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="&sOrgId&"&Party="&sPartyType,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")

	while UBound(arrTemp) = 0
		OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
	IF UBound(arrTemp) <=1 Then Exit Function
	sRetValue = OutValue
	sTemp = Split(sRetValue,":")
	sParTy = sTemp(4)
	sParSubType = sTemp(3)
	sParCode = sTemp(1)
	sPartyName = sTemp(0)
	sRetVal2 = sRetVal2&":0"
	GetPartyHeadXml sParCode,sPartyName,sRetVal2

	Set nodAccHead = AccHeadData.documentElement


	if nodAccHead.hasChildNodes then
		For Each HeaderNode In nodAccHead.childNodes
			document.formname.hPartyCode.value=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
			document.formname.txtPartyName.value=HeaderNode.Attributes.Item(3).nodeValue
		next
	else
		objAcc.selectedIndex=0
	end if 'End of Party Head Processing
End if 'End of If any Part Type Selected Check
End function

Function PopComm()
	Dim Objhttp,sOrgId
	sOrgId = document.formname.selUnitId.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLCommisionDetails.asp?OrgId="&sOrgId&"&AgentCode=" & document.formname.hPartyCode.value , false
	objhttp.send
	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text =  HeaderNode.Attributes.Item(2).nodeValue
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
		next
	end if	'End of Agent Has Any Commision Check
End Function

Function PopOthInv(sCallty)
	Dim Objhttp,sOrgId
	'sCallty = "O"
	'MsgBox "Frm "
	sOrgId = document.formname.selUnitId.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	'IF CStr(sCallty) <> "O" Then
	'	objhttp.Open "GET","XMLSalInvDetails.asp?BookCode=04&OrgId="&sOrgId&"&PartyCode=" & document.formname.hPartyCode.value&"&sCallTy="&sCallty , false
	'Else
		objhttp.Open "GET","XMLSalInvDetails.asp?BookCode=04&OrgId="&sOrgId&"&PartyCode=" & document.formname.hPartyCode.value&"&sCallTy=O" , false
	'End IF

	objhttp.send
	'MsgBox objhttp.responseText

	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			IF trim(HeaderNode.NodeName)= trim("SalInv") then
				'IF CStr(sCallty) <> "O" Then
				'	document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text =  HeaderNode.Attributes.getNamedItem("InvDetails").value
				'Else
				'	document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text =  HeaderNode.Attributes.getNamedItem("InvDetails").value
				'End IF

				IF CStr(sCallty) = CStr(HeaderNode.Attributes.Item(5).nodeValue) Then
					document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
					document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text =  HeaderNode.Attributes.Item(2).nodeValue
					document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
				End IF
			End IF ' IF trim(HeaderNode.NodeName)= trim("SalInv") then
			'document.formname.selRefNo.size = 5
			'document.formname.selRefNo.multiple = True

			IF trim(HeaderNode.NodeName)= trim("Invoice") then
				sAct = trim(HeaderNode.getAttribute("ActNo"))
				sInvNo = trim(HeaderNode.getAttribute("InvNo"))
				sInvDate = trim(HeaderNode.getAttribute("InvDate"))

				sVal = sAct &"-"& sInvNo &"-"& sInvDate
				document.formname.selRefNo.length = document.formname.selRefNo.length+1

				document.formname.selRefNo.options(document.formname.selRefNo.length-1).text =  sVal
				document.formname.selRefNo.options(document.formname.selRefNo.length-1).Value = sAct

			End IF 'IF trim(HeaderNode.NodeName)= trim("Invoice") then

		next
	end if	'End of Agent Has Any Commision Check
End Function


function setInvoiceNo()

	if document.formname.selVoucherType.value="OT" then
		document.formname.selInvoiceNo.disabled=true
		document.formname.selRefNo.disabled=true
		document.formname.btnAmend.disabled = false
	Elseif document.formname.selVoucherType.value="OP" then
		document.formname.selInvoiceNo.disabled=false
		document.formname.selRefNo.disabled=false
		document.formname.selInvoiceNo.length = 1
		document.formname.selRefNo.length = 1
		'IF document.formname.selPartyType.value = "CR?1" or document.formname.selPartyType.value = "DR?2" Then
		'	PopOthComm("P")
		'Else
			PopOthInv("P")
		'End IF
	Elseif document.formname.selVoucherType.value="SI" then
		document.formname.selInvoiceNo.disabled=false
		document.formname.selRefNo.disabled=false
		document.formname.selInvoiceNo.length = 1
		document.formname.selRefNo.length = 1
		'IF document.formname.selPartyType.value = "CR?1" or document.formname.selPartyType.value = "DR?2" Then
		'	PopOthComm("S")
		'Else
			PopOthInv("S")
		'End IF

	Elseif document.formname.selVoucherType.value="SC" then
		document.formname.selInvoiceNo.disabled=false
		document.formname.selRefNo.disabled=false
		document.formname.selInvoiceNo.length = 1
		document.formname.selRefNo.length = 1
		PopComm()
	Else
		document.formname.selInvoiceNo.disabled=false
		document.formname.selRefNo.disabled=false
		document.formname.selInvoiceNo.length = 1
		document.formname.selRefNo.length = 1
		PopOthInv("P")
	End IF
End function

Function popVoucherNo(sVal)
Dim iPurTy,iParTy,iParCode,sTemparr,sVouchTy

	IF document.formname.selUnitId.selectedIndex = 0 then
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
	iBookId = "06"
	iBookNo = document.formname.selBook.value
	iPurTy = "0"
	iParTy = document.formname.selPartyType.value
	iParCode = document.formname.hPartyCode.value
	sTrans = "DNR"
	IF document.formname.selVoucherType.selectedIndex = 0 Then
		MsgBox "Select Voucher Type "
		document.formname.selVoucherType.focus()
		Exit Function
	End IF
	sVouchTy = document.formname.selVoucherType.value



	IF CStr(sVouchTy) = "PR" Then
		sTemp = showModalDialog("VouchSelForCNDN.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans&"&sPurTy="&iPurTy&"&sParTy="&iParTy&"&iParCode="&iParCode&"&VouchTy="&sVouchTy,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	Else
		sTemp = showModalDialog("VouchSelForSalPur.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans&"&sPurTy="&iPurTy&"&sParTy="&iParTy&"&iParCode="&iParCode&"&VouchTy="&sVouchTy,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	End IF
	IF CStr(sTemp) = "0" Then
		Exit Function
	End IF

	'MsgBox sTemp

	sTemparr = Split(sTemp,"~")
	document.formname.hTransNo.value = sTemparr(0)
	document.formname.txtVouchNo.value = sTemparr(2)

	IF CStr(sVal) = "A" Then
		document.formname.btnAmend.disabled = True
		document.formname.btnDelete.disabled = True
	Else
		document.formname.btnAmend.disabled = False
		document.formname.btnDelete.disabled = False
	End IF

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
		document.formname.selInvoiceNo.length = 0

		document.formname.selInvoiceNo.size = 8
		document.formname.selInvoiceNo.multiple = True


		CommData.loadXML objhttp.responseXML.xml
		Set Root = CommData.documentElement

		For Each HeaderNode In Root.childNodes
			sSelVal = HeaderNode.Attributes.Item(0).nodeValue
			document.formname.selInvoiceNo.length = document.formname.selInvoiceNo.length+1
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
			document.formname.selInvoiceNo.options(document.formname.selInvoiceNo.length-1).Value = sSelVal
		next

	end if	'End of Agent Has Any Commision Check
End Function



Function DisButt()
	'document.formname.btnAmend.disabled = True
	document.formname.btnCreate.disabled = True
	'document.formname.btnDelete.disabled = True
	'document.formname.btnView.disabled = True
End Function

Function SelInvChoice()
	If trim(document.formname.selInvoiceNo.value) = "S" then
		document.formname.selRefNo.disabled = False
	Else
		document.formname.selRefNo.disabled = True
	End IF
End Function

Function SelRefChoice()
	If trim(document.formname.selRefNo.value) = "S" then
		document.formname.selInvoiceNo.disabled = False
	Else
		document.formname.selInvoiceNo.disabled = True
	End IF

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
	iBookId = "06";
	iBookNo = form.selBook.value;
	iPurTy = "0";
	iParTy = form.selPartyType.value;
	iParCode = form.hPartyCode.value;
	sTrans = "DNR";
	sVouchTy = form.selVoucherType.value;
	page = String(sVouchTy) === "PR" ? "VouchSelForCNDN.asp" : "VouchSelForSalPur.asp";
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
		form.hTransNo.value = parts[0] || "";
		form.txtVouchNo.value = parts[2] || "";
		if (String(sVal) === "A") {
			if (form.btnAmend) {
				form.btnAmend.disabled = true;
			}
			if (form.btnDelete) {
				form.btnDelete.disabled = true;
			}
		} else {
			if (form.btnAmend) {
				form.btnAmend.disabled = false;
			}
			if (form.btnDelete) {
				form.btnDelete.disabled = false;
			}
		}
	});
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="horgName" value="">
<input type="hidden" name="hTransNo" value="">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hVouDetails" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Debit Note Voucher Entry
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
								<td class="TabCell" valign="bottom" align="center" width="90">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Adjustments
											</td>
										</tr>
									</table>
								</td>
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
                            <td class="FieldCell">
                             <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationListDBWithVal("")%>
                              </select>
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
							<!--<option value="PR">Purchase Returns</option>-->
							<!--option value="OT">Others</option>
							<option value="OI">Other Invoices</option-->
							<option value="OT" Selected>Others</option>
							<option value="OP">Purchase Invoices</option>
							<option value="SI">Sales Invoices</option>

                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Select Reference Number</td>
                            <td class="FieldCell">
                            <select size="1" name="selInvoiceNo" class="FormElem" onchange="SelInvChoice()">
							<option value="S">Select Invoice </option>
                            </select></td>
                            <td class="FieldCell">Or</td>
                            <td class="FieldCell">
                            <select  name="selRefNo" class="FormElem" onchange="SelRefChoice()">
							<option value="S">Select Invoice </option>
                            </select></td>

                                </tr>
                                <tr>
                            <td class="FieldCell" width="168">Debit Note Number</td>
                            <td class="FieldCell">
								<input type="text" name="txtVouchNo" size="20" class="FormElem" readonly>
								<!--
								  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<a href="javascript:popVoucherNo('C')">
								<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<a href="javascript:popVoucherNo('A')">
								<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers"></a -->
                            </td>
                                </tr>

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
                                                                <!--input type="button" value="View" onClick="VouView()" name="btnView" class="ActionButton" >
                                                                <input type="button" value="Amendment" onClick="VouAmend()" name="btnAmend" class="ActionButtonX" >
                                                                <input type="button" value="Delete" onClick="VouDel()" name="btnDelete" class="ActionButton" -->
                                                                <input type="reset" value="Reset" name="btnReset" class="ActionButton" >
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
