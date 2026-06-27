<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	DebitNoteNewPage.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	S.Maheswari
	'Created On					:	Oct 22, 2008
	'Modified By                :   Ragavendran R
	'Modified On                :   Oct 24,2011
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
<xml id="PartyData"><Root /></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouTransactions.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<Script Language=vbscript>
'==============================================================================
Function FnInit()
	 'alert document.formname.hPartypeName.value
	sTemp = split(document.formname.hPartypeName.value,"?")
	'document.formname.selPartyType.options(document.formname.selPartyType.selectedIndex).value = document.formname.hPartypeName.value

	'document.formname.selPartyType.options(document.formname.selPartyType.selectedIndex).text = sTemp(2)
	document.formname.txtPartyName.value = document.formname.hParName.value

End Function
'==============================================================================
Function Validate()
	'document.formname.hUnitID.value	= document.formname.hOrgId.value
	' alert document.formname.hPartyCode.value
	sTemp = split(document.formname.hPartyCode.value,"?")
	document.formname.hParType.value = sTemp(0)
	document.formname.hSubParType.value = sTemp(1)
	document.formname.hParName.value = document.formname.txtPartyName.value 'sTemp(2)
	document.formname.hParCode.value = sTemp(3)

	document.formname.selPartyType.value   =  document.formname.hPartyCode.value
	'document.formname.hParCode.value  = document.formname.selPartyType.value
	'document.formname.txtPartyName.value =
	document.formname.submit
End Function
'==============================================================================
Function Voucher(sPara)
dim i,iCnt,sVouType,iBookNo,iChkCnt
	sVouType = sPara
	'alert  sVouType
	iCnt = document.formname.hCnt.value
	sUnit = document.formname.hUnitID.value
	'document.formname.horgName.value = document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text

	iBookNo  = document.formname.hBookNo.value
	iChkCnt = 0
	For i = 1 to iCnt
		if eval("document.formname.OptCriteria"&i).Checked = true then
			sVal = eval("document.formname.OptCriteria"&i).value
			iChkCnt = iChkCnt + 1
		end if
	Next
	IF iChkCnt = 0 then
		If trim(document.formname.selPartyType.value) = "S" or document.formname.txtPartyName.value = "" then
			MsgBox "Select 'Party Type' from Expand band to create Other type voucher.",6,"Debit Note Vouchers"
			exit function
		End IF
		sRetVal = MsgBox("This will Create 'Other' Type Voucher.Do U want to Continue?",4,"Debit Note Vouchers")
		If trim(sRetVal) = "7" then
			MsgBox "Select Any One Invoice Number"
			exit function
		Else
			document.formname.hInvVal.value = "OT"
		End If
	End IF
	document.formname.hVouType.value = sVouType
	IF trim(sVouType) = "GJ" then
		'iBookNo = 3
		'sBookName = "JOURNAL BOOK"
		set objhttp = CreateObject("MSXML2.XMLHTTP")
			objhttp.Open "GET","XMLGetOrgBookCountGJ.asp?BkCode=08&orgID=" & sUnit , false
			objhttp.send
			'alert(objhttp.responsetext)
			if objhttp.responseXML.xml <> "" then
				'alert(objhttp.responseXML.xml)
				UnitBookData.loadXML objhttp.responseXML.xml
				Set UnitRoot = UnitBookData.documentElement
					if strcomp(UnitRoot.getAttribute("Count"),"1") = 0 then
						for each subnode in UnitRoot.childNodes
							iBookNo = subnode.getAttribute("BookNumber")
							sBookName = subnode.getAttribute("BookName")
						next
					else
						OutValue = showModalDialog("BookPopUp.asp?Unit="&sUnit&"&VouType="&sVouType,"","dialogHeight:200px;dialogWidth:370px;center:Yes;help:No;resizable:No;status:No")
						sTemp = split(OutValue,"--")
						iBookNo = sTemp(0)
						sBookName = sTemp(1)
						if iBookNo = 0 then exit function
					end if
				'document.formname.hChkVal.value = "Y"
			end if

	Else
	set objhttp = CreateObject("MSXML2.XMLHTTP")
			objhttp.Open "GET","XMLGetOrgBookCountGJ.asp?BkCode=06&orgID=" & sUnit , false
			objhttp.send
			'alert(objhttp.responsetext)
			if objhttp.responseXML.xml <> "" then
				'alert(objhttp.responseXML.xml)
				UnitBookData.loadXML objhttp.responseXML.xml
				Set UnitRoot = UnitBookData.documentElement
					if strcomp(UnitRoot.getAttribute("Count"),"1") = 0 then
						for each subnode in UnitRoot.childNodes
							iBookNo = subnode.getAttribute("BookNumber")
							sBookName = subnode.getAttribute("BookName")
						next
					else
						OutValue = showModalDialog("BookPopUp.asp?Unit="&sUnit&"&VouType="&sVouType,"","dialogHeight:200px;dialogWidth:370px;center:Yes;help:No;resizable:No;status:No")
						sTemp = split(OutValue,"--")
						iBookNo = sTemp(0)
						sBookName = sTemp(1)
						if iBookNo = 0 then exit function
					end if
				'document.formname.hChkVal.value = "Y"
			end if
		'showModalDialog "CNDNOthVouchView_San.asp?TransNo="&arrTemp(0),"","dialogHeight:430px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No"
	'	OutValue = showModalDialog("BookPopUp.asp?Unit="&sUnit&"&VouType="&sVouType,"","dialogHeight:200px;dialogWidth:370px;center:Yes;help:No;resizable:No;status:No")
	 '	'alert OutValue
	'	sTemp = split(OutValue,"--")
	'	iBookNo = sTemp(0)
	'	sBookName = sTemp(1)
	End IF

	document.formname.selBook.value = iBookNo
	document.formname.selInvoiceNo.value = sVal
	'alert (document.formname.hInvVal.value)
	IF trim(document.formname.hInvVal.value) = "SI" then
		document.formname.action = "VouDNSalInvEntry.asp?CallFrom="&sVouType
	ElseIF trim(document.formname.hInvVal.value) = "PI" then
		document.formname.action = "VouDNOtherInvEntry.asp?CallFrom="&sVouType
	ElseIF trim(document.formname.hInvVal.value) = "OT" then
		document.formname.action = "VouDNOthersEntry.asp?CallFrom="&sVouType
    ElseIF trim(document.formname.hInvVal.value) = "MI" then
        document.formname.action = "VouDNThrMiscPay.asp?CallFrom="&sVouType
	End IF
	document.formname.submit

End Function
'==============================================================================
Function SetVal(Obj)
dim i,iCnt,sChk,sChkVal,sTempVal
	iCnt = document.formname.hCnt.value
	sChk = Obj.checked
	sChkVal =Obj.Value
	 ' alert iCnt &"===="& sChkVal &"===="& sTempval
	For i = 1 to iCnt
		If eval("document.formname.Optcriteria"&i).checked = True then
			sTempVal = eval("document.formname.Optcriteria"&i).value
			If trim(sChkVal) <> trim(sTempVal) then
				eval("document.formname.Optcriteria"&i).checked = false
			End IF
		End IF
	Next
End Function
'==============================================================================
FUNCTION popPartyType()

	dim iHeadCount

	iUnitNo=document.formname.hUnitID.value
	'iBkNo=document.formname.selBook.value

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

			IF document.formname.hPartyCode.value <> "" then
			'	alert "AAA"
			End IF
			iCounter=CDbl(iCounter)+1
		next
	end if

END FUNCTION
'==============================================================================
function selParty(objAcc)
dim sOrgId,sPartyType,sBookCode
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim sSelVal
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

	sRetVal2 = "0:0"
	sOrgId=document.formname.hUnitID.value
	'sBookCode=document.formname.selBook.value
	'document.formname.selVoucherType.length=1

	document.formname.txtPartyName.value=""
	document.formname.hPartyCode.value=""

	sPartyType=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
	if objAcc.selectedIndex >0 then

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

		'MsgBox sRetValue
'		alert sParTy  &" @ " & sParSubType &" @ " &   sParCode &" @ " &sPartyName
		document.formname.hParType.value = sParTy
		document.formname.hSubParType.value = sParSubType
		document.formname.hParCode.value = sParCode
		document.formname.hParName.value = sPartyName
		'document.formname.hPartyCode.value=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue&"?"&sParCode&"?"&sPartyName
		sRetVal2 = sRetVal2&":0"
		GetPartyHeadXml sParCode,sPartyName,sRetVal2

		Set nodAccHead = AccHeadData.documentElement

		if nodAccHead.hasChildNodes then
			'User Has Selected a Party
			For Each HeaderNode In nodAccHead.childNodes
				document.formname.hPartyCode.value=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
				document.formname.txtPartyName.value=HeaderNode.Attributes.Item(3).nodeValue
			next
		else
			objAcc.selectedIndex=0
		end if 'End of Party Head Processing
	End if 'End of If any Part Type Selected Check
End function
'==============================================================================
Function setInvoiceNo(sObj)
Dim sVal,i
	sVal = sObj.value

	for i = 0 to 1
		  'alert document.formname.hInvVal.value &"=="& sVal
		if trim(document.formname.voutype(i).value) = trim(sVal) then
			document.formname.voutype(i).checked = true

			document.formname.hInvVal.value  = document.formname.voutype(i).value
		else
			document.formname.voutype(i).checked = False
		end if
	next
	document.formname.submit
End Function
'==============================================================================
Function ShowVouch(iCrTransNo)
	IF trim(document.formname.hInvVal.value) <> "PI" then
		showModalDialog "SalesVouchView_San.asp?TransNo="&iCrTransNo,"","dialogHeight:410px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No"
	Else
		showModalDialog "PurchaseVouchView_San.asp?TransNo="&iCrTransNo,"","dialogHeight:410px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No"
	End IF
	Exit Function
End Function
'==============================================================================
</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="FnInit();popPartyType()">
<%

	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

	Dim sFinPeriod,Objrs,Objrs1,Objrs2,iCnt,sSql,iCrTransNo,sOptType
	Dim dcrs,sUnitLID,sUnitLName,sUnitSName,sBankType,AccVoucherNo
	Dim sFormVal,sTemparr,sUnitID,sBookNo,sFrmDate,sToDate,sFrmAmt,sToAmt
	Dim sChkSR,sChkSI,sChkPI,sChkMI
	Dim sSelVouTy,sOrgId,sPartyType,sPartypeName ,sOrgName,sPayToRecdFrom


	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")

	'sOrgId = Request("selUnitId")
	sOrgId = Session("organizationcode")
	sOrgName = Session("OrgShortName")
	'sPartyType = Request("selPartyType")
	sPartyType = Request.form("hPartyCode")
	IF trim(Request.form("hPartyCode")) <> "" then
		sTemp  = split(Request.form("hPartyCode"),"?")
		sPartypeName = sTemp(2)
		sParName =  Request.form("hParName")
		'Response.Write "PartyType="&sPartyType &"--"& sParName
	End if
	If trim(sOrgId) = "" or  trim(sOrgId) = "0" then sOrgId = "010101"
	sSelVouTy = Request("voutype")
	IF trim(sSelVouTy) = "" then
		sSelVouTy = "SI"
		sChkSI = "Checked"
		sChkSR = ""
		sChkPI = ""
		sChkMI = ""
	End IF
	IF trim(sSelVouTy) = "SI" then
		sChkSI = "Checked"
		sChkSR = ""
		sChkPI = ""
		sChkMI = ""
	ElseIF trim(sSelVouTy) = "SR" then
		sChkSR = "Checked"
		sChkSI = ""
		sChkPI = ""
		sChkMI = ""
	ElseIF trim(sSelVouTy) = "PI" then
		sChkPI = "Checked"
		sChkSI = ""
		sChkSR = ""
		sChkMI = ""
    ElseIF Trim(sSelVouTy) = "MI" then
        sChkMI = "Checked"
        sChkPI = ""
		sChkSI = ""
		sChkSR = ""
	End IF
	'Response.Write "VouType="&sSelVouTy
	iCurrentPage=CInt(Request.Form("hPageSelection"))
	'iCnt=Request.Form("hCnt")

%>
	<form method="POST" name="formname" action="DebitNoteToCreate.asp">

	<input type=hidden name="TransNo" value="">
	<input type=hidden name="hTransNo" value="">
	<input type=hidden name="hInvVal" value="<%=sSelVouTy%>">
	<input type=hidden name="hUnitID" value="<%=sOrgId%>">
	<input type=hidden name="horgName" value="<%=sOrgName%>">
	<input type=hidden name="hPartyCode" value="">
	<input type=hidden name="hParType" value="">
	<input type=hidden name="hSubParType" value="">
	<input type=hidden name="hParCode" value="">
	<input type=hidden name="hParName" value="<%=sParName%>">
	<input type=hidden name="hPartypeName" value="<%=Request.form("hPartyCode")%>">
	<input type=hidden name="hBookNo" value="">
	<input type=hidden name="selInvoiceNo" value="">
	<input type=hidden name="selBook" value="">
	<input type=hidden name="hVouType" value="">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Debit Vouchers
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
<Input type=radio name=voutype value="SI" <%=sChkSI%> onclick=setInvoiceNo(this)>Sales Invoice&nbsp;
<%
'Response.Write ("<Input type=checkbox name=voutype value="SC" onclick=setInvoiceNo() >Sales Commission&nbsp;
'Response.Write ("<Input type=radio name=voutype value="SR"  onclick=setInvoiceNo(this) >Sales Return&nbsp;%>
<Input type=radio name=voutype value="PI" <%=sChkPI%> onclick=setInvoiceNo(this) >Purchase Invoice&nbsp;
<Input type=radio name=voutype value="MI" <%=sChkMI%> onclick=setInvoiceNo(this) >Misc Payments&nbsp;
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
	<td class="FieldCellSub">Unit Name</td>
	<td class="FieldCellSub" colspan="4">
	 <select size="1" name="selUnitId" class="FormElem"> <%'onChange="DisplayBook()"%>

	  	<%populateOrganizationListDB%>
    </select> &nbsp;
	</td>
</tr>-->
<tr>
<td class="FieldCellSub" width="168px">Select Party Type</td>
<td class="FieldCellSub">
<select size="1" name="selPartyType" class="FormElem" onChange="selParty(this)">
	<option value="S">Select Party Type</option>
	</select>
	</td>
    </tr>
    <tr>
<td class="FieldCellSub" width="108px">Party Name</td>
<td class="FieldCellSub" colspan="3"> <input type="text" name="txtPartyName" size="40" class="FormElem"></td>
</tr>

<tr>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell">
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
</td>
<td class="FieldCell">
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()" >
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
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585px; height:140px;"-->
<table border="0" cellspacing="1px" class="ExcelTable" width="100%">
<tr>
<td class="ExcelHeaderCell" width="10px" >S.No.
</td>
<td class="ExcelHeaderCell" width="10px" >
</td>
<td class="ExcelHeaderCell">Number
</td>
<td class="ExcelHeaderCell">Date
</td>
<%
    if Trim(sSelVouTy)="MI" then
        Response.Write "<td class='ExcelHeaderCell'>Pay To</td>"
    end if
%>
<td class="ExcelHeaderCell">Party Name
</td>
</tr>

<SCRIPT LANGUAGE=vbscript RUNAT=Server>

</SCRIPT>
<%
Dim sQuery,sFromApp,sVouStatus,sTemp,sPartyName
Dim sParType,iParCode,iSubParType,sParName
iCnt = 0
Const iPageSize=20
'Response.Write "Test="&sPartyType
IF trim(sPartyType) = "" or trim(sPartyType) = "S" then
	IF trim(sSelVouTy) = "SI" then 'Sales Invoice
		sQuery = "select  distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
				 "H.PayToRecdFrom from Acc_T_VoucherHeader H, Acc_T_CreatedReceivables P where H.OUDefinitionID='"&sOrgId&"' "&_
				 "and H.BookCode='05' and P.AmountReceivable > P.AmountReceived Order By 5 "
	ElseIF trim(sSelVouTy) = "SR" then 'Sales Return
	sFromApp = 3
	sVouStatus = "010104"
		sQuery = "Select distinct H.CreatedTransNo,H.CreatedVoucherNo,convert(char,H.VoucherDate,103),H.VoucherAmount,H.PayToRecdFrom,P.ReceivableNumber "&_
				 "from Acc_T_CreatedVoucherHeader H,Sal_T_SalesReturnHeader S,Acc_T_CreatedReceivables P  Where H.BookCode = '05' "&_
				 "and H.CreatedVouchStatus = '"& sVouStatus &"' and H.FromApplication = "&sFromApp&" and S.SaleTransactionNo = H.OtherApplnTransNo "&_
				 "and H.CreatedTransNo = P.CreatedTransNo and P.AmountReceivable > P.AmountReceived Order By 1 "
	ElseIF trim(sSelVouTy) = "PI" then 'Purchase Invoice
		sQuery = "select Distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
	  			 "H.PayToRecdFrom,P.PayablesNumber from Acc_T_VoucherHeader H, Acc_T_Payables P "&_
				 "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='04' and H.TransactionNumber = P.TransactionNumber "&_
				 "and P.AmountPayable > P.AmountPaid Order By 1 "
    Elseif Trim(sSelVouTy) = "MI" then  ' Misc Payment
        sQuery = "Select V.CreatedTransNo,V.VoucherNumber,Convert(varchar,V.VoucherDate,103), "&_
                " V.VoucherAmount,V.PayToRecdFrom,M.MiscTransNo from Acc_T_MiscPymtRequestHeader M,Acc_T_VoucherHeader V"&_
                " where M.ReceiptNo = V.CreatedTransNo and V.OUDefinitionID = '"& sOrgId &"' and M.ApplicationCode = 2 and isNull(M.AdjustmentStatus,'N')='N'"
	End IF
ElseIF trim(sPartyType) <> "" then
'Response.Write sPartyType
	sTemp = split(sPartyType,"?")
	sParType = sTemp(0)
	iSubParType = sTemp(1)
	sParName = sTemp(2)
	iParCode = sTemp(3)

	IF trim(sSelVouTy) = "SI" then 'Sales Invoice
		sQuery = "select  distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
				 "H.PayToRecdFrom from Acc_T_VoucherHeader H, Acc_T_CreatedReceivables P where H.OUDefinitionID='"&sOrgId&"' "&_
				 "and H.BookCode='05' and P.AmountReceivable > P.AmountReceived and H.PartyType = '"& sParType &"' and "&_
				 "H.PartySubType= '"&iSubParType &"' and H.PartyCode = "& iParCode & "  Order By 5 "
	ElseIF trim(sSelVouTy) = "SR" then 'Sales Return
	sFromApp = 3
	sVouStatus = "010104"
		sQuery = "Select distinct H.CreatedTransNo,H.CreatedVoucherNo,convert(char,H.VoucherDate,103),H.VoucherAmount,H.PayToRecdFrom,P.ReceivableNumber "&_
				 "from Acc_T_CreatedVoucherHeader H,Sal_T_SalesReturnHeader S,Acc_T_CreatedReceivables P  Where H.BookCode = '05' "&_
				 "and H.CreatedVouchStatus = '"& sVouStatus &"' and H.FromApplication = "&sFromApp&" and S.SaleTransactionNo = H.OtherApplnTransNo "&_
				 "and H.CreatedTransNo = P.CreatedTransNo and P.AmountReceivable > P.AmountReceived  and H.PartyType = '"& sParType &"' and "&_
				 "H.PartySubType= '"&iSubParType &"' and H.PartyCode = "& iParCode & "	 Order By 1 "
	ElseIF trim(sSelVouTy) = "PI" then 'Purchase Invoice
		sQuery = "select Distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
	  			 "H.PayToRecdFrom,P.PayablesNumber from Acc_T_VoucherHeader H, Acc_T_Payables P "&_
				 "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='04' and H.TransactionNumber = P.TransactionNumber "&_
				 "and P.AmountPayable > P.AmountPaid  and H.PartyType = '"& sParType &"' and "&_
				 "H.PartySubType= '"&iSubParType &"' and H.PartyCode = "& iParCode & " Order By 1 "
    Elseif Trim(sSelVouTy) = "MI" then  ' Misc Payment
        sQuery = "Select V.CreatedTransNo,V.VoucherNumber,Convert(varchar,V.VoucherDate,103), "&_
                " V.VoucherAmount,V.PayToRecdFrom,M.MiscTransNo from Acc_T_MiscPymtRequestHeader M,Acc_T_VoucherHeader V"&_
                " where M.ReceiptNo = V.CreatedTransNo and V.OUDefinitionID = '"& sOrgId &"' and M.ApplicationCode = 2 and isNull(M.AdjustmentStatus,'N')='N'"
	End IF
End IF
'	Response.Write sQuery
	with Objrs
	.ActiveConnection=con
	.CursorLocation=3
	.CursorType=3
	.Source=sQuery
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
	'	Do while not Objrs.EOF
	   iCnt = iCnt + 1
	   sSql = "select PartyName from App_M_PartyMaster where PartyCode in (Select PartyCode from Acc_t_CreatedVoucherHeader where CreatedTransno = "& trim(Objrs(0)) &") "
		Objrs1.Open sSql,con
		If not Objrs1.EOF then
			sPartyName = Objrs1(0)
		End If
		Objrs1.Close
		if Trim(sSelVouTy)="MI" then
		    sQuery = "Select PayToRecdFrom from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& trim(Objrs(0))
		    Objrs1.Open sQuery,con
		    if not Objrs1.EOF then
		        sPayToRecdFrom = trim(Objrs1(0))
		    end if
		    Objrs1.Close

		    sQuery = "Select PartyName from App_M_PartyMaster where PartyCode in("&_
		             "Select PartyCode from Acc_T_MiscPymtRequestHeader where ReceiptNo ="&Trim(Objrs(0))&")"
		    objrs1.Open sQuery,con
		    if not Objrs1.EOF then
		        sPartyName = trim(Objrs1(0))
		    end if
		    objrs1.Close
		end if 'if Trim(sSelVouTy)="MI" then

	%>
<tr>
<td class="ExcelSerial"><%=iCnt%></td>
<td class="ExcelDisplayCell" width="10">
<input type="Checkbox"  name="OptCriteria<%=iCnt%>"  value="<%=trim(Objrs(0))%>:<%=trim(Objrs(1))%>:<%=trim(Objrs(2))%>:<%=trim(Objrs(4))%>" onclick="SetVal(this)">
<td class="ExcelDisplayCell" align="left" ><a href="#" LANGUAGE="VBSCRIPT" onclick="ShowVouch('<%=trim(Objrs(0))%>')" class="ExcelDisplayLink"><%=Objrs(1)%></a></td>
<td class="ExcelDisplayCell" align="left" ><%=Objrs(2)%></td>
<%
    if Trim(sSelVouTy)="MI" then
        Response.Write "<td class='ExcelDisplayCell' align='left'>"& sPayToRecdFrom &"</td>"
    end if
%>
<td class="ExcelDisplayCell" align="left" ><%=sPartyName%></td>
<!--td class="ExcelDisplayCell" align="right" ></td-->
</tr>
<%
		Objrs.MoveNext
	'loop
	if Objrs.EOF then exit for
	next
	End IF

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
<input type=hidden name="hCnt" value=<%=iCnt%>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
		if iCurrentPage = 1 then
%>
			<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
			<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
	<%	else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
	<%	end if%>

<SELECT class="FormElem" onChange="Paginate(this(this.selectedIndex).value)" id=select1 name=select1>
	<%
	For lnPage = 1 To iTotalPage
		If lnPage = iCurrentPage Then
	%>
			<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
	<%	else%>
			<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
	<%	end if
	next
	%>
</SELECT>
	<%if iCurrentPage = iTotalPage then%>
		<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
		<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>
	<%else%>
	<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
	<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotalPage%>')" id=button8 name=button8>
	<%end if
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
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td class="ActionCell">
<input type="button" value="Create GJ Voucher" name="B9" class="ActionButtonX" tabindex="3" onclick="Voucher('GJ')">
<input type="button" value="Create DR Voucher" name="B10" class="ActionButtonX" tabindex="3" onclick="Voucher('DR')">

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

