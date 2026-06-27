<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouSALBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  13, 2003
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
dim objRs,sQuery
set objRs  = server.CreateObject("adodb.recordset")
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sAccBookRel
dim sCode,sValue,sName

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

sAccBookRel = "T" 'Book and Accouhead Mapping is Enabled
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="SaleTypeData"><Book/></XML>
<XML ID="VoucherData"><Voucher/></XML>
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
	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo&"&sCallTy=P" , false
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

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=05&orgID=" & iUnitNo , false
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

Function VouCreate
dim newElem,objHeader,sTemp,sCurrDate
Dim sPartyTy,sInvNo,sInvDate,sSendVal,sRetVal,objhttp
sCurrDate = document.formname.hCurrDate.value

Set objhttp = CreateObject("MSXML2.XMLHTTP")
Set Root = VoucherData.documentElement

	if validate then

		IF document.formname.txtInvoiceNo.value = "" Then
			MsgBox "Enter Invoice Number "
			document.formname.txtInvoiceNo.focus()
			Exit Function
		End IF

		IF DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
			MsgBox "Voucher Date Should be Less than the System Date "
			Exit Function
		End IF

		sPartTy = document.formname.hPartyCode.value
		sInvNo = document.formname.txtInvoiceNo.value
		sInvDate = document.formname.ctlDate.GetDate()

		sSendVal = sPartTy&"?"&sInvNo&"?"&sInvDate&"?05"&"?"&document.formname.selUnitId.value
		'MsgBox sSendVal

		objhttp.Open "GET","CheckInvCreate.asp?sValue="&sSendVal , false
		objhttp.send
		sRetVal = objHttp.responseText
		'MsgBox sRetVal
		IF CStr(sRetVal) <> "C" Then
			MsgBox "Sales Voucher already Created for this Party,InvoiceNo and Invoice Date "
			Exit Function
		End IF


		set objHeader= VoucherData.createElement("Header")
		Root.appendChild objHeader

		document.formname.hInvDate.value=document.formname.ctlDate.GetDate
		document.formname.hOrgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).Text
		document.formname.hSalType.value=document.formname.selSaleType.options(document.formname.selSaleType.selectedIndex).Text

		Set newElem = VoucherData.createElement("Organization")
		newElem.setAttribute "OrgId",document.formname.selUnitId.value
		newElem.Text=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).Text
		objHeader.appendChild newElem

		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
				document.formname.hBkAccHead.value=HeaderNode.Attributes.Item(2).nodeValue

				Set newElem = VoucherData.createElement("Book")
				newElem.setAttribute "BookId",document.formname.selBook.value
				newElem.setAttribute "BKAccHead",HeaderNode.Attributes.Item(2).nodeValue
				newElem.setAttribute "BKOtherUnits",HeaderNode.Attributes.Item(3).nodeValue
				newElem.Text=document.formname.selBook.options(document.formname.selBook.selectedIndex).Text
				objHeader.appendChild newElem
			end if
		next


		Set newElem = VoucherData.createElement("SalesType")
		newElem.setAttribute "SalType",document.formname.selSaletype.value
		newElem.Text=document.formname.selSaletype.options(document.formname.selSaletype.selectedIndex).Text
		objHeader.appendChild newElem

		Set newElem = VoucherData.createElement("SaleInvoice")
		newElem.setAttribute "InvNo",document.formname.txtInvoiceNo.value
		newElem.setAttribute "InvDate",document.formname.ctlDate.GetDate
		newElem.setAttribute "RefNo",document.formname.txtRefNo.value

		objHeader.appendChild newElem

		Set newElem = VoucherData.createElement("Party")

		sTemp=Split(trim(document.formname.hPartyCode.value),"?")

		newElem.setAttribute "ParType",sTemp(0)
		newElem.setAttribute "ParSubType",sTemp(1)
		newElem.setAttribute "ParSubTypeName",document.formname.selPartyType.options(document.formname.selPartyType.selectedIndex).Text
		newElem.setAttribute "ParCode",sTemp(3)

		if document.formname.optAgentExist(0).checked then
			newElem.setAttribute "Agent","Y"
		else
			newElem.setAttribute "Agent","N"
		end if
		newElem.Text=document.formname.txtPartyName.value
		objHeader.appendChild newElem
	else
		exit function
	end if
	SaveXML()
End function

function validate()
	if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")
		document.formname.selUnitId.focus
		validate= false
		exit function
	end if
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select SalesBook")
		document.formname.selBook.focus
		validate= false
		exit function
	end if
	if document.formname.selSaletype.selectedIndex<1 then
		MsgBox ("Select Sales type")
		document.formname.selSaletype.focus
		validate=false
		exit function
	end if

	if document.formname.selPartyType.selectedIndex<1 then
		MsgBox ("Select Party")
		document.formname.selPartyType.focus
		validate=false
		exit function
	end if
	if trim(document.formname.txtPartyName.value)="" then
		MsgBox ("Party Name should not be blank")
		document.formname.txtPartyName.select
		validate=false
		exit function
	end if
	validate=true
End function

function selAccountHead(objAcc)
dim sOrgId,sPartyType
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp

sRetVal2 = "0:0"
sOrgId=document.formname.selUnitId.value
sPartyType=objAcc.value &"?"& objAcc.options(objAcc.selectedIndex).text

if objAcc.selectedIndex >0 then
	'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","")
	OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")
	while UBound(arrTemp) = 0
		OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend

	IF UBound(arrtemp) <= 1 Then
		objAcc.selectedIndex = 0
		Exit Function
	End IF

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
		'User Has Selected a GL Account Head
		For Each HeaderNode In nodAccHead.childNodes
			document.formname.hPartyCode.value=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
			document.formname.txtPartyName.value=HeaderNode.Attributes.Item(3).nodeValue

		next
	else
		objAcc.selectedIndex=0
	end if 'End of Party Head Processing
End if 'End of If any Account Head Selected Check
End function
'---------------------End Of Function selAccountHead----------------------
Function showAgent(bFlag)
dim Returnvalue,sExp,TempNode
if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")
		document.formname.selUnitId.focus
		document.formname.optAgentExist(2).checked=true
		exit function
	end if

Set Root = VoucherData.documentElement
For Each HeaderNode In Root.childNodes
	if HeaderNode.nodeName="AgentDetails" then
		set temp=Root.removeChild(HeaderNode)
	end if
next

if bFlag<>"N" then
		Set Returnvalue = showModalDialog ("AgentCommisionEntry.asp?OrgID="&document.formname.selUnitId.value&"&AgentType="&bFlag ,OutData,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		if Returnvalue.hasChildNodes then
			Root.appendChild Returnvalue
			sExp = "//Agent"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.Length <> 0 Then
				document.all.spAgentName.innerHTML = TempNode.Item(0).Attributes.Item(1).nodeValue
				document.formname.hCommName.Value = TempNode.Item(0).Attributes.Item(1).nodeValue
			Else
				document.all.spAgentName.innerHTML = ""
				document.formname.hCommName.Value = ""
			End IF
		else
			document.formname.optAgentExist(2).checked=true
		end if
end if
End function

Function SaveXML()
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Mod=SAL&Name=Voucher Entry", false
	objhttp.send VoucherData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		document.formname.submit()
	end if
End Function

Function popVoucherNo(sVal)
Dim iPurTy,iParTy,iParCode,sTemparr,sInvDate
	IF (validateForDel()) Then
		sCallTy = sVal
		sOrgId = document.formname.selUnitId.value
		iBookId = "05"
		iBookNo = document.formname.selBook.value
		iPurTy = document.formname.selSaleType.value
		iParTy = document.formname.selPartyType.value
		iParCode = document.formname.hPartyCode.value
		sTrans = "SJR"
		IF CStr(sVal) <> "A" Then
			sTemp = showModalDialog("VouchSelForSalPur.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans&"&sPurTy="&iPurTy&"&sParTy="&iParTy&"&iParCode="&iParCode,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
		Else
			sTemp = showModalDialog("VouchSelForSalPur.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans&"&sPurTy="&iPurTy&"&sParTy="&iParTy&"&iParCode="&iParCode,"","dialogHeight:400px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No")
		End IF

		IF CStr(sTemp) = "0" Then
			Exit Function
		End IF

		sTemparr = Split(sTemp,"~")
		document.formname.hTransNo.value = sTemparr(0)
		document.formname.txtInvoiceNo.value = sTemparr(1)
		document.formname.txtRefNo.value = sTemparr(2)
		'document.formname.hInvDate.value = sTemparr(3)

		IF CStr(sTemparr(3)) <> "" Then
			document.formname.ctlDate.setDate() = sTemparr(3)
		End IF
	End IF
End Function

Function SetDate()
	Dim sFromYr,sToYr
	sFromYr = document.formname.hFromYr.Value
	sToYr = document.formname.hToYr.Value
	sFromYr = "01/04/"&Trim(sFromYr)
	sToYr = "31/03/"&sToYr
	document.formname.ctlDate.setMinDate() = sFromYr
	document.formname.ctlDate.setMaxDate() = sToYr
End Function

Function PopulateSalTy
	Dim iUnitNo,iBookNo
	iUnitNo = document.formname.selUnitId.value
	iBookNo = document.formname.selBook.value
	Set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLGetBookSalPurType.asp?BkCode=05&orgID="& iUnitNo&"&BookNo="&iBookNo , false
	objhttp.send

	'Msgbox objhttp.responseText
	if objhttp.responseXML.xml <> "" then
		SaleTypeData.loadXML objhttp.responseXML.xml
		Set Root = SaleTypeData.documentElement
		document.formname.selSaleType.length = 1
		For Each HeaderNode In Root.childNodes
			document.formname.selSaleType.length = document.formname.selSaleType.length+1
			document.formname.selSaleType.options(document.formname.selSaleType.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
			document.formname.selSaleType.options(document.formname.selSaleType.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
		next
	end if
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
	var width;
	var url;
	if (typeof window.validateForDel === "function" && !window.validateForDel()) {
		return;
	}
	sCallTy = sVal;
	sOrgId = form.selUnitId.value;
	iBookId = "05";
	iBookNo = form.selBook.value;
	iPurTy = form.selSaleType.value;
	iParTy = form.selPartyType.value;
	iParCode = form.hPartyCode.value;
	sTrans = "SJR";
	width = String(sVal) !== "A" ? "450" : "500";
	url = "VouchSelForSalPur.asp?flag=" + encodeURIComponent(sCallTy) +
		"&orgId=" + encodeURIComponent(sOrgId) +
		"&BookCode=" + encodeURIComponent(iBookId) +
		"&BookNo=" + encodeURIComponent(iBookNo) +
		"&TransType=" + encodeURIComponent(sTrans) +
		"&sPurTy=" + encodeURIComponent(iPurTy) +
		"&sParTy=" + encodeURIComponent(iParTy) +
		"&iParCode=" + encodeURIComponent(iParCode);
	openItmsVoucherDialog(url, "dialogHeight:400px;dialogWidth:" + width + "px;center:Yes;help:No;resizable:No;status:No", function (sTemp) {
		var parts;
		var ctlDate;
		if (String(sTemp || "0") === "0") {
			return;
		}
		parts = String(sTemp).split("~");
		form.hTransNo.value = parts[0] || "";
		form.txtInvoiceNo.value = parts[1] || "";
		form.txtRefNo.value = parts[2] || "";
		ctlDate = form.ctlDate || document.getElementById("ctlDate");
		if (parts[3] && ctlDate) {
			if (typeof ctlDate.SetDate === "function") {
				ctlDate.SetDate(parts[3]);
			} else if (typeof ctlDate.setDate === "function") {
				ctlDate.setDate(parts[3]);
			} else {
				ctlDate.value = parts[3];
			}
		}
	});
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">

<form method="POST" name="formname" action="VouSALDetailsEntry.asp">
<input type="hidden" name="hTransNo" value="">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hBkAccHead" value="">
<input type="hidden" name="hInvDate" value="">
<input type="hidden" name="hOrgName" value="">
<input type="hidden" name="hSalType" value="">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hCommName" value="">
<input type="hidden" name="hAccBookRel" value="<%=sAccBookRel%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher Entry
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
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Invoice Details</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="100">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Commission</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Advance</td>
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
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="108">Organization </td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationListDBWithVal("")%>
                              </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Sales Book</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selBook" class="FormElem" onChange="PopulateSalTy()">
                        <option value="S">Select Book</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Sale Type&nbsp;</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selSaleType" class="FormElem" >
									<option value="0">Select Sale Type</option>
									<%

										IF CStr(sAccBookRel) <> "T" Then 'Book and Account Head is Not Done
											sQuery = "Select InvoiceType,InvTypeShortName,InvoiceTypeName from Sal_M_InvoiceTypes where TobeAccounted=1 and Useable = 1 Order By InvoiceTypeName "
									  		With objRs
									  			.CursorLocation = 3
									  			.CursorType = 3
									  			.Source = sQuery
									  			.ActiveConnection = con
									  			.Open
									  		End with
									  		Set objRs.Activeconnection = nothing
									  		Set sCode = objRs(0)
									  		Set sValue = objRs(1)
									  		Set sName = objRs(2)
									  		Do while not objRs.EOF

									%>
												<option value="<%Response.Write sCode%>"><%=trim(sName)%></option>
									<%
												objRs.MoveNext
											Loop
											objRs.Close
										End IF
								    %>
								</select>
                            </td>
                                </tr>

                                <tr>
                            <td class="FieldCell" width="108">Party Type</td>
                            <td class="FieldCell" colspan="3">
                            	<select size="1" name="selPartyType" class="FormElem" onChange="selAccountHead(this)">
								<option value="A">Select Party Type</option>
								</select>
                              </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Party Name</td>
                            <td class="FieldCell" colspan="3"> <input type="text" name="txtPartyName" size="40" class="FormElem"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Agent </td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td class="FieldCell">
                            <input type="radio" value="C" name="optAgentExist" onClick="showAgent('1')" class="Formelem">

                                </td>
                                <td class="FieldCell">Commission Agent</td>
                                <td class="FieldCell">
                            <input type="radio" value="D" name="optAgentExist" onClick="showAgent('2')" class="Formelem">

                                </td>
                                <td class="FieldCell">Depo Agent</td>
                                <td class="FieldCell"> <input type="radio" value="No" onClick="showAgent('N')" checked name="optAgentExist" class="Formelem">
                                </td>
                                <td class="FieldCell">
                              No Agent</td>
                              </tr>
                            </table>
                            </td>
                                </tr>
                                <tr>
										<td class="FieldCell" width="108">&nbsp;</td>
										<td class="FieldCell"><span ID="spAgentName" class="ExcelDisplayCell"></span></td>
										<td class="FieldCell" width="120"></td>
										<td class="FieldCell"></td>
                                </tr>
                                 <tr>
                            <td class="FieldCell" width="108">Reference Number</td>
                            <td class="FieldCell"><input type="text" name="txtRefNo" size="20" maxlength="30" class="FormElem"></td>
                            <td class="FieldCell" width="120"></td>
                            <td class="FieldCell"></td>
                                </tr>

                                <tr>
                            <td class="FieldCell" width="108">Invoice Number</td>
                            <td class="FieldCell">
                            <input type="text" name="txtInvoiceNo" size="20" class="FormElem">
                            <!--
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                           <a href="javascript:popVoucherNo('C')">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:popVoucherNo('A')">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt=" Accounted Vouchers "></a> -->
                            </td>
                            <td class="FieldCell" width="120"> Invoice Date</td>
                            <td class="FieldCell">
                                  <% ' Function Call to Insert Date Picker
										Response.Write InsertDatePicker("ctlDate")
									%>
                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="10" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Create" name="btnCreate" class="ActionButton" onClick="VouCreate()" >
                                                                <!--input type="button" value="View" name="btnView" class="ActionButton" onClick="VouView()" >
                                                                <input type="button" value="Amendment" name="btnAmend" class="ActionButtonX" onClick="VouAmend()" >
                                                                <input type="button" value="Delete" name="btnDel" class="ActionButton" onClick="VouDel()" -->
                                                                <input type="reset" value="Reset" name="B5" class="ActionButton"  >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="10" class="BottomPack" colspan="3">
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
