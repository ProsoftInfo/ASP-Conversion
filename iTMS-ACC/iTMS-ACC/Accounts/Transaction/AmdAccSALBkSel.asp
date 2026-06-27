<%@ Language=VBScript %>
<%	option explicit

	'Program Name				:	AmdAccSALBkSel.asp
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

dim objRs,sQuery,sUnitID
set objRs  = server.CreateObject("adodb.recordset")
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sAccBookRel
dim sCode,sValue,sName,oDOM,iTransno,sRetVal,sFormVal

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF
sFormVal = Request.Form("hFormVal")
'Response.Write sFormVal
Session("AmdSal") = sFormVal

sAccBookRel = "T" 'Book and Accouhead Mapping is Enabled
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo = Request("hTransNo")
'Response.Write iTransno &"=========== "
'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)


'IF CStr(iTransNo) <> "" Then
'	sRetVal = GetVouchXML(iTransNo)
'	oDOM.Load server.MapPath(sRetVal)
'	oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")
'End IF

oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")


sUnitID = session("organizationcode")
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
<XML id="OldVouData" src="<%="../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"%>"></XML>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT language="vbscript">
FUNCTION popPartyType()
	dim iHeadCount

	iUnitNo=document.formname.horgID.value
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

	'if objUnit.selectedIndex <> "0" then
		'iUnitNo= objUnit(objUnit.selectedIndex).value
		iUnitNo = document.formname.horgID.value

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
	'end if
end Function

Function VouCreate
	Dim newElem,objHeader,sTemp,sCurrDate,TempNode,UnitRoot,sAgChk
	Dim sPartyTy,sInvNo,sInvDate,sSendVal,sRetVal,objhttp,sExp
	sCurrDate = document.formname.hCurrDate.value

	Set objhttp = CreateObject("MSXML2.XMLHTTP")
	Set Root = OldVouData.documentElement

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

		sSendVal = sPartTy&"?"&sInvNo&"?"&sInvDate&"?05"&"?"&document.formname.horgID.value
		'MsgBox sSendVal

		document.formname.hInvDate.value=document.formname.ctlDate.GetDate
		document.formname.hOrgName.value=document.formname.horgNameN.value 'document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).Text
		document.formname.hSalType.value=document.formname.selSaleType.options(document.formname.selSaleType.selectedIndex).Text

		sExp = "//Organization"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("OrgId").Value = document.formname.horgID.value
			TempNode.Item(0).Text = document.formname.horgNameN.value 'document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).Text
		End IF

		Set UnitRoot = UnitBookData.documentElement
		For Each HeaderNode In UnitRoot.childNodes
			if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
				document.formname.hBkAccHead.value=HeaderNode.Attributes.Item(2).nodeValue

				sExp = "//Book"
				Set TempNode = Root.selectNodes(sExp)
				IF TempNode.length <> 0 Then
					TempNode.Item(0).Attributes.getNamedItem("BookId").Value = document.formname.selBook.value
					TempNode.Item(0).Attributes.getNamedItem("BKAccHead").Value = HeaderNode.Attributes.Item(2).nodeValue
					TempNode.Item(0).Attributes.getNamedItem("BKOtherUnits").Value = HeaderNode.Attributes.Item(3).nodeValue
					TempNode.Item(0).Text = document.formname.selBook.options(document.formname.selBook.selectedIndex).Text

				end if
			End IF
		next

		document.formname.hInvDate.value = document.formname.ctlDate.GetDate
		sExp = "//SalesType"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("SalType").Value = document.formname.selSaletype.value
			TempNode.Item(0).Text = document.formname.selSaletype.options(document.formname.selSaletype.selectedIndex).Text
		End IF

		sExp = "//SaleInvoice"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("InvNo").Value = document.formname.txtInvoiceNo.value
			TempNode.Item(0).Attributes.getNamedItem("InvDate").Value = document.formname.ctlDate.GetDate
			TempNode.Item(0).Attributes.getNamedItem("RefNo").Value = document.formname.txtRefNo.value
		End IF

		sExp = "//Details"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("VouDate").Value = document.formname.ctlDate.GetDate
		End IF

		if document.formname.optAgentExist(0).checked then
			sAgChk = "Y"
		else
			sAgChk = "N"
		end if

		sExp = "//Party"
		Set TempNode = Root.selectNodes(sExp)
		sTemp=Split(trim(document.formname.hPartyCode.value),"?")
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("ParType").Value = sTemp(0)
			TempNode.Item(0).Attributes.getNamedItem("ParSubType").Value = sTemp(1)
			TempNode.Item(0).Attributes.getNamedItem("ParSubTypeName").Value = document.formname.selPartyType.options(document.formname.selPartyType.selectedIndex).Text
			TempNode.Item(0).Attributes.getNamedItem("ParCode").Value = sTemp(3)
			TempNode.Item(0).Attributes.getNamedItem("Agent").Value = sAgChk
			TempNode.Item(0).Text = document.formname.txtPartyName.value
		End IF

	else
		exit function
	end if
	SaveXML()
End function

function validate()
	'if document.formname.selUnitId.selectedIndex<1 then
	'	MsgBox ("Select Unit")
	'	document.formname.selUnitId.focus
	'	validate= false
	'	exit function
	'end if
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
sOrgId=document.formname.horgID.value
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
'if document.formname.selUnitId.selectedIndex<1 then
'		MsgBox ("Select Unit")
'		document.formname.selUnitId.focus
'		document.formname.optAgentExist(2).checked=true
'		exit function
'	end if

Set Root = OldVouData.documentElement
For Each HeaderNode In Root.childNodes
	if HeaderNode.nodeName="AgentDetails" then
		set temp=Root.removeChild(HeaderNode)
	end if
next

if bFlag<>"N" then
		Set Returnvalue = showModalDialog ("AgentCommisionEntry.asp?OrgID="&document.formname.horgID.value&"&AgentType="&bFlag ,OutData,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		if Returnvalue.hasChildNodes then
			sExp = "//Voucher/Header"
			Set TempNode = Root.selectNodes(sExp)
			Set newNode = Root.insertBefore(Returnvalue,TempNode.Item(0))
			'MsgBox Returnvalue.xml
			'Root.appendChild Returnvalue
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
	objhttp.Open "POST","XMLSave.asp?Mod=SAL&Name=Voucher AMD", false
	objhttp.send OldVouData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		'alert(OldVouData.xml)
		document.formname.submit()
	end if
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
	iUnitNo = document.formname.horgID.value
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

Function DispOldVal()
	Dim sTempNode,sExp,sTemp,iCtr,Root,sTemp2,sTemp3,sAgChk,sAgType,sAgName
	Set Root = OldVouData.documentElement
	sExp = "//Organization"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("OrgId").Value
		'For iCtr = 0 To document.formname.selUnitId.length - 1
		'	IF document.formname.selUnitId.options(iCtr).value = sTemp Then
		'		document.formname.selUnitId.selectedIndex = iCtr
		'		Exit For
		'	End IF
		'Next
		DisplayBook(document.formname.horgID)
	End IF

	sExp = "//Book"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("BookId").Value
		sTemp2 = sTempNode.Item(0).Attributes.getNamedItem("BKAccHead").Value
		document.formname.hBkAccHead.value = sTemp2

		For iCtr = 0 To document.formname.selBook.length - 1
			IF document.formname.selBook.options(iCtr).value = sTemp Then
				document.formname.selBook.selectedIndex = iCtr
				Exit For
			End IF
		Next
		PopulateSalTy
	End IF

	sExp = "//Party"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("ParType").Value
		sTemp2 = sTempNode.Item(0).Attributes.getNamedItem("ParSubType").Value
		sTemp3 = sTempNode.Item(0).Attributes.getNamedItem("ParCode").Value
		sAgChk = sTempNode.Item(0).Attributes.getNamedItem("Agent").Value
		document.formname.txtPartyName.value = sTempNode.Item(0).Text


		sTemp = sTemp&"?"&sTemp2
		For iCtr = 0 To document.formname.selPartyType.length - 1
			IF document.formname.selPartyType.options(iCtr).value = sTemp Then
				document.formname.selPartyType.selectedIndex = iCtr
				Exit For
			End IF
		Next
		sTemp = sTemp&"?"&sTempNode.Item(0).Text
		sTemp = sTemp&"?"&sTemp3

		document.formname.hPartyCode.value = sTemp
	End IF

	IF CStr(sAgChk) = "Y" Then
		sExp = "//AgentDetails/Agent"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sAgType = sTempNode.Item(0).Attributes.getNamedItem("PartyType").Value
			sAgName = sTempNode.Item(0).Attributes.getNamedItem("Agentname").Value
		End IF
	Else
		sAgType = ""
		sAgName = ""
	End IF

	document.all.spAgentName.innerHTML = sAgName
	document.formname.hCommName.Value = sAgName

	IF CStr(sAgType) = "CR" Then
		document.formname.optAgentExist(0).checked = True
	Elseif CStr(sAgChk) = "DR" Then
		document.formname.optAgentExist(1).checked = True
	Else
		document.formname.optAgentExist(2).checked = True
	End IF

	sExp = "//SaleInvoice"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		document.formname.txtInvoiceNo.value = sTempNode.Item(0).Attributes.getNamedItem("InvNo").Value
		document.formname.ctlDate.setDate = sTempNode.Item(0).Attributes.getNamedItem("InvDate").Value
		document.formname.txtRefNo.value = sTempNode.Item(0).Attributes.getNamedItem("RefNo").Value

	End IF

	sExp = "//SalesType"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		For iCtr = 0 To document.formname.selSaleType.length - 1
			IF CStr(document.formname.selSaleType.options(iCtr).Value) = CStr(sTempNode.Item(0).Attributes.getNamedItem("SalType").Value) Then
				document.formname.selSaleType.selectedIndex = iCtr
				Exit For
			End IF
		Next
	End IF


	'sExp = "//PurCategory"
	'Set sTempNode = Root.selectNodes(sExp)
	'IF sTempNode.length <> 0 Then
	'	sTemp = sTempNode.Item(0).Attributes.getNamedItem("Code").Value
	'	For iCtr = 0 To document.formname.selPurCat.length - 1
	'		IF document.formname.selPurCat.options(iCtr).value = sTemp Then
	'			document.formname.selPurCat.selectedIndex = iCtr
	'			Exit For
	'		End IF
	'	Next
	'End IF


End Function


</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="DisplayBook('<%=sUnitID%>');DispOldVal()">

<form method="POST" name="formname" action="VouSALAccAmdEntry.asp?AmdType=A">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
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
<input type="hidden" name="horgID" value="<%=sUnitID%>">
<input type="hidden" name="horgNameN" value="<%=Session("orgshortname")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Accounted Sales Voucher Amendment
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
                                <!--<tr>
                            <td class="FieldCell" width="108">Organization </td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationListDBWithVal("")%>
                              </select></td>
                                </tr>-->
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
                            <td class="FieldCell" colspan="3"> <input type="text" name="txtPartyName" size="61" class="FormElem"></td>
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
                                                                <input type="button" value="Amend" name="btnCreate" class="ActionButton" onClick="VouCreate()" >
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
