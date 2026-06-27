<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouAccAmdPURBkSel.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	March 01, 2003
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
dim objRs,sQuery,sFinPeriod,sFinFun,iTransNo,sRetVal,oDOM,sFormVal

sFinPeriod = Session("FinPeriod")
sFinFun = GetfinYear(sFinPeriod)
If trim(sFinFun) = "True" Then
	Response.Redirect ("../../welcome_Welcome.asp?sFinFun="&sFinFun&"")
End If



iTransNo = Request("TransNo")
sFormVal = Request.Form("hFormVal")
'Response.Write sFormVal
Session("AmdPur") = sFormVal

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")



IF CStr(iTransNo) <> "" Then
	sRetVal = GetVouchXML(iTransNo)
	oDOM.Load server.MapPath(sRetVal)
	oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")
End IF

oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="VoucherData"><Voucher/></XML>
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<XML id="OldVouData" src="<%="../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"%>"></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
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
	Dim iUnitNo,arrTemp,Root
	document.formname.selBook.options.length = 1

	if objUnit.selectedIndex <> "0" then
		iUnitNo= objUnit(objUnit.selectedIndex).value
		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=04&orgID=" & iUnitNo , false
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
	dim sTemp,sTempNode,sTemp2,sTemp3

	Set objhttp = CreateObject("MSXML2.XMLHTTP")

	Set Root = OldVouData.documentElement
	if validate then
		sExp = "//Organization"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sTempNode.Item(0).Attributes.getNamedItem("OrgId").Value = document.formname.selUnitId.value
			sTempNode.Item(0).Attributes.getNamedItem("AccUnit").Value = document.formname.selUnitId.value
			sTempNode.Item(0).Text = document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
		End IF


		sExp = "//Book"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sTempNode.Item(0).Attributes.getNamedItem("BookId").Value = document.formname.selBook.value
			sTempNode.Item(0).Text = document.formname.selBook.options(document.formname.selBook.selectedIndex).text
		End IF

		sExp = "//Party"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sTemp2 = Split(document.formname.hPartyCode.value,"?")

			sTempNode.Item(0).Attributes.getNamedItem("ParType").Value = sTemp2(0)
			sTempNode.Item(0).Attributes.getNamedItem("ParSubType").Value = sTemp2(1)
			sTempNode.Item(0).Attributes.getNamedItem("ParCode").Value = sTemp2(3)
			sTempNode.Item(0).Text = document.formname.txtPartyName.value
		End IF

		sExp = "//PurInvoice"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sTempNode.Item(0).Attributes.getNamedItem("PurInvNo").Value = document.formname.txtInvoiceNo.value
			sTempNode.Item(0).Attributes.getNamedItem("PurInvDate").Value = document.formname.ctlDate.GetDate()
		End IF

		sExp = "//PurCategory"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sTempNode.Item(0).Attributes.getNamedItem("Code").Value = document.formname.selPurCat.value
			sTempNode.Item(0).Text = document.formname.selPurCat.value
		End IF

		sExp = "//PurchaseType"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sTempNode.Item(0).Attributes.getNamedItem("PurTypeId").Value = document.formname.selPurType.value
			sTempNode.Item(0).Text = document.formname.selPurType.Options(document.formname.selPurType.selectedIndex).Text
		End IF


	else
		exit function
	end if
	SaveXML()
End function

function validate()
	Dim sCurrDate
	sCurrDate = document.formname.hCurrDate.Value



	if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")
		document.formname.selUnitId.focus
		validate= false
		exit function
	end if
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select Purchase Book")
		document.formname.selBook.focus
		validate= false
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
	if trim(document.formname.txtInvoiceNo.value)="" then
		MsgBox ("Invoice No should not be blank")
		document.formname.txtInvoiceNo.select
		validate=false
		exit function
	end if

	IF DateDiff("d",document.formname.ctlDate.GetDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		validate=false
		Exit Function
	End IF




	validate = True
End function

function selAccountHead(objAcc)
dim sOrgId,sPartyType
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp

sRetVal2 = "0:0:0"

sOrgId=document.formname.selUnitId.value
sPartyType=objAcc.value &"?"& Replace(objAcc.options(objAcc.selectedIndex).text,"&"," and ")
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
		document.formname.txtPartyName.value = ""
		document.formname.hPartyCode.value = "0"
		Exit Function
	End IF

	sRetValue = OutValue
	sTemp = Split(sRetValue,":")
	sParTy = sTemp(4)
	sParSubType = sTemp(3)
	sParCode = sTemp(1)
	sPartyName = sTemp(0)
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
Else
	document.formname.txtPartyName.value = ""
	document.formname.hPartyCode.value = "0"
End if 'End of If any Account Head Selected Check
End function
'---------------------End Of Function selAccountHead----------------------
Function SaveXML()
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Mod=PUR&Name=Voucher AMD", false

	objhttp.send OldVouData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		document.formname.hInvDate.value = document.formname.ctlDate.GetDate()
		DisButt()
		document.formname.submit()

	end if
End Function

Function DisButt()
	'document.formname.btnAmend.disabled = True
	document.formname.btnCreate.disabled = True
	'document.formname.btnDel.disabled = True
	'document.formname.btnView.disabled = True
End Function

Function DispOldVal()
	Dim sTempNode,sExp,sTemp,iCtr,Root,sTemp2,sTemp3
	Set Root = OldVouData.documentElement

	sExp = "//Organization"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("OrgId").Value
		For iCtr = 0 To document.formname.selUnitId.length - 1
			IF document.formname.selUnitId.options(iCtr).value = sTemp Then
				document.formname.selUnitId.selectedIndex = iCtr
				Exit For
			End IF
		Next
		DisplayBook(document.formname.selUnitId)
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
	End IF

	sExp = "//Party"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("ParType").Value
		sTemp2 = sTempNode.Item(0).Attributes.getNamedItem("ParSubType").Value
		sTemp3 = sTempNode.Item(0).Attributes.getNamedItem("ParCode").Value
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

	sExp = "//PurInvoice"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		document.formname.txtInvoiceNo.value = sTempNode.Item(0).Attributes.getNamedItem("PurInvNo").Value
		document.formname.ctlDate.setDate = sTempNode.Item(0).Attributes.getNamedItem("PurInvDate").Value
	End IF

	sExp = "//PurCategory"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("Code").Value
		For iCtr = 0 To document.formname.selPurCat.length - 1
			IF document.formname.selPurCat.options(iCtr).value = sTemp Then
				document.formname.selPurCat.selectedIndex = iCtr
				Exit For
			End IF
		Next
	End IF

	sExp = "//PurchaseType"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sTemp = sTempNode.Item(0).Attributes.getNamedItem("PurTypeId").Value
			For iCtr = 0 To document.formname.selPurType.length - 1
				IF document.formname.selPurType.options(iCtr).value = sTemp Then
					document.formname.selPurType.selectedIndex = iCtr
					Exit For
				End IF
			Next
	End IF


End Function

</script>
</HEAD>
	<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="DispOldVal()">

<form method="POST" name="formname" action="VouPURAccAmdDetails.asp?AmdType=A">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hBkAccHead" value="">
<input type="hidden" name="hInvDate" value="">
<input type="hidden" name="hVouDate" value="">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Accounted Purchase Voucher Amendment </td>
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
									<%populateOrganizationList%>
                              </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Purchase Book</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selBook" class="FormElem">
                        <option value="S">Select Book</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Purchase Type&nbsp;</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selPurType" class="FormElem" >
									<option value="0">Select Purchase Type</option>
									<%
										dim sCode,sValue
										sQuery = "Select PurchaseType,PurchaseTypeName from APP_M_PurchaseTypes Where Active = 'Y' "
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
									  	Do while not objRs.EOF
									%>
											<option value="<%Response.Write sCode%>"><%Response.Write sValue%></option>
									<%

											objRs.MoveNext
										Loop
										objRs.Close
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
                            <td class="FieldCell" width="108">Invoice Number</td>
                            <td class="FieldCell">
                            <input type="text" name="txtInvoiceNo" size="20" class="FormElem">
                            <!--
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                           <a href="javascript:popVoucherNo('C')">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:popVoucherNo('A')">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers "></a>-->

                            </td>
                            <td class="FieldCell" width="120"> Invoice Date</td>
                            <td class="FieldCell">
								     <% ' Function Call to Insert Date Picker
						  						Response.Write InsertDatePicker("ctlDate")
						  			 %>
							</td>
                                </tr>

                            <tr>
                            <td class="FieldCell" width="108">Category</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selPurCat" class="FormElem" >
									<option value="0">Select Purchase Category</option>
									<%
										'dim sCode,sValue
										sQuery = "Select CategoryCode,CategoryName From APP_M_InvoiceCategory Where ApplicableFor = 'P' Order By CategoryName "
									  	With objRs
									  		.CursorLocation = 3
									  		.CursorType = 3
									  		.Source = sQuery
									  		.ActiveConnection = con
									  		.Open
									  	End with
									  	Set objRs.Activeconnection = nothing
									  	Do while not objRs.EOF
									%>
											<option value="<%Response.Write objRs(0)%>"><%Response.Write objRs(1)%></option>
									<%

											objRs.MoveNext
										Loop
										objRs.Close
								    %>
								</select></td>
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
                                                                <input type="button" value="Amendment" name="btnAmend" class="ActionButtonX" onClick="VouAmend()">
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