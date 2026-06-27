<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 27,2003
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
<%
	Dim sFinPeriod,sFromYr,sToYr,sTempYr
	sFinPeriod = Session("FinPeriod")
	IF CStr(sFinPeriod) <> "" Then
		sTempYr = Split(sFinPeriod,":")
		sFromYr = sTempYr(0)
		sToYr = sTempYr(1)
	End IF

%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="OutData"><PartyType/></XML>
<XML id="AccHeadData">
<account/>
</XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT language="vbscript">
dim sFlag

Function popPartType()
set objhttp = CreateObject("MSXML2.XMLHTTP")

if 	document.formname.selUnitId.value<>"0" then
	iUnitNo=document.formname.selUnitId.value
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
else
		document.formname.selAccHead.length=2
end if

end function
Function DisplayBook()
dim iUnitNo,arrTemp,BkCode
dim Root

	document.formname.selBook.options.length = 1
	if document.formname.selUnitId.selectedIndex <> "0" and document.formname.selVoucher.selectedIndex<>"0" then
		iUnitNo= document.formname.selUnitId.value
		BkCode= document.formname.selVoucher.value

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode="&BkCode&"&orgID=" & iUnitNo , false
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
	end if
end Function
function validate()
	if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")
		document.formname.selUnitId.focus
		exit function
	end if
	if document.formname.selVoucher.selectedIndex<1 then
		MsgBox ("Select Voucher type")
		document.formname.selVoucher.focus
		exit function
	end if
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select a Book")
		document.formname.selBook.focus
		exit function
	end if
	If sFlag="VouNo" Then
		If  document.formname.txtNoFrom.value="" Then
			Msgbox "Enter Voucher No. From "
			document.formname.txtNoFrom.select
			Exit Function
		ElseIf document.formname.txtNoTo.value ="" Then
			Msgbox "Enter Voucher No. To "
			document.formname.txtNoTo.select
			Exit Function
		End if
'------------- Coding For the case VouDate is Selected ----------------
	Elseif sFlag="VouDate" Then
			sFromDate=document.formname.CtlVouFromDate.GetDate
			sToDate=document.formname.CtlVouToDate.GetDate
			If dateDiff("d",sFromDate,sToDate)<0 Then
				Msgbox "To Date Should be Greater than From Date"
				Exit Function
			end if
'------------- Coding For the case Amount is Selected ------------------
	Elseif sFlag="Amount" Then
		If	document.formname.txtGAmount.value="" Then
				Msgbox "Enter From Amount"
				document.formname.txtGAmount.select
				Exit Function
		Elseif document.formname.txtLAmount.value="" Then
				Msgbox "Enter To Amount"
				document.formname.txtLAmount.select
				Exit Function
		Elseif not(IsNumeric(document.formname.txtGAmount.value)) Then
				Msgbox "Enter Numbers Only"
				document.formname.txtGAmount.select
				Exit Function
		Elseif not(IsNumeric(document.formname.txtLAmount.value)) Then
				Msgbox "Enter Numbers Only"
				document.formname.txtLAmount.select
				Exit Function
		Else
			dGAmount=cdbl(document.formname.txtGAmount.value)
			dLAmount=cdbl(document.formname.txtLAmount.value)
			If cdbl(dGAmount)>cdbl(dLAmount) Then
				Msgbox "To Amount Should be Greater Than From Amount "
				document.formname.txtLAmount.value =""
				document.formname.txtLAmount.select
				Exit Function
			End if
		end if
'----------- Coding For the case Account Head is Selected -------------
	Elseif sFlag ="AccHead" Then
			if document.formname.SelAccHead.value="0" then
				Msgbox "Select Account Head"
				document.formname.SelAccHead.focus
				Exit Function
			End if
	end if

	document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
	document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text
	document.formname.hVoucherName.value=document.formname.selVoucher.options(document.formname.selVoucher.selectedIndex).text
	document.formname.hFromDate.value=document.formname.CtlVouFromDate.GetDate
	document.formname.hToDate.value=document.formname.CtlVouToDate.GetDate

	document.formname.submit()
End function

Function OptSelection()

if document.formname.optCriteria(0).checked then
	sFlag=document.formname.optCriteria(0).value
	document.formname.txtNoFrom.readOnly =false
	document.formname.txtNoTo.readOnly =false
	document.formname.txtGAmount.value =""
	document.formname.txtLAmount.value =""
	document.formname.txtGAmount.readOnly=True
	document.formname.txtLAmount.readOnly=True
	document.formname.SelAccHead.disabled=true
	window.spAccHead.innerHTML =""
Elseif document.formname.optCriteria(1).checked then
	sFlag=document.formname.optCriteria(1).value
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	document.formname.txtGAmount.value =""
	document.formname.txtLAmount.value =""
	document.formname.txtNoFrom.readOnly =true
	document.formname.txtNoTo.readOnly =true
	document.formname.txtGAmount.readOnly=true
	document.formname.txtLAmount.readOnly =true
	document.formname.SelAccHead.disabled=true
	window.spAccHead.innerHTML =""

Elseif document.formname.optCriteria(2).checked then
	sFlag=document.formname.optCriteria(2).value
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	document.formname.txtNoFrom.readOnly =true
	document.formname.txtNoTo.readOnly =true
	document.formname.txtGAmount.readOnly=false
	document.formname.txtLAmount.readOnly =false
	document.formname.SelAccHead.disabled=true
	window.spAccHead.innerHTML =""

Elseif document.formname.optCriteria(3).checked then
	sFlag=document.formname.optCriteria(3).value
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	document.formname.txtGAmount.value =""
	document.formname.txtLAmount.value =""
	document.formname.txtNoFrom.readOnly =true
	document.formname.txtNoTo.readOnly =true
	document.formname.txtGAmount.readOnly=True
	document.formname.txtLAmount.readOnly =True
	document.formname.SelAccHead.disabled =false
End if

End Function

Function SelNew()
	document.formname.SelAccHead.selectedIndex=0
	document.formname.txtGAmount.value  =""
	document.formname.txtLAmount.value =""
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	window.spAccHead.innerHTML =""
End Function
'---------- Selection of Account Head From Pop up Screen --------------

Function SelectAccHead()
dim iGlHead,sOrgId,sAccHead,arrTemp,sRetVal
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
set objhttp = CreateObject("Microsoft.XMLHTTP")

If document.formname.selUnitId.selectedIndex="0" then
	Msgbox "Select Organaisation Id"
	document.formname.SelAccHead.selectedIndex=0
	document.formname.selUnitId.focus
Elseif document.formname.selVoucher.value="0" Then
		Msgbox "Select Voucher Type"
		document.formname.selVoucher.focus
		document.formname.SelAccHead.selectedIndex=0
		Exit Function
Elseif document.formname.selBook.value="S" Then
		Msgbox "Select Book"
		document.formname.selBook.focus
		document.formname.SelAccHead.selectedIndex=0
		Exit Function
Else
	sOrgId=document.formname.selUnitId.value
	sBookid=document.formname.selVoucher.value
	sBookNo=document.formname.selBook.value
	if 	document.formname.SelAccHead.value="G" then
		'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgid="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
		OutValue = showModalDialog("GLHeadSelection.asp?orgid="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")

		while UBound(arrTemp) = 0
			OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
			arrTemp = split(OutValue,":")
		wend

		sRetVal = OutValue

		if UBound(arrTemp) <= 1 then exit function
		GetGlHeadXml(sRetVal)

		Set nodAccHead = AccHeadData.documentElement


		if nodAccHead.hasChildNodes then
			For Each HeaderNode In nodAccHead.childNodes
				document.formname.hAccHead.value=HeaderNode.Attributes.getNamedItem("No").Value
				window.spAccHead.innerHTML=HeaderNode.Attributes.getNamedItem("Name").Value&"&nbsp;"
			next
		else
			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
			window.spAccHead.innerHTML=""
		End if
	else
		sPartyType=document.formname.SelAccHead.value& "?" & document.formname.SelAccHead.options(document.formname.SelAccHead.selectedIndex).text
		'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
		OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")

		while UBound(arrTemp) = 0
			OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
			arrTemp = split(OutValue,":")
		wend

		sRetValue = OutValue
		sTemp = Split(sRetValue,":")
		sParTy = sTemp(4)
		sParSubType = sTemp(3)
		sParCode = sTemp(1)
		sPartyName = sTemp(0)

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
				window.spAccHead.innerHTML=HeaderNode.Attributes.getNamedItem("Name").Value&"&nbsp;"
			next
		else
			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
			window.spAccHead.innerHTML=""
		End if
	end if
End if

End Function

Function SetDate()
	Dim sFromYr,sToYr
	sFromYr = document.formname.hFromYr.Value
	sToYr = document.formname.hToYr.Value
	sFromYr = "01/04/"&Trim(sFromYr)
	sToYr = "31/03/"&sToYr
	document.formname.ctlVouFromDate.setMinDate() = sFromYr
	document.formname.ctlVouToDate.setMinDate() = sFromYr
	document.formname.ctlVouFromDate.setMaxDate() = sToYr
	document.formname.ctlVouToDate.setMaxDate() = sToYr


End Function


</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">

<form method="POST" name="formname" action="AppVoucherList.asp">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hVoucherName" value="">
<input type="hidden" name="horgName" value="">
<input type="hidden" name="hAccHead" value="0">
<input type="hidden" name="hFromDate" value="0">
<input type="hidden" name="hToDate" value="0">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Approve Voucher
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
								<td class="TabCell" valign="bottom" align="center" width="96">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher List
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Voucher</td>
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
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="90">Organization</td>
                            <td class="FieldCell">
							<select size="1" name="selUnitId" onchange="popPartType()" class="FormElem" >
							   <OPTION value="0">Select a Unit</option>
							   <%populateOrganizationListDB%>
							</select>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="90">Voucher Type</td>
                            <td class="FieldCell">
							<select size="1" name="selVoucher" class="FormElem" onChange="DisplayBook()">
								<OPTION value="0">Select a Voucher </option>
								<OPTION value="01">Cash Voucher</option>
								<OPTION value="02">Bank Voucher</option>
								<OPTION value="04">Purchase Voucher</option>
								<OPTION value="05">Sales Voucher</option>
								<OPTION value="06">Debit Voucher</option>
								<OPTION value="07">Credit Voucher</option>
								<OPTION value="08">General Voucher</option>
							</select>
                            </td>
                                </tr>
                                <tr>
									<td class="FieldCell" width="90" valign="top">Book</td>
									<td class="FieldCell">
										<select size="1" name="selBook" class="FormElem">
											<option value="S">Select Book</option>
										</select>
									</td>
                                </tr>
                                <tr>
									<td class="FieldCell" width="90" valign="top"></td>
									<td class="FieldCell">
																<table class="ExcelTable" cellSpacing="0" cellPadding="2" border="0">
																<tbody>
																<tr>
																	<td vAlign="center" class="ExcelHeaderCell" align="center">Viewed
                                                                      By</td>
																<td vAlign="center" align="center" class="ExcelHeaderCell">From&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell">To&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell"></td>
                                                                </tr>
																<tr>
																	<td class="FieldCellSub"><input type="radio" value="VouNo" name="optCriteria" onclick="OptSelection()">
																	Voucher	Number&nbsp;</td>
																<td align="left" class="FieldCellSub"><input class="formelem"  size="11" name="txtNoFrom"></td>
                                                                  <td align="left" class="FieldCellSub"><input class="formelem"  size="11" name="txtNoTo"></td>
                                                                  <td align="left" class="FieldCellSub">&nbsp;</td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"><input onclick="OptSelection()" type="radio" value="VouDate" name="optCriteria">
                                                                    Voucher Date</td>
                                                           <td align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouFromDate")
 %>
 </td>
                                                                  <td align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouToDate")
 %>

</td>
                                                                  <td align="left" class="FieldCellSub">

&nbsp;

</td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"><input type="radio" onclick="OptSelection()"  value="Amount" name="optCriteria">
                                                                    Amount&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtGAmount"></td>
                                                                  <td align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtLAmount"></td>
                                                                  <td align="left" class="FieldCellSub"></td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"><input type="radio" onclick="OptSelection()" value="AccHead" name="optCriteria">
                                                                    Account Head</td>
                                                                  <td colSpan="3" align="left" class="FieldCellSub">
																	<select class="formelem" disabled OnChange="SelectAccHead()" size="1" name="SelAccHead">
																		  <option value="0">Select Option</option>
																		 <option value="G">General Ledger</option>
																	 </select>
                                                                   </td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"></td>
                                                                   <td colSpan="3" class="FieldCellSub"><span id="spAccHead" class="DataOnly"></span>&nbsp;</td>
                                                                </tr>
                                                              </tbody>
                                                            </table>
									</td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
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
																<input type="button" value="Next" name="B8" class="ActionButton" onClick="validate()" >
                                                                <input type="reset" value="Reset" name="B9" class="ActionButton" >
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
</BODY>
</HTML>