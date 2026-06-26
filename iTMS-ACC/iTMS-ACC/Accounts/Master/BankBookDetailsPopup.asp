<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BankBookDetailsEntry.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 02, 2011
	'Modified On				:	
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	Code
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
<%
	Dim sOrgID,sBookCode,sBookNumber,nAccHead
	sOrgID = Trim(Request("OrgCode"))
	sBookCode  = Trim(Request("BookCode"))
	sBookNumber = Trim(Request("BookNumber"))
	nAccHead = Trim(Request("FromAcc"))
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="BookData" src="../xmldata/PartyType.xml"></XML>
<xml id="GLHeadData"><Root></Root></xml>
<xml id="BankBookDet"><Root/></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT language="vbscript">

Function popBankBook()
dim Root
Set Root = BookData.documentElement
For Each HeaderNode In Root.childNodes
		document.formname.selBankBook.length = document.formname.selBankBook.length+1
		document.formname.selBankBook.options(document.formname.selBankBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
		document.formname.selBankBook.options(document.formname.selBankBook.length-1).Value =HeaderNode.Attributes.Item(0).nodeValue
next
end Function

Function DisplayBookDet()
	dim iUnitNo,iBookNo,iUnitIndex,iBookIndex
	dim Root

	iUnitNo= document.formname.hOrgID.value
	iBookNo= document.formname.hBookNo.value
	
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLBankBookDetail.asp?orgID=" & iUnitNo&"&BookNo="&iBookNo, false
	objhttp.send
	
	if objhttp.responseXML.xml <> "" then
		BookData.loadXML objhttp.responseXML.xml
		document.formname.hActionFlag.value="U"
		popBankBookDetail
	else
		document.formname.reset
		'document.formname.selUnitId.selectedIndex=iUnitIndex
		'document.formname.selBankBook.selectedIndex=iBookIndex
		document.formname.hActionFlag.value="I"
		window.spCharges.innerHTML=""
		window.spDisCount.innerHTML=""
		'MsgBox "No Details Exist Please Enter"
	end if

end Function

Function popBankBookDetail()
dim Root
Set Root = BookData.documentElement
For Each HeaderNode In Root.childNodes
			
		document.formname.txtName.value=HeaderNode.Attributes.getNamedItem("BankName").value
		document.formname.txtAddress1.value=HeaderNode.Attributes.getNamedItem("BankAddress1").value
		document.formname.txtAddress2.value=HeaderNode.Attributes.getNamedItem("BankAddress2").value
		document.formname.txtCity.value=HeaderNode.Attributes.getNamedItem("City").value
		document.formname.txtState.value=HeaderNode.Attributes.getNamedItem("State").value
		document.formname.txtCountry.value=HeaderNode.Attributes.getNamedItem("Country").value
		document.formname.txtPinCode.value=HeaderNode.Attributes.getNamedItem("PinCode").value
		document.formname.txtPhone.value=HeaderNode.Attributes.getNamedItem("PhoneNos").value
		document.formname.txtMobileNo.value=HeaderNode.Attributes.getNamedItem("MobileNos").value
		document.formname.txtFax.value=HeaderNode.Attributes.getNamedItem("FaxNos").value
		document.formname.txtEmail.value=HeaderNode.Attributes.getNamedItem("EMailId").value
		document.formname.txtWebsite.value=HeaderNode.Attributes.getNamedItem("WebSiteURL").value
		
		 if cint(HeaderNode.Attributes.getNamedItem("PrintCheques").value)=1 then
			document.formname.optCheque(0).checked=true
		 else
			document.formname.optCheque(1).checked=true
		 end if
		 
		 if cint(HeaderNode.Attributes.getNamedItem("PrintPayInSlip").value)=1 then
			document.formname.optPayIn(0).checked=true
		 else
			document.formname.optPayIn(1).checked=true
		 end if
		 
		 if HeaderNode.Attributes.getNamedItem("AccountType").value="CU" then
			document.formname.txtCreditLimit.value="0"
			document.formname.txtCreditLimit.readOnly=true
			document.formname.selAccType.selectedIndex=1
		 else
			document.formname.txtCreditLimit.readOnly=false
			document.formname.selAccType.selectedIndex=2
		 end if

		document.formname.txtAccNo.value=HeaderNode.Attributes.getNamedItem("AccountNo").value
		document.formname.txtCreditLimit.value=HeaderNode.Attributes.getNamedItem("CreditLimit").value
		document.formname.txtODLimit.value=HeaderNode.Attributes.getNamedItem("OverDraftLimit").value
		document.formname.txtDiscountLimit.value=HeaderNode.Attributes.getNamedItem("DiscountingLimit").value
		document.formname.txtLCLimit.value=HeaderNode.Attributes.getNamedItem("LCLimit").value
		document.formname.txtswitCode.value=HeaderNode.Attributes.getNamedItem("SwiftCode").value
		document.formname.hChargestHead.value=HeaderNode.Attributes.getNamedItem("ChargeHead").value
		
		window.spCharges.innerHTML=HeaderNode.Attributes.getNamedItem("ChargeHeadName").value
		document.formname.hDiscountHead.value=HeaderNode.Attributes.getNamedItem("DiscountHead").value
		window.spDisCount.innerHTML=HeaderNode.Attributes.getNamedItem("DiscountHeadName").value
next

end Function

Function popAccList_Old(bFlag)
dim iUnitNo,saTemp,iGlHead,sRetVal,OutValue


	sOrgId= document.formname.hOrgID.value
	
	OutValue= showModalDialog("BookGLHeadSelection.asp?orgId="+sOrgId,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")
	
	arrTemp = split(OutValue,":")
	while UBound(arrTemp) = 0 
		OutValue = showModalDialog("BookGLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
		
	sRetVal = OutValue
	

	if UBound(arrTemp) <= 1 then exit function
	
	if bFlag="C" then
		document.formname.hChargestHead.value=arrTemp(0)
		window.spCharges.innerHTML=arrTemp(1)
	else
		document.formname.hDiscountHead.value=arrTemp(0)
		window.spDisCount.innerHTML=arrTemp(1)
	end if

end Function

Function popAccList(bFlag)
	Dim sOrgId
	sOrgId = document.formname.hOrgID.value
	
	set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?orgId="+sOrgId,GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	    sQuery = OutValue.getAttribute("PassQuery")
	    if OutValue.getAttribute("Action")="CLOSE" then exit function

	while OutValue.getAttribute("Action")<>"Done"
		set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		    sQuery = OutValue.getAttribute("PassQuery")
		    if OutValue.getAttribute("Action")="CLOSE" then exit function
	wend
	'alert(OutValue.xml)
	if OutValue.hasChildNodes() then
	    For each ndChild in OutValue.childNodes
			
			if bFlag="C" then
				document.formname.hChargestHead.value=ndChild.getAttribute("RetField0")
				window.spCharges.innerHTML=ndChild.getAttribute("RetField5")
			else
				document.formname.hDiscountHead.value=ndChild.getAttribute("RetField0")
				window.spDisCount.innerHTML=ndChild.getAttribute("RetField5")
			end if
			
	    Next
	end if

End Function

Function CheckSubmit()
	Dim objHttp
	
	if document.formname.selAccType.selectedIndex < 1 Then
		alert("Select Account Type")
		document.formname.selAccType.focus()
		Exit Function
	Elseif trim(document.formname.txtName.value) = ""  Then
		alert("Enter Bank Name")
		document.formname.txtName.focus()
		Exit Function
	Elseif trim(document.formname.txtAccNo.value) = "" Then
		alert("Enter Account No")
		document.formname.txtAccNo.focus()
		Exit Function
	Elseif (isNaN(parseFloat(document.formname.txtCreditLimit.value))) Then
		alert("Enter Numeric Value ")
		document.formname.txtCreditLimit.value="0"
		document.formname.txtCreditLimit.select()
		Exit Function
	Elseif (parseFloat(document.formname.txtCreditLimit.value)<0) Then
		alert("Enter Value Greater than or Equal to Zero")
		document.formname.txtCreditLimit.value="0"
		document.formname.txtCreditLimit.select()
		Exit Function
	Elseif (isNaN(parseFloat(document.formname.txtODLimit.value))) Then
		alert("Enter Numeric Value ")
		document.formname.txtODLimit.value="0"
		document.formname.txtODLimit.select()
		Exit Function
	Elseif (parseFloat(document.formname.txtODLimit.value)<0) Then
		alert("Enter Value Greater than or Equal to Zero")
		document.formname.txtODLimit.value="0"
		document.formname.txtODLimit.select()
		Exit Function
	Elseif isNaN(parseFloat(document.formname.txtLCLimit.value)) Then
		alert("Enter Numeric Value ")
		document.formname.txtLCLimit.value="0"
		document.formname.txtLCLimit.select()
		Exit Function
	elseif (parseFloat(document.formname.txtLCLimit.value)<0) Then
		alert("Enter Value Greater than or Equal to Zero")
		document.formname.txtLCLimit.value="0"
		document.formname.txtLCLimit.select()
		Exit Function
	Elseif (isNaN(parseFloat(document.formname.txtDiscountLimit.value))) Then
		alert("Enter Numeric Value ")
		document.formname.txtDiscountLimit.value="0"
		document.formname.txtDiscountLimit.select()
		Exit Function
	Elseif (parseFloat(document.formname.txtDiscountLimit.value)<0) Then
		alert("Enter Value Greater than or Equal to Zero")
		document.formname.txtDiscountLimit.value="0"
		document.formname.txtDiscountLimit.select()
		Exit Function
	End IF
	
	Dim Root,NewElem
	set Root = BankBookDet.documentElement
	
	set NewElem = BankBookDet.CreateElement("BankBook")
	NewElem.setAttribute("UnitID"),document.formname.hOrgID.value
	NewElem.setAttribute("BookNo"),document.formname.hBookNo.value
	NewElem.setAttribute("ActionFlag"),document.formname.hActionFlag.value
	
	NewElem.setAttribute("BankName"),document.formname.txtName.value
	NewElem.setAttribute("Address1"),document.formname.txtAddress1.value
	NewElem.setAttribute("Address2"),document.formname.txtAddress2.value
	NewElem.setAttribute("City"),document.formname.txtCity.value
	NewElem.setAttribute("State"),document.formname.txtState.value
	NewElem.setAttribute("Country"),document.formname.txtCountry.value
	NewElem.setAttribute("Pincode"),document.formname.txtPinCode.value
	NewElem.setAttribute("Phone"),document.formname.txtPhone.value
	NewElem.setAttribute("MobileNo"),document.formname.txtMobileNo.value
	NewElem.setAttribute("Fax"),document.formname.txtFax.value
	NewElem.setAttribute("EMail"),document.formname.txtEmail.value
	NewElem.setAttribute("WebSite"),document.formname.txtWebsite.value
	
	If document.formname.optCheque(0).checked Then
		sOption = document.formname.optCheque(0).value 
	Else
		sOption = document.formname.optCheque(1).value 
	End IF
	If document.formname.optPayIn(0).checked Then
		sOptionP = document.formname.optPayIn(0).value 
	Else
		sOptionP = document.formname.optPayIn(1).value 
	End IF
	
	NewElem.setAttribute("PrintCheque"),Trim(sOption)
	NewElem.setAttribute("PrintPayInSlip"),Trim(sOptionP)
	NewElem.setAttribute("AccountType"),document.formname.selAccType.value 
	NewElem.setAttribute("AccountNo"),document.formname.txtAccNo.value
	NewElem.setAttribute("CreditLimit"),document.formname.txtCreditLimit.value
	NewElem.setAttribute("ODLimit"),document.formname.txtODLimit.value
	NewElem.setAttribute("DiscountLimit"),document.formname.txtDiscountLimit.value
	NewElem.setAttribute("LCLimit"),document.formname.txtLCLimit.value
	NewElem.setAttribute("SwiftCode"),document.formname.txtswitCode.value
	NewElem.setAttribute("ChargesHead"),document.formname.hChargestHead.value
	NewElem.setAttribute("DiscountHead"),document.formname.hDiscountHead.value
	
	Root.appendchild NewElem
	
	set objHttp = CreateObject("Microsoft.XMLHTTP")
	objHttp.open "POST","BankBookDetailsUpdate.asp",False
	objHttp.send BankBookDet.XMLDocument
	
	If objHttp.responseText <> "" Then
		alert(objHttp.responseText)
		Exit Function
	Else
		alert("Bank Details Inserted Successfully")
		window.returnvalue = "Done"
	End IF
	window.close 
	
End Function

Function checkCredit()
	If document.formname.selAccType.selectedIndex = 2  Then
		document.formname.txtCreditLimit.readOnly=false
	Else
		document.formname.txtCreditLimit.value="0"
		document.formname.txtCreditLimit.readOnly=true
	End IF
End FUnction
</script>
<script language="javascript">
window.__itmsPopupCompat = { type: "bankBookDetailsPopup" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="DisplayBookDet()">

<form method="POST" name="formname" action="">
<input type="Hidden" name="hActionFlag" value="" >
<input type="Hidden" name="hDiscountHead" value="0" >
<input type="Hidden" name="hChargestHead" value="0" >
<input type="Hidden" name="hOrgID" value="<%=sOrgID%>" >
<input type="Hidden" name="hBookNo" value="<%=sBookNumber%>" >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Bank Details</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
                                &nbsp;
                                <p>&nbsp;
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">
														
														<!--<tr>
															<td class=FieldCell width="125"> Select&nbsp;
                                                              Bank Book&nbsp;</td>
															<td class='FieldCell'>
                                                          <select size="1" name="selBankBook" class="FormElem" onChange="DisplayBookDet()">
																<OPTION value="0">Select a Bank Book</option>
														</select>
                                                            </td>
														</tr>-->
														<tr>
															<td class=FieldCell width="125"> Charges
                                                              A/C Head&nbsp;</td>
															<td class='FieldCell'>
															<a href="#">                                     
															<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" onClick="popAccList('C')" width="15" height="15">
                                                           &nbsp;<span class="DataOnly" id="spCharges"></span>
															</a>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="125"> Discounting
                                                              A/C Head&nbsp;</td>
															<td class='FieldCell'>
															<a href="#">                                     
															<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" onClick="popAccList('D')" width="15" height="15">
                                                            &nbsp;<span class="DataOnly" id="spDisCount"></span>
															</a>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="125"> Bank
                                                              Name</td>
															<td class='FieldCell'>
                                                            <input type="text" name="txtName" size="32" maxlength="30" class="Formelem">
                                                            </td>
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="60"><p align="center">Address
                                                            </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable align="left">
												<center>
                                                    <div align="left">
                                        <table cellpadding="0" cellspacing="0">
                                          <tr>
                                            <td class="MiddlePack" colspan="5"><p align="left"></td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Address</p>
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress1" size="81" class="Formelem"></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left"></td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress2" size="81" class="Formelem"></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">City</p>
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtCity" size="25" class="Formelem"></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">PIN</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtPinCode" size="7" maxlength="6" class="Formelem"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Phone</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtPhone" size="18" class="Formelem"></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">State</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtState" size="35" class="Formelem"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Fax</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtFax" size="18" class="Formelem"></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Country</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtCountry" size="25" class="Formelem"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Mobile
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtMobileNo" size="18" class="Formelem">
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub"><p align="left">E-mail ID</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtEmail" size="35" class="Formelem"></p>
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub"><p align="left">URL</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtWebsite" size="25" class="Formelem"></p>
                                          </td>
                                          </tr>
                                        <tr>
                                          <td class="MiddlePack" colspan="5"><p align="left"></td>
                                        </tr>
                                        </table>
                                                    </div>
												</center>
                                                            </td>
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                                <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                                </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
                                                            <div align="left">
                                        <table cellpadding="0" cellspacing="0">
                                          <tr>
                                            <td class="MiddlePack" colspan="5"><p align="left"></td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Account
                                              Type</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><select size="1" onchange="checkCredit()" name="selAccType" class="formelem">
                                                <option value="S" selected>Select
                                                Type</option>
                                                <option value="CU">Current
                                                Account</option>
                                                <option value="CC">Cash Credit
                                                Account</option>
                                              </select></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left"></p>
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Account
                                              Number</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtAccNo" size="31" class="Formelem"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Swift
                                            Code</p>
                                          </td>
                                          <td class="FieldCellSub">
                                            <input type="text" name="txtswitCode" size="18" class="Formelem">
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">Print Cheque
                                            </td>
                                            <td class="FieldCellSub">
                                            <table border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                                <td class="FieldCellSub"><input type="radio" value="1" name="optCheque"></td>
                                                <td class="FieldCellSub">Yes</td>
                                                <td class="FieldCellSub" width="10"></td>
                                                <td class="FieldCellSub"><input type="radio" value="0" name="optCheque" checked></td>
                                                <td class="FieldCellSub">No</td>
                                              </tr>
                                            </table>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Print Pay-in
                                            slip
                                          </td>
                                          <td class="FieldCellSub">
                                            <table border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                                <td class="FieldCellSub"><input type="radio" value="1" name="optPayIn"></td>
                                                <td class="FieldCellSub">Yes</td>
                                                <td class="FieldCellSub" width="10"></td>
                                                <td class="FieldCellSub"><input type="radio" value="0" name="optPayIn" checked></td>
                                                <td class="FieldCellSub">No</td>
                                              </tr>
                                            </table>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Credit
                                              Limit</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left">Rs.
                                              <input type="text" name="txtCreditLimit" style="text-align:right" maxlength="13" size="15" class="Formelem" value="0"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Over Draft
                                            Limit
                                          </td>
                                          <td class="FieldCellSub">Rs. <input type="text" name="txtODLimit" style="text-align:right" maxlength="13" size="15" class="Formelem" value="0">
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub">Discounting
                                            Limit
                                          </td>
                                          <td class="FieldCellSub">Rs. <input type="text" name="txtDiscountLimit" style="text-align:right" maxlength="13" size="15" class="Formelem" value="0">
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub">LC Limit
                                          </td>
                                          <td class="FieldCellSub">Rs. <input type="text" name="txtLCLimit" style="text-align:right" maxlength="13" size="15" class="Formelem" value="0">
                                          </td>
                                          </tr>
                                        <tr>
                                          <td class="MiddlePack" colspan="5"><p align="left"></td>
                                        </tr>
                                        </table>
                                                            </div>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
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
                                                                <input type="button" value="Save" name="B2" class="ActionButton" onclick="CheckSubmit()" >&nbsp;
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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
