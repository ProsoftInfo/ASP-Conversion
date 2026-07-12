<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BankDetails.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 04, 2011
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
	Dim sOrgID,sBookCode,sBookNumber,nAccHead,sQuery,objRs
	set objRs = server.CreateObject("ADODB.Recordset")
	
	sOrgID = Trim(Request("OrgCode"))
	sBookCode  = Trim(Request("BookCode"))
	sBookNumber = Trim(Request("BookNumber"))
	'nAccHead = Trim(Request("FromAcc"))
	
	
	sQuery=" SELECT isNull(BankAddress1,''), isNull(BankAddress2,''), isNull(City,''), isNull(State,''), isNull(Country,''),isNull(PinCode,''), isNull(PhoneNos,''),"&_
			"isNull(MobileNos,''), isNull(FaxNos,''), isNull(EMailId,''), isNull(WebSiteURL,''),isNull( PrintCheques,''), isNull(PrintPayInSlip,''), isNull(AccountType,''),"&_
			"isNull(AccountNo,''), isNull(CreditLimit,0), isNull(OverDraftLimit,0),isNull( DiscountingLimit,0),isNull(LCLimit,0),isNull( SwiftCode,''),isnull(BankChargesHead,0),"&_
			"isnull(BillDiscountingHead,0),BankName FROM Acc_M_BankDetails where "&_	
			" OUDefinitionID='"&sOrgID&"' and BookCode=02 and BookNumber="&sBookNumber
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
		
	set objRs.ActiveConnection = nothing
	
	Dim sBankName,sAddress1,sAddress2,sCity,sState,sCountry,sPinCode,sPhoneNo
	Dim sMobileNo,sFax,sEmailID,sWebsiteURL,sPrintCheque,sPrintPaySlip,sAccType
	Dim nAccNo,nCreditLimit,nOverDraftLimit,nDisLimit,nLCLimit,nSwiftCode
	Dim nChargeHead,nDisHead
	
	If not objRs.EOF then
			
			sBankName	= trim(objrs(22))
			sAddress1	= trim(objrs(0))
			sAddress2	= trim(objrs(1))
			sCity		= trim(objrs(2))
			sState		= trim(objrs(3))
			sCountry	= trim(objrs(4))
			sPinCode	= trim(objrs(5))
			sPhoneNo	= trim(objrs(6))
			sMobileNo	= trim(objrs(7))
			sFax		= trim(objrs(8))
			sEmailID	= trim(objrs(9))
			sWebsiteURL	= trim(objrs(10))
			
			sPrintCheque	= trim(objrs(11))
			sPrintPaySlip	= trim(objrs(12))
			sAccType		= trim(objrs(13))
			nAccNo			= trim(objrs(14))
			nCreditLimit	= trim(objrs(15))
			nOverDraftLimit	= trim(objrs(16))
			nDisLimit		= trim(objrs(17))
			nLCLimit		= trim(objrs(18))
			nSwiftCode		= trim(objrs(19))
			nChargeHead		= trim(objrs(20))
			nDisHead		= trim(objrs(21))
			
	End IF
	objRs.Close 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" ID="BookData" data-src="../xmldata/PartyType.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="BankBookDet"><Root/></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT>
function getXmlAttr(node, name, index) {
	var attr = name ? node.attributes.getNamedItem(name) : null;
	if (!attr && typeof index === "number" && node.attributes.item(index)) {
		attr = node.attributes.item(index);
	}
	return attr ? attr.value : "";
}

function setFieldValue(name, value) {
	var field = document.formname.elements[name];
	if (field) {
		field.value = value;
	}
}

function setRadioValue(name, value) {
	var fields = document.formname.elements[name];
	if (!fields) {
		return;
	}
	if (fields.length === undefined) {
		fields.checked = fields.value === String(value);
		return;
	}
	for (var i = 0; i < fields.length; i += 1) {
		fields[i].checked = fields[i].value === String(value);
	}
}

function popBankBook() {
	var root = window.BookData && BookData.documentElement;
	var select = document.formname.selBankBook;
	if (!root || !select) {
		return;
	}
	for (var i = 0; i < root.childNodes.length; i += 1) {
		var headerNode = root.childNodes[i];
		if (headerNode.nodeType !== 1) {
			continue;
		}
		select.options[select.options.length] = new Option(getXmlAttr(headerNode, null, 1), getXmlAttr(headerNode, null, 0));
	}
}

function DisplayBookDet() {
	var iUnitNo = document.formname.hOrgID.value;
	var iBookNo = document.formname.hBookNo.value;
	var objhttp = new XMLHttpRequest();
	objhttp.open("GET", "../Master/XMLBankBookDetail.asp?orgID=" + encodeURIComponent(iUnitNo) + "&BookNo=" + encodeURIComponent(iBookNo), false);
	objhttp.send(null);

	if (objhttp.responseText && objhttp.responseText.replace(/^\s+|\s+$/g, "") !== "") {
		BookData.loadXML(objhttp.responseText);
		document.formname.hActionFlag.value = "U";
		popBankBookDetail();
	} else {
		document.formname.reset();
		document.formname.hActionFlag.value = "I";
	}
}

function popBankBookDetail() {
	var root = window.BookData && BookData.documentElement;
	if (!root) {
		return;
	}
	for (var i = 0; i < root.childNodes.length; i += 1) {
		var headerNode = root.childNodes[i];
		if (headerNode.nodeType !== 1) {
			continue;
		}
		var bankName = document.getElementById("BankName");
		if (bankName) {
			bankName.innerHTML = getXmlAttr(headerNode, "BankName");
		}
		setFieldValue("txtAddress1", getXmlAttr(headerNode, "BankAddress1"));
		setFieldValue("txtAddress2", getXmlAttr(headerNode, "BankAddress2"));
		setFieldValue("txtCity", getXmlAttr(headerNode, "City"));
		setFieldValue("txtState", getXmlAttr(headerNode, "State"));
		setFieldValue("txtCountry", getXmlAttr(headerNode, "Country"));
		setFieldValue("txtPinCode", getXmlAttr(headerNode, "PinCode"));
		setFieldValue("txtPhone", getXmlAttr(headerNode, "PhoneNos"));
		setFieldValue("txtMobileNo", getXmlAttr(headerNode, "MobileNos"));
		setFieldValue("txtFax", getXmlAttr(headerNode, "FaxNos"));
		setFieldValue("txtEmail", getXmlAttr(headerNode, "EMailId"));
		setFieldValue("txtWebsite", getXmlAttr(headerNode, "WebSiteURL"));
		setRadioValue("optCheque", getXmlAttr(headerNode, "PrintCheques"));
		setRadioValue("optPayIn", getXmlAttr(headerNode, "PrintPayInSlip"));

		var accountType = getXmlAttr(headerNode, "AccountType");
		var creditLimit = document.formname.elements.txtCreditLimit;
		var selAccType = document.formname.elements.selAccType;
		if (creditLimit) {
			creditLimit.readOnly = accountType === "CU";
			if (accountType === "CU") {
				creditLimit.value = "0";
			}
		}
		if (selAccType) {
			selAccType.selectedIndex = accountType === "CU" ? 1 : 2;
		}

		setFieldValue("txtAccNo", getXmlAttr(headerNode, "AccountNo"));
		setFieldValue("txtCreditLimit", getXmlAttr(headerNode, "CreditLimit"));
		setFieldValue("txtODLimit", getXmlAttr(headerNode, "OverDraftLimit"));
		setFieldValue("txtDiscountLimit", getXmlAttr(headerNode, "DiscountingLimit"));
		setFieldValue("txtLCLimit", getXmlAttr(headerNode, "LCLimit"));
		setFieldValue("txtswitCode", getXmlAttr(headerNode, "SwiftCode"));
		setFieldValue("hChargestHead", getXmlAttr(headerNode, "ChargeHead"));
		setFieldValue("hDiscountHead", getXmlAttr(headerNode, "DiscountHead"));
	}
}

function checkCredit() {
	var selAccType = document.formname.elements.selAccType;
	var creditLimit = document.formname.elements.txtCreditLimit;
	if (!selAccType || !creditLimit) {
		return;
	}
	if (selAccType.selectedIndex === 2) {
		creditLimit.readOnly = false;
	} else {
		creditLimit.value = "0";
		creditLimit.readOnly = true;
	}
}
</script>
</HEAD>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >

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
							<!--<tr>
								<td align="center">
                                &nbsp;
                                <p>&nbsp;
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">-->
														
														<!--<tr>
															<td class=FieldCell width="125"> Select&nbsp;
                                                              Bank Book&nbsp;</td>
															<td class='FieldCell'>
                                                          <select size="1" name="selBankBook" class="FormElem" onChange="DisplayBookDet()">
																<OPTION value="0">Select a Bank Book</option>
														</select>
                                                            </td>
														</tr>-->
														<!--<tr>
															<td class=FieldCell width="125"> Charges
                                                              A/C Head&nbsp;</td>
															<td class='FieldCell'>
															<a href="#">                                     
															<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" onClick="popAccList('C'); return false;" width="15" height="15">
                                                           &nbsp;<span class="DataOnly" id="spCharges"></span>
															</a>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="125"> Discounting
                                                              A/C Head&nbsp;</td>
															<td class='FieldCell'>
															<a href="#">                                     
															<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" onClick="popAccList('D'); return false;" width="15" height="15">
                                                            &nbsp;<span class="DataOnly" id="spDisCount"></span>
															</a>
                                                            </td>
														</tr>-->
														<!--<tr>
															<td class=FieldCell width="125"> BankName</td>
															<td class='FieldCell'>
																<span class="Dataonly" Id="BankName"><%=sBankName%></span>
                                                            </td>
														</tr>-->
													<!--</table>
								</td>
								<td align="center">
								</td>
							</tr>-->
                            
							<tr>
								<td align="center">
								</td>
								<td valign="top">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
															<td class=FieldCell colspan=2> BankName
																<span class="Dataonly" Id="BankName"><%=sBankName%></span>
                                                            </td>
														</tr>
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
                                            <td class="FieldCellSub" colspan="5"><p align="left"></td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Address</p></td>
                                            <td class="FieldCellSub" colspan="4"><p align="left">
												<span class="Dataonly"><%=sAddress1%>,<%=sAddress2%></span></p>
                                            </td>
                                          </tr>
                                          
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">City</p></td>
                                            <td class="FieldCellSub" colspan="4"><p align="left">
												<span class="Dataonly"><%If sCity<> "" Then Response.Write sCity else Response.Write "-" End IF%></span></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">PIN</p></td>
                                            <td class="FieldCellSub">
												<span class="Dataonly"><%If sPinCode<> "" Then Response.Write sPinCode else Response.Write "-" End IF%></span>
                                            </td>
                                            <td class="FieldCellSub"></td>
											<td class="FieldCellSub"><p align="left">Phone</p></td>
											<td class="FieldCellSub"><p align="left"><span class="Dataonly"><%If sPhoneNo <> "" Then Response.Write sPhoneNo else Response.Write "-" End IF%></span></p></td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">State</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><span class="Dataonly"><%If sState <> "" Then Response.Write sState else Response.Write "-" End IF%></span></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Fax</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><span class="Dataonly"><%If sFax <> "" Then Response.Write sFax else Response.Write "-" End IF%></span></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Country</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><span class="Dataonly"><%If sCountry <> "" Then Response.Write sCountry else Response.Write "-" End IF%></span></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Mobile
                                          </td>
                                          <td class="FieldCellSub"><span class="Dataonly"><%If sMobileNo <> "" Then Response.Write sMobileNo else Response.Write "-" End IF%></span>
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub"><p align="left">E-mail ID</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><span class="Dataonly"><%If sEmailID<> "" Then Response.Write sEmailID else Response.Write "-" End IF%></span></p>
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub"><p align="left">URL</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><span class="Dataonly"><%If sWebsiteURL <> "" Then Response.Write sWebsiteURL else Response.Write "-" End IF%></span></p>
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
                                            <td class="FieldCellSub"><span class="DataOnly">
												<%If sAccType = "CU" Then 
													Response.Write "Current Account"
												Elseif sAccType= "CC" Then
													Response.Write "Cash Credit Account"
												End IF
												%></span>
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
                                            <td class="FieldCellSub"><span class="DataOnly"><%If nAccNo <> "" Then Response.Write nAccNo else Response.Write "-" End IF%>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Swift
                                            Code</p>
                                          </td>
                                          <td class="FieldCellSub">
                                            <span class=Dataonly><%If nSwiftCode <> "" Then Response.Write nSwiftCode Else Response.Write "-" End If%></span>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">Print Cheque
                                            </td>
                                            <td class="FieldCellSub">
                                            <table border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                                <td class="FieldCellSub"><input type="radio" value="1" name="optCheque" <%If sPrintCheque="1" Then Response.Write "checked"%>></td>
                                                <td class="FieldCellSub">Yes</td>
                                                <td class="FieldCellSub" width="10"></td>
                                                <td class="FieldCellSub"><input type="radio" value="0" name="optCheque"  <%If sPrintCheque="0" Then Response.Write "checked"%>></td>
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
                                                <td class="FieldCellSub"><input type="radio" value="1" name="optPayIn"  <%If sPrintPaySlip ="1" Then Response.Write "checked"%>></td>
                                                <td class="FieldCellSub">Yes</td>
                                                <td class="FieldCellSub" width="10"></td>
                                                <td class="FieldCellSub"><input type="radio" value="0" name="optPayIn"  <%If sPrintPaySlip="0" Then Response.Write "checked"%>></td>
                                                <td class="FieldCellSub">No</td>
                                              </tr>
                                            </table>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Credit
                                              Limit</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><span class="Dataonly">Rs.<%If nCreditLimit <> "" Then Response.Write nCreditLimit Else Response.Write "0.00" End IF%>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Over Draft
                                            Limit
                                          </td>
                                          <td class="FieldCellSub"><span class=Dataonly>Rs.<%If nOverDraftLimit  <> "" Then Response.Write nOverDraftLimit Else Response.Write "0.00" End IF%></span>
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub">Discounting
                                            Limit
                                          </td>
                                          <td class="FieldCellSub"><span class="DataOnly">Rs.<%If nDisHead  <> "" Then Response.Write nDisHead  Else Response.Write "0.00" End IF%></span>
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub">LC Limit
                                          </td>
                                          <td class="FieldCellSub"><span class=DataOnly>Rs.<%If nLCLimit <> "" Then Response.Write nLCLimit Else Response.Write "0.00" End IF%></span>
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
                                                                <input type="button" value="Close" name="B9" onClick="window.close()" class="ActionButton"  >
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
