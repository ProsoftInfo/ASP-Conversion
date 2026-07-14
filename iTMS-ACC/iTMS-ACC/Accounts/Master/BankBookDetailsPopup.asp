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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
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
<script type="application/xml" data-itms-xml-island="1" ID="BookData" data-src="../xmldata/PartyType.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="BankBookDet"><Root/></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "bankBookDetailsPopup" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
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
