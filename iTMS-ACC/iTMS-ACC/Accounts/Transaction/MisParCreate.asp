<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MisParCreate.asp
	'Module Name				:	ACCOUNTS (Transaction Creation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 23, 2004
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
Dim iRecCount
iRecCount=1


if iRecCount =0 then%>		
<HTML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<script src="../../scripts/itms-modern-compat.js"></script>
<script>
<!--
	function msgbox(strr)
	{
			alert(strr);
			window.location.href = "../AccountsHome.asp";
	}
//-->
</SCRIPT>
<BODY onLoad = "msgbox('Party Type has not been Created/Related')">
</BODY>
<HTML>
<%
Response.End
else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" ID="Party">
<Party PartyName="" PartyShortName="" Add1="" Add2="" City="" Pin="" State="" Country="" EMail="" ITPan="" Phone="" Fax="" Mobile="" Url="" />
</script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script SRC="../../scripts/cancel.js"></SCRIPT>
<script SRC="../../scripts/trim.js"></SCRIPT>
<script>
<!--

function CheckSubmit()
{
	var i,bFalg;

	bFlag=true;
	if (trim(document.formname.txtName.value)=="")
	{
		alert("Enter Party Name");
		document.formname.txtName.select();
		return false;
	}
	if (trim(document.formname.txtShortName.value) =="")
	{
		alert("Enter Party Code");
		document.formname.txtShortName.select();
		return false;
	}
	if (trim(document.formname.txtCity.value) =="")
	{
		alert("Enter Party City");
		document.formname.txtCity.select();
		return false;
	}		
	
	AddDetails()
}
//-->
</SCRIPT>

<script>
function dialogId() {
	var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
	return match ? decodeURIComponent(match[1]) : "";
}

function notifyDialogValue(id, value) {
	if (!id || !window.opener) {
		return;
	}
	try {
		if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, value);
			return;
		}
	} catch (ignoreDirectReturn) {}
	try {
		window.opener.postMessage({ type: "itms-dialog-return", id: id, value: value }, window.location.origin || "*");
	} catch (ignoreMessageReturn) {}
}

function returnModalValue(value) {
	var id;
	window.returnValue = value;
	window.returnvalue = value;
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(value);
		return;
	}
	id = dialogId();
	notifyDialogValue(id, value);
}

function AddDetails() {
	var doc = document.implementation.createDocument("", "Party", null);
	var root = doc.documentElement;
	var fields = [
		["PartyName", "txtName"],
		["PartyShortName", "txtShortName"],
		["Add1", "txtAddress1"],
		["Add2", "txtAddress2"],
		["City", "txtCity"],
		["Pin", "txtPinCode"],
		["State", "txtState"],
		["Country", "txtCountry"],
		["EMail", "txtEmail"],
		["ITPan", "txtPanNo"],
		["Phone", "txtPhone"],
		["Fax", "txtFax"],
		["Mobile", "txtMobileNo"],
		["Url", "txtWebsite"]
	];
	var xhr = new XMLHttpRequest();
	fields.forEach(function (pair) {
		root.setAttribute(pair[0], document.formname.elements[pair[1]] ? document.formname.elements[pair[1]].value : "");
	});
	xhr.open("POST", "MsiParUpdate.asp?", false);
	xhr.send(new XMLSerializer().serializeToString(doc));
	if (String(xhr.responseText || "") === "") {
		alert("Party Created ");
		returnModalValue(document.formname.txtName.value);
		window.close();
	} else {
		alert(xhr.responseText);
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type="Hidden" name="hUnitName" value="" >
<input type="Hidden" name="hUnitCode" value="" >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Miscellaneous Party Creation</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<!--TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)" height="13">
										<tr>
											<td align="center">Unit
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Group</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="72">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Contact</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="78">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Location</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="92">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Preference</td>
									</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                &nbsp;
								</td>
							</tr-->
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<!--tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr-->
							    <tr>
								<td align="center"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td width="100%" align="left">
									<table border="0" cellspacing="0"  cellpadding="0" class="ToolBarTable">
										<!--tr>
										<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
					<a href="#" onclick="popPartyList(); return false;"><span style="cursor: pointer" Title="View Contra Details" >
              						      <p align="center"><font face="Wingdings" color="#000000" size="5">4</font>
                                        </span></a>
					                    </td>
											
										</tr-->
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
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
															<td class=FieldCell width="115"> Party Name</td>
															<td class='FieldCell'><input type="text" name="txtName" size="30" maxlength="50" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="115"> Party Code</td>
															<td class='FieldCell'><input type="text" name="txtShortName" size="12" maxlength="10" class="Formelem"></td>
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
															<td class=GroupTable>
												<center>
                                                    <div align="left">
                                        <table cellpadding="0" cellspacing="0">
                                          <tr>
                                            <td class="MiddlePack" colspan="5"><p align="left"></td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Address</p>
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress1" size="25" class="Formelem" maxlength="50"></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left"></td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress2" size="25" class="Formelem" maxlength="50"></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">City</p>
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtCity" size="25" class="Formelem" maxlength="50"></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">PIN</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtPinCode" size="7"  maxlength="6" class="Formelem"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Phone</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtPhone" size="15" class="Formelem" maxlength="30"></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">State</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtState" size="20" class="Formelem" maxlength="30"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Fax</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtFax" size="15" class="Formelem" maxlength="30"></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Country</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtCountry" size="25" class="Formelem" maxlength="50"></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Mobile
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtMobileNo" size="15" class="Formelem" maxlength="30">
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub"><p align="left">E-mail ID</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtEmail" size="20" class="Formelem" maxlength="50"></p>
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub"><p align="left">URL</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtWebsite" size="20" class="Formelem" maxlength="30"></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub" width="165">IT PAN No</td>
                                            <td class="FieldCellSub"> <input type="text" name="txtPanNo" size="15" class="Formelem" maxlength="50"> </td>
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
                                                            <table border="0" cellspacing="0" class="ExcelTable" cellpadding="0" width="100%">
                                                        <tr>
                                                    <td>
                                                    <!--table border="0" cellpadding="0" cellspacing="0">
                                                <tr>
                                            <td colspan="3">
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCellSub width="165" valign="top"> Transaction Units</td>
															<td class='FieldCellSub'>
															<select size="4" name="selUnitId" multiple class="FormElem">
																<option value="0" selected > All Units</option>
																<%'populateUnit%>
    														</select>
    														</td>
														</tr>
													</table>
                                            </td>
                                                </tr-->
                                                <!--tr>
                                            <td>
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                        <tr>
                                                    <td class="FieldCellSub" width="165">Excise ECC Number</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtECCNo" size="15" class="Formelem" maxlength="50"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="165">Sales Tax Number - Local</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtSalesLocal" size="15" class="Formelem" maxlength="50"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="165">Sales Tax Number - Central</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtSalesCentral" size="15" class="Formelem" maxlength="50"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="165">IT PAN No</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtPanNo" size="15" class="Formelem" maxlength="50"> </td>
                                                        </tr>
                                                            </table >
                                            </td-->
                                            <!--td class="ClearPixel">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                            </td-->
                                            <!--td valign="top" align="right">
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                            
                                                        <tr>
                                                    <td class="FieldCellSub" rowspan="3" valign="top">
                                                    <input type="checkbox" name="chkGroupCompany" value="1" onClick="EnableGroup(this)" class="FormElem"> Group Company</td>
                                                    <td class="FieldCellSub">
                                                     <input type="radio" value="P" name="radGroupType" disabled="true" checked class="FormElem"> Parent </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub"><input type="radio" value="C" disabled="true" name="radGroupType" class="FormElem"> Child </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub"><input type="radio" value="B" disabled="true" name="radGroupType" class="FormElem"> Parent / Child </td>
                                                        </tr>
                                                            </table>
                                            </td>
                                                </tr>
                                                    </table-->
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Add" name="B2" class="ActionButton" onClick="CheckSubmit()"> 
                                                                <!--input type="button" value="Cancel" name="B3" onClick="Cancel('../AccountsHome.asp')"  class="ActionButton" -->
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
<%
end if	

%>
