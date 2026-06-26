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

<SCRIPT LANGUAGE=javascript>
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
<XML ID="Party">
<Party PartyName="" PartyShortName="" Add1="" Add2="" City="" Pin="" State="" Country="" EMail="" ITPan="" Phone="" Fax="" Mobile="" Url="" />
</XML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
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

<script language="VBScript">
Function AddDetails()
	Dim Root,newElem,objhttp
	Set Root = Party.documentElement
	Root.Attributes.Item(0).Nodevalue = document.formname.txtName.value
	Root.Attributes.Item(1).Nodevalue = document.formname.txtShortName.value
	Root.Attributes.Item(2).Nodevalue = document.formname.txtAddress1.value
	Root.Attributes.Item(3).Nodevalue = document.formname.txtAddress2.value
	Root.Attributes.Item(4).Nodevalue = document.formname.txtCity.value
	Root.Attributes.Item(5).Nodevalue = document.formname.txtPinCode.value
	Root.Attributes.Item(6).Nodevalue = document.formname.txtState.value
	Root.Attributes.Item(7).Nodevalue = document.formname.txtCountry.value
	Root.Attributes.Item(8).Nodevalue = document.formname.txtEmail.value
	Root.Attributes.Item(9).Nodevalue = document.formname.txtPanNo.value
	Root.Attributes.Item(10).Nodevalue = document.formname.txtPhone.value
	Root.Attributes.Item(11).Nodevalue = document.formname.txtFax.value
	Root.Attributes.Item(12).Nodevalue = document.formname.txtMobileNo.value
	Root.Attributes.Item(13).Nodevalue = document.formname.txtWebsite.value
	
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	
	objhttp.Open "POST","MsiParUpdate.asp?", false
	objhttp.send Party.XMLDocument
	
	IF objhttp.responseText = "" Then
		MsgBox "Party Created "
		window.returnValue= document.formname.txtName.value
		window.close()
		Exit function
	Else
		'Msgbox "Error While Creating "
		alert objhttp.responseText
	End IF
		
End Function
</SCRIPT>
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
				       	             <a href="javascript:popPartyList()"><span style="cursor: hand" Title="View Contra Details" >
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
