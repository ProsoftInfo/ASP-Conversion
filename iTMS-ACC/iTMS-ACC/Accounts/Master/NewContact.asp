<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	NewContact.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	Sep 27,2011
	'Modified By				:
	'Modified By				:
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
<!--#include file="../../include/sessionVerify.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
dim sQuery,objRs,iContactNo,sCallTy,Temparr,sAction



Set objRs = Server.CreateObject("ADODB.RecordSet")

iContactNo = Request.QueryString("ContactNo")
'Response.Write "<p>iContactNo = " & iContactNo

sCallTy = Request("hCallTy")

if trim(iContactNo)="" then
	sAction = "CREATE"
else
	sAction = "EDIT"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/Cancel.js"></SCRIPT>

<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>

<script type="application/xml" data-itms-xml-island="1" ID="OutData" ></script>
<script type="application/xml" data-itms-xml-island="1" id="ContactData"><Root/></script>

<script type="application/xml" data-itms-xml-island="1" ID="PartyData"><Party/></script>

<script>
window.__itmsPopupCompat = { type: "newContact" };
</script>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="popPartyDet('<%=iContactNo%>')">
<form method="POST" name="formname">

<input type="Hidden" name="hContactNumber" value="<%=iContactNo%>">
<input type="Hidden" name="hAction" value="<%=sAction%>">
<input type="hidden" name="hCreatedBy" value="<%=getUserID%>">

<Input type="hidden" name="hPartyName" value="">
<Input type="hidden" name="hParCode" value="">
<Input type="hidden" name="hUnitId" value="<%=Session("organizationcode")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td align="center" class="PageTitle">
		<%
			if trim(iContactNo)<>"" then
				Response.Write "Contact Amendment"
			else
				Response.Write "Contact Creation"
			end if
		%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%">
			<TR>
					<td height="20px" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="60px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Details
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="FieldCell" width="115px"> Contact Name</td>
															<td class="FieldCell">
																<Input type="text" size="71" name="txtContactName" value="" class="FormElem">&nbsp;
																<input type="checkbox" name="chkActive" value="1" class="FormElem">&nbsp;In-Active
                                                            </td>
														</tr>
														<tr>
															<td class="FieldCell" width="115px">Designation</td>
															<td class="FieldCell">
																<Input type="text" size="51" name="txtDesignation" value="" class="FormElem">
                                                            </td>
														</tr>

														<tr>
															<td class=FieldCell width="115px">ContactPerson For</td>
															<td class="FieldCell">
																<Input type="text" size="51" name="txtContactPersonFor" value="" class="FormElem">
                                                            </td>
														</tr>

														<tr>
															<td class=FieldCell width="115px"> Party Code</td>
															<td class="FieldCell" valign="bottom">
															<a href="#" onclick="SelPartyPopup(); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" ></a>
															<span id="spParty" class="DataOnly"></span> &nbsp;
															<a style="width: 1em; height: 1em;" href="#" onclick="DelParty(); return false;" >
															<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Click here to delete the Party" width="12px" height="12px">
															</a>
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
															<td class="GroupTitleLeft" width="10px">&nbsp;
                                                            </td>
															<td class="GroupTitle" width="60px"><p align="center">Address
                                                            </td>
												</center>
															<td class="GroupTitleRight"><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class="GroupTable">
												<center>
                                                    <div align="left">
                                        <table cellpadding="0" cellspacing="0">
                                          <tr>
                                            <td class="MiddlePack" colspan="5"></td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">Address</p>
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtAddress1" size="81" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left"></td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtAddress2" size="81" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">City
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtCity" size="25" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">PIN
                                            </td>
                                            <td class="FieldCellSub"><input type="text" name="txtPinCode" size="7"  maxlength="6" class="FormElem">
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Phone
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtPhone" size="18" class="FormElem">
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">State
                                            </td>
                                            <td class="FieldCellSub"><input type="text" name="txtState" size="35" class="FormElem">
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Fax
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtFax" size="18" class="FormElem">
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">Country
                                            </td>
                                            <td class="FieldCellSub"><input type="text" name="txtCountry" size="25" class="FormElem">
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Mobile
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtMobileNo" size="18" class="FormElem">
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub">E-mail ID
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtEmail" size="35" class="FormElem">
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub">URL
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtWebsite" size="25" class="FormElem">
                                          </td>
                                          </tr>
                                        <tr>
                                          <td class="MiddlePack" colspan="5"></td>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="ActionCell">
                                                                <input type="button" value="Save"		name="B2"			class="ActionButton"  onClick="PageSubmit()">
                                                                <input type="button" value="Close"		name="B3"			class="ActionButton"  onClick="GoToMain()">
                                                                <input type="button" value="Preview"	name="btnPreveiw"	class="ActionButton"  onClick="ViewData()">
                                                                <input type="reset"  value="Reset"		name="B1"			class="ActionButton"  >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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
