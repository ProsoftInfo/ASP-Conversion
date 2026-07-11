<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParDisplayContactDetails.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 20,2010
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
dim sQuery,objRs,iParty
Dim sPartyName,sOrgnPartyCode,sAddress1,sAddress2,sCity,sState
Dim sCountry,sPhoneNo,sMobileNo,sFaxNo,sEmail,sPinCode,sWebsite
iParty = Request.QueryString("PartyCode")
Set objRs = Server.CreateObject("ADODB.RecordSet")
sQuery= "Select PartyName,OrgnPartyCode,isNull(AddressLine1,''),isNull(AddressLine2,''),isNull(City,''),isNull(State,''),isNull(Country,''),isNull(PhoneNos,''),"&_
		"isNull(MobileNos,''),isNull(FaxNos,''),isNull(Email,''),isNull(PinCode,''),isNull(WebsiteURL,'') from APP_M_PartyMaster where PartyCode = "& iParty
objRs.Open sQuery,con
if not objRs.EOF then
	sPartyName = objRs(0)
	sOrgnPartyCode = objRs(1)
	sAddress1 = objRs(2)
	sAddress2 = objRs(3)
	sCity = objRs(4)
	sState = objRs(5)
	sCountry = objRs(6)
	sPhoneNo = objRs(7)
	sMobileNo = objRs(8)
	sFaxNo = objRs(9)
	sEmail = objRs(10)
	sPinCode = objRs(11)
	sWebsite = objRs(12)
end if
objRs.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/Cancel.js"></SCRIPT>

<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname">
<input type="Hidden" name="hUnitName" value="">
<input type="Hidden" name="hUnitCode" value="" >
<input type="Hidden" name="hPartyCode" value="<%=iParty%>">
<input type="Hidden" name="hOwnUnit" value="">

<input type="hidden" name="hInActive" value="0">
<input type="hidden" name="hCreatedBy" value="<%=getUserID%>">
<input type="hidden" name="hParentPartyCode" value="0">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
			Party Contact Details
		</p>
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
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
											<td class='FieldCell'>
											<Input type="text" size="71" name="txtPartyName" value="<%=sPartyName%>" class="FormElemRead" readonly>&nbsp;
											    </td>
										</tr>
										<tr>
											<td class=FieldCell width="115"> Party Code</td>
											<td class='FieldCell' valign=top><input type="text" name="txtShortName" size="20" maxlength="10" class="FormelemRead"  value="<%=sOrgnPartyCode%>" readonly >
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="1005">
								</td>
							</td>
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
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress1" size="81" class="FormelemRead" readonly value="<%=sAddress1%>" ></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left"></td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress2" size="81" class="FormelemRead" readonly value="<%=sAddress2%>"></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">City</p>
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtCity" size="25" class="FormelemRead" readonly value="<%=sCity%>" ></p>
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">PIN</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtPinCode" size="7"  maxlength="6" class="FormelemRead" readonly value="<%=sPinCode%>" ></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Phone</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtPhone" size="18" class="FormelemRead" readonly value="<%=sPhoneNo%>" ></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">State</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtState" size="35" class="FormelemRead" readonly value="<%=sState%>" ></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub"><p align="left">Fax</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtFax" size="18" class="FormelemRead" readonly value="<%=sFaxNo%>" ></p>
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left">Country</p>
                                            </td>
                                            <td class="FieldCellSub"><p align="left"><input type="text" name="txtCountry" size="25" class="FormelemRead" readonly value="<%=sCountry%>" ></p>
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Mobile
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtMobileNo" size="18" class="FormelemRead" readonly value="<%=sMobileNo%>" >
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub"><p align="left">E-mail ID</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtEmail" size="35" class="FormelemRead" readonly value="<%=sEmail%>" ></p>
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub"><p align="left">URL</p>
                                          </td>
                                          <td class="FieldCellSub"><p align="left"><input type="text" name="txtWebsite" size="25" class="FormelemRead" readonly value="<%=sWebsite%>" ></p>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Close" name="B3" class="ActionButton"  onClick="window.close()">
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
set objRs=nothing
%>
