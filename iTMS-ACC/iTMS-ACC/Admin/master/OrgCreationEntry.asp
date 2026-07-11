<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgCreationEntry.asp
	'Module Name				:	Inventory (Organization Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 11, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	orgCreationInsert.asp
	'Procedures/Functions Used	:	populateCountry
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
<!--#include File="../../include/IncludeDatePicker.asp" -->
<%
	Dim oDom,fs,Root,PGNode
	dim sorgID,sorgName,sorgShName,iorgNoUnits,sorgAddr1,sorgAddr2,sorgPIN,sorgCity
	dim sorgState,norgCountry,sorgCountry,sorgPhone,sorgFax,sorgEmail,sorgURL,sorgContactPerson
	dim norgCurrency,sorgTNGSTRCNo,sorgAreaCode,sorgCSTRCNo,dorgCSTRCDate

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set fs = CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(Server.MapPath("../xmldata/Organization.xml")) then
		oDOM.Load server.MapPath("../xmldata/Organization.xml")
		Set Root = oDOM.documentElement
		if Root.HaschildNodes() then
			For Each PGNode In Root.childNodes
				sorgID = trim(PGNode.Attributes.Item(0).nodeValue)
				sorgName = trim(PGNode.Attributes.Item(1).nodeValue)
				sorgShName = trim(PGNode.Attributes.Item(2).nodeValue)
				iorgNoUnits = trim(PGNode.Attributes.Item(3).nodeValue)
				norgCurrency = trim(PGNode.Attributes.Item(4).nodeValue)
				sorgAddr1 = trim(PGNode.Attributes.Item(5).nodeValue)
				sorgAddr2 = trim(PGNode.Attributes.Item(6).nodeValue)
				sorgPIN = trim(PGNode.Attributes.Item(7).nodeValue)
				sorgCity = trim(PGNode.Attributes.Item(8).nodeValue)
				sorgState = trim(PGNode.Attributes.Item(9).nodeValue)
				norgCountry = trim(PGNode.Attributes.Item(10).nodeValue)
				sorgCountry = trim(PGNode.Attributes.Item(11).nodeValue)
				sorgPhone = trim(PGNode.Attributes.Item(12).nodeValue)
				sorgFax = trim(PGNode.Attributes.Item(13).nodeValue)
				sorgEmail = trim(PGNode.Attributes.Item(14).nodeValue)
				sorgURL = trim(PGNode.Attributes.Item(15).nodeValue)
				sorgContactPerson = trim(PGNode.Attributes.Item(16).nodeValue)
				sorgTNGSTRCNo = trim(PGNode.Attributes.Item(17).nodeValue)
				sorgAreaCode = trim(PGNode.Attributes.Item(20).nodeValue)
				sorgCSTRCNo = trim(PGNode.Attributes.Item(18).nodeValue)
				dorgCSTRCDate = trim(PGNode.Attributes.Item(19).nodeValue)
			next
		end if
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Organization</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../scripts/orgCreate.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/OrgCreationCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init('<%=dorgCSTRCDate%>','<%=sorgCountry%>','<%=norgCurrency%>')">
<form method="POST" name="formname" action="">
<input type=hidden name=hOrgID value="<%=sorgID%>">
<input type=hidden name=hcountry value="">
<input type=hidden name=txtOrgCSTRCDate value="">
<input type=hidden name=hcountryValue value="">
<input type=hidden name=hcurrency value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Organization Creation</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="95">
                                    <div align="center">
                                      <center>
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Organization</td>
										</tr>
									</table>
                                      </center>
                                    </div>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="130">
								  <span style="cursor: pointer">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="OrgUnitCreationEntry.asp">
									  <td align="center">Organization Units</td></a>
									</tr>
								  </table>
								  </span>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="190">
								  <span style="cursor: pointer">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="OrgUnitDefinitionEntry.asp">
									  <td align="center">Organization Units Definition</td></a>
									</tr>
								  </table>
								  </span>
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
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td width="100%">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCell width="125"> Organization Name</td>
															<td class='FieldCell'><input type="text" name="txtOrgName" value="<%=sorgName%>" size="65" maxlength=255 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="125"> Short Name</td>
															<td class='FieldCell'><input type="text" name="txtOrgShortName" value="<%=sorgShName%>" size="12" maxlength=10 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="125"> Number of Levels</td>
															<td class='FieldCell'>
															<%	if not iorgNoUnits = "" then %>
																<input type="text" name="txtOrgNoUnits" value="<%=iorgNoUnits%>" READONLY size="3" maxlength=2 class="Formelem">
															<%	else %>
																<input type="text" name="txtOrgNoUnits" value="" size="3" maxlength=2 class="Formelem">
															<%	end if %>
															</td>
														</tr>
												</center>
													</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
										<tr>
											<td class="FieldCell" width="100%">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="68"><p align="center">Address
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
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="5"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub> Address</td>
															<td class='FieldCellSub' colspan="4"><input type="text" name="txtOrgAddress1" value="<%=sorgAddr1%>" size="55" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> </td>
															<td class='FieldCellSub' colspan="4"><input type="text" name="txtOrgAddress2" value="<%=sorgAddr2%>" size="55" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> PIN</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgPIN" value="<%=sorgPIN%>" size="16" maxlength=15 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>City</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgCity" value="<%=sorgCity%>" size="20" maxlength=25 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> State</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgState" value="<%=sorgState%>" size="30" maxlength=25 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>Country</td>
															<td class='FieldCellSub'>
															<%	if not sorgCountry = "" then %>
																<select size="1" name="selOrgCountry" class="FormElem" DISABLED>
															<%	else %>
																<select size="1" name="selOrgCountry" class="FormElem">
															<%	end if %>
																	<option value="select">Select</option>
																<%	'Calling the Function which populates the Country list
																	populateCountry
																%>
																</select>
                                                            </td>
														</tr>
												</center>
														<tr>
															<td class=FieldCellSub> Phone</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgPhone" value="<%=sorgPhone%>" size="18" maxlength=15 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>Fax</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgFax" value="<%=sorgFax%>" size="18" maxlength=15 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> E-mail ID</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgEmail" value="<%=sorgEmail%>" size="30" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>URL</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgURL" value="<%=sorgURL%>" size="28" maxlength=150 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> Contact Person</td>
															<td class='FieldCellSub' colspan="4"><input type="text" name="txtOrgContactPerson" value="<%=sorgContactPerson%>" size="30" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> Operating Currency</td>
															<td class='FieldCellSub' colspan="4">
															<%	if not norgCurrency = "" then %>
																<select size="1" name="selOrgCurrency" class="FormElem" DISABLED>
															<%	else %>
																<select size="1" name="selOrgCurrency" class="FormElem">
															<%	end if %>
																	<option value="select">Select</option>
																<%	'Calling the Function which populates the Currency list
																	populateCurrency
																%>
																</select>
															</td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
                                                </div>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
											</td>
										</tr>
										<tr>
											<td align="center" width="100%">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="128"><p align="center">Registration Details
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
															<td class=MiddlePack colspan="5"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub> TNGST RC Number</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgTNGSTRCNo" value="<%=sorgTNGSTRCNo%>" size="43" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left">Area Code</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgAreaCode" value="<%=sorgAreaCode%>" size="11" maxlength=10 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> CSTRC Number</td>
															<td class='FieldCellSub'><input type="text" name="txtOrgCSTRCNo" value="<%=sorgCSTRCNo%>" size="43" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left">CSTRC Date</td>
															<td class='FieldCellSub'>
															<%
																' Function Call to Insert Date Picker
																Response.Write InsertDatePicker("OrgCSTRCDate")
															%>
															</td>
														</tr>
												</center>
													</table>
                                                </div>
                                                            </td>
														</tr>
													</table>
                                                    </div>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Save" name="B1" class="ActionButton" onClick="javascript:checkSubmit('<%=FormatDate(now)%>')" >
															<%	if norgCurrency = "" then %>
																<input type="reset" value="Reset" name="B1" class="ActionButton">
															<%	end if %>
														</td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="BottomPack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
<%
	'Last Modified:
	'Tools Used: Microsoft Visual Interdev 6.0
	'Person Modified:
%>
