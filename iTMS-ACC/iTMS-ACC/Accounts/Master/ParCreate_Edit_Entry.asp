<%@ Language=VBScript %>
<%	option explicit
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ParCreate_Edit_Entry.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 15,2010
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
dim sQuery,objRs,iParty,sCallTy,Temparr,Unitarr,sAction
Dim oDOM,MainNode,Root,sOrgName
Dim iAgentCount,iPerfCount,iLocationCount,iContactCount,iUnitCount,iRepCount
Dim sGroupParentName,sGroupPartyCode,sOrgCode

iParty = Request.QueryString("PartyCode")

sOrgCode = Session("organizationcode")
Set objRs = Server.CreateObject("ADODB.RecordSet")
sCallTy = Request("hCallTy")
iUnitCount = 0
sGroupPartyCode = 0
if trim(iParty)="" then
	sAction = "CREATE"
else
	sAction = "EDIT"
end if

set oDOM = Server.CreateObject("Microsoft.XMLDOM")

set Root = oDOM.createElement("Root")
oDOM.appendChild Root

sQuery = "SELECT OUDefinitionID, OrganizationUnitId, OrgUnitDescription, OrgUnitShortDescription, "&_
		 "isNull(Address1,''), isNull(Address2,''), isNull(PostCode,''), isNull(City,''), isNull(State,''), isNull(Country,0), isNull(PhoneNumber,''),isNull(FaxNumber,''), isNull(EmailID,''), "&_
		 "isNull(WeSiteURL,'') FROM  DCS_OrganizationUnitDefinitions "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Do While Not objRs.EOF
	Set MainNode = oDom.createElement("UNIT")
	MainNode.setAttribute "UnitID", objRs(0)
	MainNode.setAttribute "ID", objRs(1)
	MainNode.setAttribute "Desc", objRs(2)
	MainNode.setAttribute "ShortDesc", objRs(3)
	MainNode.setAttribute "Add1", objRs(4)
	MainNode.setAttribute "Add2", objRs(5)
	MainNode.setAttribute "PostCode", objRs(6)
	MainNode.setAttribute "City", objRs(7)
	MainNode.setAttribute "State", objRs(8)
	MainNode.setAttribute "Country", objRs(9)
	MainNode.setAttribute "Phone", objRs(10)
	MainNode.setAttribute "Fax", objRs(11)
	MainNode.setAttribute "EmailID", objRs(12)
	MainNode.setAttribute "Web", objRs(13)
	Root.appendChild MainNode
	objRs.MoveNext
loop
objRs.Close



sQuery = "Select OrganizationName From DCS_Organization "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sOrgName = objRs(0)
End IF
objRs.Close

if sAction="EDIT" then

sQuery = "Select PartyCode,PartyName from APP_M_PartyMaster where PartyCode "&_
		 "in (Select ParentPartyCode from APP_M_PartyMaster where"&_
		 " PartyCode = "& iParty &" and PartyCode<>0)"
'		 Response.Write sQuery
	objRs.Open sQuery,con
	if not objRs.EOF then
		sGroupPartyCode = objRs(0)
		sGroupParentName = trim(objRs(1))
	end if
	objRs.Close
end if


oDOM.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/Cancel.js"></SCRIPT>

<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" ID="UNITDET" data-src="<%="../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" ID="OutData" ></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="GroupData"><Root/></script>
<SCRIPT>
<!--
function EnableGroup(objGroup)
{
	if (document.formname.chkGroupCompany.checked==true)
	{

		document.formname.radGroupType[0].disabled=false;
		document.formname.radGroupType[1].disabled=false;
		document.formname.radGroupType[2].disabled=false;
	}
	else
	{
		document.formname.radGroupType[0].disabled=true;
		document.formname.radGroupType[1].disabled=true;
		document.formname.radGroupType[2].disabled=true;
	}
}
function CheckSubmit()
{

	var i,bFalg;

	bFlag=true;
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
	return true;
}
//-->
</SCRIPT>

<script>
window.__itmsPopupCompat = { type: "partyCreateEditModals" };
</script>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="popPartyDet('<%=iParty%>')">
<form method="POST" name="formname">
<input type="Hidden" name="hUnitName" value="">
<input type="Hidden" name="hUnitCode" value="<%=sOrgCode%>">
<input type="Hidden" name="hPartyCode" value="<%=iParty%>">
<input type="Hidden" name="hOwnUnit" value="">
<input type="Hidden" name="hAction" value="<%=sAction%>">
<input type="hidden" name="hInActive" value="0">
<input type="hidden" name="hCreatedBy" value="<%=getUserID%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hParentPartyCode" value="<%=sGroupPartyCode%>">
<input type="hidden" name="hParUnit" value="N">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
		<%
			if trim(iParty)<>"" then
				Response.Write "Party Amendment"
			else
				Response.Write "Party Creation"
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
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
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
								<td class="TabCell" valign="bottom" align="center" width="60px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)" height="13px">
										<tr>
											<td align="center">
												<a href="#" onClick="ControlData(); return false;">Control</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)" height="13px">
										<tr>
											<td align="center">
												<a href="#" onClick="ViewData(); return false;">View</a>
											</td>
										</tr>
									</table>
								</td>
								<!--<td class="TabCell" valign="bottom" align="center" width="60">
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
								</td>-->
								<td class="TabCellEnd" valign="bottom" align="left">
                                &nbsp;
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
															<td class=FieldCell width="115px"> Party Name</td>
															<td class="FieldCell">
															<Input type="text" size="71" name="txtPartyName" value="" class="FormElem">&nbsp;
															<input type="checkbox" name="chkActive" value="1" class="FormElem">&nbsp;In-Active
                                                                </td>
														</tr>
														<tr>
															<td class=FieldCell width="115px"> Party Code</td>
															<td class='FieldCell' valign=top><input type="text" name="txtShortName" size="12" maxlength="10" class="Formelem">
															<input type=CheckBox name=ChkOwnUnit class=Formelem onClick=GetUnit()>Own Unit
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
									<table border="0" cellspacing="0" cellpadding="0">
									                                                        <tr>
									                                                        	<td width="115px">&nbsp;</td>
									                                                    <td class="FieldCell" rowspan="3" valign="top">
									                                                    <input type="checkbox" name="chkGroupCompany" value="1" onClick="EnableGroup(this)" class="FormElem"> Group Company</td>
									                                                    <td class="FieldCellSub">
									                                                     <input type="radio" value="P" name="radGroupType" disabled="true" checked class="FormElem" onClick="GetGroup()"> Parent </td>

									                                                    <td class="FieldCellSub"><input type="radio" value="C" disabled="true" name="radGroupType" class="FormElem" onClick="GetGroup()" > Child </td>

									                                                    <td class="FieldCellSub"><input type="radio" value="B" disabled="true" name="radGroupType" class="FormElem" onClick="GetGroup()" > Parent / Child </td>
									                                                        </tr>
									                                                        <tr>
																								<td class="FieldCell" colspan=5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" class="FormElemRead" name="ParentPartyName" size=75 readonly value="<%=sGroupParentName%>" >&nbsp;</span></td>
									                                                        </tr>
                                                            </table>
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
                                            <td class="FieldCellSub">Address
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtAddress1" size="81" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"></td>
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
								</td>
								<td valign="top">
                                    <table border="0" cellspacing="0" class="BodyTable" cellpadding="0" width="100%">
                                    	<tr>
                                           <td>
                                               <table border="0" cellpadding="0" cellspacing="0">
                                               <tr>
                                            		<td>
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                        		<tr>
																	<td class="FieldCellSub" width="165">Excise ECC Number</td>
																	<td class="FieldCellSub"> <input type="text" name="txtECCNo" size="17" class="FormElem"> </td>
																	<td class="FieldCellSub" >IT PAN No</td>
																	<td class="FieldCellSub"> <input type="text" name="txtPanNo" size="15" class="FormElem"> </td>
																</tr>
																<tr>
																	<td class="FieldCellSub" >Sales Tax Number - Local</td>
																	<td class="FieldCellSub"> <input type="text" name="txtSalesLocal" size="17" class="FormElem"> </td>
																	<td class="FieldCellSub" >TIN Number</td>
																	<td class="FieldCellSub"> <input type="text" name="txtTinNo" size="15" class="FormElem"> </td>
																</tr>
																<tr>
																	<td class="FieldCellSub" >Sales Tax Number - Central</td>
																	<td class="FieldCellSub"> <input type="text" name="txtSalesCentral" size="17" class="FormElem"> </td>
																	<td class="FieldCellSub" >&nbsp;</td>
																	<td class="FieldCellSub">

																	</td>
																</tr>

                                                            </table>
                                            		</td>
                                            		<!--td class="ClearPixel">
														<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                            		</td>
													<td valign="top" align="right">
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
													</td-->
                                                </tr>
                                              	</table>
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
									<table border="0" cellspacing="0" class="BodyTable" cellpadding="0" width="100%">
										<tr>
											<td class="ExcelHeaderCell" height="20px" >Transaction Units  (<b><a href="#" onClick="PopulatePartyTypes('<%=iParty%>','<%=sAction%>'); return false;" >Party Types</a></b>) </td>
											<td class="ExcelHeaderCell" >Contacts</td>
											<td class="ExcelHeaderCell" >Locations</td>
											<td class="ExcelHeaderCell" >Preference</td>
											<td class="ExcelHeaderCell" >Agent</td>
											<td class="ExcelHeaderCell" >Rep.</td>
										</tr>
										<tr>
											<td class="FieldCellSub">
												<%
													sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID "

													'Response.Write sQuery
													With objRs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sQuery
														.ActiveConnection = Con
														.Open
													End With
													Set objRs.ActiveConnection = Nothing
												%>
												<input type="hidden" name="hUnitrow" value="<%=objRs.RecordCount%>">
												<input type="checkbox" class="FormElem" name="chkUnitZ0" value="0:All">All Units&nbsp;
													<br>
												<%
													Do while Not objRs.EOF
														iUnitCount = iUnitCount + 1
												%>
													<input type="checkbox" class="FormElem" name="chkUnitZ<%=iUnitCount%>" value="<%=objRs(0)%>:<%=Trim(objRs(1))%>"><%=Trim(objRs(1))%> &nbsp;
													<br>
												<%
													objRs.MoveNext
													Loop
													objRs.Close
												%>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Contact('<%=iParty%>'); return false;" >
											<%
												if iParty<>"" then
													sQuery = "Select count(PartyCode) from APP_M_PartyContactPersons where PartyCode = "& iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														iContactCount = objRs(0)
													end if
													objRs.Close
												else
													iContactCount = 0
												end if 'if iParty<>"" then

												if iContactCount=0 then
													Response.Write "0 Add"
												else
													Response.Write iContactCount &" Manage"
												end if

											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Location('<%=iParty%>'); return false;">
											<%
												if iParty<>"" then
													sQuery= "Select Count(PartyCode) from APP_M_PartyLocations where PartyCode = "&iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														iLocationCount = objRs(0)
													end if
													objRs.Close
												else
													iLocationCount = 0
												end if 'if iParty<>"" then

												if iLocationCount=0 then
													Response.Write "0 Add"
												else
													Response.Write iLocationCount&" Manage"
												end if
											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Preference('<%=iParty%>'); return false;">
											<%
												if iParty<>"" then
													sQuery = "Select Count(partyCode) from APP_R_OrgParty where (isNull(PrefTransporterCode,0)<>0 OR "&_
															 "isNull(PrefDespatchMode,0)<>0 or isNull(PrefCurrencyCode,0)<>0 or "&_
															 "isNull(PrefPaymentMode,0)<>0 Or isNull(PrefBasisOfPricing,0)<>0 Or "&_
															 "isNull(PrefPaymentTerms,0)<>0) and PartyCode = "& iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														 iPerfCount = objRs(0)
													end if
													objRs.Close
												end if ' if iParty<>"" then

												if iPerfCount = 0 then
													Response.Write "0 Add"
												else
													Response.Write "1 Manage"
												end if
											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Agent('<%=iParty%>'); return false;">
											<%
												if iParty<>"" then
													sQuery = "Select Count(AgentCode) from APP_R_AgentOrgParty where PartyCode = "&iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														iAgentCount = objRs(0)
													end if
													objRs.Close
												else
													iAgentCount = 0
												end if 'if iParty<>"" then
												if iAgentCount=0 then
													Response.Write iAgentCount & "  Add"
												else
														Response.Write iAgentCount  &"  Manage"
												end if ' if iAgentCount=0 then
											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Rep('<%=iParty%>'); return false;">
											<%
												if iParty<>"" then
													sQuery = "Select isNull(RepAreaCode,0),isNull(RepAgentEntryID,0) from APP_R_OrgParty where PartyCode = "& iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
													    if Trim(objRs(0))<>"0" and Trim(objRs(1))<>"0" then
													       iRepCount  = 1
													    end if 'if Trim(objRs(0))<>"0" and Trim(objRs(1))<>"0" then
													end if
													objRs.Close
												else
													iRepCount = 0
												end if 'if iParty<>"" then
												if iRepCount=0 then
													Response.Write "Add"
												else
													Response.Write "Manage"
												end if ' if iAgentCount=0 then
											%>
											</a>
											</td>
										</td>
									</table>
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
                                                                <input type="button" value="Save" name="B2" class="ActionButton" onClick="PageSubmit()">
                                                                <input type="button" value="Close" name="B3" class="ActionButton"  onClick="GoToMain()">
                                                                <input type="button" value="Preview" name="btnPreveiw" class="ActionButton"  onClick="ViewData()">
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton"  >
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
<%
set objRs=nothing
%>
