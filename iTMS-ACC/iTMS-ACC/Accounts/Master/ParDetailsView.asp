<%@ Language=VBScript %>
<%option explicit%>
<%
	'Program Name				:	ParDetailsView.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Dec 02,2010
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
Dim oDOM,MainNode,Root,sOrgName,objRs1
Dim iAgentCount,iPerfCount,iLocationCount,iContactCount,iUnitCount
Dim sGroupParentName,sGroupPartyCode,iSNo,sChildPartyName

iParty = Request.QueryString("PartyCode")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
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

	sQuery = "Select PartyName from APP_M_PartyMaster where ParentPartyCode = "& iParty
	objRs.Open sQuery,con
	if not objRs.EOF then
		do while not objRs.EOF
			sChildPartyName	= sChildPartyName & ","& objRs(0)
			objRs.MoveNext
		loop
	end if
	objRs.Close
	if sChildPartyName<>"" then
		sChildPartyName = mid(sChildPartyName,2)
	end if

end if




oDOM.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>
<XML ID="UNITDET" src="<%="../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml"%>"></XML>
<XML ID="OutData" ></XML>
<XML id="PartyData"><Root/></XML>
<XML id="TempData"><Root/></XML>
<XML id="GroupData"><Root/></XML>
<SCRIPT LANGUAGE=vbscript>
'***************************************
Function PrintFun()
	Dim  objhttp

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.open "GET","ParPrintDetailsPopulate.asp?Action=FIND",false
	objhttp.send
	if trim(objhttp.responseText)="T" then
		objhttp.open "POST","ParPrintDetailsPopulate.asp?PartyCode="&document.formname.hPartyCode.value ,false
		objhttp.send
		if trim(objhttp.responseText)<>"" then
			alert(objhttp.responseText)
		else
			'alert("../temp/master/PartyPrinting_"& document.formname.hPartyCode.value &".xml")
			window.open "../temp/master/PartyPrinting_"& document.formname.hPartyCode.value &".xml","","Status:No"
		end if
	elseif trim(objhttp.responseText)="F" then
		alert("Please Create a Print Setup File")
		PrintSetup
	else
		alert(objhttp.responseText)
	end if

End Function
'****************************************
Function PrintSetup()
	showModalDialog "ParPrintSetup.asp","","dialogWidth:600px;dialogHeight:400;Status:No"
'	document.formname.action = "ParPrintSetup.asp"
'	document.formname.submit
End Function
'*********************************************
Function ViewData()
	document.formname.action = "ParCreate_Edit_Entry.asp?PartyCode="& document.formname.hPartyCode.value
	document.formname.submit
End Function
'**************************************************
Function ControlData()
    if Trim(document.formname.hPartyCode.value)<>"" then
        document.formname.action = "PartyControlData.asp?PartyCode="& document.formname.hPartyCode.value 
        document.formname.submit 
    else
        alert("Party Controls Cannot View because Party is not available")
        exit function
    end if
End Function

'**************************************************
Function GoToMain()
document.formname.action = "ParDisplayGrid.asp"
document.formname.submit
End Function
'***********************************
'*******************************************************************
FUNCTION popPartyDet(sTemp)
Dim sUnit,iCtr,sPartyGType,sType,Temparr,iCount,iUnitrow,Pararr,iCounter,sUseable
	if trim(sTemp)<>"" then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'Msgbox sTemp
		objhttp.Open "GET","XMLGetPartyDet.asp?PartyCode=" &sTemp , false
		objhttp.send
		'alert objhttp.responseText
		iUnitrow = document.formname.hUnitrow.value
		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentElement
			'Msgbox Root.xml
			For Each PartyNode In Root.childNodes
				document.formname.txtShortName.value=PartyNode.Attributes.getNamedItem("OrgnPartyCode").value
				document.formname.txtAddress1.value=PartyNode.Attributes.getNamedItem("AddressLine1").value
				document.formname.txtAddress2.value=PartyNode.Attributes.getNamedItem("AddressLine2").value
				document.formname.txtCity.value=PartyNode.Attributes.getNamedItem("City").value
				document.formname.txtState.value=PartyNode.Attributes.getNamedItem("State").value
				document.formname.txtCountry.value=PartyNode.Attributes.getNamedItem("Country").value
				document.formname.txtPhone.value=PartyNode.Attributes.getNamedItem("PhoneNos").value
				document.formname.txtMobileNo.value=PartyNode.Attributes.getNamedItem("MobileNos").value
				document.formname.txtFax.value=PartyNode.Attributes.getNamedItem("FaxNos").value
				document.formname.txtEmail.value=PartyNode.Attributes.getNamedItem("Email").value
				document.formname.txtWebsite.value=PartyNode.Attributes.getNamedItem("WebsiteURL").value
				document.formname.txtPinCode.value=replace(PartyNode.Attributes.getNamedItem("Pincode").value," ","")
				document.formname.txtECCNo.value=PartyNode.Attributes.getNamedItem("ExciseControlCode").value
				document.formname.txtSalesLocal.value=PartyNode.Attributes.getNamedItem("LocalSTNoandDT").value
				document.formname.txtSalesCentral.value=PartyNode.Attributes.getNamedItem("CentralSTNoandDT").value
				document.formname.txtPanNo.value=PartyNode.Attributes.getNamedItem("IncomeTaxPANNo").value
				document.formname.txtPartyName.value = PartyNode.Attributes.getNamedItem("PartyName").value
				sUnit = PartyNode.Attributes.getNamedItem("Units").value
				sPartyGType = PartyNode.Attributes.getNamedItem("PartyGroupCoyType").value
				sType = PartyNode.Attributes.getNamedItem("InTrans").value
				document.formname.txtTinNo.value = PartyNode.Attributes.getNamedItem("TINNumber").value
				sUseable = PartyNode.Attributes.getNamedItem("Useable").value
				'alert(sUseable)
				if sUseable="1" then
					chkActive.innerText = "In-Active"
				end if

			next

			Dim arrName,iUnitCountSel
			'MsgBox sUnit
			Temparr = Split(sUnit,":")

		end if

	end if' if trim(sTemp)<>"" then
END FUNCTION
</SCRIPT>

<script language="javascript">
window.__itmsPopupCompat = { type: "partyDetailsView" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="popPartyDet('<%=iParty%>')">
<form method="POST" name="formname">
<input type="Hidden" name="hUnitName" value="">
<input type="Hidden" name="hUnitCode" value="" >
<input type="Hidden" name="hPartyCode" value="<%=iParty%>">
<input type="Hidden" name="hOwnUnit" value="">
<input type="Hidden" name="hAction" value="<%=sAction%>">
<input type="hidden" name="hInActive" value="0">
<input type="hidden" name="hCreatedBy" value="<%=getUserID%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hParentPartyCode" value="<%=sGroupPartyCode%>">
<input type="hidden" name="hParUnit" value="N">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%
				Response.Write "Party Details"
		%>

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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center"><a href="#" onClick="ViewData()">Details</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)" height="13">
										<tr>
											<td align="center">
												<a href="#" onClick="ControlData()">Control</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%"  class="TabCurrentTable" height="13">
										<tr>
											<td align="center">View
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
					<TD class=TabBody>
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
									<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="60"><p align="center">Basic
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

													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCellSub width="115"> Party Code - Name</td>
															<td class='FieldCell' colspan=4>
															<input type="text" name="txtShortName" size="12" maxlength="10" class="FormelemRead" Readonly > -
															<Input type="text" size="60" name="txtPartyName" value="" class="FormElemRead" ReadOnly>&nbsp;
															<span id="chkActive" class="FieldCell" >&nbsp;</span>
                                                                </td>
														</tr>
														<!--<tr>
														    <td width="115">&nbsp;</td>
															<td class="FieldCellSub" rowspan="2" valign="top">
															<input type="checkbox" name="chkGroupCompany" value="1" onClick="EnableGroup(this)" class="FormElem" disabled > Group Company</td>
															<td class="FieldCellSub">
															 <input type="radio" value="P" name="radGroupType" disabled="true" checked class="FormElem" onClick="GetGroup()" disabled > Parent </td>

															<td class="FieldCellSub"><input type="radio" value="C" disabled="true" name="radGroupType" class="FormElem" onClick="GetGroup()" disabled > Child </td>

															<td class="FieldCellSub"><input type="radio" value="B" disabled="true" name="radGroupType" class="FormElem" onClick="GetGroup()" disabled > Parent / Child </td>
															    </tr>-->
															    <%if trim(sChildPartyName)<>"" then %>
															    <tr>
																	<td class="FieldCellSub" colspan=5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<textarea class="FormElemRead" name="ParentPartyName" rows=3 cols=75 readonly >&nbsp;<%=sChildPartyName%></textarea></td>
																</tr>
																<%end if %>
													</table>
													</div>

                                                    <div align="left">
														<table cellpadding="0" cellspacing="0">
														  <tr>
														    <td class="FieldCellSub" colspan="5"><b><u>Address</u></b></td>
														  </tr>
														  <tr>
														    <td class="MiddlePack" colspan="5"><p align="left"></td>
														  </tr>
														  <tr>
														    <td class="FieldCellSub"><p align="left">Address</p>
														    </td>
														    <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress1" size="81" class="FormelemRead" Readonly></p>
														    </td>
														  </tr>
														  <tr>
														    <td class="FieldCellSub"><p align="left"></td>
														    <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtAddress2" size="81" class="FormelemRead" Readonly></p>
														    </td>
														  </tr>
														  <tr>
														    <td class="FieldCellSub"><p align="left">City</p>
														    </td>
														    <td class="FieldCellSub" colspan="4"><p align="left"><input type="text" name="txtCity" size="25" class="FormelemRead" Readonly></p>
														    </td>
														  </tr>
														  <tr>
														    <td class="FieldCellSub"><p align="left">PIN</p>
														    </td>
														    <td class="FieldCellSub"><p align="left"><input type="text" name="txtPinCode" size="7"  maxlength="6" class="FormelemRead" Readonly></p>
														    </td>
														    <td class="FieldCellSub">
														    </td>
														  <td class="FieldCellSub"><p align="left">Phone</p>
														  </td>
														  <td class="FieldCellSub"><p align="left"><input type="text" name="txtPhone" size="18" class="FormelemRead" Readonly></p>
														  </td>
														  </tr>
														  <tr>
														    <td class="FieldCellSub"><p align="left">State</p>
														    </td>
														    <td class="FieldCellSub"><p align="left"><input type="text" name="txtState" size="35" class="FormelemRead" Readonly ></p>
														    </td>
														    <td class="FieldCellSub">
														    </td>
														  <td class="FieldCellSub"><p align="left">Fax</p>
														  </td>
														  <td class="FieldCellSub"><p align="left"><input type="text" name="txtFax" size="18" class="FormelemRead" Readonly></p>
														  </td>
														  </tr>
														  <tr>
														    <td class="FieldCellSub"><p align="left">Country</p>
														    </td>
														    <td class="FieldCellSub"><p align="left"><input type="text" name="txtCountry" size="25" class="FormelemRead" Readonly></p>
														    </td>
														    <td class="FieldCellSub">
														    </td>
														  <td class="FieldCellSub">Mobile
														  </td>
														  <td class="FieldCellSub"><input type="text" name="txtMobileNo" size="18" class="FormelemRead" Readonly>
														  </td>
														  </tr>
														  <tr>
														  <td class="FieldCellSub"><p align="left">E-mail ID</p>
														  </td>
														  <td class="FieldCellSub"><p align="left"><input type="text" name="txtEmail" size="35" class="FormelemRead" Readonly></p>
														  </td>
														  <td class="FieldCellSub">
														  </td>
														  <td class="FieldCellSub"><p align="left">URL</p>
														  </td>
														  <td class="FieldCellSub"><p align="left"><input type="text" name="txtWebsite" size="25" class="FormelemRead" Readonly></p>
														  </td>
														  </tr>
														<tr>
														  <td class="MiddlePack" colspan="5"><p align="left"></td>
														</tr>
														</table>
                                                    	<table border="0" cellspacing="0" cellpadding="0">
														   	<tr>
														   	  <td class="FieldCellSub" colspan="4"><b><u>Others</u></b></td>
														   	</tr>
														   <tr>
														   		<td class="FieldCellSub" width="165">Excise ECC Number</td>
														   		<td class="FieldCellSub"> <input type="text" name="txtECCNo" size="17" class="FormelemRead" Readonly> </td>
														   		<td class="FieldCellSub" >IT PAN No</td>
														   		<td class="FieldCellSub"> <input type="text" name="txtPanNo" size="15" class="FormelemRead" Readonly> </td>
														   	</tr>
														   	<tr>
														   		<td class="FieldCellSub" >Sales Tax Number - Local</td>
														   		<td class="FieldCellSub"> <input type="text" name="txtSalesLocal" size="17" class="FormelemRead" Readonly> </td>
														   		<td class="FieldCellSub" >TIN Number</td>
														   		<td class="FieldCellSub"> <input type="text" name="txtTinNo" size="15" class="FormelemRead" Readonly> </td>
														   	</tr>
														   	<tr>
														   		<td class="FieldCellSub" >Sales Tax Number - Central</td>
														   		<td class="FieldCellSub"> <input type="text" name="txtSalesCentral" size="17" class="FormelemRead" Readonly> </td>
														   		<td class="FieldCellSub" >&nbsp;</td>
														   		<td class="FieldCellSub">

														   		</td>
														   	</tr>

														</table>
														</div>
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
								<td valign="top" width="100%">
									<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="80"><p align="center">Other Details
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
                                                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
															<tr>
																<td class="FieldCellSub"><b><u>Transaction Units<u></b>
															</tr>
															<tr>
																<td class="FieldCellSub">
																	<table class="ExcelTable" border=0 cellspacing=1 cellpadding=0 width="100%">
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
																	<%
																		Do while Not objRs.EOF
																			iUnitCount = iUnitCount + 1

																	%>
																		<tr><td class="ExcelHeaderCell" colspan=4>
																		<%=Trim(objRs(1))%>
																		</td>
																		<td class="ExcelHeaderCell" align="center">Credit Limit</td>
																		<td class="ExcelHeaderCell" align="center">Credit Days</td>
																		</tr>
																	<%
																	      sQuery ="Select isNull(OpeningBalance,0),OpeningCDIndication,SubTypeName,isNull(CreditLimit,0),isNull(CreditDays,0) from APP_R_OrgParty OP,APP_M_PartyTypes PT where PT.PartyType =OP.PartyType and PT.PartySubType = OP.PartySubType and OP.PartyCode = "& iParty &" and OP.OUDefinitionID = '" & objRs(0) &"'"
																	      objRs1.Open sQuery,con
																	      if not objRs1.EOF then
																			iSNo= 0
																			do while not objRs1.EOF
																				iSNo= iSNo + 1
																			%>
																				<tr><td class="ExcelDisplayCell"><%=iSNo%>
																					<td class="ExcelDisplayCell"><%=Trim(objRs1(0))%></td>
																					<td class="ExcelDisplayCell">
																					<%
																					 	if objRs1(1)="C" then
																							 Response.Write "CR"
																						else
																							Response.Write "DR"
																						end if
																					%>
																					</td>
																					<td class="ExcelDisplayCell"><%=objRs1(2)%></td>
																					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(objRs1(3),2)%></td>
																					<td class="ExcelDisplayCell" align="right"><%=objRs1(4)%></td>
																					</tr>
																			<%
																				objRs1.MoveNext
																			loop
																	      end if
																	      objRs1.Close

																		objRs.MoveNext
																		Loop
																		objRs.Close
																	%>
																	</table>
																</td>
															</tr>
														</table>
													</div>

                                                    <div align="left">
                                                    <table border="0" cellspacing="0" cellpadding="0">
                                                    <tr>
														<td class="FieldCellSub" colspan=3 ><b><u>Contacts<u></b>
													</tr>
													<tr>
														<td class=MiddlePack colspan=3></td>
													</tr>
                                                	<tr>
                                                		<td width=5>&nbsp;</td>
                                                		<td width="100%">
                                                			<table border=0 cellspacing="1" cellpadding="0" class="ExcelTable" width=100% >
                                                				<tr>
                                                					<td class="ExcelHeaderCell" align=center width=10 >S.No.</td>
                                                					<td class="ExcelHeaderCell" align=center>Person Name</td>
                                                					<td class="ExcelHeaderCell" align=center>Designation</td>
                                                					<td class="ExcelHeaderCell" align=center>Person For</td>
                                                					<td class="ExcelHeaderCell" align=center>Mail ID</td>
                                                				</tr>
                                                				<%
                                                					sQuery = "Select ContactPersonName,Designation,ContactPersonFor,ContactMailID from APP_M_PartyContactPersons where PartyCode = "& iParty
                                                					objRs.Open sQuery,con
                                                					if not objRs.EOF then
                                                						iSNo = 0
                                                						do while not objRs.EOF
                                                							iSNo = iSNo + 1
                                                							%>
                                                								<tr>
                                                									<td class="ExcelSerial" align=center ><%=iSNo%></td>
                                                									<td class="ExcelDisplayCell"><%=objRs(0)%></td>
                                                									<td class="ExcelDisplayCell"><%=objRs(1)%></td>
                                                									<td class="ExcelDisplayCell"><%=objRs(2)%></td>
                                                									<td class="ExcelDisplayCell"><%=objRs(3)%></td>
                                                								</tr>
                                                							<%
                                                							objRs.MoveNext
                                                						loop
                                                					end if
                                                					objRs.Close
                                                				%>
                                                			</table>
                                                		</td>
                                                		<td width=5>&nbsp;</td>
                                                	</tr>
                                                	<tr>
														<td class=MiddlePack colspan=3></td>
													</tr>
                                                </table>
												</div>
												<div align="left">
                                                    <table border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td class="FieldCellSub" colspan=3 ><b><u>Locations<u></b>
													</tr>
													 <tr>
													 <td class=MiddlePack colspan=3></td>
														</tr>
                                                				<tr>
                                                					<td width=5>&nbsp;</td>
                                                					<td width="100%">
                                                						<table border=0 cellspacing="1" cellpadding="0" class="ExcelTable" width=100% >
                                                							<tr>
                                                								<td class="ExcelHeaderCell" align=center width=10 >S.No.</td>
                                                								<td class="ExcelHeaderCell" align=center>Location Name</td>
                                                								<td class="ExcelHeaderCell" align=center>Address</td>
                                                								<td class="ExcelHeaderCell" align=center></td>
                                                							</tr>
                                                							<%
                                                								sQuery = "Select Location,isNull(LocationAddress1,''),isNull(LocationAddress2,''),isNull(City,''),isNull(State,''),isNull(Country,''),isNull(LocalSTNoandDT,''),isNull(CentralSTNoandDT,''),isNull(ExciseControlNo,''),isNull(IncometaxPanNo,'') from APP_M_PartyLocations where PartyCode = "& iParty
                                                								objRs.Open sQuery,con
                                                								if not objRs.EOF then
                                                									iSNo = 0
                                                									do while not objRs.EOF
                                                										iSNo = iSNo + 1
                                                										%>
                                                											<tr>
                                                												<td class="ExcelSerial" align=center ><%=iSNo%></td>
                                                												<td class="ExcelDisplayCell"><%=objRs(0)%></td>
                                                												<td class="ExcelDisplayCell"><%=objRs(1)%><br><%=objRs(2)%><br><%=objrs(3)%><br><%=objrs(4)%><br><%=objrs(5)%></td>
                                                												<td class="ExcelDisplayCell">ECCNo: <%=objRs(8)%><br>Local Sale Tax: <%=objRs(6)%><br>Central Sales Tax : <%=objRs(7)%><br>IT PANNo :<%=objRs(9)%></td>
                                                											</tr>
                                                										<%
                                                										objRs.MoveNext
                                                									loop
                                                								end if
                                                								objRs.Close
                                                							%>
                                                						</table>
                                                					</td>
                                                					<td width=5>&nbsp;</td>
                                                				</tr>
                                                				<tr>
																	<td class=MiddlePack colspan=3></td>
																</tr>

																</table>
													</div>
												    <div align="left">
                                                    <table border="0" cellspacing="0" cellpadding="0">
                                                    <tr>
														<td class="FieldCellSub" colspan=3 ><b><u>Agent<u></b>
													</tr>
                                                    <tr>
													<td class=MiddlePack colspan=3></td>
													</tr>
                                                	<tr>
                                                		<td width=5>&nbsp;</td>
                                                		<td width="100%">
                                                			<table border=0 cellspacing="1" cellpadding="0" class="ExcelTable" width=100% >
                                                				<tr>
                                                					<td class="ExcelHeaderCell" align=center width=10 >S.No.</td>
                                                					<td class="ExcelHeaderCell" align=center>Agent Name - Code</td>
                                                					<td class="ExcelHeaderCell" align=center>Address</td>
                                                				</tr>
                                                				<%
                                                					sQuery = "Select OrgnPartyCode,PartyName,isNull(AddressLine1,''),isNull(AddressLine2,''),isNull(City,''),isNull(State,''),isNull(Country,'') from APP_M_PartyMaster where PartyCode in (Select AgentCode from APP_R_AgentOrgParty  where PartyCode = " & iParty & ")"
                                                					objRs.Open sQuery,con
                                                					if not objRs.EOF then
                                                						iSNo = 0
                                                						do while not objRs.EOF
                                                							iSNo = iSNo + 1
                                                							%>
                                                								<tr>
                                                									<td class="ExcelSerial" align=center ><%=iSNo%></td>
                                                									<td class="ExcelDisplayCell"><%=objRs(1)%>-<%=objRs(0)%></td>
                                                									<td class="ExcelDisplayCell"><%=objRs(2)%><br><%=objRs(3)%><br><%=objRs(4)%><br><%=objRs(5)%><br><%=objRs(6)%></td>
                                                								</tr>
                                                							<%
                                                							objRs.MoveNext
                                                						loop
                                                					end if
                                                					objRs.Close
                                                				%>
                                                			</table>
                                                		</td>
                                                		<td width=5>&nbsp;</td>
                                                	</tr>
                                                	<tr>
														<td class=MiddlePack colspan=3></td>
													</tr>
                                                    </table>
													</div>
													<div align="left">
                                                    <table border="0" cellspacing="0" cellpadding="0">
                                                    <tr>
														<td class="FieldCellSub" colspan=3 ><b><u>Preferences<u></b>
													</tr>
													<tr>
														<td class=MiddlePack colspan=3></td>
													</tr>
                                                	<tr>
                                                		<td width=5>&nbsp;</td>
                                                		<td width="100%">
                                                			<table border=0 cellspacing="1" cellpadding="0" width=100% >
																			<tr>
																				<td class="FieldCell">Payment Terms</td>
																				<td class="FieldCellSub"><span width="50" class="DataOnly" >
																			<%	sQuery = "SELECT PaymentTermsDesc, PymtTermsShortDesc FROM APP_M_PaymentTermsHeader where PaymentTermsNo in (Select PrefPaymentTerms from APP_R_OrgParty where PartyCode = "& iParty &")"
																				objRs.Open sQuery,con
																				if not objRs.EOF then
																					Response.Write objRs(0)
																				end if
																				objRs.Close
																			%>  </span></td>
																				<td class="FieldCell">Payment Mode</td>
																				<td class="FieldCellSub"><span width="50" class="DataOnly">
																			<%	sQuery = "SELECT PaymentMode, ShortPaymentMode FROM APP_M_ModeOfPayment where PaymentModeNo in (Select PrefPaymentMode from APP_R_OrgParty where PartyCode = "& iParty &")"
																				objRs.Open sQuery,con
																				if not objRs.EOF then
																					Response.Write objRs(0)
																				end if
																				objRs.Close
																			%>  </span></td>
																			</tr>
																			<tr>
																				<td class="FieldCell">Basis of Pricing</td>
																				<td class="FieldCellSub"><span width="50" class="DataOnly">
																			<%	sQuery = "SELECT BasisOfPricing, ShortBasisofPricing FROM APP_M_BasisOfPricing where BasisOfPricingNo in (Select PrefbasisofPricing from APP_R_OrgParty where PartyCode = "& iParty &")"
																				objRs.Open sQuery,con
																				if not objRs.EOF then
																					Response.Write objRs(0)
																				end if
																				objRs.Close
																			%>  </span></td>
																				<td class="FieldCell">Transporter</td>
																				<td class="FieldCellSub"><span width="50" class="DataOnly">
																			<%	sQuery = "SELECT TransporterName, TransportShortName FROM APP_M_Transporter where TransporterCode in (Select PrefTransporterCode from APP_R_OrgParty where PartyCode = "& iParty  &")"
																				objRs.Open sQuery,con
																				if not objRs.EOF then
																					Response.Write objRs(0)
																				end if
																				objRs.Close
																			%>  </span></td>
																			</tr>
																			<tr>
																				<td class="FieldCell">Despatch Mode</td>
																				<td class="FieldCellSub"><span width="50" class="DataOnly">
																			<%	sQuery = "SELECT DespatchModeDesc, ShortDespatchMode FROM APP_M_ModeOfDespatch where DespatchModeNo in (Select PrefDespatchMode from APP_R_OrgParty where PartyCode = "& iParty &")"
																				objRs.Open sQuery,con
																				if not objRs.EOF then
																					Response.Write objRs(0)
																				end if
																				objRs.Close
																			%>  </span></td>
																				<td class="FieldCell">Currency</td>
																				<td class="FieldCellSub"><span width="50" class="DataOnly">
																			<%	sQuery = "SELECT  CurrencyName, CurrencyShortName FROM Ms_CurrencyMaster where CurrencyCode in (Select PrefCurrencyCode from APP_R_OrgParty where PartyCode = "& iParty &")"
																				objRs.Open sQuery,con
																				if not objRs.EOF then
																					Response.Write objRs(0)
																				end if
																				objRs.Close
																			%>  </span></td>
																			</tr>
																		</table>
																	</td>
																	<td width=5>&nbsp;</td>
																</tr>
																<tr>
																	<td class=MiddlePack colspan=3></td>
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
																<input type="button" value="Print Setup" name="btnPrintSetup" class="ActionButtonX"  onClick="PrintSetup()">
																<input type="button" value="Print" name="btnPrint" class="ActionButtonX"  onClick="PrintFun()">
																<input type="button" value="Close" name="btnClose" class="ActionButton"  onClick="GoToMain()">
                                                               <!-- <input type="button" value="Save" name="B2" class="ActionButton" onClick="PageSubmit()">

                                                                <input type="button" value="Preview" name="btnPreveiw" class="ActionButton"  onClick="ViewData()">
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton"  >-->
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
