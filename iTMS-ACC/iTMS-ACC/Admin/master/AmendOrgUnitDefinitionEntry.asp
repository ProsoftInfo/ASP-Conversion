<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AmendOrgUnitDefinitionEntry.asp
	'Module Name				:	Inventory (Organization Creation Amendment)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 29, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	orgUnitDefinitionInsert.asp
	'Procedures/Functions Used	:	populateOrganization,populateOrganizationUnit and populateCountry
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Organization Unit Definition</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML ID="orgData" src="../xmldata/Organization.xml"></XML>
<XML ID="DivisionData" src="../xmldata/Division.xml"></XML>
<XML ID="UnitData" src="../xmldata/Unit.xml"></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/AmendOrgUnitDefinition.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
	Function openDetails()
		window.showModalDialog "XMLUnitView.asp","Unit","dialogHeight:310px;dialogWidth:320px;center:Yes;help:No;resizable:No;status:No"
	end function

	Function LoadUnits()
		dim Root, Node ,i ,j
		
		i = document.formname.selOrg.selectedIndex 
		j = document.formname.selOrgUnit.selectedIndex 
		document.formname.reset
		document.formname.selOrg.selectedIndex = i
		document.formname.selOrgUnit.selectedIndex = j
		
		i = 0	
				
		if document.formname.selOrgUnit.selectedIndex = "1" then
			set Root=DivisionData.documentElement

			document.formname.selDivisionUnit.length = 1
			for Each Node in Root.childNodes
				i=i+1
				document.formname.selDivisionUnit.length = document.formname.selDivisionUnit.length + 1
				document.formname.selDivisionUnit.options(i).value = Node.attributes.getNamedItem("OUDEFINITIONID").value
				document.formname.selDivisionUnit.options(i).text = Node.attributes.getNamedItem("ORGUNITDESCRIPTION").value	
			next
		elseif document.formname.selOrgUnit.selectedIndex = "2" then
			set Root=UnitData.documentElement

			document.formname.selDivisionUnit.length = 1
			for Each Node in Root.childNodes
				i=i+1
				document.formname.selDivisionUnit.length = document.formname.selDivisionUnit.length + 1
				document.formname.selDivisionUnit.options(i).value = Node.attributes.getNamedItem("OUDEFINITIONID").value
				document.formname.selDivisionUnit.options(i).text = Node.attributes.getNamedItem("ORGUNITDESCRIPTION").value	
			next
		end if
	End Function
	
	Function LoadDetails()
		dim sTemp, Root, HeaderNode, Node
		
		if document.formname.selOrgUnit.selectedIndex = "1" then
			set Root=DivisionData.documentElement
			sTemp = "//Division [@OUDEFINITIONID='" & document.formname.selDivisionUnit.value & "']"
		elseif document.formname.selOrgUnit.selectedIndex = "2" then
			set Root=UnitData.documentElement
			sTemp = "//Unit [@OUDEFINITIONID='" & document.formname.selDivisionUnit.value & "']"
		end if

		set Node=Root.SelectNodes(sTemp)
		
		if Node.length > 0 then

			document.formname.txtUnitName.value = Node.item(0).attributes.getNamedItem("ORGUNITDESCRIPTION").value
			document.formname.txtUnitShName.value = Node.item(0).attributes.getNamedItem("ORGUNITSHORTDESCRIPTION").value
			document.formname.txtUnitAddr1.value = Node.item(0).attributes.getNamedItem("ADDRESS1").value
			document.formname.txtUnitAddr2.value = Node.item(0).attributes.getNamedItem("ADDRESS2").value
			document.formname.txtUnitPIN.value = Node.item(0).attributes.getNamedItem("POSTCODE").value
			document.formname.txtUnitCity.value = Node.item(0).attributes.getNamedItem("CITY").value
			document.formname.txtUnitState.value = Node.item(0).attributes.getNamedItem("STATE").value
			setIndex  document.formname.selUnitCountry ,  Node.item(0).attributes.getNamedItem("COUNTRYCODE").value
			document.formname.txtUnitPhone.value = Node.item(0).attributes.getNamedItem("PHONENUMBER").value
			document.formname.txtUnitFax.value = Node.item(0).attributes.getNamedItem("FAXNUMBER").value
			document.formname.txtUnitEmail.value = Node.item(0).attributes.getNamedItem("EMAILID").value
			document.formname.txtUnitURL.value = Node.item(0).attributes.getNamedItem("WESITEURL").value
			document.formname.txtUnitContactPerson.value = Node.item(0).attributes.getNamedItem("CONTACTPERSON").value
			document.formname.txtUnitTNGSTNo.value = Node.item(0).attributes.getNamedItem("TNGSTRCNUMBER").value
			document.formname.txtUnitCSTRCNo.value = Node.item(0).attributes.getNamedItem("CSTRCNUMBER").value
			document.formname.UnitCSTRCDate.setDate = Node.item(0).attributes.getNamedItem("CSTRCDATE").value
			document.formname.txtUnitAreaCode.value = Node.item(0).attributes.getNamedItem("AREACODE").value
			document.formname.txtUnitRange.value = Node.item(0).attributes.getNamedItem("RANGE").value
			document.formname.txtUnitDivision.value = Node.item(0).attributes.getNamedItem("DIVISION").value
			document.formname.txtUnitCollectorate.value = Node.item(0).attributes.getNamedItem("COLLECTORATE").value
			document.formname.txtUnitCentralENo.value = Node.item(0).attributes.getNamedItem("CENTRALEXCISECODE").value
			document.formname.txtUnitRegNo.value = Node.item(0).attributes.getNamedItem("REGISTRATIONNUMBER").value
			document.formname.txtUnitLANo.value = Node.item(0).attributes.getNamedItem("LANUMBER").value
			document.formname.txtRangeAdd1.value = Node.item(0).attributes.getNamedItem("RANGEADDRESS1").value
			document.formname.txtRangeAdd2.value = Node.item(0).attributes.getNamedItem("RANGEADDRESS2").value
			document.formname.txtRangeAdd3.value = Node.item(0).attributes.getNamedItem("RANGEADDRESS3").value
			document.formname.txtDivisionAdd1.value = Node.item(0).attributes.getNamedItem("DIVISIONADDRESS1").value
			document.formname.txtDivisionAdd2.value = Node.item(0).attributes.getNamedItem("DIVISIONADDRESS2").value
			document.formname.txtDivisionAdd3.value = Node.item(0).attributes.getNamedItem("DIVISIONADDRESS3").value

		end if
	End Function

	Function setIndex(obj,sTemp)
		dim i
		for i = 0 to obj.length - 1
			if trim(sTemp) = trim(obj.options(i).value) then
				obj.selectedIndex = i
				exit function
			end if
		next
	End Function
	
	Function sethiddenDate()
		document.formname.txtUnitCSTRCDate.value = document.formname.UnitCSTRCDate.GetDate
	end Function
</SCRIPT>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name=hcountry value="">
<input type=hidden name=txtUnitCSTRCDate value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Organization Amendment</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" bordercolor="#000000">
				<TR>
					<TD class=TabBodywithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellspacing="0"  cellpadding="0" class="ToolBarTable">
										<tr>
											<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >

				       							<span style="cursor: hand" Title="Exisiting Units" onclick="openDetails()">
              									<p align="center"><font face="Wingdings" size="5">4</font>
												</span>
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
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
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td width="100%">
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class=MiddlePack colspan="2"> </td>
																</tr>
																<tr>
																	<td class=FieldCell width="151"> Organization</td>
																	<td class='FieldCell'>
																	<select size="1" name="selOrg" class="FormElem">
																	<%	'Calling the Function which populates the Organization list
																		populateOrganization
																	%>
																	</select>
														            </td>
																</tr>
																<tr>
																	<td class=FieldCell width="151"> Organization Unit </td>
																	<td class='FieldCell'>
																	<select size="1" name="selOrgUnit" class="FormElem" onChange="LoadUnits()">
																		<option value="select">Select</option>
																	<%	'Calling the Function which populates the Organization Units list
																		populateOrganizationUnit
																	%>
																	</select>
														            </td>
																</tr>
																<tr>
																	<td class=FieldCell width="151"> Select </td>
																	<td class='FieldCell'>
																	<select size="1" name="selDivisionUnit"  class="FormElem" onChange="LoadDetails()">
																		<option value="select">Select</option>
																	</select>
														            </td>
																</tr>
															</table>
                                                        </td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="MiddlePack" width="100%" colspan="3">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
										<tr>
											<td width="100%">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="68"><p align="center">Describe
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
															<td class=MiddlePack colspan="2"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub> Name</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitName" size="25" maxlength=255 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> Short
                                                              Name </td>
															<td class='FieldCellSub'><input type="text" name="txtUnitShName" size="25" maxlength=20 class="Formelem"></td>
														</tr>
												</center>
													</table>
                                                            </td>
														</tr>
													</table>
                                                    </div>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="MiddlePack" width="100%" colspan="3">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
										<tr>
											<td align="center" class="FieldCell" width="100%" colspan="3">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="66"><p align="center">Address
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
															<td class='FieldCellSub' colspan="4"><input type="text" name="txtUnitAddr1" size="55" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> </td>
															<td class='FieldCellSub' colspan="4"><input type="text" name="txtUnitAddr2" size="55" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> PIN</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitPIN" size="7" maxlength=15 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>City</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitCity" size="30" maxlength=25 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> State</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitState" size="35" maxlength=25 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>Country</td>
															<td class='FieldCellSub'>
															<select size="1" name="selUnitCountry" class="FormElem">
															<Option value="Select">Select</Option>
															<%	'Calling the Function which populates the Country list
																populateCountry
															%>
															</select>
                                                            </td>
														</tr>
												</center>
														<tr>
															<td class=FieldCellSub> Phone</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitPhone" size="18" maxlength=15 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>Fax</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitFax" size="18" maxlength=15 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> E-mail ID</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitEmail" size="35" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'>URL</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitURL" size="30" maxlength=150 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> Contact Person</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitContactPerson" size="35" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub'></td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
                                                    </div>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%" colspan="3">
											</td>
										</tr>
										<tr>
											<td align="center" width="100%" colspan="3">
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
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="5"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub> TNGST RC
                                                              No.&nbsp;</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitTNGSTNo" size="35" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left">Area Code</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitAreaCode" size="4" maxlength=10 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> CST RC No.</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitCSTRCNo" size="35" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left"> CST Date</td>
															<td class='FieldCellSub'>
															<%
																' Function Call to Insert Date Picker
																Response.Write InsertDatePicker("UnitCSTRCDate")
															%>	
															</td>
														</tr>
												</center>
													</table>
                                                            </td>
														</tr>
													</table>
                                                    </div>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%" colspan="3">
											</td>
										</tr>

                                        <tr>
											<td align="center" class="MiddlePack" width="100%" colspan="3">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
										<tr>
											<td class="FieldCell" width="100%" colspan="3">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;</td>
															<td class='GroupTitle' width="98"><p align="center">Excise Details </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;</td>
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
															<td class=FieldCellSub> Range</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitRange" size="25" maxlength=255 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left">Division</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitDivision" size="25" maxlength=255 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> Collectorate</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitCollectorate" size="25" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left">Central excise No.</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitCentralENo" size="25" maxlength=75 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> Registration No.</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitRegNo" size="25" maxlength=75 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left">PLA No.</td>
															<td class='FieldCellSub'><input type="text" name="txtUnitLANo" size="10" maxlength=10 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> Address</td>
															<td class='FieldCellSub'><input type="text" name="txtRangeAdd1" size="25" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left">Address</td>
															<td class='FieldCellSub'><input type="text" name="txtDivisionAdd1" size="25" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> </td>
															<td class='FieldCellSub'><input type="text" name="txtRangeAdd2" size="25" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left"></td>
															<td class='FieldCellSub'><input type="text" name="txtDivisionAdd2" size="25" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub> </td>
															<td class='FieldCellSub'><input type="text" name="txtRangeAdd3" size="25" maxlength=50 class="Formelem"></td>
															<td class='FieldCellSub'></td>
															<td class='FieldCellSub' align="left"></td>
															<td class='FieldCellSub'><input type="text" name="txtDivisionAdd3" size="25" maxlength=50 class="Formelem"></td>
														</tr>
													</table>
												</center>
													</table>
                                                    </div>
                                                            </td>
														</tr>
													</table>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%" colspan="3">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td>
											</td>
											<td width="33%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Amend" name="B2" class="ActionButton" tabindex="3" onClick="javascript:checkSubmit('<%=FormatDate(now)%>')">
																<input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
														</td>
													</tr>
												</table>
											</td>
											<td>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="BottomPack" width="100%" colspan="3">
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
</form>
</BODY>
