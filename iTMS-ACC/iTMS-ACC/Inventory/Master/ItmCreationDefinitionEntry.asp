<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmCreationDefinitionEntry.asp
	'Module Name				:	Inventory (Item Creation and Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	October 11, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	ItmCreationDefinitionInsert.asp
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Creation and Definition</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" id="OutData" data-itms-xml-island>
<Output/>
</script>
<script type="application/xml" id="Data" data-itms-xml-island>
<root/>
</script>
<script type="application/xml" id="CategoryData" data-itms-xml-island>
<Root></Root>
</script>
<%
	Dim dcrs,dcrs1, oDOM, Root, Node, sItemTypeId,sOrgID,sQuery
	sOrgID = Session("organizationcode")

	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	set oDOM = server.CreateObject("Microsoft.XMLDOM")

	set Root = oDOM.createElement("ROOT")
	oDOM.appendChild Root

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME,ClassificationCode FROM INV_M_ITEMTYPEATTRIBUTES"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	do while Not dcrs.EOF
		set Node = oDOM.createElement("ATTRIBUTES")
		'Node.setAttribute "ITEMTYPEID",trim(dcrs(0))
		Node.setAttribute "ATTRID",trim(dcrs(0))
		Node.setAttribute "ATTRNAME",trim(dcrs(1))
		Node.setAttribute "ClassCode",trim(dcrs(2))
		Root.appendChild Node
		dcrs.movenext
	loop
	dcrs.Close

	oDOM.save server.MapPath("../Temp/Master/Attribute"&Session.SessionID&".Xml")
%>
<script type="application/xml" id="AttrData" data-itms-xml-island data-src="<%="../Temp/Master/Attribute"&Session.SessionID&".Xml"%>"></script>
<script type="application/xml" id="ItemAttData" data-itms-xml-island>
<Root></Root>
</script>
<script type="application/xml" id="ClassCode" data-itms-xml-island>
<Root/>
</script>
<script type="application/xml" id="HeadData" data-itms-xml-island>
    <Root>
        <%
            sQuery = "Select H.HeaderID,ItemTypeHeaderName from Inv_M_ItemTypeHeader H,Inv_M_ItemTypeAttributes A where H.HeaderID = A.HeaderID Group By H.HeaderID,ItemTypeHeaderName Order by H.headerId "
            dcrs.open sQuery,con
            if not dcrs.eof then
                do while not dcrs.eof
                        %>
                            <Header ID="<%=dcrs(0)%>" Name="<%=dcrs(1)%>"></Header>
                        <%
                    dcrs.movenext
                loop
            end if
            dcrs.close
        %>
    </Root>
</script>
<script type="application/xml" id="OptionData" data-itms-xml-island>
    <Root>
        <%
            sQuery ="Select A.ItemTypeAttributeID,ItemTypeAttributeName,OptionValue,OptionName from Inv_M_ItemTypeAttributes A,INV_M_ItemTypeOptions O where A.ItemTypeAttributeID= O.ItemTypeAttributeID "
            dcrs.open sQuery,con
            if not dcrs.eof then
                do while not dcrs.eof
                    %>
                        <Option AttID="<%=dcrs(0)%>" AttName="<%=dcrs(1)%>" ID="<%=dcrs(2)%>" Name="<%=dcrs(3)%>"></Option>
                    <%
                    dcrs.movenext
                loop
            end if
            dcrs.close
        %>
    </Root>
</script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmCreate.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemCreationDefinition.js"></SCRIPT>
<script language="javascript">
function Help() {
    window.open("../HelpFiles/Item.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px;left=10;top=10");
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload ="Init()">
<%
	' Declaration of variables
	dim arrFin,sFinFrom,sFinTo,sTempMonYr,sMonYr,sCheckFinYear,sFromDate,sToDate
	dim sFinPeriod
	sTempMonYr = mid(FormatDate(date),4,2)

	'sMonYr = sTempMonYr&Year(date())
	'sFinPeriod = split(GetFinancialYear(sMonYr),":")
	'sFinFrom =  sFinPeriod(0)
	'sFinTo = sFinPeriod(1)


	arrFin = split(Trim(Session("FinPeriod")),":")
	sFinFrom = "04"&arrFin(0)
	sFinTo = "03"&arrFin(1)

	sFromDate = "01/04/"&arrFin(0)
	sToDate = "31/03/"&arrFin(1)
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME," & Pack(sFromDate) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sToDate) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
'Response.Write dcrs.Source
	if Not dcrs.EOF then
		sCheckFinYear = "1"
	else
		sCheckFinYear = "0"
	end if
	dcrs.Close

	'arrFin = split(sFinFrom,"/")
	'sFinFrom = arrFin(1)&arrFin(2)
	'erase arrFin
	'arrFin = split(sFinTo,"/")
	'sFinTo = arrFin(1)&arrFin(2)

	'sFinTo = sMonYr



%>
	<form method="POST" name="formname" action>
	<input type="hidden" name="hGroup">
	<input type="hidden" name="hLevel">
	<input type="hidden" name="hFinYear" value="<%=sCheckFinYear%>">
	<input type="hidden" name="hFinFrom" value="<%=sFinFrom%>">
	<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
	<input type="hidden" name="hOpeningStockUnit">
	<input type="hidden" name="hClassCode" value="">
	<input type="hidden" name="hUnitID" value="<%=sOrgID %>">
	<input type="hidden" name="selUnit" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td align="center">
			    <table>
			        <tr>
			            <td class="PageTitle" >
			                Item Creation and Definition
			            </td>
			            <td class="PageTitle" >
			                <a style="text-decoration:none;font:color:black" href="#" onclick="Help()">Help</a>
			            </td>
			        </tr>
			    </table>
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack"></td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
					<tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" >
								<tr>
								    <td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="ItemListEntryForCreate.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
									<td class="TabCurrentCell" valign="bottom" align="center" width="50">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr>
												<td align="center">Basic
												</td>
											</tr>
										</table>
									</td>
									<td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="ItmDetailedDefn.asp">
												<td align="center">Purch. & Sales
												</td></a>
											</tr>
										</table>
									</td>
									<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="ItmInvDet.asp">
											    <td align="center">Inventory
											    </td>
										    </tr>
									    </table>
								    </td>
								    <td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="ItmManufacture.asp">
											    <td align="center">Manufacturing
											    </td>
										    </tr>
									    </table>
								    </td>
									<td class="TabCellEnd" valign="bottom" align="left">
										&nbsp;
								</td>
								</tr>
							</table>
						</td>
                	</tr>
					<tr>
						<td class="TabBody">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack"></td>
								</tr>

								<tr>
									<td align="center" width="5"></td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;</td>
															<td class="GroupTitle" width="55">
																<p align="center">Details
															</td>
															<td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>
													</table>
												</td>
											</tr>

											<tr>
												<td class="GroupTable"><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="MiddlePack"></td>
														</tr>

														<tr>
															<td class="FieldCellSub">
																<table border="0" cellspacing="0" cellpadding="0" width="100%">
																    <tr>
																        <td valign="top">
																            <table border="0" cellpadding="0" cellspacing ="0">
																                <tr>
																                    <td class="FieldCell">Item Type</td>
																		            <td class="FieldCellSub" colspan=2>
																			            <select size="1" name="selIType" class="FormElem" onChange="LetIType(this);ChangeLabel(this);ChangeBOM(this);ChangeAttribute(this)">
																			            <option value="select">select</option>
																			            <%	'Calling the Function which populates the Item Type list
																				            'populateItemType
																				            popItemTypesNew
																			            %>
																			            </select>
																		            </td>
																                </tr>
																                <tr>
																                    <td class="FieldCell">Item Code</td>
																		            <td class="FieldCellSub"  colspan=2>
																			            <input type="text" name="txtitmCode" size="19" maxlength=15 class="FormElem" onblur="CheckAvailability(this,'ItemCode')">
																			            &nbsp;&nbsp;
																			            <input type="button" value="Code Create" name="btnYrnCode" class="AddButton" onClick="CreateItemCode(this.value)" disabled>
																			            <!--<input type="button" value="Existing" class="AddButtonX" onClick="DisplayItemCode()" id=button1 name=button1>-->
																		            </td>
																                </tr>
																                <tr>
																                    <td class="FieldCell">Description</td>
																		            <td class="FieldCellSub" colspan="2">
																			            <input type="text" name="txtItmDesc" size="40" maxlength=60 class="FormElem" onblur="CheckAvailability(this,'ItemName')">
																		            </td>
																                </tr>
																                <tr>
																                    <td class="FieldCell">Add. Description</td>
																		            <td class="FieldCellSub">
    																		            <input type=text name="txtItmAddDesc" maxlength="200" class="FormElem" size="40">
    																		            <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" alt="Capture Item Spec." onclick="ItemSpecPop()">
																		            </td>
																                </tr>
																            </table>
																        </td>
																        <td valign="top">
																            <table border="0" cellpadding="0" cellspacing ="0">
																                <tr>
																                    <td class="FieldCellSub">Category</td>
																		            <td class="FieldCellSub" >
																		                <span id="spanCategory" class="DataOnly"></span>
																		                <input type="hidden" name="hCategory" value="">
																			            <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="popClass()" alt="Click to Add the Classification Code">
																		            </td>
																                </tr>
																                <tr>
																                    <td class="FieldCellSub">Classification</td>
																		            <td class="FieldCellSub">
																			            <input type='text' name="txtClass" class="FormElem" disabled>
																		            </td>
																                </tr>
																                <tr>
																                    <td class="FieldCellSub">Variant Code</td>
																		            <td class="FieldCellSub">
																			            <input type="text" name="txtVariant" size="12" maxlength=10 class="FormElem">
																		            </td>
																                </tr>
																                <tr>
																		            <td class="FieldCellSub">Stores UoM</td>
																		            <td class="FieldCellSub">
																			            <select size="1" name="selUoMStores" class="FormElem">
																				            <option value="select">Select</option>
																				            <%	'Calling the Function which populates the UoM list
																					            populateUoM()
																				            %>
																			            </select>
																			            <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="UoMDetails()" alt="Alternate UoM and other UoM's">
																		            </td>
																	            </tr>
																            </table>
																        </td>
																        <td valign="top" rowspan="2">
																            <table border="0" cellpadding="0" cellspacing ="0">
																                <tr>
																                    <td class="FieldCell">
																			            <div class="frmBody" id="frm1">
																				            <table border="0" cellspacing="1" class="BodyTable" width="100%">
																					            <tr>
																						            <td class="ExcelHeaderCell" align="center">
																						                <a href="#" class="ExcelDisplayLink" alt="Click here to upload image" style="cursor:hand;" onclick="UploadImage()">Add</a>
																						            </td>
																					            </tr>
																					            <tr>
																					                <td align="center">
																					                    <img src="../../assets/images/NoImage.gif" width="150" height="170" border="1"/>
																					                </td>
																					            </tr>
																				            </table>
																			            </div>
                                                                                    </td>
																                </tr>
																            </table>
																        </td>
																    </tr>
																	<tr>
																		<td class="FieldCell" colspan="2" valign="top">
																		    <div class="FrmBody" id="divAttributes" style="width:550;height:100;">
																			    <table border="0" cellspacing="1" class="BodyTable" width="100%">
																				    <tr>
																				        <td class="ExcelDisplayCell" align="center">Attributes [<a href="#" onclick="ManageAttribute()">Manage</a>] </td>
																				    </tr>
																					<tr>
																						<td  class="FieldCellSub">
																						    <table id="tblAttribute" cellpadding="0" cellspacing="0" width="100%">
																						    </table>
																						</td>
																					</tr>
																				</table>
																			</div>
																		</td>
																	</tr>
																</table>
															</td>
														</tr>
													</table>
                                                  </center>
												</td>
											</tr>
										</table>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0" border="0" width="100%">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;</td>
															<td class="GroupTitle" width="55">
																<p align="center">Controls
															</td>
															<td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>

													</table>
												</td>
											</tr>
											<tr>
												<td class="GroupTable">
													<table cellpadding="0" cellspacing="0" border="0">
														<tr>
															<td class="MiddlePack"></td>
														</tr>

														<tr>
															<td class="FieldCellSub">
																<table border="0" cellspacing="0" cellpadding="0">
																	<tr>
																		<td class="FieldCell">Receipt Numbering</td>
																		<td class="FieldCellSub">
																			<select size="1" name="selRecNum" class="FormElem" onChange="CheckNoSeries(this)">
																				<option value="select">Select</option>
																				<option value="L">Lot</option>
																				<option value="S">Serial</option>
																				<option value="LS">Lot / Serial</option>
																				<option value="N" selected>None</option>
																			</select>
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Receipt Routing</td>
																		<td class="FieldCell">
																			<select size="1" name="selRecRout" class="FormElem">
																				<option value="select">Select</option>
																				<option value="S" selected>Stock</option>
																				<option value="IS">Inspection / Stock</option>
																			</select>
																		</td>
																	</tr>

																	<tr>
																		<td class="FieldCell">Accounting Type</td>
																		<td class="FieldCellSub">
																			<select size="1" name="selAcc" class="FormElem">
																				<option value="select">Select</option>
																				<option value="L">LIFO</option>
																				<option value="F">FIFO</option>
																				<option value="W" selected>Weighted Average</option>
																			</select>
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Modvat Eligibility</td>
																		<td class="FieldCell">
																			<input type="radio" value="1" name="radMod" class="FormElem">   Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																			<input type="radio" value="0" name="radMod" class="FormElem" checked>   No
																		</td>
																	</tr>

																	<tr>
																		<td class="FieldCell">Account Head</td>
																		<td class="FieldCellSub">
																		    <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" alt="Account Head" onclick="SelectAccHead()">
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Bill of Material</td>
																		<td class="FieldCell">
																			<input type="radio" value="1" name="radBoM" class="FormElem" disabled="true">   Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																			<input type="radio" value="0" name="radBoM" class="FormElem" checked disabled="true">   No
																			<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="GetDetails()" alt="Select BoM">
																		</td>
																	</tr>
																	<tr>
																		<td class="FieldCell">Opening Acc Head</td>
																		<td class="FieldCellSub">
																		    <span id="spanOpenAccHead" class="DataOnly">&nbsp;
																		    </span>
																		    <input type="hidden" name="hOAH" value="0" />
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Pur. Tax Type </td>
																		<td class="FieldCell">
																			<select name="SelPurTaxType" class="FormElem">
																			    <option value="0">Select</option>
																			    <%
																			        populatePurTaxType()
																			    %>
																			</select>
																		</td>
																	</tr>
																	<tr>
																		<td class="FieldCell">Closing Acc Head</td>
																		<td class="FieldCellSub">
																		    <span id="spanCloseAccHead" class="DataOnly">&nbsp;
																		    </span>
																		    <input type="hidden" name="hCAH" value="0" />
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Sales Tax Type</td>
																		<td class="FieldCell">
																			<select name="SelSalTaxType" class="FormElem">
																			    <option value="0">Select</option>
																			    <%
																			        populateSalTaxType()
																			    %>
																			</select>
																		</td>
																	</tr>

																</table>
															</td>
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

								<!--<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;</td>
															<td class="GroupTitle" width="90">
																<p align="center">Other Details
															</td>
															<td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>

													</table>
												</td>
											</tr>

											<tr>
												<td class="GroupTable">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class="MiddlePack"></td>
														</tr>

														<tr>
															<td class="FieldCellSub">
																<table border="0" cellspacing="0" cellpadding="0">
																	<tr>
																		<td class="FieldCell"><span id="idCat">Catalogue No.</span></td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtItmCat" size="20" maxlength=30 class="FormElem">
																		</td>
																		<td class="FieldCell"><span id="idDrw">Draw. Ver</span></td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtItmDrw" size="20" maxlength=20 class="FormElem">
																		</td>
																		<td class="FieldCellSub" colspan="3"></td>
																	</tr>
																</table>
															</td>
														</tr>

													</table>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>-->

 								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%"><center>
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;
															</td>
															<td class="GroupTitle" width="98">
																<p align="center">Opening Stock
															</td>
														</center><td class="GroupTitleRight">
															<p align="left">&nbsp;
														</td>
													</tr>

												</table>
											</td>
										</tr>
										<tr>
											<td class="GroupTable"><center>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="MiddlePack">
														</td>
													</tr>
														<tr>
															<td class="FieldCellSub">Storage Location &nbsp;
 																<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="GetStore()" alt="Storage Location">
																<span class="DataOnly" id="idStore"></span>
															</td>
														</tr>

													<tr>
														<td class="FieldCellSub">
															<div class="frmBody" id="frm2" style="width: 570; height:75;">
																<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
																	<tr>
																		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
																		<td class="ExcelHeaderCell" align="center">Storage</td>
																		<td class="ExcelHeaderCell" align="center">MonthYear</td>
																		<td class="ExcelHeaderCell" align="center">Quantity</td>
																		<td class="ExcelHeaderCell" align="center">Value</td>
																		<td class="ExcelHeaderCell" align="center" width="25">Lot/Serial</td>
																	</tr>

																</table>
															</div>
														</td>
													</tr>

												</center>
												</table>
											</td>
										</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td align="Left" colspan="2" class="FieldCell">
									    <input type=checkbox name="ChkAllUnit" value="ALL" class="FormElem" disabled>Applicable for other Unit(s)
									</td>
								</tr>
								<tr>
									<td class="MiddlePack">

									</td>
								</tr>
								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<input type="button" value="Save" name="B2" class="ActionButton" onClick="CheckSubmitDetails('<%=sFinFrom%>','<%=sFinTo%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
 													<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
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
</body>
</html>
<%
Function popItemTypesNew()
Dim rsTemp,ssql
set rsTemp = Server.CreateObject("ADODB.Recordset")
ssql = "Select ItemTypeID,ItemTypeDescription from INV_M_ItemTypes"
rsTemp.Open ssql,con
if not rsTemp.EOF then
    do while not rsTemp.EOF
        Response.Write "<option value="&Trim(rsTemp(0))&">"&Trim(rsTemp(1))&"</option>"

        rsTemp.MoveNext
    loop
end if
rsTemp.Close
End Function
%>
<%
Function populatePurTaxType()
Dim rsTemp,sQuery
set rsTemp = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select PurchaseType,PurchaseTypeName from APP_M_PurchaseTypes where upper(isNull(Active,'Y')) = 'Y'  ORDER BY PURCHASETYPE"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        do while not rsTemp.eof
            response.write "<option value="&trim(rsTemp(0))&">"& trim(rsTemp(1)) &"</option>"
            rsTemp.movenext
        loop
    end if
    rsTemp.close
End Function
%>

<%
Function populateSalTaxType()
Dim rsTemp,sQuery
set rsTemp = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select InvoiceType,InvoiceTypeName from Sal_M_InvoiceTypes where isNull(Useable,0)=1  Order by InvoiceType"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        do while not rsTemp.eof
            response.write "<option value="&trim(rsTemp(0))&">"& trim(rsTemp(1)) &"</option>"
            rsTemp.movenext
        loop
    end if
    rsTemp.close
End Function
%>




