<%@ Language=VBScript %>
<%option explicit	%>
<%
	'Program Name				:	ItmDetailedDefn.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:
	'Created On					:
	'Modified By				:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/ItemDisplay.asp"-->
<%
    Dim rsTemp
    Dim iItmCode,iClassCode
    Dim sItemName,sClassName,sQuery,sUnder,sOver,sUnOrder,sUoMDesc,sOrgCode,sOrgName,sSalesUOM

    set rsTemp = Server.CreateObject("ADODB.Recordset")

    iItmCode = Request("ItemCode")
    iClassCode = Request("ClassCode")
    sOrgCode = Session("organizationcode")
    sOrgName = Session("OrgShortName")
    sUnder = 1
    sOver = 1
    sUnOrder = 1

if trim(iItmCode)<>"" then
    sQuery = "Select ItemDescription,(Select GroupName from INV_M_Classification where GroupCode = V.ClassificationCode),SalesUOM from VWITEM V where ItemCode = "& iItmCode  &" and ClassificationCode = "& iClassCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sItemName = rsTemp(0)
        sClassName = rsTemp(1)
        sSalesUOM = rsTemp(2)
    end if
    rsTemp.Close
end if 'if trim(iItmCode)<>"" then
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Control Definition - Purchase</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="storageData"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemDetailedDefinition.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><root></root></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
    <Root ItemName="<%=sItemName%>" ClassName="<%=sClassName%>" OrgCode="<%=sOrgCode%>" ItemCode="<%=iItmCode%>" ClassCode="<%=iClassCode%>" OrgName="<%=sOrgName%>">
        <Purchase></Purchase>
        <Sales></Sales>
    </Root>
</script>

</HEAD>
<BODY leftMargin=0 topMargin=0 >
<form method="POST" name="formname" action="">
<input type="hidden" name="hItmCode" value="<%=iItmCode%>">
<input type="hidden" name="hClassCode" value="<%=iClassCode%>">
<input type="hidden" name="hOrgCode" value="<%=sOrgCode%>">
<input type="hidden" name="hSuppName" value="">
<input type="hidden" name="hSuppCode" value="">
<input type="hidden" name="hSuppType" value="">
<input type="hidden" name="hSuppSubType" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Detail Definition
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<tr>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
							    <td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="ItemListEntryForCreate.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center"><a href="ItmCreationDefinitionEntry.asp?Flag=O&iItmCode=<%=iItmCode%>">Basic</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" width="90">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Purch. & Sales
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmInvDet.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=iClassCode%>">
											<td align="center">Inventory
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="ItmManufacture.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=iClassCode%>">
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
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellspacing="0" cellpadding="0" class="BodyTable">
									    <tr>
											<td class="FieldCellSub" width="80">Item Name</td>
											<td>
											<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
											<td class="FieldCell" width="15"></td>
											<td class="FieldCell" width="82">Classification</td>
											<td>
											<span class="DataOnly"><%=sClassName%>&nbsp;</span>
											&nbsp;</td>
											<td></td>
									    </tr>
                                    </table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <!----------------->
                            <tr>
                            <td align="center" width="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							<td valign="top">
								<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
									<tr>
										<td>
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td align="center" colspan="3" class="MiddlePack">
														<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
													</td>
												</tr>

                            <tr>
							<td align="center" width="5" class="ClearPixel">
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							<td valign="top" width="100%">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" class=BodyTable>

                            <!----------------->
                            <tr>
                                <td ></td>
								<td align="center">
								<div>
									<table class="CollapseBand" cellspacing="0" cellpadding="0">
										<tr>
											<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(divPurRec,'')" itms_state="0">
												<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
												</a>
											</td>
											<td valign="center" class="SubTitle">&nbsp;&nbsp;
											    Purchasing and Receiving
											</td>
										</tr>
									</table>
									<table border="0" cellpadding="0" cellspacing="0" width=100% class=BodyTable>
									<tr>
										<td width="100%">
                            <div id=divPurRec style="width:100%;display:none;">
                            <table>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
   							       <div align="left">
   								<table cellpadding="0" cellspacing="0" width="100%">
   									<tr>
   										<td>
   								            <table cellpadding="0" cellspacing="0" width="100%">
   									            <tr>
   										            <td class='GroupTitleLeft' width="10">&nbsp;
                                                       </td>
   										            <td class='GroupTitle' width="80"><p align="center">
                                                       Purchasing
                                                       </td>

   										            <td class='GroupTitleRight'><p align="left">&nbsp;
                                                       </td>
   									            </tr>
   								            </table>
                                       </td>
   									</tr>
   									<tr>
   										<td class=GroupTable>
                                   <div align="left">
   								<table cellpadding="0" cellspacing="0">
   									<tr>
   										<td class=MiddlePack> </td>
   									</tr>
   									<tr>
   										<td>
                                           <table border="0" cellspacing="0" cellpadding="0">
                                       <tr>
										<td class="FieldCellSub" width="105">Buyer</td>
										<td class="FieldCell" colspan="3">
											<select size="1" name="selBuyer" class="FormElem">
												<option value="select">Select</option>
												<%	'Calling the Function which populates the Employee list
													populateEmployee
												%>
											</select>
										</td>
                                       </tr>
                                       <tr>
										<td class="FieldCellSub">Alternate Item</td>
										<td class="FieldCell" width="110">
										    <input type="button" value="Select" name="btnCheck" class="AddButton" onClick="OpenAlter()">
										</td>
										<td class="FieldCellSub" width="85">Optional UoM</td>
										<td class="FieldCell">
											<input type="button" value="Select" name="btnUoMPur" class="AddButton" onClick="OpenUoM('Pur')">
										</td>
                                       </tr>


                                       <!--<tr>
										<td class="FieldCellSub" width="110">Modvat</td>
										<td class="FieldCell" width="200">
										    <input type="radio" value="1" name="radMod" class="FormElem">   Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										    <input type="radio" value="0" name="radMod" class="FormElem" CHECKED>   No
										</td>
									   </tr>-->
                                       <tr>
										<td class="FieldCellSub" width="110">Invoice Matching</td>
										<td class="FieldCell">
                                           <select size="1" name="selInvMat" class="FormElem">
												<option value="select">Select</option>
												<option value="2" >2 Way</option>
												<option value="3" >3 Way</option>
												<option value="4" >4 Way</option>
									       </select>
										</td>
										<td class="FieldCellSub" width="110">Sub-Contracting</td>
										<td class="FieldCell">
										    <input type="radio" value="1" name="radSub" class="FormElem">   Yes&nbsp;&nbsp;
										    <input type="radio" value="0" name="radSub" class="FormElem" checked>   No
										</td>
                                       </tr>
                                       <tr>
								            <td class=FieldCellSub width="130">Substitute Receipts</td>
								            <td class='FieldCell' width="200">
                                               <input type="radio" value="1" name="radSubRec" class="FormElem"> Yes&nbsp;&nbsp;
                                               <input type="radio" value="0" name="radSubRec" class="FormElem" checked> No
                                           </td>
                                            <td class=FieldCellSub width="130">Enforce Ship To</td>
								            <td class='FieldCell'>
                                               <input type="radio" value="1" name="radShip" class="FormElem"> Yes&nbsp;&nbsp;
                                               <input type="radio" value="0" name="radShip" class="FormElem" checked> No
                                           </td>
                                       </tr>

                                           </table>
                                           </td>
   									</tr>
   									<tr>
   										<td class=MiddlePack>
                                           </td>
   									</tr>
   									<tr>
   										<td class=FieldCellSub>
                                           <table border="0" cellspacing="0" cellpadding="0">

                                           <tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p></td>
															<td class="GroupTitle" width="150">
																<p align="center"><input type="checkbox" value="Y" name="chkVendor" class="FormElem" onclick="resetVendor(this)"> Preferred Vendor
	                                                        </td>
															<td class="GroupTitleRight"><p align="left">&nbsp;</td>
														</tr>
													</table>
												</td>
											</tr>
                                       <tr>
                                       <td class=GroupTable>
                                       <table cellpadding=0 cellspacing=0 border=0>
                                      <tr>
										<td class="FieldCellSub">Vendor</td>
										<td class="FieldCellSub">
										    <span id="txtParty" class="DataOnly">&nbsp;</span>
    										<img src="../../assets/images/iTMS%20icons/Entryicon.gif" onclick="PopulateSupplier()">
										</td>
										<td class="FieldCellSub" width="5"></td>
										<td class="FieldCellSub" width="147">Supplier Item Description</td>
										<td class="FieldCellSub" width="94"> <input type="text" name="txtSupItmDesc" size="15" maxlength=100 class="Formelem"> </td>
                                       </tr>
                                       <tr>
										<td class="FieldCellSub" width="147">Supplier Item Code</td>
										<td class="FieldCellSub" width="94"> <input type="text" name="txtSupItmNo" size="15" maxlength=10 class="Formelem"> </td>
										<td class="FieldCellSub" width="5"></td>
										<td class="FieldCellSub" width="147">Supplier Drawing No</td>
										<td class="FieldCellSub" width="94"> <input type="text" name="txtSuppDrawingNo" size="15" maxlength=10 class="Formelem"></td>
                                       </tr>
                                       <tr>
										<td class="FieldCellSub" width="147">Supplier UOM</td>
										<td class="FieldCellSub" width="94">
										    <select name="selSuppUOM" class=FormElem>
										        <%
										            PopulateUOM
										        %>
										    </select>
										</td>
										<td class="FieldCellSub" width="5"></td>
										<td class="FieldCellSub" width="147">Transit Lead Time</td>
										<td class="FieldCellSub" width="94"> <input type="text" name="txtTrLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
                                       </tr>
                                       <tr>
										<td class="FieldCellSub" width="152">Purchase Lead Time</td>
										<td class="FieldCellSub" width="89"> <input type="text" name="txtPuLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
										<td class="FieldCellSub" width="5"></td>
										<td class="FieldCellSub" width="152">Purchase Warranty Period</td>
										<td class="FieldCellSub" width="89"> <input type="text" name="txtPurWarranty" size="4" maxlength=3 class="Formelem"> (in days)</td>

                                       </tr>
                                       <tr>
										<td class="FieldCellSub" width="152">Supplier Lead Time</td>
										<td class="FieldCellSub" width="89"> <input type="text" name="txtSuLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
										<td class="FieldCellSub" width="5"></td>
										<td class="FieldCellSub" width="147">Market Price</td>
										<td class="FieldCellSub" width="94"> <input type="text" name="txtMarketPrice" size="12" maxlength=10 class="Formelem"> </td>
                                       </tr>
                                       <tr>
										<td class="FieldCellSub" width="152">Preorder Lead Time</td>
										<td class="FieldCellSub" width="89"> <input type="text" name="txtPrLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
										<td class="FieldCellSub" width="5"></td>
										<td class="FieldCellSub" width="147">Market Date</td>
										<td class="FieldCellSub" width="94"> <input type="text" name="txtMarketDate" size="12" maxlength=10 class="Formelem"> </td>
                                       </tr>
                                       <tr>
										<td class="FieldCellSub" width="152">Preferred Min. Order Qty</td>
										<td class="FieldCellSub" width="89"> <input type="text" name="txtPreMinQty" size="12" maxlength=10 class="Formelem"> </td>
										<td class="FieldCellSub" width="5"></td>
										<td class="FieldCellSub" width="147">Preferred Max. Order Qty</td>
										<td class="FieldCellSub" width="94"> <input type="text" name="txtPreMaxQty" size="12" maxlength=10 class="Formelem"> </td>
                                       </tr>

                                       </table>
                                       </td>
   									</tr>
   									</table>
   									</td>
   									</tr>
   									<tr>
   										<td class=MiddlePack>
                                           </td>
   									</tr>
   								</table>
                               </div>
                                           </td>
   									</tr>
   								</table>
                                   </div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
   								 	   <div align="left">
   								 		<table cellpadding="0" cellspacing="0" width="100%">
   								 			<tr>
   								 				<td>
   								 		            <table cellpadding="0" cellspacing="0" width="100%" height="14">
   								 			            <tr>
   								 				            <td class='GroupTitleLeft' width="10" height="14">&nbsp;
                                                                </td>
   								 				            <td class='GroupTitle' width="132" height="14"><p align="center">
                                                                Additional Receiving
                                                                </td>
   								 				            <td class='GroupTitleRight' height="14"><p align="left">&nbsp;
                                                                </td>
   								 			            </tr>
   								 		            </table>
                                                </td>
   								 			</tr>
   								 			<tr>
   								 				<td class=GroupTable>
   								                 	<center>
                                            <div align="left">
                                        <table border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                <td valign="top" colspan="3" class="MiddlePack"></td>
                                    </tr>
                                    <tr>
                                <td valign="top"><div align="left">
   								 		<table cellpadding="0" cellspacing="0">
                                            <tr>
   								 				<td class=FieldCellSub width="117">Receipt Date Action</td>
   								 				<td class='FieldCell' width="215">
                                                    <input type="radio" name="radReDate" value="R" class="FormElem" > Reject&nbsp;
                                                    <input type="radio" name="radReDate" value="W" class="FormElem" > Warning&nbsp;
                                                    <input type="radio" name="radReDate" value="N" class="FormElem" checked > None
		                                        </td>
                                            </tr>
                                            <tr>
   								 				<td class=FieldCellSub width="117">Receipt Days Early (in days)</td>
   								 				<td class='FieldCellSub' width="215">
													<input type="text" name="txtRecDaysE" size="4" maxlength=3 value="0" class="Formelem" readonly>
												</td>
                                            </tr>
                                            <tr>
   								 				<td class=FieldCellSub width="117">Receipt Days Late (in days)</td>
   								 				<td class='FieldCellSub' width="215">
													<input type="text" name="txtRecDaysL" size="4" maxlength=3 value="0" class="Formelem" readonly>
		                                        </td>
                                            </tr>
   								 		</table>
                                </div></td>
                                <td width="5" valign="top"></td>
                                <td valign="top" align="right">
                            <div align="left">
   						<table cellpadding="0" cellspacing="0">
   							<tr>
   								<td>
   						<table cellpadding="0" cellspacing="0" width="100%" height="14">
   							<tr>
   								<td class='GroupTitleLeft' width="10" height="14">&nbsp;
                                    </td>
   								<td class='GroupTitle' width="100" height="14"><p align="center">Tolerance in %
                                    </td>
   								<td class='GroupTitleRight' height="14"><p align="left">&nbsp;
                                    </td>
   							</tr>
   						</table>
                                </td>
   							</tr>
   							<tr>
   								<td class=GroupTable>
                        <div align="left">
   						<table cellpadding="0" cellspacing="0">
   							<tr>
   								<td class=MiddlePack colspan="3">  </td>
   							</tr>
                            <tr>
   								<td class=FieldCellSub>
                                    </td>
   								<td class='FieldCellSub' align="center">Low</td>
   								<td class='FieldCellSub' align="center">High</td>
                            </tr>
                            <tr>
   								<td class=FieldCellSub>Under Receipts</td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnLow" size="3" maxlength=3 class="Formelem">
                        </td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnHigh" size="3" maxlength=3 class="Formelem">
                        </td>
                            </tr>
                            <tr>
   								<td class=FieldCellSub>Over Receipts</td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtOvLow" size="3" maxlength=3 class="Formelem">
                        </td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtOvHigh" size="3" maxlength=3 class="Formelem">
                        </td>
                            </tr>
                            <tr>
   								<td class=FieldCellSub>Unordered Receipts</td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnOrLow" size="3" maxlength=3 class="Formelem">
                        </td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnOrHigh" size="3" maxlength=3 class="Formelem">
                        </td>
                            </tr>
   						</table>
                        </div>
                                    </td>
   							</tr>
   						</table>
                            </div>
                        </td>
                            </tr>
                            <tr>
                        <td valign="top" colspan="3" class="MiddlePack"></td>
                            </tr>
                                </table>
                                    </div>
								</center>
                                                </td>
										</tr>
									</table>
                                    </div>
								</td>
								<td align="center">
								</td>
								</tr>
								<tr>
								    <td align="center"></td>
								    <td align="center" class="ActionCell">
									    <input type=button name="btnPurSave" value="Save" class="ActionButtonX" onclick="PurRecSubmit()">
								    </td>
								    <td align="center"></td>
                                </tr>
								</table>
								</div><!-- id=divPurRec-->
								</td>
								</tr>
								</table>
								</div><!--div-->
								</td>
								</tr>
								<!-------->
								</table>
                            </td>
                            </tr>
                            </TABLE>
                            </td>
                            </tr>
                            </table>
                            </td>
                            </tr>
								<!-------->
								                            <!----------------->
                            <tr>
                            <td align="center" width="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							<td valign="top">
								<table id="Table1" cellspacing="0" cellpadding="0" border="0" width="100%">
									<tr>
										<td >
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td align="center" colspan="3" class="MiddlePack">
														<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
													</td>
												</tr>

                            <tr>
							<td align="center" width="5" class="ClearPixel">
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							<td valign="top" width="100%">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" class=BodyTable>

                            <!----------------->
								<tr>
								<td align="center">
								<div>
									<table class="CollapseBand" cellspacing="0" cellpadding="0" >
										<tr>
											<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(divSales,'')" itms_state="0">
												<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
												</a>
											</td>
											<td valign="center" class="SubTitle">&nbsp;&nbsp;
										 	    Sales
											</td>
										</tr>
									</table>
									<table border="0" cellpadding="0" cellspacing="0" width=100% class=BodyTable>
									<tr>
										<td width="100%">
                            <div id=divSales style="width:100%;display:none;">
                            <table>

								                            <tr>
																<td align="center" width="5">
																</td>
																<td valign="top" width="100%">
								                                    <div align="left">
								                                        <table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly">
																		    <tr>
																				<td class="FieldCellSub" width="75">Basic UoM</td>
																				<td><span class="DataOnly"><%=sSalesUOM%>&nbsp;</span></td>
																				<td></td>
																				<td></td>
																				<td class="FieldCellSub" width="140">Optional UoM</td>
																				<td colspan="2">
																				<%	if sSalesUOM = "N/A" then %>
																				    <input type="button" value="Select" name="btnUoMSal" class="AddButton" DISABLED>
																				<%	else %>
																				    <input type="button" value="Select" name="btnUoMSal" class="AddButton" onClick="OpenUoM('Sal')">
																				<%	end if %>
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="75">Market Rate</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtMarketRate" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Sales Warranty Period</td>
																				<td class="FieldCell" colspan="2">
																					<input type="text" name="txtSalWarranty" size="3" maxlength=3 class="Formelem"> (in Days)
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" colspan="3" align="center"><p align="left">Selling Price:</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Minimum Sale Quantity</td>
																				<td class="FieldCell">
																					<input type="text" name="txtMinSale" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td><span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>&nbsp;</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="75">Actual</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtActual" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Unit Size</td>
																				<td class="FieldCell">
																					<input type="text" name="txtUnitSize" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCell">
																					<select size="1" name="selUoMUnit" class="FormElem">
																						<option value="select">Select</option>
																						<%	'Calling the Function which populates the UoM list
																							populateUoM
																						%>
								                                                    </select>
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="75">Minimum</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtMin" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Volume</td>
																				<td class="FieldCell">
																					<input type="text" name="txtVolume" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCell">
																					<select size="1" name="selUoMVolume" class="FormElem">
																						<option value="select">Select</option>
																						<%	'Calling the Function which populates the UoM list
																							populateUoM
																						%>
								                                                    </select>
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="75">Preferred</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtPreffered" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140"></td>
																				<td class="FieldCell" >
																				</td>
																				<td></td>
																			</tr>
								                                        </table>
								                                    </div>
																</td>
																<td align="center">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>

								                            <tr>
																<td align="center" width="5">
																</td>
																<td valign="top" width="100%">
															        <div align="left">
															        <table border="0" cellspacing="0" cellpadding="0">
																	    <tr>
																			<td class="FieldCell" width="105" valign="top">Commodity</td>
																			<td class="FieldCellSub">
																		    <select size="1" name="selCommodity" class="FormElem">
																			<option value="select">Select</option>
																			<%	'Calling the Function which populates the Commodity list
																				'populateCommodity(iCommodity)
																			%>
																			</select>
																			</td>
																	    </tr>
															        </table>
															        </div>
																</td>
																<td align="center">
																</td>
								                            </tr>

								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" width="5"></td>
																<td valign="top" width="100%">
								                                    <table cellpadding="0" cellspacing="0">
																		<tr>
																			<td>
																				<table cellpadding="0" cellspacing="0" width="100%">
																					<tr>
																						<td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p></td>
																						<td class="GroupTitle" width="120">
																							<p align="center"><input type="checkbox" value="Y" name="chkDiscount" class="FormElem" onclick="resetDiscount(this)"> Sales Discount
								                                                        </td>
																						<td class="GroupTitleRight"><p align="left">&nbsp;</td>
																					</tr>
																				</table>
																			</td>
																		</tr>
																		<tr>
																			<td class=GroupTable>
																				<div align="left">
																					        <table cellpadding="0" cellspacing="0" width="100%">
																						        <tr>
																							        <td class=MiddlePack colspan="2"> <p align="left"> </td>
																						        </tr>
																						        <tr>
																							        <td class=FieldCellSub colspan="2"> <p align="left">&nbsp;Which Discount should take precendence <input type=radio name=radQV value=Q checked>Quantity&nbsp;&nbsp;<input type=radio name=radQV value=V>Value&nbsp; </td>
																						        </tr>
																						        <tr>
																							        <td class=FieldCellSub colspan="2"> <p align="left">&nbsp;Discount applicable in <input type=radio name=radApplicable value=B checked>Basic Value&nbsp;&nbsp;<input type=radio name=radApplicable value=T>Total Value&nbsp;&nbsp;<input type=radio name=radApplicable value=P>Purchase Value</td>
																						        </tr>

																						        <tr>
																							        <td class=FieldCellSub>
																								        <div align="left">
        																								<table border=0 width=100%>
        																								<tr>
        																								<td valign=top>
																								        <table border="0" cellspacing="1" id="tblDataQty" class="ExcelTable" width=100%>
																				                            <tr>
																						                        <td class="ExcelHeaderCell" align="center" width="30" rowspan="2">S.No.</td>
																						                        <td class="ExcelHeaderCell" align="center" colspan="2">Quantity</td>
																						                        <td class="ExcelHeaderCell" align="center" width="75" rowspan="2">Discount %</td>
																						                        <td class="ExcelHeaderCell" align="center" width="75" rowspan="2">UoM</td>
																						                        <td class="ExcelHeaderCell" align="center" rowspan="2"></td>
																				                            </tr>
																				                            <tr>
																						                        <td class="ExcelHeaderCell" align="center" width="75">From</td>
																						                        <td class="ExcelHeaderCell" align="center" width="75">To</td>
																				                            </tr>
																				                            <tr>
																				                                <td class="ExcelSerial">
																				                                <td class="ExcelDisplayCell">
																							                        <input type="text" name="txtQtyFrom" size="11" maxlength=10 class="Formelem">
																					                            </td>
																					                            <td class="ExcelDisplayCell">
																							                        <input type="text" name="txtQtyTo" size="11" maxlength=10 class="Formelem">
																						                        </td>
																						                        <td class="ExcelDisplayCell">
																							                        <select size="1" name="selUoMQty" class="FormElem">
																								                        <!--<option value="select">Select</option>-->
																								                        <%	'Calling the Function which populates the UoM list
																									                        populateUOMForItem iItmCode,iClassCode,sOrgCode
																								                        %>
																							                        </select>
																						                        </td>
																						                        <td class="ExcelDisplayCell">
																							                        <input type="text" name="txtQtyDis" size="4" maxlength=3 class="Formelem">
																						                        </td>
																						                        <td class="ExcelDisplayCell" >
																			                                      <p align="center"><input type="button" value=" Add " name="B5" class="AddButtonX" onClick="CheckEntryDis()">
																						                        </td>
																				                            </tr>
																				                        </table></td><td valign=top>
																				                        <table border="0" cellspacing="1" id="tblDataVal" class="ExcelTable" width=100%>
																				                        <tr>
																						                    <td class="ExcelHeaderCell" align="center" width="30" rowspan="2">S.No.</td>
																						                    <td class="ExcelHeaderCell" align="center" colspan="2">Value</td>
																						                    <td class="ExcelHeaderCell" align="center" width="75" rowspan="2">Discount %</td>
																						                    <td class="ExcelHeaderCell" align="center" rowspan="2"></td>
																				                        </tr>
																				                        <tr>
																						                    <td class="ExcelHeaderCell" align="center" width="75">From</td>
																						                    <td class="ExcelHeaderCell" align="center" width="75">To</td>
																				                        </tr>
																				                        <tr>
																				                            <td class=ExcelSerial></td>
																				                            <td class="ExcelDisplayCell">
																							                    <input type="text" name="txtValFrom" size="11" maxlength=10 class="Formelem">
																						                    </td>
																						                    <td class="ExcelDisplayCell">
																							                    <input type="text" name="txtValTo" size="11" maxlength=10 class="Formelem">
																						                    </td>
																						                    <td class="ExcelDisplayCell">
																							                    <input type="text" name="txtValDis" size="4" maxlength=3 class="Formelem">
																						                    </td>
																						                    <td class="ExcelDisplayCell">
																			                                  <p align="center"><input type="button" value=" Add " name="B6" class="AddButtonX" onClick="CheckEntryVal()">
																						                    </td>
																				                        </tr>
																				                    </table>
																				                    </td>
																				                    </tr>
																				                    </table>
																						</div>
																							</td>
																						</tr>
																					</table>
																				</div>
																			</td>
																		</tr>
								                                    </table>
																</td>
																<td align="center"></td>
								                            </tr>
								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" width="5">
																</td>
																<td valign="top" width="100%">
								                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
																	    <tr>
																			<td>
																			<!--	<DIV class=frmBody id=frm3 style="width: 290; height:75;">-->

																			<!--	</div>-->
																			</td>
																			<td>
																				<!--<DIV class=frmBody id=frm31 style="width: 290; height:75;">-->

																				<!--</div>-->
																			</td>
																	    </tr>
								                                    </table>
																</td>
																<td align="center">
																</td>
								                            </tr>
                                                        <tr>
							                                <td align="center"></td>
							                                <td align="center" class="ActionCell">
								                                <input type=button name="btnSalSave" value="Save" class="ActionButtonX" onclick="SalSubmit()">
							                                </td>
							                                <td align="center"></td>
                                                        </tr>

                            </table>
                            </div><!-- id=divSales-->
                            </td>
                            </tr>
                            </table>
                            </div><!--div-->
                            </td>
                            </tr>
                            								<!-------->
								</table>
                            </td>
                            </tr>
                            </TABLE>
                            </td>
                            </tr>
                            </table>
                            </td>
                            </tr>
								<!-------->

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
												<p align="center">
                                                    <input type="button" value="Back" name="B5" class="ActionButton" onClick="window.location.href='ItmCreationDefinitionEntry.asp?Flag=O&iItmCode=<%=iItmCode%>'">
                                                    <input type="button" value="Cancel" name="B1" class="ActionButton" onclick="window.location.href='ITEMLISTENTRY.ASP?ACTN=L'">
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
</HTML>
<%
    Function populateUOMForItem(ItemCode,ClassCode,OrgCode)
        if trim(ItemCode)<>"" then
            sQuery = "Select SalesUOM from INV_M_ITEMMASTER where ItemCode="& ItemCode &" and ClassificationCode = "& ClassCode &" and organisationcode = "& OrgCode
            rsTemp.Open sQuery,con
            if not rsTemp.EOF then
                Response.Write "<option value="& Trim(rsTemp(0)) &">"& Trim(rsTemp(0)) &"</option>"
            end if
            rsTemp.Close
        end if
    End Function
%>
