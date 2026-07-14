<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	PackingTypes.asp
	'Module Name				:	INVENTORY 
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	April 04,2011
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
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/CheckPrevFinYear.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->
<%
    Dim rsObj
    Dim sOrgID,sOrgName,sCreatedBy,sPackingName,sQuery,sNumberingType
    Dim iSlNo
    set rsObj = Server.CreateObject("ADODB.Recordset")    
    sOrgID = Session("organizationcode")
    sOrgName = Session("OrgShortName")
	sCreatedBy = Session("userid")
	sPackingName = Request.QueryString("PackName")
	if trim(Request.QueryString)="" then
	    sPackingName = Request.Form("hPackName")
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Paking Types</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/packingTypes.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" onload="fnInit()">
	<form method="POST" name="formname" action="">
	    <input type=hidden name="hOrgID" value="<%=sOrgID%>">
		<input type=hidden name="hOrgName" value="<%=sOrgName%>">
		<input type=hidden name="hPackName" value="<%=sPackingName%>">
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" height="20">
				    <table>
			            <tr>
			                <td class="PageTitle" >
			                    <P align=center>Packing Types</P>
			                </td>
			                <td class="PageTitle" >
			                    <a style="text-decoration:none;font:color:black" href="#" onclick="Help()">Help</a>
			                </td>
			            </tr>
			        </table>
				</td>
			</tr>
			<tr>
				<td align="center" class="TopPack">
				</td>
			</tr>
			<tr>
				<td valign="top">
					<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
				        <TR>
					        <td height="20" valign="bottom">
						        <table border="0" cellpadding="0" cellspacing="0" width="100%">
							        <tr>
							            <td class="TabCell" valign="bottom" align="center" width="105">
								          <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								           <tr><a href="OrgControlDefn.asp">
									          <td width="100%" align="center">Org Defn</td></a>
									        </tr>
								          </table>
								        </td>
								        <td class="TabCell" valign="bottom" align="center" width="125">
									        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										        <tr><a href="StoreLocations.asp">
											        <td width="100%" align="center">
                                                        Storage Location
											        </td></a>
										        </tr>
									        </table>
								        </td>
								        <td class="TabCell" valign="bottom" align="center" width="82">
									        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										        <tr><a href="MASUOMENTRY.asp">
											        <td width="100%" align="center">
                                                        UOM
											        </td></a>
										        </tr>
									        </table>
								        </td>
								        <td class="TabCurrentCell" valign="bottom" width="100">
								        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
									        <tr>
										        <td width="100%" align="center">Packing Type
										        </td>
									        </tr>
								        </table>
							        </td>
							        <td class="TabCell" valign="bottom" align="center" width="105">
								          <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								           <tr><a href="../../NoSeries/InventoryNoSeriesEntry.asp">
									          <td width="100%" align="center">Number Series</td></a>
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
				            <td>
				                <table border="0" cellpadding="0" cellspacing="0" width="100%">
				                    <tr>
				                     <td class="TabCell" valign="bottom" align="center" width="82">
									        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										        <tr><a href="OrgPurchaseControlEntry.asp">
											        <td width="100%" align="center">
                                                        Purchase
											        </td></a>
										        </tr>
									        </table>
								        </td>
        							
								        <td class="TabCell" valign="bottom" align="center" width="82">
									        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										        <tr><a href="OrgInventoryControlEntry.asp">
											        <td width="100%" align="center">
                                                        Inventory
											        </td></a>
										        </tr>
									        </table>
								        </td>
								        <td class="TabCell" valign="bottom" align="center" width="85">
									        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										        <tr><a href="OrgInspectionControlEntry.asp">
											        <td width="100%" align="center">
                                                        Inspection
											        </td></a>
										        </tr>
									        </table>
								        </td>
								        <td class="TabCell" valign="bottom" align="center" width="60">
									        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										        <tr><a href="OrgSalesControlEntry.asp">
											        <td width="100%" align="center">
                                                        Sales
											        </td></a>
										        </tr>
									        </table>
								        </td>
								        <td class="TabCell" valign="bottom" align="center" width="105">
								          <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								           <tr><a href="OrgManufacturingControlEntry.asp">
									          <td width="100%" align="center">Manufacturing</td></a>
									        </tr>
								          </table>
								        </td>
								        <td class="TabCell" valign="bottom" align="center" width="150">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="MasSubCOntProcess.asp">
									  <td width="100%" align="center">Sub-Contract Process</td></a>
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
										<td align="center" colspan="3" class="MiddlePack" height="7">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
									</tr>
									<tr>
										<td align="center" width="5" class="ClearPixel">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
										<td valign="top" width="100%">
											<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
												<tr>
													<td>
														<div>
															<table class="CollapseBand" cellspacing="0" cellpadding="0">
																<tr>
																	<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
																		<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
																		</a>
																	</td>
																	<td valign="center" class="SubTitle">&nbsp;&nbsp;
																	</td>
																</tr>
															</table>
															<table border="0" cellpadding="0" cellspacing="0" class="BodyTable" width=100%>
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="width: 100%; display: none">
																		<table cellpadding="0" cellspacing="0">
	                                                                    <tr>
                                                                            <td class="FieldCellSub">Packing Name</td>
                                                                            <td class="FieldCellSub">
                                                                                <input type=text name=txtPackName value="" class=FormElem>
										                                    </td>
										                                    <td class="FieldCellSub">
										                                        <input type="button" name="btnGo" value="GO" class="ActionButton" onclick="CheckSubmit()">
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
										<td align="center" class="ClearPixel" width="5">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
									</tr>
									<tr>
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
						            <tr>
								        <td align="center">
								        </td>
								        <td valign="top" width="100%" align="left">
                                            <table border="0" cellpadding="0" cellspacing="0" width=100%>
                                                <tr>
											        <td>
												        <DIV class=frmBody style="width: 100%; height:300;">
													        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
														        <tr>
															        <td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
															        <td class="ExcelHeaderCell" align="center" ><img src="../../assets/images/iTMS%20icons/DeleteIcon.gif" onclick="DelItem()"></td>
															        <td class="ExcelHeaderCell" align="center" >Packing Name</td>
															        <td class="ExcelHeaderCell" align="center" >Numbering Type</td>
															        <td class="ExcelHeaderCell" align="center" >Alternate Label</td>
															        <td class="ExcelHeaderCell" align="center" >No of Level</td>
														        </tr>
														        
														    <%
														        sQuery = " Select PackingCode,PackingName,isNull(NumberingType,'N'),isNull(AlternateName,''),isNull(NoOfSubLevels,0) from APP_M_PackingType "
														        
														        if trim(sPackingName)<>"" then
														            sQuery = sQuery & " where PackingName like '%"& sPackingName &"%'"
														        end if
														        'Response.Write sQuery
														        rsObj.open sQuery,con 
														        If not rsObj.EOF then
														            iSlNo = 0
   												                    Do while Not rsObj.EOF
   													                    iSlNo = iSlNo + 1
   																		
   																		If rsObj(2)= "N" Then
   																			sNumberingType = "None"
   																		Elseif rsObj(2) ="L" Then
   																			sNumberingType = "Lot"
   																		Elseif rsObj(2) ="LS" Then
   																			sNumberingType = "Lot & Serial"
   																		Elseif rsObj(2) ="S" Then
   																			sNumberingType = "Serial"
   																		End IF
   													                    %>
   													                    <tr>
   														                    <td class="ExcelSerial" align="center" ><%=iSlNo%></td>
   														                    <td class="ExcelDisplayCell" align="center">
   																				<input type="checkbox" name="chkBox<%=iSlNo%>" value="<%=rsObj(0)%>">
   														                    </td>
   														                    <td class="ExcelDisplayCell" align="Left"><%=Trim(rsObj(1))%></td>
   														                    <td class="ExcelDisplayCell" align="Left"><%=sNumberingType%></td>
   														                    <td class="ExcelDisplayCell" align="Left"><%=Trim(rsObj(3))%></td>
   														                    <td class="ExcelDisplayCell" align="Left"><%=Trim(rsObj(4))%></td>
   													                    </tr>
   													                    <%
   													                    rsObj.MoveNext
												                    Loop
											                    End If 'If not dcrs.EOF then
											                    rsObj.Close %>
											                    <Input type="Hidden" Name="hCnt" value="<%=iSlNo%>">
												            </table>
												        </div>
									                </td>
									            </tr>
									            <tr>
							                        <td>
                                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
					                                    <tr>
					                                    <td valign="middle" class="ActionCell">
					                                    <p align="center">								 
						                                    <input type="button" value="Create" name="btnCreate" class="ActionButton" onclick="ShowPackingType('C')"> 
						                                    <input type="button" value="Edit" name="btnEdit" class="ActionButton" onclick="ShowPackingType('E')"> 
						                                    <input type="button" value="Lot Entry" name="btnCreate" class="ActionButtonX" onclick="ShowPackingForEntry()">
						                                    <input type="button" value="Lot Entry Static" name="btnLotEntry" class="ActionButtonX" onclick="ShowPackingForEntryStatic()">  
					                                    </td>
					                                    </tr>
					                                    </table>
					                                </td>
                                                 </tr>
									        </table>
								        </td>
								        <td align="center"></td>
                                    </tr>
                            		<tr>
										<td align="center" class="MiddlePack" colspan="3"></td>
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

