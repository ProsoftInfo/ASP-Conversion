<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasUOMEntry.asp	
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 16, 2002
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	masUOMInsert.asp
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
<%
	Dim dcrs,iSNo
	Dim sMode,sDecimal,sUOMSH,sUOMDesc,sUomCode,sAppCode
	
	
	set dcrs = Server.CreateObject("ADODB.Recordset")
	
	sAppCode = Request.QueryString("APPCODE") ' note : 6 means calling from Production
	
	'Response.Write "<p> sAppCode = " & sAppCode
	
	sMode = Request.Form("hMode")
	sUomCode = Request.QueryString("UOMCODE")
	
	if sMode="" or IsNull(sMode) then sMode ="S"
	
	if sMode="E" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source ="Select UoMCode,UoMDescription,DecimalAllowed from MS_UNITOFMEASUREMENT where UoMCode = '"& trim(sUoMCode) &"'"
			.Open 
		end with
		if not dcrs.EOF then
			sUOMSH = dcrs(0)
			sUOMDesc = dcrs(1)
			sDecimal = dcrs(2)
		end if
		dcrs.Close 
	end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Unit Of Measurement</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/masUOMCreate.js"></SCRIPT>
<script>
function Help() {
	window.open("../HelpFiles/UnitofMeasurement.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
}
</script>
<SCRIPT>
function openDetails() {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog("XMLUoMView.asp", "UoM", "dialogHeight:310px;dialogWidth:320px;center:Yes;help:No;resizable:No;status:No");
	} else {
		window.open("XMLUoMView.asp", "_blank", "height=310,width=320,resizable=no,status=no");
	}
}

function DelItem(sVal) {
	var request = new XMLHttpRequest();
	request.open("POST", "INVMASUOMDelete.asp?sTemp=" + encodeURIComponent(sVal), false);
	request.send(null);
	if (String(request.responseText || "").replace(/^\s+|\s+$/g, "") !== "") {
		alert(request.responseText);
	} else {
		alert("UOM Deleted");
		document.formname.hMode.value = "S";
		document.formname.action = "MASUOMENTRY.ASP?APPCODE=" + encodeURIComponent(document.formname.hdAppCode.value);
		document.formname.submit();
	}
}

function EditItem(sVal) {
	document.formname.hMode.value = "E";
	document.formname.action = "MASUOMENTRY.ASP?UOMCODE=" + encodeURIComponent(sVal);
	document.formname.submit();
}

function UpdateDet() {
	var request = new XMLHttpRequest();
	var decimalAllowed = document.formname.radDecimal[0].checked ? "Y" : "N";
	var tempValue = document.formname.txtUOMCode.value + ":" + document.formname.txtUOMCode.value + ":" + document.formname.txtUOMName.value + ":" + decimalAllowed;
	request.open("POST", "INVMASUOMUpdate.asp?sTemp=" + encodeURIComponent(tempValue), false);
	request.send(null);
	if (String(request.responseText || "").replace(/^\s+|\s+$/g, "") !== "") {
		alert(request.responseText);
	} else {
		alert("UOM Updated");
		document.formname.hMode.value = "S";
		document.formname.action = "MASUOMENTRY.ASP";
		document.formname.submit();
	}
}
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hMode" value="0">
<input type=hidden name="hdAppCode" value="<%=sAppCode%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" height="20">
			<table>
	            <tr>
	                <td class="PageTitle" >
	                    <p align="center">Master Creation
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
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgControlDefn.asp">
											<td width="100%" align="center">
                                                Org Defn
											</td></a>
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
								<td class="TabCurrentCell" valign="bottom" align="center" width="50">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">UoM
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
							  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
							   <tr><a href="PackingTypes.asp">
								  <td width="100%" align="center">Packing Type</td></a>
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
									<table border="0" cellpadding="0" cellspacing="0" width="20" class="TabTableEnd">
										<tr>
											<td width="100%" valign="bottom">
												<p align="center"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font></p>
											</td>
										</tr>
									</table>
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
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <!--<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellspacing="0"  cellpadding="0" class="ToolBarTable">
										<tr>
											<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
				       							
				       							<span style="cursor: hand" Title="Exisiting UoM" onclick="openDetails()">
              									<p align="center"><font face="Wingdings" size="5">4</font>
												</span>
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
                            </tr>-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td>
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td class=FieldCell> UoM Short Description</td>
														<td class='FieldCellSub'><input type="text" name="txtUOMCode" size="5" maxlength=3 class="Formelem" value="<%=sUOMSH%>"></td>
														<input type=hidden name="hUoMCode" value="<%=sUomCode%>">
													</tr>
													<tr>
														<td class=FieldCell> UoM Description</td>
														<td class='FieldCellSub'><input type="text" name="txtUOMName" size="50" maxlength=40 class="Formelem" value="<%=sUOMDesc%>"></td>
													</tr>
													<tr>
														<td class=FieldCell> Decimals Allowed</td>
														<td class='FieldCellSub'>
															<input type="radio" name="radDecimal" value="Y" class="Formelem" <%if trim(sDecimal)="Y" then Response.Write "checked" %>> Yes 
															<input type="radio" name="radDecimal" value="N" class="Formelem" <%if trim(sDecimal)="N" then Response.Write "checked" %>> No 
														</td>
													</tr>
												</table>
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
																<%if sMode="S" then%>
																	<input type="button" value="Save" name="B1" class="ActionButton" onClick="javascript:checkSubmit()">
																<%else%>
																	<input type="button" value="Update" name="B1" class="ActionButton" onClick="UpdateDet()">
																<%end if%>
																<input type="reset" value="Reset" name="B2" class="ActionButton">
																<%if trim(sAppCode) = "6" then 'calling from Production %>
																	<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('../../Production/welcome_Production.asp')">
																<%else%>	
																	<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
																<%end if%>	
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr></tr>
										<tr></tr>
										<tr></tr>
										<tr>
											<td width="100%">
												<table border="0" cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
													<tr>
														<td class="ExcelHeaderCell" align="center">S.No.</td>
														<td class="ExcelHeaderCell" align="center">Edit/Delete</td>
														<td class="ExcelHeaderCell">UOM Description</td>
														<td class="ExcelHeaderCell">Decimal Allowed</td>
													</tr>
													<%
														with dcrs
															.CursorLocation = 3
															.CursorType = 3
															.ActiveConnection = con
															.Source = "Select UoMCode,UoMDescription,DecimalAllowed from MS_UNITOFMEASUREMENT"
															.Open 
														end with
														if not dcrs.Eof then 
															iSNo = 0
															do while not dcrs.EOF 
																iSNo = iSNo + 1
															%>
																<tr>
																	<td class="ExcelDisplayCell" align="center"><%=iSNo%></td>
																	<td class="ExcelDisplayCell" align="center">
																		<input type=button name="btnEditZ<%=iSNo%>" value="Edit" onClick="EditItem('<%=dcrs(0)%>')" class="ActionButtonX">
																		<input type=button name="btnDelZ<%=iSNo%>" value="Delete" onClick="DelItem('<%=dcrs(0)%>')"  class="ActionButtonX">
																	</td>
																	<td class="ExcelDisplayCell"><%=dcrs(1)%></td>
																	<td class="ExcelDisplayCell">
																	<%
																		if trim(dcrs(2))="N" then
																			Response.Write "No"
																		else
																			Response.Write "Yes"
																		end if
																	%>
																	</td>
																</tr>
															<%
																dcrs.MoveNext 
															loop
														end if
														dcrs.Close 
													%>
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
</HTML>
