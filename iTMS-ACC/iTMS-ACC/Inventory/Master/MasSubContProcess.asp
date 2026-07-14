<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasSubContProcess.asp	
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Oct 29,2013
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	MasSubContProcessInsert.asp
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
<%
Dim dcrs
Dim sSubProDesc,sSubProName,sSubContID,sMode,sOrgCode,sQuery
Dim iSNo
set dcrs = Server.CreateObject("ADODB.Recordset")

sMode = Request("Mode")
sSubContID = Request("SubContID")
if trim(sSubContID)<>"" then
    sQuery = "Select SubConProcessName,SubConProcessDesc,OrganisationCode from APP_M_SubContractProcess where SubConProcessID = "& sSubContID
    dcrs.open sQuery,con
    if not dcrs.eof then
        sSubProName = dcrs(0)
        sSubProDesc = dcrs(1)
        sOrgCode = dcrs(2)
    end if 
    dcrs.close
end if 'if trim(sSubContID)<>"" then

if trim(sOrgCode)="" or IsNull(sOrgCode) then sOrgCode = session("organizationcode")
if trim(sMode)="" or IsNull(sMode) then sMode = "S"


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>SubContract Process</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/masUOMCreate.js"></SCRIPT>
<SCRIPT>
function DeleteItem() {
	var form = document.formname;
	var count = Number(form.hCnt.value || 0);
	var values = [];
	var checkbox;
	for (var i = 1; i <= count; i += 1) {
		checkbox = form.elements["chkZ" + i];
		if (checkbox && checkbox.checked) {
			values.push(checkbox.value);
		}
	}
	form.hSubContID.value = values.join(",");
	form.action = "MasSubContProcessInsert.asp?Mode=D";
	form.submit();
}

function ViewDetails(iSubContID) {
	document.formname.action = "MasSubContProcess.asp?SubContID=" + iSubContID + "&Mode=E";
	document.formname.submit();
}

function UpdateDet() {
	document.formname.action = "MasSubContProcessInsert.asp?Mode=E";
	document.formname.submit();
}

function SaveDet() {
	document.formname.action = "MasSubContProcessInsert.asp";
	document.formname.submit();
}

function DeleteDet() {
	document.formname.action = "MasSubContProcessInsert.asp";
	document.formname.submit();
}
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
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
								<td class="TabCell" valign="bottom" align="center" width="105">
							  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
							   <tr><a href="MASUOMENTRY.asp">
								  <td width="100%" align="center">UoM</td></a>
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
								<td class="TabCurrentCell" valign="bottom" align="center" width="150">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Sub-Contract Process
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
								</td>
                            </tr>
                           
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
														<td class="FieldCellsub">Process Name</td>
														<td class="FieldCellSub"><input type="text" name="txtSubContName" size="50" maxlength=50 class="Formelem" value="<%=sSubProName%>"></td>
														<input type="hidden" name="hSubContID" value="<%=sSubContID%>">
														<input type="hidden" name="hOrgCode" value="<%=sOrgCode%>">
													</tr>
													<tr>
														<td class="FieldCellSub">Process Description</td>
														<td class="FieldCellSub"><input type="text" name="txtSubContDesc" size="50" maxlength=50 class="Formelem" value="<%=sSubProDesc%>"></td>
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
																	<input type="button" value="Save" name="B1" class="ActionButton" onClick="SaveDet()">
																<%else%>
																	<input type="button" value="Update" name="B1" class="ActionButton" onClick="UpdateDet()">
																<%end if%>
																<input type="reset" value="Reset" name="B2" class="ActionButton">
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
														<td class="ExcelHeaderCell" align="center" width="20">S.No.</td>
														<td class="ExcelHeaderCell" align="center"  width="25">
														    <img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onclick="DeleteItem()">
														</td>
														<td class="ExcelHeaderCell" align="center">Process Name</td>
														<td class="ExcelHeaderCell" align="center">Process Description</td>
													</tr>
													<%
														with dcrs
															.CursorLocation = 3
															.CursorType = 3
															.ActiveConnection = con
															.Source = "Select SubConProcessID,SubConProcessName,SubConProcessDesc from APP_M_SubContractProcess where OrganisationCode = '"& sOrgCode & "'"
															.Open 
														end with
														if not dcrs.Eof then 
															iSNo = 0
															do while not dcrs.EOF 
																iSNo = iSNo + 1
															%>
																<tr>
																	<td class="ExcelDisplayCell" align="center"  width="20"><%=iSNo%></td>
																	<td class="ExcelDisplayCell" align="center"  width="20">
																	    <input type="checkbox" name="chkZ<%=iSNo%>" value="<%=dcrs(0)%>" />
																	</td>
																	<td class="ExcelDisplayCell">
																	    <a href="#" class="ExcelDisplayLink" onclick="ViewDetails('<%=dcrs(0)%>')"><%=dcrs(1)%></a></td>
																	<td class="ExcelDisplayCell"><%=dcrs(2)%></td>
																</tr>
															<%
																dcrs.MoveNext 
															loop
														end if
														dcrs.Close 
													%>
												</table>
												<input type="hidden" name="hCnt" value="<%=iSNo%>" />
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
