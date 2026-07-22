<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmInvDet.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran R
	'Created On					:	July 16,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
	'Procedures/Functions Used	:	populateInterUnit,populateStLocation
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
	'XML DOM Variables
	Dim oDOM,Root,objfs,PGNode,rsTemp

	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")

	dim sOrgName,sClassName,sClassCode,iItmCode,sItmDescr,sOrgCode,sSTUoM,sMAUoM,sSAUoM,sPUUoM
	dim schkSal, schkPur, schkMan,sQuery
	Dim sFastMovCri,sSlowMovCri,sNonMovCri

	iItmCode = Request("ItemCode")
    sClassCode = Request("ClassCode")
    sOrgCode = Session("organizationcode")
    sOrgName = Session("OrgShortName")

if trim(iItmCode)<>"" then
    sQuery = "Select ItemDescription,(Select GroupName from INV_M_Classification where GroupCode = "&_
             " V.ClassificationCode),StoresUOM,PurchaseUOM,ManufacturingUOM,SalesUOM,PurchaseEligible,"&_
             " ManufactureEligible,SalesEligible from VwItem V where ItemCode="& iItmCode &" and ClassificationCode = "& sClassCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sItmDescr = trim(rsTemp(0))
        sClassName = trim(rsTemp(1))
        sSTUoM = trim(rsTemp(2))
        sPUUoM = trim(rsTemp(3))
		sMAUoM = trim(rsTemp(4))
		sSAUoM = trim(rsTemp(5))
		schkPur = trim(rsTemp(6))
		schkMan = trim(rsTemp(7))
		schkSal = trim(rsTemp(8))
    end if
    rsTemp.Close
end if'if trim(iItmCode)<>"" then

		sPUUoM = DisplayUoM(sPUUoM)

%>
<BODY leftMargin=0 topMargin=0 onLoad="LoadDraftedDetails(<%=iItmCode%>)">
<%
	dim dcrs4
	dim sFIFO,sLIFO,sWA,sRep,sInter,sLoc,sStock,sABC,sFSN
	dim sdisabled,sw2,sw3,sw4,sClass

	Set dcrs4 = Server.CreateObject("ADODB.RecordSet")

	with dcrs4
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ALLOWFIFOVALUATION,ALLOWLIFOVALUATION,ALLOWWAVALUATION,ALLOWREPLENISHMENT,ALLOWINTERUNITTRANSFER,ALLOWLOCATIONTRANSFER,STOCKHOLDINGPERIOD,ABCCLASSIFICATIONEXIST,FSNCLASSIFICATIONEXIST,FastMovingDays,SlowMovingDays,NonMovingDays FROM INV_CONTROL_ORGINVENTORY WHERE OUDEFINITIONID = " & Pack(sOrgCode) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs4.ActiveConnection = nothing

	if not dcrs4.EOF then
		sFIFO = trim(dcrs4(0))
		sLIFO = trim(dcrs4(1))
		sWA = trim(dcrs4(2))
		sRep = trim(dcrs4(3))
		sInter = trim(dcrs4(4))
		sLoc = trim(dcrs4(5))
		sStock = trim(dcrs4(6))
		sABC = trim(dcrs4(7))
		sFSN = trim(dcrs4(8))
		sFastMovCri = Trim(dcrs4(9))
		sSlowMovCri = Trim(dcrs4(10))
		sNonMovCri = Trim(dcrs4(11))
	else
		sFIFO = "0"
		sLIFO = "0"
		sWA = "1"
		sRep = "0"
		sInter = "0"
		sLoc = "0"
		sStock = "0"
		sABC = "0"
		sFSN = "0"
    	sFastMovCri = "0"
		sSlowMovCri = "0"
		sNonMovCri = "0"
	end if
	dcrs4.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Control Definition - Inventory</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="storageData" data-itms-xml-island="1" data-src="../xmldata/Storage.xml"></script>
<script type="application/xml" id="OutData" data-itms-xml-island="1"><Root ItemCode="<%=iItmCode%>" ClassCode="<%=sClassCode%>" OrgCode="<%=sOrgCode%>"></Root></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT>
window.ITMS_INV_DETAIL_NEXT = "ItmManufacture.asp";
window.ITMS_INV_DETAIL_BACK = "ItmDetailedDefn.asp";
window.ITMS_INV_DETAIL_REQUIRE_ITEM = true;
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemInventoryDetail.js"></SCRIPT>

</HEAD>

<form method="POST" name="formname" action="">
<INPUT TYPE=HIDDEN NAME="hClassName" VALUE="<%=sClassName%>">
<INPUT TYPE=HIDDEN NAME="hOrgName" VALUE="<%=sOrgName%>">
<INPUT TYPE=HIDDEN NAME="hItmName" VALUE="<%=sItmDescr%>">
<INPUT TYPE=HIDDEN NAME="hClassCode" VALUE="<%=sClassCode%>">
<INPUT TYPE=HIDDEN NAME="hOrgCode" VALUE="<%=sOrgCode%>">
<INPUT TYPE=HIDDEN NAME="hItmCode" VALUE="<%=iItmCode%>">
<INPUT TYPE=HIDDEN NAME="hStock" VALUE="<%=sStock%>">

<INPUT TYPE=HIDDEN NAME="hChkPur" VALUE="<%=schkPur%>">
<INPUT TYPE=HIDDEN NAME="hChkSal" VALUE="<%=schkSal%>">
<INPUT TYPE=HIDDEN NAME="hChkMan" VALUE="<%=schkMan%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Control Definition
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
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
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
											<td align="center"><a href="ItmCreationDefinitionEntry.asp">Basic</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmDetailedDefn.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=sClassCode%>">
											<td align="center">Purch. & Sales
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="80">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Inventory
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="ItmManufacture.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=sClassCode%>">
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
                                    <table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly">
									    <tr>
											<td class="FieldCellSub" width="80">Item Name</td>
											<td>
											<span class="DataOnly"><%=sItmDescr%>&nbsp;</span>
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
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
									<center>
                                        <div align="left" style="width:100%">
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td class='GroupTitleLeft' width="10">&nbsp;
                                                </td>
												<td class='GroupTitle' width="120"><p align="center">
                                                Valuation Method
                                                </td>
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
												<td class=MiddlePack colspan="4"> </td>
											</tr>
									    <tr>
												<%	if sABC = "0" then
														sdisabled = " DISABLED "
												%>
												    <INPUT TYPE=HIDDEN NAME="hABCEli" VALUE="N">
												<%	else %>
													<INPUT TYPE=HIDDEN NAME="hABCEli" VALUE="Y">
												<%		sdisabled = ""
													end if
												%>
												<td class=FieldCellSub width="100">ABC</td>
												<td class='FieldCellSub' width="55">
													<input type="radio" value="A" name="radABC" <%=sdisabled%> class="FormElem">A
												</td>
												<td class='FieldCellSub' width="80">
													<input type="radio" value="B" name="radABC" <%=sdisabled%> class="FormElem">B
												</td>
												<td class='FieldCellSub'>
													<input type="radio" value="C" name="radABC" <%=sdisabled%> class="FormElem">C
												</td>
                                        </tr>
                                        <tr>
												<%	if sFSN = "0" then
														sdisabled = " DISABLED "
												%>
												<INPUT TYPE=HIDDEN NAME="hFSNEli" VALUE="N">
												<%	else %>
													<INPUT TYPE=HIDDEN NAME="hFSNEli" VALUE="Y">
												<%		sdisabled = ""
													end if
												%>
												<td class=FieldCellSub width="100">FSN</td>
												<td class='FieldCellSub' >
													<!--<input type="radio" value="F" name="radFSN" <%=sdisabled%> class="FormElem">-->Fast
													<input type="text" name="txtFastMovCriteria" class="FormElem" value="<%=sFastMovCri%>" size=5 <%=sdisabled%>>
												</td>
												<td class='FieldCellSub' >
													<!--<input type="radio" value="S" name="radFSN" <%=sdisabled%> class="FormElem">-->Slow
													<input type="text" name="txtSlowMovCriteria" class="FormElem" value="<%=sSlowMovCri%>" size=5 <%=sdisabled%>>
												</td>
												<td class='FieldCellSub'>
													<!--<input type="radio" value="N" name="radFSN" <%=sdisabled%> class="FormElem">-->Non-Moving
													<input type="text" name="txtNonMovCriteria" class="FormElem" value="<%=sNonMovCri%>" size=5 <%=sdisabled%>>
												</td>
                                        </tr>
                                        <tr>
												<td class=FieldCellSub width="100">VED</td>
												<td class='FieldCellSub' width="55">
													<input type="radio" value="V" name="radVED" class="FormElem">Vital
												</td>
												<td class='FieldCellSub' width="80">
													<input type="radio" value="E" name="radVED" class="FormElem">Essential
												</td>
												<td class='FieldCellSub'>
													<input type="radio" value="D" name="radVED" class="FormElem">Desirable
												</td>
                                        </tr>
										<%
											if (instr(1,sFIFO,"1")) then sw2 = " CHECKED "
											if (instr(1,sLIFO,"1")) then sw3 = " CHECKED "
											if (instr(1,sWA,"1")) then sw4 = " CHECKED "
										%>
										</table>
										</center>
                                                </td>
											</tr>
										</table>
                                    </div>
                                    </center>
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
                                     <div align="left" style="width:100%">
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
        										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td class='GroupTitleLeft' width="10">&nbsp;
                                                </td>
												<td class='GroupTitle' width="120"><p align="center">
                                               Replenishment
                                                </td>
												<td class='GroupTitleRight'><p align="left">&nbsp;
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
																		<td class="FieldCell">Reorder Level</td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtReLvl" value="0" style="text-align=right" size="10" maxlength=10 class="FormElem" onkeypress="DoKeyPress('Y',10,3)">
																		</td>
																		<td class="FieldCell">Reorder Quantity</td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtReQty" value="0" style="text-align=right" size="10" maxlength=10 class="FormElem" onkeypress="DoKeyPress('Y',10,3)">
																		</td>
																		<td class="FieldCell">Economic Order Quantity</td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtEcQty" size="10" value="0" style="text-align=right" maxlength=10 class="FormElem" onkeypress="DoKeyPress('Y',10,3)">
																		</td>
																	</tr>
																</table>
															</td>
														</tr>
													</table>
									            </td>
											</tr>
										</table>
									</div>
                                </center>
                            </td>
                            <td align="center">
                            </td>
                        </tr>
                         <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                        </tr>
                        <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
								    <table border=0 cellpadding=0 cellspacing =0>
								        <tr>
								            <td class="FieldCellSub">Stock Reserved
								            </td>
								            <td class="FieldCellSub"><span id="txtStkReserved" class="DataOnly">&nbsp;</span>
								            </td>
								        </tr>
								    </table>
                        		</td>
								<td align="center">
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
                                                    <input type="button" value="Back" name="B5" class="ActionButton" onClick="CheckBack()" >
                                                    <input type="button" value="Save" name="B4" class="ActionButton" onClick="CheckSubmit()" >
													<input type="button" value="Cancel" name="B1" class="ActionButton">
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
	con.close
	set con = nothing
%>

<%
	' Function to populate UoM
	Function DisplayUoM(sUoM)
		' Declaration of variables
		Dim dcrs,sUoMDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = '" & sUoM & "'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMDesc = dcrs(0)

		if Not dcrs.EOF then
			DisplayUoM = sUoMDesc
		else
			DisplayUoM = "N/A"
		end if
		dcrs.Close
	End Function
%>
