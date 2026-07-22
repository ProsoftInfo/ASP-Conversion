<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MRGenerationEntry.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 27, 2005
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Material Requisition Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData"></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData"></script>
<script type="application/xml" data-itms-xml-island="1" id="UoMData" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><ROOT></ROOT></script>
<script type="application/xml" data-itms-xml-island="1" id="OutSelectData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutCost"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTypeData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root/></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></script>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/TempItem.js"></script>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<script language="javascript" src="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrEntryModern.js"></SCRIPT>
</head>
<%
	dim sUnit,sIType,sCreatedBy,iuserid,rsUser
	dim sFinPeriod,Arr,dFrmDate,dToDate,sQuery,dcrs

	sUnit = Request.QueryString("sOrg")
	sIType = Request.QueryString("sIType")

	if trim(sUnit)="" or IsNull(sUnit) then sUnit = session("organizationcode")

	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	dFrmDate = "01/04/"& Arr(0)
	dToDate = "31/03/"& Arr(1)

	iuserid = Session("userid")
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	set dcrs = Server.CreateObject("ADODB.Recordset")

    sCreatedBy = session("username")


%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="popCC();Init()">
<form method="POST" name="formname">
	<input type=hidden name="hCreatedBy" value="<%=Session("userid")%>">
	<input type=hidden name="hFrmDate" value="<%=dFrmDate%>">
	<input type=hidden name="hToDate" value="<%=dToDate%>">
	<input type=hidden name="hUnit" value="<%=sUnit%>">
	<input type=hidden name="hItemType" value="">
	<input type=hidden name="hIssTo" value="">
	<input type=hidden name="hIssForType" value="">
	<input type=hidden name="hPartyCode" value="">
	<input type=hidden name="hRefType" value="">
	<input type=hidden name="hRefNo" value="">
	<input type=hidden name="hRefDate" value="">
	<input type="hidden" name="hRequestedByUnit" value="<%=sUnit%>">

	<input type="hidden" name="hIssueToType" value="">
	<input type="hidden" name="hIssueToCode" value="">
	<input type="hidden" name="hIssueToSubCode" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Material Requisition
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				    		    <tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" >
								<tr>
								   	<td class="TabCell" valign="bottom" align="center" width="50">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="MRSMGMTList.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
									<td class="TabCurrentCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="MRGENERATIONENTRY.asp">
												<td align="center">Basic
												</td></a>
											</tr>
										</table>
									</td>
									<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="MRApprovalEntry.asp">
											    <td align="center">Edit/Approval
											    </td></a>
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
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center">
									</td>
									<td width="100%" colspan="2">
										<div align="left">
											<table border="0" cellspacing="0" cellpadding="0" width="100%">
											    <tr>
											        <td class="FieldCellSub" style="width:125px" >Requested By</td>
													<td class="FieldCellSub">
													    <select name="selIssueTo" class="FormElem" onChange="popIssueTo()">
													        <option value="0">Select</option>
													        <%
														        populateIssueToSel(sUnit)
													        %>
													        </select>
													    <span id="txtParty" class="DataOnly"></span>
													</td>
													<td class="FieldCellSub">MR Date</td>
													<td class="FieldCellSub" valign="middle">
														<object id="ctlCDDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"    codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
															<param name="_ExtentX" value="2355">
															<param name="_ExtentY" value="529">
														</object>
													</td>
											    </tr>
												<tr>
                                                    <td class="FieldCellSub" style="width:125px">Reference Name</td>
													<td class="FieldCellSub">
														<select name="selRefName" class="FormElem">
														<%
														    RefTypePop 2,4
														%>
														</select>
													    <a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click Here to Edit Usage Information" width="11" height="11" onClick="GetDetails()"></a>
													</td>
													<td class="FieldCellSub">Requested For</td>
                                                    <td class="FieldCellSub">
                                                        <select id="cmbIssType" class="FormElem">
                                                            <option value="SEL">Select</option>
                                                            <%
                                                                sQuery = "Select ReceiptIssueTypeCode,ReceiptIssueTypeDesc from APP_M_ReceiptIssueTypes where ApplicableFor in ('B','I')"
                                                                dcrs.open sQuery,con
                                                                if not dcrs.eof then
                                                                    do while not dcrs.eof
                                                                            response.write "<option value="& trim(dcrs(0)) &">"& trim(dcrs(1)) &"</option>"
                                                                        dcrs.movenext
                                                                    loop
                                                                end if
                                                                dcrs.close
                                                            %>
                                                        </select>&nbsp;
                                                    </td>
												</tr>
                                                   <tr>
                                                    <td class="FieldCellSub" style="width:125px">Reference No - Date</td>
													<td class="FieldCellSub">

														<span class="DataOnly" align=center id="RefNoDate">NA</span>
												    </td>
												    <td class="FieldCellSub">Created By</td>
													<td class="FieldCellSub">
														<span class="dataonly"><%=sCreatedBy%></span>
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Acc. Head</td>
													<td class="FieldCellSub">
														<select size="1" name="selAccHead" class="FormElem" onChange="CreateNew(this)">
															<option value="select">Select</option>
															<option value="NEW">< NEW ></option>
														</select>
														<!--Div to display the Consumption head mapping pop up layer-->
														<div class=frmbody id="idConsumption" style="Z-INDEX: 1; POSITION: absolute" style="width=350;display: none" >
															<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
															<tr>
																<td valign="top">
																	<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
																		<TR>
																			<TD class="TabBodyWithTopLine">
																				<table border="0" cellpadding="0" cellspacing="0" width="100%">
																					<tr><td height="1px"></td></tr>
																					<tr>
																						<td colspan=2 class="PageTitle">Consumption Head Configuration</td>
																					</tr>
																					<tr>
																						<td align="center"></td>
																						<td width="100%">
																							<table border="0" cellpadding="0" cellspacing="0">
																								<tr>
																									<td class="FieldCell">Consumption Head</td>
																									<td class="FieldCellSub">
																										<input type="text" name="txtCHead" size="20" maxlength=50 class="FormElem">
																									</td>
																								</tr>
																								<tr>
																									<td class="FieldCell" valign="top">Account Head</td>
																									<td class="FieldCellSub">
																										<select size="1" name="selAcc" class="FormElem">
																											<option value="select">Select</option>
																											<%	'Calling the Function which populates Account Head List
																												populateAccountHead
																											%>
																										</select>
																										<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
																									</td>
																								</tr>
																							</table>
																						</td>
																						<td align="center"></td>
																					</tr>
																					<tr>
																						<td align="center" colspan="3" class="MiddlePack"></td>
																					</tr>
																					<tr>
																						<td align="center"></td>
																						<td width="100%">
																							<div class="frmbody" id="frm2" style="width: 100%; height:130;">
																								<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
																									<tr>
																										<td class="ExcelHeaderCell" align="center">Consumption</td>
																										<td class="ExcelHeaderCell" align="center" width="100">Account Head</td>
																									</tr>
																								</table>
																							</div>
																						</td>
																						<td align="center"></td>
																					</tr>

																					<tr>
																						<td align="center" colspan="3" class="MiddlePack"></td>
																					</tr>

																					<tr>
																						<td align="center">
																						    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																						</td>
																						<td valign="top">
																							<table border="0" cellpadding="0" cellspacing="0" width="100%">
																								<tr>
																									<td valign="middle" class="ActionCell">
																										    <input type="button" value="Done" name="B1" class="ActionButton" onClick="PopDone()">
																										    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="hideDiv()">
																									</td>
																								</tr>
																							</table>
																						</td>
																						<td align="center">
																						    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																						</td>
																					</tr>
																					<tr>
																						<td align="center" colspan="3" class="BottomPack"></td>
																					</tr>
																				</table>
																			</td>
																		</tr>
																	</table>
																</td>
															</tr>
															</table>
														</div>
													</td>
													<td class="FieldCellSub" style="width:125px">Cost Center</td>
													<td class="FieldCellSub" valign="top">
														<select size="1" name="selCC" class="FormElem">
															<option value="select">Select</option>
														<%	'Calling the Function which populates Cost Center List
															populateCostCenter
														%>
														</select>
													</td>

													<!--<td class="FieldCellSub">Type</td>
													<td class="FieldCell" valign="top">
														<input type="checkbox" name="chkReqType" class="FormElem" value="Returnable">Returnable
													</td>-->
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack"></td>
								</tr>

								<tr>
									<td align="center"></td>
									<td width="100%" colspan="2">
										<div class="frmBody" id="frm1" style="width: 100%; height:280;">
											<table border="0" cellspacing="1" class="ExcelTable" width="100%" id=tblLot>
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">
														<p align="center">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center">
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" alt="Delete's the Selected Item (s)" height="15" onClick="DeleteItems()"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" >
														Item Description
														<!--<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item (s)" width="11" height="11" onClick="GetItems('<%=FormatDate(date)%>')"></a>-->
													</td>
													<td class="ExcelHeaderCell" align="center">Quantity</td>
													<td class="ExcelHeaderCell" align="center" >UoM</td>
													<td class="ExcelHeaderCell" align="center" >Required By</td>
													<!--<td class="ExcelHeaderCell" align="center" >Add Spec</td>-->
													<td class="ExcelHeaderCell" align="center" >Stock</td>
													<td class="ExcelHeaderCell" align="center" >Remarks</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
								    <td align="center"></td>
									<td class="FieldCellSub"> Approver</td>
									<td class="FieldCellSub">
                                      <select size="1" name="selApprover" class="FormElem">
											<option value="0">Select</option>
											<option value="IM">Immediate Approver</option>
											<%	'Calling the Function which populates the User list
												populateEmployee
											%>
										</select>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
								    <td align="center"></td>
									<td class="FieldCellSub">Remarks</td>
									<td class="FieldCellSub">
										<textarea name="txtRemarks" cols="100" class="FormElem"></textarea>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" colspan="2">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<input type="button" value="Save" name="BtnSubmit" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="BottomPack">
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
	' Function to populate Usage
	Function populateUsage()
		' Declaration of variables
		Dim dcrs,sUsageCode,sUsageDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR WHERE ISSUEDFORCODE <> 'INV' ORDER BY ISSUEDFORCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUsageCode = dcrs(0)
		set sUsageDesc = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUsageCode)&""">"&trim(sUsageDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Cost Center list
	Function populateCostCenter()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT COSTCENTERHEAD,CCACCOUNTDESCRIPTION FROM VWORGCOSTCENTER WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND USEABLE = 1 ORDER BY COSTCENTERHEAD"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>

<%
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTDESCRIPTION,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sUnit) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(2)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function

	Function Issue()
		MsgBox "ok"
	End Function
%>
