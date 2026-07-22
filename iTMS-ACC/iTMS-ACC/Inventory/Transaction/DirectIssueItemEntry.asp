<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	DirectIssueItemEntry.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	MAHESWARI
	'Created On					:	March 21, 2008
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	Sep 13,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsIssueInsert.asp
	'Procedures/Functions Used	:	populateStore
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
<!-- #include File="../../include/ItemDisplay.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!--#include file="../../include/CommonFunctions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>MR Issue - Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%

	Dim oDom,objfs,Root,HeaderNode,iMRSNo,newElem,sUnit
	Dim arrSSelected,arrSSelectedName,iCtr,arrItemClass
	dim dcrs,dcrs1

	dim sOrgName,dMRSDate,sIssue,sItmTypeName,sUsage,sItmType,sOrgID
	dim sReqType,sUsageCode,sITypeName,arrSchTemp,sSchTemp,sSchTempValue
	dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore,sItemName
	dim arrQty,iQtyReq,iQtyIssued,iQtyPending,iQtyAppr,iQtyTrans,iQtyPur
	dim iUnitQty,iOthUnitQty,iMarkQty
	dim iQtyRes,iQtyOnHold,iQtyRej
	dim arrUoM,sUoMDesc,sUoMCode,sRecBy
	dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin
	dim iuserId,sCreatedBy,rsUser
	dim sFinPeriod,arr,sMaxDate,sMinDate,sAutoConsumption,sQuery

	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	sMinDate = "01/04/"& Arr(0)
	sMaxDate = "31/03/"& Arr(1)

	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")

	if len(Month(date())) = 1 then
		sTempMonYr = "0"&Month(date())
	else
		sTempMonYr = Month(date())
	end if
	sMonYr = sTempMonYr&Year(date())

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	sUnit = Request("hUnit")
'	Response.Write "sUnit="&sUnit
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID= '"& sUnit&"' "
		.ActiveConnection = con
		.Open
	end with
	if not dcrs.EOF then
		sOrgName = dcrs(0)
	end if
	dcrs.Close
	sUsage = Request.Form("selUsage")
	sIssue = Request.Form("selReqType")
	sUsageCode = Request.Form("selIssType")
	
	if sIssue = "0" then
		sIssue = "Returnable"
	else
		sIssue = "Non Returnable"
	end if
	Set Root = oDOM.createElement("MRSApproval")
	oDOM.appendChild Root
	Set newElem = oDOM.createElement("MRSHeader")
	newElem.setAttribute "MRSNO",""
	newElem.setAttribute "MRSDATE",""
	newElem.setAttribute "ORGID",sUnit
	newElem.setAttribute "ORGNAME", sOrgName
	newElem.setAttribute "REQTYPE", sIssue

	newElem.setAttribute "USAGE", sUsageCode
	newElem.setAttribute "USAGENAME", sUsage
	newElem.setAttribute "ITYPE", ""
	newElem.setAttribute "ITYPENAME",""

	Root.appendChild newElem
	oDOM.Save server.MapPath("../temp/transaction/MRSIssue"&Session.SessionID&".xml")
	
	sQuery = "Select isNull(AutomaticConsumptionEntry,'N') from APP_M_ApplicationSetup"
	dcrs.open sQuery,con
	if not dcrs.eof then
	    sAutoConsumption = trim(dcrs(0))
	end if
	dcrs.close

	iuserid = Session("userid")
	Set rsUser = Server.CreateObject("ADODB.RecordSet")

	'To get User name
	with rsUser
		.CursorLocation = 3
		.CursorType = 3
		.Source ="Select isNull(UserName,'') from DCS_User where EmployeeNumber = "&iuserid&" "
		.ActiveConnection = con
		.Open
	end with
	if not rsUser.EOF then
		sCreatedBy = trim(rsUser(0))
	end if
	rsUSer.Close
	if objfs.FileExists(Server.MapPath("../temp/transaction/MRS"&iMRSNo&".xml")) then

%>
<script type="application/xml" data-itms-xml-island="1" id="Data" data-src="<%="../temp/transaction/MRS"&iMRSNo&".xml"%>"></script>
<%	else %>
<script type="application/xml" data-itms-xml-island="1" id="Data"><root/></script>
<%	end if %>
<script type="application/xml" data-itms-xml-island="1" id="OutData1"><root/></script> <%'src="<%="../temp/transaction/MRISSUEDETAILS"&Session.SessionID&".xml"%>
<script type="application/xml" data-itms-xml-island="1" id="OutData2"><root></root></script>
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData"></script>
<script type="application/xml" data-itms-xml-island="1" id="UoMData" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutSelectData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData"></script>
<script type="application/xml" data-itms-xml-island="1" id="NewData"></script>
<script type="application/xml" data-itms-xml-island="1" id="RefData" data-src="<% Response.Write("../temp/transaction/UsageSelection"&Session.SessionID&".xml")%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="POrder"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="POConfirm"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="SalesInvoice"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="GatePass"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="RefBasedItem"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTypeData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ConfData"><Root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsIssueItemDetails.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="LoadData();SetDate('<%=FormatDate(Date())%>');init();">

<form method="POST" name="formname">
<!--OBJECT id=penDet type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" VIEWASTEXT>
<PARAM name="Command" value="HH Version"-->
</OBJECT>
<input type=hidden name="hMRSNo" value="<%=iMRSNo%>">
<input type=hidden name="hReqDate" value="<%=sMaxDate%>">
<input type=hidden name="hMinDate" Value="<%=sMinDate%>">
<input type=hidden name="hMaxDate" Value="<%=sMaxDate%>">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hItemType" value="">
<input type=hidden name="hUserId" Value="<%=iuserId%>">
<input type=hidden name="hISSFORTYPE" value="">
<input type="hidden" name="mrs" value="<%=iMRSNo%>">
<input type="hidden" name="sAct" value="mrsIssueItemEntry.asp">
<input type="hidden" name="hUsage" value="<%=sUsageCode%>">
<input type="hidden" name="hUnit" value="<%=sUnit%>">
<input type="hidden" name="hCtr" value="">
<input type="hidden" name="hRefCodes" value="">
<input type="hidden" name="hSupplier" value="">
<input type="hidden" name="hJobWorkNo" value="">
<input type="hidden" name="hSubCon" value="">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hRefType" value="">
<input type="hidden" name="hDocType" value="">
<input type=hidden name="hRefNo" value="" >
<input type=hidden name="hRefDate" value="" >
<input type=hidden name="hAutoConsumption" value="<%=sAutoConsumption%>">
<input type=hidden name="hIssueToCode" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Material Issue Details
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
					<TD class="TabBodywithtopline">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                                    <td class="FieldCellSub">Reference Name</td>
													<td class="FieldCellSub">
														<select name="selRefName" class="FormElem">
														<%
														    RefTypePop 2,4
														%>
														</select>
													<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click Here to Edit Usage Information" width="11" height="11" onClick="GetDetails()"></a>
													</td>
												</td>
                                                    <td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Issue Date</td>
													<td class="FieldCellSub" valign="middle">
														<object id="ctlIssDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"     codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
															<param name="_ExtentX" value="2355">
															<param name="_ExtentY" value="529">
														</object>
													</td>
												</tr>
                                                   <tr>
                                                    <td class="FieldCellSub">Reference No - Date</td>
													<td class="FieldCellSub">

														<span class="DataOnly" align=center id="RefNoDate">NA</span>


													<!--<span class="DataOnly">N/A&nbsp;</span>-->
												</td>
													<td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Created By</td>
														<td class="FieldCellSub">
															<span class="dataonly"><%=sCreatedBy%></span>
														</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Issue For</td>
													<td class="FieldCellSub" valign="top">
														<span id="UsageName" class="DataOnly"></span>

													</td>
													<td class="FieldCellSub" width="2"></td>
													<td class="FieldCellSub" width="75">Acc. Head</td>
													<td class="FieldCellSub">
														<select size="1" name="selAccHead" class="FormElem" onChange="CreateNew(this)">
															<option value="select">Select</option>
															<option value="NEW">< NEW ></option>
														</select>

<DIV class=frmBody id=idConsumption style="Z-INDEX: 1; POSITION: absolute" style="width=350;display: none" >
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td colspan=2 align="center" class=PageTitle height="20"><p align="center">Consumption Head Configuration</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Usage of Item</td>
											<td class="FieldCellSub">
												<span class="DataOnly" id="idUsage">&nbsp;</span>
											</td>
										</tr>
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
									<div class="frmBody" id="frm2" style="width: 100%; height:100;">
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
												<p align="center">
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
												</tr>

												<tr>
													<td class="FieldCellSub">Issue To</td>
													<td class="FieldCellSub" valign="top">
														<span id="selIssueFor" class="DataOnly"></span>
													</td>
													<td class="FieldCellSub" width="2"></td>
													<td class="FieldCellSub">Cost Center</td>
													<td class="FieldCellSub" valign="top">
														<select size="1" name="selCC" class="FormElem">
															<option value="select">Select</option>
														<%	'Calling the Function which populates Cost Center List
															populateCostCenter
														%>
														</select>
													</td>
												</tr>

										    <td class="FieldCellSub">Received By</td>&nbsp;
											<td class="FieldCellSub">
												<input type="text" name="txtRecBy" size="35" class="FormElem" maxlength=35 style="text-align:left">
											</td>
											<td></td>
											<td class="FieldCellSub">Type</td>
											<td class="FieldCellSub" valign="top">
												<input type="checkbox" name="chkReqType" class="FormElem" value="Returnable">Returnable
											</td>
										</tr>
										<tr>
										    <td class="FieldCellsub">Remarks</td>

										    <td class="FieldCellSub" colspan="4">
										    <textarea name="Remarks" cols="90" class="Formelem" maxlength="100"></textarea>

										    <!--td class="FieldCell">Issue Date</td>
										    <td class="ExcelInutCell">
											  <%'Response.Write InsertDatePicker("ctlIssDate")%>
											</td-->
										<!--	<td class="FieldCellSub"></td>
											<td class="FieldCellSub">Issue Type</td>
											<td class="FieldCellSub">
												<input type="checkbox" name="selIssType" class="FormElem" value="Marked">Marked
											</td>						-->
										</tr>

                                    </table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel"></td>
								<td valign="top" class="FieldCell" width="100%"><center>
                                    <div align="left">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;</td>
															<td class='GroupTitle' width="50"><p align="center">Items</td></center>
															<td class='GroupTitleRight'><p align="left">&nbsp;</td>
														</tr>
													</table>
                                                </td>
											</tr>
											<tr>
												<td class=GroupTable><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="3"> </td>
														</tr>
														<tr>
															<td class=ClearPixel width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
															<td class=FieldCell>
																<DIV class=frmBody id=frm2 style="height:260;">
																	<table  id="tblLot" border="0" cellspacing="1" class="ExcelTable" width="100%">
																		<tr>
																			<td class="ExcelHeaderCell" align="center" width="10" rowspan="3">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="3">Item Description
<!--																			<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item (s)" width="11" height="11" onClick="GetItems('<%=FormatDate(date)%>')"></a>-->
																			</td>
																			<!--td class="ExcelHeaderCell" align="center" rowspan="3">Store</td-->
																			<td class="ExcelHeaderCell" align="center" colspan="3">Quantity Availability</td>
																			<td class="ExcelHeaderCell" align="center" colspan="2">Stock</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="3">Additional<br> Details</td>
																		</tr>
																	    <tr>

																			<td class="ExcelHeaderCell" align="center">In Unit</td>
																			<td class="ExcelHeaderCell" align="center">Reserved</td>
																			<td class="ExcelHeaderCell" align="center">Transit</td>
																			<td class="ExcelHeaderCell" align="center">Issue</td>
																			<!--td class="ExcelHeaderCell" align="center">Transfer</td>
																			<td class="ExcelHeaderCell" align="center">Purchase</td-->
																			<td class="ExcelHeaderCell" align="center">Total</td>
																	    </tr>
																		<tr>

																			<td class="ExcelHeaderCell" align="center">Other Unit</td>
																			<td class="ExcelHeaderCell" align="center">On Hold</td>
																			<td class="ExcelHeaderCell" align="center">Rejected</td>
																			<!--td class="ExcelHeaderCell" align="center">Date</td>
																			<td class="ExcelHeaderCell" align="center">By Date</td-->
																			<td class="ExcelHeaderCell" align="center">By Date</td>
																			<td class="ExcelHeaderCell" align="center">UoM</td>
																		</tr>
																		<% 'refer(1---1) Row 1 %>
																		<% 'refer(1---1) Row 2 %>

																		<%' End %>

																	</table>
																</div>
															</td>
															<td class=ClearPixel width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
														</tr>
														<tr>
															<td class=MiddlePack width="267" colspan="3"></td>
														</tr>
													</table>
                                                </td>
											</tr>
										</table>
                                    </div>
								</td>
								<td align="center" class="ClearPixel"></td>
							</tr>
                            <tr>
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <!--input type="button" value="Back" name="B2" class="ActionButton" onClick="Back()"-->
                                                    <input type="button" value="Issue" name="B15" class="ActionButton" onClick="CheckSubmit('<%=formatDate(date())%>')">
                                                    <!--input type="button" value="Issue" name="B15" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(Date())%>')"-->
                                                    <!--input type="reset" value="Reset" name="B16" class="ActionButton"-->
                                                    <input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('ISSUEMGMT.ASP?ACTN=L')">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel" colspan="3">
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
