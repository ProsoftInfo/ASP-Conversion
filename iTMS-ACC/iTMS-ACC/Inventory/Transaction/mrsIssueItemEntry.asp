<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsIssueItemEntry.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 18, 2003
	'Modified On				:
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>MR Issue - Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	Dim oDom,objfs,Root,HeaderNode,iMRSNo,newElem
	Dim arrSSelected,arrSSelectedName,iCtr,arrItemClass
	dim dcrs,dcrs1,rsTemp

	dim sOrgName,dMRSDate,sIssue,sItmTypeName,sUsage,sItmType,sOrgID
	dim sReqType,sUsageCode,sITypeName,arrSchTemp,sSchTemp,sSchTempValue
	dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore,sItemName
	dim arrQty,iQtyReq,iQtyIssued,iQtyPending,iQtyAppr,iQtyTrans,iQtyPur
	dim iUnitQty,iOthUnitQty,iMarkQty,sIssueTo,sIssueForType,sPartyCode
	dim iQtyRes,iQtyOnHold,iQtyRej,dCreatedDate
	dim arrUoM,sUoMDesc,sUoMCode,sRecBy,sAttList
	dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin
	dim iuserId,sCreatedBy,rsUser
	dim sFinPeriod,arr,sMaxDate,sMinDate,sQuery,sArrList

	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	sMinDate = "01/04/"& Arr(0)
	sMaxDate = "31/03/"& Arr(1)

	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set rsTemp = Server.CreateObject("ADODB.RecordSet")
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

	iMRSNo = trim(Request.Form("mrs"))
	dCreatedDate = trim(Request.Form("ctlIssDate"))
	IF dCreatedDate = "" then dCreatedDate = sFinTo
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT CONVERT(CHAR,MRSDATE,103),ORGUNITSHORTDESCRIPTION,ISSUEDFORDESCRIPTION,'',MRSTYPE,'',MRSFORUNIT,ISSUEDFORCODE,CREATEDBY,ISSUEFOR,isNull(PartyCode,0) FROM VWMRSLIST WHERE MRSNUMBER = " & iMRSNo & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		dMRSDate = trim(dcrs(0))
		sOrgName = trim(dcrs(1))
		sUsage = trim(dcrs(2))

		sItmTypeName = trim(dcrs(3))
		sIssue = trim(dcrs(4))
		sItmType = trim(dcrs(5))
		if trim(dcrs(10))<>"0" then
		    sPartyCode = trim(dcrs(10))
		end if

		if IsNull(sItmType) or sItmType = "" then sItmType = "0"
		if IsNull(sItmTypeName) or sItmTypeName = "" then sItmTypeName = "0"

		sOrgID = trim(dcrs(6))
		sUsageCode = trim(dcrs(7))
		iuserId = dcrs(8)
		sIssueTo = trim(dcrs(9))
		sIssueForType = sIssueTo
		'IF sIssueTo = "S" then sIssueForType = "Stores"

		with rsUser
			.CursorLocation = 3
			.CursorType = 3
			.Source ="Select isNull(UserName,'') from DCS_User where EmployeeNumber = "&iuserId&" "
			.ActiveConnection = con
			.Open
		end with
		if not rsUser.EOF then
			sCreatedBy = trim(rsUser(0))
		end if
		rsUSer.Close

		if sIssue = "0" then
			sIssue = "Returnable"
		else
			sIssue = "Non Returnable"
		end if
	end if
	dcrs.Close


	Set Root = oDOM.createElement("MRSApproval")
	oDOM.appendChild Root
	Set newElem = oDOM.createElement("MRSHeader")
	newElem.setAttribute "MRSNO",iMRSNo
	newElem.setAttribute "MRSDATE",dMRSDate
	newElem.setAttribute "ORGID",sOrgID
	newElem.setAttribute "ORGNAME", sOrgName
	newElem.setAttribute "REQTYPE", sIssue

	newElem.setAttribute "USAGE", sUsageCode
	newElem.setAttribute "USAGENAME", sUsage
	newElem.setAttribute "ITYPE", sItmType
	newElem.setAttribute "ITYPENAME", sItmTypeName

	Root.appendChild newElem
	oDOM.Save server.MapPath("../temp/transaction/MRSIssue"&Session.SessionID&".xml")

	arrSSelected = split(trim(populateClassItem (iMRSNo,sItmType,"Y")),"|")
	arrSSelectedName = split(trim(populateClassItem (iMRSNo,sItmType,"N")),"|")

	'Response.Write trim(populateClassItem (iMRSNo,sItmType,"Y"))

	set oDom = nothing

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	Set Root = oDOM.createElement("ISSTYPE")
	oDOM.appendChild Root

	for iCtr = 0 to UBound(arrSSelected) - 1
		arrItemClass = split(arrSSelected(iCtr),":")

		'Response.Write "arrItemClass(4) = "&arrItemClass(4)
		Set newElem = oDOM.createElement("ITEM")
		newElem.setAttribute "ENTRYNO",trim(arrItemClass(2))
		newElem.setAttribute "ITMCODE",trim(arrItemClass(1))
		newElem.setAttribute "CLACODE",trim(arrItemClass(0))
		newElem.setAttribute "ITMNAME",trim(arrSSelectedName(iCtr))
		newElem.setAttribute "SSTORE", ""
		newElem.setAttribute "REQQTY", "0"
		newElem.setAttribute "ISSQTY", ""
		newElem.setAttribute "TRAQTY", "0"
		newElem.setAttribute "PRQTY", "0"
		newElem.setAttribute "IVALUE", "0"
		newElem.setAttribute "ORGCODE", ""
		newElem.setAttribute "MRSNO", iMRSNo
		newElem.setAttribute "MRSDATE", dMRSDate
		newElem.setAttribute "ATTRIBUTELIST",trim(arrItemClass(4))
		newElem.setAttribute "CREATEDBY", iuserId
		newElem.setAttribute "CREATEDON", dCreatedDate
		newElem.setAttribute "RefNo", iMRSNo

		Root.appendChild newElem
	next
	oDOM.Save server.MapPath("../temp/transaction/MRISSUEDETAILS"&Session.SessionID&".xml")

	if objfs.FileExists(Server.MapPath("../temp/transaction/MRS"&iMRSNo&".xml")) then
%>
<script type="application/xml" data-itms-xml-island="1" id="Data" data-src="<%="../temp/transaction/MRS"&iMRSNo&".xml"%>"></script>
<%	else %>
<script type="application/xml" data-itms-xml-island="1" id="Data"><root/></script>
<%	end if %>
<script type="application/xml" data-itms-xml-island="1" id="OutData1" data-src="<%="../temp/transaction/MRISSUEDETAILS"&Session.SessionID&".xml"%>"></script>

<script type="application/xml" data-itms-xml-island="1" id="OutData"><root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsIssueItemDetails.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="LoadData();SetDate('<%=FormatDate(Date())%>')">

<form method="POST" name="formname">
<!--OBJECT id=penDet type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" VIEWASTEXT>
<PARAM name="Command" value="HH Version">
</OBJECT-->
<input type=hidden name="hMRSNo" value="<%=iMRSNo%>">
<input type=hidden name="hReqDate" value="<%=dMRSDate%>">
<input type=hidden name="hMinDate" Value="<%=sMinDate%>">
<input type=hidden name="hMaxDate" Value="<%=sMaxDate%>">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hItmType" value="<%=sItmType%>">
<input type="hidden" name="mrs" value="<%=iMRSNo%>">
<input type="hidden" name="sAct" value="mrsIssueItemEntry.asp">
<input type="hidden" name="hUsage" value="<%=sUsageCode%>">
<input type="hidden" name="hIssForType" Value="<%=sIssueTo%>">
<input type=hidden name="hPartyCode" value="<%=sPartyCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Material Issue
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
					<TD class=TabBodywithtopline>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" class="ClearPixel">&nbsp;

								</td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                                    <td class="FieldCell">Reference Name</td>
													<td class="FieldCellSub">

														<span class="DataOnly" align=center>Material Requisition</span>


													<!--<span class="DataOnly">N/A&nbsp;</span>-->
												</td>

                                                    <td class="FieldCell">Issue Date</td>
													<td class="FieldCellSub" valign="middle">
														<object id="ctlIssDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"      codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
															<param name="_ExtentX" value="2355">
															<param name="_ExtentY" value="529">
														</object>
													</td>
												</tr>
                                                   <tr>
                                                    <td class="FieldCell">Reference No - Date</td>
													<td class="FieldCellSub">

														<span class="DataOnly" align=center><%=iMRSNo%>-<%=dMRSDate%></span>


													<!--<span class="DataOnly">N/A&nbsp;</span>-->
												</td>

                                                    <td class="FieldCell">Created By</td>
														<td class="FieldCellSub">
															<span class="dataonly"><%=sCreatedBy%></span>
														</td>
												</tr>

										<!--tr>
										    <td class="FieldCell">Requisition by Unit&nbsp;</td>
										    <td class="FieldCellSub" width="175"><span class="DataOnly" id="idOrgName"><%=sOrgName%>&nbsp;</span></td>
										    <!--td class="FieldCell">Requisition Date</td>
										    <td class="FieldCellSub"><span class="DataOnly"><%=dMRSDate%>&nbsp;</span></td-->
										     <!--td class="FieldCell">Requisition Type</td>
										    <td class="FieldCellSub"><span class="DataOnly"><%=sIssue%>&nbsp;</span></td>
										</tr>
										<tr>


										</tr-->
										<tr>
										    <!--class="FieldCell">Item Type</td>
										    <td class="FieldCellSub"><span class="DataOnly"><%=sItmTypeName%>&nbsp;</span></td-->
										    <td class="FieldCell">Usage of Item</td>
										    <td class="FieldCellSub"><span class="DataOnly"><%=sUsage%>&nbsp;</span></td>

										   <td class="FieldCell">Issue To</td>
										    <td class="FieldCellSub"><span class="DataOnly"><%=sIssueForType%>&nbsp;</span></td>
										</tr>
										<tr>
										    <td class="FieldCell">Received By</td>&nbsp;
											<td class="FieldCellSub">
												<input type="text" name="txtRecBy" size="35" class="FormElem" maxlength=35 style="text-align:left">
											</td>
											 <td class="FieldCell">Issue Type</td>
										    <td class="FieldCellSub">
												<%	if sItmType = "STO" then %>
													<select name="selIssType" class="FormElem">

														<option value="F">Firm</option>
														<option value="M">Marked</option>
													</select>
												<%	else %>
													<select name="selIssType" class="FormElem" disabled>
														<option value="select">Select</option>
														<option value="F">Firm</option>
														<option value="M" selected>Marked</option>
													</select>
												<%	end if %>
										    </td>
										</tr>
										<tr>
										    <td class="FieldCell">Remarks</td>

										    <td class="FieldCellSub" >
										    <textarea rows="2" name="Remarks" cols="32" class="Formelem" maxlength="100"></textarea>
										    <!--td class="FieldCell">Issue Date</td>
										    <td class="ExcelInutCell">
											  <%'Response.Write InsertDatePicker("ctlIssDate")%>
											</td-->
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
										<table cellpadding="0" cellspacing="0" width="100%">
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
																<DIV class=frmBody id=frm2 style="width: 750; height:285;">
																	<table border="0" cellspacing="1" class="ExcelTable" width="100%">
																		<tr>
																			<td class="ExcelHeaderCell" align="center" width="10" rowspan="3">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" colspan="4">Requisition Details</td>
																			<!--td class="ExcelHeaderCell" align="center" rowspan="3">Store</td-->
																			<td class="ExcelHeaderCell" align="center" colspan="3">Quantity Availability</td>
																			<td class="ExcelHeaderCell" align="center" colspan="4">Stock</td>
																		</tr>
																	    <tr>
																			<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
																			<td class="ExcelHeaderCell" align="center">Approved</td>
																			<td class="ExcelHeaderCell" align="center">Pending</td>
																			<td class="ExcelHeaderCell" align="center">By Date's</td>
																			<td class="ExcelHeaderCell" align="center">In Unit</td>
																			<td class="ExcelHeaderCell" align="center">Reserved</td>
																			<td class="ExcelHeaderCell" align="center">Transit</td>
																			<td class="ExcelHeaderCell" align="center">Issue</td>
																			<td class="ExcelHeaderCell" align="center">Transfer</td>
																			<td class="ExcelHeaderCell" align="center">Purchase</td>
																			<td class="ExcelHeaderCell" align="center">Total</td>
																	    </tr>
																		<tr>
																			<td class="ExcelHeaderCell" align="center">Issued</td>
																			<td class="ExcelHeaderCell" align="center">Tra. / PR</td>
																			<td class="ExcelHeaderCell" align="center">Quality</td>
																			<td class="ExcelHeaderCell" align="center">Other Unit</td>
																			<td class="ExcelHeaderCell" align="center">On Hold</td>
																			<td class="ExcelHeaderCell" align="center">Rejected</td>
																			<td class="ExcelHeaderCell" align="center">Date</td>
																			<td class="ExcelHeaderCell" align="center">By Date</td>
																			<td class="ExcelHeaderCell" align="center">By Date</td>
																			<td class="ExcelHeaderCell" align="center">UoM</td>
																		</tr>
																	<%	for iCtr = 0 to UBound(arrSSelectedName) - 1%>
																			<%	'Calling the Function which populates the Store list
																				iOthUnitQty = 0
																				arrItemClass = split(arrSSelected(iCtr),":")
																				'Response.Write arrSSelected(0)
																				'arrLocation = split(populateStore(sOrgID,arrItemClass(0),arrItemClass(1),iMRSNo),"|")
																				'sStoreName = arrLocation(0)
																				'arrStore = split(arrLocation(1),"-")
																				'sStoreCode = arrStore(0)
																				'if UBound(arrStore) > 0 then
																				'	sBinCode = arrStore(1)
																				'else
																				'	sBinCode = "N"
																				'end if
																				
																				arrUoM = split(DisplayUoM (sOrgID,arrItemClass(0),arrItemClass(1)),":")
																				sUoMCode = arrUoM(0)
																				sUoMDesc = arrUoM(1)

																				arrQty = split(DisplayQty(sOrgID,arrItemClass(0),arrItemClass(1),arrItemClass(2),iMRSNo),":")

																				iQtyReq = cdbl(trim(arrQty(0)))
																				iQtyAppr = cdbl(trim(arrQty(1)))
																				iQtyIssued = cdbl(trim(arrQty(2)))
																				iQtyPur = cdbl(trim(arrQty(3)))
																				iQtyTrans = cdbl(trim(arrQty(4)))

																				'Response.Write iQtyReq & "<BR>" & iQtyAppr & "<BR>" & iQtyIssued & "<BR>" & iQtyPur & "<BR>" & iQtyTrans & "<BR>"

																				iQtyPending = cdbl(iQtyAppr - (iQtyIssued + (iQtyPur+iQtyTrans)))
																				sArrList = arrItemClass(4)
																				if trim(sArrList)<>"" and trim(sArrList)<>"NULL" then
																				    sAttList = split(sArrList,"#")
																				end if
																				
																				
																				sQuery = " SELECT isNull(SUM(AVAILABLENETSTOCK),0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & arrItemClass(1) & " AND ClassificationCode = "& arrItemClass(0)
																				'Response.Write sQuery
																				if trim(sArrList)<>"" and trim(sArrList)<>"NULL" then
																				    if UBound(sAttList)>0 then
																				        if trim(sAttList(1))<>"0" and trim(sAttList(1))<>"" then
		  	                                                                                   sQuery  = sQuery & " and AttributeList in ('"& sAttList(1) &"')"
	                                                                                    end if
	                                                                                else
																						if trim(sAttList(0))<>"0" and trim(sAttList(0))<>"" then
																							sQuery  = sQuery & " and AttributeList in ('"& sAttList(0) &"')"
																						end if
	                                                                                end if
	                                                                            end if  'if trim(sArrList)<>"" then
	                                                                            
	                                                                           ' Response.Write sQuery

																				with dcrs
																					.CursorLocation = 3
																					.CursorType = 3
																					.Source = sQuery
																					'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0) FROM VWItemStockStatus WHERE ITEMCODE = " & arrItemClass(1) & " AND CLASSIFICATIONCODE = " & arrItemClass(0) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
																					.ActiveConnection = con
																					.Open
																				end with
																				'Response.Write dcrs.Source
																				set dcrs.ActiveConnection = nothing

																				iUnitQty = "0"
																				if Not dcrs.EOF then
																					iUnitQty = dcrs(0)
																				end if
																				dcrs.Close

																				with dcrs
																					.CursorLocation = 3
																					.CursorType = 3
																					.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0),ISNULL(SUM(QUANTITYISSUED),0) FROM INV_T_MRSISSUEPICK WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & arrItemClass(1) & " AND CLASSIFICATIONCODE = " & arrItemClass(0) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
																					.ActiveConnection = con
																					.Open
																				end with
																				set dcrs.ActiveConnection = nothing

																				iMarkQty = "0"
																				if Not dcrs.EOF then
																					iMarkQty = cdbl(dcrs(0)) - cdbl(dcrs(1))
																				end if
																				dcrs.Close

																				with dcrs
																					.CursorLocation = 3
																					.CursorType = 3
																					.Source = "SELECT ISNULL(SUM(RESERVED),0),ISNULL(SUM(ONHOLD),0),ISNULL(SUM(REJECTED),0) FROM VWITEMSTOCKSTATUS WHERE ITEMCODE = " & arrItemClass(1) & " AND CLASSIFICATIONCODE = " & arrItemClass(0) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
																					.ActiveConnection = con
																					.Open
																				end with
																				set dcrs.ActiveConnection = nothing

																				iQtyRes = "0"
																				iQtyOnHold = "0"
																				iQtyRej = "0"

																				if Not dcrs.EOF then
																					iQtyRes = trim(dcrs(0))
																					iQtyOnHold = trim(dcrs(1))
																					iQtyRej = trim(dcrs(2))
																				end if
																				dcrs.Close

																			'	with dcrs1
																			'		.CursorLocation = 3
																			'		.CursorType = 3
																			'		.Source = "SELECT DISTINCT OTHERORGANISATIONCODE FROM INV_M_ITEMORGTRANSFERUNITS WHERE ITEMCODE = " & arrItemClass(1) & " AND CLASSIFICATIONCODE = " & arrItemClass(0) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
																			'		.ActiveConnection = con
																			'		.Open
																			'	end with
																			'	set dcrs1.ActiveConnection = nothing

																			'	if not dcrs1.EOF then
																			'		Do While Not dcrs1.EOF

																			'			with dcrs
																			'				.CursorLocation = 3
																			'				.CursorType = 3
																			'				.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0) FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & arrItemClass(1) & " AND CLASSIFICATIONCODE = " & arrItemClass(0) & " AND ORGANISATIONCODE = " & Pack(trim(dcrs1(0))) & " AND APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1"
																			'				.ActiveConnection = con
																			'				.Open
																			'			end with
																			'			set dcrs.ActiveConnection = nothing

																			'			if Not dcrs.EOF then
																			'				iOthUnitQty = cdbl(iOthUnitQty) + cdbl(dcrs(0))
																			'			end if
																			'			dcrs.Close

																			'		dcrs1.MoveNext
																			'		loop
																			'	else
																			'		iOthUnitQty = "0"
																			'	end if
																			'	dcrs1.Close

																			iOthUnitQty = "0"
																				if iOthUnitQty = "" then iOthUnitQty = "0"

																				arrSchTemp = split(GetSchedule(sOrgID,arrItemClass(0),arrItemClass(1),arrItemClass(2),iMRSNo),":")
																				sSchTemp = ""
																				if UBound(arrSchTemp) > 0 then
																					sSchTemp = arrSchTemp(0)
																					sSchTempValue = arrSchTemp(1)
																				end if
																				'sItemName = ItemDisplay(arrItemClass(1),arrItemClass(0))
																				with dcrs
																					.CursorLocation = 3
																					.CursorType = 3
																					.Source = "SELECT ITEMDESCRIPTION  FROM VWITEM WHERE ITEMCODE = " & arrItemClass(1) & " "
																					.ActiveConnection = con
																					.Open
																				end with
																				set dcrs.ActiveConnection = nothing

																				if Not dcrs.EOF then

																					sItemName = dcrs(0)
																				end if
																				dcrs.Close
																				if arrItemClass(3)  <> "" THEN sItemName = sItemName& arrItemClass(3)


																			%>
																			<input type="hidden" name="hSchA<%=cstr(arrItemClass(0))%>A<%=cstr(arrItemClass(1))%>A<%=cstr(arrItemClass(2))%>" value="<%=sSchTemp%>">
																			<input type="hidden" name="hStoA<%=sStoreCode%>A<%=sBinCode%>" value="<%=sStoreCode%>:<%=sBinCode%>">
																			<input type="hidden" name="hStoName<%=cstr(arrItemClass(0))%>A<%=cstr(arrItemClass(1))%>A<%=cstr(arrItemClass(2))%>" value="<%=sStoreName%>">
																		<tr>
																			<td class="ExcelSerial" align="center" rowspan="2"><%=iCtr + 1%></td>
																			<td class="ExcelDisplayCell" rowspan="2"><p align="left">
																				<a href="javascript:void(0)" class="ExcelDisplayLink" name="lnkA<%=cstr(arrItemClass(0))%>A<%=cstr(arrItemClass(1))%>A<%=sOrgID%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%></a>
																			</td>
																			<td class="ExcelDisplayCell" width="10"><p align="right">
																				<input type="text" name="txtQtyA<%=cstr(arrItemClass(0))%>A<%=cstr(arrItemClass(1))%>A<%=cstr(arrItemClass(2))%>" size="12" value="<%=iQtyAppr%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			</td>
																			<td class="ExcelDisplayCell" width="10"><p align="right">
																				<input type="text" name="txtQtyPenX<%=cstr(arrItemClass(0))%>X<%=cstr(arrItemClass(1))%>X<%=cstr(arrItemClass(2))%>" size="12" value="<%=iQtyPending%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			</td>
																			<td class="ExcelDisplayCell" width="91"><p align="right">
																				<%
																					if sSchTemp = "I" then
																						Response.Write trim(sSchTempValue)
																					elseif sSchTemp = "W" then
																						Response.Write "Within " &sSchTempValue& " Days"
																					elseif sSchTemp = "D" then
																						Response.Write sSchTempValue
																					elseif sSchTemp = "S" then
																						Response.Write sSchTempValue
																				%>
																					<a href="javascript:void(0)">
																						<img name="btn:<%=arrSSelected(iCtr)%>" border="0" src="../../assets/images/iTMS%20Icons/Details.gif" width="15" height="15" alt="Schedule Details" onClick="CheckSch(this)">
																					</a>
																				<%	end if %>
																			</td>
																			<!--td class="ExcelDisplayCell" rowspan="2"><p align="left"><%=sStoreName%></td-->
																			<td class="ExcelDisplayCell" width="10"><p align="right">
																				<input type="text" name="txtStockZ<%=cstr(arrItemClass(0))%>Z<%=cstr(arrItemClass(1))%>Z<%=cstr(arrItemClass(2))%>Z<%=cstr(arrItemClass(3))%>Z<%=sStoreCode%>Z<%=sBinCode%>" size="10" value="<%=cdbl(iUnitQty)%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right;cursor:hand;FONT-WEIGHT: bold" alt="In Unit Stock Details" onClick="DisplayStock(this)">
																			</td>
																			<td class="ExcelDisplayCell">
																				<p align="right"><%=iQtyRes%><!--a href="mrsIssuePopupB.html" class="ExcelDisplayLink" target="_blank">-</a-->
																			</td>
																			    <td class="ExcelDisplayCell">
																			    <p align="right"><!--a href="mrsIssuePopupB.html" class="ExcelDisplayLink" target="_blank">-</a-->
																			</td>
																			<td class="ExcelInputCell" width="91"><p align="left">

																				<a href="javascript:void(0)">
																				<%if trim(sArrList)<>"" and trim(sArrList)<>"NULL" then %>
																				    <%if UBound(sAttList)>0  then%>
																					    <img name="btnZ<%=cstr(arrItemClass(0))%>Z<%=cstr(arrItemClass(1))%>Z<%=cstr(arrItemClass(2))%>Z<%=cstr(arrItemClass(3))%>Z<%=sStoreCode%>Z<%=sBinCode%>Z<%=cstr(sAttList(1))%>" border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" width="15" height="15" alt="Pick Details" onClick="CheckLot(this,'<%=iQtyPending%>')">
																					<%else %>
																					    <img name="btnZ<%=cstr(arrItemClass(0))%>Z<%=cstr(arrItemClass(1))%>Z<%=cstr(arrItemClass(2))%>Z<%=cstr(arrItemClass(3))%>Z<%=sStoreCode%>Z<%=sBinCode%>Z<%=cstr(sAttList(0))%>" border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" width="15" height="15" alt="Pick Details" onClick="CheckLot(this,'<%=iQtyPending%>')">
																					<%end if %>
																			    <%else %>
																			        <img name="btnZ<%=cstr(arrItemClass(0))%>Z<%=cstr(arrItemClass(1))%>Z<%=cstr(arrItemClass(2))%>Z<%=cstr(arrItemClass(3))%>Z<%=sStoreCode%>Z<%=sBinCode%>Z" border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" width="15" height="15" alt="Pick Details" onClick="CheckLot(this,'<%=iQtyPending%>')">
																			    <%end if%>
																				</a>
																				<input type="text" name="txtQtyPPX<%=cstr(arrItemClass(0))%>X<%=cstr(arrItemClass(1))%>X<%=cstr(arrItemClass(2))%>" size="12" class="FormElem" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" style="text-align:right" onBlur="GetSch(this)" READONLY>
																			</td>
																			<td class="ExcelInputCell"><p align="left">
																				<input type="text" name="txtQtyTraX<%=cstr(arrItemClass(0))%>X<%=cstr(arrItemClass(1))%>X<%=cstr(arrItemClass(2))%>" value="0" size="14" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" style="text-align:right" class="Formelem" READONLY>
																				<a href="javascript:void(0)">
																					<img name="btn:<%=arrSSelected(iCtr)%>" border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" width="15" height="15" alt="Stock Transfer" onClick="CheckST(this)">
																				</a>
																			</td>
																			<td class="ExcelInputCell"><p align="left">
																				<input type="text" name="txtQtyPrX<%=cstr(arrItemClass(0))%>X<%=cstr(arrItemClass(1))%>X<%=cstr(arrItemClass(2))%>" value="0" size="18" class="FormElem" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" style="text-align:right">
																			</td>
																			<td class="ExcelInputCell" width="10"><p align="left">
																				<input type="text" name="txtQtyTotX<%=cstr(arrItemClass(0))%>X<%=cstr(arrItemClass(1))%>X<%=cstr(arrItemClass(2))%>" size="12" class="FormElem" maxlength=10 readonly style="text-align:right">
																			</td>
																		</tr>
																		<tr>
																			<td class="ExcelDisplayCell" width="10"><p align="right">
																				<input type="text" name="txtQtyIssX<%=cstr(arrItemClass(0))%>X<%=cstr(arrItemClass(1))%>X<%=cstr(arrItemClass(2))%>" size="12" value="<%=iQtyIssued%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			</td>
																			<td class="ExcelDisplayCell" width="10"><p align="right">
																				<input type="text" name="txtQtyMarX<%=cstr(arrItemClass(0))%>X<%=cstr(arrItemClass(1))%>" size="10" value="<%=cdbl(iQtyPur)+cdbl(iQtyTrans)%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right;cursor:hand;FONT-WEIGHT: bold" alt="BreakUp Details" onClick="DisplayDet('<%=iQtyPur%>|<%=iQtyTrans%>')">
																			</td>
																			<td class="ExcelDisplayCell"><p align="right">
																				<a href="javascript:void(0)">
																					<img name="btn:<%=arrSSelected(iCtr)%>" border="0" src="../../assets/images/iTMS%20Icons/Details.gif" width="15" height="15" alt="Quality Parameters" onClick="CheckQty(this)">
																				</a>
																			</td>
																			<td class="ExcelDisplayCell">
																				<p align="right"><%=iOthUnitQty%><!--a href="mrsIssuePopupB.html" class="ExcelDisplayLink" target="_blank"></a></p-->
																			</td>
																			<td class="ExcelDisplayCell">
																				<p align="right"><%=iQtyOnHold%><!--a href="mrsIssuePopupB.html" class="ExcelDisplayLink" target="_blank">-</a></p-->
																			</td>
																			<td class="ExcelDisplayCell">
																				<p align="right"><%=iQtyRej%><!--a href="mrsIssuePopupB.html" class="ExcelDisplayLink" target="_blank">-</a></p-->
																			</td>
																			<td class="ExcelDisplayCell">
																				<!--p align="right">xDate<a href="mrsIssueSchedule.html" target="_blank"><img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" width="15" height="15" alt="Change Alt Tag"></a></p-->
																			</td>
																			<td class="ExcelFieldCell" align="left" width=30>
																			    <select size="1" name="selSTSchZ<%=cstr(arrItemClass(0))%>Z<%=cstr(arrItemClass(1))%>Z<%=cstr(arrItemClass(2))%>Z<%=cstr(arrItemClass(3))%>" class="FormElem" onChange="CheckSTPRSch(this,'<%=FormatDate(date())%>','ST')">
																					<!--option value="select">Select</option-->
																					<option value="ID">Immediate</option>
																					<option value="WD">Within x Days</option>
																					<option value="SD">Specific Date</option>
																					<option value="S">Scheduled</option>
																			    </select>
																			</td>
																			<td class="ExcelFieldCell" align="left" width=30>
																			    <select size="1" name="selPRSchZ<%=cstr(arrItemClass(0))%>Z<%=cstr(arrItemClass(1))%>Z<%=cstr(arrItemClass(2))%>Z<%=cstr(arrItemClass(3))%>" class="FormElem" onChange="CheckSTPRSch(this,'<%=FormatDate(date())%>','PR')">
																					<!--option value="select">Select</option-->
																					<option value="ID">Immediate</option>
																					<option value="WD">Within x Days</option>
																					<option value="SD">Specific Date</option>
																					<option value="S">Scheduled</option>
																			    </select>
																			</td>
																			<td class="ExcelDisplayCell"><%=sUoMDesc%></td>
																		</tr>

																	<%	next %>
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
                                                    <input type="reset" value="Reset" name="B16" class="ActionButton">
                                                    <!--input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('mrsMgmtEntry.asp?sCheck=M')"-->
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
	' Function to populate Store
	Function populateStore(sOrgID,iClass,iItem,iMRSNoP)
		' Declaration of variables
		Dim dcrs1,sLoc,sBin,sBinName,sLocName,sLocAmd,sBinAmd
		'Declaration of Objects
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNoP & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
'Response.Write dcrs1.Source
		if not dcrs1.EOF then
			sLocAmd = dcrs1(0)
			sBinAmd = dcrs1(1)
		end if
		dcrs1.close

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT LOCATIONNAME,LOCATIONCODE FROM INV_M_STORAGE WHERE LOCATIONNUMBER = " & sLocAmd & ""
			.ActiveConnection = con
			.Open
		end with
		if not dcrs1.EOF then
			sLoc = dcrs1(0)
		end if
		dcrs1.close

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM INV_M_STOREBINDETAILS WHERE LOCATIONNUMBER = " & sLocAmd & " AND BINNUMBER = " & sBinAmd & " ORDER BY BINNUMBER"
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		if not dcrs1.EOF then
			populateStore = trim(sLoc)&" -- "&trim(dcrs1(2))&"|"&sLocAmd&"-"&trim(dcrs1(0))
		else
			populateStore = trim(sLoc)&"|"&sLocAmd
		end if
		dcrs1.Close
	End Function
%>

<%
	' Function to populate Quantity Requested
	Function DisplayQty(sOrgID,iClass,iItem,iEntNo,iMRSNoP)
		' Declaration of variables
		Dim dcrs,iQtyReq,iQtyIssued,iQtyAppr,iQtyPur,iQtyTransfer
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(QUANTITYREQUESTED,0),ISNULL(QUANTITYAPPROVED,0),ISNULL(QUANTITYISSUED,0),(ISNULL(QUANTITYTOPURCHASE,0) - ISNULL(QUANTITYPURCHASED,0)),(ISNULL(QUANTITYFORTRANSFER,0) - ISNULL(QUANTITYTRANSFERRED,0)) FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNoP & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND ICOUNTER = " & iEntNo & " "
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.Source
		set dcrs.ActiveConnection = nothing

		set iQtyReq = dcrs(0)
		set iQtyAppr = dcrs(1)
		set iQtyIssued = dcrs(2)
		set iQtyPur = dcrs(3)
		set iQtyTransfer = dcrs(4)

		if Not dcrs.EOF then
			DisplayQty = iQtyReq&":"&iQtyAppr&":"&iQtyIssued&":"&iQtyPur&":"&iQtyTransfer
		end if
		dcrs.Close

	End Function
%>

<%
	' Function to populate Schedule Type
	Function GetSchedule(sOrgID,iClass,iItem,iEntNo,iMRSNoP)
		' Declaration of variables
		Dim dcrs,dcrs1,sReqBy,sReqValue
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT REQUIREDBY,REQUIREDVALUE FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNoP & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND ICOUNTER = " & iEntNo & " "
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing
		set sReqBy = dcrs(0)
		set sReqValue = dcrs(1)
		if Not dcrs.EOF then
			if not sReqValue = "S" then
				GetSchedule = sReqBy&":"&sReqValue
			else
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT SCHEDULEDON FROM INV_T_MRSITEMSCHEDULES WHERE MRSNUMBER = " & iMRSNoP & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " ORDER BY SCHEDULENO"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
				set sReqValue = dcrs1(0)
				if Not dcrs1.EOF then
					GetSchedule = sReqBy&":"&sReqValue
				end if
				dcrs1.Close
			end if
		end if
		dcrs.Close

	End Function
%>

<%
	' Function to populate Store
	Function DisplayUoM(sOrgID,iClass,iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate Classification and Item
	Function populateClassItem(iMRSNoP,sItmTypeP,sFlagP)
		' Declaration of variables
		Dim dcrs,sItemDesc,sItemShDesc,sClassDesc,sTempClassCode,iTempItmCode
		dim sTemp,iAttributeList,iEntNo,sOptName,sTempAttList
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ITEMDESCRIPTION,SHORTDESCRIPTION,GROUPNAME,CLASSIFICATIONCODE,ITEMCODE FROM VWMRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNoP & " AND (ISNULL(QUANTITYISSUED,0) + ISNULL(QUANTITYTOPURCHASE,0) + ISNULL(QUANTITYFORTRANSFER,0)) < ISNULL(QUANTITYAPPROVED,0) ORDER BY ITEMCODE"
			.Source = "SELECT ITEMDESCRIPTION,SHORTDESCRIPTION,GROUPNAME,CLASSIFICATIONCODE,ITEMCODE,ISNULL(ITEMATTRIBUTES,''),ISNULL(ICOUNTER,0) FROM VWMRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNoP & " AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) ORDER BY ITEMCODE"
			'Response.Write dcrs.source
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.Source
		set dcrs.ActiveConnection = nothing

		Do While Not dcrs.EOF
			sItemDesc = dcrs(0)
			sItemShDesc = dcrs(1)
			sClassDesc = dcrs(2)
			sTempClassCode = dcrs(3)
			iTempItmCode = dcrs(4)
			sAttList = dcrs(5)
			iAttributeList = split(sattList,":")
			iEntNo = dcrs(6)
			if trim(sAttList)<>"" then
			    'Response.Write "iAttributeList = "& iAttributeList(0)
			    IF (iAttributeList(0)) <> "" then
			        sTempAttList = split(iAttributeList(0),"#")
			        if UBound(sTempAttList)>0 then
				        sOptName = FunAttribName(iAttributeList(0))
				    else
				        sQuery = "Select OptionName from INV_M_ITEMTYPEOPTIONS where OptionValue = "& iAttributeList(0)
				        rsTemp.Open sQuery,con
				        if not rsTemp.EOF then
				            sOptName = "["& trim(rsTemp(0))&"]"
				        end if
				        rsTemp.Close 
				    end if
			    Else
				    sOptName = ""
			    End IF
			 end if

			IF trim(sOptName) <> "" then sItemDesc = sItemDesc & sOptName
			'Response.Write iAttributeList &"--"& iEntNo

			if sFlagP = "Y" then
			    if trim(sAttList)<>"" then
				    sTemp = sTemp & trim(sTempClassCode)&":"&trim(iTempItmCode) &":"&trim(iEntNo)&":"&Trim(sOptName)&":"&iAttributeList(0) & "|"
				else
				    sTemp = sTemp & trim(sTempClassCode)&":"&trim(iTempItmCode) &":"&trim(iEntNo)&":"&Trim(sOptName)&":"& "|"
				end if
			else
				sTemp = sTemp & trim(sItemDesc) & "|"
			end if
			dcrs.MoveNext
		Loop
		dcrs.Close
		populateClassItem = mid(sTemp,1,len(sTemp))
	End Function
%>

