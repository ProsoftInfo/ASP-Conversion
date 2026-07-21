<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	MRSMGMTList.asp
	'Module Name				:	INVENTORY (Requisition List)
	'Author Name				:
	'Created On					:
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	Jan 05,2011
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Grid</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root/></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/printwindow.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsMgmt.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
  Function Init(dDate)
	document.formname.ctlFromDate.SetDate = document.formname.hFrmDate.value
	document.formname.ctlToDate.SetDate = document.formname.hToDate.value
'	if document.formname.hCheck.value <> "Z" then document.formname.ctlFromDate.SetDate = cdate(ddate) - 15
  end Function
  '********************************************************************************************
  Function checkMR()
	Dim OutValue,objhttp
	Dim sOrgID,sItemType
		sOrgID= document.formname.hUnit.value
		'sItemType=document.formname.selItemType.value

	'    set OutValue =  showModalDialog("IssueUsageSelPop.asp?OrgID="& document.formname.hUnit.value,RefData,"dialogHeight:340px;dialogwidth=500px;center:yes;help:no;resizable:no;status:no")
	 '   if OutValue.getAttribute("Done")="Y" then
	'	    set objhttp = CreateObject("Microsoft.XMLHTTP")
	'	    objhttp.open "POST","XMLSave.asp?Name=UsageSelection&SessionFlag=true",false
	'	    objhttp.send RefData.XMLDocument
    '
	'	    document.formname.action="MRGENERATIONENTRY.ASP?sOrg="&sOrgID&"&sIType="& sItemType
	'	    document.formname.submit()
	 '   end if

    document.formname.action="MRGENERATIONENTRY.ASP?sOrg="&sOrgID
    document.formname.submit()
  End Function
  '*****************************************************
  Function MinDate()

		Dim sMinDate,sFinPeriod,sSelDate,sMaxDate
		'alert("date check")
		'sFinPeriod = document.formname.hFinPeriod.value
		sMinDate = document.formname.hFrmDate.value
		sMaxDate = document.formname.hToDate.value
		RngFrom = document.formname.ctlFromDate.getdate
		RngTo =document.formname.ctlToDate.getDate
		'alert(RngFrom &"="& sMinDate)
		If DateValue(RngFrom) < DateValue(sMinDate) or  DateValue(RngFrom) > DateValue(sMaxDate) then
			Alert("Date Should be With in the Range "& sMinDate & " to " & sMaxDate)
			document.formname.ctlFromDate.Setdate = sMinDate
			Exit function
		End If
		If DateValue(RngTo) < DateValue(sMinDate) or  DateValue(RngTo) > DateValue(sMaxDate) then
			Alert("Date Should be With in the Range "& sMinDate & " to " & sMaxDate)
			document.formname.ctlToDate.Setdate = sMaxDate
			Exit function
		End If
	End Function

'********************************************************************************************
  Function CheckSubmit()
		if document.formname.hCheck.value <> "Z" then
			document.formname.hFrmDate.value = document.formname.ctlFromDate.GetDate()
			document.formname.hToDate.value = document.formname.ctlToDate.GetDate()
			sUsage = ""
		end if

		sIssType = document.formname.cmbIssType(document.formname.cmbIssType.selectedIndex).value
		document.formname.action = "mrsMgmtList.asp?ISSTYPE="& sIssType
		document.formname.submit()

End Function
'*******************************************************************************
Function DirectIssue()
	Dim OutValue,objhttp
	set OutValue =  showModalDialog("IssueUsageSelPop.asp?OrgID="& document.formname.hUnit.value,RefData,"dialogHeight:340px;dialogwidth=500px;center:yes;help:no;resizable:no;status:no")
	if OutValue.getAttribute("Done")="Y" then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "POST","XMLSave.asp?Name=UsageSelection&SessionFlag=true",false
		objhttp.send RefData.XMLDocument

		document.formname.action = "DirectIssueItemEntry.asp"
		document.formname.submit()
	end if
End Function
'*******************************************************************************
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<%
	dim dcrs,dcrs1,iCtr,sAmendedBy,sCheck,sStatus,sOrgID,sOrgName,sAction,iCnt
	dim dFrmDate,dToDate,sType,sUsage,sCreatedBy,sItemType,sIssTypeCodestr,sArrRef,sReferenceName
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	dim sFinPeriod,Arr,sIssType
	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	dFrmDate = "01/04/"& Arr(0)
	dToDate = "31/03/"& Arr(1)

	sIssType = Request("ISSTYPE")
    if trim(sIssType)="" then sIssType="SEL"


 sOrgID = Session("organizationcode")
	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sOrgID
		.open
	End With
	if not dcrs.eof then
		sOrgName = trim(dcrs(0))
	end if
	dcrs.close

	sCheck = trim(Request("hCheck"))
	'Response.Write sCheck
	IF trim(sCheck) = "" then sCheck = "M"
'	sOrgID = trim(Request.Form("selUnit"))
'	sOrgName = trim(Request.Form("hOrgName"))
	'dFrmDate = trim(Request.Form("hFrmDate"))
	'dToDate = trim(Request.Form("hToDate"))
	sType = trim(Request.Form("selType"))
	sUsage = trim(Request.Form("hUsage"))
	sCreatedBy = trim(Request.Form("selUser"))
	sItemType = trim(Request.Form("selItemType"))
	if dFrmDate = "" then dFrmDate = FormatDate(date -15)
	if dToDate = "" then dToDate = FormatDate(date)
	if sOrgID = "" then sOrgID = "010101"
	if sItemType = "" then sItemType = "STO"
	'if sSortBy = "" then sSortBy = "D"

	if sType = "select" then sType = ""
	if sCreatedBy = "select" then sCreatedBy = ""

	sAction = StatusEligible
	''''''''''''''''''''' Paging Declaration ''''''''''''''''''''''''''''''''''''''''
    Const iPageSize=15	'How many records to show
    Dim iCurrentPage	'Current Page No.
    Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
    Dim iPageCtr		'Counter
	Dim lnPage

    iCurrentPage = Request.Form("hPageSelection")
    if iCurrentPage = "" or iCurrentPage = "0" then iCurrentPage = "1"
    'iCtr = (Cint(iPageSize) * (iCurrentPage - 1))

    con.CursorLocation = 3

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  onLoad="Init('<%=FormatDate(date)%>')">
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
		<input type=hidden name="hSelected" value="">
		<input type=hidden name="hWhichMRS" value="">
		<input type=hidden name="hAction" value="">
		<input type=hidden name="mrs" value="">
		<input type=hidden name="sAct" value="">
		<input type=hidden name="hUnit" value="<%=sOrgID%>">

		<input type=hidden name="hOrgName" value="<%=sOrgName%>">
		<input type=hidden name="hFrmDate" value="<%=dFrmDate%>">
		<input type=hidden name="hToDate" value="<%=dToDate%>">
		<input type=hidden name="hUsage" value="">
		<input type=hidden name="hCheck" value="<%=sCheck%>">

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr><td height="1px"></td></tr>
			<tr>
				<td class="PageTitle">
				<% if sCheck = "I" then %>
					Material Requisition For Issues
				<% else %>
					Material Requisition
				<% end if %>
				</td>
			</tr>
			<tr>
				<td align="center" class="TopPack">
				</td>
			</tr>
			<tr>
				<td valign="top">
					<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					    <tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" >
								<tr>
								   	<td class="TabCurrentCell" valign="bottom" align="center" width="50">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr>
												<td align="center">List
												</td>
											</tr>
										</table>
									</td>
									<td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
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
										<td align="center" colspan="3" class="MiddlePack" height="7">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
									</tr>
									<tr>
										<td align="center" class="ClearPixel">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
										<td valign="top" width="100%">
											<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
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
															<table border="0" cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="width: 100%; display: none">
																		<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
																			<!--<tr>
																				<td class="FieldCellSub">Unit</td>
																				<td class="FieldCellSub">
																					<select size="1" name="selUnit" class="FormElem">
																						<option value="select">Select</option>
																						<%	'Calling the Function which populates Organization Unit list
																							populateUnitSelected sOrgID
																						%>
																					</select>
																				<td class="FieldCellSub"></td>
																			</tr>-->
																			<!--<tr>
																			    <td class="FieldCellSub">Item Type</td>
																			    <td class="FieldCellSub">
																					<select size="1" name="selItemType" class="FormElem">
																						<option value="select">Select</option>
																						<%	'Calling the Function which populates Item Type list
																						'	populateItemTypeSelected sItemType
																						%>
																				    </select>
																				</td>
																			</tr>-->
																		<% if sCheck <> "Z"	then %>
																			<tr>
																			    <td class="FieldCellSub">Date Range From</td>
																			    <td class="FieldCellSub">
																				  <object id="ctlFromDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"     codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
																						<param name="_ExtentX" value="2355">
																						<param name="_ExtentY" value="529">
																					</object>
																				<%
																					' Function Call to Insert Date Picker
																				'	Response.Write InsertDatePicker("ctlFromDate")
																				%>
																			    </td>
																			    <td class="FieldCellSub"> To</td>
																			    <td class="FieldCellSub">
																					<object id="ctlToDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"     codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
																						<param name="_ExtentX" value="2355">
																						<param name="_ExtentY" value="529">
																					</object>
																				<%
																					' Function Call to Insert Date Picker
																				'	Response.Write InsertDatePicker("ctlToDate")
																				%>
																			    </td>
																			</tr>
																			<!--<tr>
																			    <td class="FieldCellSub">Issue Type</td>
																			    <td class="FieldCellSub">
																					<select size="1" name="selType" class="FormElem">
																						<option value="select">Select</option>
																						<option value="0">Returnable</option>
																						<option value="1">Non returnable</option>
																				    </select>
																				</td>
																			</tr>
																			<tr>
																			    <td class="FieldCellSub">Usage of Item</td>
																			    <td class="FieldCellSub" colspan="3">
																					<select size="8" name="selUsage" class="FormElem" multiple>
																						<%	'Calling the Function which populates the Usage of Item
																							'populateUsage
																						%>
																					</select>
																			    </td>
																			</tr>-->
																			<tr>
																			    <td class="FieldCellSub">CreatedBy</td>
																			    <td class="FieldCellSub" colspan="3">
																					<select size="1" name="selUser" class="FormElem">
																						<option value="select">Select</option>
																						<%	'Calling the Function which populates the User List
																							populateEmployee
																						%>
																					</select>
																			    </td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub">Issue Type</td>
						                                                        <td class="FieldCellSub">
						                                                            <select id="cmbIssType" class="FormElem">
						                                                                <option value="SEL" <%if sIssType="SEL" then Response.write "Selected" %>>Select</option>
						                                                                <%
						                                                                    sQuery = "Select ReceiptIssueTypeCode,ReceiptIssueTypeDesc from APP_M_ReceiptIssueTypes where ApplicableFor in ('B','I')"
						                                                                    dcrs.open sQuery,con
						                                                                    if not dcrs.eof then
						                                                                        do while not dcrs.eof
						                                                                            if trim(sIssType)=trim(dcrs(0)) then
						                                                                                response.write "<option value="& trim(dcrs(0)) &" selected>"& trim(dcrs(1)) &"</option>"
						                                                                            else
						                                                                                response.write "<option value="& trim(dcrs(0)) &">"& trim(dcrs(1)) &"</option>"
						                                                                            end if
						                                                                            dcrs.movenext
						                                                                        loop
						                                                                    end if
						                                                                    dcrs.close
						                                                                %>
						                                                            </select>&nbsp;
						                                                            <input type="button" value="Search" class="ActionButton" onClick="CheckSubmit()">
						                                                        </td>
																			</tr>

																		<% end if %>
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
						<%

						'Response.write "<textarea>"&sCheck&"</textarea>"
						if sCheck = "M"	then %>
							<tr>
								<td></td>
								<td>
								<%
								    Dim sQuery,sIssToStr

    								sQuery = "Select MRSNumber,Convert(varchar,MRSDate,103),EmployeeName,MRSHeaderStatus,IsNull(ApprovedBy,0),MRSISSUESTATUS,IsNulL(MRSCODE,MRSNUMBER),IsNull(AppRefType,''),AppRefNo,IssToType,IssToCode,IssToSubCode,IsNull(IssueTypeCode,'GEN') from VWMRSLIST"
									sQuery = sQuery &" WHERE MRSFORUNIT = " & Pack(sOrgID) & " AND GENERATEDFROM = 4 AND (ISNULL(QUANTITYISSUED,0) + ISNULL(QUANTITYTOPURCHASE,0) + ISNULL(QUANTITYFORTRANSFER,0)) <= ISNULL(QUANTITYAPPROVED,0) AND (MRSHEADERSTATUS IN (040101,040107,040105) OR MRSISSUESTATUS IN (040102,040108,040112,040113,040114,040115,040116,040117,040118,040130))"
									sQuery =sQuery & " Group By MRSNumber,Convert(varchar,MRSDate,103),EmployeeName,MRSHeaderStatus,IsNull(ApprovedBy,0),MRSISSUESTATUS,IsNulL(MRSCODE,MRSNUMBER),IsNull(AppRefType,''),AppRefNo,IssToType,IssToCode,IssToSubCode,IsNull(IssueTypeCode,'GEN') "
								    'Response.write "<textarea>"& sQuery &"</textarea>"

									iCtr = 0
									with dcrs
										.CursorLocation = 3
										.CursorType = 3
										.source =sQuery
										.ActiveConnection = con
										.Open
									end with
									set dcrs.ActiveConnection = nothing
								%>
								</td>
								<td></td>
							</tr>


                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%" align="left">
                                    <table border="0" cellpadding="0" cellspacing="0" width=100%>
                                        <tr>
											<td>
												<DIV class=frmBody style="width: 100%; height:350;">
													<table border="0" cellspacing="1" class="ExcelTable" width="100%">
														<tr>
															<td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
															<td class="ExcelHeaderCell" align="center">MR No. - Date</td>
															<td class="ExcelHeaderCell" align="center">Requested By</td>
															<td class="ExcelHeaderCell" align="center">Requested For</td>
															<td class="ExcelHeaderCell" align="center">Ref Name(No - Date)</td>
															<td class="ExcelHeaderCell" align="center">Created / (Amended or Approved) by</td>
															<td class="ExcelHeaderCell" align="center">Status</td>
															<td class="ExcelHeaderCell" align="center">Action</td>
														</tr>
												<%
										'			if sAction <> "NO" then
														Do While Not dcrs.EOF

															with dcrs1
																.CursorLocation = 3
																.CursorType = 3
																if instr(1,"040101,040107,040105",trim(dcrs(3))) > 0 then
																	sStatus = trim(dcrs(3))
																	.Source = "SELECT EMPLOYEENAME FROM MS_EMPLOYEEMASTER WHERE EMPLOYEENUMBER IN (SELECT AMENDENDBY FROM INV_A_MRSHEADER WHERE MRSNUMBER = " & trim(dcrs(0)) & ")"
																else
																	sStatus = trim(dcrs(5))
																	.Source = "SELECT EMPLOYEENAME FROM MS_EMPLOYEEMASTER WHERE EMPLOYEENUMBER = " & trim(dcrs(4)) & ""
																end if
																.ActiveConnection = con
																'Response.Write dcrs1.Source
																'Response.Write  " sStatus = " & sStatus
																.Open
															end with
															set dcrs1.ActiveConnection = nothing
															if not dcrs1.EOF then
																sAmendedBy = trim(dcrs1(0))
															else
																sAmendedBy = "-"
															end if
															dcrs1.Close

															sIssToStr = IssuedToString(dcrs(9),dcrs(10),dcrs(11))
															sIssTypeCodestr = GetRcptIssName(dcrs(12))

															if trim(dcrs(7))<>"" and trim(dcrs(8))<>"" then
															    sArrRef = split(GetInfoRefType(dcrs(7),dcrs(8),sOrgID),":")
															    if trim(sArrRef(0))<>"" then
															        sReferenceName= sArrRef(0) & " ("& sArrRef(1) &" - "& sArrRef(2) &")"
															    else
															        sReferenceName = "None"
															    end if
															else
															    sReferenceName = "None"
															end if


															iCtr = iCtr + 1
													%>
														<tr>
															<td class="ExcelSerial" align="center" width="30"><%=iCtr%></td>
															<!--td class="ExcelDisplayCell" width="100"><%=trim(dcrs(0))%> - <%=trim(dcrs(1))%></td-->
															<td class="ExcelDisplayCell" width="100">

																<a href="javascript:void(0)" Class="ExcelDisplayLink" onClick="showModalDialog('mrsPopup.asp?MRSNO=<%=trim(dcrs(0))%>','','dialogHeight:470px;dialogwidth=615px;center:yes;help:no;resizable:no;status:no')"><%=trim(dcrs(6))%> - <%=trim(dcrs(1))%></a>
															</td>
															<td class="ExcelDisplayCell"><%=sIssToStr%></td>
															<td class="ExcelDisplayCell"><%=sIssTypeCodestr%></td>
															<td class="ExcelDisplayCell"><%=sReferenceName%></td>
															<td class="ExcelDisplayCell"><%=trim(dcrs(2))%> / <%=sAmendedBy%></td>
															<td class="ExcelDisplayCell" align="left"><%=GetStatus(sStatus)%></td>
															<td class="ExcelInputCell" align="center">
																<select name="selAmendZ<%=trim(dcrs(0))%>" class="FormElem" onChange="RequisitionAction(this)">
																	<option value="select"></option>
																<%	' Function to populate Action
																	populateAction "AA",sStatus,trim(dcrs(0)),sAction
																%>
																</select>
															</td>
														</tr>
												<%
														dcrs.MoveNext
														loop
														dcrs.Close
										'			end if
												%>
									                </table>
												</div>
									        </td>
									    </tr>
									</table>
								</td>
								<td align="center"></td>
                            </tr>
                            <% elseif sCheck = "I"	then %>
							<tr>
								<td></td>
								<td>
								<%
										iCtr = 0
										'Response.Write sType & " +++ "& sUsage &" +++ "& sCreatedBy
										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											'.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS FROM VWMRSLIST WHERE MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118) AND GENERATEDFROM = 4 AND (ISNULL(QUANTITYISSUED,0) + ISNULL(QUANTITYTOPURCHASE,0) + ISNULL(QUANTITYFORTRANSFER,0)) < ISNULL(QUANTITYAPPROVED,0) ORDER BY MRSNUMBER DESC"
											'.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE ISSUEDFORCODE <> 'JWK' AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) ORDER BY MRSNUMBER DESC"
											if sType <> "" and sUsage = "" and sCreatedBy = "" then
												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND MRSTYPE = " & sType & " AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) ORDER BY MRSNUMBER DESC"

											elseif sType = "" and sUsage <> "" and sCreatedBy = "" then

												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND ISSUEDFORCODE IN (" & sUsage & ") AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) ORDER BY MRSNUMBER DESC"

											elseif sType <> "" and sUsage <> "" and sCreatedBy = "" then
												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND MRSTYPE = " & sType & " AND ISSUEDFORCODE IN (" & sUsage & ") AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) ORDER BY MRSNUMBER DESC"

											elseif sType <> "" and sUsage = "" and sCreatedBy <> "" then
												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND MRSTYPE = " & sType & " AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) AND CREATEDBY = " & sCreatedBy & " ORDER BY MRSNUMBER DESC"

											elseif sType = "" and sUsage <> "" and sCreatedBy <> "" then
												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND ISSUEDFORCODE IN (" & sUsage & ") AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) AND CREATEDBY = " & sCreatedBy & " ORDER BY MRSNUMBER DESC"

											elseif sType = "" and sUsage = "" and sCreatedBy <> "" then
												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) AND CREATEDBY = " & sCreatedBy & " ORDER BY MRSNUMBER DESC"

											elseif sType <> "" and sUsage <> "" and sCreatedBy <> "" then
												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND MRSTYPE = " & sType & " AND ISSUEDFORCODE IN (" & sUsage & ") AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) AND CREATEDBY = " & sCreatedBy & " ORDER BY MRSNUMBER DESC"

											else
												.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE (ITEMTYPEID = " & Pack(sItemType) & " OR ITEMTYPEID IS NULL) AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) AND CONVERT(DATETIME,MRSDATE,103) BETWEEN CONVERT(DATETIME," & Pack(dFrmDate) & ",103) AND CONVERT(DATETIME," & Pack(dToDate) & ",103) ORDER BY MRSNUMBER DESC"
											end if
											'.Source = "SELECT DISTINCT MRSNUMBER,CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,EMPLOYEENAME,APPROVEDBY,MRSISSUESTATUS,SOURCEREFNO FROM VWMRSLIST WHERE ISSUEDFORCODE <> 'JWK' AND MRSFORUNIT = " & Pack(sOrgID) & " AND MRSISSUESTATUS IN (040102,040105,040108,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129) AND GENERATEDFROM = 4 AND ISNULL(QUANTITYISSUED,0) < ISNULL(QUANTITYAPPROVED,0) ORDER BY MRSNUMBER DESC"
										'	Response.Write dcrs.Source
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing
								%>
								</td>
								<td></td>
							</tr>
                            <tr>
								<td align="center"></td>
								<td valign="top" width="100%" align="left">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
											<td>
												<DIV class=frmBody style="width: 585; height:320;">
													<table border="0" cellspacing="1" class="ExcelTable" width="100%">
														<tr>
															<td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
															<td class="ExcelHeaderCell" align="center">MR No. - Date</td>
															<td class="ExcelHeaderCell" align="center">Usage of Item</td>
															<!--td class="ExcelHeaderCell" align="center">Requested Unit</td-->
															<td class="ExcelHeaderCell" align="center">Created / Approved by</td>
															<td class="ExcelHeaderCell" align="center">Status</td>
															<td class="ExcelHeaderCell" align="center">Action</td>
														</tr>
													<%
														Do While Not dcrs.EOF
															iCtr = iCtr + 1
															with dcrs1
																.CursorLocation = 3
																.CursorType = 3
																.Source = "SELECT EMPLOYEENAME FROM MS_EMPLOYEEMASTER WHERE EMPLOYEENUMBER = " & trim(dcrs(5)) & ""
																.ActiveConnection = con
																.Open
															end with
															set dcrs1.ActiveConnection = nothing
															if not dcrs1.EOF then
																sAmendedBy = trim(dcrs1(0))
															else
																sAmendedBy = "-"
															end if
															dcrs1.Close
													%>
														<tr>
															<td class="ExcelSerial" align="center" width="30"><%=iCtr%></td>
															<!--td class="ExcelDisplayCell" width="100"><%=trim(dcrs(0))%> - <%=trim(dcrs(1))%></td-->
															<td class="ExcelDisplayCell" width="100">
																<a href="javascript:void(0)" Class="ExcelDisplayLink" onClick="showModalDialog('mrsPopup.asp?MRSNO=<%=trim(dcrs(0))%>','','dialogHeight:470px;dialogwidth=615px;center:yes;help:no;resizable:no;status:no')"><%=trim(dcrs(0))%> - <%=trim(dcrs(1))%> [<%=trim(dcrs(7))%>]</a>
															</td>

															<td class="ExcelDisplayCell"><%=trim(dcrs(2))%></td>
															<!--td class="ExcelDisplayCell"><%=trim(dcrs(3))%></td-->
															<td class="ExcelDisplayCell"><%=trim(dcrs(4))%> / <%=sAmendedBy%></td>
															<td class="ExcelDisplayCell"><p align="left"><%=GetStatus(trim(dcrs(6)))%></td>
															<td class="ExcelInputCell"><p align="center">
																<select name="selIssueZ<%=trim(dcrs(0))%>" class="FormElem" onChange="IssueAction(this)">
																	<option value="select"></option>
																<%	' Function to populate Action
																	populateAction "IS",trim(dcrs(6)),trim(dcrs(0)),sAction
																%>
																</select>
															</td>
														</tr>
													<%	dcrs.MoveNext
														loop
														dcrs.Close
													%>
									                </table>
												</div>
									        </td>
									    </tr>
									</table>
								</td>
								<td align="center">
								</td>
                            </tr>

                            <% end if %>
									<tr>
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
									<tr>
										<td align="center" width="5" class="ClearPixel">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
										<td valign="top">
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
								                        <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
								                        <Input Type=Hidden name="hCtr" Value="<%=iCnt%>" >
								                        <Input Type=Hidden name="hPageSelection" Value="" >
														<%	If iTotPage >= 2 Then
																if iCurrentPage = 1 then
														%>
														<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
														<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
														<%		else	%>
														<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
														<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
    													<%		end if	%>
    													<SELECT class="FormElem" onChange="Paginate(this(this.selectedIndex).value)" id=select1 name=select1>
    													<%
															For lnPage = 1 To iTotPage
																If lnPage = iCurrentPage Then
														%>
															<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotPage%></OPTION>
														<%		else	%>
															<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
    													<%		end if
    														next
    													%>
    													</SELECT>
    													<%
    															if iCurrentPage = iTotPage then
    													%>
														<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
														<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

    													<%		else	%>
														<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
														<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotPage%>')" id=button8 name=button8>
    													<%		end if
															End If
														%>

												</td>

											</tr>
											 <tr>
                                           <table border="0" cellpadding="0" cellspacing="0" width="100%">
									        <tr>
									        <td valign="center" class="ActionCell">
										        <input type=button name='btnMR' value='New Material Requisition' class='ActionButtonX'  onClick='checkMR()'>&nbsp;
										        <!--<input type="button" value="Direct Issue" name="BtnDI" class="ActionButtonX" tabindex="3" onclick = "DirectIssue()">&nbsp;-->
										        <input type="button" value="Close" name="B4" class="ActionButton" onClick="window.location.href='../welcome_Inventory.asp'">
									        </td>
									        </tr>
									        </table>
                                 </tr>

								<tr>
									<td align="center" class="BottomPack" colspan="3"></td>
								</tr>
										</table>
									</td>
									<td align="center" class="ClearPixel" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" class="BottomPack" colspan="3"></td>
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
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR ORDER BY ISSUEDFORCODE"
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

	' Function to populate Action
	Function populateAction(sWho,sStatus,iMRNumber,sAction)
		if sWho = "AA" then
			'Created or Changed or Partially Approved
			if sStatus = "040101" or sStatus = "040107" or sStatus = "040105" then
				if sAction <> "NO" then
					'Amend
					if InStr(1,sAction,"AME") > 0 then
						Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsAmendmentItemEntry.asp"">Amend</OPTION>" &vbcrlf)
					end if
					' Approve
					if InStr(1,sAction,"APP") > 0 then
						Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsApprovalItemSelEntry.asp"">Approve</OPTION>" &vbcrlf)
					end if
					' Cancel
					if InStr(1,sAction,"CAN") > 0 and sStatus <> "040105" then
						Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsCancelDetailsEntry.asp"">Cancel</OPTION>" &vbcrlf)
					end if
				else
					Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsAmendmentItemEntry.asp"">Amend</OPTION>" &vbcrlf)
					Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsApprovalItemSelEntry.asp"">Approve</OPTION>" &vbcrlf)
					Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsCancelDetailsEntry.asp"">Cancel</OPTION>" &vbcrlf)
					'Response.Write("<OPTION VALUE=""O"">On Hold</OPTION>" &vbcrlf)
				end if

			'On Hold
			elseif sStatus = "040130" then
				'Un Hold
				Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsOnHoldDetailsEntry.asp"">Un Hold</OPTION>" &vbcrlf)

			'Approved / Partially Approved
			'Partially Issued
			'Partial Issued / PR
			'Partial Issued / ST
			'Partial PR / ST
			'Partial Issued / PR / ST
			'Purchase Request
			'Request For Transfer
			else
				'Amend
				if InStr(1,sAction,"AME") > 0 then
					Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsManageEntry.asp"">Amend</OPTION>" &vbcrlf)
				end if
				'Close
				if InStr(1,sAction,"CAN") > 0 and sStatus = "040102" then
					Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsCloseDetailsEntry.asp"">Close</OPTION>" &vbcrlf)
				end if
				'Short Close
				if InStr(1,sAction,"SHO") > 0 and sStatus = "040112" then
					Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsShortCloseDetailsEntry.asp"">Short Close</OPTION>" &vbcrlf)
				end if
				'On Hold
				if InStr(1,sAction,"OHO") > 0 and sStatus = "040112" then
					Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsOnHoldDetailsEntry.asp"">On Hold</OPTION>" &vbcrlf)
				end if
			end if
		end if

		if sWho = "IS" then
			'Approved / Partially Approved
			'Partially Issued
			'Partial Issued / PR
			'Partial Issued / ST
			'Partial PR / ST
			'Partial Issued / PR / ST
			'Purchase Request
			'Request For Transfer
			if InStr(1,"040102,040105,040112,040113,040114,040115,040116,040117,040118,040122,040123,040124,040125,040126,040127,040128,040129",sStatus) > 0 then
				Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsIssueItemEntry.asp"">Create</OPTION>" &vbcrlf)
			end if
		end if

		if sWho = "IM" then
			'Approved / Partially Approved
			if sStatus = "040102" or sStatus = "040105" then
				Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsIssueAmendEntry.asp"">Amend</OPTION>" &vbcrlf)
			'Partially Issued
			'Partial Issued / PR
			'Partial Issued / ST
			'Partial PR / ST
			'Partial Issued / PR / ST
			'Purchase Request
			'Request For Transfer
			elseif sStatus = "040112" or sStatus = "040113" or sStatus = "040114" or sStatus = "040115" or sStatus = "040116" or sStatus = "040117" or sStatus = "040118" then
				Response.Write("<OPTION VALUE="""&iMRNumber&"?mrsIssueAmendEntry.asp"">Amend</OPTION>" &vbcrlf)
			end if
		end if

	End Function

	' Function to return back what status is available for the Activity
	Function StatusEligible
		dim sTemp
		dim iSAApplicationPop,iSAProcessPop,iSAActivityPop,iEmpNoPopulate

		iSAApplicationPop = Session("iApplication")
		iSAProcessPop = Session("iProcess")
		iSAActivityPop = Session("iActivity")
		iEmpNoPopulate = Session("employeenumber")

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ACTIONS FROM MS_USERACTIONS WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & " ORDER BY 1"
			'Response.Write dcrs1.Source
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		if not dcrs1.EOF then
			do while not dcrs1.EOF
				sTemp = sTemp & "," & trim(dcrs1(0))
			dcrs1.MoveNext
			loop
		else
			sTemp = ",NO"
		end if
		dcrs1.Close

		sTemp = mid(sTemp,2)
		StatusEligible = sTemp
	End Function
%>
