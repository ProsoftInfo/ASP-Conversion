<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	mrsIssuePickList.asp
	'Module Name				:	INVENTORY (Requisition List)
	'Author Name				:
	'Created On					:
	'Modified By				:	Ragavendran R
	'Modified On				:	Feb 27,2013
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
<title>MR Issue Pick</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/printwindow.js"></script>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsMgmt.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
'*****************************************
Function ResetData()
    document.formname.hFromDate.value =""
    document.formname.hToDate.value = ""
    document.formname.hIssueToType.value = ""
    document.formname.hIssueToCode.value =""
    document.formname.hIssueToSubCode.value =""
    document.formname.action = "mrsIssuePickList.asp"
    document.formname.submit
End Function
'****************************************
Function Validate()
    document.formname.hFromDate.value = document.formname.ctlFromDate.getdate()
    document.formname.hToDate.value = document.formname.CtlToDate.getdate()
    document.formname.action = "mrsIssuePickList.asp"
    document.formname.submit
End Function
'*************************************
Function PickBag(sIssueEntryNo)
    document.formname.action = "mrsIssuePickItemList.asp?IssueNo="&sIssueEntryNo
    document.formname.submit
End Function
'***************************************
Function PickBagSchedule(sIssueEntryNo,sScheduledNo)
    document.formname.action = "mrsIssuePickItemList.asp?IssueNo="&sIssueEntryNo&"&ScheduleNo="&sScheduledNo
    document.formname.submit
End Function
'***************************************
Function CheckPick()
Dim sValue
    if document.formname.radPick(0).checked = true then
        sValue = document.formname.radPick(0).value
    else
        sValue = document.formname.radPick(1).value
    end if
    document.formname.hCallFrom.value = sValue
    document.formname.submit
End Function
'*************************************
Function Init(dDate)
	document.formname.ctlFromDate.SetDate = document.formname.hFromDate.value
	document.formname.ctlToDate.SetDate = document.formname.hToDate.value

	sIssToType = document.formname.hIssueToType.value
	sIssToCode = document.formname.hIssueToCode.value
	'alert(sIssToType&":"&sIssToCode)
	if trim(sIssToType)<>"" and trim(sIssToCode)<>"" then
	    for iCnt = 0 to (document.formname.selIssueTo.length-1)
	        if lcase(document.formname.selIssueTo.options(iCnt).value) = lcase(trim(sIssToType)&":"&trim(sIssToCode)) then
	            document.formname.selIssueTo.selectedIndex=iCnt
	        end if
	    next
	elseif trim(sIssToType)<>"" and trim(sIssToCode)="" then
	    for iCnt = 0 to (document.formname.selIssueTo.length-1)
	        if lcase(document.formname.selIssueTo.options(iCnt).value) = lcase(trim(sIssToType)) then
	            document.formname.selIssueTo.selectedIndex=iCnt
	        end if
	    next
	else
	    document.formname.selIssueTo.selectedIndex = 0
	end if
end Function
'********************************************************************************************
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

		document.formname.action = "mrsIssuePickList.asp"
		document.formname.hOrgName.value = document.formname.hOrgName.value

			document.formname.hFrmDate.value = document.formname.ctlFromDate.GetDate()
			document.formname.hToDate.value = document.formname.ctlToDate.GetDate()
			sUsage = ""
			l=document.formname.selUsage.length
			for ii=0 to l-1
				if document.formname.selUsage.options(ii).selected then
					sUsage = sUsage & "'" & document.formname.selUsage.options(ii).value & "',"
				end if
			next
			if sUsage <> "" then
				sUsage = left(sUsage,len(sUsage)-1)
			end if


		document.formname.submit()

End Function

</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<%
	dim dcrs,dcrs1,iCtr,sCheck,sOrgID,sOrgName
	dim dFrmDate,dToDate
	dim arrFin,sFinFrom,sFinTo,sCallFrom,sSql
	Dim iIssueEntryNo,sIssCode,dIssue,sIssType,sPartyCode,sAppRefType,sAppRefNo,sPartyName
	Dim sMRSCode,sMrsNumber,dMRS
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")


	dim sFinPeriod,Arr
	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	sFinFrom = "01/04/"& Arr(0)
	sFinTo = "31/03/"& Arr(1)

	sCheck = trim(Request("hCheck"))
	IF trim(sCheck) = "" then sCheck = "M"

	sOrgID = Session("organizationcode")
	sCallFrom = Request("hCallFrom")
	if trim(sCallFrom)="" or isnull(sCallFrom) then sCallFrom="T"
	with dcrs
		.cursorlocation = 3
		.cursortype = 3
		.activeconnection = con
		.source = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sOrgID
		.open
	end with
	if not dcrs.eof then
		sOrgName = trim(dcrs(0))
	end if
	dcrs.close

	dFrmDate = trim(Request.Form("hFromDate"))
	dToDate = trim(Request.Form("hToDate"))

	sIssToCode = trim(Request("hIssueToCode"))
	sIssToSubCode = trim(Request("hIssueToSubCode"))
	sIssToType = trim(Request("hIssueToType"))


	if dFrmDate = "" then dFrmDate = sFinFrom
	if dToDate = "" then dToDate = sFinTo
	if sOrgID = "" then sOrgID = "010101"

%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  onLoad="Init('<%=FormatDate(date)%>')">
	<form method="POST" name="formname" >
		<input type=hidden name="hUnit" value="<%=sOrgID%>">
		<input type=hidden name="hFromDate" value="<%=sFinFrom%>">
		<input type=hidden name="hToDate" value="<%=sFinTo%>">
		<input type=hidden name="hOrgName" value="<%=sOrgName%>">

		<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

		<input type="Hidden" name="hIssueToCode" value="<%=sIssToCode%>">
	    <input type="Hidden" name="hIssueToSubCode" value="<%=sIssToSubCode%>">
	    <input type="hidden" name="hIssueToType" value="<%=sIssToType%>">

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr><td height="1px"></td></tr>
			<tr>
				<td class="PageTitle">
					MR Issue Pick List
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
							<td class="TabBodyWithTopLine">
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
																	    <input type="radio" name="radPick" value="T" <%if trim(sCallFrom)="T" then Response.write "Checked" %> onclick="CheckPick()">To Pick&nbsp;&nbsp;
																	    <input type="radio" name="radPick" value="P" <%if trim(sCallFrom)="P" then Response.write "Checked" %> onclick="CheckPick()">Picked

																	</td>
																</tr>
															</table>
															<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="width: 575; display: none">
																		<table cellpadding="0" cellspacing="0" border="0">

																		<% if sCheck <> "Z"	then %>
																			<tr>
																			    <td class="FieldCellSub">Date Range From</td>
																			    <td class="FieldCellSub">
																					<object id="ctlFromDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"      codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
																						<param name="_ExtentX" value="2355">
																						<param name="_ExtentY" value="529">
																					</object>
																				<%
																					' Function Call to Insert Date Picker
																					'Response.Write InsertDatePicker("ctlFromDate")
																				%>
																			    </td>
																			    <td class="FieldCellSub"> To</td>
																			    <td class="FieldCellSub">
																					<object id="ctlToDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"      codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
																						<param name="_ExtentX" value="2355">
																						<param name="_ExtentY" value="529">
																					</object>
																				<%
																					' Function Call to Insert Date Picker
																					'Response.Write InsertDatePicker("ctlToDate")
																				%>
																			    </td>
																			</tr>
																			<tr>
																			    <td class="FieldCellSub"> Issue To</td>
																			    <td class="FieldCellSub">
																					<select size="1" name="selIssueTo" class="FormElem" onchange="popIssueToWithOutSubLevel()">
																					<option value="select">Select</option>
																						<%	'Calling the Function which populates the Issued To From Common Functions
																							populateIssueToSelWithOutSubLevel sOrgID
																						%>
																					</select>
																			    </td>
																			    <td class="FieldCellSub">
																			        <input type="button" name="btnGO" value="GO" onclick="Validate()" class="ActionButton">
																			        <input type="button" name="btbReset" value="Reset" onclick="ResetData()" class="ActionButton">
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
												<DIV class=frmBody style="width: 100%; height:425;">
												    <%if trim(sCallFrom)="T" then %>
													    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
													        <tr>
															        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
															        <td class="ExcelHeaderCell" align="center">Issue No - Date</td>
															        <td class="ExcelHeaderCell" align="center">Issued To</td>
															        <td class="ExcelHeaderCell" align="center">Reference Name</td>
															        <td class="ExcelHeaderCell" align="center">Reference No - Date</td>
													        </tr>
													        <tr>
															        <td class="ExcelHeaderCell" align="center"></td>
															        <td class="ExcelHeaderCell" align="center" colspan="2">Item Description</td>
															        <td class="ExcelHeaderCell" align="center">Scheduled On</td>
															        <td class="ExcelHeaderCell" align="center">Scheduled Quantity</td>
													        </tr>
												            <%
												                Dim sIssToString,sIssToCode,sIssToType,sIssToSubCode

														            sSql = "Select IssueEntryNo,OrganisationCode,IssueEntryCode,Convert(varchar,IssueDate,103),IssuedToCode,IssuedToType,MarkPackFlag,AppRefType,AppRefNo,IssuedToSubCode from INV_T_MaterialIssueHeader where IssueType = 'M' and OrganisationCode = "& pack(sOrgID) &" and Convert(datetime,IssueDate,103) Between Convert(datetime,'"& dFrmDate &"',103) and Convert(datetime,'"& dToDate &"',103) "
            														sSql = sSql & " and IssueEntryNo in (Select distinct IssueEntryNo from INV_T_MaterialIssuedForPick where QuantityForPick > QuantityPicked) "

            														If trim(sIssToCode) <>"" and trim(sIssToType)<>"" then
                                                                        sSql = sSql & " and IssuedToCode in ('"& sIssToCode &"') and IssuedToType in ('"& sIssToType &"')"
                                                                    elseif trim(sIssToCode)="" and trim(sIssToType)<>"" then
                                                                        sSql = sSql & " and IssuedToType in ('"& sIssToType &"')"
                                                                    end if


														          'Response.write "<textarea>"&sSql &"</textarea>"
														            with dcrs
															            .CursorLocation = 3
															            .CursorType = 3

															            .Source = sSql

															            .ActiveConnection = con
															            .Open
														            end with
														            set dcrs.ActiveConnection = nothing
														            if not dcrs.EOF then
															            Do While Not dcrs.EOF
																            iCtr = iCtr + 1
																            sMRSCode=""
																            dMRS=""
																            iIssueEntryNo = dcrs(0)
																            sIssCode = dcrs(2)
																            dIssue = dcrs(3)

																            sAppRefType = dcrs(7)
																            sAppRefNo = dcrs(8)

																            sIssToCode = dcrs(4)
																            sIssToType = dcrs(5)
																            sIssToSubCode = dcrs(9)

																            sIssToString = IssuedToString(sIssToType,sIssToCode,sIssToSubCode)

																            if lcase(trim(sIssType))=lcase("Party") then
																                sPartyCode = sIssToCode
																            end if

																            if trim(sPartyCode)<>"" and trim(sPartyCode)<>"0" then
																                sSql = "Select PartyName from APP_M_PartyMaster where PartyCode ="& sPartyCode
																                dcrs1.open sSql,con
																                if not dcrs1.eof then
																                    sPartyName = trim(dcrs1(0))
																                end if
																                dcrs1.close
																            end if 'if trim(sPartyCode)<>"" and trim(sPartyCode)<>"0" then

																            if trim(sAppRefType)="11" then
																                sSql = "Select MRSCode,Convert(varchar,MRSDate,103) from INV_T_MRSHeader where MRSNumber = "& sAppRefNo
																                dcrs1.open sSql,con
																                if not dcrs1.eof then
																                    sMRSCode = dcrs1(0)
																                    dMRS = dcrs1(1)
																                end if
																                dcrs1.close
																            end if

																            Dim sRefValues,sArrRefValues,sRefName,sRefCode,sRefDate
																            if trim(sAppRefType)<>"" then
																                sRefValues = GetInfoRefType(sAppRefType,sAppRefNo,sOrgID)
																                sArrRefValues = split(sRefValues,":")

																                sRefName = sArrRefValues(0)
																                sRefCode = sArrRefValues(1)
																                sRefDate= sArrRefValues(2)
																            else
																                sRefName = "None"
																                sRefCode = ""
																                sRefDate = ""
																            end if

																            sSql = "Select ItemCode,ClassificationCode,ItemAttributes,ScheduleNo,Convert(varchar,Scheduledon,103),ScheduledQty from Inv_T_IssueForPickSchedule where IssueEntryNo = "& iIssueEntryNo
													                        dcrs1.open sSql,con
													                        if dcrs1.eof then
													                            %>
														                            <tr>
															                            <td class="ExcelSerial" align="center"><%=iCtr%></td>
															                            <td class="ExcelDisplayCell"><a href="#" class="ExcelDisplayLink" onclick="PickBag('<%=iIssueEntryNo%>')"><%=trim(sIssCode)%>-<%=trim(dIssue)%></a></td>
															                            <td class="ExcelDisplayCell" align="center"><%=trim(sIssToString)%></td>
															                            <td class="ExcelDisplayCell" align="center"><%=sRefName%></td>
															                            <td class="ExcelDisplayCell" align="center">
															                                <%=sRefCode%>-<%=sRefDate%>
															                            </td>
														                            </tr>
													                            <%
													                        else
													                            %>
														                            <tr>
															                            <td class="ExcelSerial" align="center"><%=iCtr%></td>
															                            <td class="ExcelDisplayCell"><%=trim(sIssCode)%>-<%=trim(dIssue)%></td>
															                            <td class="ExcelDisplayCell"><%=trim(sIssToString)%></td>
															                            <td class="ExcelDisplayCell" align="center"><%=sRefName%></td>
															                            <td class="ExcelDisplayCell" align="center">
															                                <%=sRefCode%>-<%=sRefDate%>
															                            </td>
														                            </tr>
													                            <%
													                                do while not dcrs1.eof
													                                    %>
													                                        <tr>
															                                    <td class="ExcelSerial" align="center"></td>
															                                    <td class="ExcelDisplayCell" colspan="2"><%=GetItemDesc(dcrs1(0),dcrs1(1))%></td>
															                                    <td class="ExcelDisplayCell"><a href="#" class="ExcelDisplayLink" onclick="PickBagSchedule('<%=iIssueEntryNo%>','<%=dcrs1(3)%>')"><%=dcrs1(4)%></a></td>
															                                    <td class="ExcelDisplayCell" align="right"><%=dcrs1(5)%></td>
														                                    </tr>
													                                    <%
													                                    dcrs1.movenext
													                                loop
													                        end if
													                        dcrs1.close


															            dcrs.MoveNext
															            loop
														            end if
														            dcrs.Close
													            %>
									                    </table>
									                <%else ' if trim(sCallFrom)="P" then %>
									                    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
													        <tr>
															        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
															        <td class="ExcelHeaderCell" align="center">Pick Number - Date</td>
															        <td class="ExcelHeaderCell" align="center">Issue No - Date</td>
													        </tr>
												            <%

												                    sSql = "Select PickNumber,Convert(varchar,PickedOn,103),IssueEntryNo,OrganisationCode from INV_T_IssuePick where Convert(datetime,PickedOn,103) Between Convert(datetime,'"& dFrmDate &"',103) and Convert(datetime,'"& dToDate &"',103) "
												                    'Response.write"<textarea>"& sSql &"</textarea>"
												                    dcrs.open sSql,con
												                    if not dcrs.EOF then
															            Do While Not dcrs.EOF
																            iCtr = iCtr + 1
																            sSql = "Select IsNull(IssueEntryCode,IssueEntryNo),Convert(varchar,IssueDate,103) from INV_T_MaterialIssueHeader where IssueEntryNo="& trim(dcrs(2))
																            dcrs1.open sSql,con
																            if not dcrs1.eof then
																                sIssCode = trim(dcrs1(0))
																                dIssue = trim(dcrs1(1))
																            end if
																            dcrs1.close
													            %>
														            <tr>
															            <td class="ExcelSerial" align="center"><%=iCtr%></td>
															            <td class="ExcelDisplayCell" align="left"><a href="IssuePickedDetails.asp?PickNo=<%=trim(dcrs(0))%>" class="ExcelDisplayLink"><%=dcrs(0)%> - <%=dcrs(1)%></a></td>
															            <td class="ExcelDisplayCell"><%=trim(sIssCode)%> - <%=trim(dIssue)%></td>
														            </tr>
													            <%
															            dcrs.MoveNext
															            loop
														            end if
														            dcrs.Close
													            %>
									                    </table>
									                <%end if %>
												</div>
									        </td>
									    </tr>
									</table>
								</td>
								<td align="center"></td>
                            </tr>

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
Function GetItemDesc(ItemCode,ClassCode)
    Dim dcrs,sSql
    set dcrs = server.createObject("ADODB.Recordset")
    sSql =  "Select ItemDescription from INV_M_ItemMaster where ItemCode = "& ItemCode &" and ClassificationCode = "& ClassCode
    dcrs.open sSql,con
    if not dcrs.eof then
        GetItemDesc = trim(dcrs(0))
    end if
    dcrs.close
End Function
%>
