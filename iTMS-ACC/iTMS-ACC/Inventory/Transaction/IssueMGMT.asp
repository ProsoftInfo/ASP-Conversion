<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	IssueMGMT.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	RAGAVENDRAN
	'Created On					:	APRIL 22,2010
	'Modified On				:	Feb 27,2013
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/Purpopulate.asp" -->
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/PurChkItemSpecPack.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Purchase Receipt Status</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/PurchaseCCIDivClick.Js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></script>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsMgmt.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
'*****************************************************
Function DirectIssue()
sIssType =document.formname.cmbIssType(document.formname.cmbIssType.selectedIndex).value
if UCase(trim(sIssType))="SEL" then sIssType="GEN"
	document.formname.action = "MATERIALISSUEENTRY.ASP?TYPE="&sIssType
	document.formname.submit
End Function

'------------------------------------------------------------------------------------------

Function checkSubmit()
Dim obkChk,sOrgID,sIssDate,sUsage,sUsageDesc,sIType,iRefNo
Dim sIssNo,sTempValues,i
	sOrgID = document.formname.hUnit.Value
	if document.formname.hCtr.value = 0 then exit function
	for i = 1 to document.formname.hCtr.value
		set obkChk = eval("document.formname.rSelect(i-1)")
		If obkChk.Checked = True Then
			sIssNo = document.formname.rSelect(i-1).Value

			sTempValues = sOrgID&":"&sIssNo
			PrintWindow( "../reports/PRNDICreateDetails.asp?sTemp=" + sTempValues)
		End If
	next
	if sTempValues = "" then
		alert("Select any one Issue")
		exit function
	end if
End Function
'------------------------------------------------------------------------------------------
Function fninit()
	dim i,sOrgID
	sOrgID = document.formname.hUnit.Value
	document.formname.hUnit.Value = sOrgID

	document.formname.ctlRcptFromDate.SetDate = document.formname.hFromDate.value
	document.formname.ctlRcptToDate.SetDate = document.formname.hToDate.value

	sIssToType = document.formname.hIssueToType.value
	sIssToCode = document.formname.hIssueToCode.value

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


	If document.formname.hUser.Value <> "" Then
		for i=0 to (document.formname.selUser.length-1)
			if document.formname.selUser.options(i).Value=document.formname.hUser.value then
				document.formname.selUser.selectedIndex=i
			end if
		next
	Else
		document.formname.selUser.selectedIndex = 0
	End If
	if document.formname.hStatus.value="Y" then
		document.formname.rStatus(0).checked=true
	elseif document.formname.hStatus.value="N" then
		document.formname.rStatus(1).checked =true
	elseif document.formname.hStatus.value="A" then
		document.formname.rStatus(2).checked=true
	end if

	document.formname.ctlRcptFromDate.disabled = TRUE
End Function

''------------------------------------------------------------------------------------------
Function Clear()
	clearTblItem
End Function
'------------------------------------------------------------------------------------------
Function clearTblItem
 dim iNum
 For iNum = 1 to document.all.tblDetail.rows.length -1
 	document.all.tblDetail.deleteRow(1)
 Next
End Function
'------------------------------------------------------------------------------------------
Function Validate()
	dim sOrgID,sSupplier
	sOrgID = document.formname.hUnit.value


	If Ucase(document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value) <> "SELECT" Then
		IssVal = document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value
		sArrIssVal = split(IssVal,":")

		if Ubound(sArrIssVal)>0 then
		    document.formname.hIssueToType.value = sArrIssVal(0)
		    document.formname.hIssueToCode.value = sArrIssVal(1)
		else
		    document.formname.hIssueToType.value = sArrIssVal(0)
		end if

	Else
	    document.formname.hIssueToType.value = ""
		document.formname.hIssueToCode.value = ""

	End If

	If document.formname.rStatus(0).checked=true Then
		document.formname.hStatus.value = "Y"
	ElseIf document.formname.rStatus(1).checked=true Then
		document.formname.hStatus.value = "N"
	ElseIf document.formname.rStatus(2).checked=true Then
		document.formname.hStatus.value = "A"
	End If
	If document.formname.selUser.selectedIndex > 0 then
		document.formname.hUser.value =  document.formname.selUser(document.formname.selUser.selectedIndex).Value
	Else
		document.formname.hUser.value = ""
	End If

	document.formname.hFromDate.value = document.formname.ctlRcptFromDate.getDate
	document.formname.hToDate.value = document.formname.ctlRcptToDate.getDate

	sIssType =document.formname.cmbIssType(document.formname.cmbIssType.selectedIndex).value

	document.formname.action =  "IssueMGMT.asp?ACTN=L&ISSTYPE="&sIssType
	document.formname.submit
End Function
'------------------------------------------------------------------------------------------
Function ResetData()

	document.formname.hFromDate.value = ""
	document.formname.hToDate.value = ""
	document.formname.hStatus.value = ""
	document.formname.hUser.value = ""
	document.formname.hIssueToCode.value = ""
	document.formname.hIssueToSubCode.value = ""
	document.formname.hIssueToType.value = ""


	document.formname.action =  "IssueMGMT.asp"
	document.formname.submit
End Function

'------------------------------------------------------------------------------------------
Function DisplayExistingCriteria()
	sImageSrc = ucase(trim(document.formname.ImgSearch.src))

	if instr(1,sImageSrc,"MINUS.GIF") > 0 then



	end if 'if instr(1,sImageSrc,"MINUS.GIF") > 0 then

End Function

Function DoAction(sAction)
Dim i , IssDate
	sOrgID = document.formname.hUnit.Value
	if document.formname.hCtr.value = 0 then exit function
	if document.formname.hCtr.value > 1 then
	    for i = 1 to document.formname.hCtr.value
		    set obkChk = eval("document.formname.rSelect(i-1)")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect(i-1).Value
			    IssDate = Eval("document.formname.hIssDate"&i).value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	    next
	else
	        set obkChk = eval("document.formname.rSelect")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect.Value
			    IssDate = document.formname.hIssDate1.value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	end if 'if document.formname.hCtr.value > 1 then
	if sTempValues = "" then
		alert("Select any one Issue")
		exit function
	Else
		If sAction = "C" Then
			document.formname.action =  "MatConsDetailsEntry.asp?RefDet="&sTempValues
			document.formname.submit
		ElseIf sAction = "R" Then
			document.formname.action =  "IssReturnEntry.asp?RefDet="&sTempValues
			document.formname.submit
		End If
	End If
End Function
'------------------------------------------------------------------------------------------
Function ViewIssDet()
Dim i , IssDate
	sOrgID = document.formname.hUnit.Value
	if document.formname.hCtr.value = 0 then exit function
	if document.formname.hCtr.value > 1 then
	    for i = 1 to document.formname.hCtr.value
		    set obkChk = eval("document.formname.rSelect(i-1)")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect(i-1).Value
			    IssDate = Eval("document.formname.hIssDate"&i).value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	    next
	else
	        set obkChk = eval("document.formname.rSelect")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect.Value
			    IssDate = document.formname.hIssDate1.value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	end if 'if document.formname.hCtr.value > 1 then
	if sTempValues = "" then
		alert("Select any one Issue")
		exit function
	Else
		OutValue = showModalDialog("ViewIssueEntDetail.asp?IssNo="&sIssNo,"","dialogHeight:500px;dialogWidth:600px;Status:No")
	End If
End Function
'------------------------------------------------------------------------------------------
Function EditIssue()
Dim i , IssDate
	sOrgID = document.formname.hUnit.Value
	if document.formname.hCtr.value = 0 then exit function
	if document.formname.hCtr.value > 1 then
	    for i = 1 to document.formname.hCtr.value
		    set obkChk = eval("document.formname.rSelect(i-1)")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect(i-1).Value
			    sIssTypeCode = eval("document.formname.hIssTypeCode"&i).value
			    IssDate = Eval("document.formname.hIssDate"&i).value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	    next
	else
	        set obkChk = eval("document.formname.rSelect")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect.Value
			    IssDate = document.formname.hIssDate1.value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	end if 'if document.formname.hCtr.value > 1 then
	if sTempValues = "" then
		alert("Select any one Issue")
		exit function
	Else
		document.formname.action = "MATERIALISSUEENTRY.ASP?ISSNO="&sIssNo&"&TYPE="&sIssTypeCode
		document.formname.submit
	End If
End Function
'******************************
'------------------------------------------------------------------------------------------
Function DeleteIssue()
Dim i , IssDate
	sOrgID = document.formname.hUnit.Value
	if document.formname.hCtr.value = 0 then exit function
	if document.formname.hCtr.value > 1 then
	    for i = 1 to document.formname.hCtr.value
		    set obkChk = eval("document.formname.rSelect(i-1)")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect(i-1).Value
			    sIssTypeCode = eval("document.formname.hIssTypeCode"&i).value
			    IssDate = Eval("document.formname.hIssDate"&i).value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	    next
	else
	        set obkChk = eval("document.formname.rSelect")
		    If obkChk.Checked = True Then
			    sIssNo = document.formname.rSelect.Value
			    IssDate = document.formname.hIssDate1.value
			    sTempValues = sOrgID&":"&sIssNo&":"&IssDate
		    End If
	end if 'if document.formname.hCtr.value > 1 then


	if sTempValues = "" then
		alert("Select any one Issue")
		exit function
	Else
        document.formname.action = "ViewIssueEntDetail.asp?IssNo="&sIssNo&"&CallFor=D"
	    document.formname.submit
	End If
End Function
</Script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<%

Dim sForUnit,sPartyCode,sPartyType,sPartySubType,sPartyName,objRS,sTransUnit, rsTemp
dim sRefAgnst,sDCDesc,sDCNo,sMarkInv,sType
dim sFromDate,sToDate,sString,sDateFlag,dcrs
Dim sIssType,sIssTypeCode


Dim sRstatus,sUser

Dim sCallFrom,sAction

Dim iCnt,iSql,obRs,obRs1,sSql,iGRNNo,iRcptNo,iInvNo,iInspNo,sGrnCode,sGrnAgainst,sGRNAgainstStr,sReceiptCode,sInvCode
Dim     iTempGrNNo

Dim iIssueNo,sIssCode,sIssdate,sRecBy,sIssueType,iRefNo,sIssuedFor,sDept,dLastPicked
Dim sIssuedToCode,sIssuedToType,sIssuedToSubCode,sIssuedToString,sQuery

Dim sIssToCode,sIssToSubCode,sIssToType,sFinPeriod,sArrFinPeriod

Set objRS = server.CreateObject("Adodb.recordset")
Set rsTemp = server.CreateObject("Adodb.recordset")
set obRs = server.CreateObject("ADODB.RecordSet")
set obRs1 = server.CreateObject("ADODB.RecordSet")
set dcrs = Server.CreateObject("ADODB.RecordSet")



sForUnit = Session("organizationcode")


sFromDate = trim(request("hFromDate"))
sToDate = trim(request("hToDate"))

sRstatus = trim(request("hStatus"))
sUser = trim(request("hUser"))

sIssToCode = trim(Request("hIssueToCode"))
sIssToSubCode = trim(Request("hIssueToSubCode"))
sIssToType = trim(Request("hIssueToType"))

if trim(sFromDate)="" then
    sFinPeriod = session("FinPeriod")
    sArrFinPeriod = split(sFinPeriod,":")
    sFromDate = "01/04/"&sArrFinPeriod(0)
    sToDate = "31/03/"&sArrFinPeriod(1)
end if

sAction = Request("ACTN")

if sCallFrom="" then sCallFrom ="IS"

sIssType = Request("ISSTYPE")
if trim(sIssType)="" then sIssType="SEL"

sString = ""
if sForUnit = "" then sForUnit = "010101"


If sUser <> "" Then
		sString = sString & " Where (IssuedBy = "&sUser&") and organisationCode = '" & sForUnit & "'"
Else
		sString = sString & " Where (IssuedBy is Not Null) and organisationCode = '" & sForUnit & "'"
End If

If trim(sIssToCode) <>"" and trim(sIssToType)<>"" then
    sString = sString & " and IssuedToCode in ('"& sIssToCode &"') and IssuedToType in ('"& sIssToType &"')"
elseif trim(sIssToCode)="" and trim(sIssToType)<>"" then
    sString = sString & " and IssuedToType in ('"& sIssToType &"')"
end if

	sString = sString & " and (CONVERT(DATETIME,IssueDate,103)>= Convert(DATETIME,'"&sFromDate&"',103) and CONVERT(DATETIME,IssueDate,103)<=  Convert(DATETIME,'"&sToDate&"',103)  or CONVERT(DATETIME,IssueDate,103) is Null)"

If sRstatus = "Y" Then
		sString = sString &	" and ISNULL(Returnable,'') <> 'N' "
ElseIf sRstatus = "N" Then
		sString = sString &	" and ISNULL(Returnable,'') = 'N' "
End If


'Response.Write sString

if sPartyCode <> "" then
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = "Select PartyName from APP_M_PartyMaster where PartyCode="&sPartyCode&" "
		.Open
	end with
	set rsTemp.ActiveConnection = nothing
	if not rsTemp.EOF then
		sPartyName = rsTemp(0)
	end if
	rsTemp.Close
end if

''''''''''''''''''''' Paging Declaration ''''''''''''''''''''''''''''''''''''''''
Const iPageSize=12	'How many records to show
Dim iCurrentPage	'Current Page No.
Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
Dim iPageCtr		'Counter
Dim lnPage
iCurrentPage = 0
if Request.Form("hPageSelection") <> "" then iCurrentPage = CInt(Request.Form("hPageSelection"))

con.CursorLocation = 3

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>


<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="fninit()">
	<form method="POST" name="formname" action="">
	<input type="hidden" name="hPage">
	<input type="hidden" name="hUnit" value="<%=sForUnit%>">

	<input type="hidden" name="hFromDate" value="<%=sFromDate%>">
	<input type="hidden" name="hToDate" value="<%=sToDate%>">
	<input type="hidden" name="hStatus" value="<%=sRstatus%>">
	<input type="hidden" name="hUser" value="<%=sUser%>">
	<input type="Hidden" name="hIssueToCode" value="<%=sIssToCode%>">
	<input type="Hidden" name="hIssueToSubCode" value="<%=sIssToSubCode%>">
	<input type="hidden" name="hIssueToType" value="<%=sIssToType%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				List of Issues
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
							<table border="0" cellpadding="0" cellspacing="0">
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
											<tr><a href="MaterialIssueEntry.asp">
												<td align="center">Basic
												</td></a>
											</tr>
										</table>
									</td>
									<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="IssReturnEntry.asp">
											    <td align="center">Return
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
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable" >
											<tr>
												<td>
													<div>
														<table class="CollapseBand" cellspacing="0" cellpadding="0" >
															<tr>
																<td valign="center"><a style="width: 1em; height: 1em;" title onclick="Div_OnClick(idUnprocessed,'');DisplayExistingCriteria()">
																	<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
																	</a>
																</td>
																<td valign="center" class="SubTitle">&nbsp;&nbsp;
																</td>
															</tr>
														</table>
														<table border="0" width="100%" cellpadding="0" cellspacing="0" class="BodyTable">
															<tr>
																<td width="100%">
																	<div align="left" id="idUnprocessed" style="width : 100%; display: none">
																		<table border="0" cellpadding="0" cellspacing="0">

																			<tr>
																			    <td class="FieldCellSub"> Issue To</td>
																			    <td class="FieldCellSub">
																					<select size="1" name="selIssueTo" class="FormElem" onchange="popIssueToWithOutSubLevel()">
																					<option value="select">Select</option>
																						<%	'Calling the Function which populates the Issued To From Common Functions
																							populateIssueToSelWithOutSubLevel sForUnit
																						%>
																					</select>
																			    </td>
																			    <td class="FieldCellSub">
																				</td>
																			    <td class="FieldCellSub">Created By</td>
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
																				<td class="FieldCellSub">

																					Date Range From
																				</td>
																				<td class="FieldCellSub">
																					<object id="ctlRcptFromDate" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
																						<param name="_ExtentX" value="2355">
																						<param name="_ExtentY" value="529">
																					</object>
																				</td>
																				<td class="FieldCellSub">
																				</td>
																				<td class="FieldCellSub">To
																				</td>
																				<td class="FieldCellSub">
																					<object id="ctlRcptToDate" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
																						<param name="_ExtentX" value="2355">
																						<param name="_ExtentY" value="529">
																					</object>
																				</td>
																			</tr>

																			<tr>
																				<td class="FieldCellSub" valign="top">Returnable
																				</td>
																				<td class="FieldCell">
																					<input type="radio" name="rStatus" value="Y">Yes
																					<input type="radio" name="rStatus" value="N">No
																					<input type="radio" name="rStatus" value="B" checked>All&nbsp;
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
						                                                            </select>
						                                                        </td>
																				<td class="FieldCell">
																					<input type="button" value="Go" name="ButGo" class="ActionButton" onclick="Validate()">
 																					<input type="button" value="Reset" name="ButReset" class="ActionButton" onclick="ResetData()">
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
									<td align="center" class="ClearPixel">
								        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								    </td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7" width="700">
									</td>
								</tr>

								<tr>
								    <td align="center" class="ClearPixel">
								        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								    </td>
									<td valign="top" >
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="top">
													<div class="frmbody" id="frm3">
														<table id="tblDetail" border="0" cellspacing="1" class="ExcelTable" width="100%">
															<tr>
																<td width="10" align="center" class="ExcelHeaderCell">
																	<p align="center">S.No.
																</td>
																<td align="center" class="ExcelHeaderCell">
																</td>
															<%if sCallFrom="IS" then%>
																<td align="center" class="ExcelHeaderCell">Issue Number-Date</td>
																<td align="center" class="ExcelHeaderCell">Issue To</td>
																<td align="center" class="ExcelHeaderCell">Received By
																</td>
																<td align="center" class="ExcelHeaderCell">Returnable
																</td>
																<td align="center" class="ExcelHeaderCell">Reference Type
																</td>
																<td align="center" class="ExcelHeaderCell">Reference No - Date
																</td>
															<%elseif (sCallFrom="TP" or sCallFrom="TI") then%>
																<td width="120" align="center" class="ExcelHeaderCell">MR Number-Date</td>
																<td width="160" align="center" class="ExcelHeaderCell">Issued For</td>
																<td width="160" align="center" class="ExcelHeaderCell">Approved By</td>
																<td  align="center" class="ExcelHeaderCell">IssueType</td>
															<%elseif sCallFrom="PI" then%>
																<td width="120" align="center" class="ExcelHeaderCell">MR Code - Date</td>
																<td width="160" align="center" class="ExcelHeaderCell">Issued For</td>
																<td width="160" align="center" class="ExcelHeaderCell">Marked On</td>
																<td  align="center" class="ExcelHeaderCell">Last Picked</td>
															<%end if%>

															</tr>

											<%


													if sRefAgnst = "" then sRefAgnst = "11"


												IF sCallFrom="IS" THEN

													sSql = "Select IssueEntryNo,isNull(IssueEntryCode,IssueEntryNo),Convert(VarChar,IssueDate,103),MaterialReceivedBy,referencetype,0,IssuedToType,IssuedToCode,IssuedToSubCode,Returnable,AppRefType,AppRefNo,IssueTypeCode From INV_T_MaterialIssueHeader " & sString

													if trim(sIssType)<>"" and trim(sIssType)<>"SEL"  then
													    sSql = sSql & " and IssueTypeCode ='"& sIssType &"'"
													end if

													sSql = sSql &" order by issueentryno desc"
													'Response.Write "<textarea>"& sSql &"</textarea>"

													With objRS
														.CursorLocation = 3
														.CursorType = 3
														.Source = sSql
														.ActiveConnection = con
														.Open
													End with

													Set objRS.ActiveConnection = Nothing
													Set iIssueNo = objRS(0)
													set sIssCode = objRS(1)
													Set sIssdate  = objRS(2)
													Set sRecBy  = objRS(3)
													set sIssueType =objRS(4)
													Set iRefNo = objRS(5)
													iTempGrNNo = 0
												If Not objRS.EOF then

												'''''''''''''''''''''''''''''''''''''''''''''''''''''''
   													objRS.PageSize = iPageSize
													If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
													objRS.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
													iTotPage = objRS.PageCount					'stores total no. of pages
												'''''''''''''''''''''''''''''''''''''''''''''''''''''''
													For iPageCtr = 1 to objRS.PageSize
														iTempGrNNo = iTempGrNNo + 1


														sIssuedToType = objrs("IssuedToType")
														sIssuedToCode = objrs("IssuedToCode")
														sIssuedToSubCode = objrs("IssuedToSubCode")
														sIssuedToString=IssuedToString(sIssuedToType,sIssuedToCode,sIssuedToSubCode)
														sIssTypeCode = objRS("IssueTypeCode")





											%>
															<tr>
																<input type="hidden" name="hRefNo<%=iTempGrNNo%>" value="<%=iRefNo%>" >
																<input type="hidden" name="hIssDate<%=iTempGrNNo%>" value="<%=sIssdate%>" >
																<input type="hidden" name="hIssType<%=iTempGrNNo%>" value="<%=sIssueType%>" >
																<input type="hidden" name="hIssTypeCode<%=iTempGrnNo%>" value="<%=sIssTypeCode%>" />
																<td class="ExcelSerial" align="center" width="31"><%=iTempGrNNo%></td>
																<td class="ExcelDisplayCell" width="10">
																	<input type="radio" name="rSelect" value="<%=iIssueNo%>">
																</td>
																<td class="ExcelDisplayCell">
																    <p align="center">
																    <% If objRS(10)="11" then
    																    response.Write "<font color=Red>*</font>"
																       End If %>
																    <% if sIssdate <> "" Then %>
																    <%=sIssCode%> - <%=sIssdate%>
																	<%Else%>
																	<p align="center"><%=sIssCode%>
																	<%End If%>
																	&nbsp;&nbsp;

																	<a>
																	</a></p>
																</td>
																<td class="ExcelDisplayCell">
																	<%=sIssuedToString%>
																</td>
																<td class="ExcelDisplayCell">
																	<p align="Left"><%=sRecBy%>
																	&nbsp;</p>
																</td>
																<td class="ExcelDisplayCell" align="center">
																	<%
																	    if trim(objrs("Returnable"))="Y" then
																	        response.write "Yes"
																	    else
																	        response.write "No"
																	    end if

																	%>
																</td>
																<%
																Dim sRefValues,sArrRefValues,sRefName,sRefCode,sRefDate
																if trim(objrs("AppRefType"))<>"" then
																    sRefValues = GetInfoRefType(objrs("AppRefType"),objrs("AppRefNo"),sForUnit)
																    sArrRefValues = split(sRefValues,":")

																    sRefName = sArrRefValues(0)
																    sRefCode = sArrRefValues(1)
																    sRefDate= sArrRefValues(2)
																else
																    sRefName = "None"
																    sRefCode = ""
																    sRefDate = ""
																end if
																%>
																<td class="ExcelDisplayCell" align="center">
																	<%=sRefName%>
																</td>
																<td class="ExcelDisplayCell" align="center">
																	<%=sRefCode%>-<%=sRefDate%>
																</td>
															</tr>

											<%
														Response.Flush
													'	End If
														objRS.MoveNext
														If objRS.EOF Then Exit For


														next
														'Loop
													end if
													objRS.Close
												END IF
											%>

														</table>
													</div>
												</td>
											</tr>

										</table>
									</td>
									<td align="center" class="ClearPixel">
								        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								    </td>
								</tr>
								<tr>
								    <td align="center" class="ClearPixel">
								        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								    </td>
								    <td align="left" class="FieldCell">
								        <table width="100%">
								            <tr>
								                <td class="FieldCell">
								                    <%if sCallFrom="IS" then%>
							    	                &nbsp;&nbsp;<font color=Red>*</font> Indicate MR based Issue
							                        <%end if  %>
								                </td>
								                <td class="FieldCell" >
								                    <p align="right">
									                    <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
                                                        <Input Type=Hidden name="hCtr" Value="<%=iTempGrNNo%>" >
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
								        </table>
									</td>
									<td align="center" class="ClearPixel">
								        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								    </td>
								</tr>
								<tr>
								    <td align="center" colspan="3" class="MiddlePack"></td>
							    </tr>
							    <tr>
							        <td align="center" class="ClearPixel">
							            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							        </td>
								    <td valign="top">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%">
										    <tr>
											    <td valign="middle" class="ActionCell" align=center width=761>
											    <%if sCallFrom="IS" then%>
											        <%if trim(sAction)="P" then %>
											            <input type="button" value="Print" name="BtnPrint" class="ActionButton" onclick="CheckSubmit()">
											        <%else %>
												        <input type="button" value="New Issue" name="BtnDI" class="ActionButtonX" tabindex="3" onclick = "DirectIssue()">&nbsp;
												        <input type="button" value="Print" name="BtnPrint" class="ActionButton" onclick="CheckSubmit()">
												        <input type="button" value="Edit" name="BtnEdit" class="ActionButton" onclick="EditIssue()">
												        <input type="button" value="Delete" name="BtnDelete" class="ActionButton" onclick="DeleteIssue()">
												        <input type="button" value="Cancel" name="BtnCancel" class="ActionButton" onclick="CancelIssue()">
												        <input type="button" value="Issue Return" name="BtnIssue" class="ActionButton2" onclick="DoAction('R')">
												        <input type="button" value="View Details" name="BtnView" class="ActionButton2" onclick="ViewIssDet()">
												    <%end if %>
                                                <%end if%>
											    </td>

										    </tr>
									    </table>
								    </td>
								    <td align="center" class="ClearPixel">
							            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							        </td>
							    </tr>
								<tr>
									<td align="center" class="MiddlePack" colspan="3">
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
</body>
</html>

<%
Function getUserName(sUserId)
Dim dcrs1
set dcrs1 = server.CreateObject("ADODB.RecordSet")
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT EMPLOYEENAME FROM MS_EMPLOYEEMASTER WHERE EMPLOYEENUMBER = " & sUserId & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing
	if not dcrs1.EOF then
		getUserName = trim(dcrs1(0))
	else
		getUserName = "-"
	end if
	dcrs1.Close
End Function
%>
