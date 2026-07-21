
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
'Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassServiceAccountEntry.asp
	'Module Name				:	Gate Pass - Service
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	April 10, 2006
	'Modified By				:	Ragavendran R
	'Modified On				:	Sep 21,2010
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include File="../../include/sessionVerify.asp" -->
<!--#include File="../../include/populate.asp" -->
<!--#include File="../../include/purpopulate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Purchase Receipt - Service Receipt Status</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData"><ROOT></ROOT></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/PurchaseCCIDivClick.Js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,iCounter1,ssupcode,iser,iTempCode,Partyname,iRcptNo
j = 0
iser = 0
iCounter1 = 0
'******************************************************
Function btnServiceBill_Click()
	document.formname.action = "GATEPASSSERVICEENTRY.ASP?SerType=I"
	document.formname.submit
End Function
'******************************************************
Function popSuppAgent()

	'OutValue = showModalDialog("../../sales/transaction/PartySelectionTrans.asp","","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'arrTemp = split(OutValue,":")

	'while UBound(arrTemp) = 0
	'	OutValue = showModalDialog("../../sales/transaction/PartySelectionTrans.asp?"&OutValue,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'	arrTemp = split(OutValue,":")
	'wend
	'if UBound(arrTemp) <= 1 then exit function

	'ssupcode = arrTemp(1)
	'Partyname = arrTemp(0)

	sForUnit = document.formname.cmbUnit.value+":O"
	nFlag = 2
	'note : value for hSelectMode is M( Multiple ) / S(Single)
	set OutValue=showModalDialog("SupplierSelect.asp?OrgId="+sForUnit+"&hSelectMode=S&Flag="+cstr(nFlag) & "&OrderTo=S",OutData,"status:no")
	'msgbox OutValue.xml


	sAct = UCase(trim(OutValue.getAttribute("Action")))
	sQuery = trim(OutValue.getAttribute("PassQuery"))
	if ucase(trim(sAct)) <> "CLOSE" then
		do while sAct <> "DONE"
			set OutValue=showModalDialog("SupplierSelect.asp?" & sQuery,OutData,"status:no")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			sQuery = trim(OutValue.getAttribute("PassQuery"))

			if ucase(Trim(sAct)) = "CLOSE" then exit do
		loop
	end if 'if ucase(trim(sAct)) <> "CLOSE" then
	'alert(OutValue.xml)

	If not OutValue.hasChildNodes Then 	exit function


	For each Node2 in OutValue.childNodes
		if ucase(Node2.nodename) = ucase("Supplier") then

			Partyname		= Node2.getAttribute("SuppName")
			'ssupcode	= trim(Node2.getAttribute("SuppCode")) & "#" & trim(Node2.getAttribute("AgentCode"))
			ssupcode	= trim(Node2.getAttribute("SuppCode"))

		end if
	Next


	document.formname.txtRefName.value = Partyname
	document.formname.txtSupplier.value = Partyname
	document.formname.hSupplierCode.value = sSupCode
End Function

Function clearXML()
	Set Root = OutData.documentElement
	For Each HeaderNode In Root.childNodes
		set a=Root.removeChild(HeaderNode)
	next
end Function

Function ClearTable()
	dim i
	for i=1 to document.all.tblDetails.rows.length - 1
		document.all.tblDetails.deleteRow(1)
	next
	j = 0
end Function
'**************************************************************
Function InvoiceDetails_onClick(sGPNo,sOrgID)
	document.formname.action ="InvComeFromGatePass.asp?iGPNo="& sGPNo & "&ForUnit=" & sOrgID
	document.formname.submit
End Function
'**************************************************************

Function Check(sPassGP)
	Dim OutValue,ndInv
	Dim sRateType
	sForUnit = document.formname.hUnit.value
	set OutValue = showModalDialog("GatePassServiceAccountPop.asp?iGPNo="& sPassGP & "&ForUnit=" & sForUnit,"","dialogHeight:350px;dialogWidth:600px;help:no;status:no")
	if OutValue.hasChildNodes() then
		for each ndInv in OutValue.childNodes
			if ndInv.nodeName="WithInv" then
				sRateType  = ndInv.getAttribute("Rate")
			end if
		next
		document.formname.action ="InvComeFromGatePass.asp?iGPNo="& sPassGP & "&ForUnit=" & sForUnit&"&RateType="&sRateType
	else
		document.formname.action = "GatePassServiceAccountEntry.asp"
	end if
		document.formname.submit()
End Function

'------------------------------------------------------------------------------------------
function AssignPage(nPage)
	document.formname.hPage.value = nPage
	document.formname.action = "GatePassServiceAccountEntry.asp"
	document.formname.submit()
end function
'------------------------------------------------------------------------------------------

Function Validate()

	dim sOrgID,sSupplier,sItemType
	sOrgID = document.formname.cmbUnit.value


	'sItemType = document.formname.cmbItemType.value
	if sOrgID = "0" then
		MsgBox "Select Unit"
		document.formname.cmbUnit.focus
		exit function
'	elseif sItemType = "0" then
'		MsgBox "Select Itemtype"
'		document.formname.cmbItemType.focus
'		exit function
	end if

	document.formname.hFromDate.value = document.formname.ctlRcptFromDate.getDate
	document.formname.hToDate.value = document.formname.ctlRcptToDate.getDate
	document.formname.hItemTypeName.value =""' document.formname.cmbItemType(document.formname.cmbItemType.selectedIndex).text
	document.formname.action =  "GatePassServiceAccountEntry.asp"
	document.formname.submit

End Function
''------------------------------------------------------------------------------------------
Function fninit()
	dim i
	for i=0 to (document.formname.cmbUnit.length-1)
		if document.formname.cmbUnit.options(i).value=document.formname.hUnit.value then
			document.formname.cmbUnit.selectedIndex=i
		end if
	next

'	for i=0 to (document.formname.cmbItemType.length-1)
'		if document.formname.cmbItemType.options(i).value=document.formname.hItemTypeName.value then
''			document.formname.cmbItemType.selectedIndex=i
'		end if
'	next

	if document.formname.hMonthDate.value="D" then
		document.formname.rMonthDate.checked=true
		document.formname.ctlRcptFromDate.SetDate=document.formname.hFromDate.value
		document.formname.ctlRcptToDate.SetDate=document.formname.hToDate.value
	end if

	document.formname.txtRefName.value =document.formname.txtSupplier.value

	document.formname.ctlRcptFromDate.disabled = TRUE
	document.formname.ctlRcptToDate.disabled = TRUE
End Function
'------------------------------------------------------------------------------------------
Function DisplayExistingCriteria()
	sImageSrc = ucase(trim(document.formname.ImgSearch.src))

	if instr(1,sImageSrc,"MINUS.GIF") > 0 then


		'document.formname.mTxtSupplierName.value  = document.formname.hSupplierName.value

	end if 'if instr(1,sImageSrc,"MINUS.GIF") > 0 then

End Function
'------------------------------------------------------------------------------------------

Function gotoPage()
	document.formname.action = "GatePassServiceAccountEntry.asp"
	document.formname.submit()
End Function
'------------------------------------------------------------------------------------------
Function ResetForm()
	if document.formname.cmbUnit.selectedIndex = "0" then exit function
	'document.formname.cmbItemType.value = 0
	document.formname.action =  "GatePassServiceAccountEntry.asp"
	document.formname.submit
End Function

'------------------------------------------------------------------------------------------


</Script>
<%
Dim sForUnit, sPartyCode, sPartyName,sSqlTemp,sDateTo,sStatus
Dim sItemType, sMonthDate, sMonthDateVal, sDateFrom,sPassSupplierCode
Dim sInvoiceData,sAppRefNo,sAppRefType
Dim nRecCtr

Dim iPageNo,iTotalPages,iPrevPage,iNextPage,nPageSize
Dim iTotalRecords, iStartRec,iEndRec,nPageCtr,DtRcvd,DtRcvdP

Dim rsTemp,rsInvoice,rsObj
Dim sFinPeriod,sArrFin,sFinFrom,sFinTo


Set rsTemp = server.CreateObject("Adodb.recordset")
Set rsObj = server.CreateObject("Adodb.recordset")
Set rsInvoice = Server.CreateObject("ADODB.Recordset")




nPageSize = 20
iPageNo = trim(Request("hPage"))
if trim(iPageNo) = "" then iPageNo = 1
'Response.Write "<p> iPageNo = " & iPageNo




If sForUnit = "" Then sForUnit = Trim(request("cmbUnit"))
'If sItemType = "" Then sItemType = Trim(request("cmbItemType"))
If sPassSupplierCode = "" Then sPassSupplierCode = Trim(request("hSupplierCode"))
If sPartyName = "" Then sPartyName = Trim(request("txtSupplier"))

If sMonthDate = "" Then sMonthDate = Trim(request("rMonthDate"))

'sStatus = trim(Request("rdStatus"))
'if sStatus = "" then sStatus = "T"
'Response.Write "<p> sItemType = " & sItemType

If sMonthDate = "D" Then
	If sDateFrom = "" Then sDateFrom = Trim(request("hFromDate"))
	If sDateTo = "" Then sDateTo = Trim(request("hToDate"))
End If

sFinPeriod = Session("FinPeriod")
sArrFin = Split(sFinPeriod,":")
sFinFrom = "01/04/"&sArrFin(0)
sFinTo = "31/03/"&sArrFin(1)



if trim(sPassSupplierCode) <> "" then
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = "Select PartyCode from Map_Supplier where OUDefinitionID='" & sForUnit & "' and SupplierCode = " & sPassSupplierCode & ""
		.ActiveConnection = con
		.Open
	end with
	set rsTemp.ActiveConnection = nothing
	if not rsTemp.EOF then
		sPartyCode  = rsTemp(0)
	else
		sPartyCode = 0
	end if
	rsTemp.close
end if





if trim(sForUnit) = "" then


	If iSAApplicationPop1 <> "" then ' this is session variable definded in Purpopulate.asp
		sSqlTemp = " SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate1 & " AND APPLICATIONCODE = " & iSAApplicationPop1 & " AND PROCESSCODE = " & iSAProcessPop1 & " AND ACTIVITYCODE = " & iSAApplicationPop1 & ") ORDER BY OUDEFINITIONID"
	Else
		sSqlTemp = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
	End If
	'response.write sSqlTemp


	Set rsTemp = Server.CreateObject("ADODB.RecordSet")
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSqlTemp
		.ActiveConnection = con
		.Open
	end with
	set rsTemp.ActiveConnection = nothing

	If not rsTemp.EOF then
		sForUnit = rsTemp(0)
	end if
	rsTemp.Close


'	If iSAApplicationPop1 <> "" then ' this is session variable definded in Purpopulate.asp
'		sSqlTemp = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID IN (SELECT DISTINCT ITEMTYPEID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate1 & " AND APPLICATIONCODE = " & iSAApplicationPop1 & " AND PROCESSCODE = " & iSAProcessPop1 & " AND ACTIVITYCODE = " & iSAActivityPop1 & ") ORDER BY ITEMTYPENO"
'	Else
'		sSqlTemp = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
'	End If
'	'response.write sSqlTemp
'
'
'	Set rsTemp = Server.CreateObject("ADODB.RecordSet")
'	with rsTemp
'		.CursorLocation = 3
'		.CursorType = 3
'		.Source = sSqlTemp
'		.ActiveConnection = con
'		.Open
'	end with
'	set rsTemp.ActiveConnection = nothing
'
'	If not rsTemp.EOF then
'		sItemType = rsTemp(0)
'	end if
'	rsTemp.Close
'

end if 'if trim(sForUnit) = "" then

%>

<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fninit()">
<form method="POST" name="formname">

<input type="hidden" name="hPage" value="">

<Input type=hidden name="hUnit" value="<%=sForUnit%>">
<Input type=hidden name="hSupplierCode" value="<%=sPassSupplierCode%>">
<Input type=hidden name="txtSupplier" value="<%=sPartyName%>">

<Input type=hidden name="hItemTypeName" value="<%=sItemType%>">
<Input type=hidden name="hFromDate" value="<%=sDateFrom%>">
<Input type=hidden name="hToDate" value="<%=sDateTo%>">


<Input type=hidden name="hMonthDate" value="<%=sMonthDate%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Gate Pass - Service</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine >
						<table border="0" cellpadding="0" cellspacing="0" width="100%" >
                            <tr>
								<td align="center" colspan="4" class="MiddlePack" height="7" width="2"></td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../../assets/images/clearpixel.gif" width="2" height="5">
								</td>
								<td valign="top" colspan="2" align="left">

									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable" align=left >
										<tr>
											<td>
										<div align="left">

										<table class="CollapseBand" cellspacing="0" cellpadding="0">
											<tr>
												<td valign="center">
													<a style="width: 1em; height: 1em;" title="" onclick="Div_OnClick(idUnprocessed,'');DisplayExistingCriteria()" >
													<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
													</a>
												</td>
												<td valign="center" class="SubTitle">&nbsp;&nbsp;

												</td>
											</tr>

										</table>


										<table border="0" width="100%" cellpadding="0" cellspacing="0"  class="BodyTable">
										<tr>
											<td width="100%">
												<div align="left" id="idUnprocessed" style="width: 575; display: none">


													<table border="0" cellpadding="0" cellspacing="0">
													<tr>
														<td class="FieldCellSub">For Unit</td>
														<td class="FieldCellSub" colspan="4"><select size="1" name="cmbUnit" class="FormElem">
																<option value="0">Select</option>
														<%
															populateUnit(sForUnit)
														%>
															</select>
														</td>
													</tr>
													<!--<tr>
														<td class="FieldCellSub">Item type</td>
														<td class="FieldCellSub"><select size="1" name="cmbItemType" class="FormElem">
															<option value="0">Select</option>
														<%
															''Populate Item type based on the UserId
														'	popSelItemType(sItemType)
														%>
															</select>
														</td>
													</tr>-->
													<tr>
														<td class="FieldCell"><input type="radio" value="D" name="rMonthDate">Date Range From</td>
														<td class="FieldCellSub">
														<% ' Function Call to Insert Date Picker
															Response.Write InsertDatePicker("ctlRcptFromDate")
														%>
														</td>
														<td class="FieldCellSub"></td>
														<td class="FieldCellSub">To</td>
														<td class="FieldCellSub">
														<% ' Function Call to Insert Date Picker
															Response.Write InsertDatePicker("ctlRcptToDate")
														%>
														</td>
													</tr>
													<tr>
														<td class="FieldCellSub">Party</td>
														<td class="FieldCellSub" colspan="4">
															<input type="text" name="txtRefName" value="" size="60" class="formelemread" readonly>&nbsp;
															<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Party" width="11" height="11" onClick="popSuppAgent()"></a>
														</td>
													</tr>
													<!--<tr>
														<td class="FieldCellSub">Status</td>
														<td class="FieldCellSub" colspan="4">
															<input type="radio" value="R" name="rdStatus" <%if sStatus = "R" then Response.Write " Checked " %> >Returned
															<input type="radio" value="P" name="rdStatus" <%if sStatus = "P" then Response.Write " Checked " %> >Partial Return
															<input type="radio" value="T" name="rdStatus" <%if sStatus = "T" then Response.Write " Checked " %> >To Return
														</td>
													</tr>-->

														<tr>
															<td class="FieldCell"></td>
															<td class="FieldCell" >
																<input type="button" value="Go" name="ButGo" class="ActionButton" onclick="Validate()">
																<input type="button" value="Reset" name="ButReset" class="ActionButton" onclick="ResetData()" >
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
								<td align="center" class="ClearPixel" width="6" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="4" class="MiddlePack" height="7" width="2"></td>
                            </tr>

                            <tr>
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" colspan="2">
                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                      <tr>
                                        <td valign="top">
                                          <div class="frmBody" id="frm3" >
                                            <table id="tblDetails" border="0" cellspacing="1" class="ExcelTable" width="100%">
                                              <tr>
													<td class="ExcelHeaderCell" align="center" width="10">
														<p align="center">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center">
														Party Name
													</td>
													<td class="ExcelHeaderCell" align="center">
														DC No. & Date
													</td>
													<td class="ExcelHeaderCell" align="center">Status</td>
													<%
														'if sStatus ="R" then
															Response.Write "<td class=ExcelHeaderCell align=center>Reference Type (No-Date)</td>"
														'end if 'if sStatus ="R" then
													%>
												</tr>

												<%
												Dim nCount1,nCountR,nCountP,nCountT

												'sSqlTemp = "Select Distinct G.GatePassNo,isNull(P.PartyName,''),isNull(G.DCCode,''),Convert(varchar,G.MarkedOn,103),isNull(G.Status,'N'),G.MarkedOn From ForGatePassHeader G, App_M_PartyMaster P Where G.PartyCode = P.PartyCode And G.OrganisationCode = '" & sForUnit & "' And G.TypeOfItems =  '" & sItemType & "' and GatePassType='E'"

												sSqlTemp = "Select Distinct G.GatePassNo,isNull(P.PartyName,''),isNull(G.DCCode,''),Convert(varchar,G.MarkedOn,103),isNull(G.Status,'N'),G.MarkedOn,IsNull(G.AppRefNo,''),IsNull(G.AppRefType,'') From ForGatePassHeader G, App_M_PartyMaster P Where G.PartyCode = P.PartyCode And G.OrganisationCode = '" & sForUnit & "' And InvoiceType='V'"

												If trim(sPartyCode) <> "" then
													sSqlTemp = sSqlTemp & " and G.PartyCode = " & sPartyCode & " "
												End If

												sSqlTemp = sSqlTemp & " and Convert(datetime,G.MarkedOn,103)	>= Convert(datetime,'" & sFinFrom & "',103) and Convert(datetime,G.MarkedOn,103) <= Convert(datetime,'" & sFinTo & "',103)"

												If sMonthDate = "D" Then
													sSqlTemp = sSqlTemp & " and Convert(datetime,G.MarkedOn,103)	>= Convert(datetime,'" & sDateFrom & "',103) and Convert(datetime,G.MarkedOn,103) <= Convert(datetime,'" & sDateTo & "',103)"
												End If

												'if sStatus = "R" then
												'	sSqlTemp = sSqlTemp & " and (select count(Det.GatePassNo) from ForGatePassDetails Det where G.GatePassNo = Det.GatePassNo) = " & _
												'			" (select count(Det.GatePassNo) from ForGatePassDetails Det where G.GatePassNo = Det.GatePassNo and isNull(Det.MaterialRcvd,'N')= 'Y')"
												'elseif sStatus = "P" then
												'	sSqlTemp = sSqlTemp & " and (select count(Det.GatePassNo) from ForGatePassDetails Det where G.GatePassNo = Det.GatePassNo) >= " & _
												'			" (select count(Det.GatePassNo) from ForGatePassDetails Det where G.GatePassNo = Det.GatePassNo and isNull(Det.MaterialRcvd,'N')= 'Y')"
												'elseif sStatus = "T" then
												'	sSqlTemp = sSqlTemp & " and (select count(Det.GatePassNo) from ForGatePassDetails Det where G.GatePassNo = Det.GatePassNo) = " & _
												'			" (select count(Det.GatePassNo) from ForGatePassDetails Det where G.GatePassNo = Det.GatePassNo and isNull(Det.MaterialRcvd,'N')= 'N')"
												'end if

												sSqlTemp = sSqlTemp & " Order By G.MarkedOn "
												'Response.Write "<textarea>" & sSqlTemp&"</textarea>"
												with rsTemp
													.CursorLocation = 3
													.CursorType = 3
													.Source = sSqlTemp
													.PageSize = nPageSize
													.ActiveConnection = con
													.Open
												end with

												set rsTemp.ActiveConnection = nothing

												nRecCtr = 1

												If not rsTemp.EOF Then
													iTotalPages = rsTemp.PageCount
													iTotalRecords = rsTemp.RecordCount
													rsTemp.AbsolutePage = iPageNo
												Else
													iTotalPages = 0
													iTotalRecords = 0

													iStartRec = 0
													iEndRec = 0
												End If

												if trim(iPageNo) = 1 then
													iPrevPage = 0
												else
													iPrevPage = iPageNo - 1
												end if


												if iTotalPages >= iPageNo + 1 then
													iNextPage = iPageNo + 1
												else
													iNextPage = 0
												end if

												if not rsTemp.EOF then
													do while not rsTemp.EOF and nRecCtr <= rsTemp.PageSize

														%>
														<tr>
															<td class="ExcelSerial" align="center" width="31"><%=nRecCtr%> </td>
															<td class="ExcelDisplayCell"><%=rsTemp(1)%></td>
															<td class="ExcelDisplayCell">
																<a href="#" onclick="Check(<%=rsTemp(0)%>)" class="ExcelDisplayLink">
																	<%=rsTemp(2) & " - " & rsTemp(3) %>
																</a>
															</td>
															<td class="ExcelDisplayCell"><%

															sAppRefNo = rsTemp(6)
															sAppRefType = rsTemp(7)

																sSqlTemp = "SELECT MAX(CONVERT(datetime, MaterialRcvdOn, 103)) FROM ForGatePassDetails WHERE (ISNULL(MaterialRcvd, 'N') = 'Y') AND (GatePassNo = "& rsTemp(0) &")"
																rsObj.open sSqlTemp ,con
																if not rsObj.eof then
																	DtRcvd = rsObj(0)
																end if
																rsObj.Close



																nCount1 = 0
																nCountR = 0
																nCountP = 0
																nCountT = 0

																sSqlTemp = " select count(Det.GatePassNo) from ForGatePassDetails Det where Det.GatePassNo = "& rsTemp(0)&"  "
																rsObj.open sSqlTemp ,con
																if not rsObj.eof then
																	nCount1 = rsObj(0)
																End IF
																rsObj.Close

																sSqlTemp = " select count(Det.GatePassNo) from ForGatePassDetails Det where Det.GatePassNo = "& rsTemp(0)&" and isNull(Det.MaterialRcvd,'N')= 'Y'"
																rsObj.open sSqlTemp ,con
																if not rsObj.eof then
																	nCountR = rsObj(0)
																End IF
																rsObj.Close

																sSqlTemp = " select count(Det.GatePassNo) from ForGatePassDetails Det where Det.GatePassNo = "& rsTemp(0)&" and isNull(Det.MaterialRcvd,'N')= 'Y'"
																rsObj.open sSqlTemp ,con
																if not rsObj.eof then
																	nCountP = rsObj(0)
																End IF
																rsObj.Close

																sSqlTemp = " select count(Det.GatePassNo) from ForGatePassDetails Det where Det.GatePassNo = "& rsTemp(0)&" and isNull(Det.MaterialRcvd,'N')= 'N'"
																rsObj.open sSqlTemp ,con
																if not rsObj.eof then
																	nCountT = rsObj(0)
																End IF
																rsObj.Close

																If nCount1 = nCountR Then
																	sStatus = "R"
																Elseif nCount1 > nCountP Then
																	sStatus = "P"
																Elseif nCount1 = nCountT Then
																	sStatus = "T"
																End IF

															if trim(sStatus)="T" then
																Response.Write "Not Received"
															else
																if trim(DtRcvd)<>"" then
																	Response.Write "Returned On "& DtRcvd
																else
																	Response.Write "Not Received"
																end if
															end if
														'	if ucase(trim(rsTemp(4))) = "Y" then
														'		Response.Write "Returned"
														'	else
														'		Response.Write "To Return"
														'	end if
															%></td>
															<%
																sInvoiceData = ""

																if trim(sAppRefNo)<>"" and trim(sAppRefType)<>"" then
																    sInvoiceData = GetRefNoDate(sAppRefType,sAppRefNo)
																    sInvoiceData = Replace(sInvoiceData,",","(") &")"

																end if

															'	if sStatus ="R" then
																	Response.Write "<td class=ExcelDisplayCell>"
																	Response.Write "<span style='width:150px'>"&sInvoiceData&"</span>"
																	'Response.Write "<img src='../../assets/images/iTMS%20icons/Entry.gif' onClick=InvoiceDetails_onClick('"& trim(rsTemp(0)) &"','"& sForUnit &"') alt='Click here to enter invoice details'>"
																	Response.Write "</td>"
															'	Else
															'		Response.Write "<td class=ExcelDisplayCell>"
															'		Response.Write "<span style='width:150px'>&nbsp;</span>"
															'		Response.Write "</td>"
															'	end if ' if sStatus ="R" then
															%>
														</tr>
														<%
														nRecCtr = nRecCtr + 1
														rsTemp.MoveNext
													loop
												end if
												rsTemp.Close

												%>

                                            </table>
                                          </div>
                                        </td>
                                      </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="6" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>


							<%
								if iTotalPages > 0 then
									if iPageNo  = 1 then
										iStartRec = 1
									else
										iStartRec  = (iPageNo - 1 ) * nPageSize + 1
									end if
									iEndRec = iStartRec + nPageSize -1
									if iTotalRecords < nPageSize  then
										iEndRec = iTotalRecords
									end if

									if iEndRec > iTotalRecords then
										iEndRec = iTotalRecords
									end if
								end if 	'if iTotalPages > 0 then

							%>

								<!--
								<tr></tr><tr></tr>
								<tr>
									<td align="center"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"></td>
									<td align="center">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="button" value="New Search" class="ActionButton" onClick="gotoPage()">
												</td>
											</tr>
										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								-->


								 <tr>
									<td align="center" class="BottomPack" colspan="4" width="761">
									</td>
                                </tr>

                                <tr>
									<td align="right" colspan="2" width="100%" >
										<table border="0" cellspacing="0" cellpadding="0" >
											<td valign="top" align="center">

											</td>
											<td valign="top" align="right">

											<input type="button" value=" |< " class="ActionButtonX" id=ButFirst name=ButFirst onClick="AssignPage('1')">

											<%if trim(iPrevPage) = "0" then  %>
												<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev >
											<%else%>
												<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev onClick="AssignPage('<%=iPrevPage%>')">
											<%end if %>


											<SELECT class="FormElem" onChange="AssignPage(this.value)"  id="mCmbPage" name="mCmbPage">

											<%for nPageCtr= 1 to iTotalPages %>
												<option value="<%=nPageCtr%>" <%if trim(iPageNo) = trim(nPageCtr) then Response.Write "Selected" %> >Page <%=nPageCtr%> of <%=iTotalPages %></option>
											<%next%>

											</SELECT>
											<%if trim(iNextPage) = "0" then  %>
												<input type="button" value=" >> " class="ActionButtonX" id=ButNext name=ButNext >
											<%else%>
												<input type="button" value=" >> " class="ActionButtonX" onclick="AssignPage('<%=iNextPage%>')" id=ButNext name=ButNext >
											<%end if%>

											<input type="button" value=" >| " class="ActionButtonX" id=ButLast name=ButLast OnClick="AssignPage('<%=iTotalPages %>')">

											</td>
										</table>
									</td>

									<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<tr>
																	<td align="center" class="MiddlePack" colspan="3">
																	</td>
								</tr>
								<tr>
									<td align="center" class="ClearPixel" width="5">
																		<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td align="center" width="100%" class="ActionCell">
										<table border="0" cellspacing="0" cellpadding="0" width="100%">
											<td valign="top" align="center">
												<input type="button" value="Misc.Service Bill" name="btnServiceBill" onClick="btnServiceBill_Click()" class="ActionButtonX">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											</td>
										</table>
									</td>

									<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<tr>
									<td align="center" class="BottomPack" colspan="3">
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

