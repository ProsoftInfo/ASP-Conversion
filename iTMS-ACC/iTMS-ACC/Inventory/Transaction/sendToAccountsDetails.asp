<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	sendToAccountsDetails.asp
	'Module Name				:	Inventory (Send Closing Stock to Accounts)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	September 24,2003
	'Modified By				:	TAJUDEEN S
	'Modified On				:	June 8, 2004
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	sendToAccountsInsert.asp
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
<!--#include file="../../include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Closing Stock Details to Accounts</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<%
	'XML DOM Variables
	Dim oDOM
	dim RootNode,HeaderNode,PageNode,EntryNode

	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	' Declaration of variables
	Dim dcrs,dcrs1
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

	dim sUnitName,dClosingDate,sFor,sUnit,sForName,iCtr
	dim iAccHead,sAccHeadDesc,iAccHeadPL,sAccHeadPLDesc,iTotRecValue,iTotIssValue,iTotValue
	dim sMonYr,arrFin,sFinFrom,sTempMonYr,iTotPLValue,iOpeningValue,iClosingValue
	dim dCrTotal,dDrTotal,iNdCtr,sExp,sTemp,bFlag,dLastSentDate,iNo
	dim iTotalOpeningValue,iTotalRecValue,iTotalIssValue,iTotalClosingValue
    Dim sFinPeriod
	iTotPLValue = 0
	bFlag = false

	sFinPeriod = Session("FinPeriod")
	arrFin = split(sFinPeriod,":")
	sFinFrom = "01/04/"&arrFin(0)

	iCtr = 0
	sUnit = trim(Request.Form("selUnit"))
	sUnitName = trim(Request.Form("hUnitName"))
	dClosingDate = trim(Request.Form("hClosingDate"))
	sFor = trim(Request.Form("selFor"))
	sForName = trim(Request.Form("hForName"))

	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT MAX(CONVERT(DATETIME,TRANSACTIONDATE,103)) FROM INV_T_ITEMLEDGER WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND SENTTOACCOUNTS = 'S'"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing
	if not dcrs1.EOF then
		dLastSentDate = FormatDate(dcrs1(0))
	else
		dLastSentDate = "N/A"
	end if
	dcrs1.Close
	
	if dLastSentDate = "//" then dLastSentDate = "N/A"

	set RootNode = oDOM.createElement("voucher")
	RootNode.setAttribute "UnitNo",sUnit
	RootNode.setAttribute "UnitName",sUnitName
	RootNode.setAttribute "BookNo",""
	RootNode.setAttribute "BookName",""
	RootNode.setAttribute "CRDR",""
	RootNode.setAttribute "VouDate",dClosingDate
	RootNode.setAttribute "Approver","Y"
	RootNode.setAttribute "TransNo",""
	RootNode.setAttribute "VoucherNo",""

	oDOM.appendChild(RootNode)

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ACCOUNTHEAD,ACCOUNTHEADCODE,ACCOUNTDESCRIPTION FROM ACC_M_GLACCOUNTHEAD WHERE ACCOUNTHEAD = (SELECT ACCOUNTHEAD FROM INV_CONTROL_ORGACCOUNTHEADS WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND ACCOUNTHEADFOR = 'L') ORDER BY ACCOUNTHEAD"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iAccHeadPL = trim(dcrs(0))
		sAccHeadPLDesc = trim(dcrs(2))
		bFlag = true
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ACCOUNTHEAD,ACCOUNTHEADCODE,ACCOUNTDESCRIPTION FROM ACC_M_GLACCOUNTHEAD WHERE ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM INV_CONTROL_ORGACCOUNTHEADS WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND ACCOUNTHEADFOR = 'C') ORDER BY ACCOUNTHEAD"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		do while not dcrs.EOF
			iTotIssValue = 0
			iTotRecValue = 0
			iCtr = iCtr + 1
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(TRANSACTVALUE,0),TRANSACTIONTYPE FROM INV_T_ITEMLEDGER WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(DATETIME,TRANSACTIONDATE,103) >= CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME,TRANSACTIONDATE,103) <= CONVERT(DATETIME," & Pack(dClosingDate) & ",103) AND (STR(ITEMCODE)+STR(CLASSIFICATIONCODE)) IN (SELECT (STR(ITEMCODE)+STR(CLASSIFICATIONCODE)) FROM INV_M_ITEMORGACCOUNTHEAD WHERE ACCOUNTHEAD = " & trim(dcrs(0)) & " AND ORGANISATIONCODE = " & Pack(sUnit) & " AND ACCOUNTHEADFOR = 'C') AND SENTTOACCOUNTS = 'T' AND TRANSACTIONTYPE <> 'RO'"
				.ActiveConnection = con
			'	response.write "<textarea>"& dcrs1.source&"</textarea>"
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			do while not dcrs1.EOF
				if left(trim(dcrs1(1)),1) = "I" then
					iTotIssValue = cdbl(iTotIssValue) + cdbl(dcrs1(0))
				elseif left(trim(dcrs1(1)),1) = "R" then
					iTotRecValue = cdbl(iTotRecValue) + cdbl(dcrs1(0))
				end if
			dcrs1.MoveNext
			loop
			dcrs1.Close

			'Fetching Opening Value
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(SUM(YEAROPENINGVALUE),0) FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME,FINANCIALYEARFROM,103) = CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND (STR(ITEMCODE)+STR(CLASSIFICATIONCODE)) IN (SELECT (STR(ITEMCODE)+STR(CLASSIFICATIONCODE)) FROM INV_M_ITEMORGACCOUNTHEAD WHERE ACCOUNTHEAD = " & trim(dcrs(0)) & " AND ORGANISATIONCODE = " & Pack(sUnit) & " AND ACCOUNTHEADFOR = 'C') AND ORGANISATIONCODE = " & Pack(sUnit)
				.ActiveConnection = con
				'response.write dcrs1.source
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then 
				iOpeningValue = cdbl(dcrs1(0))
			else
				iOpeningValue = 0
			end if
			dcrs1.Close 
			dim sCRDR

			iClosingValue = cdbl(iOpeningValue) + cdbl(iTotRecValue) - cdbl(iTotIssValue)
			iTotPLValue = cdbl(iTotPLValue) + (cdbl(iTotRecValue) - cdbl(iTotIssValue))

			if (cdbl(iTotRecValue) - cdbl(iTotIssValue)) < 0 then
				sCRDR = "C"
			else
				sCRDR = "D"
			end if

			set HeaderNode = oDOM.createElement("Entry")
			HeaderNode.setAttribute "No",iCtr
			HeaderNode.setAttribute "CRDR",sCRDR
			HeaderNode.setAttribute "Payto",""
			HeaderNode.setAttribute "Amount",FormatNumber(abs(cdbl(iTotRecValue) - cdbl(iTotIssValue)),2,,,0)
			HeaderNode.setAttribute "AccUnit",sUnit
			HeaderNode.setAttribute "AccName",sUnitName
			'Newly added Parameters
			HeaderNode.setAttribute "OpeningValue",iOpeningValue
			HeaderNode.setAttribute "ReceiptValue",iTotRecValue
			HeaderNode.setAttribute "IssueValue",iTotIssValue
			HeaderNode.setAttribute "ClosingValue",iClosingValue

			RootNode.appendChild(HeaderNode)

			set PageNode = oDOM.createElement("AccHead")
			PageNode.setAttribute "No",trim(dcrs(0))
			PageNode.setAttribute "Name",trim(dcrs(2))
			PageNode.setAttribute "CostCenter","0"
			PageNode.setAttribute "Analytical","0"
			PageNode.setAttribute "Type","G"
			PageNode.setAttribute "TransFlag","W"

			HeaderNode.appendChild(PageNode)

			set PageNode = oDOM.createElement("Narration")
			PageNode.Text = "Closing Stock As on " & dClosingDate & " for " & trim(dcrs(2))
			HeaderNode.appendChild(PageNode)

		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	iCtr = iCtr + 1
	if iTotPLValue < 0 then
		sCRDR = "D"
	else
		sCRDR = "C"
	end if

	set HeaderNode = oDOM.createElement("Entry")
	HeaderNode.setAttribute "No",iCtr
	HeaderNode.setAttribute "CRDR",sCRDR
	HeaderNode.setAttribute "Payto",""
	HeaderNode.setAttribute "Amount",FormatNumber(abs(iTotPLValue),2,,,0)
	HeaderNode.setAttribute "AccUnit",sUnit
	HeaderNode.setAttribute "AccName",sUnitName

	RootNode.appendChild(HeaderNode)

	set PageNode = oDOM.createElement("AccHead")
	PageNode.setAttribute "No",iAccHeadPL
	PageNode.setAttribute "Name",sAccHeadPLDesc
	PageNode.setAttribute "CostCenter","0"
	PageNode.setAttribute "Analytical","0"
	PageNode.setAttribute "Type","G"
	PageNode.setAttribute "TransFlag","W"

	HeaderNode.appendChild(PageNode)

	set PageNode = oDOM.createElement("Narration")
	PageNode.Text = "Closing Stock As on " & dClosingDate & " for " & sAccHeadPLDesc
	HeaderNode.appendChild(PageNode)


	oDOM.save server.MapPath("../temp/transaction/Creation_GJ_"&Session.SessionID&".xml")
	iCtr = 0

%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/Creation_GJ_"&Session.SessionID&".xml"%>"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
FUNCTION CheckSubmit()
	document.formname.action = "sendToAccountsInsert.asp"
	document.formname.submit()
END FUNCTION
'*******************************************************
	Function CheckSetup()
	    showModalDialog "SetupInvBooks.asp","","dialogWidth:500px;dialogHeight:250px;Status:No;"
	End Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%	if not bFlag then %>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	alert("Account Head for Profit and Loss is not defined for the Selected Organization")
	window.location.href = "sendToAccounts.asp"
</SCRIPT>
<%	end if
	'if cdbl(iTotRecValue) = 0 and cdbl(iTotIssValue) = 0 then
	if cdbl(iTotPLValue) = 0 then
%>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	alert("Closing Stock has already been updated to Accounts for the Selected Date")
	'window.location.href = "sendToAccounts.asp"
</SCRIPT>
<%	end if %>

<form method="POST" name="formname" action="">
<input type=hidden name="hFinFrom" value="<%=sFinFrom%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Closing Stock Details to Accounts
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
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onmouseover="tabrollover(this)" onmouseout="tabrollout(this)">
										<tr>
											<td width="100%" align="center">Header</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Details</td>
										</tr>
									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center" width="5"></td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Unit Name</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=sUnitName%>&nbsp;</span>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Closing For</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=sForName%>&nbsp;</span>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Last Transferred On</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=dLastSentDate%>&nbsp;</span>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Closing as on</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=dClosingDate%>&nbsp;</span>
											</td>
										</tr>
									</table>
								</td>
								<td align="center" width="5"></td>
							</tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>

							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td valign="top">
												<div class="frmBody" id="frm3" style="width: 585; height: 315">
													<table border="0" cellspacing="1" class="ExcelTable" width="100%">
														<tr>
															<td class="ExcelHeaderCell" align="center" rowspan="2" width="30"><p align="center">S.No.</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2" >Account Head</td>
															<td class="ExcelHeaderCell" align="center" colspan="4" width="300">Value</td>
														</tr>
														<tr>
															<td class="ExcelHeaderCell" align="center" width="75">Opening</td>
															<td class="ExcelHeaderCell" align="center" width="75">Receipt</td>
															<td class="ExcelHeaderCell" align="center" width="75">Issue</td>
															<td class="ExcelHeaderCell" align="center" width="75">Closing</td>
														</tr>
													<%
														dCrTotal=0
														dDrTotal=0
														iTotalOpeningValue = 0
														iTotalRecValue = 0
														iTotalIssValue = 0
														iTotalClosingValue = 0
														
														sExp ="//voucher/Entry"
														Set EntryNode = RootNode.Selectnodes(sExp)

														for iNdCtr =  0 to EntryNode.length - 2
																sTemp = EntryNode.Item(iNdCtr).Attributes.getNamedItem("CRDR").Value
																iCtr = iCtr + 1
																iTotValue = EntryNode.Item(iNdCtr).Attributes.getNamedItem("Amount").Value

																iTotValue = FormatNumber(iTotValue,2,,,0)

																iOpeningValue = EntryNode.Item(iNdCtr).Attributes.getNamedItem("OpeningValue").Value
																iTotRecValue = EntryNode.Item(iNdCtr).Attributes.getNamedItem("ReceiptValue").Value
																iTotIssValue = EntryNode.Item(iNdCtr).Attributes.getNamedItem("IssueValue").Value
																iClosingValue = EntryNode.Item(iNdCtr).Attributes.getNamedItem("ClosingValue").Value
																
																iTotalOpeningValue = cdbl(iTotalOpeningValue) + cdbl(iOpeningValue)
																iTotalRecValue = cdbl(iTotalRecValue) + cdbl(iTotRecValue)
																iTotalIssValue = cdbl(iTotalIssValue) + cdbl(iTotIssValue)
																iTotalClosingValue = cdbl(iTotalClosingValue) + cdbl(iClosingValue)

																'Response.Write iOpeningValue
																for each PageNode in EntryNode.Item(iNdCtr).childNodes
																	if PageNode.nodeName="AccHead" then
																		sAccHeadDesc = trim(PageNode.Attributes.getNamedItem("Name").value)
																		iNo = trim(PageNode.Attributes.getNamedItem("No").value)
																	end if
																	exit for
																next
													%>

														<tr>
															<td class="ExcelSerial" align="center"><%=iCtr%></td>
															<td class="ExcelDisplayCell">
															<a href="javascript:void(0)" class="ExcelDisplayLink" onClick="showModalDialog('SendToAccountsPopup.asp?AccountHead=<%=iNo%>&Unit=<%=sUnit%>&Date=<%=dClosingDate%>','','dialogHeight:370px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No')"><%=sAccHeadDesc%></a>
															</td>
													<%
																if sTemp = "D" then
																	dDrTotal = CDbl(dDrTotal) + CDbl(iTotValue)
													%>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iOpeningValue),2,-1,-1,0)%></td>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iTotRecValue),2,-1,-1,0)%></td>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iTotIssValue),2,-1,-1,0) %></td>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iClosingValue),2,-1,-1,0)%></td>
													<%
																else
																	dCrTotal = CDbl(dCrTotal) + CDbl(iTotValue)
													%>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iOpeningValue),2,-1,-1,0)%></td>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iTotRecValue),2,-1,-1,0)%></td>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iTotIssValue),2,-1,-1,0) %></td>
															<td class="ExcelDisplayCell" align="right"><%=FormatNumber(cdbl(iClosingValue),2,-1,-1,0)%></td>
													<%			end if	%>
														</tr>
													<%
														next
													%>
														<tr>
															<td class="ExcelSerial" align="right" colspan="2"><B>Total</B></td>
															<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(cdbl(iTotalOpeningValue),2,-1,-1,0)%></B></td>
															<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(cdbl(iTotalRecValue),2,-1,-1,0)%></B></td>
															<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(cdbl(iTotalIssValue),2,-1,-1,0) %></B></td>
															<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(cdbl(iTotalClosingValue),2,-1,-1,0)%></B></td>
														</tr>
														<!--tr>
																<td class="ExcelSerial" align="center"></td>
																<td class="ExcelDisplayCell" align="right">Total</td>
																<td class="ExcelDisplayCell" align="right"><%=dDrTotal%></td>
																<td class="ExcelDisplayCell" align="right"><%=dCrTotal%></td>
														</tr-->
													</table>
												</div>
											</td>
										</tr>
									</table>
								</td>
								<td align="center" width="5"></td>
							</tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>

							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
												<input type="button" value="Setup" name="btnSetup" class="ActionButton" onclick="CheckSetup()" />
												<input type="button" value="Account" name="B2" class="ActionButton" onclick="CheckSubmit()">
 												<input type="button" value="Back" name="B1" class="ActionButton" onclick="window.history.back(1)">
 												<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('sendToAccounts.asp')">
											</td>
										</tr>

									</table>
								</td>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

							<tr>
								<td align="center" width="10" colspan="3" class="BottomPack">
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
