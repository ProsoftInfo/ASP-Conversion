<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SendToAccountsPopup.asp
	'Module Name				:	Inventory (Send Closing Stock to Accounts)
	'Author Name				:	TAJUDEEN S
	'Created On					:	June 07, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	SendToAccountsDetails.asp
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
<!-- #include File="../../include/ItemDisplay.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%
	dim dcrs,dcrs1,i,sSql,iNo, sOrgID,oDOM,objFs, Root, Node, sExp, sAccName
	dim iItemCode, iClassCode, sItemName, iIssueValue, iReceiptValue, iClosingValue
	dim iOpeningValue, sTempMonYr,sMonYr,arrFin,sFinFrom,dClosingDate
	dim iTotOpeningValue, iTotIssueValue, iTotReceiptValue, iTotClosingValue
	Dim sFinPeriod

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objFs = server.CreateObject("Scripting.FileSystemObject")

    sFinPeriod = Session("FinPeriod")
	iNo = Request.QueryString("AccountHead")
	sOrgID = Request.QueryString("Unit")
	dClosingDate = Request.QueryString("Date")
	
	if objFs.FileExists(server.MapPath("../temp/transaction/Creation_GJ_"&Session.SessionID&".xml")) then
		oDOM.load server.MapPath("../temp/transaction/Creation_GJ_"&Session.SessionID&".xml")
		
		set Root = oDOM.documentElement
		sExp = "//voucher//Entry//AccHead [@No = " & iNo & "]"
		set Node = Root.selectNodes(sExp)
		
		if Node.length > 0  then
			sAccName = Node.Item(0).attributes.getNamedItem("Name").value
		end if
		
		set oDOM = Nothing
	end if


%>

<form method="POST" name="formname" Action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Details
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="TopPack">
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td >
                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                        <tr><td class="FieldCell"></td></tr>
 										<tr>
											<td class="FieldCell" align="center">Account Head : <Span Class="DataOnly"><%=sAccName%></Span></td>
										</tr>

                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" >
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td >
									<div class="frmBody" id="frm2" style="width: 100%; height:240px;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="10" >S.No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="200">Item Description</td>
												<td class="ExcelHeaderCell" align="center" colspan="4" width="300">Value</td>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center" width="75">Opening</td>
												<td class="ExcelHeaderCell" align="center" width="75">Receipt</td>
												<td class="ExcelHeaderCell" align="center" width="75">Issue</td>
												<td class="ExcelHeaderCell" align="center" width="75">Closing</td>
											</tr>
									<%

                                        
										arrFin = split(sFinPeriod,":")
										sFinFrom = "01/04/"&arrFin(0)

										iTotOpeningValue = 0
										iTotIssueValue = 0
										iTotReceiptValue = 0
										iTotClosingValue = 0
										
										SSQL="SELECT ITEMCODE, CLASSIFICATIONCODE FROM INV_M_ITEMORGACCOUNTHEAD WHERE ACCOUNTHEAD = " & iNo & " AND ORGANISATIONCODE = " &Pack(sOrgID) & " AND ACCOUNTHEADFOR = 'C'"
										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = sSql
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing
										do while not dcrs.eof
											i = i + 1
											iItemCode = trim(dcrs(0))
											iClassCode = trim(dcrs(1))
											sItemName = ItemDisplay(iItemCode,iClassCode)
'-------------------------------------------------------------------------------------------
'Stock Details											
											iIssueValue = 0
											iReceiptValue = 0
											'Fetching Receipt and Issue Value
											with dcrs1
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT ISNULL(TRANSACTVALUE,0),TRANSACTIONTYPE FROM INV_T_ITEMLEDGER WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME,TRANSACTIONDATE,103) >= CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME,TRANSACTIONDATE,103) <= CONVERT(DATETIME," & Pack(dClosingDate) & ",103) AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND TRANSACTIONTYPE <> 'RO'" 
												.ActiveConnection = con
												.Open
											end with
											set dcrs1.ActiveConnection = nothing
											do while not dcrs1.EOF
												if left(trim(dcrs1(1)),1) = "I" then
													iIssueValue = cdbl(iIssueValue) + cdbl(dcrs1(0))
												elseif left(trim(dcrs1(1)),1) = "R" then
													iReceiptValue = cdbl(iReceiptValue) + cdbl(dcrs1(0))
												end if
											dcrs1.MoveNext
											loop
											dcrs1.Close

											'Fetching Opening Value
											with dcrs1
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT ISNULL(YEAROPENINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME,FINANCIALYEARFROM,103) = CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID)
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

											iClosingValue = cdbl(iOpeningValue) + cdbl(iReceiptValue) - cdbl(iIssueValue)
											
											iTotOpeningValue = cdbl(iTotOpeningValue) + cdbl(iOpeningValue)
											iTotReceiptValue = cdbl(iTotReceiptValue) + cdbl(iReceiptValue) 
											iTotIssueValue = cdbl(iTotIssueValue ) + cdbl(iIssueValue)  
											iTotClosingValue = cdbl(iTotClosingValue) + cdbl(iClosingValue) 
'-------------------------------------------------------------------------------------------

									%>
											<tr>
												<td class="ExcelSerial" align="center"><%=i%></td>
												<td class="ExcelDisplayCell"><%=sItemName%></td>
												<td class="ExcelDisplayCell" align="right"><%=FormatNumber(iOpeningValue,2,-1,-1,0) %></td>
												<td class="ExcelDisplayCell" align="right"><%=FormatNumber(iReceiptValue,2,-1,-1,0)%></td>
												<td class="ExcelDisplayCell" align="right"><%=FormatNumber(iIssueValue,2,-1,-1,0)%></td>
												<td class="ExcelDisplayCell" align="right"><%=FormatNumber(iClosingValue,2,-1,-1,0)%></td>
											</tr>
									<% 
											dcrs.movenext
										loop
										dcrs.close
										
										if i = 0 then
									%>
											<tr>
												<td class="ExcelDisplayCell" Colspan="6" align="center">No Items Defined</td>
											</tr>
									<%	else	%>
											<tr>
												<td class="ExcelSerial" align="right" colspan="2"><B>Total</B></td>
												<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(iTotOpeningValue,2,-1,-1,0) %></B></td>
												<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(iTotReceiptValue,2,-1,-1,0)%></B></td>
												<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(iTotIssueValue,2,-1,-1,0)%></B></td>
												<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(iTotClosingValue,2,-1,-1,0)%></B></td>
											</tr>									
									<%	end if	%>
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
											<td valign="middle" class="ActionCell" align="center">
                                                    <input type="button" value="Close" name="Close" class="ActionButton" onclick="window.close()">
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
	Function GetPackingName(iCode)

		if icode = "" or isnull(iCode) then
			GetPackingName ="N/A"
			exit function
		end if

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PACKINGNAME FROM APP_M_PACKINGTYPE WHERE PACKINGCODE =" & iCode
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		if not dcrs.EOF then
			GetPackingName = trim(dcrs1(0))
		else
			GetPackingName = "N/A"
		end if
		dcrs1.Close
	End Function

	Function GetSellingName(iCode)
		if icode = "" or isnull(iCode) then
			GetSellingName ="N/A"
			exit function
		end if

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS WHERE CODE = " & iCode
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		if not dcrs1.EOF then
			GetSellingName = dcrs1(0)
		else
			GetSellingName = "N/A"
		end if
		dcrs1.Close

	End Function


%>