<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	itmStockDetailsPop.asp
	'Module Name				:	Inventory (Item Stock Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 14, 2003
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
<!-- #include File="../../include/ItemDisplay.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Stock Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%
	dim dcrs,dcrs1,sSql,iCtr,iQty
	dim iItem,iClass,sOrg,sItemName,sClassName,sOrgName,sUoM,sBinName
	dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sBin

	iItem = trim(Request.QueryString("sItem"))
	iClass = trim(Request.QueryString("sClass"))
	sOrg = trim(Request.QueryString("sOrg"))
	sItemName = trim(Request.QueryString("ItemName"))

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT SHORTDESCRIPTION,GROUPNAME,ORGUNITSHORTDESCRIPTION,ITEMDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrg) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if Not dcrs.EOF then
		'sItemName = trim(dcrs(0))
		sClassName = trim(dcrs(1))
		sOrgName = trim(dcrs(2))
		'sItemName = trim(dcrs(3))
	end if
	dcrs.Close

	'sItemName = ItemDisplay(iItem,iClass)
'	sItemName =  sItemName & iOptName
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrg) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if Not dcrs.EOF then
		sUoM = trim(dcrs(0))
	end if
	dcrs.Close

	if len(Month(date())) = 1 then
		sTempMonYr = "0"&Month(date())
	else
		sTempMonYr = Month(date())
	end if
	sMonYr = sTempMonYr&Year(date())

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)

%>
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr><td height="1px"></td></tr>

	<tr>
		<td class="PageTitle">
			Stock Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >

				<TR>
					<TD class=TabBodyWithTopLine>
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
									<table border="0" cellspacing="0" cellpadding="0">
                                        <tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub" colspan=4>
                                                <span class="DataOnly"><%=sItemName%>&nbsp;</span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Organization</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">UoM</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly"><%=sUoM%>&nbsp;</span>
                                            </td>
											<td class="FieldCellSub"></td>
                                        </tr>
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5"></td>
								<td valign="top" width="100%">
									<div class="frmBody" id="frm2" style="width: 100%; height:250;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Store</td>
												<td class="ExcelHeaderCell" align="center" colspan="4">Quantity Status</td>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center" width="80">Clean</td>
												<td class="ExcelHeaderCell" align="center" width="80">Reserve</td>
												<td class="ExcelHeaderCell" align="center" width="80">On Hold</td>
												<td class="ExcelHeaderCell" align="center" width="80">Rejected</td>
											</tr>
										<%
												with dcrs
													.CursorLocation = 3
													.CursorType = 3
													.Source = "SELECT LOCATIONNAME,ISNULL(BINNUMBER,0),ISNULL(RESERVED,0),ISNULL(ONHOLD,0),ISNULL(REJECTED,0),LOCATIONNUMBER FROM VWITEMSTOCKSTATUS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrg) & " AND APPLICABLEFOR = 'IN' AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
													.ActiveConnection = con
													.Open
												end with
												set dcrs.ActiveConnection = nothing
												'Response.Write dcrs.Source
												if Not dcrs.EOF then
													do while not dcrs.EOF
														sBin = trim(dcrs(1))

														if sBin = "0" then sBin = "NULL"

														with dcrs1
															.CursorLocation = 3
															.CursorType = 3
															'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrg) & " AND LOCATIONNUMBER = " & trim(dcrs(5)) & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
															.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrg) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
															.ActiveConnection = con
															.Open
														end with
														set dcrs1.ActiveConnection = nothing

														if not dcrs1.EOF then
															iQty = cdbl(dcrs1(0))
														end if
														dcrs1.close

														if sBin <> "NULL" then
															with dcrs1
																.CursorLocation = 3
																.CursorType = 3
																.Source = "SELECT BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & trim(dcrs(5)) & " AND BINNUMBER = " & sBin & ""
																.ActiveConnection = con
																.Open
															end with
															set dcrs1.ActiveConnection = nothing
															if not dcrs1.EOF then
																sBinName = " - " & trim(dcrs1(0))
															end if
															dcrs1.Close
														else
															sBinName = ""
														end if

														if iQty >= 0 then
															iCtr = iCtr + 1
										%>
											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell"><%=trim(dcrs(0))%> <%=sBinName%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelDisplayCell" align="right"><%=trim(dcrs(2))%></td>
												<td class="ExcelDisplayCell" align="right"><%=trim(dcrs(3))%></td>
												<td class="ExcelDisplayCell" align="right"><%=trim(dcrs(4))%></td>
											</tr>
										<%
														end if
													dcrs.MoveNext
													loop
												end if
												dcrs.Close
										%>
										</table>
									</div>
								</td>
								<td align="center"></td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                    <input type="button" value="Close" name="B3" class="ActionButton" onClick="window.close()">
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
</form>
</BODY>
</HTML>