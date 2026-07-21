<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsIssueSchedulePoP.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 18, 2003
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
dim dcrs,iItem,iClass
dim i,sSchType,iQty,bFlag,iCounter
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,sItemName
dim iTotSchQty,iTotMarQty,iTotIssQty,iEntNo,sOptName

set dcrs = server.CreateObject("Adodb.recordset")

arrTemp = split(trim(Request.QueryString("sTemp")),":")
iQty	= arrTemp(0)
iClass	= arrTemp(2)
iItem	= arrTemp(1)
iMRSNo = arrTemp(3)
sOrgID = arrTemp(4)
iEntNo = arrTemp(5)
sOptName = arrTemp(6)
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT ORGUNITSHORTDESCRIPTION,CONVERT(CHAR,MRSDATE,103) FROM VWMRSLIST WHERE MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sOrgName = trim(dcrs(0))
	dMRSDate = trim(dcrs(1))
end if
dcrs.Close

sItemName = ItemDisplay(iItem,iClass)

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT SCHEDULETYPE FROM INV_T_MRSITEMSCHEDULES WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sSchType = trim(dcrs(0))
end if
dcrs.Close

select case sSchType
	case "D"
		sSchType = "Date"
	case "M"
		sSchType = "MonthYear"
	case "Y"
		sSchType = "MonthWeekYear"
	case "W"
		sSchType = "WeekYear"

end select
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : MR Issue - Schedule Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Material Requisition - Schedule Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%"  >
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell" width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
                                            <td class="FieldCell">MR No.&nbsp;- Date</td>
                                            <td class="FieldCellSub"><span class="DataOnly"><%=iMRSNo%>&nbsp;</span> - <span class="DataOnly"><%=dMRSDate%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Unit Name</td>
                                            <td class="FieldCellSub"><span class="DataOnly"><%=sOrgName%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell"> Description</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="idItemName"><%=sItemName%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Schedule Type</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="idSch"><%=sSchType%>&nbsp;</span></td>
                                        </tr>
							       </table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                        <tr>
							<td align="center" class="ClearPixel"></td>
							<td valign="top" class="FieldCell" width="100%" align="center">
								<DIV class=frmBody id=frm3 style="width: 350; height:275;">
                                    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                        <tr>
											<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
											<td class="ExcelHeaderCell" align="center" rowspan="2">Need by</td>
											<td class="ExcelHeaderCell" align="center" colspan="3">Quantity</td>
										</tr>
										<tr>
											<td class="ExcelHeaderCell" align="center">Scheduled</td>
											<td class="ExcelHeaderCell" align="center">Marked</td>
											<td class="ExcelHeaderCell" align="center">Issued</td>
                                        </tr>
										<%
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT SCHEDULEDON,SCHEDULEDQTY,ISNULL(MARKEDQTY,0),ISNULL(ISSUEDQTY,0) FROM INV_T_MRSITEMSCHEDULES WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " ORDER BY 1"
												.ActiveConnection = con
												.Open
											end with
											set dcrs.ActiveConnection = nothing

											if not dcrs.EOF then
												do while not dcrs.EOF
													iCounter = iCounter + 1
													iTotSchQty = cdbl(iTotSchQty) + cdbl(trim(dcrs(1)))
													iTotMarQty = cdbl(iTotMarQty) + cdbl(trim(dcrs(2)))
													iTotIssQty = cdbl(iTotIssQty) + cdbl(trim(dcrs(3)))
										%>
                                        <tr>
											<td class="ExcelSerial" align="center"><%=iCounter%></td>
											<td class="ExcelDisplayCell" width="10"><input type="text" name="txtD<%=iCounter%>" value="<%=trim(dcrs(0))%>" size="12" class="FormElemRead" READONLY></td>
											<td class="ExcelDisplayCell" width="10"><input type="text" name="txtQ<%=iCounter%>" value="<%=trim(dcrs(1))%>" size="12" class="FormElemRead" READONLY style="text-align:right"></td>
											<td class="ExcelDisplayCell" width="10"><input type="text" name="txtM<%=iCounter%>" value="<%=trim(dcrs(2))%>" size="12" class="FormElemRead" READONLY style="text-align:right"></td>
											<td class="ExcelDisplayCell" width="10"><input type="text" name="txtI<%=iCounter%>" value="<%=trim(dcrs(3))%>" size="12" class="FormElemRead" READONLY style="text-align:right"></td>
                                        </tr>
                                        <%
												dcrs.MoveNext
												loop
											end if
											dcrs.Close
                                        %>
                                        <tr>
											<td class="ExcelSerial" colspan=2 align="right">Total &nbsp;</td>
											<td class="ExcelDisplayCell" align="right"><B><%=trim(iTotSchQty)%></B></td>
											<td class="ExcelDisplayCell" align="right"><B><%=trim(iTotMarQty)%></B></td>
											<td class="ExcelDisplayCell" align="right"><B><%=trim(iTotIssQty)%></B></td>
                                        </tr>
                                    </table>
								</div>
							</td>
							<td align="center" class="ClearPixel"></td>
                        </tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                    <input type="button" value="OK" name="B4" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
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
