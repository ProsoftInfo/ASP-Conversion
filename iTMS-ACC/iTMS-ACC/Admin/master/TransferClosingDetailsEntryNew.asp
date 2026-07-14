<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	TransferClosingDetailsEntryNew.asp
	'Module Name				:	Transfer Closing Values
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 13, 2004
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Transfer Closing Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AdminTransferClosingCompat.js"></SCRIPT>
<SCRIPT>
ITMSAdminTransferClosingCompat.installTransferDetails();
</SCRIPT>
</HEAD>
<%
	dim sUnit,sUnitName,dPreFinStartDate,dPreFinEndDate,dCurFinStartDate,dCurFinEndDate
	Dim iCurPeriodFrom,iCurPeriodTo,sQuery,dcrs,sNSClosed
	dim arrTemp,sWho,sFor, sItemTypeId
	sUnit = trim(Request.Form("hUnitCode"))
	sUnitName = trim(Request.Form("hOrgName"))
	sItemTypeId = trim(Request.Form("hItemType"))
	dCurFinStartDate = trim(Request.Form("hCFinStartDate"))
	dCurFinEndDate = trim(Request.Form("hCFinEndDate"))
	sWho = trim(Request.Form("hWho"))
	IF Request.Form("selPFinStartDate") <> "" then
		arrTemp = split(trim(Request.Form("selPFinStartDate")),":")
		dPreFinStartDate = arrTemp(0)
		dPreFinEndDate = arrTemp(1)
	End IF
	sWho = "Y"
	sFor = Request("Frm")
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
%>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type=hidden name="hUnit" value="<%=sUnit%>">
<input type=hidden name="hOrgName" value="<%=sUnitName%>">
<input type=hidden name="hCFinStartDate" value="<%=dCurFinStartDate%>">
<input type=hidden name="hCFinEndDate" value="<%=dCurFinEndDate%>">
<input type=hidden name="hPFinStartDate" value="<%=dPreFinStartDate%>">
<input type=hidden name="hPFinEndDate" value="<%=dPreFinEndDate%>">
<input type=hidden name="hApplication" value="">
<input type=hidden name="hItemTypeId" value="<%=sItemTypeId%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td align="center" class=PageTitle height="20">
			<p align="center">Transfer Closing Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >

				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%" align="left">
                                    <table BORDER="0" CELLSPACING="1" CELLPADDING="0">
										<tr>
											<td class="FieldCell" valign="top">For Unit</td>
											<td class="FieldCellSub" valign="top">
												<span class="Dataonly"><%=sUnitName%></span></td>
                                            </tr>
                                    </table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class='GroupTitleLeft' width="10">&nbsp;</td>
														<td class='GroupTitle' width="147"><p align="center">Transfer Closing Values</td>
														<td class='GroupTitleRight'><p align="left">&nbsp;</td>
													</tr>
												</table>
                                            </td>
										</tr>
										<tr>
											<td class=GroupTable>
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td class=MiddlePack colspan="7"> </td>
													</tr>
													<tr>
														<td class=FieldCellSub valign="top"> Previous Financial Year</td>
														<td class="FieldCellSub" valign="top">Start Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly"><%=dPreFinStartDate%></span></td>
														<td class='FieldCellSub'></td>
														<td class='FieldCellSub'>End Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly"><%=dPreFinEndDate%></span>
														</td>
														<td class='FieldCellSub'></td>
													</tr>
													<tr>
														<td class=MiddlePack valign="top" colspan="7">
														</td>
													</tr>
													<tr>
														<td class=FieldCellSub valign="top">Current Financial Year</td>
														<td class="FieldCellSub" valign="top">Start Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly"><%=dCurFinStartDate%></span>
														</td>
														<td class='FieldCellSub'></td>
														<td class='FieldCellSub'>End Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly"><%=dCurFinEndDate%></span>
                                                        </td>
														<td class='FieldCellSub'></td>
													</tr>
												</table>
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
								<td valign="top">
									<table border="0" cellspacing="1" class="ExcelTable">
										<tr>
											<td class="ExcelHeaderCell" align="center" width="10" valign="middle">S.No.</td>
											<td class="ExcelHeaderCell" align="center" colspan="2" valign="middle">Activity</td>
											<td class="ExcelHeaderCell" align="center" valign="middle">Transferred<br>/ Initialised On</td>
											<td class="ExcelHeaderCell" align="center" valign="middle">Transfer</td>
										</tr>
										<tr>
											<td class="ExcelHeaderCell" align="center" valign="middle" colspan="5">ACCOUNTS - CARRY FORWARD</td>
										</tr>
									<%
										sNSClosed = CheckNoSeriesTransfer(1)
										if CStr(sFor) = "NS" then
									%>

										<tr>
											<td class="ExcelSerial" align="center" valign="middle">1</td>
											<td class="ExcelDisplayCell" valign="middle">Number Series</td>
											<td class="ExcelDisplayCell" valign="top"></td>
											<td class="ExcelDisplayCell" valign="middle"><%If sNSClosed > 1 Then Response.Write "No Series transferred"%></td>
											<td class="ExcelDisplayCell" valign="middle" align="center">&nbsp;
 												<input type="button" value="Transfer" name="btnAcc" onClick="Transfer('AC','1','1')" class="AddButtonX" <%If sNSClosed > 1 Then Response.Write "Disabled"%>>&nbsp;
											</td>
										</tr>
										<tr>
											<td class="ExcelHeaderCell" align="center" valign="middle" colspan="5">STORES - CARRY FORWARD (ALL ITEM TYPES)</td>
										</tr>
									<%sNSClosed = CheckNoSeriesTransfer(4)%>
										<tr>
											<td class="ExcelSerial" align="center" valign="middle">1</td>
											<td class="ExcelDisplayCell" valign="middle">Number Series</td>
											<td class="ExcelDisplayCell" valign="top"></td>
											<td class="ExcelDisplayCell" valign="middle"><%If sNSClosed > 1 Then Response.Write "No Series transferred"%></td>
											<td class="ExcelDisplayCell" valign="middle">&nbsp;
 												<input type="button" value="Transfer" name="btnSto" class="AddButtonX" onClick="Transfer('IN','1','4')" <%If sNSClosed > 1 Then Response.Write "Disabled"%>>&nbsp;
											</td>
										</tr>

										<!--tr>
											<td class="ExcelSerial" align="center" valign="middle">2</td>
											<td class="ExcelDisplayCell" valign="middle">Closing Stock Details</td>
											<td class="ExcelDisplayCell" valign="top"></td>
											<td class="ExcelDisplayCell" valign="middle"></td>
											<td class="ExcelDisplayCell" valign="middle">&nbsp;
 												<input type="button" value="Transfer" name="B8" class="AddButtonX" onClick="Transfer('IN','2','4')">&nbsp;
											</td>
										</tr-->
										<%sNSClosed = CheckNoSeriesTransfer(2)%>
										<tr>
											<td class="ExcelHeaderCell" align="center" valign="middle" colspan="5">PURCHASE - CARRY FORWARD (ALL ITEM TYPES)</td>
										</tr>

										<tr>
											<td class="ExcelSerial" align="center" valign="middle">1</td>
											<td class="ExcelDisplayCell" valign="middle">Number Series</td>
											<td class="ExcelDisplayCell" valign="top"></td>
											<td class="ExcelDisplayCell" valign="middle"><%If sNSClosed > 1 Then Response.Write "No Series transferred"%></td>
											<td class="ExcelDisplayCell" valign="middle">&nbsp;
 												<input type="button" value="Transfer" onClick="Transfer('PU','1','2')" name="btnPur" class="AddButtonX" <%If sNSClosed > 1 Then Response.Write "Disabled"%>>&nbsp;
											</td>
										</tr>

										<tr>
											<td class="ExcelHeaderCell" align="center" valign="middle" colspan="5">SALES - CARRY FORWARD (ALL ITEM TYPES)</td>
										</tr>
										<%sNSClosed = CheckNoSeriesTransfer(3)%>
										<tr>
											<td class="ExcelSerial" align="center" valign="middle">1</td>
											<td class="ExcelDisplayCell" valign="middle">Number Series</td>
											<td class="ExcelDisplayCell" valign="top"></td>
											<td class="ExcelDisplayCell" valign="middle"><%If sNSClosed > 1 Then Response.Write "No Series transferred"%></td>
											<td class="ExcelDisplayCell" valign="middle">&nbsp;
 												<input type="button" value="Transfer" name="btnSal" onClick="Transfer('SA','1','3')" class="AddButtonX" <%If sNSClosed > 1 Then Response.Write "Disabled"%>>&nbsp;
											</td>
										</tr>



										<tr>
											<td class="ExcelHeaderCell" align="center" valign="middle" colspan="5">PRODUCTION - CARRY FORWARD (ALL ITEM TYPES)</td>
										</tr>
										<%sNSClosed = CheckNoSeriesTransfer(6)%>
										<tr>
											<td class="ExcelSerial" align="center" valign="middle">1</td>
											<td class="ExcelDisplayCell" valign="middle">Number Series</td>
											<td class="ExcelDisplayCell" valign="top"></td>
											<td class="ExcelDisplayCell" valign="middle"><%If sNSClosed > 1 Then Response.Write "No Series transferred"%></td>
											<td class="ExcelDisplayCell" valign="middle">&nbsp;
 												<input type="button" value="Transfer" name="btnPro" onClick="Transfer('PD','1','6')" class="AddButtonX" <%If sNSClosed > 1 Then Response.Write "Disabled"%>>&nbsp;
											</td>
										</tr>

									<%Elseif CStr(sFor) = "IS" Then %>
											<tr>
												<td class="ExcelSerial" align="center" valign="middle">2</td>
												<td class="ExcelDisplayCell" valign="middle">Closing Stock Details</td>
												<td class="ExcelDisplayCell" valign="top"></td>
												<td class="ExcelDisplayCell" valign="middle"></td>
												<td class="ExcelDisplayCell" valign="middle">&nbsp;
 													<input type="button" value="Transfer" name="btnStock" class="AddButtonX" onClick="Transfer('IN','2','4')">&nbsp;
												</td>
											</tr>

									<%	end if %>
									</table>
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
												<input type="button" value="Cancel" name="cancel" class="ActionButton" OnClick="window.location.href='CloseEntry.asp?Frm=<%=sFor%>'">
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
Function CheckNoSeriesTransfer(iApplCode)
	iCurPeriodFrom = right(dCurFinStartDate,4)&mid(dCurFinStartDate,4,2)
	iCurPeriodTo = right(dCurFinEndDate,4)&mid(dCurFinEndDate,4,2)
	sQuery = "SELECT Count(1) FROM APP_R_NOSERIESMODULEENTRY WHERE (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) IN (SELECT (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) FROM APP_R_NOSERIESMODULES WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND APPLICATIONCODE = "&iApplCode&" AND (PERIOD >= " & iCurPeriodFrom & " AND PERIOD <= " & iCurPeriodTo & "))"
	'Response.Write sQuery
	'Response.End
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		CheckNoSeriesTransfer = dcrs(0)
	Else
		CheckNoSeriesTransfer = 0
	end if
	dcrs.Close
End Function
%>
