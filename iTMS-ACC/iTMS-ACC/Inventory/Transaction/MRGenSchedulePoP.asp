<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MRGenSchedulePoP.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 28, 2005
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
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
dim oDom,Root,PageNode,HeaderNode,PGNode

dim dcrs,sql,sUnitName,iItem,iClass
dim sSchType,iQty,iCounter,EntryNode,iEntNo,sOptName
dim arrTemp,iMRSNo,sOrgID,sOrgName,sItemName
dim arrUoM,sUoMDesc,sUoMCode
dim sFinFrom,sFinTo,sTodaysDate

arrTemp =  split(session("FinPeriod"),":")
sFinFrom = "01/04/"&arrTemp(0)
sFinTo = "31/03/"&arrTemp(1)
sTodaysDate  = FormatDate(Date)

set dcrs = server.CreateObject("Adodb.recordset")
set oDom = server.CreateObject("Microsoft.xmlDom")

sSchType="D"
arrTemp = split(trim(Request.QueryString("sTemp")),":")
iQty	= arrTemp(0)
iItem	= arrTemp(1)
iClass	= arrTemp(2)
iEntNo 	= arrTemp(3)
sOrgID  = arrTemp(4)
sUoMCode = arrTemp(5)
sOptName = arrTemp(6)
sUoMDesc = sUoMCode
'sItemName = ItemDisplay(iItem,iClass)
'Response.Write iEntNo
sql = "Select ItemDescription from VwItem where ItemCode = "& iItem &" and ClassificationCode = "& iClass &" and OrganisationCode = '"& sOrgID &"' "
dcrs.Open sql,Con
If not dcrs.EOF then
	sItemName = dcrs(0)
End If
dcrs.Close
if sOptName <> "" then sItemName = sItemName & sOptName
'
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : MR - Schedule Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><ROOT></ROOT></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/MRCreateSch.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=iItem%>','<%=iClass%>','<%=iQty%>','<%=iEntNo%>')">

<form method="POST" name="formname">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">MR - Schedule
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
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
									<table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCellSub">Description</td>
                                            <td class="FieldCellSub"><span class="DataOnly"><%=sItemName%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Quantity</td>
                                            <td class="FieldCellSub"><span class="DataOnly"><%=iQty%>&nbsp;</span> - <span class="DataOnly"><%=sUoMDesc%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Schedule Type</td>

                                            <td class="FieldCellSub">
												<SELECT NAME="selSchtype" class="FormElem" size="1" onchange="setMax(this)">
													<option value="select">Select</option>
													<option value="D">Date</option>
													<option value="M">MonthYear</option>
													<option value="W">WeekYear</option>
													<option value="Y">MonthWeekYear</option>
												</SELECT>
											</td>
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
								<DIV class=frmBody id=frm3 style="width: 220; height:275;">
                                    <table border="0" cellspacing="1" class="ExcelTable">
                                        <tr>
											<td class="ExcelHeaderCell" align="center">S.No.</td>
											<td class="ExcelHeaderCell" align="center">Need by</td>
											<td class="ExcelHeaderCell" align="center">Quantity</td>
                                        </tr>
										<%for iCounter=1 to 12%>
                                        <tr>
											<td class="ExcelSerial" align="center"><%=iCounter%></td>
											<td class="ExcelInputCell"><input type="text" name="txtD<%=iCounter%>" size="12" maxlength=10 class="FormElem"></td>
											<td class="ExcelInputCell"><input type="text" name="txtQ<%=iCounter%>" size="12" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" class="FormElem" style="text-align=right"></td>
                                        </tr>
                                        <%next%>
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
                                                    <input type="button" value="Done" name="B4" class="ActionButton" onClick="CheckSubmit('<%=sTodaysDate%>')">
                                                    <input type="reset" value="Reset" name="B5" class="ActionButton">
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

