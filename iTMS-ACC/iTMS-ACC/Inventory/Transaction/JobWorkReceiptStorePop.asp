<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	JobWorkReceiptStorePop.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	TAJUDEEN S
	'Created On					:	September 16, 2004
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
<!-- #include file="../../include/NoSeries.asp"-->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
Response.Write "<font color=#000000#>"
dim dcrs,iItem,iClass,sOrgID,iMRSNo,arrTemp,iIssNo,sRecName,sLotNumber
dim sItmName,sClassName,iQty,sRecNumbering,sRec,iCtr,sStoreName
dim sItmType,iSeriesNo,iSeriesCode,bNoSerFlag,bPackNoSerFlag
dim sStoUoM,sPurUoM,sOperator,iConvRate,iAccQty
dim arrUoM,sUoMDesc,sUoMCode,sUoMCheck
iCtr = 1
bNoSerFlag = false
bPackNoSerFlag = false
set dcrs = Server.CreateObject("ADODB.Recordset")

arrTemp = split(trim(Request.QueryString("sTemp")),":")

iItem	= arrTemp(0)
iClass	= arrTemp(1)
sOrgID	= arrTemp(2)
iQty = arrTemp(3)
sRec = arrTemp(5)

if trim(sRec) = "N" then
	sRecNumbering = "DISABLED"
else
	sRecNumbering = ""
end if

arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoMDesc = arrUoM(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Receipts</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="XmlData"><Root/></script>
<%
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT GROUPNAME,SHORTDESCRIPTION FROM VWITEM WHERE CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItem & " AND ORGANISATIONCODE = " & Pack(sorgID) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sClassName = trim(dcrs(0))
	sItmName = trim(dcrs(1))
end if
dcrs.Close

sItmName = ItemDisplay(iItem,iClass)

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ITEMTYPEID,STORESUOM,ISNULL(PURCHASEUOM,'N/A'),PURTOSTOREOPERATOR,ISNULL(PURTOSTORERATE,0) FROM VWITEM WHERE ClassificationCode = " & iClass & " AND ITEMCODE = " & iItem & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sItmType = trim(dcrs(0))
	sStoUoM = trim(dcrs(1))
	sPurUoM = trim(dcrs(2))
	sOperator = trim(dcrs(3))
	iConvRate = trim(dcrs(4))

	sUoMCheck = UoMDecimal(sPurUoM)

	if lcase(sStoUoM) = lcase(sPurUoM) then
		iAccQty = iQty
	else
		if sOperator = "0" then iAccQty = cdbl(iQty) / cdbl(iConvRate)
		if sOperator = "1" then iAccQty = cdbl(iQty) * cdbl(iConvRate)
		if sUoMCheck = "Y" then
			iAccQty = FormatNumber(iAccQty,3)
		else
			iAccQty = Round(iAccQty)
		end if
	end if
end if
dcrs.Close

' Packing Number - Number Series
' OrganisationCode,SeriesNumber,SeriesCode,Date
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	'.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'PN' AND ITEMTYPE = " & Pack(sItmType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
	.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'PN' AND ITEMTYPE = 'YRN' AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
if dcrs.EOF then
	bPackNoSerFlag = true
end if
dcrs.close

if sRec = "N" then
	sRecName = "None"
elseif sRec = "L" then
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'LO' AND ITEMTYPE = " & Pack(sItmType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		iSeriesNo = trim(dcrs(0))
		iSeriesCode = trim(dcrs(1))

		' Lot Number - Number Series
		' OrganisationCode,SeriesNumber,SeriesCode,Date
		'sLotNumber = GetSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,FormatDate(date())) & "[" & replace(FormatDate(date()),"/","") & "]"
	else
		bNoSerFlag = true
	end if
	dcrs.close

	sRecName = "Lot"
elseif sRec = "S" then
	sRecName = "Serial"
elseif sRec = "LS" then
	' Lot Number - Number Series
	' OrganisationCode,SeriesNumber,SeriesCode,Date
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'LO' AND ITEMTYPE = " & Pack(sItmType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		iSeriesNo = trim(dcrs(0))
		iSeriesCode = trim(dcrs(1))

		'sLotNumber = GetSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,FormatDate(date())) & "[" & replace(FormatDate(date()),"/","") & "]"
	else
		bNoSerFlag = true
	end if
	dcrs.close

	sRecName = "Lot and Serial"
end if
%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ReceiptStoreBase.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/JobWorkReceiptStore.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">
<%	if bNoSerFlag and not bPackNoSerFlag then %>
<SCRIPT LANGUAGE=javascript>
	ITMSReceiptStore.closeWithDialogArguments("Number Series for Lot Number has not been defined for this Item Type.");
</SCRIPT>
<%	elseif not bNoSerFlag and bPackNoSerFlag then %>
<SCRIPT LANGUAGE=javascript>
	ITMSReceiptStore.closeWithDialogArguments("Number Series for Packing Number has not been defined for this Item Type.");
</SCRIPT>
<%	elseif bNoSerFlag and bPackNoSerFlag then %>
<SCRIPT LANGUAGE=javascript>
	ITMSReceiptStore.closeWithDialogArguments("Number Series for Lot and Packing Number has not been defined for this Item Type.");
</SCRIPT>
<%	end if %>
<form method="POST" name="formname" action="">
<input type="hidden" name="hRec" value="<%=sRec%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Storage Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idClassName"><%=sItmName%></span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity Received</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly"><%=iQty%>&nbsp;</span>
                                                <span class="DataOnly"><%=sPurUoM%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity to Account</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idQty"><%=iAccQty%>&nbsp;</span>
                                                <span class="DataOnly"><%=sStoUoM%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Receipt Numbering</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly"><%=sRecName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell"></td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idLotNumber"><%=sLotNumber%></span>
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
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<DIV class=frmBody id=frm6 style="width: 350; height:150;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Store -- Bin</td>
												<td class="ExcelHeaderCell" align="center">Lot / Serial</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
											</tr>
											<%
													with dcrs
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 ORDER BY 1,2"
														.ActiveConnection = con
														.Open

													end with
													set dcrs.ActiveConnection = nothing

													if not dcrs.EOF then
														Do While Not dcrs.EOF
															'
															'Response.Write "A"
															sStoreName = DisplayStore(trim(dcrs(0)),trim(dcrs(1)))
											%>
											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell" align="left"><%=sStoreName%></td>
												<td class="ExcelDisplayCell" align="center">
													<input type="button" value=" Yes " name="btn:<%=iClass%>:<%=iItem%>:<%=sOrgID%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(1))%>" <%=sRecNumbering%> class="AddButtonX" onClick="GetLot(this,'<%=sRec%>','<%=iCtr%>','<%=sStoreName%>')">
												</td>
											<%	if not sRecNumbering = "DISABLED" then %>
												<td class="ExcelAverageCell" width="10">
													<input type="text" name="txtQty<%=iCtr%>" size="10" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElemRead" READONLY style="text-align:right">
											<%	else %>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtQty<%=iCtr%>" size="10" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align:right">
											<%	end if %>
												</td>
											</tr>
											<%			iCtr = iCtr + 1
														dcrs.MoveNext
														loop
													else
											%>
											<tr>
												<td class="ExcelDisplayCell" align="center" colspan=5><b>No Inventory Storage related for this Item. Relate first and account.</b></td>
											</tr>
											<%
													end if
													dcrs.Close
											%>
										</table>
									</div>
								</td>
								<input type="hidden" name="hiCtr" value="<%=iCtr-1%>">
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
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
	' Function to get Store
	Function DisplayStore(sLoc,sBin)
		' Declaration of variables
		Dim dcrs,sBinName,sLocName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT LOCATIONNAME,LOCATIONCODE FROM INV_M_STORAGE WHERE LOCATIONNUMBER = " & sLoc & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			sLocName = trim(dcrs(0))
		else
			sLocName = "-"
		end if
		dcrs.close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT BINNAME,BINCODE FROM INV_M_StoreBINDETAILS WHERE LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY BINNUMBER"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			DisplayStore = trim(sLocName)&" -- "&trim(dcrs(0))
		else
			DisplayStore = trim(sLocName)
		end if
		dcrs.Close
	End Function
%>
<%
	' Function to populate Store
	Function DisplayUoM(sOrgID,iClass,iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>
