<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsPickLotPoP.asp
	'Module Name				:	Inventory (MRS Issue Pick Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 25, 2003
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include file="../../include/GetSerialDetail.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
' Declaration of variables
Dim dcrs,dcrs1,dcrs2,iCtr
iCtr = 0
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

dim iItem,iClass,sItemName,iLot,iLineNo,iPickNo
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,iQty,iInvRecNo,iDINo
dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore
dim arrUoM,sUoMDesc,sUoMCode,sQuery,sPackingNo

arrTemp = split(trim(Request.QueryString("sTemp")),"`")
iLineNo = arrTemp(1)
iPickNo = arrTemp(2)
iLot = arrTemp(3)
sStoreCode = arrTemp(4)
sBinCode = arrTemp(5)
iInvRecNo = arrTemp(6)
iMRSNo = arrTemp(7)
iQty = arrTemp(8)
iDINo = arrTemp(9)

with dcrs2
	.CursorLocation = 3
	.CursorType = 3
	if Trim(iMRSNo) <> "" then
		.Source = "SELECT DISTINCT GROUPNAME,SHORTDESCRIPTION,ORGUNITSHORTDESCRIPTION,CONVERT(CHAR,MRSDATE,103),ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE FROM VWMRSISSUEPICK WHERE MRSNUMBER = " & iMRSNo & " AND PICKNUMBER = " & iPickNo & ""
	else
		.Source = "SELECT DISTINCT GROUPNAME,SHORTDESCRIPTION,ORGUNITSHORTDESCRIPTION,CONVERT(CHAR,DIDATE,103),ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE FROM VWMRSISSUEPICK WHERE DINUMBER = " & iDINo & " AND PICKNUMBER = " & iPickNo & ""
	end if
	.ActiveConnection = con
	.Open
end with
set dcrs2.ActiveConnection = nothing

if not dcrs2.EOF then
	sItemName = trim(dcrs2(0)) & " -- " & trim(dcrs2(1))
	sOrgName = trim(dcrs2(2))
	dMRSDate = trim(dcrs2(3))
	sOrgID = trim(dcrs2(4))
	iClass = trim(dcrs2(5))
	iItem = trim(dcrs2(6))
end if
dcrs2.Close

sItemName = ItemDisplay(iItem,iClass)

arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoMDesc = arrUoM(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - MR Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/mrsLotSerial.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>','<%=FormatDate(date())%>')">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Serial Details
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
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item</td>
                                            <td class="FieldCellSub" colspan="4">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity</td>
                                            <td class="FieldCellSub" colspan="4">
                                                <span class="DataOnly" id="idQty"><%=iQty%>&nbsp;</span>
                                                <span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Lot Number</td>
                                            <td class="FieldCellSub" colspan="4">
												<span class="DataOnly"><%=iLot%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Date</td>
                                            <td class="FieldCellSub" colspan="4">
												<input type="radio" name="radDate" value="M" onClick="ChangeDet('M')"> Multiple
												<input type="radio" name="radDate" value="S" onClick="ChangeDet('S')"> Single &nbsp;
												<%
													' Function Call to Insert Date Picker
													Response.Write ValidateDatePicker("ctlDDate")
												%>
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
								<td>
									<div class="frmBody" id="frm2" style="width: 100%; height:200;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Packing Number</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Quantity Issue</td>
												<td class="ExcelHeaderCell" align="center">Date</td>
											</tr>
										<%
											with dcrs2
												.CursorLocation = 3
												.CursorType = 3
												if iLot = "N/A" then
													'.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0) FROM INV_T_RECEIPTLOTDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 ORDER BY 1"
													.Source = "SELECT SERIALNUMBER,ISNULL(AVAILABLENETSTOCK,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND (LOTNUMBER IS NULL or LotNumber = '0' or LotNumber = 'N/A') AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL)  ORDER BY 1"
													'Response.Write dcrs2.Source
												else
													.Source = "SELECT SERIALNUMBER,ISNULL(AVAILABLENETSTOCK,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL)  ORDER BY 1"
												end if
												.ActiveConnection = con
												.Open
											end with

											set dcrs2.ActiveConnection = nothing

											if not dcrs2.EOF then
												Do While Not dcrs2.EOF

													if not cdbl(trim(dcrs2(1))) = 0 then
													iCtr = iCtr + 1
													
													
													sQuery = "Select isNull(PackingNumber,SerialNumber) from INV_T_LocationLot where SerialNumber = "& Trim(dcrs2(0))
													dcrs1.Open sQuery,con
													if not dcrs1.EOF then
														sPackingNo  = Trim(dcrs1(0))
													end if
													dcrs1.Close 
													
										%>

											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell">
													<input type="text" name="txtSerial<%=iCtr%>" size="11" maxlength=10 value="<%=sPackingNo%>" class="FormElemRead" READONLY>
													<input type="hidden" name="hSerial<%=iCtr%>" value="<%=trim(dcrs2(0))%>">
												</td>
												<td class="ExcelDisplayCell" width="10">
													<input type="text" name="txtStQty<%=iCtr%>" size="11" maxlength=10 value="<%=cdbl(trim(dcrs2(1)))%>" class="FormElemRead" READONLY style="text-align:right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtQty<%=iCtr%>" size="11" onkeypress="return DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3,event)" value="0" class="FormElem" style="text-align:right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtDate<%=iCtr%>" size="12" maxlength=10 value="" class="FormElem">
												</td>
											</tr>
										<%
													end if
												dcrs2.MoveNext
												Loop
											end if
											dcrs2.Close
										%>
										</table>
										<input type=hidden name="hiCtr" value="<%=iCtr%>">
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
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date())%>')">
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton">
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
