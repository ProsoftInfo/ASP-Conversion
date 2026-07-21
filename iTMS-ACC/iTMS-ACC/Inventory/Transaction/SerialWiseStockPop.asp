<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	SerialWiseStockPoP.asp
	'Module Name				:	Inventory (Direct Issue)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 09, 2003
	'Modified By				:	KUMAR K A
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
<!-- #include file="../../include/GetSerialDetail.asp" -->
<%
dim RootNode,HeaderNode,objfs
dim oDOM

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

' Declaration of variables
Dim dcrs,dcrs1,dcrs2,iCtr
iCtr = 0
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

dim iItem,iClass,sItemName,iLot,iLineNo,iPickNo
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,iQty,iInvRecNo
dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore
dim arrUoM,sUoMDesc,sUoMCode,sSubCon,sSalesInv

arrTemp = split(trim(Request.QueryString("sTemp")),"`")
iLot = arrTemp(1)
iClass = arrTemp(2)
iItem = arrTemp(3)
sOrgID = arrTemp(4)
sStoreCode = arrTemp(5)
sBinCode = arrTemp(6)
'iInvRecNo = arrTemp(8)

with dcrs2
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT GROUPNAME,SHORTDESCRIPTION,ORGUNITSHORTDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs2.ActiveConnection = nothing

if not dcrs2.EOF then
	sItemName = trim(dcrs2(0)) & " -- " & trim(dcrs2(1))
	sOrgName = trim(dcrs2(2))
end if
dcrs2.Close

sItemName = ItemDisplay(iItem,iClass)

arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoMDesc = arrUoM(1)

Set objfs = CreateObject("Scripting.FileSystemObject")
if objfs.FileExists(Server.MapPath("../temp/transaction/DI"&Session.SessionID&".xml")) then
	oDOM.Load server.MapPath("../temp/transaction/DI"&Session.SessionID&".xml")
	Set RootNode = oDOM.documentElement
	if RootNode.HaschildNodes() then
		For Each HeaderNode In RootNode.childNodes
			if StrComp(HeaderNode.nodeName,"Header") = 0 then
				sSalesInv = HeaderNode.Attributes.Item(5).nodeValue
				sSubCon = HeaderNode.Attributes.Item(11).nodeValue
				if sSubCon = "" then sSubCon = "NO"
			end if
		next
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/SerialWiseStock.js"></SCRIPT>
<!--SCRIPT LANGUAGE=vbscript SRC="../scripts/DirectIssueLotSerial.vbs"></SCRIPT-->
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">
<form method="POST" name="formname" action="">
<input type=hidden name="hSUBC" value="<%=sSubCon%>">
<input type=hidden name="hSalesInv" value="<%=sSalesInv%>">
<input type=hidden name="hLot" value="<%=iLot%>">
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
                                            <td class="FieldCell">Lot Number</td>
                                            <td class="FieldCellSub" colspan="4">
												<span class="DataOnly"><%=iLot%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub" colspan="4">
												<span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>
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
									<div class="frmBody" id="frm2" style="width: 100%; height:180;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<!--td class="ExcelHeaderCell" align="center" width="10">S.No.</td-->
												<td class="ExcelHeaderCell" align="center">
												<input type="Checkbox" name="ChkAll" class="Formelem" onClick="SelectAll()">All</td>
												<td class="ExcelHeaderCell" align="center">Serial Number</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
											</tr>
										<%
											with dcrs2
												.CursorLocation = 3
												.CursorType = 3
											'	if iLot = "N/A" or iLot = "-" then
											'		.Source = "SELECT DISTINCT SERIALNUMBER,(ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)),Cast(isNull(PackingNumber,0) as Numeric) FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL AND SERIALNUMBER NOT IN (SELECT ISNULL(SERIALNO,0) FROM INV_T_MRSISSUEPICKSERIAL) ORDER BY Cast(isNull(PackingNumber,0) as Numeric)"
											'	else
											'		.Source = "SELECT DISTINCT SERIALNUMBER,(ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)),Cast(isNull(PackingNumber,0) as Numeric) FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL AND SERIALNUMBER NOT IN (SELECT ISNULL(SERIALNO,0) FROM INV_T_MRSISSUEPICKSERIAL) ORDER BY Cast(isNull(PackingNumber,0) as Numeric)"
											'	end if
												if iLot = "N/A" or iLot = "-" then
													.Source = "SELECT DISTINCT SERIALNUMBER,(ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)),Cast(isNull(PackingNumber,0) as Numeric) FROM INV_T_RECEIPTLOTDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL AND SERIALNUMBER NOT IN (SELECT ISNULL(SERIALNO,0) FROM INV_T_MRSISSUEPICKSERIAL) ORDER BY Cast(isNull(PackingNumber,0) as Numeric)"
												else
													.Source = "SELECT DISTINCT SERIALNUMBER,(ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)),Cast(isNull(PackingNumber,0) as Numeric) FROM INV_T_RECEIPTLOTDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL AND SERIALNUMBER NOT IN (SELECT ISNULL(SERIALNO,0) FROM INV_T_MRSISSUEPICKSERIAL) ORDER BY Cast(isNull(PackingNumber,0) as Numeric)"
												end if
												'Response.Write dcrs2.Source
												.ActiveConnection = con
												.Open
											end with
											set dcrs2.ActiveConnection = nothing

											if not dcrs2.EOF then
												Do While Not dcrs2.EOF
													iCtr = iCtr + 1
													if not cdbl(dcrs2(1)) = 0 then
										%>

											<tr>
												<!--td class="ExcelSerial" align="center"><%=iCtr%></td-->
												<td class="ExcelHeaderCell" align="center">
												<input type="Checkbox" name="ChkItem<%=iCtr%>" class="Formelem" onClick="SelectItem(<%=iCtr%>)"></td>
												<td class="ExcelDisplayCell">
													<input type="text" name="txtSerial<%=iCtr%>" size="25" value="<%=GetSerialDetail(trim(dcrs2(0)))%>" class="FormElemRead" READONLY>
													<input type="hidden" name="hSerial<%=iCtr%>" value="<%=trim(dcrs2(0))%>">
												</td>
												<td class="ExcelDisplayCell" width="10">
													<input type="text" name="txtStQty<%=iCtr%>" size="11" maxlength=10 value="<%=cdbl(dcrs2(1))%>" class="FormElemRead" READONLY style="text-align:right">
												</td>
												<td class="ExcelInputCell" width="10" align=center>
												<%'	if lcase(sSubCon) = "no" and Trim(sSalesInv) <> "INV" then %>
													<input type="text" name="txtQty<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align:right">
												<%'	else %>
													<!--input type="checkbox" name="chkQty<%=iCtr%>" value="<%=trim(dcrs2(0))%>" class="FormElem"-->
												<%'	end if %>
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
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
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
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
