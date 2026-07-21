<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MatConOPackSelPop.asp
	'Module Name				:	Inventory (Consumption)
	'Author Name				:	R. Ragavendran
	'Created On					:
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
<!-- #include file="../../include/CommonFunctions.asp"-->
<%
' Declaration of variables
Dim dcrs,dcrs1,dcrs2,iCtr
Dim arrTemp,sIntRcptNo,iIntItemCode,sAttList,iIntClassCode,iIntQty
Dim sOrgID,sItemName,arrUOM,sUoMCode,sUoMDesc,sQuery,sRequest,sRcptNum
Dim iQty,iSerial,iLotNo
iCtr = 0
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
sRequest = Request.QueryString("sTemp")
arrTemp = split(trim(sRequest),":")
Response.Write "<font color=red>"
sIntRcptNo = arrTemp(0)
iIntItemCode= arrTemp(1)
sAttList = arrTemp(2)
iIntQty = arrTemp(3)
sRcptNum = arrTemp(4)

iIntClassCode = split(GetClassification(iIntItemCode),":")(0)

sOrgID = session("organizationcode")

sItemName = ItemDisplay(iIntItemCode,iIntClassCode)

arrUoM = split(DisplayUoM(sOrgID,iIntClassCode,iIntItemCode),":")
sUoMCode = arrUoM(0)
sUoMDesc = arrUoM(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : Lot / Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/MatConOPackSel.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Lot / Serial Details
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
                                            <td class="FieldCell">Item</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idQty"><%=iIntQty%>&nbsp;</span>
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
									<div class="frmbody" id="frm2" style="width: 100%; height:160;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Serial Number</td>
												<td class="ExcelHeaderCell" align="center">Quantity Remaining</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
											</tr>
										<%
										    sQuery = "Select IsNull(LotQuantityNett,0),LotNumber,SerialNumber,PackingNumber from INV_T_LocationLot L join APP_T_InternalReceiptHeader H on L.InventoryReceiptNo = H.InvRecNo where H.InternalReceiptNo = "& sIntRcptNo &" and L.ItemCode = "& iIntItemCode
										    sQuery = sQuery & " and L.SerialNumber not in (Select Serialno from INV_T_MaterialConsumptionOutput H join INV_T_MaterialConsumptionOutputDet D on H.ConsumptionNo = D.ConsumptionNo  and H.IssueEntryNo = D.IssueEntryNo and H.LineNumber = D.LineNumber)"
										    if trim(sAttList)<>"" then
										        sQuery = sQuery & "  and L.AttributeList ="& sAttList
										    end if
										    dcrs2.Open sQuery,con
										    'Response.Write sQuery
										    if not dcrs2.EOF then
											    Do While Not dcrs2.EOF
											        iQty = dcrs2(0)
											        iSerial = dcrs2(2)
													iCtr = iCtr + 1
													'Response.Write iCtr
										        %>

											        <tr>
												        <td class="ExcelSerial" align="center"><%=iCtr%></td>
												        <td class="ExcelDisplayCell">
													        <input type="text" name="txtSerial<%=iCtr%>" value="<%=trim(dcrs2(3))%>" class="FormElemRead" READONLY>
													        <input type="hidden" name="hSerial<%=iCtr%>" value="<%=trim(iSerial)%>">
												        </td>
												        <td class="ExcelDisplayCell" width="10">
													        <input type="text" name="txtStQty<%=iCtr%>" size="11" maxlength=10 value="<%=trim(iQty)%>" class="FormElemRead" READONLY style="text-align:right">
												        </td>
												        <td class="ExcelInputCell" width="10">
													        <input type="text" name="txtQty<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align:right">
												        </td>
											        </tr>
										        <%
												dcrs2.MoveNext
												Loop
											end if
											dcrs2.Close
										%>
										</table>
										<input type=hidden name="hCtr" value="<%=iCtr%>">
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit('<%=sRequest%>')">
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
