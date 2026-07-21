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
	'Program Name				:	IssRetSerialPop.asp
	'Module Name				:	Inventory 
	'Author Name				:	R. Ragavendran
	'Created On					:	Jun 27,2011
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
<!-- #include file="../../include/GetSerialDetail.asp" -->
<%
' Declaration of variables
Dim dcrs,dcrs1,dcrs2,iCtr
iCtr = 0
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

dim iItem,iClass,sItemName,iLot,dIssDate
dim arrTemp,iMRSNo,sOrgID,iIssNo,iQty,iSerial,iLineNo
dim arrUoM,sUoMDesc,sUoMCode,sType,iDINo,sQuery
Dim iPackCode,iPackNum,iAttributeList,sAttID
arrTemp = split(trim(Request.QueryString("sTemp")),"`")
Response.Write "<font color=red>"
'Response.Write Request.QueryString 
sType = arrTemp(1)
sOrgID = arrTemp(2)
iItem = arrTemp(3)
iClass = arrTemp(4)
iLot = arrTemp(5)
iQty = arrTemp(7)
iIssNo = arrTemp(8)
dIssDate = arrTemp(9)
sAttID = arrTemp(10)

sItemName = ItemDisplay(iItem,iClass)

arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
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
<script LANGUAGE=javascript SRC="../scripts/IssRetSerial.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0  onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot / Serial Details
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
                                                <span class="DataOnly" id="idQty"><%=iQty%>&nbsp;</span>
                                                <span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Lot Number&nbsp;</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=iLot%>&nbsp;</span>
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
									<div class="frmBody" id="frm2" style="width: 100%; height:160;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Serial Number</td>
												<td class="ExcelHeaderCell" align="center">Quantity Remaining</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
											</tr>
										<%
										    if iLot <>"N/A" then
										        sQuery = "SELECT SUM(ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))),LOTNO,SERIALNO,ITEMENTRYNO FROM INV_T_MaterialIssueDetails D,INV_T_MaterialIssueHeader H WHERE D.IssueEntryNo = H.IssueEntryNo and D.IssueEntryNo = "& iIssNo &" AND (ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))) > 0 And LotNo = "& Pack(iLot)&" and Convert(Varchar,IssueDate,103) = Convert(Varchar,'"& dIssDate &"',103) Group By LotNo,SerialNo,ItemEntryNo"
										    else
										        sQuery = "SELECT SUM(ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))),LOTNO,SERIALNO,ITEMENTRYNO FROM INV_T_MaterialIssueDetails D,INV_T_MaterialIssueHeader H WHERE D.IssueEntryNo = H.IssueEntryNo and D.IssueEntryNo = "& iIssNo &" AND (ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))) > 0 And  (LotNo is Null or LotNo='NULL' or LotNo ='0') and Convert(Varchar,IssueDate,103) = Convert(Varchar,'"& dIssDate &"',103) Group By LotNo,SerialNo,ItemEntryNo"
										    end if
										    dcrs2.Open sQuery,con
										    'Response.Write sQuery
										    if not dcrs2.EOF then
											    Do While Not dcrs2.EOF
											        iQty = dcrs2(0)
											        iSerial = dcrs2(2)
													iCtr = iCtr + 1
													'Response.Write iCtr
													
													sQuery = "Select PackingCode,PackingNumber,AttributeList from INV_T_LocationLot Where SerialNumber in ("& iSerial &")Order by InventoryReceiptNo desc"
													dcrs.open sQuery,con
													if not dcrs.eof then
													    iPackCode   = dcrs(0)
													    iPackNum    = dcrs(1)
													    iAttributeList = dcrs(2)
													end if
													dcrs.close 
													
										%>

											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell">
													<!--<input type="text" name="txtSerial<%=iCtr%>" value="<%=GetSerialDetail(trim(iSerial))%>" class="FormElemRead" READONLY>-->
													<input type="text" name="txtSerial<%=iCtr%>" value="<%=trim(iSerial)%>" class="FormElemRead" READONLY>
													<input type="hidden" name="hSerial<%=iCtr%>" value="<%=trim(iSerial)%>">
													<input type="hidden" name="hPackCode<%=iCtr%>" value="<%=trim(iPackCode)%>">
													<input type="hidden" name="hPackNumber<%=iCtr%>" value="<%=trim(iPackNum)%>">
													<input type="hidden" name="hAttributeList<%=iCtr%>" value="<%=trim(iAttributeList)%>">
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
