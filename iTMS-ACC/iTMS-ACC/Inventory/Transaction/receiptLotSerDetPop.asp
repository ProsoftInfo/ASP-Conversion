<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	receiptLotSerDetPop.asp - Display the Accounted receipt Informations
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	Ragavendran R
	'Created On					:	Aug 31,2012
	'Modified By				:	
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
<%
Dim rsObj,rsTemp
Dim sStoreName,sStoresUom,sOrgCode,sItemName,sArrValues,sTempValues,sStoresBinName,sQuery,sLotNo
Dim sRcptNumbering
Dim iItemCode,iClassCode,iRecNo,iSNo,iInvReceiptNo,iLocNo,iBinNo

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")

sTempValues = Request.QueryString("Temp")
sArrValues = Split(sTempValues,":")

iClassCode = sArrValues(1)
iItemCode = sArrValues(2)
sOrgCode = sArrValues(3)
iRecNo = sArrValues(4)

sItemName = GetItemName(iItemCode)

sQuery = "Select InventoryRecNo from RCV_T_ActualReceiptHeader where ReceiptNumber = "& iRecNo
rsTemp.Open sQuery,con
if not rsTemp.EOF then
	iInvReceiptNo = rsTemp(0)
end if
rsTemp.Close 

sQuery = "Select ReceiptNumbering,StoresUom from VWItem where ItemCode = "& iItemCode
rsTemp.Open sQuery,con
if not rsTemp.EOF then
	sRcptNumbering = rsTemp(0)
	sStoresUom = rsTemp(1)
end if
rsTemp.Close 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock - Lot / Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
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
											<td class="FieldCell"></td>
											<td class="FieldCellSub"></td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell" colspan="2"></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Item</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idItemName"><%=sItemName%></span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idItemName"><%=sStoresUom%></span>
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
										<table border="0" cellspacing="1" id="tblLot" class="ExcelTable" width="100%">
											<%if Trim(sRcptNumbering)="LS" or Trim(sRcptNumbering)="L" then%>
												<tr>
													<td class="ExcelHeaderCell" align="Left" colspan="10" >Lot No.</td>
												</tr>
											<%end if%>
											<tr>
												<td class="ExcelHeaderCell" align="Left" colspan="10" >Stores - Bin</td>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" >Packing No</td>
												<td class="ExcelHeaderCell" align="center" >Gross Weight</td>
												<td class="ExcelHeaderCell" align="center" >Nett Weight</td>
												<td class="ExcelHeaderCell" align="center" >Tare Weight</td>
											</tr>
											<%
											Response.Write "<font color=red>"
											iSNo = 0
												sQuery = "Select InventoryReceiptNo,isNull(LotNumber,''),StorageLocationNo,isNull(StorageBinNumber,0) from INV_T_LocationLot where InventoryReceiptNo = "& iInvReceiptNo &" and ItemCode = "& iItemCode &" Group By LotNumber,InventoryReceiptNo,StorageLocationNo,StorageBinNumber"
											'	Response.Write "<textarea>"& sQuery &"</textarea>"
												rsObj.Open sQuery,con
												if not rsObj.EOF then
													do while not rsObj.EOF 
														sLotNo = rsObj(1)
														iLocNo = rsObj(2)
														iBinNo = rsObj(3)
														
														if Trim(sRcptNumbering)="LS" or Trim(sRcptNumbering)="L" then
														%>
															<tr>
																<td class="ExcelDisplayCell" align="Left" colspan="10" ><%=sLotNo%></td>
															</tr>
														<%
														end if
														
														sStoresBinName = DisplayStore(iLocNo,iBinNo)
														%>
															<tr>
																<td class="ExcelDisplayCell" align="Left" colspan="10" ><%=sStoresBinName%></td>
															</tr>
														<%
														if Trim(sRcptNumbering)="L" or Trim(sRcptNumbering)="LS" then
															sQuery = "Select PackingNumber,LotQuantityGross,LotQuantityNett,LotQuantityTare from INV_T_LocationLot where InventoryReceiptNo = "& iInvReceiptNo &" and ItemCode = "& iItemCode &" and LotNumber = '"& sLotNo &"'"
														else
															sQuery = "Select PackingNumber,LotQuantityGross,LotQuantityNett,LotQuantityTare from INV_T_LocationLot where InventoryReceiptNo = "& iInvReceiptNo &" and ItemCode = "& iItemCode 
														end if
													'	Response.Write "<textarea>"& sQuery &"</textarea>"
														rsTemp.Open sQuery,con
														if not rsTemp.EOF then
															do while not rsTemp.EOF 
																iSNo = iSNo+1
															%>
																<tr>
																	<td class="ExcelSerial" align="center" ><%=iSNo%></td>
																	<td class="ExcelDisplayCell" align="center" ><%=rsTemp(0)%></td>
																	<td class="ExcelDisplayCell" align="center" ><%=rsTemp(1)%></td>
																	<td class="ExcelDisplayCell" align="center" ><%=rsTemp(2)%></td>
																	<td class="ExcelDisplayCell" align="center" ><%=rsTemp(3)%></td>
																</tr>
															<%
																rsTemp.MoveNext 
															loop
														end if
														rsTemp.Close 
														
														rsObj.MoveNext 
													loop
												end if
												rsObj.Close 
											%>
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
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Close" name="B1" class="ActionButton" onclick="window.close()">
                                                    <!--<input type="reset" value="Reset" name="B2" class="ActionButton" >-->
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
	' Function to populate Item
	Function GetItemName(iItem)
		' Declaration of variables
		Dim dcrs,sItemDesc,sItemShDesc,sClassDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMDESCRIPTION,SHORTDESCRIPTION FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sItemDesc = dcrs(0)
		set sItemShDesc = dcrs(1)

		if Not dcrs.EOF then
			GetItemName = trim(sItemDesc)
		end if
		dcrs.Close

	End Function
%>

<%
	' Function to Display UoM
	Function DisplayUoM(iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ")"
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

<%
	' Function to Get Selling Form
	Function GetSellingForm()
		' Declaration of variables
		Dim dcrs,dcrs1,iCodeLen,iCodeSize,sForm
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CODELENGTH FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form' AND ITEMTYPEID = " & Pack(sItmType) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		If not dcrs.EOF Then
			iCodeLen = cint(dcrs(0))
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(SUM(CODELENGTH)+1,0) FROM APP_M_CODETYPES WHERE DISPLAYORDER < (SELECT DISPLAYORDER FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form') AND ITEMTYPEID = " & Pack(sItmType) & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			If not dcrs1.EOF Then
				iCodeSize = cint(dcrs1(0))
			end if
			dcrs1.Close
		end if
		dcrs.Close

		if iCodeSize > 0 and iCodeLen > 0 then
			sForm = trim(mid(sComCode,iCodeSize,iCodeLen))

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS WHERE CODE = " & Pack(sForm) & " AND ITEMTYPEID = " & Pack(sItmType) & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			If not dcrs.EOF Then
				do while not dcrs.EOF
					'Response.Write "<option value="""&trim(dcrs(0))&""">"&trim(trim(dcrs(1)))&"</option>" & vbCrLf

					Set newElem = oDOM.createElement("SELLFORM")
					newElem.setAttribute "UNITID",trim(dcrs(0))
					newElem.setAttribute "UNITNAME",trim(dcrs(1))

					Root.appendChild newElem

				dcrs.MoveNext
				loop
			end if
			dcrs.Close
		end if
	End Function
%>
<%
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
			.Source = "SELECT BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY BINNUMBER"
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