<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	JobWorkReceiptItemEntry.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	TAJUDEEN S
	'Created On					:	September 04, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	JobWorkReceiptInsert.asp
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
<!-- #include File="../../include/PurchaseInvFnc.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
	Dim oDom,Root,HeaderNode,newElem,newElem1

	'Declaration of Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	dim dcrs,dcrs1,iRecNo,iCtr,iQty,iQtyRec,iQtyAcc,iQtyRej,iQtyInsRec,iQtyIns
	dim sRec,sRecNumbering,iItmRate,sOrgName,sOrgID,sItmType,sValue,dGDate
	dim sStoUoM,sPurUoM,sOperator,iConvRate,iAccQty,sUoMCheck,sItemName
	
	Dim iEntryNo,iSrcRefNO,iReceiptAgainst
	
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	
	iRecNo = trim(Request.QueryString("rcptNo"))
	sOrgName = trim(Request.QueryString("sOrg"))
	dGDate = trim(Request.QueryString("gDate"))
	sOrgID = trim(Request.QueryString("sOrgID"))
	iSrcRefNO = trim(Request.QueryString("iSrcRefNO")) 
	iReceiptAgainst = trim(Request.QueryString("ReceiptAgainst"))
	
	Set Root = oDOM.createElement("ROOT")
	Root.setAttribute "RECNO", iRecNo
	Root.setAttribute "SrcRefNO", iSrcRefNO
	Root.setAttribute "ReceiptAgainst", iReceiptAgainst
	oDOM.appendChild Root

	'Adding Received Items
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT FROMITEMCODE, FROMCLASSIFICATIONCODE, ORGANISATIONCODE, RECEIPTQUANTITY, ACTUALRECEIPTENTRYNO FROM SUB_T_SALEARDETAILS WHERE ITEMSTATUS = 'A' AND ACTUALRECEIPTNO = " & iRecNo & " ORDER BY ACTUALRECEIPTNO"
'		Response.Write dcrs.Source 
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		do while not dcrs.EOF
			iQty = trim(dcrs(3))
			iEntryNo = iEntryNo & "," & trim(dcrs(4))
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ORDEREDRATE FROM SUB_T_SALEORDERDETAILS WHERE ORDERNUMBER in (SELECT SOURCEREFNO FROM SUB_T_SALEARHEADER WHERE ACTUALRECEIPTNO = " & iRecNo & ") AND FROMITEMCODE = " & Trim(dcrs(0)) & " AND FROMCLASSIFICATIONCODE = " & Trim(dcrs(1))
'				Response.Write dcrs1.Source 
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				iItmRate = trim(dcrs1(0))
				sValue = "O"
			end if
			dcrs1.Close

			sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))
			if (iItmRate ="" or isnull(iItmRate)) then iItmRate = 0
			if (iQty="" or isNull(iQty)) then iQty = 0
			'if 

			Set newElem = oDOM.createElement("ITEM")
			newElem.setAttribute "ICODE",trim(dcrs(0))
			newElem.setAttribute "CCODE",trim(dcrs(1))
			newElem.setAttribute "OCODE",trim(dcrs(2))
			newElem.setAttribute "ITEMNAME", sItemName
			newElem.setAttribute "QTY", trim(iQty)
			newElem.setAttribute "QTYREC", trim(iQty)
			newElem.setAttribute "ITMRATE", iItmRate
			newElem.setAttribute "VALUETYPE", sValue
			'Newly Added
			newElem.setAttribute "PRUNFROM", ""
			newElem.setAttribute "PRUNTO", ""
			newElem.setAttribute "QTYIN", ""
			newElem.setAttribute "TARE", ""
			newElem.setAttribute "PRESS", ""
			newElem.setAttribute "CROP", ""
			newElem.setAttribute "BALES", ""
			newElem.setAttribute "TAREWEIGHT", ""

			'Adding Location Details
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = " & Pack(trim(dcrs(2))) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 ORDER BY 1,2"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			Do While Not dcrs1.EOF
				Set newElem1 = oDOM.createElement("STORAGE")
				newElem1.setAttribute "STORE",trim(dcrs1(0))
				newElem1.setAttribute "BIN",trim(dcrs1(1))
				newElem1.setAttribute "QTY","0"

				newElem.appendChild newElem1

			dcrs1.MoveNext
			loop
			dcrs1.Close

			Root.appendChild newElem
			dcrs.MoveNext
		loop
	end if
	dcrs.Close

	iEntryNo = mid(iEntryNo,2)

	'Adding BOM Items
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORDBOMITEMCODE, ORDBOMCLASSCODE, SUM(ORDBOMITEMQTY) FROM SUB_T_SALEARITEMBOM WHERE ACTUALRECEIPTNO = " & iRecNo & "  AND ACTUALRECEIPTENTRYNO IN (" & iEntryNo & ") GROUP BY ORDBOMITEMCODE, ORDBOMCLASSCODE"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		do while not dcrs.EOF
			iQty = trim(dcrs(2))
			sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))

			Set newElem = oDOM.createElement("ITEM")
			newElem.setAttribute "ICODE",trim(dcrs(0))
			newElem.setAttribute "CCODE",trim(dcrs(1))
			newElem.setAttribute "OCODE",sOrgID 
			newElem.setAttribute "ITEMNAME", sItemName
			newElem.setAttribute "QTY", trim(iQty)
			newElem.setAttribute "QTYREC", trim(iQty)
			newElem.setAttribute "ITMRATE", ""
			newElem.setAttribute "VALUETYPE", ""
			'Newly Added
			newElem.setAttribute "PRUNFROM", ""
			newElem.setAttribute "PRUNTO", ""
			newElem.setAttribute "QTYIN", ""
			newElem.setAttribute "TARE", ""
			newElem.setAttribute "PRESS", ""
			newElem.setAttribute "CROP", ""
			newElem.setAttribute "BALES", ""
			newElem.setAttribute "TAREWEIGHT", ""

			'Adding Location Details
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_ORGSTORAGE IC,INV_M_ITEMORGSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 ORDER BY 1,2"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			Do While Not dcrs1.EOF
				Set newElem1 = oDOM.createElement("STORAGE")
				newElem1.setAttribute "STORE",trim(dcrs1(0))
				newElem1.setAttribute "BIN",trim(dcrs1(1))
				newElem1.setAttribute "QTY","0"

				newElem.appendChild newElem1

			dcrs1.MoveNext
			loop
			dcrs1.Close

			Root.appendChild newElem
			dcrs.MoveNext
		loop
	end if
	dcrs.Close

	'Adding PROCESSBOM Items
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT BOMITEMCODE, BOMCLASSCODE, SUM(BOMRECQTY) FROM SUB_T_SALEARPROCESSBOM WHERE ACTUALRECEIPTNO = " & iRecNo & "  AND ACTUALRECEIPTENTRYNO IN (" & iEntryNo & ") GROUP BY BOMITEMCODE, BOMCLASSCODE"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		do while not dcrs.EOF
			iQty = trim(dcrs(2))
			sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))

			Set newElem = oDOM.createElement("ITEM")
			newElem.setAttribute "ICODE",trim(dcrs(0))
			newElem.setAttribute "CCODE",trim(dcrs(1))
			newElem.setAttribute "OCODE",sOrgID 
			newElem.setAttribute "ITEMNAME", sItemName
			newElem.setAttribute "QTY", trim(iQty)
			newElem.setAttribute "QTYREC", trim(iQty)
			newElem.setAttribute "ITMRATE", ""
			newElem.setAttribute "VALUETYPE", ""
			'Newly Added
			newElem.setAttribute "PRUNFROM", ""
			newElem.setAttribute "PRUNTO", ""
			newElem.setAttribute "QTYIN", ""
			newElem.setAttribute "TARE", ""
			newElem.setAttribute "PRESS", ""
			newElem.setAttribute "CROP", ""
			newElem.setAttribute "BALES", ""
			newElem.setAttribute "TAREWEIGHT", ""

			'Adding Location Details
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_ORGSTORAGE IC,INV_M_ITEMORGSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 ORDER BY 1,2"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			Do While Not dcrs1.EOF
				Set newElem1 = oDOM.createElement("STORAGE")
				newElem1.setAttribute "STORE",trim(dcrs1(0))
				newElem1.setAttribute "BIN",trim(dcrs1(1))
				newElem1.setAttribute "QTY","0"

				newElem.appendChild newElem1

			dcrs1.MoveNext
			loop
			dcrs1.Close

			Root.appendChild newElem
			dcrs.MoveNext
		loop
	end if
	dcrs.Close

	if CheckFinYear(dGDate) <> "0" then 
%>
<SCRIPT LANGUAGE=javascript>
	alert("Since Year End closing has been done / Transaction date entered is in current Financial Year, this transaction cannot be performed for this current Financial Year.")
	window.history.back(1)
</SCRIPT>
<%	
	Response.End
	end if 
	oDOM.Save server.MapPath("../temp/transaction/JOBWORKRECEIPTEX"&Session.SessionID&".xml")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receipt Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data" data-src="<%="../temp/transaction/JOBWORKRECEIPTEX"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../scripts/JobWorkReceiptItem.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hiRecNo" value="<%=iRecNo%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<input type="hidden" name="hGDate" value="<%=dGDate%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Receipt
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
								</td>
                            </tr>
							<tr>
                                <td class="FieldCell" colspan="3">&nbsp;&nbsp;Received On &nbsp;
							<%
								' Function Call to Insert Date Picker
								Response.Write InsertDatePicker("ctlDDate")
							%>
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td>
	                                <div class="frmBody" id="frm2" style="width: 570; height:410;">
					                    <table border="0" cellspacing="1" class="ExcelTable" width="675">
                                            <tr>
												<td class="ExcelHeaderCell" align="center" width="10" rowspan="4">S.No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="4">Item Description</td>
												<td class="ExcelHeaderCell" align="center" colspan="2">Quantity</td>
												<td class="ExcelHeaderCell" align="center" colspan="2">Press</td>
												<td class="ExcelHeaderCell" align="center" rowspan="4">Quantity In</td>
												<!--td class="ExcelHeaderCell" align="center" rowspan="4">Tare Weight</td-->
												<td class="ExcelHeaderCell" align="center" rowspan="2">Account Details</td>
                                            </tr>
                                            <tr>
												<td class="ExcelHeaderCell" align="center" rowspan="2">As per DC</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Accepted</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Mark No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Running From</td>
                                            </tr>
                                            <tr>
												<!--td class="ExcelHeaderCell" align="center" rowspan="2">No.of Bales</td-->
												<td class="ExcelHeaderCell" align="center" rowspan="2">UoM</td>
                                            </tr>
                                            <tr>
												<td class="ExcelHeaderCell" align="center">Recd.</td>
												<td class="ExcelHeaderCell" align="center">Rejected</td>
												<td class="ExcelHeaderCell" align="center">Crop Year</td>
												<td class="ExcelHeaderCell" align="center">Running To</td>
                                            </tr>
										<%
											'Fetching Received Items
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT FROMITEMCODE, FROMCLASSIFICATIONCODE, ORGANISATIONCODE, RECEIPTQUANTITY, ACTUALRECEIPTENTRYNO FROM SUB_T_SALEARDETAILS WHERE ITEMSTATUS = 'A' AND ACTUALRECEIPTNO = " & iRecNo & " ORDER BY ACTUALRECEIPTNO"
												.ActiveConnection = con
												.Open
											end with
											set dcrs.ActiveConnection = nothing
											if not dcrs.EOF then
												do while not dcrs.EOF
													iCtr = iCtr + 1
													iQty = trim(dcrs(3))

													sRec = ""

													with dcrs1
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT RECEIPTNUMBERING FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = " & Pack(trim(dcrs(2))) & ""
														.ActiveConnection = con
														.Open
													end with
													set dcrs1.ActiveConnection = nothing
													if not dcrs1.EOF then
														sRec = trim(dcrs1(0))
													else
														sRec = "N"
													end if
													dcrs1.close
													sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))
										%>

											<tr>
												<input type="hidden" name="hRecNo" value="<%=sRec%>">
												<td class="ExcelSerial" align="center" rowspan="2"><%=iCtr%></td>
												<td class="ExcelDisplayCell" align="left" rowspan="2">
													<a href="javascript:void(0)" class="ExcelDisplayLink" name="lnkA<%=cstr(trim(dcrs(1)))%>A<%=cstr(trim(dcrs(0)))%>A<%=trim(dcrs(2))%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%></a>
												</td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelInputCell" align="center" width=20>
													<input type="text" name="txtPressMark<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelInputCell" align="center" width=20>
	                                                <input type="text" name="txtPressRunFro<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelDisplayCell" align="center" rowspan="2">
													<select name="selQtyIn<%=iCtr%>" class="FormElem">
														<option value="select">Select</option>
														<option value="G">Gross</option>
														<option value="N">Nett</option>
													</select>
												</td>
												<td class="ExcelDisplayCell" align="center" width="50">
													<input type="button" value=" Yes " name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=iQty%>:<%=sRec%>" class="AddButtonX" onClick="GetLot(this)">
												</td>
									        </tr>
											<tr>
												<td class="ExcelDisplayCell" align="right"><%=iQtyRec%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQtyRej%></td>
												<td class="ExcelInputCell" align="center" width=20>
													<input type="text" name="txtCropYear<%=iCtr%>" size="12" maxlength=4 class="FormElem">
												</td>
												<td class="ExcelInputCell" align="center" width=20>
	                                                <input type="text" name="txtPressRunTo<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelDisplayCell" align="center"><%=DisplayUOM(trim(dcrs(0)),trim(dcrs(1)),sOrgID)%></td>
									        </tr>
										<%
												dcrs.MoveNext
												loop
											end if
											dcrs.Close
										%>
											<tr>
												<td class="ExcelHeaderCell" align="left"></td>
												<td class="ExcelHeaderCell" align="center" colspan="7" width="230">Bill of Materials</td>
											</tr>
										<%
											'Fetching BOM Items
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT ORDBOMITEMCODE, ORDBOMCLASSCODE, SUM(ORDBOMITEMQTY) FROM SUB_T_SALEARITEMBOM WHERE ACTUALRECEIPTNO = " & iRecNo & "  AND ACTUALRECEIPTENTRYNO IN (" & iEntryNo & ") GROUP BY ORDBOMITEMCODE, ORDBOMCLASSCODE"
												.ActiveConnection = con
												.Open
											end with
											'Response.Write "<p> " & dcrs.Source 
											set dcrs.ActiveConnection = nothing
											if not dcrs.EOF then
												do while not dcrs.EOF
													iCtr = iCtr + 1
													iQty = trim(dcrs(2))
													sRec = "N"
													sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))
										%>

											<tr>
												<input type="hidden" name="hRecNo" value="<%=sRec%>">
												<td class="ExcelSerial" align="center" rowspan="2"><%=iCtr%></td>
												<td class="ExcelDisplayCell" align="left" rowspan="2">
													<a href="javascript:void(0)" class="ExcelDisplayLink" name="lnkA<%=cstr(trim(dcrs(1)))%>A<%=cstr(trim(dcrs(0)))%>A<%=sOrgID%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%></a>
												</td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelInputCell" align="center" width=20>
													<input type="text" name="txtPressMark<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelInputCell" align="center" width=20>
	                                                <input type="text" name="txtPressRunFro<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelDisplayCell" align="center" rowspan="2">
													<select name="selQtyIn<%=iCtr%>" class="FormElem" DISABLED>
														<option value="select">Select</option>
														<option value="G">Gross</option>
														<option value="N" Selected>Nett</option>
													</select>
												</td>
												<td class="ExcelDisplayCell" align="center" width="50">
													<input type="button" value=" Yes " name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=sOrgID%>:<%=iQty%>:<%=sRec%>" class="AddButtonX" onClick="GetLot(this)">
												</td>
									        </tr>
											<tr>
												<td class="ExcelDisplayCell" align="right"><%=iQtyRec%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQtyRej%></td>
												<td class="ExcelInputCell" align="center" width=20>
													<input type="text" name="txtCropYear<%=iCtr%>" size="12" maxlength=4 class="FormElem">
												</td>
												<td class="ExcelInputCell" align="center" width=20>
	                                                <input type="text" name="txtPressRunTo<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelDisplayCell" align="center"><%=DisplayUOM(trim(dcrs(0)),trim(dcrs(1)),sOrgID)%></td>
									        </tr>
										<%
												dcrs.MoveNext
												loop
											end if
											dcrs.Close
										%>


										<%
											'Fetching Process BOM Items
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT BOMITEMCODE, BOMCLASSCODE, SUM(BOMRECQTY) FROM SUB_T_SALEARPROCESSBOM WHERE ACTUALRECEIPTNO = " & iRecNo & "  AND ACTUALRECEIPTENTRYNO IN (" & iEntryNo & ") GROUP BY BOMITEMCODE, BOMCLASSCODE"
												.ActiveConnection = con
												.Open
											end with
											'Response.Write "<p> " & dcrs.Source 
											set dcrs.ActiveConnection = nothing
											if not dcrs.EOF then
												do while not dcrs.EOF
													iCtr = iCtr + 1
													iQty = trim(dcrs(2))
													sRec = "N"
													sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))
										%>

											<tr>
												<input type="hidden" name="hRecNo" value="<%=sRec%>">
												<td class="ExcelSerial" align="center" rowspan="2"><%=iCtr%></td>
												<td class="ExcelDisplayCell" align="left" rowspan="2">
													<a href="javascript:void(0)" class="ExcelDisplayLink" name="lnkA<%=cstr(trim(dcrs(1)))%>A<%=cstr(trim(dcrs(0)))%>A<%=sOrgID%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%></a>
												</td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelInputCell" align="center" width=20>
													<input type="text" name="txtPressMark<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelInputCell" align="center" width=20>
	                                                <input type="text" name="txtPressRunFro<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelDisplayCell" align="center" rowspan="2">
													<select name="selQtyIn<%=iCtr%>" class="FormElem" DISABLED>
														<option value="select">Select</option>
														<option value="G">Gross</option>
														<option value="N" Selected>Nett</option>
													</select>
												</td>
												<td class="ExcelDisplayCell" align="center" width="50">
													<input type="button" value=" Yes " name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=sOrgID%>:<%=iQty%>:<%=sRec%>" class="AddButtonX" onClick="GetLot(this)">
												</td>
									        </tr>
											<tr>
												<td class="ExcelDisplayCell" align="right"><%=iQtyRec%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQtyRej%></td>
												<td class="ExcelInputCell" align="center" width=20>
													<input type="text" name="txtCropYear<%=iCtr%>" size="12" maxlength=4 class="FormElem">
												</td>
												<td class="ExcelInputCell" align="center" width=20>
	                                                <input type="text" name="txtPressRunTo<%=iCtr%>" size="12" maxlength=20 class="FormElem">
												</td>
												<td class="ExcelDisplayCell" align="center"><%=DisplayUOM(trim(dcrs(0)),trim(dcrs(1)),sOrgID)%></td>
									        </tr>
										<%
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton">
                                                    <input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('JobWorkReceiptEntry.asp')">
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
