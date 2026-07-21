<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	receiptItemEntry.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 25, 2003
	'Modified By				:	KUMAR K A
	'Modified By				:	Ragavendran R.	'Only add one Parameter to redirect the Page
	'Modified On				:	Dec 16,2009
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptInsert.asp
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
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
	Dim oDom,Root,HeaderNode,newElem,newElem1


	'Declaration of Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	dim dcrs,dcrs1,dcrs2,rsTemp
	dim iRecNo,iCtr,iQtyOrd,iQtyRec,iQtyAcc,iQtyRej,iQtyInsRec,iQtyIns
	dim sRec,sRecNumbering,iItmRate,sOrgName,sItmType,sValue,dGDate,sReferenceNo
	dim sStoUoM,sPurUoM,sOperator,iConvRate,iAccQty,sUoMCheck,sItemName
	dim	sRecAgainst,sTraAgainst,iTransferNo,iTransferValue,iTransferQty
	dim dActionQty, sFinancialYearTo, sNo,Sql,sRcptAgnst,sDocumntType,sRcptType
	Dim ReqDate, sFinPeriod, sFinPeriodFrom, sFinPeriodTo, ChkFlg, ChkStr
	Dim sFinFrom, sFinTo, sTempMonYr, sMonYr, arrFin,iActRcptCode,sRcptCode
	Dim sBillNo,sBillDesc,sDCNo,sDCDesc,sDocType,iGrnnum,sRcptDate,sOtherDesc
	Dim sPartyCode,sPartyType,sPartyName,sSubClassCode,sSubItemCode
	Dim sInvNo,sInvDate,sPoNumber,sOCNumber,sDrgStoreNo,sReturnValue,sOrgID,sQuery,sStoreName
	Dim iCountStore,iCtrStore
	Dim iStoreNo,iBinNo,iRecQuantity,iItemEntryNo,iCreatedBy,iStockQuality,sItmAttID
	Dim iInvReceiptNo
	Dim iItemQty

	ChkFlg = False
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
	Set rsTemp = Server.CreateObject("ADODB.RecordSet")

	iRecNo = trim(Request.QueryString("rcptNo"))
	sOrgName = trim(Request.QueryString("sOrg"))
	dGDate = trim(Request.QueryString("gDate"))
	sReturnValue=trim(Request.QueryString("ReturnTo"))
	'Response.write "sReturnValue="&sReturnValue
	dGDate = FormatDate(dGDate)
	sOrgID = Session("organizationcode")
	'Response.Write dGDate
	iCreatedBy = getUserID()

	sFinPeriod = Session("FinPeriod")
	sFinPeriodFrom = FormatDate("04/01/" & Mid(sFinPeriod,1,4))
	sFinPeriodTo = FormatDate("03/31/" & Mid(sFinPeriod,6,4))
	sFinFrom = "01/04/" & Mid(sFinPeriod,1,4)
	sFinTo = "31/03/" & Mid(sFinPeriod,6,4)
'	Response.Write sFinFrom
'	Response.Write sFinTo
	'Response.Write sFinPeriodFrom & sFinPeriodTo

	Set Root = oDOM.createElement("STOCK")
	Root.setAttribute "TRANSDATE",dGDate
	Root.setAttribute "FINFROMDATE",sFinFrom
	Root.setAttribute "FINTODATE",sFinTo
	Root.setAttribute "UNIT",sOrgID
	Root.setAttribute "SRCTYPE","R"
	Root.setAttribute "TRANSACTIONTYPE","R"
	Root.setAttribute "REFTYPE",""
	Root.setAttribute "RECEIPTFOR",""
	Root.setAttribute "RECNO", iRecNo
	Root.setAttribute "RECEIVEDON", dGDate

	oDOM.appendChild Root


	sql = " Select WayBillNo,WayBillDescription,DCNumber,DCDescription,DocumentType,GRNNumber, "
    sql = sql+" convert(varchar,ReceiptDate,103),OtherDescription,ReceiptType,isNull(ReceiptCode,'') from RCV_T_ActualReceiptHeader "
    sql = sql+" where ReceiptNumber='"&iRecNo&"' "
    'Response.Write "<p>"& sql
	with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sql
			.ActiveConnection = con
			.Open
	end with
	if not dcrs.EOF then
		sBillNo = dcrs(0)
		sBillDesc = dcrs(1)
		sDCNo = dcrs(2)
		sDCDesc = dcrs(3)
		sDocType = dcrs(4)
		iGrnnum = dcrs(5)
		sRcptDate = dcrs(6)
		sOtherDesc = dcrs(7)
		sRcptCode = dcrs(8)
		iActRcptCode = dcrs(9)
	  end if
	dcrs.Close


	sql = " Select ReceiptType from Cex_M_ReceiptTypes where ReceiptCode='"&sRcptCode&"' "
	with dcrs
	 		.CursorLocation = 3
	 		.CursorType = 3
	 		.Source = sql
			.ActiveConnection = con
	 		.Open
	end with
	if not dcrs.EOF then
		sRcptType = dcrs(0)
	end if
	dcrs.Close

	sRcptAgnst = sRcptType


	sql = "Select partycode,partytype,isNull(InvoiceNumber,'0'),convert(varchar,InvoiceDate,103) from RCV_T_GateReceiptHeader "
	sql = sql+" where GRNNumber='"&iGrnnum&"' "
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sql
		.ActiveConnection = con
		.Open
	end with
	if not dcrs.EOF then
		sPartyCode = dcrs(0)
		sPartyType = dcrs(1)
		If trim (dcrs(2)) = "0" Then
			'sInvNo = "NA"
			'sInvDate = "NA"
			sInvNo = ""
		Else
			sInvNo = dcrs(2)
			sInvDate = dcrs(3)
		End if
	end if
	dcrs.Close
	sql = " Select PartyName from VwOrgParty where partycode='"&sPartyCode&"' "
	sql = sql+" and partytype='"&sPartyType&"' "
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sql
		.ActiveConnection = con
		.Open
	end with
	if not dcrs.EOF then
		sPartyName = dcrs(0)
	end if
	dcrs.Close

	sql = " Select DocumentType from Cex_M_DocumentTypes where "
	sql = sql+" DocumentTypeNo='"&sDocType&"' "
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sql
		.ActiveConnection = con
		.Open
	end with
	if not dcrs.EOF then
		sDocumntType = dcrs(0)
	end if
	dcrs.Close
iItemEntryNo = 0
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,ISNULL(QUANTITYRECEIVED,0),ISNULL(QUANTITYACCEPTED,0),ISNULL(QUANTITYREJECTED,0),StockQuality,IsNull(ItemRate,0),IsNull(ItemAttributes,'') FROM RCV_T_ACTUALRCPTITEMDET WHERE RECEIPTNUMBER = " & iRecNo & " ORDER BY RECEIPTNUMBER"
		.ActiveConnection = con
		.Open
	end with
	'Response.Write "<textarea>"& dcrs.source&"</textarea>"
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		iCtr = 0
		do while not dcrs.EOF
		    iItemEntryNo = iItemEntryNo +1
			iCtr = iCtr + 1
			iQtyOrd = trim(dcrs(3))
			iQtyRec = trim(dcrs(3))
			iQtyAcc = trim(dcrs(4))
			iQtyRej = trim(dcrs(5))
			iStockQuality = trim(dcrs(6))
			iItmRate = trim(dcrs(7))
			sItmAttID = trim(dcrs(8))

			iQtyAcc = cdbl(iQtyRec)-cdbl(iQtyRej)

			if Trim(iQtyRej)="" or IsNull(iQtyRej) then iQtyRej = 0
			if Trim(iQtyAcc)<>"" then
			    iItemQty = iQtyAcc
			else
			    iItemQty = cdbl(iQtyRec)-CDbl(iQtyRej)
			end if

			sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))

			Set newElem = oDOM.createElement("ITEM")
			newElem.setAttribute "ITEMENTRYNO",iItemEntryNo
			newElem.setAttribute "ITEM",trim(dcrs(0))
			newElem.setAttribute "CLASS",trim(dcrs(1))
			newElem.setAttribute "ITEMQTY", trim(iItemQty)
			newElem.setAttribute "ITEMVALUE", cdbl(iItemQty)*cdbl(iItmRate)
			newElem.setAttribute "ATTRIBUTE", sItmAttID
			newElem.setAttribute "SUMQTY", trim(iItemQty)
			newElem.setAttribute "CREATEDBY", iCreatedBy
			newElem.setAttribute "CREATEDON", date()
			newElem.setAttribute "STOCKQUALITY",iStockQuality
			newElem.setAttribute "ITEMNAME",sItemName

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = " & Pack(trim(dcrs(2))) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 ORDER BY 1,2"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			Do While Not dcrs1.EOF


				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT RECEIPTNUMBERING FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = " & Pack(trim(dcrs(2))) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs2.ActiveConnection = nothing
				'Response.Write dcrs2.source
				if not dcrs2.EOF then
					if Trim(dcrs2(0))="N" then
						sRec = trim(dcrs2(0))
						sRecNumbering ="DISABLED"
					else
						sRec = trim(dcrs2(0))
						sRecNumbering =""
					end if

				else
					sRec = "N"
					sRecNumbering ="DISABLED"
				end if
				dcrs2.close

				Set newElem1 = oDOM.createElement("STORAGE")
				newElem1.setAttribute "STOENTRYNO",iItemEntryNo
				newElem1.setAttribute "ITEM",trim(dcrs(0))
			    newElem1.setAttribute "CLASS",trim(dcrs(1))
				newElem1.setAttribute "STORE",trim(dcrs1(0))
				newElem1.setAttribute "BIN",trim(dcrs1(1))
				newElem1.setAttribute "STOREQTY",trim(iItemQty)
				newElem1.setAttribute "STOREVALUE",cdbl(iItemQty)*cdbl(iItmRate)
				newElem1.setAttribute "DATERECEIVED",date
				newElem.appendChild newElem1


				dcrs1.MoveNext
			loop
			dcrs1.Close

			Root.appendChild newElem

			dcrs.MoveNext
		loop
	end if
	dcrs.Close

	oDOM.Save server.MapPath("../temp/transaction/RECEIPTEX"&Session.SessionID&".xml")

	'storing Lot Details
	'**************
	server.execute("XMLGetLotDetailsForSingleStore.asp")
	'**************

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receipt Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data" data-src="<%="../temp/transaction/RECEIPTEX"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"></script>
<script type="application/xml" data-itms-xml-island="1" id="XmlData"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../scripts/receiptItem.js"></SCRIPT>
</head>
<% ChkStr = CheckFinYr(dGDate)
'Response.Write ChkStr
if ChkStr  = "3" then
%>
	<SCRIPT LANGUAGE=javascript>
		alert("This transaction cannot be performed for this current Financial Year.")
		window.history.back(1)
	</SCRIPT>
<%
elseif ChkStr = "2" then
%>
	<SCRIPT LANGUAGE=javascript>

	</SCRIPT>
<%
elseif ChkStr = "1"  then
%>
	<SCRIPT LANGUAGE=javascript>
		if (!confirm("Since Year End closing has been done and transaction date is in last FY this transaction will be accounted in current financial year. Do you want to proceed?")) {
			window.history.back(1);
		}
	</SCRIPT>
<%
end if


'	Response.Write dGDate
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();">
<form method="POST" name="formname" action="">
<input type="hidden" name="hiRecNo" value="<%=iRecNo%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hGDate" value="<%=dGDate%>">
<input type="hidden" name="hFormattedGDate" value="<%=FormatDate(dGDate)%>">
<input type="hidden" name="hFinFrom" value="<%=sFinPeriodFrom%>">
<input type="hidden" name="hFinTo" value="<%=sFinPeriodTo%>">
<input type="hidden" name="hRec" value="<%=sRec%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Item Receipt
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
								</td>
                            </tr>


                            <tr>
									<td align="middle" class="ClearPixel">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td class="FieldCell">Supplier Name
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sPartyName%></span>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Arrival Date
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sRcptDate%></span>
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Receipt Code
												</td>

												<td class="FieldCellSub">
													<span class="DataOnly"><%=iActRcptCode%></span>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Receipt Date
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sRcptDate%></span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">WayBill No. / Description
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sBillNo%></span>
 													<span class="DataOnly"><%=sBillDesc%></span>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">DC No. / Description
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sDCNo%></span>
 													<span class="DataOnly"><%=sDCDesc%></span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Invoice No. / Date
												</td>
												<td class="FieldCellSub">
													<%if trim(sInvNo) = "" then%>
														<span class="DataOnly">Not Received</span>
													<%else%>
														<span class="DataOnly"><%=sInvNo%></span>
 														<span class="DataOnly"><%=sInvDate%></span>
 													<%end if %>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Other Document
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sOtherDesc%></span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Receipt Against
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sRcptAgnst%></span>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Document Type
												</td>
												<%


												%>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sDocumntType%></span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Receipt Type
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sRcptType%></span>
												</td>
												<td class="FieldCellSub"></td>
												<td class="FieldCellSub">LR Details</td>
												<td class="FieldCellSub">
													<%'added by kalai selvi on  06/09/2004
														' to check LR Detail are Exist or Not
														Sql = "Select LRNumber,GrossWeight,convert(varchar,LRDate,103),NettWeight,PackingCases,VehicleNumber from RCV_T_TransportDetail where RefferenceNo="&iRecNo&" "
														with dcrs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sql
																.ActiveConnection = con
																.Open
														end with
														if not dcrs.EOF then
														%>

														<A onclick="showLRDetailsPopup(' <%=sPartyCode%>','<%=iRecNo%>')" name=txtSamples href="#" >
														<IMG height=11 alt="View LR Details" src  ="../../assets/images/iTMS%20Icons/DetailsIcon.gif" width=11 align=center border=0 >
														</A>
														<%else%>
															<span class="DataOnly">NA</span>
														<%
														end if
													dcrs.Close
												   %>
												</td>

											</tr>
											<tr>
												<td class="FieldCell">Received On
												</td>
												<td class="FieldCellSub">
													<%
														' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDDate")
													%>
												</td>
												<td class="FieldCellSub"></td>
												<td class="FieldCellSub"></td>
												<td class="FieldCellSub">
												</td>

											</tr>

										</table>
									</td>
									<td align="middle" class="ClearPixel">
									</td>
								</tr>


                            <tr>
								<td align="center"></td>
								<td>
	                                <div class="frmBody" id="frm2" style="width: 750; height:300;">
					                    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
												<td class="ExcelHeaderCell" align="center" width="10" rowspan="3">S.No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="3">Item Description</td>
												<td class="ExcelHeaderCell" align="center" colspan="2">Quantity</td>
												<!--<td class="ExcelHeaderCell" align="center" colspan="2">Press</td>
												<td class="ExcelHeaderCell" align="center" rowspan="4">Quantity In</td>-->
												<td class="ExcelHeaderCell" align="center" colspan="3">Account Details</td>
                                            </tr>
                                            <tr>
												<td class="ExcelHeaderCell" align="center">As per DC</td>
												<td class="ExcelHeaderCell" align="center">Accepted</td>
												<!--<td class="ExcelHeaderCell" align="center" rowspan="2">Mark No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Running From</td>-->
												<td class="ExcelHeaderCell" align="center" rowspan="2">Store-Bin</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Quantity</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">UoM</td>
                                            </tr>
                                            <tr>
												<td class="ExcelHeaderCell" align="center">Recd.</td>
												<td class="ExcelHeaderCell" align="center">Rejected & Returned</td>
												<!--td class="ExcelHeaderCell" align="center" rowspan="2">No.of Bales</td-->

                                            </tr>
										<%
											Dim sItemStatus
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,ISNULL(QUANTITYRECEIVED,0),ISNULL(QUANTITYACCEPTED,0),ISNULL(QUANTITYREJECTED,0),ITEMSTATUS FROM RCV_T_ACTUALRCPTITEMDET WHERE RECEIPTNUMBER = " & iRecNo & " ORDER BY RECEIPTNUMBER"
												.ActiveConnection = con
												.Open
											end with
											'Response.Write dcrs.source
											set dcrs.ActiveConnection = nothing
											if not dcrs.EOF then
											iCtr = 0
												do while not dcrs.EOF
													iCtr = iCtr + 1
													iQtyOrd = trim(dcrs(3))
													iQtyRec = trim(dcrs(3))
													iQtyAcc = trim(dcrs(4))
													iQtyRej = trim(dcrs(5))
													sItemStatus = Trim(dcrs(6))
													iQtyAcc = cdbl(iQtyRec)-cdbl(iQtyRej)

													sRec = ""

													with dcrs1
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT isNull(Sum(actiononQty),0) FROM RCV_T_ItemActionTaken WHERE ReceiptNumber = " & iRecNo & " AND EntryNo = " & iCtr & " AND actionon='R'"
														.ActiveConnection = con
														.Open
													end with
'													Response.Write dcrs1.source
													set dcrs1.ActiveConnection = nothing
													if not dcrs1.EOF then
														dActionQty = dcrs1(0)
													end if
													dcrs1.close
													'Response.Write dActionQty

													with dcrs1
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT RECEIPTNUMBERING FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = " & Pack(trim(dcrs(2))) & ""
														.ActiveConnection = con
														.Open
													end with
													set dcrs1.ActiveConnection = nothing
													'Response.Write dcrs1.source
													if not dcrs1.EOF then
														if Trim(dcrs1(0))="N" then
															sRec = trim(dcrs1(0))
															sRecNumbering = "DISABLED"
														else
															sRec = trim(dcrs1(0))
															sRecNumbering = ""
														end if
													else
														sRec = "N"
														sRecNumbering = "DISABLED"
													end if
													dcrs1.close
													sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1)))

													'Response.Write "<textarea>"& sItemStatus &"</textarea>"

													sQuery = "Select IsNull(InventoryRecNo,0) from RCV_T_ActualReceiptHeader where ReceiptNumber = "& iRecNo
													'Response.Write "<textarea>"& sQuery &"</textarea>"
													rsTemp.Open sQuery,con
													if not rsTemp.EOF then
														iInvReceiptNo = rsTemp(0)
													end if
													rsTemp.Close



													if ucase(Trim(iInvReceiptNo))="0" then
													%>
														<tr>
															<input type="hidden" name="hRecNo" value="<%=sRec%>">
															<td class="ExcelSerial" align="center" rowspan="2"><%=iCtr%></td>
															<td class="ExcelDisplayCell" align="left" rowspan="2">
																<a href="javascript:void(0)" class="ExcelDisplayLink" name="lnkA<%=cstr(trim(dcrs(1)))%>A<%=cstr(trim(dcrs(0)))%>A<%=trim(dcrs(2))%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%></a>
															</td>
															<td class="ExcelDisplayCell" align="right"><%=iQtyOrd%></td>
															<td class="ExcelDisplayCell" align="right"><%=cDbl(iQtyAcc)-cDbl(dActionQty)%></td>
															<!--<td class="ExcelInputCell" align="center" width=20>
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
															</td>-->
															<!--td class="ExcelDisplayCell" align="center">
																<select name="selQtyIn<%=iCtr%>" class="FormElem" onChange="DisableTxt(this)">
																	<option value="select">Select</option>
																	<option value="U">Uniform</option>
																	<option value="I">Individual</option>
																</select>
															</td-->
															<td class="ExcelDisplayCell" align="center" rowspan="2">
															<%

															sStoreName=""
															iCtrStore = 0
																sQuery = "SELECT count(IM.LOCATIONNUMBER) FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & dcrs(0) & " AND CLASSIFICATIONCODE = " & dcrs(1) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1"
																rsTemp.Open sQuery,con
																if not rsTemp.EOF then
																	iCountStore =  rsTemp(0)
																end if
																rsTemp.Close
																if iCountStore > 1 then
																	sStoreName = "( Select Store )"
																else
																	sQuery = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & dcrs(0) & " AND CLASSIFICATIONCODE = " & dcrs(1) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 ORDER BY 1,2"
																	rsTemp.Open sQuery,con
																	if not rsTemp.EOF then
																		Do While Not rsTemp.EOF
																			iCtrStore  = iCtrStore  + 1
																			iStoreNo = rsTemp(0)
																			iBinNo = rsTemp(1)
																			sStoreName = DisplayStore(trim(rsTemp(0)),trim(rsTemp(1)))
																			'Response.Write sStoreName
																		rsTemp.MoveNext
																		loop
																	end if
																	rsTemp.Close

																end if' if iCountStore > 1 then
														%>
															<span id="spanStoreName<%=iCtr%>" class="DataOnly"><%=sStoreName%></span>
															<%if iCountStore > 1 then%>
                                                                <img alt="" src="../../assets/images/iTMS%20icons/EntryIcon.gif" onclick="SelectStore('<%=iCtr%>','<%=dcrs(0)%>','<%=dcrs(2)%>')" />
                                                                <input type="hidden" name="hStoreNo<%=iCtr%>" value="" />
                                                                <input type="hidden" name="hBinNo<%=iCtr%>" value="" />
															<%end if%>
															</td>

															<td class="ExcelDisplayCell" align="center"  rowspan="2">
																<%
																	if iCountStore > 1 then
																%>
																	<input type="text" size="12" name="txtZ<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>" class="FormElemRead" Disabled="true">
																<%
																	else
																		if Trim(sRec)="N" then
																		%>
																			<input type="text" size="12" name="txtZ<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>" class="FormElem">
																		<%
																		else
																		%>
																			<input type="text" size="12" name="txtZ<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>" class="FormElem" Disabled="true">
																		<%
																		end if
																	end if
																%>
																<!--<input type="button" value=" Yes " name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>" class="AddButtonX" onClick="GetLot(this)" >-->
																<%
																	if iCountStore > 1 then
																%>
																	<input type="button" value=" Yes " name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>" class="AddButtonX" onClick="GetLotSingle(this,'<%=iCtr%>')" >
																	<!--<img src="../../assets/images/iTMS%20icons/EntryIcon.gif" onClick="GetLot(this)" name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>">-->
																<%
																	else
																%>
																	<input type="button" value=" Yes " name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>:<%=trim(iStoreNo)%>:<%=trim(iBinNo)%>" class="AddButtonX" onClick="GetLot(this,'<%=sRec%>','1','<%=sStoreName%>')"  <%=sRecNumbering%> >
																	<!--<img src="../../assets/images/iTMS%20icons/EntryIcon.gif" onClick="GetLot(this,'<%=sRec%>','1','<%=sStoreName%>')" name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=trim(iStoreNo)%>:<%=trim(iBinNo)%>:<%=cDbl(iQtyAcc)-cDbl(dActionQty)%>:<%=trim(dGDate)%>" <%=sRecNumbering%>>-->
																<%
																	end if
																%>

															</td>
															<td class="ExcelDisplayCell" align="center" rowspan="2"><%=sPurUoM%></td>
														</tr>
														<tr>
															<td class="ExcelDisplayCell" align="right"><%=iQtyRec%></td>
															<td class="ExcelDisplayCell" align="right"><%=cDbl(iQtyRej)+cDbl(dActionQty)%></td>
															<!--<td class="ExcelInputCell" align="center" width=20>
																<input type="text" name="txtCropYear<%=iCtr%>" size="12" maxlength=4 class="FormElem">
															</td>
															<td class="ExcelInputCell" align="center" width=20>
														        <input type="text" name="txtPressRunTo<%=iCtr%>" size="12" maxlength=20 class="FormElem">
															</td>-->
															<!--td class="ExcelInputCell" align="center" width=10>
																<input type="text" name="txtBales<%=iCtr%>" size="11" maxlength=10 class="FormElem" style="text-align=right">
															</td>
															<td class="ExcelInputCell" align="center" width=10>
																<input type="text" name="txtTare<%=iCtr%>" size="11" maxlength=10 class="FormElem" style="text-align=right">
															</td-->
															<!--<td class="ExcelDisplayCell" align="left"><%=sPurUoM%></td>-->
														</tr>
													<%
													else
													%>
														<tr>
															<input type="hidden" name="hRecNo" value="<%=sRec%>">
															<td class="ExcelSerial" align="center" rowspan="2"><%=iCtr%></td>
															<td class="ExcelDisplayCell" align="left" rowspan="2">
																<a href="javascript:void(0)" class="ExcelDisplayLink" name="lnkA<%=cstr(trim(dcrs(1)))%>A<%=cstr(trim(dcrs(0)))%>A<%=trim(dcrs(2))%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%></a>
															</td>
															<td class="ExcelDisplayCell" align="right"><%=iQtyOrd%></td>
															<td class="ExcelDisplayCell" align="right"><%=cDbl(iQtyAcc)-cDbl(dActionQty)%></td>
															<!--<td class="ExcelInputCell" align="center" width=20>
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
															</td>-->
															<!--td class="ExcelDisplayCell" align="center">
																<select name="selQtyIn<%=iCtr%>" class="FormElem" onChange="DisableTxt(this)">
																	<option value="select">Select</option>
																	<option value="U">Uniform</option>
																	<option value="I">Individual</option>
																</select>
															</td-->
															<td class="ExcelDisplayCell" align="center" rowspan="2">
															<%

															sStoreName=""
															iCtrStore = 0

																sQuery = "Select StorageLocationNo,isNull(StorageBinNumber,0) from INV_T_LocationLot where InventoryReceiptNo = "& iInvReceiptNo &" and ItemCode = "& dcrs(0) &" Group By LotNumber,InventoryReceiptNo,StorageLocationNo,StorageBinNumber"
																rsTemp.Open sQuery,con
																if not rsTemp.EOF then
																	Do While Not rsTemp.EOF
																		iCtrStore  = iCtrStore  + 1
																		iStoreNo = rsTemp(0)
																		iBinNo = rsTemp(1)
																		sStoreName = DisplayStore(trim(rsTemp(0)),trim(rsTemp(1)))
																		Response.Write sStoreName
																	rsTemp.MoveNext
																	loop
																end if
																rsTemp.Close

																sQuery = "SELECT STORESUOM FROM VWITEM WHERE ClassificationCode = " & trim(dcrs(1)) & " AND ITEMCODE = " & trim(dcrs(0)) & ""
																rsTemp.Open sQuery,con
																if not rsTemp.EOF then
																	sPurUoM = rsTemp(0)
																end if
																rsTemp.Close

																sQuery = "Select SUM(LotQuantityNett) from INV_T_LocationLot where InventoryReceiptNo = "& iInvReceiptNo &" and ItemCode ="& dcrs(0)
																rsTemp.Open sQuery,con
																if not rsTemp.EOF then
																	iRecQuantity = rsTemp(0)
																end if
																rsTemp.Close

																if trim(iRecQuantity)="" or isnull(iRecQuantity) then iRecQuantity=0.00

														%>

															</td>

															<td class="ExcelDisplayCell" align="center"  rowspan="2">
															<span id="txtQuantity" class="DataOnly"><%=FormatNumber(iRecQuantity,2,0,0,0)%></span>
															<img src="../../assets/images/iTMS%20icons/DetailsIcon.gif" onClick="ViewLotDetails(this)" name="btn:<%=trim(dcrs(1))%>:<%=trim(dcrs(0))%>:<%=trim(dcrs(2))%>:<%=iRecNo%>" alt='Click here to view the Lot Serial Details'>
															</td>
															<td class="ExcelDisplayCell" align="center" rowspan="2"><%=sPurUoM%></td>
														</tr>
														<tr>
															<td class="ExcelDisplayCell" align="right"><%=iQtyRec%></td>
															<td class="ExcelDisplayCell" align="right"><%=cDbl(iQtyRej)+cDbl(dActionQty)%></td>
															<!--<td class="ExcelInputCell" align="center" width=20>
																<input type="text" name="txtCropYear<%=iCtr%>" size="12" maxlength=4 class="FormElem">
															</td>
															<td class="ExcelInputCell" align="center" width=20>
														        <input type="text" name="txtPressRunTo<%=iCtr%>" size="12" maxlength=20 class="FormElem">
															</td>-->
															<!--td class="ExcelInputCell" align="center" width=10>
																<input type="text" name="txtBales<%=iCtr%>" size="11" maxlength=10 class="FormElem" style="text-align=right">
															</td>
															<td class="ExcelInputCell" align="center" width=10>
																<input type="text" name="txtTare<%=iCtr%>" size="11" maxlength=10 class="FormElem" style="text-align=right">
															</td-->
															<!--<td class="ExcelDisplayCell" align="left"><%=sPurUoM%></td>-->
														</tr>

													<%
													end if'if ucase(Trim(sItemStatus))="A" then
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
													<%if Trim(iInvReceiptNo)<>"0" and Trim(iInvReceiptNo)<>"" then%>
														<input type="button" value="Account" name="B1" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(sFinancialYearTo)%>','<%=sReturnValue%>')" disabled>
                                                    <%else%>
														<input type="button" value="Account" name="B1" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(sFinancialYearTo)%>','<%=sReturnValue%>')">
                                                    <%end if%>
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton">
                                                    <input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('receiptEntry.asp')">
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
	' Function to Check for Fin. Year
	Function CheckFinYr(dDate)
		' Declaration of variables
		Dim dcrs
		dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sMonYrNew
		dim sCurYear, sCurYearFrom, sCurYearTo
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		'Response.Write dDate & "        "
		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())
		'Response.Write dDate &"," & sFinPeriodFrom
		arrFin = split(GetFinancialYear(sMonYr),":")
		sCurYearFrom  = arrFin(0)
		sCurYearTo = arrFin(1)
		'Response.Write sFinPeriodFrom & "--" & formatdate(dDate) & "--"& FormatDate(sFinPeriodTo) &"<br>"
		'Response.Write DateDiff("d",FormatDate(sFinPeriodFrom),formatdate(dDate)) & "--"&DateDiff("d",FormatDate(sFinPeriodTo),formatdate(dDate))
		If (DateDiff("d",FormatDate(sFinPeriodFrom),formatdate(dDate)) >= 0) and (DateDiff("d",FormatDate(sFinPeriodTo),formatdate(dDate))<= 0) Then
			CheckFinYr = "2"
		ElseIf (DateDiff("d",FormatDate(sFinPeriodFrom),dDate)<=0) Then
			CheckFinYr = "1"
		Else
			CheckFinYr = "3"
		End If
		If CheckFinYr = "1" Then
			If (DateDiff("d",FormatDate(sFinPeriodTo),date()))<=0 Then
				dGDate = FormatDate(date())
			Else
				dGDate = sFinancialYearTo
			End If
		End If
		sFinancialYearTo = arrFin(1)
		'Response.Write dGDate
	End Function
%>



<%
	' Function to Check for Fin. Year
	Function CheckFinYer(dDate)
		' Declaration of variables
		Dim dcrs
		dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sMonYrNew
		dim sCurYear, sCurYearFrom, sCurYearTo
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		'Response.Write dDate

		'Response.Write "C" & sMonYr & vbCrLf
		sMonYrNew = Mid(ReqDate,4,2) & Mid(ReqDate,7,4)
		'Response.Write "P" & sMonYrNew
'		arrFin = split(GetFinancialYear(sMonYrNew),":")
'		sCurYearFrom = arrFin(0)
'		sCurYearTo = arrFin(1)

		sTempMonYr = Mid(dDate,4,2) & Mid(dDate,7,4)
		sMonYr = sTempMonYr

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)
		sFinancialYearTo = arrFin(1)
		Response.Write sCurYearFrom & "  " & sFinFrom & " " & dDate
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.source
		set dcrs.ActiveConnection = nothing

		if Not dcrs.EOF then
		'Response.Write DateDiff("d",formatdate(sFinFrom),formatdate(dDate))
			if DateDiff("d",formatdate(sFinFrom),formatdate(dDate)) < 0 then
				'Response.Write DateDiff("d",formatdate(sCurYearFrom),formatdate(dDate)) & "ok"
				CheckFinYer = "1"
			else
				CheckFinYer = "2"
			end if
		else
			if DateDiff("d",formatdate(dDate),formatdate(sFinFrom)) > 0 then
				CheckFinYer = "0"
			else
				'dGDate = sFinTo

				CheckFinYer = "1"

				ReqDate = FormatDate(Date())
				'Response.Write "C" & sMonYr & vbCrLf
				sMonYrNew = Mid(ReqDate,4,2) & Mid(ReqDate,7,4)
				'Response.Write "P" & sMonYrNew
				arrFin = split(GetFinancialYear(sMonYrNew),":")
				sFinFrom = arrFin(0)
				sFinTo = arrFin(1)

				sFinancialYearTo = arrFin(1)
			end if
		end if
		dcrs.Close
		'Response.Write sFinancialYearTo
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
