<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	newreceiptItemEntry.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 20, 2003
	'Modified By				:	KUMAR K A
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	newreceiptInsert.asp
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
<!-- #include file="../../include/NoSeries.asp"-->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
    Response.Write "<font color=#000000#>"
	dim dcrs,dcrs1,dcrs2,iRecNo,iCtr,iQtyOrd,iQtyRec,iQtyAcc,iQtyRej,iQtyInsRec,iQtyIns
	dim iItem,iClass,sOrgID,iMRSNo,iLot,iIssueNo,sItemRate,sLot,dDate,sStage
	dim sRec,sRecNumbering,sSrcType,sSellingType,iWeight,sPackingType,spackingForm
	dim sTempLot,iSeriesNo,iSeriesCode,sTempLotDate,sItemName,sOrgName,iValue,iDINo
	dim sSql, sIssueEntryNo,sQuery,iInvRecNo
	Dim sFinPeriod, sFinPeriodFrom, sFinPeriodTo, sFinancialYearFrom,sFinancialYearto, ChkStr
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
	
	Response.write "<font color=red>"

	sFinPeriod = Session("FinPeriod")
	sFinPeriodFrom = "01/04/" & Mid(sFinPeriod,1,4)
	sFinPeriodTo = "31/03/" & Mid(sFinPeriod,6,4)
	'Response.Write  sFinPeriodFrom & "," & sFinPeriodTo
	iRecNo = trim(Request.QueryString("rcptNo"))
	sOrgID = trim(Request.QueryString("sOrg"))
	dDate = trim(Request.QueryString("iDate"))
	'dDate = FormatDate(dDate)
	iCtr = "0"
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = " & Pack(sOrgID) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sOrgName = trim(dcrs(0))
	end if
	dcrs.close

	''''''''''''''''''''''''''''''''''''''' only for Finished Goods '''''''''''''''''
	' Lot Number - Number Series
	' OrganisationCode,SeriesNumber,SeriesCode,Date
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'LO' AND ITEMTYPE = 'YRN' AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		iSeriesNo = trim(dcrs(0))
		iSeriesCode = trim(dcrs(1))
		sTempLot = GenSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,FormatDate(date())) & "[" & replace(FormatDate(date()),"/","") & "]"
	else
		sTempLotDate = FormatDate(date())
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(RIGHT(LOTNUMBER,5)+1),1) FROM INV_T_LocationLOT WHERE LEFT(LOTNUMBER,10) = " & Pack(sTempLotDate) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		if not dcrs1.EOF then
			sTempLot = sTempLotDate & "-" & string(5-len(trim(dcrs1(0))),"0") & trim(dcrs1(0))
		end if
		dcrs1.Close
	end if
	dcrs.close
	
	sQuery = "Select IsNull(InvRecNo,0) from APP_T_InternalReceiptHeader where InternalReceiptNo = "& iRecNo
    dcrs.open sQuery,con
    if not dcrs.eof then
        iInvRecNo = dcrs(0)
    end if
    dcrs.close


	''''''''''''''''''''''''''''''''''''''' only for Finished Goods '''''''''''''''''

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Internal Receipt Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data">
<root RECNO="<%=iRecNo%>" FGLOT="<%=sTempLot%>">
<%
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT REFTYPE FROM APP_T_INTERNALRECEIPTHEADER WHERE INTERNALRECEIPTNO = " & iRecNo & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sSrcType = trim(dcrs(0))
	end if
	dcrs.close
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT ITEMCODE,CLASSIFICATIONCODE,SUM(QUANTITYRETURN),ISNULL(MRSNUMBER,0),LOTNO,ISSUENO,ISNULL(DINUMBER,0) FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " GROUP BY ITEMCODE,CLASSIFICATIONCODE,MRSNUMBER,LOTNO,ISSUENO,DINUMBER"
		sSql = "SELECT DISTINCT ITEMCODE,CLASSIFICATIONCODE,SUM(QUANTITYRETURN),ISNULL(MRSNUMBER,0),LOTNO,ISSUENO,ISNULL(DINUMBER,0) FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " GROUP BY ITEMCODE,CLASSIFICATIONCODE,MRSNUMBER,LOTNO,ISSUENO,DINUMBER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		do while not dcrs.EOF
			iCtr = iCtr + 1
			sLot = trim(dcrs(4))
			iDINo = trim(dcrs(6))

		'	with dcrs2
		'		.CursorLocation = 3
		'		.CursorType = 3
		'		.Source = "SELECT DISTINCT IssueEntryNo FROM Inv_T_DepartmentStock WHERE IssueNo = " & dcrs(5) & ""
		'		sSql = "SELECT DISTINCT IssueEntryNo FROM Inv_T_DepartmentStock WHERE IssueNo = " & dcrs(5) & ""
		'		.ActiveConnection = con
		'		.Open
		'	end with
		'	set dcrs2.ActiveConnection = nothing
		'		If not dcrs2.EOF then sIssueEntryNo = dcrs2(0)
		'	dcrs2.Close

			if IsNull(sLot) or sLot = "" then
				sLot = "NULL"
				Response.Write "<ITEM ICODE="""&trim(dcrs(0))&""" CCODE="""&trim(dcrs(1))&""" OCODE="""&trim(sOrgID)&""" ITEMNAME ="""" MRSNUMBER="""&trim(dcrs(3))&""" LOTNO="""&sLot&""" ISSUENO="""&trim(dcrs(5))&""" QTYREC="""&trim(dcrs(2))&""" STORE="""" IRATE="""" CTR="""&iCtr&""" REFTYPE="""&trim(sSrcType)&""" DINUMBER="""&iDINo&""">" & vbCrLf
			else
				Response.Write "<ITEM ICODE="""&trim(dcrs(0))&""" CCODE="""&trim(dcrs(1))&""" OCODE="""&trim(sOrgID)&""" ITEMNAME ="""" MRSNUMBER="""&trim(dcrs(3))&""" LOTNO="""&sLot&""" ISSUENO="""&trim(dcrs(5))&""" QTYREC="""&trim(dcrs(2))&""" STORE="""" IRATE="""" CTR="""&iCtr&""" REFTYPE="""&trim(sSrcType)&""" DINUMBER="""&iDINo&""">" & vbCrLf
				sLot = Pack(sLot)
			end if

			with dcrs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT ISNULL(STAGEID,'0'),ISNULL(SUM(QUANTITYRETURN),0) FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " GROUP BY STAGEID"
				.ActiveConnection = con
				.Open
			end with
			set dcrs2.ActiveConnection = nothing
			if not dcrs2.EOF then
				do while not dcrs2.EOF
					sStage = trim(dcrs2(0))
					if sStage = "0" then
						sStage = "NULL"
					else
						sStage = Pack(mid(sStage,1,InStr(1,sStage,":")-1))
					end if

					iValue = "0"

					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						if sStage = "NULL" then
							.Source = "SELECT RATE FROM INV_M_PACKINGQUALITYRATES WHERE STAGEID IS NULL AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & ""
						else
							.Source = "SELECT RATE FROM INV_M_PACKINGQUALITYRATES WHERE STAGEID = " & sStage & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & ""
						end if
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing
					if not dcrs1.EOF then
						iValue = trim(dcrs1(0))
					else
						iValue = "0"
					end if
					dcrs1.close

					iValue = cdbl(iValue) * cdbl(dcrs2(1))

					if sSrcType = "C" then
						with dcrs1
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ISNULL(ITEMVALUE,0) FROM RCV_T_CCIRELEASEDET WHERE RELEASENUMBER = (SELECT SRCREFERENCENO FROM APP_T_INTERNALRECEIPTHEADER WHERE INTERNALRECEIPTNO = " & iRecNo & ")"
							.ActiveConnection = con
							.Open
						end with
						set dcrs1.ActiveConnection = nothing
						if not dcrs1.EOF then
							iValue = trim(dcrs1(0))
						else
							iValue = "0"
						end if
						dcrs1.close
					end if
					sStage = trim(dcrs2(0))
					Response.Write "<STAGE STAGEID="""&sStage&""" IVALUE="""&cdbl(iValue)&""" IQTY="""&cdbl(dcrs2(1))&""">" & vbCrLf

					if sStage = "0" then
						sStage = "NULL"
					else
						sStage = Pack(sStage)
					end if

					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						if sLot = "NULL" then
							.Source = "SELECT SERIALNO,QUANTITYRETURN,PACKINGNUM,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,ISNULL(GROSSQUANTITYRETURN,0),PACKINGFORM FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND (RTRIM(LOTNO) = '' OR LOTNO IS NULL) AND (STAGEID = " & sStage & " OR STAGEID = 0 OR STAGEID IS NULL) ORDER BY SERIALNO"
						else
							.Source = "SELECT SERIALNO,QUANTITYRETURN,PACKINGNUM,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,ISNULL(GROSSQUANTITYRETURN,0),PACKINGFORM FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND LOTNO = " & sLot & " AND (STAGEID = " & sStage & " OR STAGEID = 0 OR STAGEID IS NULL) ORDER BY SERIALNO"
						end if
						.ActiveConnection = con
						.Open
					end with

					set dcrs1.ActiveConnection = nothing
					do while not dcrs1.EOF
						sPackingType = trim(dcrs1(3))
						sSellingType = trim(dcrs1(4))
						iWeight = trim(dcrs1(5))
						spackingForm = trim(dcrs1(7))

						if IsNull(sSellingType) or sSellingType = "" then sSellingType = "NULL"
						if IsNull(iWeight) or iWeight = "" then iWeight = "NULL"
						if IsNull(sPackingType) or sPackingType = "" then sPackingType = "NULL"
						if IsNull(spackingForm) or spackingForm = "" then spackingForm = "NULL"

						Response.Write "<LOTSERIAL SERIAL="""&trim(dcrs1(0))&""" QTY="""&trim(dcrs1(1))&""" PACKNUM="""&trim(dcrs1(2))&""" PACKCODE="""&sPackingType&""" SELLNUM="""&sSellingType&""" WEIGHT="""&iWeight&""" GROSSWEIGHT="""&trim(dcrs1(6))&""" PACKFORM="""&spackingForm&"""/>" & vbCrLf
					dcrs1.MoveNext
					loop
					dcrs1.Close

					Response.Write "</STAGE>"
				dcrs2.MoveNext
				loop
			end if
			dcrs2.Close

			Response.Write "</ITEM>"
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	iCtr = "0"

%>
</root>
</script>
<% ' Response.Write "SELECT SERIALNO,QUANTITYRETURN,PACKINGNUM,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,ISNULL(GROSSQUANTITYRETURN,0),PACKINGFORM FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " AND ITEMCODE = " & trim(dcrs(0)) & " AND CLASSIFICATIONCODE = " & trim(dcrs(1)) & " AND (MRSNUMBER = " & trim(dcrs(3)) & " OR MRSNUMBER IS NULL) AND (RTRIM(LOTNO) = '' OR LOTNO IS NULL) AND ISSUENO = " & trim(dcrs(5)) & " AND (STAGEID = " & sStage & " OR STAGEID = 0 OR STAGEID IS NULL) ORDER BY SERIALNO"  %>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/newreceiptItemAcc.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
function EditRcpt(iRcptNo) {
	document.formname.action = "RECEIPTINTERNALENTRY.ASP?RcptNo=" + encodeURIComponent(iRcptNo);
	document.formname.submit();
}
function DeleteRcpt(iRcptNo) {
	document.formname.action = "InternalReceiptDelete.ASP?RcptNo=" + encodeURIComponent(iRcptNo);
	document.formname.submit();
}
function CloseWin() {
	document.formname.action = "MATERIALRECEIPTS.ASP?RCPT=A";
	document.formname.submit();
}
</script>
<% ChkStr = CheckFinYr(dDate)
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
'Response.Write dDate
%>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hRecNo" value="<%=iRecNo%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hDate" value="<%=dDate%>">
<input type="hidden" name="hFinFrom" value="<%=sFinPeriodFrom%>">
<input type="hidden" name="hFinTo" value="<%=sFinPeriodTo%>">
<input type="hidden" name="hOrgID" value="<%=sOrgID%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" >
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Internal Item Receipt
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
                                <span id="spanDate" class="DataOnly">&nbsp;</span>
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td>
	                                <div class="frmBody" id="frm2" style="width: 760; height:410;">
					                    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
												<td class="ExcelHeaderCell" align="center" colspan="2">Quantity</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">To Account in</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Packing Details</td>
									        </tr>
									        <tr>
												<td class="ExcelHeaderCell" align="center">Actually Received</td>
												<td class="ExcelHeaderCell" align="center">Accepted - Rejected</td>
									        </tr>
										<%
											'Response.Write sSql
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT DISTINCT ITEMCODE,CLASSIFICATIONCODE,SUM(QUANTITYRETURN),ISNULL(MRSNUMBER,0),LOTNO,ISSUENO FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " GROUP BY ITEMCODE,CLASSIFICATIONCODE,MRSNUMBER,LOTNO,ISSUENO"
												'Response.Write "<textarea>"& dcrs.source &"</textarea>"
												.ActiveConnection = con
												.Open
											end with
											set dcrs.ActiveConnection = nothing
											if not dcrs.EOF then
												do while not dcrs.EOF
													iCtr = iCtr + 1
													iItem = trim(dcrs(0))
													iClass = trim(dcrs(1))
													iQtyRec = trim(dcrs(2))
													iMRSNo = trim(dcrs(3))
													sLot = trim(dcrs(4))
													iIssueNo = trim(dcrs(5))
													
													

													if IsNull(sLot) or sLot = "" then
														sLot = "NULL"
													end if
													
													sSql = "Select IsNull(SUM(L.LotQuantityNett),0) from APP_T_InternalReceiptHeader H join INV_T_LocationLot L on H.InvRecNo = L.InventoryReceiptNo  where H.InternalReceiptNo = "& iRecNo&" and L.ItemCode ="& iItem 
													'Response.write "<textarea>"& sSql &"</textarea>"
													dcrs1.open sSql,con
													if not dcrs1.eof then
													    iQtyRec = dcrs1(0)
													    iQtyAcc = iQtyRec
													end if
													dcrs1.close
													
													if cdbl(iQtyRec)=0 then
													    sSql = "Select IsNull(SUM(QuantityReturn),0) from APP_T_InternalReceiptDetails where InternalReceiptNo = "& iRecNo&" and ItemCode ="& iItem 
													    dcrs1.open sSql,con
													    if not dcrs1.eof then
													        iQtyRec = dcrs1(0)
													        iQtyAcc = "0"
													    end if
													    dcrs1.close
													end if

													with dcrs1
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT RECEIPTNUMBERING FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & trim(iItem) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
														.ActiveConnection = con
														.Open
													end with
													set dcrs1.ActiveConnection = nothing
													if not dcrs1.EOF then
														if trim(dcrs1(0)) = "N" then
															sRec = "N"
															sRecNumbering = " DISABLED "
														else
															sRecNumbering = ""
															'if sSrcType = "F" then sRecNumbering = " DISABLED "
															sRec = trim(dcrs1(0))
														end if
													else
														sRec = "N"
														sRecNumbering = " DISABLED "
													end if
													dcrs1.close

													if not iMRSNo = "0" then
														sItemRate = " DISABLED "
													else
														'sItemRate = ""
														sItemRate = " DISABLED "
													end if
													
													iQtyRej = "0"
													sItemName = ItemDisplay(iItem,iClass)
										%>
									        <tr>
											    <td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell" align="left" >
													<a href="javascript:void(0)" class="ExcelDisplayLink" name="lnkA<%=cstr(iClass)%>A<%=cstr(iItem)%>A<%=trim(sOrgID)%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%></a>
												</td>
											    <td class="ExcelDisplayCell" align="right"><%=iQtyRec%></td>
											    <td class="ExcelDisplayCell" align="right"><%=iQtyAcc%> - <%=iQtyRej%></td>
											    <!--<td class="ExcelInputCell" align="center">
													<input type="text" name="txtRate<%=iCtr%>" size="11" style="text-align=right" onkeypress="DoKeyPress('Y',7,3)" class="FormElem" <%=sItemRate%>>
												</td>-->
											    <td class="ExcelFieldCell" align="center">
													<select size="1" name="selStore<%=iCtr%>" class="FormElem">
														<%	'Calling the Function which populates the Store list
															populateStore sOrgID,iItem
														%>

													</select>
												</td>
											    <td class="ExcelFieldCell"><p align="center">
												<%' 	if sSrcType = "F" or sSrcType = "C" or sSrcType = "R" then %>
												<!--		<input type="button" value=" Yes " name="btn:<%=iRecNo%>:<%=iClass%>:<%=iItem%>:<%=sOrgID%>:<%=iQtyRec%>" class="AddButtonX" onClick="GetRate(this)">-->
												<%'	else %>
														<input type="button" value=" Yes " name="btn:<%=iRecNo%>:<%=iClass%>:<%=iItem%>:<%=sOrgID%>:<%=iMRSNo%>:<%=iIssueNo%>:<%=sLot%>" <%=sRecNumbering%> class="AddButtonX" onClick="GetLot(this)">
												<%'	end if %>
												</td>
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
                            <input type=hidden name=hiCtr value="<%=iCtr%>">
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
											    <%if trim(iInvRecNo)="0" then %>
                                                    <input type="button" value="Account" name="B1" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
                                                <%end if %>
                                                    <input type="button" value=" Edit " name="btnEdit" class="ActionButton" onClick="EditRcpt('<%=iRecNo%>')">
                                                    <input type="button" value=" Delete " name="btnEdit" class="ActionButton" onClick="DeleteRcpt('<%=iRecNo%>')">
                                                    <input type="button" value="Back" name="btnBack" class="ActionButton" onClick="CloseWin()">
                                                    
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
	' Function to populate Store
	Function populateStore(sOrgID,iItem)
		' Declaration of variables
		Dim dcrs,dcrs1,sLoc,sBin,sBinName,sLocName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0),LOCATIONNAME,LOCATIONCODE FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & iItem & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' ORDER BY 1,2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sLoc = dcrs(0)
		set sBin = dcrs(1)
		set sLocName = dcrs(2)

		Do While Not dcrs.EOF
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM INV_M_StoreBINDETAILS WHERE LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY BINNUMBER"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = Nothing

			if not dcrs1.EOF then
				Response.Write("<OPTION VALUE="""&trim(sLoc)&":"&trim(dcrs1(0))&""">"&trim(sLocName)&" -- "&trim(dcrs1(2))&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(sLoc)&""">"&trim(sLocName)&"</OPTION>" &vbcrlf)
			end if
			dcrs1.Close

		dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>


<%
	' Function to Check for Fin. Year
	Function CheckFinYr(dDate)
		' Declaration of variables
		Dim dcrs
		dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sMonYrNew
		dim sCurYear, sCurYearFrom, sCurYearTo,sCurDate,sCurDay,sCurMonth

		If Month(Date()) < 10 Then
			sCurMonth = "0"& Month(Date())
		Else
			sCurMonth = Month(Date())
		End If

		If Day(Date()) < 10 Then
			sCurDay = "0" & Day(Date())
		Else
			sCurDay = Day(Date())
		End If
		sCurDate = sCurMonth & "/" & sCurDay & "/" & Year(Date())

		'Response.write sCurDate
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
		'Response.Write DateDiff("d",FormatDate(sFinPeriodFrom),dGDate)
		If (DateDiff("d",FormatDate(sFinPeriodFrom),FormatDate(dDate)) >= 0) and (DateDiff("d",FormatDate(sFinPeriodTo),FormatDate(dDate))<= 0) Then
			CheckFinYr = "2"
		ElseIf (DateDiff("d",FormatDate(sFinPeriodFrom),dDate)<=0) Then
			CheckFinYr = "1"
		Else
			CheckFinYr = "3"
		End If
		'Response.write DateDiff("d",FormatDate(sFinPeriodTo),sCurDate)
		If CheckFinYr = "1" Then
			If (DateDiff("d",FormatDate(sFinPeriodTo),date()))<=0 Then
			'If (DateDiff("d",FormatDate(sFinPeriodTo),sCurDate))<=0 Then 'New
				dDate = FormatDate(date())
				'dDate = FormatDate(sCurDate) 'New
			Else
				dDate = sFinancialYearTo
			End If
		End If
		sFinancialYearTo = arrFin(1)
		'Response.Write dGDate
	End Function
%>

