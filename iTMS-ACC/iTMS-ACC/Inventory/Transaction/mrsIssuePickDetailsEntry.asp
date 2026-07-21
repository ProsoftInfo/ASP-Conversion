<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsIssuePickDetailsEntry.asp
	'Module Name				:	Inventory (MRS Issue Pick Details)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 21,2012
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsIssuePickInsert.asp
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
' Declaration of variables
Dim dcrs,dcrs1,dcrs2,iCtr,bexists
'Declaration of Objects
iCtr = 0
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

dim sql,sUnitName,iItem,iClass,sItemName,iLot,iInvRecNo,sStatus,iSchNo,dSchMinDate,iSerialNo
dim arrTemp,iIssueEntryNo,sOrgID,sOrgName,dMRSDate,bFlag,iQty,iLotQty,iEntryNo,iQtyForIssue
dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore,iQtyIssue,iTotLotQty,bSerialFlag
dim dSchMaxDate,iLotQtyReserved,sReqBy,sReqValue,sTempStore,sUoM,iQtyWOLot,iQtyWithLot
dim arrUoM,sUoMCode,iDINo,sIssueCode,sAttributeList
dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sAttID,sOptName,sQuery,sReceiptNumbering

if len(Month(date())) = 1 then
	sTempMonYr = "0"&Month(date())
else
	sTempMonYr = Month(date())
end if
sMonYr = sTempMonYr&Year(date())

arrFin = split(GetFinancialYear(sMonYr),":")
sFinFrom = arrFin(0)
sFinTo = arrFin(1)

iTotLotQty = 0
iLotQtyReserved = 0
bSerialFlag = false
Response.Write "<font color=#000000>"
arrTemp = split(trim(Request.QueryString("sTemp")),"|")
iClass = arrTemp(0)
iItem = arrTemp(1)
iQtyIssue = arrTemp(2)
iIssueEntryNo = arrTemp(3)
sAttID = arrTemp(4)
iEntryNo =  arrTemp(5)

if sAttID="0" then sAttID = ""
sQuery = "Select GroupName,IsNull(ItemDescription,''),OrgUnitShortDescription,Convert(varchar,IssueDate,103),"
sQuery = sQuery &" OrganisationCode,IsNull(IssueEntryCode,IssueEntryNo),ItemEntryNo from VW_INV_IssuedForPick "
sQuery = sQuery &" where IssueEntryNo = "&iIssueEntryNo &" and ItemCode="& iItem &" and ClassificationCode ="& iClass
if trim(sAttID)<>"" then
    sQuery = sQuery &" and ItemAttributes in ("& sAttID &")"
end if
if trim(iEntryNo)<>"" then
    sQuery = sQuery & " and ItemEntryNo = "& iEntryNo
end if
'Response.write "<textarea>"& sQuery &"</textarea>"
dcrs2.open sQuery,con
if not dcrs2.EOF then
	sItemName = trim(dcrs2(0)) & " -- " & trim(dcrs2(1))
	sOrgName = trim(dcrs2(2))
	dMRSDate = trim(dcrs2(3))
	sOrgID = trim(dcrs2(4))
	sIssueCode = trim(dcrs2(5))
	iEntryNo = trim(dcrs2(6))
end if
dcrs2.Close

sItemName = ItemDisplay(iItem,iClass)
if Trim(sAttID)<>"" and Trim(sAttID)<>"NULL" then
    sQuery = "Select OptionName from INV_M_ITEMTYPEOPTIONS where OptionValue = "& sAttID
    dcrs2.Open sQuery,con
    if not dcrs2.EOF then
        sOptName = " ["&trim(dcrs2(0))&"]"
    end if
    dcrs2.Close
end if

if trim(sOptName)<>"" then
    sItemName = sItemName & sOptName
end if
arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoM = arrUoM(1)
'Response.Write "sUoMCode ="& sUoMCode &"***"& sUoM
iQtyWOLot = 0
iQtyWithLot = 0

sQuery = "Select IsNull(SUM(QuantityForPick),0),IsNull(SUM(QuantityPicked),0) from VW_INV_IssuedForPick "
sQuery = sQuery & " WHERE IssueEntryNo = "& iIssueEntryNo &" AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass
sQuery = sQuery & " AND (ISNULL(QuantityForPick,0) - ISNULL(QuantityPicked,0) > 0) AND LocationNumber IS NULL AND BINNUMBER IS NULL"
'Response.write "<textarea>"& sQuery &"</textarea>"
dcrs2.open sQuery,con
if not dcrs2.EOF then
	iQtyWOLot = cdbl(dcrs2(0)) - cdbl(dcrs2(1))
end if
dcrs2.close

sQuery = "Select ReceiptNumbering from VWItem where ItemCode = "& iItem
dcrs2.open sQuery,con
if not dcrs2.eof then
    sReceiptNumbering = trim(dcrs2(0))
end if
dcrs2.close

'Response.Write iQtyWithLot
iQtyWithLot = cdbl(iQtyIssue) - cdbl(iQtyWOLot)
'Response.Write iQtyWithLot
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - MR Issue Pick Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	'XML DOM Variables
	Dim oDOM

	dim RootNode,HeaderNode,PageNode,EntryNode
	dim iSerNo
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	set RootNode = oDOM.createElement("PickDet")
	RootNode.setAttribute "CLAS",iClass
	RootNode.setAttribute "ITM",iItem
	RootNode.setAttribute "TOT",""

	oDOM.appendChild(RootNode)

	sQuery = "Select LocationNumber,BinNumber,LotNo,IsNull(SUM(QuantityForPick),0),IsNull(SUM(QuantityPicked),0),IsNull(ItemAttributes,'')"
    sQuery = sQuery & " from VW_INV_IssuedForPick where IssueEntryNo =" & iIssueEntryNo & " and ItemCode = "& iItem &""
    sQuery = sQuery & " and ClassificationCode = "& iClass &" and (IsNull(QuantityForPick,0)-IsNull(QuantityPicked,0)>0)"

	if trim(sAttID)<>"" and trim(sAttID)<>"NULL" then
		sQuery = sQuery & " and ItemAttributes = '"& sAttID&"'"
	end if
	sQuery = sQuery &" Group By LocationNumber,BinNumber,LotNo,ItemAttributes"
	'Response.write "<textarea>"& sQuery &"</textarea>"
'	response.end
	dcrs2.open sQuery,con
	if not dcrs2.EOF then
		Do While Not dcrs2.EOF
			sStoreCode = trim(dcrs2(0))
			sBinCode = trim(dcrs2(1))
			iLot = trim(dcrs2(2))
			iQtyForIssue = cdbl(dcrs2(3)) - cdbl(dcrs2(4))
			sReqBy = "I"
			sReqValue = FormatDate(Date)
			sAttributeList = trim(dcrs2(5))

			if sReqBy = "S" then
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT MIN(CONVERT(CHAR,SCHEDULEDON,103)),MAX(CONVERT(CHAR,SCHEDULEDON,103)) FROM INV_T_MRSITEMSCHEDULES WHERE SCHEDULETYPE = 'D' AND MRSNUMBER = " & iIssueEntryNo & " AND CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItem & " AND CONVERT(DATETIME,SCHEDULEDON,103) <= CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103)"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
				if not dcrs1.EOF then
					dSchMinDate = trim(dcrs1(0))
					dSchMaxDate = trim(dcrs1(1))
				end if
				dcrs1.Close
			elseif sReqBy = "I" or sReqBy = "D" then
				dSchMinDate = sReqValue
				dSchMaxDate = sReqValue
			elseif sReqBy = "W" then
				dSchMinDate = FormatDate(DateAdd("d",cdbl(sReqValue),FormatDate(dMRSDate)))
				dSchMaxDate = FormatDate(DateAdd("d",cdbl(sReqValue),FormatDate(dMRSDate)))
			end if
			'Response.Write "sStoreCode="&sStoreCode
			if not (sStoreCode = "0" or IsNull(sStoreCode)) then
				if IsNull(iLot) or iLot = "" then iLot = "NULL"
				if sBinCode = "" or IsNull(sBinCode) then sBinCode = "0"

                if trim(iLot)="NULL" then
                sQuery = " SELECT isNull(Sum(AVAILABLENETSTOCK),0),ISNULL(SERIALNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND  "&_
	                     " ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND CONVERT(DATETIME,DATEOFRECEIPT,103) BETWEEN "&_
		                 " CONVERT(DATETIME,'" & sFinFrom & "',103) AND  CONVERT(DATETIME,'" &sFinTo & "',103) and (LotQuantityNett - QuantityIssued) > 0 GROUP BY AVAILABLENETSTOCK,INVENTORYRECEIPTNO,LOTNUMBER,SRCTYPE,SERIALNUMBER "
                else
                sQuery = " SELECT isNull(SUM(AVAILABLENETSTOCK),0),ISNULL(SERIALNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND LOTNUMBER = " & Pack(iLot) & " AND  "&_
   		                 " ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND CONVERT(DATETIME,DATEOFRECEIPT,103) BETWEEN "&_
		                 " CONVERT(DATETIME,'" & sFinFrom & "',103) AND  CONVERT(DATETIME,'" & sFinTo & "',103) and (LotQuantityNett - QuantityIssued) > 0 GROUP BY AVAILABLENETSTOCK,INVENTORYRECEIPTNO,LOTNUMBER,SRCTYPE,SERIALNUMBER "
                end if
                'Response.Write "<textarea>"& sQuery &"</textarea>"
                dcrs.open sQuery,con
				Do while not dcrs.EOF
					iLotQty = cdbl(trim(dcrs(0)))
					iSerNo = cdbl(trim(dcrs(1)))

					iTotLotQty = cdbl(iTotLotQty) + cdbl(iLotQty)
					'Response.Write "iTotLotQty="&iTotLotQty
					IF iSerNo <> 0 then
						bSerialFlag = True
					Else
						bSerialFlag = false
					End IF

					dcrs.movenext
				loop
				dcrs.close
				'Response.Write "bSerialFlag="&bSerialFlag
				iSerialNo = iSerialNo + 1
				set HeaderNode = oDOM.createElement("PICK")
				HeaderNode.setAttribute "ITEMENTRYNO",trim(iEntryNo)
				HeaderNode.setAttribute "ISSUEENTRYNO",trim(iIssueEntryNo)
				HeaderNode.setAttribute "LOC",trim(sStoreCode)
				HeaderNode.setAttribute "BIN",trim(sBinCode)
				HeaderNode.setAttribute "LOTNO",iLot
				HeaderNode.setAttribute "QTYFORISS",iQtyForIssue
				HeaderNode.setAttribute "ISSQTY",""
				HeaderNode.setAttribute "SCHMINDATE",dSchMinDate
				HeaderNode.setAttribute "SCHMAXDATE",dSchMaxDate
				HeaderNode.setAttribute "STOCK",iTotLotQty
				HeaderNode.setAttribute "SNO",iSerialNo
				HeaderNode.setAttribute "ATTID",sAttributeList

				RootNode.appendChild(HeaderNode)
			else

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNull(BINNUMBER,0) FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN'"
					'Response.Write dcrs1.Source
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				Do While Not dcrs1.EOF
					sStoreCode = trim(dcrs1(0))
					sBinCode = trim(dcrs1(1))

					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM inv_t_itemLocationstock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
						.ActiveConnection = con
					'	Response.Write dcrs.source
						.Open
					end with
					set dcrs.ActiveConnection = nothing

					if not dcrs.EOF then
						iQty = cdbl(trim(dcrs(0)))
					end if
					dcrs.close

					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT DISTINCT ISNULL(SUM(AVAILABLENETSTOCK),0),STORAGELOCATIONNO,STORAGEBINNUMBER,INVENTORYRECEIPTNO,ISNULL(LOTNUMBER,NULL) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND (STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL)) AND CONVERT(DATETIME,DATEOFRECEIPT,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) GROUP BY STORAGELOCATIONNO,STORAGEBINNUMBER,INVENTORYRECEIPTNO,LOTNUMBER ORDER BY 2,3"
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing

					if not dcrs.EOF then
						Do While Not dcrs.EOF
							iTotLotQty = cdbl(iTotLotQty) + cdbl(trim(dcrs(0)))

							if isNull(dcrs(4)) then
								iLot = "NULL"
							else
								iLot = trim(dcrs(4))
							end if

							iSerialNo = iSerialNo + 1
							set HeaderNode = oDOM.createElement("PICK")
							HeaderNode.setAttribute "ITEMENTRYNO",trim(iEntryNo)
							HeaderNode.setAttribute "ISSUEENTRYNO",trim(iIssueEntryNo)
							HeaderNode.setAttribute "LOC",trim(sStoreCode)
							HeaderNode.setAttribute "BIN",trim(sBinCode)
							HeaderNode.setAttribute "LOTNO",iLot
							HeaderNode.setAttribute "QTYFORISS","-"
							HeaderNode.setAttribute "ISSQTY",""
							HeaderNode.setAttribute "SCHMINDATE",dSchMinDate
							HeaderNode.setAttribute "SCHMAXDATE",dSchMaxDate
							HeaderNode.setAttribute "STOCK",iTotLotQty
							HeaderNode.setAttribute "SNO",iSerialNo
							HeaderNode.setAttribute "ATTID",sAttributeList

							RootNode.appendChild(HeaderNode)

						dcrs.MoveNext
						Loop
					end if
					dcrs.Close

					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing
					if Not dcrs.EOF then
						if not (cdbl(trim(dcrs(0))) - cdbl(iTotLotQty)) <= 0 then

							iSerialNo = iSerialNo + 1
							set HeaderNode = oDOM.createElement("PICK")
							HeaderNode.setAttribute "ITEMENTRYNO",trim(iEntryNo)
							HeaderNode.setAttribute "ISSUEENTRYNO",trim(iIssueEntryNo)
							HeaderNode.setAttribute "LOC",trim(sStoreCode)
							HeaderNode.setAttribute "BIN",trim(sBinCode)
							HeaderNode.setAttribute "LOTNO","NULL"
							HeaderNode.setAttribute "QTYFORISS","-"
							HeaderNode.setAttribute "ISSQTY",""
							HeaderNode.setAttribute "SCHMINDATE",dSchMinDate
							HeaderNode.setAttribute "SCHMAXDATE",dSchMaxDate
							HeaderNode.setAttribute "STOCK", iTotLotQty
							HeaderNode.setAttribute "SNO",iSerialNo
							HeaderNode.setAttribute "ATTID",sAttributeList

							RootNode.appendChild(HeaderNode)
						end if
					end if
					dcrs.Close
				dcrs1.MoveNext
				loop
				dcrs1.Close

			end if
		dcrs2.MoveNext
		Loop
	end if
	dcrs2.Close

	dim sExp,tempNode,iNdCtr,sNodeLoc,sNodeBin,sTemp,iLocBinCtr,sTempLocBin
	dim arrTempLocBin,iarrCtr,arrLocBin,sTempBin,sTempLoc,arrTempLot,iLotCtr,sTempLot
	dim sNodeLot,iTempCtr,bLotFlag,bQtyFlag,iSerialCtr

	iLocBinCtr = 0
	sExp = "//PickDet/PICK"
	Set EntryNode = RootNode.Selectnodes(sExp)

	for iNdCtr =  0 to EntryNode.length - 1
		sNodeLoc = EntryNode.Item(iNdCtr).Attributes.getNamedItem("LOC").Value
		sNodeBin = EntryNode.Item(iNdCtr).Attributes.getNamedItem("BIN").Value
		if sTemp = "" then
			iLocBinCtr = iLocBinCtr + 1
			sTemp = sNodeLoc&"`"&sNodeBin
			sTempLocBin = sTemp
		elseif sTemp <> sNodeLoc&"`"&sNodeBin then
			iLocBinCtr = iLocBinCtr + 1
			sTemp = sNodeLoc&"`"&sNodeBin
			sTempLocBin = sTempLocBin&"|"&sTemp
		end if
	next

	arrTempLocBin = split(sTempLocBin,"|")

	sTemp = ""
	iNdCtr = 0
	for iarrCtr = 0 to UBound(arrTempLocBin)
		arrLocBin = split(arrTempLocBin(iarrCtr),"`")
		sTempLoc = arrLocBin(0)
		sTempBin = arrLocBin(1)

		sExp = "//PickDet/PICK [@LOC = "&sTempLoc&" and @BIN = '"&sTempBin&"']"
		Set tempNode = RootNode.Selectnodes(sExp)
		for iNdCtr =  0 to tempNode.length - 1
			sNodeLot = tempNode.Item(iNdCtr).Attributes.getNamedItem("LOTNO").Value
			if sTemp = "" then
				iLotCtr = iLotCtr + 1
				sTemp = sNodeLot
				sTempLot = sTemp
			elseif instr(1,sTempLot,sNodeLot) <= 0 then
				iLotCtr = iLotCtr + 1
				sTemp = sNodeLot
				sTempLot = sTempLot&"|"&sTemp
			end if
		next
		'Response.Write "Store >>>> " & arrTempLocBin(iarrCtr) & " Lot >>> " & sTempLot & "<BR>"
	next


		oDOM.Save server.MapPath("../temp/transaction/Pick"&iIssueEntryNo&".xml")
%>
<script type="application/xml" data-itms-xml-island="1" id="PickData" data-src="<%="../temp/transaction/Pick"&iIssueEntryNo&".xml"%>"></script>

<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/mrsIssuePickDetails.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=iClass%>','<%=iItem%>','<%=iEntryNo%>')">

<form method="POST" name="formname" action="">
<input type="hidden" name="hMRSNo" value="<%=iIssueEntryNo%>">
<input type="hidden" name="hMRSDate" value="<%=dMRSDate%>">
<input type="hidden" name="hOrgID" value="<%=sOrgID%>">

<%
	if Trim(iIssueEntryNo) <> "" then
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT MRSITEMSTATUS FROM VWMRSITEMDETAILS WHERE MRSNUMBER = " & iIssueEntryNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing
		if not dcrs2.EOF then
			sStatus = trim(dcrs2(0))
		end if
		dcrs2.Close
		' Status MRS Amended
		if sStatus = "040117" then
%>
		<SCRIPT LANGUAGE=javascript>
			if (confirm("This MR has been amended, so Amend the Issue before Pick Issue.\n Do you want to Amend now?")) {
				window.location.href = "mrsAmendedEntry.asp?mrs=" + document.formname.hMRSNo.value;
			} else {
				window.location.href = "mrsMgmtEntry.asp";
			}
		</SCRIPT>

<%		end if
	end if
%>

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Issue Pick Details
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
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">MR No. Date</td>
                                            <td width="200" class="FieldCellSub">
                                            <%	if Trim(iIssueEntryNo) <> "" then	%>
                                                <span class="DataOnly" id="idMRSNo"><%=sIssueCode%>&nbsp;</span>
                                            <%	else	%>
                                                <span class="DataOnly" id="idMRSNo"><%=sIssueCode%>&nbsp;</span>
                                            <%	end if	%>
                                                <span class="DataOnly"><%=dMRSDate%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Unit Name</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Item Description&nbsp;</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
										<tr>
											<td class="FieldCell">Stores UoM</td>
											<td class="FieldCellSub">
												<span class="DataOnly"><%=sUoM%></span>
											</td>
										</tr>
                                        <tr>
                                            <td class="FieldCell">Quantity W/o Lot Selection</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idWOQty"><%=iQtyWOLot%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity with Lot Selection</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idWQty"><%=iQtyWithLot%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Total Quantity available for Pick</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idQty"><%=iQtyIssue%>&nbsp;</span>
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
												<td class="ExcelHeaderCell" align="center" width="150">Store / Bin</td>
												<td class="ExcelHeaderCell" align="center">Lot Number</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Quantity Marked</td>
												<td class="ExcelHeaderCell" align="center" width="50">Serial Number</td>
												<td class="ExcelHeaderCell" align="center">Quantity Issue</td>
											</tr>
										<%
											iNdCtr = 0

											for iarrCtr = 0 to UBound(arrTempLocBin)
												arrLocBin = split(arrTempLocBin(iarrCtr),"`")
												sStoreCode = arrLocBin(0)
												sBinCode = arrLocBin(1)

												sExp = "//PickDet/PICK [@LOC = "&sStoreCode&" and @BIN = '"&sBinCode&"']"
												Set tempNode = RootNode.Selectnodes(sExp)
												'Response.Write "tempNode="&tempNode.length
												sStoreName = DisplayStore(sStoreCode,sBinCode)
												'Response.Write sExp & " -- " & tempNode.length & "<BR>"


										%>
											<tr>
												<td class="ExcelSerial" align="center" rowspan="<%=tempNode.length%>"><%=iarrCtr+1%></td>
												<td class="ExcelDisplayCell" align="left" rowspan="<%=tempNode.length%>"><%=sStoreName%></td>
										<%
												bLotFlag=false
												bQtyFlag=false
												arrTempLot = split(sTempLot,"|")
												iLotQtyReserved = 0
												iTotLotQty = 0
												iQty = 0

												for iNdCtr =  0 to UBound(arrTempLot)
													sExp = "//PickDet/PICK [@LOC = "&sStoreCode&" and @BIN = '"&sBinCode&"' and @LOTNO = '"&trim(arrTempLot(iNdCtr))&"']"
													Set tempNode = RootNode.Selectnodes(sExp)
													iLot = trim(arrTempLot(iNdCtr))

													if trim(iLot)="" or IsNull(iLot) then iLot = "NULL"
													'Response.Write "<td>iLot = "& iLot &"</td>"
													if iLot = "NULL" then
														iLot = "N/A"
														with dcrs
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
														'	Response.Write DCRS.SOURCE
															.ActiveConnection = con
															.Open
														end with
														set dcrs.ActiveConnection = nothing

														if not dcrs.EOF then
															iQty = cdbl(trim(dcrs(0)))
														end if
														dcrs.close


														with dcrs
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT SERIALNUMBER FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND SERIALNUMBER IS NOT NULL"
															'Response.Write "<td>"& dcrs.source&"</td>"
															.ActiveConnection = con
															.Open
														end with
														set dcrs.ActiveConnection = nothing
														if not dcrs.EOF then
															bSerialFlag = true
														end if
														dcrs.close

                                                        sQuery = "SELECT ISNULL(SUM(QUANTITYFORPick),0) - ISNULL(SUM(QUANTITYPicked),0) FROM VW_INV_IssuedForPick "
                                                        sQuery = sQuery & "WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND "
                                                        sQuery = sQuery & " ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND "
                                                        sQuery = sQuery & "(BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND (LOTNO IS NULL or LOTNO = 'N/A' or LOTNO = '0')"
                                                        dcrs.open sQuery,con
														if not dcrs.EOF then
															iLotQtyReserved = cdbl(dcrs(0))
														end if
														dcrs.close

													    if trim(sReceiptNumbering)="LS" then
    													    sQuery = "Select distinct isNull(sum(LotQuantityNett),0) - isNull(sum(QuantityIssued),0) from INV_T_LocationLot where ItemCode ="& iItem &" and ClassificationCode = "& iClass &" and OrganisationCode ="& Pack(sOrgID)  &" and (StorageLocationNo ="& sStoreCode &" or storageLocationNo is Null) and (StorageBinNumber = "& sBinCode &" or StorageBinNumber is NUll) and isNUll(LotQuantityNett,0) - isNUll(QuantityIssued,0) > 0 "
	    												else
		    											    sQuery = "Select distinct isNull(sum(LotQuantityNett),0) - isNull(sum(QuantityIssued),0) from INV_T_LocationLot where (LotNumber is Null or LotNumber = 'N/A' or LotNumber = '0' or LotNumber = '') and ItemCode ="& iItem &" and ClassificationCode = "& iClass &" and OrganisationCode ="& Pack(sOrgID)  &" and (StorageLocationNo ="& sStoreCode &" or storageLocationNo is Null) and (StorageBinNumber = "& sBinCode &" or StorageBinNumber is NUll) and isNUll(LotQuantityNett,0) - isNUll(QuantityIssued,0) > 0 "
			    										end if
			    										'Response.write "<textarea>"&sQuery&"</textarea>"
												        dcrs.open sQuery,con
														if not dcrs.EOF then
															iLotQty = cdbl(trim(dcrs(0)))
															iTotLotQty = cdbl(iTotLotQty) + cdbl(iLotQty)
														end if
														dcrs.close

													else
													    sQuery = "SELECT ISNULL(SUM(QUANTITYFORPICK),0) - ISNULL(SUM(QUANTITYPICKED),0) FROM VW_INV_IssuedForPick WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND LOTNO = " & Pack(iLot) & ""
													    dcrs.open sQuery,con
														if not dcrs.EOF then
															iLotQtyReserved = cdbl(dcrs(0))
														end if
														dcrs.close


														sQuery = "SELECT DISTINCT ISNULL(AVAILABLENETSTOCK,0),ISNULL(SERIALNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE LOTNUMBER = " & Pack(iLot) & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & "  AND  ORGANISATIONCODE = " & Pack(sOrgID) & "  AND  (STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL)) "
													    dcrs.open sQuery,con
														do while not dcrs.EOF
															iLotQty = cdbl(dcrs(0))
															iSerNo =  cdbl(dcrs(1))
															'Response.Write "iTotLotQty="&iTotLotQty
															iTotLotQty = cdbl(iTotLotQty) + cdbl(iLotQty)
															If iSerNo <> 0 then
																bSerialFlag = True
															Else
																bSerialFlag = false
															End IF
															dcrs.movenext
														loop

														dcrs.close

													end if
													'Response.Write "1="&iLotQty
													'Response.Write "2="&iTotLotQty
													if bLotFlag	then

										%>
											</tr>
											<tr>
										<%
														bLotFlag=true
													end if

													if iLot <> "N/A" then
														'if iLotQty > 0 then
										%>
												<td class="ExcelDisplayCell" align="left" rowspan="<%=tempNode.length%>"><%=iLot%></td>
												<td class="ExcelDisplayCell" width="10" rowspan="<%=tempNode.length%>">
													<input type="text" name="txtStk`<%=sStoreCode%>`<%=sBinCode%>`<%=iLot%>" size="11" value="<%=cdbl(iTotLotQty)%>" maxlength=10 class="FormElemRead" READONLY style="text-align:right">
												</td>

										<%				'end if
													else
														'if iLotQty > 0 then

										%>
												<td class="ExcelDisplayCell" align="left" rowspan="<%=tempNode.length%>"><%=iLot%></td>
												<td class="ExcelDisplayCell" width="10" rowspan="<%=tempNode.length%>">
													<%
													    'Response.Write "iTotLotQty = "& iTotLotQty

														if bSerialFlag then %>
														<input type="text" name="txtStk`<%=sStoreCode%>`<%=sBinCode%>`<%=iLot%>" size="11" value="<%=iTotLotQty%>" maxlength=10 class="FormElemRead" READONLY style="text-align:right">
													<%	else %>
														<input type="text" name="txtStk`<%=sStoreCode%>`<%=sBinCode%>`<%=iLot%>" size="11" value="<%= cdbl(iTotLotQty)%>" maxlength=10 class="FormElemRead" READONLY style="text-align:right">
													<%	end if %>
												</td>

										<%				'end if
													end if

													for iTempCtr=0 to tempNode.length-1
														iSerialCtr = iSerialCtr + 1
														iSerialNo = tempNode.Item(iTempCtr).Attributes.getNamedItem("SNO").Value
														iEntryNo = tempNode.Item(iTempCtr).Attributes.getNamedItem("ITEMENTRYNO").Value
														iIssueEntryNo = tempNode.Item(iTempCtr).Attributes.getNamedItem("ISSUEENTRYNO").Value
														iQtyForIssue = tempNode.Item(iTempCtr).Attributes.getNamedItem("QTYFORISS").Value

														if bQtyFlag	then
										%>
											<tr>
										<%
															bQtyFlag=true
														end if
														if iLot <> "N/A" then
										%>
												<!--<input type="hidden" name="hStock<%=iSerialNo%>" value="<%=cdbl(iLotQty)%>">-->
												<input type="hidden" name="hStock<%=iSerialNo%>" value="<%=cdbl(iTotLotQty)%>">
												<%		else %>
												<input type="hidden" name="hStock<%=iSerialNo%>" value="<%= cdbl(iTotLotQty)%>">
												<%		end if %>
												<td class="ExcelDisplayCell" width="10">
													<input type="text" name="txtQty<%=iSerialNo%>" size="11" value="<%=trim(iQtyForIssue)%>" maxlength=10 class="FormElemRead" READONLY style="text-align:right">
												</td>
												<%'Response.Write bSerialFlag
												if bSerialFlag then %>
														<td class="ExcelFieldCell" align="center">
														<%	if iQtyForIssue = "-" then %>
															<input type="button" name="btn`<%=iEntryNo%>`<%=iIssueEntryNo%>`<%=iLot%>`<%=sStoreCode%>`<%=sBinCode%>`0" value="Pick" class="ActionButtonX" onClick="CheckLot(this,'<%=iSerialNo%>','<%=trim(iQtyWOLot)%>','<%=iItem%>','<%=iClass%>','<%=sAttID%>')">
														<%	else %>
															<input type="button" name="btn`<%=iEntryNo%>`<%=iIssueEntryNo%>`<%=iLot%>`<%=sStoreCode%>`<%=sBinCode%>`0" value="Pick" class="ActionButtonX" onClick="CheckLot(this,'<%=iSerialNo%>','<%=trim(iQtyForIssue)%>','<%=iItem%>','<%=iClass%>','<%=sAttID%>')">
														<%	end if %>
														</td>
														<td class="ExcelInputCell" width="10">
															<input type="text" name="txtIss<%=iSerialNo%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" class="FormElem" READONLY style="text-align:right">
														</td>
												<%	else %>
														<td class="ExcelFieldCell" align="center">
															<input type="button" name="btn" value="Pick" class="ActionButtonX" DISABLED>
														</td>
														<td class="ExcelInputCell" width="10">
															<input type="text" name="txtIss<%=iSerialNo%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" class="FormElem" style="text-align:right">
														</td>
												<%	end if %>
											</tr>
										<%
													next

												next
											next
										%>

										</table>
									</div>
								<input type=hidden name="hiCtr" value="<%=iSerialCtr%>">
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
                                                    <input type="button" value="Issue" name="B1" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date())%>','<%=iItem%>','<%=iClass%>','<%=iEntryNo%>')">
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
			.Source = "SELECT BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " AND BINNUMBER = " & sBin & " ORDER BY BINNUMBER"
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
	' Function to get Item Name
	Function GetItem(iIssueEntryNo)
		' Declaration of variables
		Dim dcrs
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ITEMCODE,GROUPNAME,SHORTDESCRIPTION FROM VWMRSITEMDETAILS WHERE STR(ITEMCODE)+STR(CLASSIFICATIONCODE) IN (SELECT DISTINCT STR(ITEMCODE)+STR(CLASSIFICATIONCODE) FROM INV_T_MRSISSUEPICK WHERE MRSNUMBER = " & iIssueEntryNo & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			Do While Not dcrs.EOF

				Response.Write "<option value="""&trim(dcrs(0))&""">"&trim(trim(dcrs(1)))&"</option>" & vbCrLf

			dcrs.MoveNext
			Loop
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
