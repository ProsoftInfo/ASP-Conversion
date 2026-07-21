<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	itmStatusXMLSelect.asp
	'Module Name				:	Inventory (Stock Management - Status Management)
	'Author Name				:	KUMAR K A
	'Created On					:	06 MAY 2008
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

<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->

<%
	dim dcrs,dcrs1,dcrs2,dcrs3,sSql,OutData,sorgID,Root,newElem,newElem1,newElem2,iItem,iClass
	Dim sLoc,sBin,sBinName,sLocName,sUoM,sUoMDesc,iTotLotQty,sDecimal
	dim sTempMonYr,sMonYr,arrFin,sFinFrom,sFinTo,iStockNo,sLot, iRcptNumbering
	'XML DOM Variables
	Dim oDOM
	dim RootNode,HeaderNode,PageNode,EntryNode
	Dim sStoreCode, sBinCode, iQty, iInvRecNo, iLot, sSrcType, iLotQtyReserved, iSer, iSerialNo
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	if len(Month(date())) = 1 then
		sTempMonYr = "0"&Month(date())
	else
		sTempMonYr = Month(date())
	end if

	sMonYr = sTempMonYr&Year(date())

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
	set dcrs3 = Server.CreateObject("ADODB.Recordset")

	sorgID = Request("orgID")
	iItem = Request("iItem")
	iClass = Request("iClass")
	iRcptNumbering = ""
	
'---------------------------------------------------------------------------------------
	set RootNode = OutData.createElement("PickDet")
	RootNode.setAttribute "CLAS",iClass
	RootNode.setAttribute "ITM",iItem
	RootNode.setAttribute "TOT",""

	OutData.appendChild(RootNode)

	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN'"
		.ActiveConnection = con
		'Response.Write dcrs1.Source 
		.Open
	end with
	set dcrs1.ActiveConnection = nothing

	Do While Not dcrs1.EOF
		sStoreCode = trim(dcrs1(0))
		sBinCode = trim(dcrs1(1))

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
			.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM Inv_T_ItemYearlyStock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
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
			if iInvRecNo = "" then
				'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),ISNULL(IL.LOTNUMBER,'-'),SRCTYPE,IR.INVENTORYRECEIPTNO FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sStoreCode & " AND (IL.STORAGEBINNUMBER = " & sBinCode & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL  AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) GROUP BY IL.LOTNUMBER,SRCTYPE,IR.INVENTORYRECEIPTNO ORDER BY 2"
				'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),ISNULL(IL.LOTNUMBER,'-'),SRCTYPE,IR.INVENTORYRECEIPTNO FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & "  AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sStoreCode & " AND (IL.STORAGEBINNUMBER = " & sBinCode & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL GROUP BY IL.LOTNUMBER,SRCTYPE,IR.INVENTORYRECEIPTNO ORDER BY 2"
				.Source = " SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),ISNULL(LOTNUMBER,'-'),SRCTYPE,INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND (STORAGELOCATIONNO = 2 AND (STORAGEBINNUMBER = " & sStoreCode & " OR STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0  AND SERIALNUMBER IS NOT NULL GROUP BY LOTNUMBER,SRCTYPE,INVENTORYRECEIPTNO ORDER BY 2 "
			else
				'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),ISNULL(IL.LOTNUMBER,'-'),SRCTYPE,IR.INVENTORYRECEIPTNO FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sStoreCode & " AND (IL.STORAGEBINNUMBER = " & sBinCode & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL  AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) AND IR.INVENTORYRECEIPTNO = " & iInvRecNo & " AND IL.INVENTORYRECEIPTNO = " & iInvRecNo & " GROUP BY IL.LOTNUMBER,SRCTYPE,IR.INVENTORYRECEIPTNO ORDER BY 2"
				'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),ISNULL(IL.LOTNUMBER,'-'),SRCTYPE,IR.INVENTORYRECEIPTNO FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sStoreCode & " AND (IL.STORAGEBINNUMBER = " & sBinCode & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL AND IR.INVENTORYRECEIPTNO = " & iInvRecNo & " AND IL.INVENTORYRECEIPTNO = " & iInvRecNo & " GROUP BY IL.LOTNUMBER,SRCTYPE,IR.INVENTORYRECEIPTNO ORDER BY 2"
				.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),ISNULL(LOTNUMBER,'-'),SRCTYPE,INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItem & " AND  CLASSIFICATIONCODE = " & iClass & " AND (STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL AND INVENTORYRECEIPTNO = " & iInvRecNo & " GROUP BY LOTNUMBER,SRCTYPE,INVENTORYRECEIPTNO ORDER BY 2"
			end if
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.Source 
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			Do While Not dcrs.EOF
				iLot = trim(dcrs(1))
				sSrcType = trim(dcrs(2))
				iLotQtyReserved = 0
				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					if iLot = "-" then
						.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0) - ISNULL(SUM(QUANTITYISSUED),0) FROM VWPICKDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNO = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND LOTNUMBER IS NULL"
					else
						.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0) - ISNULL(SUM(QUANTITYISSUED),0) FROM VWPICKDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNO = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND LOTNUMBER = " & Pack(iLot) & ""
					end if
					.ActiveConnection = con
					.Open
				end with
				set dcrs2.ActiveConnection = nothing

				if not dcrs2.EOF then
					iLotQtyReserved = cdbl(dcrs2(0))
				end if
				dcrs2.close

				if (cdbl(dcrs(0)) - cdbl(iLotQtyReserved)) > 0 then

					if iLot = "-" then iLot = "NULL"

					if iLot <> "NULL" then
						with dcrs2
							.CursorLocation = 3
							.CursorType = 3
							'.Source = "SELECT SERIALNUMBER FROM INV_T_RECEIPTLOTDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL"
							.Source = "SELECT SERIALNUMBER FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL"
							.ActiveConnection = con
							.Open
						end with
						set dcrs2.ActiveConnection = nothing
						if not dcrs2.EOF then
							iSer = trim(dcrs2(0))
						end if
						dcrs2.close
					'elseif iLot = "NULL" then
					'	iSer = "NULL"
					end if
				'	Response.Write dcrs2.Source 
				'	if not iSer = "NULL" then
						iTotLotQty = cdbl(iTotLotQty) + (cdbl(dcrs(0)) - cdbl(iLotQtyReserved))

						iSerialNo = iSerialNo + 1

						set HeaderNode = oDOM.createElement("PICK")
						HeaderNode.setAttribute "LOC",trim(sStoreCode)
						HeaderNode.setAttribute "BIN",trim(sBinCode)
						HeaderNode.setAttribute "LOTNO",iLot
						HeaderNode.setAttribute "ISSQTY",""
						HeaderNode.setAttribute "STOCK",(cdbl(dcrs(0)) - cdbl(iLotQtyReserved))
						'HeaderNode.setAttribute "STOCK",cdbl(trim(dcrs(0)))
						HeaderNode.setAttribute "SNO",iSerialNo
						HeaderNode.setAttribute "SERIALNO",iSer
						HeaderNode.setAttribute "SRCTYPE",sSrcType
						HeaderNode.setAttribute "INVRECNO",trim(dcrs(3))

						RootNode.appendChild(HeaderNode)
				'	end if
				end if
			dcrs.MoveNext
			Loop
		else
			if iInvRecNo = "" then
				iTotLotQty = 0

				iLotQtyReserved = 0
				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0) - ISNULL(SUM(QUANTITYISSUED),0) FROM VWPICKDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNO = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND LOTNUMBER IS NULL"
					.ActiveConnection = con
					.Open
				end with
				set dcrs2.ActiveConnection = nothing

				if not dcrs2.EOF then
					iLotQtyReserved = cdbl(dcrs2(0))
				end if
				dcrs2.close

				with dcrs3
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
					.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM Inv_T_ItemYearlyStock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				if Not dcrs3.EOF then
					if not (cdbl(trim(dcrs3(0))) - cdbl(iTotLotQty)) <= 0 then
						iLotQtyReserved = 0
						with dcrs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0) - ISNULL(SUM(QUANTITYISSUED),0) FROM VWPICKDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNO = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND LOTNUMBER IS NULL"
							.ActiveConnection = con
							.Open
						end with
						set dcrs2.ActiveConnection = nothing

						if not dcrs2.EOF then
							iLotQtyReserved = cdbl(dcrs2(0))
						end if
						dcrs2.close

						iSerialNo = iSerialNo + 1
						set HeaderNode = oDOM.createElement("PICK")
						HeaderNode.setAttribute "LOC",trim(sStoreCode)
						HeaderNode.setAttribute "BIN",trim(sBinCode)
						HeaderNode.setAttribute "LOTNO","NULL"
						HeaderNode.setAttribute "ISSQTY",""
						HeaderNode.setAttribute "STOCK",((cdbl(trim(dcrs3(0))) - cdbl(iLotQtyReserved)) - cdbl(iTotLotQty))
						'HeaderNode.setAttribute "STOCK",(iQty - iTotLotQty)
						HeaderNode.setAttribute "SNO",iSerialNo
						HeaderNode.setAttribute "SERIALNO","NULL"
						HeaderNode.setAttribute "SRCTYPE","NULL"
						HeaderNode.setAttribute "INVRECNO","NULL"

						RootNode.appendChild(HeaderNode)
					end if
				end if
				dcrs3.Close
			end if
		end if
		dcrs.Close
	dcrs1.MoveNext
	loop
	dcrs1.Close

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
'------------------------------------------------------------------------------------------------

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
