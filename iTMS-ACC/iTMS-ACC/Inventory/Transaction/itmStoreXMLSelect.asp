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
	'Program Name				:	itmStoreXMLSelect.asp
	'Module Name				:	Inventory (Stock Management - Stock Transfer)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 27, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 21, 2010
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
	dim dcrs,dcrs1,dcrs2,sSql,OutData,sorgID,Root,newElem,newElem1,iItem,iClass,ItemDetails
	Dim sLoc,sBin,sBinName,sLocName,sUoM,sUoMDesc,iTotLotQty,iStockQty,sDecimal
	dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iLotQtyReserved,sLot,iSer,ndRoot,ndItem,ItemNode,sItemName
	dim sTemp,sReceiptNumberStatus,iRate,sAttributeList,sFinPeriod
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	Set ItemDetails = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
	
	ItemDetails.load(Server.MapPath("../temp/transaction/Inv_ItemDetails_"&Session.SessionID&".xml"))
	set ndRoot = ItemDetails.documentElement
	
	sFinPeriod = Session("FinPeriod")
	arrFin = split(sFinPeriod,":")
	sFinFrom = "01/04/"&arrFin(0)
	sFinTo = "31/03/"&arrFin(1)
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	if ndRoot.hasChildNodes() then
		For each ndItem in ndRoot.childNodes
			sorgID = ndItem.getAttribute("Unit")
			iItem = ndItem.getAttribute("ICode")
			iClass = ndItem.getAttribute("CCode")
			
			sSql = "Select ItemDescription,ReceiptNumbering from VW_INV_ITEMS where ItemCode = "& iItem &" and ClassificationCode = "& iClass 
			dcrs.Open sSql,con
			if not dcrs.EOF then
				sItemName = dcrs(0)
				sReceiptNumberStatus = trim(dcrs(1))
			end if
			dcrs.Close
			
			set ItemNode  = OutData.createElement("Item")
				ItemNode.setAttribute "ICode",iItem
				ItemNode.setAttribute "CCode",iClass
				ItemNode.setAttribute "Unit",sorgID
				ItemNode.setAttribute "IName",sItemName
				ItemNode.setAttribute "RecNumStatus",sReceiptNumberStatus
			Root.appendChild ItemNode
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0),LOCATIONNAME,LOCATIONCODE,YearClosingValue,YearClosingStock FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
				'Response.Write dcrs.Source 
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			Do While Not dcrs.EOF
				iStockQty = 0
				sLoc = trim(dcrs(0))
				sBin = trim(dcrs(1))
				sLocName = trim(dcrs(2))
				if cdbl(dcrs(5))>0 then
				    iRate = cdbl(dcrs(4))/cdbl(dcrs(5))
				else
				    iRate = 0
				end if

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " AND BINNUMBER = " & sBin & " ORDER BY BINNUMBER"
					'Response.Write "<p>Sql="&dcrs1.Source 
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					Set newElem = OutData.createElement("LOCDET")
					newElem.setAttribute "LOC", trim(sLoc)
					newElem.setAttribute "BIN", trim(dcrs1(0))
					newElem.setAttribute "LOCNAME", trim(sLocName)
					newElem.setAttribute "BINNAME", trim(dcrs1(1))
					newElem.setAttribute "AVGRATE",iRate
				else
					Set newElem = OutData.createElement("LOCDET")
					newElem.setAttribute "LOC", trim(sLoc)
					newElem.setAttribute "BIN", "0"
					newElem.setAttribute "LOCNAME", trim(sLocName)
					newElem.setAttribute "BINNAME", "-"
					newElem.setAttribute "AVGRATE",iRate
				end if
				dcrs1.Close
				sAttributeList = ""

				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,ISNULL(IL.LOTNUMBER,'-') FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = '" & sorgID & "' AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sLoc & " AND (IL.STORAGEBINNUMBER = " & sBin & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) GROUP BY IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,IL.LOTNUMBER ORDER BY 2,3"
					'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,ISNULL(IL.LOTNUMBER,0) FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sLoc & " AND (IL.STORAGEBINNUMBER = " & sBin & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) AND SERIALNUMBER IS NOT NULL GROUP BY IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,IL.LOTNUMBER ORDER BY 2,3"
					'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),STORAGELOCATIONNO,STORAGEBINNUMBER,INVENTORYRECEIPTNO,ISNULL(LOTNUMBER,0) FROM INV_T_LOCATIONLOT  WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND (STORAGELOCATIONNO = "& sLoc &" AND (STORAGEBINNUMBER = "& sBin &" OR STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 GROUP BY STORAGELOCATIONNO,STORAGEBINNUMBER,INVENTORYRECEIPTNO,LOTNUMBER ORDER BY 2,3"
					.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),STORAGELOCATIONNO,STORAGEBINNUMBER,ISNULL(LOTNUMBER,0),IsNull(AttributeList,'-') FROM INV_T_LOCATIONLOT  WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND (STORAGELOCATIONNO = "& sLoc &" AND (STORAGEBINNUMBER = "& sBin &" OR STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 GROUP BY STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,AttributeList ORDER BY 2,3"
					'Response.Write dcrs2.Source
					.ActiveConnection = con
					.Open
				end with
				set dcrs2.ActiveConnection = nothing

				if not dcrs2.EOF then
					Do While Not dcrs2.EOF
						iLotQtyReserved = "0"
						'sLot = trim(dcrs2(4))
						sLot = trim(dcrs2(3))
						if sLot = "0" then sLot = "NULL"
						sAttributeList= trim(dcrs2(4))

						if sLot <> "NULL" then
							with dcrs1
								.CursorLocation = 3
								.CursorType = 3
								'.Source = "SELECT SERIALNUMBER FROM INV_T_RECEIPTLOTDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOTNUMBER = " & Pack(sLot) & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL"
								.Source = "SELECT SERIALNUMBER FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND (LOTNUMBER = " & Pack(sLot) & " OR LOTNUMBER = '0' OR LOTNUMBER = 'N/A') AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER IS NOT NULL"
								'Response.write dcrs1.source
								.ActiveConnection = con
								.Open
							end with
							set dcrs1.ActiveConnection = nothing
							if not dcrs1.EOF then
								iSer = trim(dcrs1(0))
							end if
							dcrs1.close
						elseif sLot = "NULL" then
							'iSer = "NULL"
						end if

						with dcrs1
							.CursorLocation = 3
							.CursorType = 3
							if sLot = "NULL" then
								'.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0) - ISNULL(SUM(QUANTITYISSUED),0) FROM VWPICKDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNO = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND LOTNUMBER IS NULL"
								.Source = "Select IsNull(Sum(QuantityForPick),0)-IsNull(Sum(QuantityPicked),0) from VW_INV_IssuedForPick WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNumber = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND LotNo IS NULL"
							else
								'.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0) - ISNULL(SUM(QUANTITYISSUED),0) FROM VWPICKDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNO = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND LOTNUMBER = " & Pack(sLot) & ""
								.Source = "Select IsNull(Sum(QuantityForPick),0)-IsNull(Sum(QuantityPicked),0) from VW_INV_IssuedForPick WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNumber = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND LotNO = "& Pack(sLot) &""
							end if
							.ActiveConnection = con
							.Open
						end with
						set dcrs1.ActiveConnection = nothing

						if not dcrs1.EOF then
							iLotQtyReserved = cdbl(dcrs1(0))
						end if
						dcrs1.close

						'if not iSer = "NULL" then

							iTotLotQty = cdbl(iTotLotQty) + cdbl(trim(dcrs2(0)))
							if (cdbl(dcrs2(0)) - cdbl(iLotQtyReserved)) > 0 then
								Set newElem1 = OutData.createElement("PICK")
								if sLot = "-" then
									newElem1.setAttribute "LOTNO", "NULL"
								else
									newElem1.setAttribute "LOTNO", sLot
								end if

								newElem1.setAttribute "INVRECNO",""' trim(dcrs2(3))
								newElem1.setAttribute "QTYSTK", (cdbl(dcrs2(0)) - cdbl(iLotQtyReserved))
								newElem1.setAttribute "QTYISS", ""
								newElem1.setAttribute "TYPE", ""
								newElem1.setAttribute "ENABLE", "Y"
								if trim(sAttributeList)="-" then
								    newElem1.setAttribute "ATTLIST", ""
								else
								    newElem1.setAttribute "ATTLIST", sAttributeList
								end if
								newElem.appendChild newElem1
								iStockQty = iStockQty + (cdbl(dcrs2(0)) - cdbl(iLotQtyReserved))
							end if
						'end if
					dcrs2.MoveNext
					Loop
				else
					
					iTotLotQty = "0"
					iLotQtyReserved = "0"
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						'.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0) - ISNULL(SUM(QUANTITYISSUED),0) FROM VWPICKDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNO = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND LOTNUMBER IS NULL"
						.Source = "Select IsNull(Sum(QuantityForPick),0)-IsNull(Sum(QuantityPicked),0) from VW_INV_IssuedForPick WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LocationNumber = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND LotNO IS NULL"
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing

					if not dcrs1.EOF then
						iLotQtyReserved = cdbl(dcrs1(0))
					end if
					dcrs1.close

					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
						.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & "  AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
						'Response.Write "<p>sql="&dcrs1.Source 
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing
					
					If Not dcrs1.EOF Then	
						do while Not dcrs1.EOF
							if not (cdbl(trim(dcrs1(0))) - cdbl(iTotLotQty)) <= 0 then
								Set newElem1 = OutData.createElement("PICK")
								newElem1.setAttribute "LOTNO", "NULL"
								newElem1.setAttribute "INVRECNO", "NULL"
								newElem1.setAttribute "QTYSTK",(cdbl(dcrs1(0)) - cdbl(iLotQtyReserved))
								newElem1.setAttribute "QTYISS", ""
								newElem1.setAttribute "TYPE", ""
								newElem1.setAttribute "ENABLE", "Y"
								if trim(sAttributeList)="-" then
								    newElem1.setAttribute "ATTLIST", ""
								else
								    newElem1.setAttribute "ATTLIST", sAttributeList
								end if
								newElem.appendChild newElem1
								iStockQty = iStockQty + (cdbl(dcrs1(0)) - cdbl(iLotQtyReserved))
							Else
								'Added Newly By UmaMaheswari S,on MAY 27
								Set newElem1 = OutData.createElement("PICK")
								newElem1.setAttribute "LOTNO", "NULL"
								newElem1.setAttribute "INVRECNO", "NULL"
								newElem1.setAttribute "QTYSTK","0"
								newElem1.setAttribute "QTYISS", ""
								newElem1.setAttribute "TYPE", ""
								newElem1.setAttribute "ENABLE", "Y"
								if trim(sAttributeList)="-" then
								    newElem1.setAttribute "ATTLIST", ""
								else
								    newElem1.setAttribute "ATTLIST", sAttributeList
								end if
								newElem.appendChild newElem1
							end if
						dcrs1.MoveNext
						Loop
					Else
						'Added Newly By UmaMaheswari S,on MAY 27
						Set newElem1 = OutData.createElement("PICK")
						newElem1.setAttribute "LOTNO", "NULL"
						newElem1.setAttribute "INVRECNO", "NULL"
						newElem1.setAttribute "QTYSTK","0"
						newElem1.setAttribute "QTYISS", ""
						newElem1.setAttribute "TYPE", ""
						newElem1.setAttribute "ENABLE", "Y"
						if trim(sAttributeList)="-" then
								    newElem1.setAttribute "ATTLIST", ""
								else
								    newElem1.setAttribute "ATTLIST", sAttributeList
								end if
						newElem.appendChild newElem1
					End IF
					dcrs1.close
				end if
				dcrs2.Close

				newElem.setAttribute "STOCK", iStockQty
				ItemNode.appendChild newElem

			dcrs.MoveNext
			Loop
			dcrs.Close

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION,DECIMALALLOWED FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "')"
				.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION,DECIMALALLOWED FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "')"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
	
			set sUoM = dcrs(0)
			set sUoMDesc = dcrs(1)
			set sDecimal = dcrs(2)

			if Not dcrs.EOF then
				Set newElem = OutData.createElement("UOM")
				newElem.setAttribute "UoMCode", trim(sUoM)
				newElem.setAttribute "UoMName", trim(sUoMDesc)
				newElem.setAttribute "UoMDecimal", trim(sDecimal)
				ItemNode.appendChild newElem
			end if
			dcrs.Close
		Next
	end if 

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
