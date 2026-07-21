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
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 03, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 21,2010
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
	dim dcrs,dcrs1,dcrs2,dcrs3,sSql,OutData,sorgID,Root,newElem,newElem1,newElem2,iItem,iClass,ItemData
	Dim sLoc,sBin,sBinName,sLocName,sUoM,sUoMDesc,iTotLotQty,sDecimal,ndRootItem,ndChildItem
	dim sTempMonYr,sMonYr,arrFin,sFinFrom,sFinTo,iStockNo,sLot, iRcptNumbering,ndItem,sItemName
	dim sTemp,iValue,sFinPeriod
	
	sFinPeriod = Session("FinPeriod")
	arrFin = split(sFinPeriod,":")
	sFinFrom ="01/04/"&arrFin(0)
	sFinTo = "31/03/"&arrFin(1)

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	Set ItemData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
	set dcrs3 = Server.CreateObject("ADODB.Recordset")
	
	Set Root = OutData.createElement("Root")
	OutData.appendChild Root
		
	ItemData.load(Server.MapPath("../temp/transaction/Inv_ItemDetails_"& Session.SessionID &".xml"))
	'Response.Write ItemData.xml
	
	set ndRootItem = ItemData.documentElement
	if ndRootItem.hasChildNodes then
		for each ndChildItem in ndRootItem.childNodes
				sorgID = ndChildItem.getAttribute("Unit")
				iItem = ndChildItem.getAttribute("ICode")
				iClass = ndChildItem.getAttribute("CCode")
				iRcptNumbering = ""
				
				sSql = "Select ItemDescription from VW_INV_ITEMS where ItemCode = "& iItem &" and ClassificationCode = "& iClass
				dcrs.Open sSql,con
				if not dcrs.EOF then
					sItemName  = dcrs(0)
				end if
				dcrs.Close 
				
				set ndItem = OutData.createElement("Item")
				ndItem.setAttribute "ICode",iItem
				ndItem.setAttribute "CCode",iClass
				ndItem.setAttribute "Unit",sorgID
				ndItem.setAttribute "IName",sItemName
				Root.appendChild ndItem
				
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0),LOCATIONNAME,LOCATIONCODE,YearClosingValue FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
					'Response.Write vbCrLf + dcrs.Source + vbCrLf 
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				set sLoc = dcrs(0)
				set sBin = dcrs(1)
				set sLocName = dcrs(2)
				Set iValue = dcrs(4)

				Do While Not dcrs.EOF
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY BINNUMBER"
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing

					with dcrs2
						.CursorLocation = 3
						.CursorType = 3
						'.Source = "Select ReceiptNumbering From inv_M_ItemOrgPurchase where ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sorgID & "'"
						.Source = "Select ReceiptNumbering From inv_M_ItemMaster where ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sorgID & "'"  
						.ActiveConnection = con
						.Open
					end with
					set dcrs2.ActiveConnection = nothing
					If Not dcrs2.EOF Then
						iRcptNumbering = dcrs2(0)
					End If
					dcrs2.Close 

					if not dcrs1.EOF then
						Set newElem = OutData.createElement("LOCDET")
						newElem.setAttribute "LOC", trim(sLoc)
						newElem.setAttribute "BIN", trim(dcrs1(0))
						newElem.setAttribute "LOCNAME", trim(sLocName)
						newElem.setAttribute "BINNAME", trim(dcrs1(1))
						newElem.setAttribute "RECNUM", iRcptNumbering
						newElem.setAttribute "ITEMCODE", iItem
						newElem.setAttribute "VALUE", iValue
						ndItem.appendChild newElem
					else
						Set newElem = OutData.createElement("LOCDET")
						newElem.setAttribute "LOC", trim(sLoc)
						newElem.setAttribute "BIN", "0"
						newElem.setAttribute "LOCNAME", trim(sLocName)
						newElem.setAttribute "BINNAME", "-"
						newElem.setAttribute "RECNUM", iRcptNumbering 
						newElem.setAttribute "ITEMCODE", iItem			
						newElem.setAttribute "VALUE", iValue
						ndItem.appendChild newElem
					end if
					dcrs1.Close

					with dcrs2
						.CursorLocation = 3
						.CursorType = 3
						'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,ISNULL(IL.LOTNUMBER,'-') FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sorgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sLoc & " AND (IL.STORAGEBINNUMBER = " & sBin & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) GROUP BY IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,IL.LOTNUMBER ORDER BY 2,3"
						'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,ISNULL(IL.LOTNUMBER,0) FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sLoc & " AND (IL.STORAGEBINNUMBER = " & sBin & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) AND SERIALNUMBER IS NOT NULL GROUP BY IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,IL.LOTNUMBER ORDER BY 2,3"
						.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),STORAGELOCATIONNO,STORAGEBINNUMBER,INVENTORYRECEIPTNO,ISNULL(LOTNUMBER,0) FROM INV_T_LOCATIONLOT  WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND (STORAGELOCATIONNO = "& sLoc &" AND (STORAGEBINNUMBER = "& sBin &" OR STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 GROUP BY STORAGELOCATIONNO,STORAGEBINNUMBER,INVENTORYRECEIPTNO,LOTNUMBER ORDER BY 2,3"
						'Response.Write dcrs2.Source & vbCrLf
						.ActiveConnection = con
						.Open
					end with
					set dcrs2.ActiveConnection = nothing

					if not dcrs2.EOF then
						Do While Not dcrs2.EOF
							iTotLotQty = cdbl(iTotLotQty) + cdbl(trim(dcrs2(0)))
							Set newElem1 = OutData.createElement("PICK")

							sLot = trim(dcrs2(4))
							if sLot = "0" then sLot = "NULL"

							newElem1.setAttribute "LOTNO", sLot
							newElem1.setAttribute "INVRECNO", trim(dcrs2(3))
							newElem1.setAttribute "QTYSTK", cdbl(trim(dcrs2(0)))
							newElem.appendChild newElem1
							
							iStockNo = "0"
							
							if not iStockNo = "0" then
							'if iStockNo = "0" then
								with dcrs3
									.CursorLocation = 3
									.CursorType = 3
									if sLot = "NULL" or sLot = "0" then
										'.Source = "SELECT ISNULL(SUM(CLEAN),0),ISNULL(SUM(RESERVED),0),ISNULL(SUM(ONHOLD),0),ISNULL(SUM(REJECTED),0) FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL AND SERIALNO IN (SELECT SERIALNO FROM INV_T_MRSISSUEPICKSERIAL WHERE SERIALNO IS NOT NULL AND QUANTITYFORISSUE - QUANTITYISSUED > 0)"
										.Source = "Select 0,IsNull(Sum(reserved),0),IsNull(sum(Onhold),0),Isnull(sum(Rejected),0) from INV_T_LocationLot Where (LotNumber is Null or LotNumber = 'N/A' or LotNumber = 0) and LotQuantityNett - QuantityIssued > 0 and ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &""
									else
										'.Source = "SELECT ISNULL(SUM(CLEAN),0),ISNULL(SUM(RESERVED),0),ISNULL(SUM(ONHOLD),0),ISNULL(SUM(REJECTED),0) FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND LOTNO = " & Pack(sLot) & ""
										.Source = "Select 0,IsNull(Sum(reserved),0),IsNull(sum(Onhold),0),Isnull(sum(Rejected),0) from INV_T_LocationLot Where LotNumber = " & Pack(sLot) & " and LotQuantityNett - QuantityIssued > 0 and ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" "
									end if
									.ActiveConnection = con
									.Open
								end with
								'Response.Write dcrs3.Source 
								set dcrs3.ActiveConnection = nothing

								if not dcrs3.EOF then
									Set newElem2 = OutData.createElement("STATUS")
									newElem2.setAttribute "STOCKNO", iStockNo
									newElem2.setAttribute "LOTNO", sLot
									newElem2.setAttribute "INVRECNO", trim(dcrs2(3))
									newElem2.setAttribute "QTYCLE", cdbl(dcrs3(0))
									newElem2.setAttribute "QTYRES", cdbl(dcrs3(1))
									newElem2.setAttribute "QTYOHO", cdbl(dcrs3(2))
									newElem2.setAttribute "QTYREJ", cdbl(dcrs3(3))
									newElem2.setAttribute "A", "1"
									newElem1.appendChild newElem2
								else
									Set newElem2 = OutData.createElement("STATUS")
									newElem2.setAttribute "STOCKNO", iStockNo
									newElem2.setAttribute "LOTNO", sLot
									newElem2.setAttribute "INVRECNO", trim(dcrs2(3))
									newElem2.setAttribute "QTYCLE", "0"
									newElem2.setAttribute "QTYRES", "0"
									newElem2.setAttribute "QTYOHO", "0"
									newElem2.setAttribute "QTYREJ", "0"
									newElem2.setAttribute "B", "2"
									newElem1.appendChild newElem2
								end if
								dcrs3.Close
							else
								Set newElem2 = OutData.createElement("STATUS")
								newElem2.setAttribute "STOCKNO", iStockNo
								newElem2.setAttribute "LOTNO", sLot
								newElem2.setAttribute "INVRECNO", trim(dcrs2(3))
								newElem2.setAttribute "QTYCLE", "0"
								newElem2.setAttribute "QTYRES", "0"
								newElem2.setAttribute "QTYOHO", "0"
								newElem2.setAttribute "QTYREJ", "0"
								newElem2.setAttribute "C", "3"
								newElem1.appendChild newElem2
							end if

						dcrs2.MoveNext
						Loop
					else
						iTotLotQty = "0"

						with dcrs1
							.CursorLocation = 3
							.CursorType = 3
							'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "' AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
							.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "' AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
							.ActiveConnection = con
							.Open
						end with
						set dcrs1.ActiveConnection = nothing

						do while Not dcrs1.EOF
							if not (cdbl(trim(dcrs1(0))) - cdbl(iTotLotQty)) <= 0 then
								Set newElem1 = OutData.createElement("PICK")
								newElem1.setAttribute "LOTNO", "NULL"
								newElem1.setAttribute "INVRECNO", "NULL"
								newElem1.setAttribute "QTYSTK",(cdbl(trim(dcrs1(0))) - cdbl(iTotLotQty))
								newElem.appendChild newElem1
								
								iStockNo = "0"

								if not iStockNo = "0" then
									with dcrs3
										.CursorLocation = 3
										.CursorType = 3
										'.Source = "SELECT ISNULL(SUM(CLEAN),0),ISNULL(SUM(RESERVED),0),ISNULL(SUM(ONHOLD),0),ISNULL(SUM(REJECTED),0) FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL"
										.Source = "Select 0,IsNull(Sum(reserved),0),IsNull(sum(Onhold),0),Isnull(sum(Rejected),0) from INV_T_LocationLot Where (LotNumber is Null or LotNumber = 'N/A' or LotNumber = 0) and ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '" & sorgID & "'"
										.ActiveConnection = con
										.Open
									end with
									set dcrs3.ActiveConnection = nothing

									if not dcrs3.EOF then
										Set newElem2 = OutData.createElement("STATUS")
										newElem2.setAttribute "STOCKNO", iStockNo
										newElem2.setAttribute "LOTNO", "NULL"
										newElem2.setAttribute "INVRECNO", "NULL"
										newElem2.setAttribute "QTYCLE", cdbl(dcrs3(0))
										newElem2.setAttribute "QTYRES", cdbl(dcrs3(1))
										newElem2.setAttribute "QTYOHO", cdbl(dcrs3(2))
										newElem2.setAttribute "QTYREJ", cdbl(dcrs3(3))
										newElem1.appendChild newElem2
									else
										Set newElem2 = OutData.createElement("STATUS")
										newElem2.setAttribute "STOCKNO", iStockNo
										newElem2.setAttribute "LOTNO", "NULL"
										newElem2.setAttribute "INVRECNO", "NULL"
										newElem2.setAttribute "QTYCLE", "0"
										newElem2.setAttribute "QTYRES", "0"
										newElem2.setAttribute "QTYOHO", "0"
										newElem2.setAttribute "QTYREJ", "0"
										newElem1.appendChild newElem2
									end if
									dcrs3.Close
								else
									Set newElem2 = OutData.createElement("STATUS")
									newElem2.setAttribute "STOCKNO", iStockNo
									newElem2.setAttribute "LOTNO", "NULL"
									newElem2.setAttribute "INVRECNO", "NULL"
									newElem2.setAttribute "QTYCLE", "0"
									newElem2.setAttribute "QTYRES", "0"
									newElem2.setAttribute "QTYOHO", "0"
									newElem2.setAttribute "QTYREJ", "0"
									newElem1.appendChild newElem2
								end if
							end if
						dcrs1.MoveNext
						Loop
						dcrs1.close
					end if
					dcrs2.Close

				dcrs.MoveNext
				Loop
				dcrs.Close

				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION,DECIMALALLOWED FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "')"
					.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION,DECIMALALLOWED FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM VW_INV_ITEMS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "')"
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
					ndItem.appendChild newElem
				end if
				dcrs.Close
		next 'for each ndChildItem in ndRootItem.childNodes
	end if 'if ndRootItem.hasChilNodes() then

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
