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
	'Program Name				:	ArrangeBinDetXML.asp
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
	dim sTemp,sReceiptNumberStatus,sBinStatus,sItemTypeID
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	Set ItemDetails = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
	
	ItemDetails.load(Server.MapPath("../temp/transaction/Inv_ItemDetails_"&Session.SessionID&".xml"))
	set ndRoot = ItemDetails.documentElement
	
	if len(Month(date())) = 1 then
		sTempMonYr = "0"&Month(date())
	else
		sTempMonYr = Month(date())
	end if
	sMonYr = sTempMonYr&Year(date())

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	if ndRoot.hasChildNodes() then
		For each ndItem in ndRoot.childNodes
			sorgID = ndItem.getAttribute("Unit")
			iItem = ndItem.getAttribute("ICode")
			iClass = ndItem.getAttribute("CCode")
			sItemTypeID = ndItem.getAttribute("ItemTypeID")
			
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
				'.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0),LOCATIONNAME,LOCATIONCODE FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1"
				.Source = "SELECT DISTINCT LOCATIONNUMBER,0,LOCATIONNAME,LOCATIONCODE FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			Do While Not dcrs.EOF
				iStockQty = 0
				sLoc = trim(dcrs(0))
				sBin = trim(dcrs(1))
				sLocName = trim(dcrs(2))
				sBinStatus = "N"

				Set newElem = OutData.createElement("LOCDET")
				newElem.setAttribute "LOC", trim(sLoc)
				newElem.setAttribute "LOCNAME", trim(sLocName)
				
				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "select StorageBinNumber from INV_T_LOCATIONLOT WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND STORAGELOCATIONNO = "& sLoc &" AND StorageBinNumber is NOT NULL AND StorageBinNumber <> 0"
					.Source = "select BinNumber from Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = "& sLoc &" "
					.ActiveConnection = con
					.Open
				end with
				set dcrs2.ActiveConnection = nothing
				If Not dcrs2.EOF Then
					sBinStatus = "Y"
				End IF
				dcrs2.Close 
				
				If sBinStatus = "Y" Then
					with dcrs2
						.CursorLocation = 3
						.CursorType = 3
						'.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),STORAGELOCATIONNO,isNull(STORAGEBINNUMBER,0) FROM INV_T_LOCATIONLOT  WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND (STORAGELOCATIONNO = "& sLoc &") AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND Year(DateofReceipt) >= Year('"&sFinFrom&"') AND Year(DateofReceipt) <= '"&Year(sFinTo)&"' GROUP BY STORAGELOCATIONNO,STORAGEBINNUMBER ORDER BY 2,3"
						.source = "SELECT DISTINCT (ISNULL(SUM(YEAROPENINGSTOCK),0)+ISNULL(SUM(YEARRECEIPTQUANTITY),0) - ISNULL(SUM(YEARISSUEQUANTITY),0)),isNull(BINNUMBER,0) FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND LOCATIONNUMBER = "& sLoc &" AND CONVERT(Datetime,FinancialYearFrom,103) >= CONVERT(Datetime,'"& sFinFrom &"',103) AND CONVERT(Datetime,FinancialYearTo,103) <= CONVERT(Datetime,'"& sFinTo &"',103) GROUP BY BINNUMBER ORDER BY 2"
						'Response.Write dcrs2.Source 
						.ActiveConnection = con
						.Open
					end with
					set dcrs2.ActiveConnection = nothing
					If Not dcrs2.EOF Then
						Do while Not dcrs2.EOF 
							Set newElem1  = OutData.createElement("BINDET")
							newElem1.setAttribute "BINNO", trim(dcrs2(1))'trim(dcrs2(2))
							newElem1.setAttribute "QTY", trim(dcrs2(0))
							newElem.appendchild newElem1 
							dcrs2.MoveNext 
						Loop
					End IF
					dcrs2.Close 
					
				End IF
				
				'Response.Write "<p>sItemTypeID="&sItemTypeID
				sSql = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
						" where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) "&_
						" FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMTYPEID = '"& sItemTypeID &"' AND VI.ORGANISATIONCODE = '"& sorgID &"' "&_
						" AND VI.PURCHASEELIGIBLE = 1 AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 "&_
						" AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND "&_ 
						" VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND "&_
						" CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"& sFinFrom &"',103) "&_
						" AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"& sFinTo &"',103)"&_
						" AND VI.ITEMCODE = "& iItem &" AND VI.CLASSIFICATIONCODE = "& iClass &" "
				dcrs1.Open sSql,con
				If Not dcrs1.EOF Then
					iStockQty = dcrs1(3)
				End IF
				dcrs1.Close 

				newElem.setAttribute "STOCK", iStockQty
				newElem.setAttribute "BINSTATUS", sBinStatus 
				ItemNode.appendChild newElem
				
				'Store Bin Det
				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT DISTINCT BINCODE,BINNUMBER FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER="& sLoc &" Order by BINNUMBER"
					.ActiveConnection = con
					.Open
				end with
				set dcrs2.ActiveConnection = nothing
				If Not  dcrs2.EOF Then
					Set newElem1  = OutData.createElement("STOREBINDET")
					newElem.appendchild newElem1
					Do while Not dcrs2.EOF 
						Set newElem  = OutData.createElement("BIN")
						newElem.setAttribute "NO", trim(dcrs2(1))
						newElem.setAttribute "CODE", trim(dcrs2(0))
						newElem.setAttribute "QTY","0"
						newElem1.appendchild newElem
						dcrs2.MoveNext 
					Loop
				End If
				dcrs2.Close 				
				

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
