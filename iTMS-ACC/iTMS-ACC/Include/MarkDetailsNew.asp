<%
	'Program Name				:	MarkDetailsNew.asp
	'Module Name				:	Inventory (Mark Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	September 04, 2003
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

<%	' Function for Inventory Application for Marking the Details of an MR
	Function MarkMR(iItemCode,iClass,iEntNo,sOrgID,sPickLoc,sPickBin,sPickLot,sPickQty,iInvRecNo)
		dim dcrs,dcrs1,sSql
		dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iValue,iStockNo,iItmRate
	
		if sPickBin = "0" then sPickBin = "NULL"
	
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if

		sMonYr = sTempMonYr&Year(date())
		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			'Response.Write dcrs.Source 
			If Not dcrs(0) = "0" then 
				iValue = sPickQty * (cdbl(dcrs(1)) / cdbl(dcrs(0)))
			Else
				iValue = sPickQty 
			End IF
		end if
		dcrs.Close

		if sPickLot <> "NULL" then
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				 .Source = "SELECT ISNULL(RATE,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ISNULL(ITEMENTRYNO,0) = " & iEntNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sPickLoc & " AND (STORAGEBINNUMBER = " & sPickBin & " OR STORAGEBINNUMBER IS NULL) AND LOTNUMBER = " & sPickLot & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iItmRate = cdbl(dcrs(0))
			end if
			dcrs.Close
			iValue = sPickQty * iItmRate
		end if

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			if cdbl(dcrs(0)) < iValue then
				iValue = cdbl(dcrs(0))
			else
				iValue = iValue
			end if
		end if
		dcrs.Close
''for the double entry in issue based mr mark case so blocked by ragav on Sep 05,2012
	'	sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET " &_
	'		"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & sPickQty & "), " &_
	'		"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & "), "&_
	'		"YEARRESERVED = (YEARRESERVED + " & sPickQty & ") WHERE " &_
	'		"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
	'		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
	'		"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
	'		"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
	'	Response.Write "1="&sSql & vbCrLf & vbCrLf
	'	con.Execute sSql
	''end
	End Function
%>

<%	' Function for Inventory Application for UnMarking the Details of an MR
	Function UnMarkMR(iItemCode,iClass,sOrgID,sPickLoc,sPickBin,sPickLot,iMarkedQty,iMarkQty)
		dim dcrs,dcrs1,sSql
		dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iValue,iStockNo

		if sPickBin = "0" then sPickBin = "NULL"

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iValue = (cdbl(iMarkedQty) - cdbl(iMarkQty)) * (cdbl(dcrs(1)) / cdbl(dcrs(0)))
		end if
		dcrs.Close

		if sPickLot <> "NULL" then
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(RATE,0) FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sPickLoc & " AND (STORAGEBINNUMBER = " & sPickBin & " OR STORAGEBINNUMBER IS NULL) AND LOTNUMBER = " & sPickLot & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iItmRate = cdbl(dcrs(0))
			end if
			dcrs.Close
			iValue = (cdbl(iMarkedQty) - cdbl(iMarkQty)) * iItmRate
		end if

		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET " &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(iMarkedQty) - cdbl(iMarkQty) & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write "2="&sSql & vbCrLf & vbCrLf
		con.Execute sSql
		if 1 = 2 then
		sSql = "UPDATE INV_T_ITEMLOCYEARLYSTOCK SET " &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(iMarkedQty) - cdbl(iMarkQty) & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKFIFO SET " &_
			"FIFOYEARCLSGSTOCK = (FIFOYEARCLSGSTOCK + " & cdbl(iMarkedQty) - cdbl(iMarkQty) & "), " &_
			"FIFOYEARCLSGVALUE = (FIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FIFOFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FIFOFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKLIFO SET " &_
			"LIFOYEARCLSGSTOCK = (LIFOYEARCLSGSTOCK + " & cdbl(iMarkedQty) - cdbl(iMarkQty) & "), " &_
			"LIFOYEARCLSGVALUE = (LIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,LIFOFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,LIFOFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKWA SET " &_
			"WAYEARCLSGSTOCK = (WAYEARCLSGSTOCK + " & cdbl(iMarkedQty) - cdbl(iMarkQty) & "), " &_
			"WAYEARCLSGVALUE = (WAYEARCLSGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,WAFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,WAFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
		end if 'if 1 = 2 then
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STOCKNO FROM INV_M_STOCKSTATUS WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iStockNo = trim(dcrs(0))
			sSql = "UPDATE INV_M_STOCKSTATUS SET RESERVED = (ISNULL(RESERVED,0) + (" & cdbl(iMarkQty) - cdbl(iMarkedQty) & ")) WHERE STOCKNO = " & iStockNo & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STOCKNO FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND (LOTNO = " & sPickLot & " OR LOTNO IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			if sPickLot = "NULL" then
				sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) + (" & cdbl(iMarkQty) - cdbl(iMarkedQty) & ")) " &_
					"WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL"
			else
				sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) + (" & cdbl(iMarkQty) - cdbl(iMarkedQty) & ")) " &_
					"WHERE STOCKNO = " & iStockNo & " AND LOTNO = " & sPickLot & ""
			end if
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

	End Function
%>

<%	' Function for Sales Application for Reserving the Details of an Sales Order
	Function ReserveLocationStockQty(iRefNumber,iItemCode,iClass,sOrgID,sPickLoc,sPickBin,sPickQty)
		dim dcrs,dcrs1,sSql
		dim iStockNo,iPickNo
		dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iValue

		if sPickBin = "0" then sPickBin = "NULL"

		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(FormatDate(date()))

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(PICKNUMBER)+1,1) FROM INV_T_MRSISSUEPICK"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iPickNo = trim(dcrs(0))
		end if
		dcrs.Close

		sSql = "INSERT INTO INV_T_MRSISSUEPICK (PICKNUMBER,ORGANISATIONCODE," &_
			"CLASSIFICATIONCODE,ITEMCODE,QUANTITYFORISSUE,MARKEDON,SALESREFNUMBER,MARKEDFROM) VALUES " &_
			"(" & iPickNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
			"" & sPickQty & ",CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103)," & Pack(iRefNumber) & ",'S')"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STOCKNO FROM INV_M_STOCKSTATUS WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iStockNo = trim(dcrs(0))
			sSql = "UPDATE INV_M_STOCKSTATUS SET RESERVED = (ISNULL(RESERVED,0) + " & sPickQty & ") WHERE STOCKNO = " & iStockNo & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.close

		ReserveLocationStockQty = iPickNo&":"&iStockNo

	End Function
%>

<%	' Function for Sales Application for Reserving the Lot Details of an Sales Order
	Function ReserveLotQty(iPickNumber,sPickLoc,sPickBin,sPickLot,sPickQty,sWhatFor, iProcNo)
		dim dcrs,sSql
		dim iLineNo
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		if sPickLot = "" or sPickLot = "0" then sPickLot = "NULL"
		if sPickLot <> "NULL" then sPickLot = Pack(sPickLot)
		if sPickBin = "0" then sPickBin = "NULL"

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(LINENUMBER)+1,1) FROM INV_T_MRSISSUEPICKDETAILS"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iLineNo = trim(dcrs(0))
		end if
		dcrs.close

		sSql = "INSERT INTO INV_T_MRSISSUEPICKDETAILS (LINENUMBER,PICKNUMBER,LOCATIONNO," &_
			"BINNUMBER,LOTNUMBER,QUANTITYFORISSUE,MARKEDFOR, PROCESSINGNO) VALUES " &_
			"(" & iLineNo & "," & iPickNumber & "," & sPickLoc & "," & sPickBin & "," &_
			"" & sPickLot & "," & sPickQty & "," & Pack(sWhatFor) & ", "&iProcNo&")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		ReserveLotQty = iLineNo

	End Function
%>

<%	' Function for Sales Application for Reserving the Lot / Serial Details of an Sales Order
	Function ReserveLotSerialStatusQty(iLineNumber,iStockNumber,sPickLot,iSerialNo,sPickQty,sWhatFor,iItemCode,iClass,sOrgID,sPickLoc,sPickBin)
		dim dcrs,dcrs1,sSql
		dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iValue,iYrCloValue,iYrCloQty

		if sPickLot = "0" or sPickLot = "" then sPickLot = "NULL"
		if sPickLot <> "NULL" then sPickLot = Pack(sPickLot)
		if sPickBin = "0" then sPickBin = "NULL"

		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(FormatDate(date()))

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)

		if iSerialNo = "0" then
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iYrCloQty = cdbl(dcrs(0))
				iYrCloValue = cdbl(dcrs(1))
			end if
			dcrs.Close
			iValue = cdbl(sPickQty) * (iYrCloValue / iYrCloQty)
			iSerialNo = "NULL"
		else
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(RATE,0) FROM INV_T_RECEIPTLOTDETAILS WHERE SERIALNUMBER = " & iSerialNo & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iValue = cdbl(sPickQty) * cdbl(dcrs(0))
			end if
			dcrs.Close
		end if

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			if cdbl(dcrs(0)) < iValue then
				iValue = cdbl(dcrs(0))
			else
				iValue = iValue
			end if
		end if
		dcrs.Close

		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET " &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & sPickQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write "3="&sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_ITEMLOCYEARLYSTOCK SET " &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & sPickQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKFIFO SET " &_
			"FIFOYEARCLSGSTOCK = (FIFOYEARCLSGSTOCK - " & sPickQty & "), " &_
			"FIFOYEARCLSGVALUE = (FIFOYEARCLSGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FIFOFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FIFOFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKLIFO SET " &_
			"LIFOYEARCLSGSTOCK = (LIFOYEARCLSGSTOCK - " & sPickQty & "), " &_
			"LIFOYEARCLSGVALUE = (LIFOYEARCLSGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,LIFOFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,LIFOFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKWA SET " &_
			"WAYEARCLSGSTOCK = (WAYEARCLSGSTOCK - " & sPickQty & "), " &_
			"WAYEARCLSGVALUE = (WAYEARCLSGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,WAFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,WAFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "INSERT INTO INV_T_MRSISSUEPICKSERIAL (LINENUMBER,SERIALNO,QUANTITYFORISSUE,MARKEDFOR) VALUES " &_
			"(" & iLineNumber & "," & iSerialNo & "," & sPickQty & "," & Pack(sWhatFor) & ")"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STOCKNO FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND (LOTNO = " & sPickLot & " OR LOTNO IS NULL) AND (SERIALNO = " & iSerialNo & " OR SERIALNO IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_M_STOCKSTATUSDETAILS (STOCKNO,LOTNO,SERIALNO,RESERVED) VALUES " &_
				"(" & iStockNumber & "," & sPickLot & "," & iSerialNo & "," & sPickQty & ")"
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			if sPickLot = "NULL" and iSerialNo = "NULL" then
				sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) + " & sPickQty & ") " &_
					"WHERE STOCKNO = " & iStockNumber & " AND LOTNO IS NULL AND SERIALNO IS NULL"
			elseif sPickLot <> "NULL" and iSerialNo = "NULL" then
				sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) + " & sPickQty & ") " &_
					"WHERE STOCKNO = " & iStockNumber & " AND LOTNO = " & sPickLot & " AND SERIALNO IS NULL"
			elseif sPickLot = "NULL" and iSerialNo <> "NULL" then
				sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) + " & sPickQty & ") " &_
					"WHERE STOCKNO = " & iStockNumber & " AND LOTNO IS NULL AND SERIALNO = " & iSerialNo & ""
			else
				sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) + " & sPickQty & ") " &_
					"WHERE STOCKNO = " & iStockNumber & " AND LOTNO = " & sPickLot & " AND SERIALNO = " & iSerialNo & ""
			end if
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

	End Function
%>

<%	' Function for Sales Application for UnReserving the Details of an Sales Order
	Function UnReserveLocationStockQty(iRefNumber,iItemCode,iClass,sOrgID,sPickLoc,sPickBin,sPickQty,sWhatFor)
		dim dcrs,dcrs1,dcrs2,sSql
		dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iValue,iStockNo,iLineNo,iPickNo
		dim sPickLot

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
		Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with

		'Response.Write dcrs.source
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iValue = cdbl(sPickQty) * (cdbl(dcrs(1)) / cdbl(dcrs(0)))
		Else
			iValue = 0
		end if
		dcrs.Close

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)

		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET " &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(sPickQty) & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write "Chk="&sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_ITEMLOCYEARLYSTOCK SET " &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(sPickQty) & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKFIFO SET " &_
			"FIFOYEARCLSGSTOCK = (FIFOYEARCLSGSTOCK + " & cdbl(sPickQty) & "), " &_
			"FIFOYEARCLSGVALUE = (FIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FIFOFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FIFOFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKLIFO SET " &_
			"LIFOYEARCLSGSTOCK = (LIFOYEARCLSGSTOCK + " & cdbl(sPickQty) & "), " &_
			"LIFOYEARCLSGVALUE = (LIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,LIFOFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,LIFOFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_YEARLYSTOCKWA SET " &_
			"WAYEARCLSGSTOCK = (WAYEARCLSGSTOCK + " & cdbl(sPickQty) & "), " &_
			"WAYEARCLSGVALUE = (WAYEARCLSGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,WAFINYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,WAFINYEARTO,103)"
		Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STOCKNO FROM INV_M_STOCKSTATUS WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iStockNo = trim(dcrs(0))
			sSql = "UPDATE INV_M_STOCKSTATUS SET RESERVED = (ISNULL(RESERVED,0) - (" & cdbl(sPickQty) & ")) WHERE STOCKNO = " & iStockNo & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.close


		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PICKNUMBER FROM INV_T_MRSISSUEPICK WHERE SALESREFNUMBER = " & Pack(iRefNumber) & " AND MARKEDFROM = 'S'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				iPickNo = trim(dcrs(0))

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT LINENUMBER FROM INV_T_MRSISSUEPICKDETAILS WHERE PICKNUMBER = " & iPickNo & " AND MARKEDFOR = " & Pack(sWhatFor) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					iLineNo = trim(dcrs1(0))

					sSql = "DELETE INV_T_MRSISSUEPICKSERIAL WHERE LINENUMBER = " & iLineNo & " AND " &_
						"MARKEDFOR = " & Pack(sWhatFor) & ""
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql

					with dcrs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(LOTNUMBER,'-') FROM INV_T_MRSISSUEPICKDETAILS WHERE LINENUMBER = " & iLineNo & " AND PICKNUMBER = " & iPickNo & " AND MARKEDFOR = " & Pack(sWhatFor) & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs2.ActiveConnection = nothing

					if not dcrs2.EOF then
						sPickLot = trim(dcrs2(0))
					end if
					dcrs2.Close

					if sPickLot = "-" then
						sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) - " & sPickQty & ") " &_
							"WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL"
						Response.Write sSql & vbCrLf & vbCrLf
					elseif sPickLot <> "-" then
						sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) - " & sPickQty & ") " &_
							"WHERE STOCKNO = " & iStockNo & " AND LOTNO = " & Pack(sPickLot) & ""
						Response.Write sSql & vbCrLf & vbCrLf
					end if
					con.Execute sSql

					sSql = "DELETE INV_T_MRSISSUEPICKDETAILS WHERE LINENUMBER = " & iLineNo & " AND " &_
						"PICKNUMBER = " & iPickNo & " AND MARKEDFOR = " & Pack(sWhatFor) & ""
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
				dcrs1.Close

				sSql = "DELETE INV_T_MRSISSUEPICK WHERE PICKNUMBER = " & iPickNo & ""
				Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql

			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	End Function
%>

<%	' Function for Short Closing MR from Inventory Application
	Function MRShortClose(iRefNumber,iItemCode,iClass,sOrgID)
		dim dcrs,dcrs1,dcrs2,sSql
		dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iValue,iStockNo,iLineNo,iPickNo
		dim sPickLot,iQtyForIssue,iQtyIssued,iQtyDetForIssue,iQtyDetIssued
		dim sPickLoc,sPickBin,sPickQty

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
		Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PICKNUMBER,ISNULL(QUANTITYFORISSUE,0),ISNULL(QUANTITYISSUED,0) FROM INV_T_MRSISSUEPICK WHERE MRSNUMBER = " & iRefNumber & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND MARKEDFROM = 'I' AND (ISNULL(QUANTITYFORISSUE,0) - ISNULL(QUANTITYISSUED,0)) > 0"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				iPickNo = trim(dcrs(0))
				iQtyForIssue = trim(dcrs(1))
				iQtyIssued = trim(dcrs(2))

				if cdbl(iQtyIssued) > 0 then
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT LINENUMBER,ISNULL(QUANTITYFORISSUE,0),ISNULL(QUANTITYISSUED,0),LOCATIONNO,ISNULL(BINNUMBER,0) FROM INV_T_MRSISSUEPICKDETAILS WHERE PICKNUMBER = " & iPickNo & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing

					if not dcrs1.EOF then
						do while not dcrs1.EOF
							iLineNo = trim(dcrs1(0))
							iQtyDetForIssue = trim(dcrs1(1))
							iQtyDetIssued = trim(dcrs1(2))
							sPickLoc = trim(dcrs1(3))
							sPickBin = trim(dcrs1(4))

							if sPickBin = "0" then sPickBin = "NULL"

							sPickQty = cdbl(iQtyDetForIssue) - cdbl(iQtyDetIssued)

							sSql = "UPDATE INV_T_MRSISSUEPICKDETAILS SET QUANTITYFORISSUE = QUANTITYISSUED, " &_
								"QUANTITYSHORTCLOSED = " & cdbl(sPickQty) & " WHERE LINENUMBER = " & iLineNo & " AND PICKNUMBER = " & iPickNo & ""

							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql


			''''''''''''''''''''''''STOCK UPDATION ''''''''''''''''''''''''''''''''''''''''''''''
							with dcrs2
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
								.ActiveConnection = con
								.Open
							end with
							set dcrs2.ActiveConnection = nothing

							if not dcrs2.EOF then
								iValue = cdbl(sPickQty) * (cdbl(dcrs2(1)) / cdbl(dcrs2(0)))
							end if
							dcrs2.Close

							sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET " &_
								"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(sPickQty) & "), " &_
								"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							Response.Write "Test="&sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_ITEMLOCYEARLYSTOCK SET " &_
								"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(sPickQty) & "), " &_
								"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_YEARLYSTOCKFIFO SET " &_
								"FIFOYEARCLSGSTOCK = (FIFOYEARCLSGSTOCK + " & cdbl(sPickQty) & "), " &_
								"FIFOYEARCLSGVALUE = (FIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FIFOFINYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FIFOFINYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_YEARLYSTOCKLIFO SET " &_
								"LIFOYEARCLSGSTOCK = (LIFOYEARCLSGSTOCK + " & cdbl(sPickQty) & "), " &_
								"LIFOYEARCLSGVALUE = (LIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,LIFOFINYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,LIFOFINYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_YEARLYSTOCKWA SET " &_
								"WAYEARCLSGSTOCK = (WAYEARCLSGSTOCK + " & cdbl(sPickQty) & "), " &_
								"WAYEARCLSGVALUE = (WAYEARCLSGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,WAFINYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,WAFINYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							with dcrs2
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT STOCKNO FROM INV_M_STOCKSTATUS WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
								.ActiveConnection = con
								.Open
							end with
							set dcrs2.ActiveConnection = nothing

							if not dcrs2.EOF then
								iStockNo = trim(dcrs2(0))
								sSql = "UPDATE INV_M_STOCKSTATUS SET RESERVED = (ISNULL(RESERVED,0) - (" & cdbl(sPickQty) & ")) WHERE STOCKNO = " & iStockNo & ""
								'Response.Write sSql & "<BR><BR>"
								con.Execute sSql
							end if
							dcrs2.close

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

							'sSql = "DELETE INV_T_MRSISSUEPICKSERIAL WHERE LINENUMBER = " & iLineNo & ""
							'Response.Write sSql & "<BR><BR>"
							'con.Execute sSql

							with dcrs2
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT ISNULL(LOTNUMBER,'-') FROM INV_T_MRSISSUEPICKDETAILS WHERE LINENUMBER = " & iLineNo & " AND PICKNUMBER = " & iPickNo & ""
								.ActiveConnection = con
								.Open
							end with
							set dcrs2.ActiveConnection = nothing

							if not dcrs2.EOF then
								sPickLot = trim(dcrs2(0))
							end if
							dcrs2.Close

							if sPickLot = "-" then
								sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) - " & cdbl(sPickQty) & ") " &_
									"WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL"
							elseif sPickLot <> "-" then
								sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) - " & cdbl(sPickQty) & ") " &_
									"WHERE STOCKNO = " & iStockNo & " AND LOTNO = " & Pack(sPickLot) & ""
							end if
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

						dcrs1.MoveNext
						loop

					end if
					dcrs1.Close

					sSql = "UPDATE INV_T_MRSISSUEPICK SET QUANTITYFORISSUE = QUANTITYISSUED, " &_
						"QUANTITYSHORTCLOSED = " & cdbl(iQtyForIssue) - cdbl(iQtyIssued) & " WHERE PICKNUMBER = " & iPickNo & ""
					'Response.Write sSql & "<BR><BR>"
					con.Execute sSql

				elseif cdbl(iQtyIssued) = 0 then
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT LINENUMBER,ISNULL(QUANTITYFORISSUE,0),LOCATIONNO,ISNULL(BINNUMBER,0) FROM INV_T_MRSISSUEPICKDETAILS WHERE PICKNUMBER = " & iPickNo & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing

					if not dcrs1.EOF then
						do while not dcrs1.EOF
							iLineNo = trim(dcrs1(0))
							iQtyDetForIssue = trim(dcrs1(1))
							sPickLoc = trim(dcrs1(2))
							sPickBin = trim(dcrs1(3))

							if sPickBin = "0" then sPickBin = "NULL"

							sSql = "UPDATE INV_T_MRSISSUEPICKDETAILS SET QUANTITYFORISSUE = 0, " &_
								"QUANTITYSHORTCLOSED = " & cdbl(iQtyDetForIssue) & " WHERE LINENUMBER = " & iLineNo & " AND PICKNUMBER = " & iPickNo & ""
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

			''''''''''''''''''''''''STOCK UPDATION ''''''''''''''''''''''''''''''''''''''''''''''
							with dcrs2
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
								.ActiveConnection = con
								.Open
							end with
							set dcrs2.ActiveConnection = nothing

							if not dcrs2.EOF then
								iValue = cdbl(iQtyDetForIssue) * (cdbl(dcrs2(1)) / cdbl(dcrs2(0)))
							end if
							dcrs2.Close

							sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET " &_
								"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(iQtyDetForIssue) & "), " &_
								"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_ITEMLOCYEARLYSTOCK SET " &_
								"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & cdbl(iQtyDetForIssue) & "), " &_
								"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL) AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_YEARLYSTOCKFIFO SET " &_
								"FIFOYEARCLSGSTOCK = (FIFOYEARCLSGSTOCK + " & cdbl(iQtyDetForIssue) & "), " &_
								"FIFOYEARCLSGVALUE = (FIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FIFOFINYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FIFOFINYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_YEARLYSTOCKLIFO SET " &_
								"LIFOYEARCLSGSTOCK = (LIFOYEARCLSGSTOCK + " & cdbl(iQtyDetForIssue) & "), " &_
								"LIFOYEARCLSGVALUE = (LIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,LIFOFINYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,LIFOFINYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							sSql = "UPDATE INV_T_YEARLYSTOCKWA SET " &_
								"WAYEARCLSGSTOCK = (WAYEARCLSGSTOCK + " & cdbl(iQtyDetForIssue) & "), " &_
								"WAYEARCLSGVALUE = (WAYEARCLSGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,WAFINYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,WAFINYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

							with dcrs2
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT STOCKNO FROM INV_M_STOCKSTATUS WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) AND LOCATIONNUMBER = " & sPickLoc & " AND (BINNUMBER = " & sPickBin & " OR BINNUMBER IS NULL)"
								.ActiveConnection = con
								.Open
							end with
							set dcrs2.ActiveConnection = nothing

							if not dcrs2.EOF then
								iStockNo = trim(dcrs2(0))
								sSql = "UPDATE INV_M_STOCKSTATUS SET RESERVED = (ISNULL(RESERVED,0) - (" & cdbl(iQtyDetForIssue) & ")) WHERE STOCKNO = " & iStockNo & ""
								'Response.Write sSql & "<BR><BR>"
								con.Execute sSql
							end if
							dcrs2.close

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

							with dcrs2
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT ISNULL(LOTNUMBER,'-') FROM INV_T_MRSISSUEPICKDETAILS WHERE LINENUMBER = " & iLineNo & " AND PICKNUMBER = " & iPickNo & ""
								.ActiveConnection = con
								.Open
							end with
							set dcrs2.ActiveConnection = nothing

							if not dcrs2.EOF then
								sPickLot = trim(dcrs2(0))
							end if
							dcrs2.Close

							if sPickLot = "-" then
								sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) - " & cdbl(iQtyDetForIssue) & ") " &_
									"WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL"
							elseif sPickLot <> "-" then
								sSql = "UPDATE INV_M_STOCKSTATUSDETAILS SET RESERVED = (ISNULL(RESERVED,0) - " & cdbl(iQtyDetForIssue) & ") " &_
									"WHERE STOCKNO = " & iStockNo & " AND LOTNO = " & Pack(sPickLot) & ""
							end if
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql

						dcrs1.MoveNext
						loop
					end if
					dcrs1.Close

					sSql = "UPDATE INV_T_MRSISSUEPICK SET QUANTITYFORISSUE = 0, " &_
						"QUANTITYSHORTCLOSED = " & cdbl(iQtyForIssue) & " WHERE PICKNUMBER = " & iPickNo & ""
					'Response.Write sSql & "<BR><BR>"
					con.Execute sSql

				end if
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	End Function
%>
