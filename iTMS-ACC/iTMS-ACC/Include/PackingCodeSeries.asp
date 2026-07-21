<%
	'Program Name				:	PackingCodeSeries.asp
	'Module Name				:	Number Series
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 25, 2003
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
<%
	' Function for Selecting / Updating the Packing Code

	' Returns Packing Code in case the unit has been defined for,
	' else a string "No Packing Code number series defined for the Unit"
	' so accordingly adjust the program

	'1.	Application code (Production, Inventory, etc) from which application the Packing Code is asked
	'2.	Module code - For which module the Packing Code. For eg Inventory Module Code is 4 - Packing Code
	'3.	Product Code for which the Packing is been checked
	'4.	Item Code
	'5.	Classification Code
	'6.	Organization Code
	'7.	Packing Type
	'8.	Item Type
	'9.	What For - "I" - Insert, "D" - Display

	Function PackingCodeSeries(iApplication,iModuleCode,sProductCode,iItemCode,iClassCode,sUnitCode,sPackingNumber,sItemType,sWhatFor)
		If sUnitCode = "" Then
			PackingCodeSeries = "Err : Unit Code not passed"
			Exit Function
		End If

		Dim dcrs,dcrs1,dcrs2,sSql,sActivity,iSeries,iSeriesCode,iSeriesType,sActName,iLength
		Dim iOldSeriesCode, sProductWise

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
		Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

		With dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT NUMBERINGTYPE,SERIESNO FROM PRD_M_PACKINGNUMBERSERIESCHECK WHERE ORGANISATIONCODE = " & Pack(sUnitCode) & ""
			.ActiveConnection = con
			.Open
		End With
		If Not dcrs1.EOF Then
			sActivity = trim(dcrs1(0))
			iSeries = trim(dcrs1(1))

			With dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT COUNTERTYPE,NUMBERLENGTH FROM MS_NUMBERSERIES WHERE SERIESNO = " & iSeries & ""
				.ActiveConnection = con
				.Open
			End With
			If Not dcrs.EOF Then
				iSeriesType = trim(dcrs(0))
				iLength = trim(dcrs(1))
			End If
			dcrs.Close

			' Unit Wise Packing Code
			If sActivity = "U" Then
				With dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ORGANISATIONCODE,SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = " & Pack(sUnitCode) & ""
					.ActiveConnection = con
					.Open
				End With
				If Not dcrs.EOF Then
					iSeriesCode = trim(dcrs(1))
					' Return the Packing Code for the Unit
					' OrganisationCode,SeriesNumber,SeriesCode,Date
					If sWhatFor = "D" Then
						PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					ElseIf sWhatFor = "I" Then
						PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					End If
				End If
				dcrs.Close
			' Product Wise Packing Code
			ElseIf sActivity = "P" Then
				If sProductCode = "" Then
					PackingCodeSeries = "Err : Product Code not passed"
					Exit Function
				End If
				With dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT PRODUCTCODE,SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE PRODUCTCODE = " & Pack(sProductCode) & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & ""
					.ActiveConnection = con
					.Open
				End With
				If dcrs.EOF Then
					sActName = "Product Wise Packing Code"

					With dcrs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = " & Pack(sUnitCode) & ""
						.ActiveConnection = con
						.Open
					End With
					dcrs2.ActiveConnection = Nothing

					If Not dcrs2.EOF Then
						iOldSeriesCode = dcrs2(0)
					End If
					dcrs2.Close

					'sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen
					iSeriesCode = GenSeriesCodeFor(sUnitCode,iApplication,iModuleCode,iSeries,iSeriesType,"",sActName,iLength, iOldSeriesCode)
					'Response.Write iSeriesCode & vbCrLf

					sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,PRODUCTCODE,SERIESNO,SERIESCODE) VALUES " &_
						"(" & Pack(sUnitCode) & "," & Pack(sProductCode) & ","& iSeries & "," & iSeriesCode & ")"
					con.Execute sSql
					'Response.Write sSql & vbCrLf

					' Return the Packing Code for the Product
					' OrganisationCode,SeriesNumber,SeriesCode,Date
					If sWhatFor = "D" Then
						PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					ElseIf sWhatFor = "I" Then
						PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					End If
				Else
					iSeriesCode = trim(dcrs(1))
					' Return the Packing Code for the Product
					' OrganisationCode,SeriesNumber,SeriesCode,Date
					If sWhatFor = "D" Then
						PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					ElseIf sWhatFor = "I" Then
						PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					End If
				End If
				dcrs.Close
			' Packing Type Wise Packing Code
			ElseIf sActivity = "K" Then
				If sPackingNumber = "" Then
					PackingCodeSeries = "Err : Activity Type not passed"
					Exit Function
				End If
				'Response.Write "SELECT PRODUCTCODE,SERIESCODE,PACKINGTYPE FROM PRD_M_PACKINGNUMBERSERIES WHERE PRODUCTCODE = " & Pack(sProductCode) & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & " AND PACKINGTYPE = " & sPackingNumber & ""
				With dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT PRODUCTCODE,SERIESCODE,PACKINGTYPE FROM PRD_M_PACKINGNUMBERSERIES WHERE PRODUCTCODE = " & Pack(sProductCode) & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & " AND PACKINGTYPE = " & sPackingNumber & ""
					.ActiveConnection = con
					.Open
				End With

				If dcrs.EOF Then
					sActName = "Product cum Packing Type Wise Packing Code"

					With dcrs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = " & Pack(sUnitCode) & ""
						.ActiveConnection = con
						.Open
					End With
					dcrs2.ActiveConnection = Nothing

					If Not dcrs2.EOF Then
						iOldSeriesCode = dcrs2(0)
					End If
					dcrs2.Close

					'sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen
					iSeriesCode = GenSeriesCodeFor(sUnitCode,iApplication,iModuleCode,iSeries,iSeriesType,"",sActName,iLength, iOldSeriesCode)
					'Response.Write iSeriesCode & vbCrLf

					sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,PRODUCTCODE,PACKINGTYPE,SERIESNO,SERIESCODE) VALUES " &_
						"(" & Pack(sUnitCode) & "," & Pack(sProductCode) & "," & sPackingNumber & "," & iSeries & "," & iSeriesCode & ")"
					con.Execute sSql
					'Response.Write sSql & vbCrLf

					' Return the Packing Code for the Product and Packing Type
					' OrganisationCode,SeriesNumber,SeriesCode,Date
					If sWhatFor = "D" Then
						PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					ElseIf sWhatFor = "I" Then
						PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					End If
				Else
					iSeriesCode = trim(dcrs(1))
					' Return the Packing Code for the Product and Packing Type
					' OrganisationCode,SeriesNumber,SeriesCode,Date
					If sWhatFor = "D" Then
						PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					ElseIf sWhatFor = "I" Then
						PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
					End If
				End If
				dcrs.Close
			' Item Type Wise Packing Code
			ElseIf sActivity = "I" Then
			'	If sItemType = "" Then
			'		PackingCodeSeries = "Err : Item Type not passed"
			'		Exit Function
			'	End If

				With dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT PRODUCTWISE FROM PRD_M_PACKINGNUMBERSERIESITEMTYPE WHERE ORGANISATIONCODE = '"&sUnitCode&"' "
					.ActiveConnection = con
					.Open
				End With
				dcrs.ActiveConnection = Nothing

				If Not dcrs.EOF Then
					sProductWise = dcrs(0)
				End If
				dcrs.Close

				If sProductWise = "1" Then
					With dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ITEMTYPEID,SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = " & Pack(sUnitCode) & " AND PRODUCTCODE = "&Pack(sProductCode)&""
						.ActiveConnection = con
						.Open
					End With
					If dcrs.EOF Then
						sActName = "Item Type Wise Product Packing Code"

						With dcrs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = " & Pack(sUnitCode)
							.ActiveConnection = con
							.Open
						End With
						dcrs2.ActiveConnection = Nothing

						If Not dcrs2.EOF Then
							iOldSeriesCode = dcrs2(0)
						End If
						dcrs2.Close

						'sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen
						iSeriesCode = GenSeriesCodeFor(sUnitCode,iApplication,iModuleCode,iSeries,iSeriesType,"",sActName,iLength, iOldSeriesCode)
						'Response.Write iSeriesCode & vbCrLf

						sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,ITEMTYPEID,SERIESNO,SERIESCODE, PRODUCTCODE) VALUES " &_
							"(" & Pack(sUnitCode) & "," & Pack(sItemType) & "," & iSeries & "," & iSeriesCode & ","&Pack(sProductCode)&")"
						con.Execute sSql
						'Response.Write sSql & vbCrLf

						' Return the Packing Code for the Item Type
						If sWhatFor = "D" Then
							PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						ElseIf sWhatFor = "I" Then
							PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						End If
					Else
						iSeriesCode = trim(dcrs(1))
						' Return the Packing Code for the Item Type
						' OrganisationCode,SeriesNumber,SeriesCode,Date
						If sWhatFor = "D" Then
							PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						ElseIf sWhatFor = "I" Then
							PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						End If
					End If
					dcrs.Close
				ElseIf sProductWise = "0" Then
					With dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ITEMTYPEID,SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = " & Pack(sUnitCode) & ""
						.ActiveConnection = con
						.Open
					End With
					If dcrs.EOF Then
						sActName = "Item Type Wise Packing Code"

						With dcrs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = " & Pack(sUnitCode) 
							.ActiveConnection = con
							.Open
						End With
						dcrs2.ActiveConnection = Nothing

						If Not dcrs2.EOF Then
							iOldSeriesCode = dcrs2(0)
						End If
						dcrs2.Close

						'sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen
						iSeriesCode = GenSeriesCodeFor(sUnitCode,iApplication,iModuleCode,iSeries,iSeriesType,"",sActName,iLength, iOldSeriesCode)
						'Response.Write iSeriesCode & vbCrLf

						sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,ITEMTYPEID,SERIESNO,SERIESCODE) VALUES " &_
							"(" & Pack(sUnitCode) & "," & Pack(sItemType) & "," & iSeries & "," & iSeriesCode & ")"
						con.Execute sSql
						'Response.Write sSql & vbCrLf

						' Return the Packing Code for the Item Type
						If sWhatFor = "D" Then
							PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						ElseIf sWhatFor = "I" Then
							PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						End If
					Else
						iSeriesCode = trim(dcrs(1))
						' Return the Packing Code for the Item Type
						' OrganisationCode,SeriesNumber,SeriesCode,Date
						If sWhatFor = "D" Then
							PackingCodeSeries = GetSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						ElseIf sWhatFor = "I" Then
							PackingCodeSeries = GenSeriesNumber(sUnitCode,iSeries,iSeriesCode,FormatDate(date()))
						End If
					End If
					dcrs.Close
				End If
			End If
		Else
			PackingCodeSeries = "Err : No Packing Code number series defined for the Unit"
		End If
		dcrs1.Close

	End Function
%>

<%
	Function GenSeriesCodeFor(sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen, iOldSeriesCode)
	Dim iSeriesCode
	Dim objRsSeries,sQuery
	Dim iEntryNo,sPeriod,iNumber,sPrefix,sSufix

		Set objRsSeries = Server.CreateObject("ADODB.RecordSet")
		sQuery = "SELECT ISNULL(MAX(SERIESCODE),0)+1 FROM APP_R_NOSERIESMODULES WHERE OUDEFINITIONID='"&sOrgid&"' and SeriesNo="&iSeriesNo
		'Response.Write sQuery &"<br>"
		With objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End With
		iSeriesCode =objRsSeries(0)
		objRsSeries.close

		sQuery = "INSERT INTO APP_R_NOSERIESMODULES(OUDEFINITIONID, SERIESNO, SERIESCODE, "&_
				"MODULECODE, APPLICATIONCODE, DESCRIPTION, LASTUSEDDATE,COUNTERTYPE,NUMBERLENGTH)"&_
				"VALUES('"&sOrgid&"',"&iSeriesNo&","&iSeriesCode&","&iModuleCode&","&iAppcode&",'"&sDescription&"',"&_
				"NULL,'"&sType&"',"&iLen&")"
		'Response.Write sQuery &"<br><br>"
		con.Execute sQuery

		sQuery = "SELECT DISTINCT ENTRYNO,PERIOD,NUMBER,PREFIX,SUFFIX FROM APP_R_NOSERIESMODULEENTRY WHERE SERIESNO="&iSeriesNo&" AND OUDEFINITIONID = '"&sOrgid&"' AND SERIESCODE = "&iOldSeriesCode&""
'Response.Write sQuery
		With objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End With

		If Not objRsSeries.EOF Then
			iEntryNo = Cint(trim(objRsSeries(0)))
			sPeriod = GetPeriod(getFromFinYearLocal(),CDbl(objRsSeries(1))-1,sType)
			iNumber = trim(objRsSeries(2))
			sPrefix = trim(objRsSeries(3))
			sSufix = trim(objRsSeries(4))
			sQuery = "INSERT INTO APP_R_NoSeriesModuleEntry(OUDefinitionID, SeriesNo, SeriesCode, "&_
					"EntryNo,Period, Number, Prefix, Suffix)VALUES('"&sOrgid&"',"&iSeriesNo&","&iSeriesCode&","&_
					""&iEntryNo&",'"&sPeriod&"',"&iNumber&",'"&sPrefix&"','"&sSufix&"')"
			'Response.Write sQuery &"<br>"
			con.Execute sQuery
		End If
		Set objRsSeries = Nothing
		GenSeriesCodeFor = iSeriesCode
	End Function
%>
<%
	Function getFromFinYearLocal()
		dim arrFin,sFinFrom,sMonYr,sTempMonYr,sFinTo

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)

		sTempMonYr = Mid(sFinFrom,4,2)
		sMonYr = sTempMonYr&Year(sFinFrom)

		sFinTo = arrFin(1)

		getFromFinYearLocal = sMonYr
	End Function
%>