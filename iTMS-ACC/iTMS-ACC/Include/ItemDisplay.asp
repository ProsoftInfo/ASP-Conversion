<%
	'Program Name				:	ItemDisplay.asp
	'Module Name				:	Inventory (Item Display)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 28, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	Item Code and Classification Code
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
	'Description				:	Function to get the values based on display order
%>
<%
	'Function to get the display order
	Function ItemDisplay(iItemCode,iClassCode) 'Internal ItemCode, ClassificationCode
		dim dcrs, dcrs1, sTemp
		dim sItemCode, sItemName, sClassName, iDrawingNo, iCatalogueNo

		sTemp = ""

		'If classcode is null or 0 then class name will not be returned
		if trim(iClassCode) = "0" or trim(iClassCode) = "NULL" or trim(iClassCode) = "" then iClassCode = 0

		set dcrs = server.CreateObject("ADODB.Recordset")
		set dcrs1 = server.CreateObject("ADODB.Recordset")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT  COMPANYITEMCODE, ITEMDESCRIPTION, ISNULL(DRAWINGNUMBER,'N/A'), ISNULL(CATALOGUENO,'N/A') FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode

			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = Nothing
		if not dcrs.EOF then
			sItemCode = dcrs(0)
			sItemName = ucase(trim(dcrs(1)))
			iDrawingNo = dcrs(2)
			iCatalogueNo = dcrs(3)

			sTemp = sTemp & sItemName & "-" 

			if sTemp<>"" then sTemp = left(sTemp,len(sTemp)-1)

		end if
		dcrs.Close

		ItemDisplay = sTemp
	End Function
%>

<%
	'Function to get the display order
	Function ItemDisplayWOClass(iItemCode,iClassCode) 'Internal ItemCode, ClassificationCode
		dim dcrs, dcrs1, sTemp
		dim sItemCode, sItemName, sClassName, iDrawingNo, iCatalogueNo

		sTemp = ""

		'If classcode is null or 0 then class name will not be returned
		if trim(iClassCode) = "0" or trim(iClassCode) = "NULL" or trim(iClassCode) = "" then iClassCode = 0

		set dcrs = server.CreateObject("ADODB.Recordset")
		set dcrs1 = server.CreateObject("ADODB.Recordset")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT  COMPANYITEMCODE, ITEMDESCRIPTION, ISNULL(DRAWINGNUMBER,''), ISNULL(CATALOGUENO,'') FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode

			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = Nothing
		if not dcrs.EOF then
			sItemCode = dcrs(0)
			sItemName = ucase(trim(dcrs(1)))
			iDrawingNo = dcrs(2)
			iCatalogueNo = dcrs(3)

			sTemp = sTemp & sItemName & "-"
			if sTemp<>"" then sTemp = left(sTemp,len(sTemp)-1)

		end if
		dcrs.Close

		ItemDisplayWOClass = sTemp
	End Function
%>
