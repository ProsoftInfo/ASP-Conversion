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
	'Program Name				:	ItmMergeStockDet.asp
	'Module Name				:	Inventory 
	'Author Name				:	Ragavendran R
	'Created On					:	April 22,2011
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
	dim sTemp
	
	if len(Month(date())) = 1 then
		sTempMonYr = "0"&Month(date())
	else
		sTempMonYr = Month(date())
	end if

	sMonYr = sTempMonYr&Year(date())

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	
	sorgID = Session("organizationcode")
	iItem = Request.QueryString("ItemCode")
	iClass = Request.QueryString("ClassCode")

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	Set ItemData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
	set dcrs3 = Server.CreateObject("ADODB.Recordset")
	
	Set Root = OutData.createElement("Root")
	OutData.appendChild Root
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0),LOCATIONNAME,LOCATIONCODE FROM VWITEMSTORAGEDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1"
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
						Root.appendChild newElem
					else
						Set newElem = OutData.createElement("LOCDET")
						newElem.setAttribute "LOC", trim(sLoc)
						newElem.setAttribute "BIN", "0"
						newElem.setAttribute "LOCNAME", trim(sLocName)
						newElem.setAttribute "BINNAME", "-"
						newElem.setAttribute "RECNUM", iRcptNumbering 
						newElem.setAttribute "ITEMCODE", iItem			
						Root.appendChild newElem
					end if
					dcrs1.Close
					
					sSql = "Select YearClosingStock from VwYearlyStock where ItemCode = "& iItem &" and Convert(datetime,FinancialYearFrom,103) = Convert(datetime,'"& sFinFrom &"',103)"
					dcrs2.Open sSql,con
					if not dcrs2.EOF then
						Do While Not dcrs2.EOF
							    iTotLotQty = cdbl(iTotLotQty) + cdbl(trim(dcrs2(0)))
							    Set newElem1 = OutData.createElement("PICK")

							    newElem1.setAttribute "QTYSTK", cdbl(trim(dcrs2(0)))
							    newElem.appendChild newElem1
							    iStockNo = "0"
						    dcrs2.MoveNext
						loop
					end if
					dcrs2.Close
    				dcrs.MoveNext
				Loop
				dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
