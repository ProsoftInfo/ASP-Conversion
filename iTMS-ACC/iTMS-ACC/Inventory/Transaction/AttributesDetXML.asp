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
	'Program Name				:	AttributesDetXML.asp
	'Module Name				:	Inventory (Stock Management - Attribute Wise Stock)
	'Author Name				:	UmaMaheswari S
	'Created On					:	June 08, 2011
	'Modified By				:	
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
	dim dcrs,dcrs1,dcrs2,sSql,OutData,sorgID,Root,newElem,newElem1,iItem,iClass,ItemDetails
	Dim sLoc,sBin,sBinName,sLocName,sUoM,sUoMDesc,iTotLotQty,iStockQty,sDecimal
	dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,iLotQtyReserved,sLot,iSer,ndRoot,ndItem,ItemNode,sItemName
	dim sTemp,sReceiptNumberStatus,sItemTypeID,nItemQty
	
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
			end if
			dcrs.Close
			
			set ItemNode  = OutData.createElement("Item")
				ItemNode.setAttribute "ICode",iItem
				ItemNode.setAttribute "CCode",iClass
				ItemNode.setAttribute "Unit",sorgID
				ItemNode.setAttribute "IName",sItemName
				ItemNode.setAttribute "ItemTypeID",sItemTypeID 
			Root.appendChild ItemNode
			
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
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
			
			'LOT Details
			If Trim(sItemTypeID) = "FAB" or Trim(sItemTypeID) = "YRN" Then
				
				sSql = " Select L.ItemCode,L.ClassificationCode,I.ItemDescription From Inv_T_LocationLot L,Inv_M_ItemMaster I where L.SerialNumber in (select SerialNumber from Inv_T_LocationLot where ItemCode = "& iItem &" and classificationcode= "& iClass &" and SerialNumber is Not NULL and (AttributeList is NULL OR AttributeList='') )and L.ItemCode <> "& iItem &" and L.ItemCode = I.ItemCode and L.classificationcode =I.classificationcode group by L.itemcode,L.classificationcode,I.ItemDescription "
				dcrs.Open sSql,con
				
				
				Do while Not dcrs.EOF 
					
					Set newElem = OutData.createElement("BaseItem")
					newElem.setAttribute "ICode", trim(dcrs(0))
					newElem.setAttribute "CCode", trim(dcrs(1))
					newElem.setAttribute "Desc", trim(dcrs(2))
					newElem.setAttribute "OptValue", ""
					newElem.setAttribute "Selection", "N"
					
					nItemQty = Cdbl("0")
					
					sSql = "select SerialNumber,isNULL(LotQuantityNett,0) from Inv_T_LocationLot where ItemCode = "& dcrs(0) &" and Classificationcode = "& dcrs(1) &"  and SerialNumber is Not NULL "
					dcrs1.Open sSql,con
					
					Do while Not dcrs1.EOF 
						nItemQty = cdbl(nItemQty) + cdbl(dcrs1(1))
					
						Set newElem1 = OutData.createElement("Serial")
							newElem1.setAttribute "No", trim(dcrs1(0))
							newElem.appendchild newElem1 
							
						dcrs1.MoveNext 
					Loop
					dcrs1.Close 
					newElem.setAttribute "Qty", nItemQty 
					ItemNode.appendChild newElem
				
					dcrs.MoveNext 
				Loop
				dcrs.Close 
			Else
				sSql = " select LotNumber from Inv_T_LocationLot where ItemCode = "& iItem &" and classificationcode= "& iClass &" and LotNumber is Not NULL and (AttributeList is NULL OR AttributeList='')group by LotNumber "
			
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
			
				Do While Not dcrs.EOF 
					
					Set newElem = OutData.createElement("Lot")
					newElem.setAttribute "No", trim(dcrs(0))
					newElem.setAttribute "OptValue", ""
					newElem.setAttribute "Selection", "N"
					ItemNode.appendChild newElem
					
					dcrs.MoveNext 
				Loop
				dcrs.close 
				
			End IF	'If Trim(sItemTypeID) = "FAB" Then
				   
			
		Next
	End IF	
	
	set newElem = OutData.createElement("AttributeDet")
	Root.appendChild newElem 
	
	sSql = " Select a.OptionValue,a.OptionName From Inv_M_ItemTypeOptions a, Inv_M_ItemTypeAttributes b "&_
		   " where a.ItemTypeAttributeID = b.ItemTypeAttributeID and b.ItemTypeID = '"& sItemTypeID &"' Order by a.OptionValue" 		
			
	dcrs.Open sSql,con
			
	Do While Not dcrs.EOF
				
		set newElem1 = OutData.createElement("Attribute")
		newElem1.setAttribute "OptValue",dcrs(0)
		newElem1.setAttribute "OptName",dcrs(1)
		newElem.appendChild newElem1 
				
		dcrs.MoveNext 
	Loop
	dcrs.Close
	
	
	
	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
