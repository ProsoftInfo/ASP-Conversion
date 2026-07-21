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
	'Program Name				:	XMLGetLotDetailsForSingleStore.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	TAJUDEEN S
	'Created On					:	May 25, 2004
	'Modified By				:	KUMAR K A
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
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/Populate.asp" -->

<%
	dim dcrs,dcrs1,dcrs2,sSql,OutData,sType,Root,newElem,NewNode,TempNode,StoreNode,ItemNode
	dim iRecNo, iItem, iClass, sOrgID, sRecType,iStoreQty
	dim dActionQty,iCount,iNoOfStores,sTemp,iStoreEntryNo
	Dim sStore,sBin,iRate,iStoreValue
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")

	
	OutData.Load server.MapPath("../temp/transaction/RECEIPTEX"&Session.SessionID&".xml")
	

	Set Root = OutData.documentElement
	

	iRecNo = Root.getAttribute("RECNO")
	sOrgID	= Root.getAttribute("UNIT")
	For each ItemNode in Root.ChildNodes
		
		if ItemNode.NodeName = "ITEM" then
		
			iItem	= ItemNode.getAttribute("ITEM")
			iClass	= ItemNode.getAttribute("CLASS")
			
	
	
			iNoOfStores = 0
			
			sTemp = "./ITEM[@ITEM='"& iItem &"' and @CLASS='" & iClass & "']/STORAGE"
			set TempNode = Root.selectnodes(stemp)
			iNoOfStores =  TempNode.length
			
				
		'	if iNoOfStores = 1  then
			
			
			
				For each StoreNode in ItemNode.ChildNodes
			
					if StoreNode.NodeName = "STORAGE" then
					    iStoreEntryNo = StoreNode.getAttribute("STOENTRYNO")
						iStoreQty = StoreNode.getAttribute("STOREQTY")
						sRecType = StoreNode.getAttribute("RECTYPE")
						iItem = StoreNode.getAttribute("ITEM")
						iClass = StoreNode.getAttribute("CLASS")
						sStore = StoreNode.getAttribute("STORE")
						sBin = StoreNode.getAttribute("BIN")
						iStoreValue= StoreNode.getAttribute("STOREVALUE")
						if trim(iStoreValue)="" or IsNull(iStoreValue) then iStoreValue = "0"
						if cdbl(iStoreValue)>0 then
						    iRate = cdbl(iStoreValue)/cdbl(iStoreQty)
						end if
						
						if trim(iRate)="" or IsNull(iRate) then iRate = "0"
						
						with dcrs1
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ISNull(MILLLOTNO,''), MILLSERIALNO, MILLNETTWEIGHT, MILLTAREWEIGHT, ISNULL(SELLINGUNITID,0), ISNULL(WEIGHTPERSELLINGFORM,0) , PACKINGCODE, ISNULL(PACKINGFORM,0), ISNULL(MILLPACKINGNUMBER,''), MILLGROSSWEIGHT,IsNull(AttributeList,''),STOCKQUALITY FROM VW_PURCHASE_RCPT_SERIAL WHERE RECEIPTNUMBER = " & iRecNo & " AND  ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID) & ""
							.ActiveConnection = con
							.Open
						end with
						set dcrs1.ActiveConnection = nothing

						do while not dcrs1.eof
						
							Set newElem = OutData.createElement("LOT")	
							newElem.setAttribute "LOTENTRYNO",iStoreEntryNo
							newElem.setAttribute "ITEM",iItem
							newElem.setAttribute "CLASS",iClass
							newElem.setAttribute "STORE",sStore
							newElem.setAttribute "BIN",sBin
							newElem.setAttribute "LOT",trim(dcrs1(0))
							newElem.setAttribute "QTY",trim(dcrs1(2))
							newElem.setAttribute "RATE",iRate
							newElem.setAttribute "GROSSQTY",trim(dcrs1(9))
							newElem.setAttribute "PACKINGNUMBER",trim(dcrs1(8))
							newElem.setAttribute "PACKINGCODE",trim(dcrs1(6))
							newElem.setAttribute "SELLINGNUMBER",trim(dcrs1(4))
							newElem.setAttribute "WEIGHTPERSELLINGFORM",trim(dcrs1(5))
							newElem.setAttribute "SELLINGFORM",trim(dcrs1(7))
							newElem.setAttribute "STAGE","0"
							newElem.setAttribute "ATTRIBUTE",trim(dcrs1(10))
							newElem.setAttribute "SERIALNO",trim(dcrs1(1))
							newElem.setAttribute "SQ",trim(dcrs1(11))
							newElem.setAttribute "FLAG","Y"
							
							StoreNode.appendChild newElem
							dcrs1.MoveNext
						loop
						dcrs1.Close

					end if 'if StoreNode.NodeName = "STORAGE" then
				Next 'For each StoreNode in ItemNode.ChildNodes	
				
		'	end if 'if iNoOfStores = 1  then
			
		end if 'if ItemNode.NodeName = "ITEM" then
	Next 'For each ItemNode in Root.ChildNodes
					
	OutData.Save server.MapPath("../temp/transaction/RECEIPTEX"&Session.SessionID&".xml")
	
%>
