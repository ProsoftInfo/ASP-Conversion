<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtSTInsert.asp
	'Module Name				:	Inventory (Stock Management Stock Transfer)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 28, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 22,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	stkMgmtSTEntry.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
'XML DOM Variables
Dim oDOM,Root,objfs,RootO

dim OutData

Set OutData = Server.CreateObject("Microsoft.XMLDOM")
Set RootO = OutData.createElement("Root")
OutData.appendChild RootO

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim newxml,RootNode,HeaderNode,PickNode,PickDetNode,ndItem

Set newxml = Server.CreateObject("Microsoft.XMLDOM")

newxml.async = false
newxml.load(Request)

dim dcrs,dcrs1,sSql,bFlag
dim iItemCode,iClass,arrStore,sBin,sLoc,iReqQty,iIssQty,sToBin,sToLoc,iInvRecNoRec
dim iValue,sOrgID,iTransBy,dTraDate,sMonYr,iRecQty,iItmRate,sPickLot
dim arrFin,sFinFrom,sFinTo,sTempMonYr,iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue
dim sType,iTraNo,iSerialQty,iSerialNo,iInvRecNo,iTotalValue,sInvTotValue
dim iSerialRate

iTransBy = getUserid
bFlag = true

dTraDate = FormatDate(date())

sTempMonYr = mid(dTraDate,4,2)
sMonYr = sTempMonYr&Year(dTraDate)

arrFin = split(GetFinancialYear(sMonYr),":")
sFinFrom = arrFin(0)
sFinTo = arrFin(1)


Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

Set RootNode = newxml.documentElement

'iItemCode = trim(RootNode.Attributes.getNamedItem("ITEM").Value)
'iClass = trim(RootNode.Attributes.getNamedItem("CLASS").Value)
'sOrgID = trim(RootNode.Attributes.getNamedItem("ORG").Value)
newxml.save server.MapPath("../temp/Transaction/STdata.xml")
con.beginTrans

if RootNode.HaschildNodes() then
	For Each ndItem In RootNode.childNodes
		iItemCode = ndItem.getAttribute("ICode")
		iClass = ndItem.getAttribute("CCode")
		sOrgID = ndItem.getAttribute("Unit")
		For Each HeaderNode in ndItem.childNodes
			if StrComp(HeaderNode.nodeName,"LOCDET") = 0 then
				sLoc = trim(HeaderNode.Attributes.getNamedItem("LOC").Value)
				sBin = trim(HeaderNode.Attributes.getNamedItem("BIN").Value)

				bFlag = true

				if sBin <> "0" then
					sBin = sBin
				else
					sBin = "NULL"
				end if

				For Each PickNode In HeaderNode.childNodes
					if StrComp(PickNode.nodeName,"PICK") = 0 then
						sPickLot = trim(PickNode.Attributes.getNamedItem("LOTNO").Value)
						iInvRecNo = trim(PickNode.Attributes.getNamedItem("INVRECNO").Value)
						iIssQty = cdbl(PickNode.Attributes.getNamedItem("QTYISS").Value)
						sType = trim(PickNode.Attributes.getNamedItem("TYPE").Value)
						arrStore = split(trim(PickNode.Attributes.getNamedItem("STORE").Value),":")

						if iInvRecNo = "" or IsNull(iInvRecNo) then iInvRecNo = "NULL"

						if sPickLot <> "NULL" then sPickLot = Pack(sPickLot)

						if UBound(arrStore) > 0 then
							sToLoc = arrStore(0)
							sToBin = arrStore(1)
						else
							sToLoc = arrStore(0)
							sToBin = "NULL"
						end if

						with dcrs
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ISNULL(MAX(TRANSFERNO)+1,1) FROM INV_T_INTERNALSTOCKTRANSFER"
							.ActiveConnection = con
							.Open
						end with
						set dcrs.ActiveConnection = nothing

						if not dcrs.EOF then
							iTraNo = dcrs(0)
						end if
						dcrs.close
						sInvTotValue = "0"
						' If Transfer is Full Lot
						
						if sType = "F" then
							if not iIssQty = 0 then

								with dcrs
									.CursorLocation = 3
									.CursorType = 3
									'.Source = "SELECT ITEMRATE FROM INV_T_RECEIPTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
									.Source = "SELECT RATE FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
									.ActiveConnection = con
									.Open
								end with
								set dcrs.ActiveConnection = nothing

								if not dcrs.EOF then
									iItmRate = cdbl(dcrs(0))
								end if
								dcrs.Close

								sInvTotValue = LotFullInsert()

								ReceiptFullInsert sInvTotValue
							end if
						' If Transfer is from Stock
						elseif sType = "" then
							if not iIssQty = 0 then
								StockInsert
							end if
						' If Transfer is Partial Lot
						elseif sType = "P" then
							if PickNode.hasChildNodes() then
								' To insert Header Values
								For Each PickDetNode In PickNode.childNodes
									iSerialNo = trim(PickDetNode.Attributes.Item(0).nodeValue)
									iSerialQty = trim(PickDetNode.Attributes.Item(1).nodeValue)
									if not cdbl(iSerialQty) = 0 then
										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = "SELECT ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE SERIALNUMBER = " & iSerialNo & ""
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing
										if not dcrs.EOF then
											iSerialRate = trim(dcrs(0))
										end if
										dcrs.close

										sInvTotValue = cdbl(sInvTotValue) + (iSerialQty * cdbl(iSerialRate))
									end if
								next

								with dcrs
									.CursorLocation = 3
									.CursorType = 3
									'.Source = "SELECT ITEMRATE FROM INV_T_RECEIPTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
									.Source = "SELECT RATE FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
									.ActiveConnection = con
									.Open
								end with
								set dcrs.ActiveConnection = nothing

								if not dcrs.EOF then
									iItmRate = cdbl(dcrs(0))
								end if
								dcrs.Close

								iInvRecNoRec = LotPartialInsert("H",iIssQty,sInvTotValue,"N")

								ReceiptFullInsert iInvRecNoRec&":"&sInvTotValue

								For Each PickDetNode In PickNode.childNodes
									iSerialNo = trim(PickDetNode.Attributes.Item(0).nodeValue)
									iSerialQty = trim(PickDetNode.Attributes.Item(1).nodeValue)
									if not cdbl(iSerialQty) = 0 then

										LotPartialInsert "D",iSerialQty,iInvRecNoRec,iSerialNo

										'ReceiptPartialInsert "D",iInvRecNoRec,iSerialQty,iSerialNo

									end if
								next
							end if
						end if
					end if
				next
			' end if for LOCDET Node check
			end if
		next
	next
end if

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing

' Function to Insert in case of Full Lot / Serial Selection [FROM STORE]
Function LotFullInsert()

	' Declaration of variables
	Dim dcrs,dcrs1,dcrs2
	dim iSerialNo,iSerialQty,iRate,iInvRecNewNo,iValue
	dim iNewSerNo,iGross,iNett,iTare,iQtyIssued,sPackNum,iPackCode,iSellNum,iWtPerSellFm
	dim iSellForm,sStage,iAltGross,iAltNett
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

	sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFER (TRANSFERNO,ITEMCODE,ORGANISATIONCODE," &_
		"CLASSIFICATIONCODE,FROMLOCATIONNUMBER,TOLOCATIONNUMBER," &_
		"TOBINNUMBER,FROMBINNUMBER,TRANSFERTYPE,TRANSFERREDON,TRANSFERREDBY) VALUES " &_
		"(" & iTraNo & "," & iItemCode & "," & Pack(sOrgID) & "," & iClass & "," & sLoc & "," &_
		"" & sToLoc & "," & sToBin & "," &_
		"" & sBin & ",'ST',CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & ")"
	''Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	'if sPickLot <> "NULL" then
	if 1 = 1 then

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			if sPickLot <> "NULL" then
				'.Source = "SELECT INVENTORYRECEIPTNO FROM INV_T_RECEIPTLOTDETAILS WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & sPickLot & " AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL)"
				.Source = "SELECT INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & sPickLot & " AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL)"
			else
				'.Source = "SELECT INVENTORYRECEIPTNO FROM INV_T_RECEIPTLOTDETAILS WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL)"
				.Source = "SELECT INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL)"
			end if
			Response.Write dcrs.source & vbCrLf & vbCrLf
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		' Check whether the TO Storage Location has the Lot Number, if not then insert
		if dcrs.EOF then

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_RECEIPTDETAILS"
				.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				iInvRecNewNo = dcrs1(0)
			end if
			dcrs1.close
			
			'INV_T_RECEIPTDETAILS
			sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,RECEIPTNUMBER,ORGANISATIONCODE," &_
				"ITEMCODE,CLASSIFICATIONCODE,RECEIPTQUANTITY,ITEMRATE,SRCTYPE,ACCOUNTEDON,ACCOUNTEDBY,RECEIVEDON) VALUES " &_
				"(" & iInvRecNewNo & "," & iTraNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
				"" & iIssQty & "," & iItmRate & ",'RT'," &_
				"CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103))"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				if sPickLot = "NULL" then
					.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYGROSS,0),ISNULL(LOTQUANTITYNETT,0),ISNULL(LOTQUANTITYTARE,0),ISNULL(QUANTITYISSUED,0),ISNULL(PACKINGNUMBER,'-'),ISNULL(PACKINGCODE,0),ISNULL(SELLINGNUMBER,0),ISNULL(WEIGHTPERSELLINGFORM,0),ISNULL(SELLINGFORM,0),ISNULL(STAGE,'-'),ISNULL(ALTGROSS,0),ISNULL(ALTNETT,0),ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER NOT IN (SELECT SERIALNO FROM INV_T_MRSISSUEPICKSERIAL WHERE SERIALNO IS NOT NULL) ORDER BY 1"
				else
					.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYGROSS,0),ISNULL(LOTQUANTITYNETT,0),ISNULL(LOTQUANTITYTARE,0),ISNULL(QUANTITYISSUED,0),ISNULL(PACKINGNUMBER,'-'),ISNULL(PACKINGCODE,0),ISNULL(SELLINGNUMBER,0),ISNULL(WEIGHTPERSELLINGFORM,0),ISNULL(SELLINGFORM,0),ISNULL(STAGE,'-'),ISNULL(ALTGROSS,0),ISNULL(ALTNETT,0),ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & sPickLot & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER NOT IN (SELECT SERIALNO FROM INV_T_MRSISSUEPICKSERIAL WHERE SERIALNO IS NOT NULL) ORDER BY 1"
				end if
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				Do While Not dcrs1.EOF

					iSerialNo = trim(dcrs1(0))
					iGross = trim(dcrs1(1))
					iNett = trim(dcrs1(2))
					iTare = trim(dcrs1(3))
					iQtyIssued = trim(dcrs1(4))
					sPackNum = trim(dcrs1(5))
					iPackCode = trim(dcrs1(6))
					iSellNum = trim(dcrs1(7))
					iWtPerSellFm = trim(dcrs1(8))
					iSellForm = trim(dcrs1(9))
					sStage = trim(dcrs1(10))
					iAltGross = trim(dcrs1(11))
					iAltNett = trim(dcrs1(12))
					iRate = trim(dcrs1(13))

					if sPackNum = "-" then
						sPackNum = "NULL"
					else
						sPackNum = Pack(sPackNum)
					end if
					if iPackCode = "0" then iPackCode = "NULL"
					if iSellNum = "0" then iSellNum = "NULL"
					if iWtPerSellFm = "0" then iWtPerSellFm = "NULL"
					if iSellForm = "0" then iSellForm = "NULL"
					if sStage = "-" then
						sStage = "NULL"
					else
						sStage = Pack(sStage)
					end if

					if iRate = "0" then iRate = "NULL"

					iSerialQty = cdbl(iNett) - cdbl(iQtyIssued)

					' if for Serial Quantity > 0
					if not iSerialQty = 0 then
						iValue = cdbl(iValue) + (iSerialQty * cdbl(iRate))

						sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,LOTNUMBER,SERIALNUMBER," &_
							"QUANTITY) VALUES " &_
							"(" & iTraNo & "," & sPickLot & "," & iSerialNo & "," & iSerialQty & ")"
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql

						with dcrs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LOCATIONLOT"
							.ActiveConnection = con
							.Open
						end with
						set dcrs2.ActiveConnection = nothing

						if not dcrs2.EOF then
							iNewSerNo = dcrs2(0)
							sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
								"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
								"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE) VALUES " &_
								"(" & iInvRecNewNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
								"" & sToLoc & "," & sToBin & "," & sPickLot & "," & iNewSerNo & "," &_
								"" & (cdbl(iSerialQty) + cdbl(iTare)) & "," & iSerialQty & "," & cdbl(iTare) & "," & sPackNum & "," &_
								"" & iPackCode & "," & iSellNum & "," & iWtPerSellFm & "," & iSellForm & "," & sStage & "," & iRate & ")"
							'Response.Write sSql & vbCrLf & vbCrLf
							con.Execute sSql

						end if
						dcrs2.Close

						sSql = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iSerialQty & ")" &_
							" WHERE SERIALNUMBER = " & iSerialNo & ""
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql

					' end if for Serial Quantity > 0
					end if
				dcrs1.MoveNext
				loop
			end if
			dcrs1.Close

		' If Lot Exists then check in transfer table for which serial number and Update the corresponding
		else
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				if sPickLot = "NULL" then
					.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),ISNULL(PACKINGNUMBER,'-'),ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER NOT IN (SELECT SERIALNO FROM INV_T_MRSISSUEPICKSERIAL WHERE SERIALNO IS NOT NULL) ORDER BY 1"
				else
					.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),ISNULL(PACKINGNUMBER,'-'),ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & sPickLot & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER NOT IN (SELECT SERIALNO FROM INV_T_MRSISSUEPICKSERIAL WHERE SERIALNO IS NOT NULL) ORDER BY 1"
				end if
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				Do While Not dcrs1.EOF

					iSerialNo = trim(dcrs1(0))
					iNett = trim(dcrs1(1))
					iQtyIssued = trim(dcrs1(2))
					sPackNum = trim(dcrs1(3))
					iRate = trim(dcrs1(4))

					if sPackNum = "-" then
						sPackNum = "NULL"
					else
						sPackNum = Pack(sPackNum)
					end if

					iSerialQty = cdbl(iNett) - cdbl(iQtyIssued)

					' if for Serial Quantity > 0
					if not iSerialQty = 0 then
						iValue = cdbl(iValue) + (iSerialQty * cdbl(iRate))

						sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,LOTNUMBER,SERIALNUMBER," &_
							"QUANTITY) VALUES " &_
							"(" & iTraNo & "," & sPickLot & "," & iSerialNo & "," & iSerialQty & ")"
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql

						sSql = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iSerialQty & ")" &_
							" WHERE SERIALNUMBER = " & iSerialNo & ""
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql

						with dcrs2
							.CursorLocation = 3
							.CursorType = 3
							if sPickLot = "NULL" then
								.Source = "SELECT SERIALNUMBER,INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE PACKINGNUMBER = " & sPackNum & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL) AND SERIALNUMBER <> " & iSerialNo & ""
							else
								.Source = "SELECT SERIALNUMBER,INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE PACKINGNUMBER = " & sPackNum & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & sPickLot & " AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL) AND SERIALNUMBER <> " & iSerialNo & ""
							end if
							.ActiveConnection = con
							.Open
						end with
						set dcrs2.ActiveConnection = nothing

						if not dcrs2.EOF then
							iSerialNo = trim(dcrs2(0))
							iInvRecNewNo = trim(dcrs2(1))

							' if for Serial Quantity > 0
							if not iSerialQty = 0 then

								sSql = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) - " & iSerialQty & ")" &_
									" WHERE SERIALNUMBER = " & iSerialNo & ""
								'Response.Write sSql & vbCrLf & vbCrLf
								con.Execute sSql

							' end if for Serial Quantity > 0
							end if
						end if
						dcrs2.Close

					' end if for Serial Quantity > 0
					end if
				dcrs1.MoveNext
				loop
			end if
			dcrs1.Close

		end if
		dcrs.close
	end if

	sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
		"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
		"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
		"'IT'," & iTraNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iIssQty & "," & iValue & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if dcrs.EOF then
		sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
			"LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
			"" & sLoc & "," & sBin & "," & iIssQty & "," & iValue & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	else
		sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
			"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	end if
	dcrs.Close
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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

	sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
		"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
		"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
		"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
		"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
		"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
		"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	sSql = "UPDATE Inv_T_ItemLocationStock SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
		"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
		"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
		"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
		"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
		"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
		"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
		"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql
	
	LotFullInsert = iInvRecNewNo&":"&iValue
end function
'''***************************************************************************''''''''

' Function to Insert Marked and Picked Serial Values
Function LotPartialInsert(bCheck,iIssQtyP,iValueP,iSerialNoP)

	' Declaration of variables
	Dim dcrs,dcrs1,dcrs2,iInvRecNewNo,iValue,iSerialNo,iSerialQty,iIssQty,iNewSerNo,iRate
	dim iGross,iNett,iTare,iQtyIssued,sPackNum,iPackCode,iSellNum,iWtPerSellFm
	dim iSellForm,sStage,iAltGross,iAltNett
	iInvRecNewNo = "0"
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

	iValue = iValueP
	iSerialNo = iSerialNoP
	iIssQty = iIssQtyP

	' to insert only details
	if bCheck = "D" then

		iSerialQty = cdbl(iIssQtyP)
		iInvRecNewNo = iValueP

		sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,LOTNUMBER,SERIALNUMBER," &_
			"QUANTITY) VALUES " &_
			"(" & iTraNo & "," & sPickLot & "," & iSerialNo & "," & iSerialQty & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		if iInvRecNewNo <> "0" then
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYGROSS,0),ISNULL(LOTQUANTITYNETT,0),ISNULL(LOTQUANTITYTARE,0),ISNULL(QUANTITYISSUED,0),ISNULL(PACKINGNUMBER,'-'),ISNULL(PACKINGCODE,0),ISNULL(SELLINGNUMBER,0),ISNULL(WEIGHTPERSELLINGFORM,0),ISNULL(SELLINGFORM,0),ISNULL(STAGE,'-'),ISNULL(ALTGROSS,0),ISNULL(ALTNETT,0),ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE SERIALNUMBER = " & iSerialNo & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				iGross = trim(dcrs1(1))
				iNett = trim(dcrs1(2))
				iTare = trim(dcrs1(3))
				iQtyIssued = trim(dcrs1(4))
				sPackNum = trim(dcrs1(5))
				iPackCode = trim(dcrs1(6))
				iSellNum = trim(dcrs1(7))
				iWtPerSellFm = trim(dcrs1(8))
				iSellForm = trim(dcrs1(9))
				sStage = trim(dcrs1(10))
				iAltGross = trim(dcrs1(11))
				iAltNett = trim(dcrs1(12))
				iRate = trim(dcrs1(13))

				if sPackNum = "-" then
					sPackNum = "NULL"
				else
					sPackNum = Pack(sPackNum)
				end if
				if iPackCode = "0" then iPackCode = "NULL"
				if iSellNum = "0" then iSellNum = "NULL"
				if iWtPerSellFm = "0" then iWtPerSellFm = "NULL"
				if iSellForm = "0" then iSellForm = "NULL"
				if sStage = "-" then
					sStage = "NULL"
				else
					sStage = Pack(sStage)
				end if

				if iRate = "0" then iRate = "NULL"
			end if
			dcrs1.close

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LOCATIONLOT"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iNewSerNo = dcrs(0)
				sSql = "INSERT INTO	INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
					"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
					"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE) VALUES " &_
					"(" & iInvRecNewNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
					"" & sToLoc & "," & sToBin & "," & sPickLot & "," & iNewSerNo & "," &_
					"" & (cdbl(iSerialQty) + cdbl(iTare)) & "," & iSerialQty & "," & cdbl(iTare) & "," & sPackNum & "," &_
					"" & iPackCode & "," & iSellNum & "," & iWtPerSellFm & "," & iSellForm & "," & sStage & "," & iRate & ")"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql

			end if
			dcrs.Close

			sSql = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iSerialQty & ")" &_
				" WHERE SERIALNUMBER = " & iSerialNo & ""
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),ISNULL(PACKINGNUMBER,'-'),ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE SERIALNUMBER = " & iSerialNo & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				Do While Not dcrs1.EOF

					iSerialNo = trim(dcrs1(0))
					iNett = trim(dcrs1(1))
					iQtyIssued = trim(dcrs1(2))
					sPackNum = trim(dcrs1(3))
					iRate = trim(dcrs1(4))

					if sPackNum = "-" then
						sPackNum = "NULL"
					else
						sPackNum = Pack(sPackNum)
					end if

					iSerialQty = cdbl(iNett) - cdbl(iQtyIssued)

					' if for Serial Quantity > 0
					if not iSerialQty = 0 then
						iValue = cdbl(iValue) + (iSerialQty * cdbl(iRate))

						sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,LOTNUMBER,SERIALNUMBER," &_
							"QUANTITY) VALUES " &_
							"(" & iTraNo & "," & sPickLot & "," & iSerialNo & "," & iSerialQty & ")"
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql

						sSql = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iSerialQty & ")" &_
							" WHERE SERIALNUMBER = " & iSerialNo & ""
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql

						with dcrs2
							.CursorLocation = 3
							.CursorType = 3
							if sPickLot = "NULL" then
								.Source = "SELECT SERIALNUMBER,INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE PACKINGNUMBER = " & sPackNum & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL) AND SERIALNUMBER <> " & iSerialNo & ""
							else
								.Source = "SELECT SERIALNUMBER,INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE PACKINGNUMBER = " & sPackNum & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & sPickLot & " AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL) AND SERIALNUMBER <> " & iSerialNo & ""
							end if
							.ActiveConnection = con
							.Open
						end with
						set dcrs2.ActiveConnection = nothing

						if not dcrs2.EOF then
							iSerialNo = trim(dcrs2(0))
							iInvRecNewNo = trim(dcrs2(1))

							' if for Serial Quantity > 0
							if not iSerialQty = 0 then

								sSql = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) - " & iSerialQty & ")" &_
									" WHERE SERIALNUMBER = " & iSerialNo & ""
								'Response.Write sSql & vbCrLf & vbCrLf
								con.Execute sSql

							' end if for Serial Quantity > 0
							end if
						end if
						dcrs2.Close

					' end if for Serial Quantity > 0
					end if
				dcrs1.MoveNext
				loop
			end if
			dcrs1.Close

		end if
	' to insert only header values
	elseif bCheck = "H" then

		sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFER (TRANSFERNO,ITEMCODE,ORGANISATIONCODE," &_
			"CLASSIFICATIONCODE,FROMLOCATIONNUMBER,TOLOCATIONNUMBER," &_
			"TOBINNUMBER,FROMBINNUMBER,TRANSFERTYPE,TRANSFERREDON,TRANSFERREDBY) VALUES " &_
			"(" & iTraNo & "," & iItemCode & "," & Pack(sOrgID) & "," & iClass & "," & sLoc & "," &_
			"" & sToLoc & "," & sToBin & "," &_
			"" & sBin & ",'ST',CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		'if sPickLot <> "NULL" then
		if 1 = 1 then
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				if sPickLot <> "NULL" then
					.Source = "SELECT INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & sPickLot & " AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL)"
				else
					.Source = "SELECT INVENTORYRECEIPTNO FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sToLoc & " AND (STORAGEBINNUMBER = " & sToBin & " OR STORAGEBINNUMBER IS NULL)"
				end if
				'Response.Write dcrs.source & vbCrLf & vbCrLf
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			' Check whether the TO Storage Location has the Lot Number, if not then insert
			if dcrs1.EOF then

				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing

				if not dcrs.EOF then
					iInvRecNewNo = dcrs(0)
				end if
				dcrs.close

				sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,RECEIPTNUMBER,ORGANISATIONCODE," &_
					"ITEMCODE,CLASSIFICATIONCODE,RECEIPTQUANTITY,ITEMRATE,SRCTYPE,ACCOUNTEDON,ACCOUNTEDBY,RECEIVEDON) VALUES " &_
					"(" & iInvRecNewNo & "," & iTraNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
					"" & iIssQty & "," & iItmRate & ",'RT'," &_
					"CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103))"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs1.close
		end if

		sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
			"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
			"'IT'," & iTraNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iIssQty & "," & iValue & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
				"" & sLoc & "," & sBin & "," & iIssQty & "," & iValue & ")"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
				"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
				
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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

		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE Inv_T_ItemLocationStock SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		LotPartialInsert = iInvRecNewNo
	end if
	' end if for Check
end function
'''***************************************************************************''''''''

' Function to Insert in case of Clean Stock Transfer
Function StockInsert()
' Declaration of variables
	Dim dcrs,dcrs1,iValue
	dim iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue,iRecQty,iQtyRec

	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(YEARRECEIPTQUANTITY,0),YEAROPENINGSTOCK,YEARISSUEQUANTITY,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iRecQty = cdbl(dcrs(0))
		iYrOpStock = cdbl(dcrs(1))
		iYrIssQty = cdbl(dcrs(2))
		iYrCloQty = cdbl(dcrs(3))
		iYrCloValue = cdbl(dcrs(4))
	end if
	dcrs.Close

	if iYrCloQty > 0 then

		iValue = iIssQty * (iYrCloValue / iYrCloQty)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFER (TRANSFERNO,ITEMCODE,ORGANISATIONCODE," &_
			"CLASSIFICATIONCODE,FROMLOCATIONNUMBER,TOLOCATIONNUMBER," &_
			"TOBINNUMBER,FROMBINNUMBER,TRANSFERTYPE,TRANSFERREDON,TRANSFERREDBY) VALUES " &_
			"(" & iTraNo & "," & iItemCode & "," & Pack(sOrgID) & "," & iClass & "," & sLoc & "," &_
			"" & sToLoc & "," & sToBin & "," &_
			"" & sBin & ",'ST',CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,QUANTITY) VALUES " &_
			"(" & iTraNo & "," & iIssQty & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
			"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
			"'IT'," & iTraNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iIssQty & "," & iValue & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
				"" & sLoc & "," & sBin & "," & iIssQty & "," & iValue & ")"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
				"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
				
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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

		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE Inv_T_ItemLocationStock SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		iQtyRec = iIssQty

		sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
			"TRANSACTIONTYPE,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
			"'RT',CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iQtyRec & "," & iValue & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
				"" & sToLoc & "," & sToBin & "," & iQtyRec & "," & iValue & ")"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
				"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL) "
				
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
			"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO Inv_T_ItemLocationStock (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
				"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
				"YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
				"" & sToLoc & "," & sToBin & "," &_
				"" & iQtyRec & "," & iValue & "," & iQtyRec & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		
			If 1 = 2 Then
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(STOCKNO)+1,1) FROM INV_M_STOCKSTATUS"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				sSql = "INSERT INTO INV_M_STOCKSTATUS (STOCKNO,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER) VALUES " &_
					"(" & trim(dcrs1(0)) & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
					"" & sToLoc & "," & sToBin & ")"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs1.Close
			End  If ' If 1 = 2 Then
		else
			sSql = "UPDATE Inv_T_ItemLocationStock SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
				"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
				"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
				"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
				"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
				"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL) AND " &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
				"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close
		
		
	end if

end function

' Function to Insert in case of Full Lot / Serial Selection [TO STORE]
Function ReceiptFullInsert(sPara)
	' Declaration of variables
	Dim dcrs,dcrs1,iInvRecNew,arrTemp,iTValue
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

	arrTemp = split(sPara,":")
	iInvRecNew = arrTemp(0)
	iTValue = arrTemp(1)

	sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
		"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
		"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
		"'RT'," & iInvRecNew & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iIssQty & "," & iTValue & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL) "
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if dcrs.EOF then
		sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
			"LOCATIONNUMBER,BINNUMBER,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
			"" & sToLoc & "," & sToBin & "," & Pack(sMonYr) & "," & iIssQty & "," & iTValue	& ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	else
		sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iIssQty & ")," &_
			"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iTValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
			"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL)"
			
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	end if
	dcrs.Close
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iIssQty & ")," &_
		"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iTValue & ")," &_
		"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iIssQty & "), " &_
		"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iTValue & ") WHERE " &_
		"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
		"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
		"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing

	if dcrs.EOF then
		sSql = "INSERT INTO Inv_T_ItemLocationStock (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
			"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
			"YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
			"" & sToLoc & "," & sToBin & "," &_
			"" & iIssQty & "," & iTValue & "," & iIssQty & "," & iTValue & ")"
		'objTxt.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
		
		If 1 = 2 Then
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(STOCKNO)+1,1) FROM INV_M_STOCKSTATUS"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				sSql = "INSERT INTO INV_M_STOCKSTATUS (STOCKNO,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER) VALUES " &_
					"(" & trim(dcrs1(0)) & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
					"" & sToLoc & "," & sToBin & ")"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs1.Close
		End IF
	else
		sSql = "UPDATE Inv_T_ItemLocationStock SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iIssQty & ")," &_
			"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iTValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iIssQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iTValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sToLoc & " AND (BINNUMBER = " & sToBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	end if
	dcrs.Close

end function
%>

