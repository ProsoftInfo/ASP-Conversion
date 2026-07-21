<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	JobWorkReceiptInsert.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	TAJUDEEN S
	'Created On					:	September 16, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	JobWorkReceiptEntry.asp
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
<!--#include file="../../include/NoSeries.asp"-->
<!--#include file="../../include/PurchaseInvFnc.asp"-->
<!--#include file="../../include/mrsStatus.asp"-->

<%
	'XML DOM Variables
	Dim objfs,newxml,RootNode,HeaderNode,PageNode,SchNode,StorageNode
	dim dcrs,dcrs1,dcrs2,sSql,adoCmd
	dim iItemCode,iClass,arrStore,sBin,sLoc,sOrgCode,iSrcRefNO,iReceiptAgainst
	dim iRecNo,iInvRecNo,iQtyOrd,iQtyRec,iQtyIns,iQtyAcc,iQtyRej,iItmRate,iValue
	dim sPressFrom,sPressTo,sQtyIn,iTareQty,sPress,sCrop,iBales,iLot,iSerialFrom,iSerialTo
	dim arrLotSerial,iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sTareIn,iQtyGross,iQtyNett
	dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,bLotFlag,sRecType,sLotNumber,sPackingNo
	dim sSellingType,iWeight,sPackingType,sNoofPacks
	dim sItmType,iSeriesNo,iSeriesCode,sRefType,sSellingFormType,sStage
	dim iProcessNo,iOrderConfNo,sValue,iAccountedBy,dReceivedOn,dtFinFrom,dtFinTo

	Dim Node,subNode,selNode
	' Create our DOM Document Objects
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

	newxml.async = false
	newxml.load(Request)
	'newxml.Save Server.MapPath("../Temp/Transaction/JobWorkReceipt.xml")

	bLotFlag = false
	iAccountedBy = getUserid
	sRefType = "'RW'"	'Receipt For Job Work
	
	if len(Month(date())) = 1 then
		sTempMonYr = "0"&Month(date())
	else
		sTempMonYr = Month(date())
	end if
	sMonYr = sTempMonYr&Year(date())
	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	
'	Response.Write newxml.xml
dtFinFrom =  sFinFrom
dtFinTo =  sFinTo

	con.beginTrans

	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LocationLot"
'		Response.Write DCRS1.Source 
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing

	if not dcrs1.EOF then
		iInvRecNo = dcrs1(0)
		Response.Write "Y<BR>" ' pls do not block this line , this is referred in calling prg.

		Set RootNode = newxml.documentElement
		iRecNo = trim(RootNode.Attributes.getNamedItem("RECNO").Value)
		dReceivedOn = trim(RootNode.Attributes.getNamedItem("RECEIVEDON").Value)
		iSrcRefNO = trim(RootNode.Attributes.getNamedItem("SrcRefNO").Value)
		iReceiptAgainst = trim(RootNode.Attributes.getNamedItem("ReceiptAgainst").Value)

		'Response.Write "<p> " & iSrcRefNO
		'Fetching Order Confirmation No
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT SOURCEREFNO FROM SUB_T_SALEARHEADER WHERE ACTUALRECEIPTNO = " & iRecNo
			'Response.Write DCRS2.Source 
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing

		if not dcrs2.EOF then
			iOrderConfNo = dcrs2(0)
		else
			iOrderConfNo = "NULL"
		end if
		dcrs2.Close

		'Fetching Processing No
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PROCESSINGNO FROM SAL_T_PROCESSEDOC WHERE ORDERCONFIRMATIONNO = " & iOrderConfNo
			'Response.Write "<P>"& DCRS2.Source 
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing

		if not dcrs2.EOF then
			iProcessNo = dcrs2(0)
		else
			iProcessNo = "NULL"
		end if
		dcrs2.Close

		if RootNode.HaschildNodes() then
			For Each HeaderNode In RootNode.childNodes
				iItemCode = trim(HeaderNode.Attributes.getNamedItem("ICODE").Value)
				iClass = trim(HeaderNode.Attributes.getNamedItem("CCODE").Value)
				sOrgCode = trim(HeaderNode.Attributes.getNamedItem("OCODE").Value)

				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ITEMTYPEID FROM VWITEM WHERE ClassificationCode = " & iClass & " AND ITEMCODE = " & iItemCode & ""
					'Response.Write "<P>"& DCRS.Source 
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing

				if not dcrs.EOF then
					sItmType = trim(dcrs(0))
				end if
				dcrs.Close

				iQtyOrd = 0
				iQtyRec = 0
				iQtyIns = 0
				iQtyAcc = 0
				iQtyRej = 0
				iValue = 0
				iItmRate = 0

				iQtyOrd = trim(HeaderNode.Attributes.getNamedItem("QTY").Value)
				iQtyRec = trim(HeaderNode.Attributes.getNamedItem("QTYREC").Value)
				iQtyAcc = trim(HeaderNode.Attributes.getNamedItem("QTYREC").Value)
				iItmRate = trim(HeaderNode.Attributes.getNamedItem("ITMRATE").Value)
				sValue = trim(HeaderNode.Attributes.getNamedItem("VALUETYPE").Value)

				sPressFrom = trim(HeaderNode.Attributes.getNamedItem("PRUNFROM").Value)
				sPressTo = trim(HeaderNode.Attributes.getNamedItem("PRUNTO").Value)
				sQtyIn = trim(HeaderNode.Attributes.getNamedItem("QTYIN").Value)
				sTareIn = trim(HeaderNode.Attributes.getNamedItem("TARE").Value)
				sPress = trim(HeaderNode.Attributes.getNamedItem("PRESS").Value)
				sCrop = trim(HeaderNode.Attributes.getNamedItem("CROP").Value)
				iBales = trim(HeaderNode.Attributes.getNamedItem("BALES").Value)
				iTareQty = trim(HeaderNode.Attributes.getNamedItem("TAREWEIGHT").Value)

				if iQtyAcc = "Nil" then
					iQtyAcc = "0"
				end if

				if iQtyRej = "Nil" then
					iQtyRej = "0"
				end if

				if Trim(iItmRate) = "" then
					iItmRate = 0
					iValue = 0
				else
					iValue = cdbl(iItmRate) * cdbl(iQtyAcc)
				end if

				if iTareQty = "" or IsNull(iTareQty) or IsEmpty(iTareQty) then
					iTareQty = "0"
				end if

				if sPressFrom = "" or IsNull(sPressFrom) or IsEmpty(sPressFrom) then
					sPressFrom = "NULL"
				else
					sPressFrom = Pack(sPressFrom)
				end if

				if sPressTo = "" or IsNull(sPressTo) or IsEmpty(sPressTo) then
					sPressTo = "NULL"
				else
					sPressTo = Pack(sPressTo)
				end if

				if sPress = "" or IsNull(sPress) or IsEmpty(sPress) then
					sPress = "NULL"
				else
					sPress = Pack(sPress)
				end if

				if sCrop = "" or IsNull(sCrop) or IsEmpty(sCrop) then
					sCrop = "NULL"
				else
					sCrop = Pack(sCrop)
				end if

				if iBales = "" or IsNull(iBales) or IsEmpty(iBales) then
					iBales = "NULL"
				end if

				if lcase(sQtyIn) = "select" or IsNull(sQtyIn) or sQtyIn = "" then
					sQtyIn = "NULL"
				else
					sQtyIn = Pack(sQtyIn)
				end if

				if lcase(sTareIn) = "select" or IsNull(sTareIn) or sTareIn = "" then
					sTareIn = "NULL"
				else
					sTareIn = Pack(sTareIn)
				end if
				''Added By Ragavendran on Dec 05,2009 for Fibre Items 
					sNoofPacks = 0
					
					For each Node in HeaderNode.childNodes
						if strcomp(Node.nodeName,"STORAGE") =0 then
							For each subNode in Node.childNodes
								if strcomp(subNode.nodeName,"LotSerial")=0 then
									For each selNode in subNode.childNodes
										if	strcomp(selNode.nodeName,"LotSerialDetails")= 0 then
											'Response.Write "QTYGRO="&selNode.getAttribute("QTYGRO")
											if cdbl(selNode.getAttribute("QTYGRO")) > 0 then
												sNoofPacks = sNoofPacks +1
											end if
										end if
									next
								end if
							next
						end if
					next
				
					'Response.Write "sNoofPacks="& sNoofPacks

				sSql = "UPDATE SUB_T_SALEARDETAILS SET ITEMSTATUS = 'Y' WHERE FROMITEMCODE = " & iItemCode &_
					" AND FROMCLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) &_
					" AND ACTUALRECEIPTNO = " & iRecNo & ""
			'	Response.Write sSql & "<BR><BR>"
				con.Execute sSql
				
				sSql = "Update PUR_T_RefferenceNumberDet set InventoryRcptNo ="&iInvRecNo&" where ReceiptNumber="&iRecNo&""
				'Response.Write "<P>"&sSql
				con.execute sSql


				iQtyRec = iQtyAcc

				sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
					"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NoOfPacks) VALUES " &_
					"(" & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
					"'RW'," & iInvRecNo & ",CONVERT(DATETIME," & Pack(dReceivedOn) & ",103)," & iQtyRec & "," & iValue & ","& sNoofPacks&")"
				
			'	Response.Write sSql & "<BR><BR>"
				con.Execute sSql

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				For Each StorageNode In HeaderNode.childNodes
					sLoc = trim(StorageNode.Attributes.getNamedItem("STORE").Value)
					sBin = trim(StorageNode.Attributes.getNamedItem("BIN").Value)
					iQtyRec = trim(StorageNode.Attributes.getNamedItem("QTY").Value)
					sRecType = trim(StorageNode.Attributes.getNamedItem("RECTYPE").Value)

					if sBin = "0" then sBin = "NULL"

					if cdbl(iQtyRec) <> 0 then
						iValue = cdbl(iItmRate) * cdbl(iQtyRec)

						' In case of Receipt Numbering is None
						if sRecType = "N" then
							sLotNumber = "NULL"
							sPackingNo = "NULL"
							iSerialEntry = "NULL"

							'if sQtyIn = "N" then
							'	iQtyGross = cdbl(iQtyRec) + cdbl(iTareQty)
							'	iQtyNett = iQtyRec
							'else
							'	iQtyGross = iQtyRec
							'	iQtyNett = cdbl(iQtyRec) - cdbl(iTareQty)
							'end if
							iQtyNett = iQtyRec
							iQtyGross = iQtyRec

							sSql = "INSERT INTO INV_T_LocationLot (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
								"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER," &_
								"LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,SRCTYPE,DateOfReceipt) VALUES " &_
								"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
								"" & sLoc & "," & sBin & "," & sLotNumber & "," & iSerialEntry & "," &_
								"" & iQtyGross & "," & iQtyNett & "," & iTareQty & "," & sPackingNo & ","&sRefType & ",CONVERT(DATETIME," & Pack(dReceivedOn) & ",103))"
							Response.Write sSql & "<BR><BR>"
							con.Execute sSql
						else

							' In case of Receipt Numbering is (LOT and SERIAL) or SERIAL
							For Each PageNode In StorageNode.childNodes
								if StrComp(PageNode.nodeName,"LotSerial") = 0 then
									if sRecType = "S" then
										iLot = "NULL"
									else
										iLot = Pack(trim(PageNode.Attributes.getNamedItem("LOT").Value))
										bLotFlag = true
									end if

									For Each SchNode In PageNode.childNodes
										if StrComp(SchNode.nodeName,"LotSerialDetails") = 0 then
											bLotFlag = true
											iLotEntry = iLot
											'iSerialEntry = trim(SchNode.Attributes.getNamedItem("LOTSERIAL").Value)
											iQtyRecEntry = trim(SchNode.Attributes.getNamedItem("QTYREC").Value)
											iTareQtyEntry = trim(SchNode.Attributes.getNamedItem("TAREREC").Value)

											sSellingType = trim(SchNode.Attributes.getNamedItem("SELLINGTYPE").Value)
											iWeight = trim(SchNode.Attributes.getNamedItem("WEIGHTSTYPE").Value)
											sPackingType = trim(SchNode.Attributes.getNamedItem("PACKINGTYPE").Value)
											sSellingFormType = trim(SchNode.Attributes.getNamedItem("SELLINGFORM").Value)

											sPackingNo = trim(SchNode.Attributes.getNamedItem("PACKNUMBER").Value)
											iQtyGross = trim(SchNode.Attributes.getNamedItem("QTYGRO").Value)

											if sSellingType = "select" then sSellingType = "NULL"
											if iWeight = "" then iWeight = "NULL"
											if sPackingType = "select" then sPackingType = "NULL"
											if sPackingNo = "" or sPackingNo = "NULL" then
												sPackingNo = "NULL"
											else
												sPackingNo = Pack(sPackingNo)
											end if

											iQtyNett = iQtyRecEntry
											'if sQtyIn = "N" then
											'	iQtyGross = cdbl(iQtyRecEntry) + cdbl(iTareQtyEntry)
											'	iQtyNett = iQtyRecEntry
											'else
											'	iQtyGross = iQtyRecEntry
											'	iQtyNett = cdbl(iQtyRecEntry) - cdbl(iTareQtyEntry)
											'end if

											sStage = "NULL"
											'if sSellingType <> "NULL" then
											'	if sSellingType = sSellingFormType then
											'		sStage = "'F'"
											'	else
											'		sStage = "'SF'"
											'	end if
											'end if

											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LocationLot"
												'Response.Write "<P>"&dcrs.Source 
												.ActiveConnection = con
												.Open
											end with
											set dcrs.ActiveConnection = nothing

											if not dcrs.EOF then
												iSerialEntry = dcrs(0)
											end if
											dcrs.Close

											sSql = "INSERT INTO INV_T_LocationLot (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
												"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
												"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,SRCTYPE,DateofReceipt) VALUES " &_
												"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
												"" & sLoc & "," & sBin & "," & iLotEntry & "," & iSerialEntry & "," &_
												"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & "," & sPackingNo & "," &_
												"" & sPackingType & "," & sSellingType & "," & iWeight & "," & sSellingFormType & "," & sStage & ","&sRefType &",CONVERT(DATETIME," & Pack(dReceivedOn) & ",103))"
											'Response.Write sSql & "<BR><BR>"
											con.Execute sSql

										end if
									next
								end if
							next
						end if

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						arrFin = split(GetFinancialYear(sMonYr),":")
						sFinFrom = arrFin(0)
						sFinTo = arrFin(1)

						with dcrs
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(dtFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(dtFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
							'Response.Write "<P>"&DCRS.Source 
							.ActiveConnection = con
							.Open
						end with
						set dcrs.ActiveConnection = nothing
						if dcrs.EOF then
							sSql = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
								"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
								"(" & Pack(sOrgCode) & "," & iClass & "," & iItemCode & "," &_
								"CONVERT(DATETIME," & Pack(dtFinFrom) & ",103),CONVERT(DATETIME," & Pack(dtFinTo) & ",103)," & iQtyRec & "," & iValue & "," & iQtyRec & "," & iValue & ")"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql
						else
							sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
								"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
								"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
								"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
								"CONVERT(DATETIME," & Pack(dtFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(dtFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql
						end if
						dcrs.Close

						with dcrs
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(dtFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(dtFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
							'Response.Write "<P>"& DCRS.Source 
							.ActiveConnection = con
							.Open
						end with

						set dcrs.ActiveConnection = nothing

						if dcrs.EOF then
							sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
								"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
								"YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
								"(" & Pack(sOrgCode) & "," & iClass & "," & iItemCode & "," &_
								"CONVERT(DATETIME," & Pack(dtFinFrom) & ",103),CONVERT(DATETIME," & Pack(dtFinTo) & ",103)," &_
								"" & sLoc & "," & sBin & "," &_
								"" & iQtyRec & "," & iValue & "," & iQtyRec & "," & iValue & ")"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql
						else
							sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
								"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
								"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
								"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
								"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
								"CONVERT(DATETIME," & Pack(dtFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
								"CONVERT(DATETIME," & Pack(dtFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							'Response.Write sSql & "<BR><BR>"
							con.Execute sSql
						end if
						dcrs.Close


			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
					end if
				next
			next

			' Check for Lot entry and Update the Number Series
			if bLotFlag then
				' Lot Number - Number Series
				' OrganisationCode,SeriesNumber,SeriesCode,Date
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'LO' AND ITEMTYPE = " & Pack(sItmType) & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
					'Response.Write "<P>"& DCRS.Source 
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				if not dcrs.EOF then
					iSeriesNo = trim(dcrs(0))
					iSeriesCode = trim(dcrs(1))
				end if
				dcrs.close

				'GenSeriesNumber sOrgCode,iSeriesNo,iSeriesCode,FormatDate(date())
				'Response.Write "OK"
			end if
		end if
	end if
	dcrs1.Close
	
	'added on Nov 17,2009
	'Calling Store Procedure to Job Order
	if trim(iSrcRefNO) <> "" then
		Set adoCmd = Server.CreateObject("ADODB.Command")

		Set adoCmd.ActiveConnection = con

		adoCmd.CommandText = "JWK_CreateProductionOrder"
		adoCmd.CommandType = 4 'adCmdStoredProc
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@iJobOrderNo", 129,1,250, cint(iSrcRefNO))

		adoCmd.Execute()

	end if 
	
	
	
	if con.Errors.count <> 0 then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
		'Response.Write "ffff"
	'	con.RollbackTrans
	'	Response.End 
		con.CommitTrans

		if objfs.FileExists(server.MapPath("../temp/transaction/JOBWORKRECEIPTEX"&Session.SessionID&".xml")) then
			objfs.DeleteFile server.MapPath("../temp/transaction/JOBWORKRECEIPTEX"&Session.SessionID&".xml")
		end if

	end if

	con.close
	set con = nothing
%>
