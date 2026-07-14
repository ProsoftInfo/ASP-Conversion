<%
	Function checkLeapYear(iYear)
		If iYear Mod 4 = 0 Then
			If iYear Mod 100 = 0 Then
				If iYear Mod 400 = 0 Then
					checkLeapYear = True
				Else
					checkLeapYear = False
				End If
			Else
				checkLeapYear = True
			End If
		Else
			checkLeapYear = False
		End If
	End Function
	
	Function InvItemCreationInsert()
	
		dim dcrs,dcrs1,dcrs2,sSql,oDOM,objfs,RootNode,ItemNode,Node,UnitNode,StorageNode,LotNode
		dim newxml,LotSerialNode,AssetNode,BoMNode,FabNode,FabDetNode,ClassNode,cSubNode
		dim iItmCode,sitmShDesc,sitmDesc,sitmAddDesc,sitmUsage,iitmController,sStUoM,sPuUoM
		dim sMaUoM,sSaUoM,iStToPur,iStToMan,iStToSal,sStToPurOp,sStToManOp,sStToSalOp,sitmType
		dim sFormCode,sItmWho,sItmActive,sItmStock,sSalEli,sManEli,sPurEli
		dim iCodeLen,iCodeSize,sDrwVerNo,sCatalogue,sExp,sExp1,iClass,sParentCode,sitmCode
		dim iBoMClassCode,iCtrBoM,iBoMItemCode,iBoMQty,sBoMUoM,sBoMType,sBoMUnit
		dim sWeave,iWidth,iReedCount,sReedSpace,sVariety,iNODent,iNOTotal,iNOEnds
		dim iNOPicks,sAverageWC,iTapeLength,sWarpYarn,sWeftYarn,iWarpCount,iWeftCount,iCtr
		dim arrWarpYarn,arrWeftYarn,i,j,sGroupPath,sDateOfRec,iStoEntNo,iItemEntNo
		Dim ItmNode,iItmCtr,iBomUOM,iBomType,iBomConsumable

		dim iVar,jVar,isize,iCount,sTempitmCode

		dim sVariantCode,sOrgCode,sCategory,sOpenStockOrgCode
		dim sRecNum,sRecRout,sAccType,sABC,sModvat,sAPur,sASubCon,sASales
		dim sTemp,sLoc,sBin,sMonYr,iQty,iValue,sAppli
		dim arrFin,arrFin1,sFinFrom,sFinTo,dFinFrom,dFinTo,arrTemp,iInvRecNo,sTempMonYr
		dim sPurUoM,sManUoM,sSalUoM,iPurRate,iManRate,iSalRate,sUoMDesc,sOpUoMFor
		dim sQtyIn,iTareQty,iLot,iSerialFrom,iSerialTo,sSellingType,iWeight,sPackingType,sPackingNumber
		dim iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sTareIn,iQtyGross,iQtyNett
		dim iSeriesNo,iSeriesCode,sProductCode,iItmRate,iCounter
		dim iLotCtr,sSellingFormType,sStage,sForm,arrAttr,sAttr,sBarCode
		dim iLotQty,iLotValue,iEntryCtr,sReceivedOn,iReOrderLvl,iReOrderQty,iEcoOrderQty
		dim iFabCtr,iPieceNo,iPieceQty,arrFab,iAltGross,iAltNett,sAltUoM,sCheck,iRate
		dim sOpUoMCode,iOpUoMFactor,sOpUoMOperator,sAttrFlag,iBOMApplicable
		dim sTempType,sTempCat,iParentCode,arrParent,arrGroup,sAttribute,arrLevel,iLevel,iGroup
		dim objFSO,objTxt
		Dim lastDateStr,lastDateArray
	
		
		Dim HeaderNode,newElem,newElem1,newElem2,StockNode,iItemTotalQty,iItemTotalValue

		set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		iCodeLen = 0
		iCodeSize = 0

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
		Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

		set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		set objfs = Server.CreateObject("Scripting.FileSystemObject")

		'Create our DOM Document Objects
		Set newxml = Server.CreateObject("Microsoft.XMLDOM")

		iitmController = getUserid
		iItemEntNo = 1
		'newxml.async = false
		'newxml.load(Request)
		'Response.Write newxml.xml
		'newxml.save Server.MapPath("../temp/Master/ItemCreate.xml")
		
		newxml.Load Server.MapPath("../temp/Master/AssetItem.xml")
		Set RootNode = newxml.documentElement

		sExp ="//Output/DETAILS"
		Set ItemNode = RootNode.selectSingleNode(sExp)
		sitmType = trim(ItemNode.Attributes.getNamedItem("ITYPE").Value)
		sitmCode = trim(ItemNode.Attributes.getNamedItem("ICODE").Value)
		sitmShDesc = trim(ItemNode.Attributes.getNamedItem("SHDESC").Value)
		sitmDesc = trim(ItemNode.Attributes.getNamedItem("DESC").Value)
		sitmAddDesc = trim(ItemNode.Attributes.getNamedItem("ADDDESC").Value)
		sDrwVerNo = trim(ItemNode.Attributes.getNamedItem("DRAWVER").Value)
		sCatalogue = trim(ItemNode.Attributes.getNamedItem("CATALOUGE").Value)
		sVariantCode = trim(ItemNode.Attributes.getNamedItem("VARIANT").Value)
		sStUoM = trim(ItemNode.Attributes.getNamedItem("UOM").Value)
		sCategory = trim(ItemNode.Attributes.getNamedItem("CATEGORY").Value)
		sAttribute = trim(ItemNode.Attributes.getNamedItem("ATTRIBUTES").Value)
		arrLevel = split(trim(ItemNode.Attributes.getNamedItem("LEVEL").Value),",")
		arrGroup = split(trim(ItemNode.Attributes.getNamedItem("GROUP").Value),",")
		sOrgCode = trim(ItemNode.Attributes.getNamedItem("UNIT").Value)
		sOpenStockOrgCode = trim(ItemNode.Attributes.getNamedItem("OPSTOCKUNIT").Value)
	
		'Response.Write  "sAttribute = "& sAttribute &" ****"
	
		if sAttribute <> "NULL" then sAttribute = Pack(sAttribute)

		sExp ="//CONTROLS"
		Set Node = RootNode.selectSingleNode(sExp)

		sRecNum = trim(Node.Attributes.getNamedItem("RECNUM").Value)
		sRecRout = trim(Node.Attributes.getNamedItem("ROUTING").Value)
		sAccType = trim(Node.Attributes.getNamedItem("ACCOUNTING").Value)
		sModvat = trim(Node.Attributes.getNamedItem("MODVAT").Value)
		iReOrderLvl = trim(Node.Attributes.getNamedItem("REORDERLEVEL").Value)
		iReOrderQty = trim(Node.Attributes.getNamedItem("REORDERQTY").Value)
		iEcoOrderQty = trim(Node.Attributes.getNamedItem("ECOORDERQTY").Value)
		iBOMApplicable = trim(Node.Attributes.getNamedItem("BOMAPPLICABLE").Value)

		if iReOrderLvl = "" then iReOrderLvl = "0"
		if iReOrderQty = "" then iReOrderQty = "0"
		if iEcoOrderQty = "" then iEcoOrderQty = "0"

		Set Node = nothing

		sExp ="//UOMDETAILS"
		Set Node = RootNode.selectNodes(sExp)
		if Node.length > 0 then
			sPuUoM = trim(Node.Item(0).Attributes.getNamedItem("PUR").Value)
			sMaUoM = trim(Node.Item(0).Attributes.getNamedItem("MAN").Value)
			sSaUoM = trim(Node.Item(0).Attributes.getNamedItem("SAL").Value)
			iStToPur = trim(Node.Item(0).Attributes.getNamedItem("PURFAC").Value)
			iStToMan = trim(Node.Item(0).Attributes.getNamedItem("MANFAC").Value)
			iStToSal = trim(Node.Item(0).Attributes.getNamedItem("SALFAC").Value)
			sStToPurOp = trim(Node.Item(0).Attributes.getNamedItem("PUROPE").Value)
			sStToManOp = trim(Node.Item(0).Attributes.getNamedItem("MANOPE").Value)
			sStToSalOp = trim(Node.Item(0).Attributes.getNamedItem("SALOPE").Value)
		end if
	
			
		sExp ="//Output/STORAGE [@UNIT = '"&sOrgCode&"']"
					
		Set StorageNode = RootNode.Selectnodes(sExp)
		'	Response.Write "<P>"& StorageNode.length & vbcrlf
		for iCtr =  0 to StorageNode.length  -1					
			sMonYr = StorageNode.Item(iCtr).Attributes.getNamedItem("MONTHYEAR").Value
		next
	
		''Added by Ragav for Classification Code 
		sExp="//Output/CLASSIFICATION [@CATEGORY = '"&sCategory&"']"
		set ClassNode = RootNode.selectSinglenode(sExp)
		'
		'Response.Write "ClassNode.length = "& ClassNode.length & vbCrLf 
		iClass= ClassNode.Attributes.getNamedItem("CODE").value
	
	
	
	
	
		'Response.Clear
		'sMonYr = mid(FormatDate(date()),4,2)&right(FormatDate(date()),4)
		'arrFin1 = split(GetFinancialYear(sMonYr),":")
		'dFinFrom = arrFin1(0)
		'dFinTo = arrFin1(1)
	
		arrFin = split(Session("FinPeriod"),":")
		'sFinFrom = "01/04"&arrFin(0)
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)
		'sFinTo = "31/03"&arrFin(1)
		'Response.Write sFinFrom & " " & sFinTo
		'Response.End 
	
		If checkLeapYear(sFinTo) Then
			lastDateStr = "31,29,31,30,31,30,31,31,30,31,30,31"
		Else
			lastDateStr = "31,28,31,30,31,30,31,31,30,31,30,31"
		End If
		lastDateArray = split(lastDateStr,",")
	
		Set StockNode = oDOM.createElement("STOCK")
		StockNode.setAttribute "UNIT", sOpenStockOrgCode
		StockNode.setAttribute "TRANSDATE", FormatDate(date())
		'StockNode.setAttribute "FINFROMDATE", dFinFrom
		'StockNode.setAttribute "FINTODATE", dFinTo
		StockNode.setAttribute "FINFROMDATE", "01/04/"&sFinFrom
		StockNode.setAttribute "FINTODATE", "31/03/"&sFinTo
		StockNode.setAttribute "SRCTYPE", "RO"
		StockNode.setAttribute "TRANSACTIONTYPE", "RO"
		StockNode.setAttribute "REFTYPE", "0"
		StockNode.setAttribute "RECEIPTFOR", "0"
		oDOM.appendChild StockNode

		'Response.Write "A"

		'con.beginTrans

		if sitmType = "PLA" then
			sItmActive = "'Y'"
			sItmStock = "'N'"
		else
			sItmActive = "'Y'"
			sItmStock = "'S'"
		end if

		if sVariantCode = "" then
			sVariantCode = "NULL"
		else
			sVariantCode = Pack(sVariantCode)
		end if

		if sitmShDesc = "" then
			sitmShDesc = "NULL"
		else
			sitmShDesc = Pack(sitmShDesc)
		end if

		if sitmAddDesc = "" then
			sitmAddDesc = "NULL"
		else
			sitmAddDesc = Pack(sitmAddDesc)
		end if

		if sDrwVerNo = "" then
			sDrwVerNo = "NULL"
		else
			sDrwVerNo = Pack(sDrwVerNo)
		end if

		if sCatalogue = "" then
			sCatalogue = "NULL"
		else
			sCatalogue = Pack(sCatalogue)
		end if

		if iStToPur = "" then
			iStToPur = "NULL"
		elseif not iStToPur = "" and sPuUoM = "select" then
			iStToPur = "NULL"
		end if

		if iStToMan = "" then
			iStToMan = "NULL"
		elseif not iStToMan = "" and sMaUoM = "select" then
			iStToMan = "NULL"
		end if

		if iStToSal = "" then
			iStToSal = "NULL"
		elseif not iStToSal = "" and sSaUoM = "select" then
			iStToSal = "NULL"
		end if

		if sPuUoM = "" or sPuUoM = "select" then
			sPuUoM = "NULL"
			sPurEli = "'0'"
		else
			sPuUoM = Pack(sPuUoM)
			sPurEli = "'1'"
		end if

		if sMaUoM = "" or sMaUoM = "select" then
			sMaUoM = "NULL"
			sManEli = "'0'"
		else
			sMaUoM = Pack(sMaUoM)
			sManEli = "'1'"
		end if

		if sSaUoM = "" or sSaUoM = "select" then
			sSaUoM = "NULL"
			sSalEli = "'0'"
		else
			sSaUoM = Pack(sSaUoM)
			sSalEli = "'1'"
		end if

		if sStToPurOp = "" or sStToPurOp = "select" then
			sStToPurOp = "'0'"
		else
			sStToPurOp = Pack(sStToPurOp)
		end if

		if sStToManOp = "" or sStToManOp = "select" then
			sStToManOp = "'0'"
		else
			sStToManOp = Pack(sStToManOp)
		end if

		if sStToSalOp = "" or sStToSalOp = "select" then
			sStToSalOp = "'0'"
		else
			sStToSalOp = Pack(sStToSalOp)
		end if

		if instr(1,sPuUoM,sStUoM) > 0 then
			sStToPurOp = "'0'"
			iStToPur = "1"
		end if

		if instr(1,sMaUoM,sStUoM) > 0 then
			sStToManOp = "'0'"
			iStToMan = "1"
		end if

		if instr(1,sSaUoM,sStUoM) > 0 then
			sStToSalOp = "'0'"
			iStToSal = "1"
		end if

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iInvRecNo = dcrs(0)
		end if
		dcrs.Close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(ITEMCODE),0) + 1 FROM INV_M_ITEMMASTER WHERE (ITEMCODE = (SELECT ISNULL(MAX(ITEMCODE), 0) FROM INV_M_ITEMMASTER))"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set iCount = dcrs(0)

		iItmCode = iCount
		dcrs.close
	''''''''''''''''' Item definition ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		i = 0
	'	iClass = 0
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			if sOrgCode = "ALL" then
				.Source = "SELECT OUDEFINITIONID FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
			else
				.Source = "SELECT OUDEFINITIONID FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = " & Pack(sOrgCode) & " ORDER BY OUDEFINITIONID"
			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing
		 
		do while not dcrs2.EOF
			sOrgCode = trim(dcrs2(0))

			If cint(iItmCode) = 1 Then
				sSql = "INSERT INTO INV_M_ITEMMASTER (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,COMPANYITEMCODE,SHORTDESCRIPTION," &_
					"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,DRAWINGNUMBER,PURCHASEELIGIBLE,MANUFACTUREELIGIBLE,SALESELIGIBLE,STORESUOM, " &_
					"PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR,MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE, " &_
					"SALETOSTOREOPERATOR,ITEMCONTROLLER,ITEMDEFINEDON,ITEMDEFINEDBY,ITEMDELETED,ITEMSOURCE,STOCKNONSTOCK,ITEMACTIVE,CATALOGUENO, " &_
					"ITEMTYPEID,CATEGORYCODE,RECEIPTNUMBERING,RECEIPTROUTING,ACCOUNTINGTYPE,ATTRIBUTELIST,REORDERLEVEL,REORDERQTY,ECOORDERQTY,BOMAPPLICABILITY) " &_
					" VALUES (" & iItmCode & "," & iClass & "," & Pack(sOrgCode) & "," & Pack(ucase(sitmCode)) & "," & sitmShDesc & "," & Pack(sitmDesc) & "," &_
					"" & sitmAddDesc & "," & sDrwVerNo & "," & sPurEli & "," & sManEli & "," & sSalEli & "," & Pack(sStUoM) & "," & sPuUoM & "," & iStToPur & "," &_
					"" & sStToPurOp & "," & sMaUoM & "," & iStToMan & "," & sStToManOp & "," & sSaUoM & "," & iStToSal & "," & sStToSalOp & "," & iitmController & "," &_
					" CONVERT(DATETIME,GETDATE(),103)," & iitmController & ",'0','O'," & sItmStock & "," & sItmActive & "," & sCatalogue & "," & Pack(sitmType) & ", "&_
					"" & Pack(sCategory) & "," & Pack(sRecNum) & "," & Pack(sRecRout) & "," & Pack(sAccType) & "," & sAttribute & "," & iReOrderLvl & "," & iReOrderQty & "," & iEcoOrderQty & "," & iBOMApplicable & ") "
				Response.Write sSql
				con.Execute sSql

			else
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ITEMDESCRIPTION FROM INV_M_ITEMMASTER WHERE (LOWER(ITEMDESCRIPTION) = " & Pack(lcase(sitmDesc)) & " OR LOWER(COMPANYITEMCODE) = " & Pack(lcase(sitmCode)) & ") and ORGANISATIONCODE = "&Pack(sOrgCode)
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				If Not dcrs.EOF Then
					Response.Write("N")
					Response.End
			    End If
			    dcrs.Close

				sSql = "INSERT INTO INV_M_ITEMMASTER (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,COMPANYITEMCODE,SHORTDESCRIPTION," &_
					"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,DRAWINGNUMBER,PURCHASEELIGIBLE,MANUFACTUREELIGIBLE, " &_
					"SALESELIGIBLE,STORESUOM,PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR, " &_
					"MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE,SALETOSTOREOPERATOR, " &_
					"ITEMCONTROLLER,ITEMDEFINEDON,ITEMDEFINEDBY,ITEMDELETED,ITEMSOURCE,STOCKNONSTOCK,ITEMACTIVE,CATALOGUENO, " &_
					"ITEMTYPEID,CATEGORYCODE,RECEIPTNUMBERING,RECEIPTROUTING,ACCOUNTINGTYPE,ATTRIBUTELIST,REORDERLEVEL,REORDERQTY,ECOORDERQTY,BOMAPPLICABILITY) VALUES " &_
					"(" & iItmCode & "," & iClass & "," & Pack(sOrgCode) & "," & Pack(ucase(sitmCode)) & "," & sitmShDesc & "," &_
					"" & Pack(sitmDesc) & "," & sitmAddDesc & "," & sDrwVerNo & "," & sPurEli & "," & sManEli & "," &_
					"" & sSalEli & "," & Pack(sStUoM) & "," & sPuUoM & "," & iStToPur & "," & sStToPurOp & "," &_
					"" & sMaUoM & "," & iStToMan & "," & sStToManOp & "," & sSaUoM & "," &_
					"" & iStToSal & "," & sStToSalOp & "," & iitmController & ",CONVERT(DATETIME,GETDATE(),103)," &_
					"" & iitmController & ",'0','O'," & sItmStock & "," & sItmActive & "," & sCatalogue & ","&_
					"" & Pack(sitmType) & "," & Pack(sCategory) & "," & Pack(sRecNum) & "," & Pack(sRecRout) & ","&_
					"" & Pack(sAccType) & "," & sAttribute & "," & iReOrderLvl & "," & iReOrderQty & "," & iEcoOrderQty & "," & iBOMApplicable & ")"
				Response.Write sSql
				con.Execute sSql

			end if

			if sAttrFlag then
				for iCtr = 0 to UBound(sAttribute)
					sSql = "INSERT INTO INV_M_CATALOGITEMATTRIBUTES (ITEMCODE,ORGANISATIONCODE," &_
						"ATTRIBUTEID) VALUES " &_
						"(" & iItmCode & "," & Pack(sOrgCode) & "," & trim(sAttribute(iCtr)) & ")"
					'Response.Write sSql & vbCrLf & vbCrLf
					'con.Execute sSql
				next
			end if

			' Catalog Details Insert
			iCtr = 0
			sGroupPath = ""
			for iCtr =  0 to UBound(arrLevel)
				iLevel = trim(arrLevel(iCtr))
				iGroup = trim(arrGroup(iCtr))
				sGroupPath = sGroupPath & ":" & iGroup
				sSql = "INSERT INTO INV_M_CATALOGRELATION (ITEMCODE,ORGANISATIONCODE," &_
					"LEVELID,GROUPCODE) VALUES " &_
					"(" & iItmCode & "," & Pack(sOrgCode) & "," & iLevel & "," & Pack(iGroup) & ")"

				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			next
			sGroupPath = mid(sGroupPath,2)
			iCtr = 0
			for j = 0 to 0
				'iClass = "0"

				'Optional UOM Details
				iCtr = 0
				sExp = "//UOMDETAILS/OPUOMENTRY"
				Set Node = RootNode.Selectnodes(sExp)

				if Node.length > 0 then
					for iCtr =  0 to Node.length - 1
						sOpUoMCode = Node.Item(iCtr).Attributes.getNamedItem("UCODE").Value
						iOpUoMFactor = Node.Item(iCtr).Attributes.getNamedItem("BRATE").Value
						sOpUoMOperator = Node.Item(iCtr).Attributes.getNamedItem("OPERATOR").Value
						sOpUoMFor = Node.Item(iCtr).Attributes.getNamedItem("FOR").Value
						if sOpUoMFor = "Sales" then
							sOpUoMFor = "S"
						elseif sOpUoMFor = "Purchase" then
							sOpUoMFor = "P"
						end if

						with dcrs
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ITEMCODE FROM INV_M_ITEMOPTIONALUOM WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND UOMCODE = " & Pack(sOpUoMCode) & " AND OPTIONALUOMFOR = " & Pack(sOpUoMFor) & ""
							.ActiveConnection = con
							.Open
						end with

						set dcrs.ActiveConnection = nothing

						if dcrs.EOF then
							sSql = "INSERT INTO INV_M_ITEMOPTIONALUOM (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
								"OPTIONALUOMFOR,UOMCODE,OPTIONTOBASERATE,OPTIONTOBASEOPERATOR) VALUES" &_
								"(" & iItmCode & "," & iClass & "," & Pack(sOrgCode) & "," & Pack(sOpUoMFor) & "," &_
								"" & Pack(sOpUoMCode) & "," & iOpUoMFactor & "," & Pack(sOpUoMOperator) & ")"

							'Response.Write sSql & vbCrLf & vbCrLf
							con.Execute sSql
						end if
						dcrs.Close
					next
				end if

				' Storage Location Details Insert
				iCtr = 0
				iItemTotalQty = 0
				iItemTotalValue = 0
				 
			'	Response.Write "sOrgCode="&sOrgCode
				sExp ="//Output/STORAGE [@UNIT = '"&sOrgCode&"']"
				iStoEntNo = 0 
				Set StorageNode = RootNode.Selectnodes(sExp)
			'	Response.Write "<P>"& StorageNode.length & vbcrlf
				for iCtr =  0 to StorageNode.length  -1	
				iStoEntNo	 = iStoEntNo + 1
					sLoc = StorageNode.Item(iCtr).Attributes.getNamedItem("STORE").Value
					sBin = StorageNode.Item(iCtr).Attributes.getNamedItem("BIN").Value
					sMonYr = StorageNode.Item(iCtr).Attributes.getNamedItem("MONTHYEAR").Value
					iQty = StorageNode.Item(iCtr).Attributes.getNamedItem("QTY").Value
					iValue = StorageNode.Item(iCtr).Attributes.getNamedItem("STORAGEVALUE").Value
					
							
					iItemTotalQty = cdbl(iItemTotalQty) + cdbl(iQty)
					iItemTotalValue = cdbl(iItemTotalValue) + cdbl(iValue)
					'sOrgCode = StorageNode.Item(iCtr).Attributes.getNamedItem("UNIT").Value
					
					sDateOfRec= LastDayOfMonth(Left(sMonYr,2))&"/"&Left(sMonYr,2)&"/"&Right(sMonYr,4)
					'Response.Write "GetFinancialYear(sMonYr)="&sDateOfRec
			'		Response.Write "sBin="& sBin
					'if sBin = "NULL" then sBin = cint(0)

					Set newElem1 = oDOM.createElement("STORAGE")
					newElem1.setAttribute "STOENTRYNO",iStoEntNo
					newElem1.setAttribute "ITEM",iItmCode
					newElem1.setAttribute "CLASS",iClass
					newElem1.setAttribute "STORE",sLoc
					newElem1.setAttribute "BIN",sBin
					newElem1.setAttribute "STOREQTY",iQty
					newElem1.setAttribute "STOREVALUE",iValue
					newElem1.setAttribute "DATERECEIVED",lastDateArray(Cint(mid(sMonYr,1,2))-1)&"/"&mid(sMonYr,1,2)&"/"&mid(sMonYr,3,4)
					StockNode.appendChild newElem1
					'Response.Write "AAA"	
					
					if cint(sBin) = 0 then sBin = "NULL"
					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT APPLICABLEFOR FROM INV_M_STORAGE WHERE OUDEFINITIONID = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing
					if not dcrs.EOF then
						sAppli = trim(dcrs(0))
					end if
					dcrs.Close

					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ITEMCODE FROM INV_M_ITEMSTORAGE WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND APPLICABLEFOR = " & Pack(sAppli) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing
					if dcrs.EOF then
						sSql = "INSERT INTO INV_M_ITEMSTORAGE (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
							"APPLICABLEFOR,LOCATIONNUMBER,BINNUMBER,ALLOWTRANSFERS) VALUES " &_
							"(" & iItmCode & "," & iClass & "," & Pack(sOrgCode) & "," & Pack(sAppli) & "," &_
							"" & sLoc & "," & sBin & ",'1')"
						'objTxt.write sSql & vbCrLf & vbCrLf
						'Response.Write sSql & vbCrLf
						con.Execute sSql
					end if
					dcrs.Close
					
					sExp1 = "//Output/BOM/Item"
					set ItmNode = RootNode.Selectnodes(sExp1)
					if ItmNode.Length > 0 then
					    For iItmCtr = 0 to ItmNode.Length - 1
					        iBoMItemCode = ItmNode.Item(iItmCtr).Attributes.getNamedItem("ItemCode").Value
							iBoMClassCode = ItmNode.Item(iItmCtr).Attributes.getNamedItem("ClassCode").Value
							iBoMQty = ItmNode.Item(iItmCtr).Attributes.getNamedItem("Qty").Value
							iBOMUOM = ItmNode.Item(iItmCtr).Attributes.getNamedItem("UoM").Value
							iBomType = ItmNode.Item(iItmCtr).Attributes.getNamedItem("Type").Value
							iBomConsumable = trim(ItmNode.Item(iItmCtr).Attributes.getNamedItem("Consumable").Value)
							
							sSql = "Insert into INV_M_ITEMMASTERBOM (ItemCode,ClassificationCode,OrganisationCode,"&_
			                       " BOMClassificationCode,BOMItemCode,Quantity,UOM,Type,Consumable) values "&_
			                       "("& iItmCode &","& iClass &",'"&sOrgCode&"',"& iBoMClassCode &","& iBoMItemCode &","& iBoMQty &",'"& iBOMUOM &"','"& iBomType &"','"& iBomConsumable &"')"
			                       
			                Response.Write sSql
			                 
					        con.execute sSql
					    Next
					end if
					

					sReceivedOn = "01/"&left(sMonYr,2)&"/"&right(sMonYr,4)
	'Response.Write vbCrLf 
	'Response.Write vbCrLf & "iClass="& iClass
					if sBin = "NULL" then sBin = cint(0)

					sExp1 ="//Output/STORAGE [ @STORE = "&sLoc&" and @BIN = '"&sBin&"' and @CLASSIFICATION = "&iClass&" and @UNIT = '"&sOrgCode&"']/LotSerial"

					if cint(sBin) = 0 then sBin = "NULL"

					Set LotNode = RootNode.Selectnodes(sExp1)
					if LotNode.Length > 0 then
						For iLotCtr = 0 to LotNode.Length - 1

							iCounter = 0
							'Response.Write "Lot="&LotNode.Item(iLotCtr).Attributes.getNamedItem("LOT").Value & vbCrLf
							iLot = LotNode.Item(iLotCtr).Attributes.getNamedItem("LOT").Value
							iSerialFrom = LotNode.Item(iLotCtr).Attributes.getNamedItem("SERIALFROM").Value
							iSerialTo = LotNode.Item(iLotCtr).Attributes.getNamedItem("SERIALTO").Value
							iLotQty = LotNode.Item(iLotCtr).Attributes.getNamedItem("QTY").Value
							iEntryCtr = LotNode.Item(iLotCtr).Attributes.getNamedItem("COUNTER").Value
							iLotValue = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("IVALUE").Value)

							if sitmType <> "GAR" and sRecNum <> "N" then
								iAltGross = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTGROSS").Value)
								iAltNett = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTNETT").Value)
								sAltUoM = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTUOM").Value)
								sQtyIn = LotNode.Item(iLotCtr).Attributes.getNamedItem("QTYIN").Value
								iTareQty = LotNode.Item(iLotCtr).Attributes.getNamedItem("TARE").Value
								sTareIn = LotNode.Item(iLotCtr).Attributes.getNamedItem("TAREWEIGHT").Value
								sStage = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("STAGE").Value)
							elseif sitmType = "GAR" or sRecNum = "N" then
								sAttr = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ATTRIBUTE").Value)
							end if

							if sAttr = "" or sAttr = "NULL" then
								sAttr = "0"
							else
								sAttr = sAttr
							end if

							if sQtyIn = "" or sQtyIn = "NULL" then
								sQtyIn = "NULL"
							else
								sQtyIn = Pack(sQtyIn)
							end if

							if sStage = "select" or sStage = "" then
								sStage = "0"
							else
								sStage = sStage
							end if

							if iTareQty = "" then iTareQty = "0"
							if iAltGross = "" then iAltGross = "0"
							if iAltNett = "" then iAltNett = "0"

							if sAltUoM = "" or lcase(sAltUoM) = "select" or sAltUoM = "NULL" then
								sAltUoM = "NULL"
							else
								sAltUoM = Pack(sAltUoM)
							end if

							iItmRate = cdbl(iValue) / cdbl(iQty)

							sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
								"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
								"(" & Pack(sOrgCode) & "," & iItmCode & "," & iClass & "," &_
								"'RO'," & iInvRecNo & ",CONVERT(DATETIME," & Pack(sReceivedOn) & ",103)," & iLotQty & "," & iLotValue & ")"
							'objTxt.write sSql & vbCrLf & vbCrLf
							'Response.Write sSql & vbCrLf & vbCrLf
							'Response.Write "iLot="& iLot 
							'con.Execute sSql

							if sBin = "NULL" then sBin = cint(0)
							sExp1 ="//Output/STORAGE [ @STORE = "&sLoc&" and @BIN = '"&sBin&"']/LotSerial [ @LOT = '"&iLot&"' and @COUNTER = "&iEntryCtr&"]/LotSerialDetails"
							if sBin = 0 then sBin = "NULL"
							Set LotSerialNode = RootNode.Selectnodes(sExp1)

							For iCounter = 0 to LotSerialNode.Length - 1
								iLotEntry = iLot
								iSerialEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value)
								if sitmType = "GAR" or sRecNum = "N" then
									iQtyRecEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SERIALQTY").Value)
									sBarCode = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("BARCODE").Value)

									if sBarCode = "" then
										sBarCode = "NULL"
									else
										sBarCode = Pack(sBarCode)
									end if
								else
									if sRecNum <> "N" then
										iQtyRecEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("QTYREC").Value)
										iTareQtyEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("TAREREC").Value)
										sSellingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGTYPE").Value)
										iWeight = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("WEIGHTSTYPE").Value)
										sPackingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKINGTYPE").Value)
										sSellingFormType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGFORM").Value)
										sAttr = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").Value)
									end if
								end if
								if trim(sAttr)="" then sAttr = "NULL"

								if sRecNum <> "N" then
									sPackingNumber = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKNUMBER").Value)
								else
									sPackingNumber = "NULL"
								end if
								iLotValue = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value)

								if sSellingType = "select" or sSellingType = "" then
									sSellingType = "0"
									iWeight = "0"
								end if

								if sPackingType = "select" or sPackingType = "" then
									sPackingType = "0"
								end if

								if sSellingFormType = "select" or sSellingFormType = "" then
									sSellingFormType = "0"
								end if

								if iLotEntry <> "N/A" and iSerialEntry = "0" then iSerialEntry = "NULL"
								
								if iLotEntry = "N/A" or iLotEntry = "NULL" or trim(iLotEntry)="" or IsNull(iLotEntry) or  trim(iLotEntry)="0"  then
									iLotEntry = "NULL"
								else
									iLotEntry = iLotEntry
								end if

								if sQtyIn = "N" then
									iQtyGross = cdbl(iQtyRecEntry) + cdbl(iTareQtyEntry)
									iQtyNett = iQtyRecEntry
								else
									iQtyGross = iQtyRecEntry
									iQtyNett = cdbl(iQtyRecEntry) - cdbl(iTareQtyEntry)
								end if

								if sItmType = "FAB" then iQtyNett = iQtyRecEntry
								if iTareQtyEntry = "" then iTareQtyEntry = "0"
								with dcrs
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LOCATIONLOT"
									.ActiveConnection = con
									.Open
								end with
								set dcrs.ActiveConnection = nothing

								if not dcrs.EOF then
									iSerialEntry = dcrs(0)
								end if
								dcrs.Close
					
								iRate = FormatNumber(cdbl(iLotValue) / cdbl(iQtyNett),4,,,0)
								if sBarCode = "" then sBarCode = "NULL"

								if sBin = "NULL" then sBin = cint(0)
								'Response.Write "LOT"
								Set newElem2 = oDOM.createElement("LOT")
								newElem2.setAttribute "ITEM",iItmCode
								newElem2.setAttribute "CLASS",iClass
								newElem2.setAttribute "STORE",sLoc
								newElem2.setAttribute "BIN",sBin
								newElem2.setAttribute "LOT",iLot
								newElem2.setAttribute "QTY",iQtyNett
								newElem2.setAttribute "RATE",iRate
								newElem2.setAttribute "GROSSQTY",iQtyGross
								newElem2.setAttribute "PACKINGNUMBER",sPackingNumber
								newElem2.setAttribute "PACKINGCODE",sPackingType
								newElem2.setAttribute "SELLINGNUMBER",sSellingType
								newElem2.setAttribute "WEIGHTPERSELLINGFORM",iWeight
								newElem2.setAttribute "SELLINGFORM",sSellingFormType
								newElem2.setAttribute "STAGE",sStage
								newElem2.setAttribute "ATTRIBUTE",sAttr

								newElem1.appendChild newElem2

								sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
									"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
									"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,"&_
									"SELLINGFORM,STAGE,RATE,BARCODETAG,ATTRIBUTELIST,SRCTYPE,DATEOFRECEIPT) VALUES " &_
									"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItmCode & "," & iClass& "," &_
									"" & sLoc & "," & sBin & "," & iLotEntry & "," & iSerialEntry & "," &_
									"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & ",'" & sPackingNumber & "'," &_
									"" & sPackingType & "," & sSellingType & "," & iWeight & "," & sSellingFormType & ", "&_
									"" & sStage & "," & iRate & "," & sBarCode & ",'" & sAttr & "','RO',CONVERT(DATETIME,"&sDateOfRec&",103))"
								'objTxt.write sSql & vbCrLf & vbCrLf
							'	Response.Write sSql & vbCrLf & vbCrLf
								 'con.Execute sSql
							next
						next
						
						
					else
						if iQty <> 0 then
							iItmRate = FormatNumber(cdbl(iValue) / cdbl(iQty),4,,,0)
						else
							iItmRate = 0
						end if
				
						sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
							"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
							"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,"&_
							"SELLINGFORM,STAGE,RATE,BARCODETAG,ATTRIBUTELIST,SRCTYPE,DATEOFRECEIPT) VALUES " &_
							"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItmCode & "," & iClass& "," &_
							"" & sLoc & "," & sBin & ",NULL,NULL," & iQty & "," & iQty & ",0,NULL," &_
							"NULL,NULL,NULL,NULL,NULL," & iItmRate & ",NULL,NULL,'RO',CONVERT(DATETIME,"&sDateOfRec&",103))"
						'objTxt.write sSql & vbCrLf & vbCrLf
							'Response.Write "1="&sSql & vbCrLf & vbCrLf
						 'con.Execute sSql

						sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
							"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
							"(" & Pack(sOrgCode) & "," & iItmCode & "," & iClass & "," &_
							"'RO',NULL,CONVERT(DATETIME," & Pack(sReceivedOn) & ",103)," & iQty & "," & iValue & ")"
						'objTxt.write sSql & vbCrLf & vbCrLf
						'Response.Write sSql & vbCrLf & vbCrLf
						' con.Execute sSql

					end if

				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					'arrFin = split(GetFinancialYear(sMonYr),":")
					'sFinFrom = arrFin(0)
					'sFinTo = arrFin(1)
					arrFin = split(Session("FinPeriod"),":")
					'sFinFrom = "01/04"&arrFin(0)
					'sFinTo = "31/03"&arrFin(1)
					sFinFrom = arrFin(0)
					sFinTo = arrFin(1)


					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing
					if dcrs.EOF then
						sSql = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
							"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
							"(" & Pack(sOrgCode) & "," & iClass & "," & iItmCode & "," &_
							"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
						'Response.Write sSql & vbCrLf & vbCrLf
						'con.Execute sSql
					else
						sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEAROPENINGSTOCK = (YEAROPENINGSTOCK + " & iQty & ")," &_
							"YEAROPENINGVALUE = (YEAROPENINGVALUE + " & iValue & ")," &_
							"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQty & "), " &_
							"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
							"ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
							"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
							"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
							"CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
						objTxt.write sSql & vbCrLf & vbCrLf
						'Response.Write sSql & vbCrLf & vbCrLf
						'con.Execute sSql
					end if
					dcrs.Close

					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
						.ActiveConnection = con
						.Open
					end with

					set dcrs.ActiveConnection = nothing

					if dcrs.EOF then
						sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
							"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
							"YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
							"(" & Pack(sOrgCode) & "," & iClass & "," & iItmCode & "," &_
							"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
							"" & sLoc & "," & sBin & "," &_
							"" & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
						'objTxt.write sSql & vbCrLf & vbCrLf
						'Response.Write sSql & vbCrLf & vbCrLf
						'con.Execute sSql
					end if
					dcrs.Close

				next
				
			next
			
		dcrs2.MoveNext
		
		loop
		dcrs2.Close
		IF iItemTotalQty = "" then iItemTotalQty = 0
		IF iItemTotalValue = "" then iItemTotalValue  = 0
		IF sAttribute = "NULL" then sAttribute = ""
		Set newElem = oDOM.createElement("ITEM")
		newElem.setAttribute "ITEMENTRYNO",iItemEntNo
		newElem.setAttribute "ITEM",iItmCode
		newElem.setAttribute "CLASS",iClass
		newElem.setAttribute "ITEMQTY",iItemTotalQty
		newElem.setAttribute "ITEMVALUE",iItemTotalValue
		newElem.setAttribute "ATTRIBUTE",sAttribute
		newElem.setAttribute "SUMQTY",iItemTotalQty
		StockNode.appendChild newElem
		
		'Response.Clear 
		'Response.Write oDOM.xml
		'Response.End

		dim adoCmd
		'Declaration of Objects
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection = con
		Response.Write "XML="&oDOM.xml
		oDOM.save server.MapPath("../Temp/Master/ItemStock.xml")
		adoCmd.CommandText = "StockUpdationItemCreate"
		adoCmd.CommandType = 4
		adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(oDOM.xml),oDOM.xml)
		adoCmd.Execute()

		set adoCmd = nothing
		'con.RollbackTrans
		
		'Response.Clear
		'Response.End
	End Function

%>
