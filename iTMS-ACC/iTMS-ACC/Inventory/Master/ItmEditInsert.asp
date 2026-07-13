<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmEditInsert.asp
	'Module Name				:	Inventory (Item Creation and Definition)
	'Author Name				:	S.MAHESHWARI
	'Created On					:	September 07, 2007
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	ItmEditEntry.asp
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
<!-- #include File="../../include/NoSeries.asp" -->
<!-- #include File="../../include/NoSeriesCommonFunctions.asp"-->
<%
	dim dcrs,dcrs1,dcrs2,sSql,oDOM,objfs,RootNode,ItemNode,Node,UnitNode,StorageNode,LotNode
	dim newxml,LotSerialNode,AssetNode,BoMNode,FabNode,FabDetNode
	dim sitmCode,sitmShDesc,sitmDesc,sitmAddDesc,sitmUsage,iitmController,sStUoM,sPuUoM
	dim sMaUoM,sSaUoM,iStToPur,iStToMan,iStToSal,sStToPurOp,sStToManOp,sStToSalOp,sitmType
	dim sFormCode,sItmWho,sItmActive,sItmStock,sSalEli,sManEli,sPurEli
	dim iCodeLen,iCodeSize,sDrwVerNo,sCatalogue,sExp,sExp1,iClass,sParentCode,iitmCode
	dim iBoMClassCode,iCtrBoM,iBoMItemCode,iBoMQty,sBoMUoM,sBoMType,sBoMUnit
	dim sWeave,iWidth,iReedCount,sReedSpace,sVariety,iNODent,iNOTotal,iNOEnds
	dim iNOPicks,sAverageWC,iTapeLength,sWarpYarn,sWeftYarn,iWarpCount,iWeftCount,iCtr
	dim arrWarpYarn,arrWeftYarn,i,j,sGroupPath,sDateOfRec,iStoEntNo,iItemEntNo,ItmNode
	Dim iItmCtr,iBomUom,iBomType,IBomConsumable,ndAttNode,sAutoGenDet

	dim iVar,jVar,isize,iCount,sTempitmCode,iAttributeList
	Dim Storedetail,sstorageNode,sLotNode,sAutoGen,iSNo,iSCode,dDate2

	dim sVariantCode,sOrgCode,sCategory
	dim sRecNum,sRecRout,sAccType,sABC,sModvat,sAPur,sASubCon,sASales
	dim sTemp,sLoc,sBin,sMonYr,iQty,iValue,sAppli
	dim arrFin,sFinFrom,sFinTo,arrTemp,iInvRecNo,sTempMonYr
	dim sPurUoM,sManUoM,sSalUoM,iPurRate,iManRate,iSalRate,sUoMDesc,sOpUoMFor
	dim sQtyIn,iTareQty,iLot,iSerialFrom,iSerialTo,sSellingType,iWeight,sPackingType,sPackingNumber
	dim iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sTareIn,iQtyGross,iQtyNett
	dim iSeriesNo,iSeriesCode,sProductCode,iItmRate,iCounter
	dim iLotCtr,sSellingFormType,sStage,sForm,arrAttr,sAttr,sBarCode
	dim iLotQty,iLotValue,iEntryCtr,sReceivedOn
	dim iReOrderLvl,iReOrderQty,iEcoOrderQty,iLedEntNo,iSrcTypeCode,sTransType
	dim iFabCtr,iPieceNo,iPieceQty,arrFab,iAltGross,iAltNett,sAltUoM,sCheck,iRate
	dim sOpUoMCode,iOpUoMFactor,sOpUoMOperator,sAttrFlag,sCmpItmCode,iBOMApplicable
	dim sTempType,sTempCat,iParentCode,arrParent,arrGroup,sAttribute,arrLevel,iLevel,iGroup,sStatus
	dim objFSO,objTxt,sModVatEligibility,sPurTaxType,sSalTaxType

	Dim HeaderNode,newElem,newElem1,newElem2,StockNode,iItemTotalQty,iItemTotalValue
	Dim LotNode1,iLSerQty,iLPackNo,iLVal,iLAttbID,iLAttbList,iLBarCode,iLCtr,iLLotNo
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/master/Lot.txt"))

	iCodeLen = 0
	iCodeSize = 0
	iItemEntNo = 1
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

	set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	set objfs = Server.CreateObject("Scripting.FileSystemObject")

	'Create our DOM Document Objects
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")

	iitmController = getUserid

	newxml.async = false
	newxml.load(Request)
	newxml.save server.mappath("../temp/master/itemstockedit.xml")
	

	Set RootNode = newxml.documentElement
'	Response.Write "newxml="&newxml.xml
sAutoGen=""
	sExp ="//DETAILS"
	Set ItemNode = RootNode.selectSingleNode(sExp)
	sitmType = trim(ItemNode.Attributes.getNamedItem("ITYPE").Value)
	sitmCode = trim(ItemNode.Attributes.getNamedItem("ICODE").Value)
	sCmpItmCode = trim(ItemNode.Attributes.getNamedItem("COMPITEMCODE").Value)
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
	sStatus  = trim(ItemNode.Attributes.getNamedItem("ITEMSTATUS").value)
	sModVatEligibility = trim(ItemNode.Attributes.getNamedItem("MODVAT").value)
	sPurTaxType = trim(ItemNode.Attributes.getNamedItem("PURTAX").value)
	sSalTaxType = trim(ItemNode.Attributes.getNamedItem("SALTAX").value)

	sitmDesc = Replace(sitmDesc,"~~",chr(39))
	sitmDesc = Replace(sitmDesc,"``",chr(34))
	
	if trim(sPurTaxType)="0" or trim(sPurTaxtype)="" then sPurTaxType = "NULL"
	if trim(sSalTaxType)="0" or trim(sSalTaxType)="" then sSalTaxType = "NULL"

	'Response.Write "sitmCode="&sitmCode
	if sAttribute <> "NULL" then sAttribute = Pack(sAttribute)
	if trim(sModVatEligibility)="" or isNull(sModVatEligibility) then sModVatEligibility="NULL"
	if trim(sModVatEligibility)<>"NULL" then sModVatEligibility=pack(sModVatEligibility)

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

	sExp = "//CLASSIFICATION"
	Set Node = RootNode.selectSingleNode(sExp)
	iClass = Node.getAttribute("CODE")

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

	sMonYr = mid(FormatDate(date()),4,2)&right(FormatDate(date()),4)
	'Response.Write "sMonYr=" & sMonYr
'	arrFin = split(GetFinancialYear(sMonYr),":")
	arrFin = split(session("FinPeriod"),":")
	sFinFrom = "01/04/"&arrFin(0)
	sFinTo = "31/03/"&arrFin(1)
	sDateOfRec= LastDayOfMonth(Left(sMonYr,2))&"/"&Left(sMonYr,2)&"/"&Right(sMonYr,4)
	'Response.Write sFinFrom & " " & sFinTo
'	Response.End
	Set StockNode = oDOM.createElement("STOCK")
	StockNode.setAttribute "UNIT", sOrgCode
	StockNode.setAttribute "TRANSDATE", FormatDate(date())
	StockNode.setAttribute "FINFROMDATE", sFinFrom
	StockNode.setAttribute "FINTODATE", sFinTo
	StockNode.setAttribute "SRCTYPE", "RO"
	StockNode.setAttribute "TRANSACTIONTYPE", "RO"
	StockNode.setAttribute "REFTYPE", "0"
	StockNode.setAttribute "RECEIPTFOR", "0"
	oDOM.appendChild StockNode

	'Response.Write "Testing............"

	con.beginTrans

	if sitmType = "6" then
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
	sTransType  = "RO"
	IF sTransType = "RO" then  iSrcTypeCode = 22
	IF sTransType = "RR" then  iSrcTypeCode = 25
	IF sTransType = "R"  then  iSrcTypeCode = 7
	IF sTransType = "I"  then  iSrcTypeCode = 1


''''''''''''''''' Item definition - UPDATION''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	i = 0
	'iClass = 0


		if sAttrFlag then
			for iCtr = 0 to UBound(sAttribute)
				sSql = "UPDATE INV_M_CATALOGITEMATTRIBUTES SET ORGANISATIONCODE=" & Pack(sOrgCode) & "," &_
					"ATTRIBUTEID=" & trim(sAttribute(iCtr)) & " WHERE ITEMCODE=" & sitmCode & " "
				'Response.Write sSql &"<br><br>" & vbCrLf & vbCrLf
				con.Execute sSql
			next
		end if

		' Catalog Details Insert
		iCtr = 0
		sGroupPath = ""
		for iCtr =  0 to UBound(arrLevel)
			iLevel = trim(arrLevel(iCtr))
			iGroup = trim(arrGroup(iCtr))
			sGroupPath = sGroupPath & ":" & iGroup
			sSql = "UPDATE INV_M_CATALOGRELATION ORGANISATIONCODE=" & Pack(sOrgCode) & ",LEVELID=" & iLevel & "," &_
				   "GROUPCODE=" & Pack(iGroup) & " WHERE ITEMCODE = " & sitmCode & " "

			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		next
		sGroupPath = mid(sGroupPath,2)
		iCtr = 0

		sSql = "DELETE FROM INV_T_LOCATIONLOT WHERE	ITEMCODE="& sitmCode &" and ClassificationCode = "& iClass &" and SrcType='RO'"
	    'Response.write sSql & vbCrLf & vbCrLf
	    con.Execute sSql
	  
		sSql = "Delete from INV_M_ITEMSTORAGE where ItemCode = "& sitmCode
			    'Response.Write sSql & vbCrLf
			    con.Execute sSql

'		sSql = "UPDATE INV_M_ITEMMASTER  SET COMPANYITEMCODE=" & Pack(ucase(sCmpItmCode)) & ",SHORTDESCRIPTION=" & sitmShDesc & ",ITEMDESCRIPTION=" & Pack(sitmDesc) & "," &_
'				"ADDITIONALDESCRIPTION=" & sitmAddDesc & ",DRAWINGNUMBER=" & sDrwVerNo & ",PURCHASEELIGIBLE=" & sPurEli & ",MANUFACTUREELIGIBLE=" & sManEli & " ," &_
'				"SALESELIGIBLE=" & sSalEli & ",STORESUOM=" & Pack(sStUoM) & ",PURCHASEUOM=" & sPuUoM & ",PURTOSTORERATE=" & iStToPur & ",PURTOSTOREOPERATOR=" & sStToPurOp & ", " &_
'				"MANUFACTURINGUOM=" & sMaUoM & ",MANTOSTORERATE=" & iStToMan & ",MANTOSTOREOPERATOR=" & sStToManOp & ",SALESUOM=" & sSaUoM & ",SALETOSTORERATE=" & iStToSal & ", "&_
'				"SALETOSTOREOPERATOR=" & sStToSalOp & ",ITEMCONTROLLER=" & iitmController & ",ITEMDEFINEDON=CONVERT(DATETIME,GETDATE(),103),ITEMDEFINEDBY=" & iitmController & ","&_
'				"ITEMDELETED='0',ITEMSOURCE='0',STOCKNONSTOCK=" & sItmStock & ",ITEMACTIVE=" & sItmActive & ",CATALOGUENO=" & sCatalogue & ",ITEMTYPEID=" & Pack(sitmType) & ", " &_
'				"CATEGORYCODE=" & Pack(sCategory) & ",RECEIPTNUMBERING=" & Pack(sRecNum) & ",RECEIPTROUTING=" & Pack(sRecRout) & ",ACCOUNTINGTYPE=" & Pack(sAccType) & ", "&_
'				"ATTRIBUTELIST=" & sAttribute & ",REORDERLEVEL=" & iReOrderLvl & ",REORDERQTY=" & iReOrderQty & ",ECOORDERQTY=" & iEcoOrderQty & ",BOMAPPLICABILITY=" & iBOMApplicable & ", ClassificationCode = "& iClass &" WHERE ITEMCODE = " & sitmCode & " "
	    
	    sSql = "UPDATE INV_M_ITEMMASTER  SET COMPANYITEMCODE=" & Pack(ucase(sCmpItmCode)) & ",SHORTDESCRIPTION=" & sitmShDesc & ",ITEMDESCRIPTION=" & Pack(sitmDesc) & "," &_
				"ADDITIONALDESCRIPTION=" & sitmAddDesc & ",DRAWINGNUMBER=" & sDrwVerNo & ",PURCHASEELIGIBLE=" & sPurEli & ",MANUFACTUREELIGIBLE=" & sManEli & " ," &_
				"SALESELIGIBLE=" & sSalEli & ",STORESUOM=" & Pack(sStUoM) & ",PURCHASEUOM=" & sPuUoM & ",PURTOSTORERATE=" & iStToPur & ",PURTOSTOREOPERATOR=" & sStToPurOp & ", " &_
				"MANUFACTURINGUOM=" & sMaUoM & ",MANTOSTORERATE=" & iStToMan & ",MANTOSTOREOPERATOR=" & sStToManOp & ",SALESUOM=" & sSaUoM & ",SALETOSTORERATE=" & iStToSal & ", "&_
				"SALETOSTOREOPERATOR=" & sStToSalOp & ",ITEMCONTROLLER=" & iitmController & ",ITEMDEFINEDON=CONVERT(DATETIME,GETDATE(),103),ITEMDEFINEDBY=" & iitmController & ","&_
				"ITEMDELETED='0',ITEMSOURCE='0',STOCKNONSTOCK=" & sItmStock & ",ITEMACTIVE=" & sItmActive & ",CATALOGUENO=" & sCatalogue & ",CATEGORYCODE=" & sCategory & ",RECEIPTNUMBERING=" & Pack(sRecNum) & ",RECEIPTROUTING=" & Pack(sRecRout) & ",ACCOUNTINGTYPE=" & Pack(sAccType) & ", "&_
				"ATTRIBUTELIST=" & sAttribute & ",BOMAPPLICABILITY=" & iBOMApplicable & ", ClassificationCode = "& iClass &", AllowModvatCredit ="& sModVatEligibility &",ItemType="& sitmType &" WHERE ITEMCODE = " & sitmCode & " "

		'Response.Write sSql &"<br><br>"
		con.Execute sSql





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
						.Source = "SELECT ITEMCODE FROM INV_M_ITEMOPTIONALUOM WHERE ITEMCODE = " & sitmCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND UOMCODE = " & Pack(sOpUoMCode) & " AND OPTIONALUOMFOR = " & Pack(sOpUoMFor) & ""
						.ActiveConnection = con
						.Open
					end with

					set dcrs.ActiveConnection = nothing

					if dcrs.EOF then
						sSql = "INSERT INTO INV_M_ITEMOPTIONALUOM (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
							"OPTIONALUOMFOR,UOMCODE,OPTIONTOBASERATE,OPTIONTOBASEOPERATOR) VALUES" &_
							"(" & sitmCode & "," & iClass & "," & Pack(sOrgCode) & "," & Pack(sOpUoMFor) & "," &_
							"" & Pack(sOpUoMCode) & "," & iOpUoMFactor & "," & Pack(sOpUoMOperator) & ")"

					'	Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
					Else
						sSql = "UPDATE INV_M_ITEMOPTIONALUOM SET CLASSIFICATIONCODE=" & iClass & ",ORGANISATIONCODE=" & Pack(sOrgCode) & "," &_
							   "OPTIONALUOMFOR=" & Pack(sOpUoMFor) & ",UOMCODE=" & Pack(sOpUoMCode) & ",OPTIONTOBASERATE=" & iOpUoMFactor & "," &_
							   "OPTIONTOBASEOPERATOR=" & Pack(sOpUoMOperator) & " WHERE ITEMCODE = " & sitmCode & " "

					'	Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
					end if
					dcrs.Close
				next


			end if

	Dim StoDetNode,StoNode
			' Storage Location Details Insert
			iCtr = 0
			iStoEntNo = 0

			sSql = "Delete from INV_M_ITEMMASTERBOM where ItemCode = "& sitmCode
			'Response.Write sSql
			con.execute sSql


			sExp1 = "//BOM/Item"
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
                               "("& sitmCode &","& iClass &",'"&sOrgCode&"',"& iBoMClassCode &","& iBoMItemCode &","& iBoMQty &",'"& iBOMUOM &"','"& iBomType &"','"& iBomConsumable &"')"

                        'Response.Write sSql

				        con.execute sSql
				    Next
				end if



        'Response.Write RootNode.xml
		for each StoDetNode in RootNode.childNodes
		    if StoDetNode.nodeName="STOREDET" then
			    for each StoNode in StoDetNode.childNodes
				    iStoEntNo = iStoEntNo + 1
				    'Response.Write StoNode.xml
				    'Response.Write "sLocation = "& StoNode.getAttribute("STORE")
				    sLoc = StoNode.getAttribute("STORE")
				    sBin = StoNode.getAttribute("BIN")
				    sMonYr = StoNode.getAttribute("MONTHYEAR")
				    iQty = StoNode.getAttribute("QTY")
				    iValue = StoNode.getAttribute("STORAGEVALUE")
				    'Response.Write "sLoc = "& sLoc
				    if trim(iItemTotalQty)="" or IsNull(iItemTotalQty) then iItemTotalQty = 0
				    if trim(iItemTotalValue)="" or IsNull(iItemTotalValue) then iItemTotalValue = 0
				    if trim(iQty)="" or IsNull(iQty) then iQty = 0
				    if Trim(iValue)="" or IsNull(iValue) then iValue = 0
				    'Response.Write "iQty = "& iQty
				    iItemTotalQty = cdbl(iItemTotalQty) + cdbl(iQty)
				    iItemTotalValue = cdbl(iItemTotalValue) + cdbl(iValue)
				    'sOrgCode = StorageNode.Item(iCtr).Attributes.getNamedItem("UNIT").Value
				    sDateOfRec = LastDayOfMonth(Left(sMonYr,2))&"/"&Left(sMonYr,2)&"/"&Right(sMonYr,4)
				    'Response.Write "DATEOFRECEIPT="&sMonYr
				    'Response.Write "Left(sMonYr,2)="&sDateOfRec
				    Set newElem1 = oDOM.createElement("STORAGE")
				    newElem1.setAttribute "STOENTRYNO",iStoEntNo
				    newElem1.setAttribute "ITEM",sitmCode
				    newElem1.setAttribute "CLASS",iClass
				    newElem1.setAttribute "STORE",sLoc
				    newElem1.setAttribute "BIN",sBin
				    newElem1.setAttribute "STOREQTY",iQty
				    newElem1.setAttribute "STOREVALUE",iValue

				    if sBin = "0" then sBin = "NULL"

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
					    .Source = "SELECT ITEMCODE FROM INV_M_ITEMSTORAGE WHERE ITEMCODE = " & sitmCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND APPLICABLEFOR = " & Pack(sAppli) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
					    'Response.Write "<p>"&dcrs.Source &"--"
					    .ActiveConnection = con
					    .Open
				    end with
				    set dcrs.ActiveConnection = nothing
				    if dcrs.EOF then
					    sSql = "INSERT INTO INV_M_ITEMSTORAGE (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
						    "APPLICABLEFOR,LOCATIONNUMBER,BINNUMBER,ALLOWTRANSFERS) VALUES " &_
						    "(" & sitmCode & "," & iClass & "," & Pack(sOrgCode) & "," & Pack(sAppli) & "," &_
						    "" & sLoc & "," & sBin & ",'1')"
					    'Response.Write sSql & vbCrLf
					    con.Execute sSql
					else
					    sSql = "INSERT INTO INV_M_ITEMSTORAGE (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
						    "APPLICABLEFOR,LOCATIONNUMBER,BINNUMBER,ALLOWTRANSFERS) VALUES " &_
						    "(" & sitmCode & "," & iClass & "," & Pack(sOrgCode) & "," & Pack(sAppli) & "," &_
						    "" & sLoc & "," & sBin & ",'1')"
					    'Response.Write sSql & vbCrLf
					    con.Execute sSql

				    end if

				    dcrs.Close

				    sReceivedOn = "01/"&left(sMonYr,2)&"/"&right(sMonYr,4)

				    if sBin = "NULL" then sBin = "0"

			    '	sExp1 ="//STOREDET/STORAGE [ @STORE = "&sLoc&" and @BIN = '"&sBin&"' and @CLASSIFICATION = "&iClass&" and @UNIT = '"&sOrgCode&"']/LotSerial"

			    'Response.Write RootNode.xml

							    for each Storedetail in RootNode.childNodes
								    for each StorageNode in Storedetail.childNodes
									   ' Response.Write "sLoc = "& sLoc & " StorageNode.getAttribute = "& StorageNode.getAttribute("STORE")
									   ' Response.Write "<br>sBIn = "& sBin & " StorageNode.getAttribute = "& StorageNode.getAttribute("BIN")
									   ' Response.Write "<br>sClass = "& iClass & " StorageNode.getAttribute = "& StorageNode.getAttribute("CLASSIFICATION")
									    if strcomp(sLoc,StorageNode.getAttribute("STORE"))=0 and strcomp(sBin,StorageNode.getAttribute("BIN"))=0 and strcomp(sOrgCode,StorageNode.getAttribute("UNIT"))=0 then
									    'Response.Write "****Hello Success*****"
										    for each LotNode in StorageNode.childNodes
										       ' Response.Write ">>>>>>>>>LotNode"

											    iCounter = 0

											    iLot = LotNode.getAttribute("LOT")
											    iSerialFrom = LotNode.getAttribute("SERIALFROM")
											    iSerialTo = LotNode.getAttribute("SERIALTO")
											    iLotQty = LotNode.getAttribute("QTY")
											    iEntryCtr = LotNode.getAttribute("COUNTER")
											    iLotValue = trim(LotNode.getAttribute("IVALUE"))
											    sAutoGen = Trim(LotNode.getAttribute("AUTOGEN"))

											    if sitmType <> "GAR" and sRecNum <> "N" then
												    iAltGross = trim(LotNode.getAttribute("ALTGROSS"))
												    iAltNett = trim(LotNode.getAttribute("ALTNETT"))
												    sAltUoM = trim(LotNode.getAttribute("ALTUOM"))
												    sQtyIn = LotNode.getAttribute("QTYIN")
												    iTareQty = LotNode.getAttribute("TARE")
												    sTareIn = LotNode.getAttribute("TAREWEIGHT")
												    sStage = trim(LotNode.getAttribute("STAGE"))
												    sAttr = trim(LotNode.getAttribute("ATTLIST"))
											    elseif sitmType = "GAR" or sRecNum = "N" then
												    sAttr = trim(LotNode.getAttribute("ATTLIST"))
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

											    if iValue > 0 and iQty > 0 then
												    iItmRate = cdbl(iValue) / cdbl(iQty)
											    else
												    iItmRate = 0
											    end if

											    sSql = "DELETE FROM INV_T_ITEMLEDGER WHERE ITEMCODE =" & sitmCode & " AND TransactionType<>'RO'"
											    Con.execute sSql

											   ' Response.Write "*** iLot No = "& ilot
											   ' Response.Write "*** iSerial No = "& iSerialEntry
											   
											   
											   Dim sArrTempSerial,sTempSerial,sSQLQuery,rsTemp,sNumClassName
				                                set rsTemp = Server.CreateObject("ADODB.Recordset")
                                				
                                					
					                                If sAutoGen = "AUTO" and (sRecNum = "LS" or sRecNum = "L") Then
				                                '	Response.Write vbCrLf 
                                				
				                                        sSQLQuery = "Select GroupName from INV_M_Classification where GroupCode = "& iClass 
	                                                    rsTemp.Open sSQLQuery,con
	                                                    if not rsTemp.EOF then
	                                                        sNumClassName = Trim(rsTemp(0))
	                                                    end if
	                                                    rsTemp.Close 
                                	                    
                                                    	
				                                       sTempSerial = GetInvNumberSeriesCodes("LO",sOrgCode,iClass)
				                                       sArrTempSerial = Split(sTempSerial,":")
				                                       iSNo = sArrTempSerial(0)
				                                       iSCode = sArrTempSerial(1)
				                                       if Trim(iSNo)="0" and Trim(iSCode)="0" then
	                                                        Response.Clear 
	                                                        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Lot Entry - "& sNumClassName &"  Classification </H2></p>"
	                                                        Response.End 
	                                                    else
	                                                        iLotEntry = GenSeriesNumber(sOrgCode,iSNo,iSCode,FormatDate(sFinTo))
	                                                    end if
                                	                    
                                				            
                                				       	
						                                'Response.Write vbCrLf + vbCrLf + iLotEntry + vbCrLf
					                                else
					                                    iLotEntry =""
					                                End If
					
											   
											   

											    if sRecNum<>"N" then
												    sSql = "UPDATE INV_T_ITEMLEDGER SET ORGANISATIONCODE=" & Pack(sOrgCode) & ",CLASSIFICATIONCODE=" & iClass & ",TRANSACTIONTYPE='"& sTransType &"'," &_
													       "TRANSACTIONNO=" & iInvRecNo & ",TRANSACTIONDATE=CONVERT(DATETIME," & Pack(sReceivedOn) & ",103),TRANSACTQUANTITY=" & iLotQty & ", "&_
													       "TRANSACTVALUE=" & iLotValue & "	WHERE ITEMCODE =" & sitmCode & " AND TransactionType='RO' "
											    else
												    sSql = "UPDATE INV_T_ITEMLEDGER SET ORGANISATIONCODE=" & Pack(sOrgCode) & ",CLASSIFICATIONCODE=" & iClass & ",TRANSACTIONTYPE='"& sTransType &"'," &_
													       "TRANSACTIONNO=" & iInvRecNo & ",TRANSACTIONDATE=CONVERT(DATETIME," & Pack(sReceivedOn) & ",103),TRANSACTQUANTITY=" & iQty & ", "&_
													       "TRANSACTVALUE=" & iValue & "	WHERE ITEMCODE =" & sitmCode & " AND TransactionType='RO'"
											    end if
											    'Response.Write sSql & vbCrLf & vbCrLf
											    con.Execute sSql

											    if sBin = "NULL" then sBin = "0"
											    sExp1 ="//STOREDET/STORAGE [ @STORE = "&sLoc&" and @BIN = '"&sBin&"']/LotSerial [ @LOT = '"&iLot&"' and @COUNTER = "&iEntryCtr&" and @ATTLIST='"& sAttr &"']/LotSerialDetails"
											    Response.Write " >>>>>>>>>>>>sExp1 = "& sExp1
											    if sBin = "0" then sBin = "NULL"
											    Set LotSerialNode = RootNode.Selectnodes(sExp1)

											    iItemEntNo = 0

											    For iCounter = 0 to LotSerialNode.Length-1
										    'set LotSerialNode = LotNode
												    if Trim(sAutoGen)<>"AUTO" then iLotEntry = iLot
												    iItemEntNo	 = iItemEntNo + 1
												    'iSerialEntry = trim(LotSerialNode.getAttribute("LOTSERIAL"))
												    iSerialEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value)
												    Response.Write "sRecNum = "& sRecNum
												    
												    if sitmType = "GAR" or sRecNum = "N" then
													    iQtyRecEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("QTYREC").Value)
													    iAttributeList = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").Value)
													    
													    if trim(sBarCode) = "" or IsNull(sBarCode) then
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
														    iAttributeList = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").Value)
														    sAutoGenDet = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("AUTOGENDET").Value)
													    end if
												    end if
												    
												    if trim(iAttributeList)="" or IsNull(iAttributeList) or IsNull(iAttributeList)="NULL" then iAttributeList ="0"
												    'if trim(iAttributeList)<>"0" then iAttributeList=pack(iAttributeList)

												    if sRecNum <> "N" then
												        sPackingNumber = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKNUMBER").Value)
												    else
													    sPackingNumber = "NULL"
												    end if
												    'iLotValue = trim(LotSerialNode.getAttribute("IVALUE"))
												    'iLotValue = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value)
												    
												    
												    
												    If sAutoGenDet = "Y" and (sRecNum="S") Then
							                            Response.Write "Welcome to Packing Number Gen"
                            							    
						                                sTempSerial  =  GetPRDNumberSeriesCodes(sOrgCode,iClass)
					                                    sArrTempSerial = Split(sTempSerial,":")
					                                    iSNo = sArrTempSerial(0)
					                                    iSCode = sArrTempSerial(1)
                        						  
					                                    sSQLQuery = "Select GroupName from INV_M_Classification where GroupCode = "& iClass
                                                        rsTemp.Open sSQLQuery,con
                                                        if not rsTemp.EOF then
                                                            sNumClassName = Trim(rsTemp(0))
                                                        end if
                                                        rsTemp.Close 
                        	                    
                        	                    
                	                                    if Trim(iSNo)="0" and Trim(iSCode)="0" then
                                                            Response.Clear 
                                                            Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Packing Number - "& sNumClassName &"  Classification </H2></p>"
                                                            Response.End 
                                                        else
                                                            sPackingNumber = GenSeriesNumber(sOrgCode,iSNo,iSCode,FormatDate(sFinTo))										 			
                                                        end if
                                                        'Response.Write sPackingNumber 
			                                        End If
												    

												    if sSellingType = "select" or sSellingType = "" then
													    sSellingType = "0"
													    iWeight = "0"
												    end if
												    if trim(iWeight)="" then iWeight = "0"

												    if sPackingType = "select" or sPackingType = "" then
													    sPackingType = "0"
												    end if

												    if sSellingFormType = "select" or sSellingFormType = "" then
													    sSellingFormType = "0"
												    end if

												    if iLotEntry <> "N/A" and iSerialEntry = "0" then iSerialEntry = "NULL"
												    if iLotEntry = "N/A" or iLotEntry = "NULL" then
													    iLotEntry = "NULL"
												    else
													    iLotEntry = iLotEntry
												    end if
												    If iQtyRecEntry = "" or IsNull(iQtyRecEntry) then iQtyRecEntry = 0
												    IF iTareQtyEntry ="" or IsNull(iTareQtyEntry) then iTareQtyEntry = 0
												    if sQtyIn = "N" then
													    iQtyGross = cdbl(iQtyRecEntry) + cdbl(iTareQtyEntry)
													    iQtyNett = iQtyRecEntry
												    else
													    iQtyGross = iQtyRecEntry
													    iQtyNett = cdbl(iQtyRecEntry) - cdbl(iTareQtyEntry)
												    end if

												   ' Response.Write "<br> iQtyRecEntry = "& iQtyRecEntry & "<br>"
												   ' Response.Write "<br> iTareQtyEntry = "& iTareQtyEntry & "<br>"
												   ' Response.Write "<br> iQtyGross = "& iQtyGross & "<br>"
												   ' Response.Write "<br> iQtyNett = "& iQtyNett & "<br>"

												    if sItmType = "FAB" then iQtyNett = iQtyRecEntry
												    if iTareQtyEntry = "" then iTareQtyEntry = "0"

												    if sRecNum<>"N" then
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
												    else
													    iSerialEntry ="NULL"
												    end if


												    if iLotValue > 0 and iQtyNett > 0 then
													    iRate = FormatNumber(cdbl(iLotValue) / cdbl(iQtyNett),4,,,0)
												    else
													    iRate = 0
												    end if

												    if sBarCode = "" then sBarCode = "NULL"

												    if sBin = "NULL" then sBin = "0"

												    if trim(sBin)="0" or trim(sBin)="NULL" or trim(sBin)="N/A" or IsNull(sBin) then sBin = "NULL"

													        if iLBarCode = "" or IsNull(iLBarCode) then 	iLBarCode = 0
														    if iLSerQty = "" or IsNull(iLSerQty) then iLSerQty = 0

														    Set newElem2 = oDOM.createElement("LOT")
														    newElem2.setAttribute "ITEMENTNO",iItemEntNo
														    newElem2.setAttribute "ITEM",sitmCode
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
														    newElem2.setAttribute "ATTRIBUTE",iAttributeList

														    newElem1.appendChild newElem2
														    'Response.Write "<BR>====="&sDateOfRec&"=====<BR>"
														    if trim(iLLotNo)="0" then iLLotNo ="N/A"
													    
														    if (trim(iLLotNo)="N/A" or trim(iLLotNo)="NULL" or trim(iLLotNo)="0" or trim(iLLotNo)="" or IsNull(iLLotNo) ) and  (trim(iSerialEntry)="0" or trim(iSerialEntry)="NULL") then
															    sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE,STORAGELOCATIONNO," &_
																    "STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE," &_
																    "SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE,BARCODETAG,ATTRIBUTELIST,SRCTYPE,DATEOFRECEIPT,ITEMENTRYNO,SRCTYPECODE) VALUES " &_
																    "(" & iInvRecNo & "," & Pack(sOrgCode) & "," & sitmCode & "," & iClass& "," & sLoc & "," & sBin & ",'" & iLLotNo & "'," &_
																    "" & iSerialEntry & "," & iQtyGross & ","& iQtyNett & ",0,NULL," & sPackingType & "," &_
																    "" & sSellingType & ",0," & sSellingFormType & "," & sStage & "," & cdbl(iRate) & ",NULL, "&_
																    "" & pack(iAttributeList) & ",'"& sTransType &"',CONVERT(DATETIME,"& pack(sReceivedOn) &",103)," & iItemEntNo & "," & iSrcTypeCode & " )"
															    'Response.write sSql & vbCrLf & vbCrLf
														    else
															    sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE,STORAGELOCATIONNO," &_
																    "STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE," &_
																    "SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE,BARCODETAG,ATTRIBUTELIST,SRCTYPE,DATEOFRECEIPT,ITEMENTRYNO,SRCTYPECODE) VALUES " &_
																    "(" & iInvRecNo & "," & Pack(sOrgCode) & "," & sitmCode & "," & iClass& "," & sLoc & "," & sBin & ",'" & iLot & "'," &_
																    "" & iSerialEntry & "," & iQtyGross  & ","& iQtyNett & "," & iTareQtyEntry & ",'" & sPackingNumber & "'," & sPackingType & "," &_
																    "" & sSellingType & "," & iWeight & "," & sSellingFormType & "," & sStage & "," & cdbl(iRate) & "," & iLBarCode & ", "&_
																    "" & pack(iAttributeList) & ",'"& sTransType &"',CONVERT(DATETIME,"& pack(sReceivedOn) &",103)," & iItemEntNo & "," & iSrcTypeCode & " )"
															    'Response.write sSql & vbCrLf & vbCrLf
														    end if
														    Response.write mid(sSql,200) & vbCrLf & vbCrLf
														    con.Execute sSql


											    next
										    next
								    else

									    'Response.Write "Else part"
										    if iQty <> 0 then
											    iItmRate = FormatNumber(cdbl(iValue) / cdbl(iQty),4,,,0)
										    else
											    iItmRate = 0
										    end if

										    with dcrs
											    .CursorLocation = 3
											    .CursorType = 3
											    .Source  = "SELECT  ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
											    .ActiveConnection = con
											    .Open
										    end with
										    set dcrs.ActiveConnection = nothing

										    if not dcrs.EOF then
											    iLedEntNo  = dcrs(0)
										    end if
										    dcrs.Close

										    'sExp ="//STOREDET/STORAGE/LotSerial/LotSerialDetails"

									    '	Set LotNode1 = RootNode.selectSingleNode(sExp)

										    for each StockNode in RootNode.childNodes
											    for each sstorageNode in StockNode.childNodes
											    iItemEntNo = 0
											    iLCtr = 0
												    for each LotNode in sstorageNode.childNodes

													    if strcomp(LotNode.nodeName,"LotSerialDetails")=0 then
														    set LotNode1 = LotNode
													    iItemEntNo = iItemEntNo + 1
													    iLLotNo		= trim(LotNode1.getAttribute("LOT"))
													    iLSerQty	= trim(LotNode1.getAttribute("SERIALQTY"))
													    iLPackNo	= trim(LotNode1.getAttribute("PACKNUMBER"))
													    iLVal		= trim(LotNode1.getAttribute("IVALUE"))
													    iLAttbID	= trim(LotNode1.getAttribute("ATTRIBUTEID"))
													    iLAttbList	= trim(LotNode1.getAttribute("ATTRIBUTELIST"))
													    iLBarCode	= trim(LotNode1.getAttribute("BARCODE"))
													    'Response.Write "iLAttbList="&iLAttbList & vbCrLf
													    if trim(iLLotNo)="0" then iLLotNo ="N/A"
													    if trim(sBin)="0" or trim(sBin)="NULL" or trim(sBin)="N/A" or IsNull(sBin) then sBin = "NULL"
													    'if trim(iLLotNo)="N/A" then iLLotNo="NULL"
													    sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE,STORAGELOCATIONNO," &_
														    "STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE," &_
														    "SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE,BARCODETAG,ATTRIBUTELIST,SRCTYPE,DATEOFRECEIPT,ITEMENTRYNO,SRCTYPECODE) VALUES " &_
														    "(" & iInvRecNo & "," & Pack(sOrgCode) & "," & sitmCode & "," & iClass& "," & sLoc & "," & sBin & "," & iLLotNo & "," &_
														    "" & iSerialEntry & "," & iLSerQty & ","& iLSerQty & "," & iTareQtyEntry & ",'" & iLPackNo & "'," & sPackingType & "," &_
														    "" & sSellingType & "," & iWeight & "," & sSellingFormType & "," & sStage & "," & iLVal & "," & iLBarCode & ", "&_
														    "'" & iLAttbList  & "','"& sTransType &"',CONVERT(DATETIME,'"&sReceivedOn&"',103)," & iItemEntNo & "," & iSrcTypeCode & ")"
													    'objTxt.write sSql & vbCrLf & vbCrLf
													    Response.write sSql & vbCrLf & vbCrLf
													    con.Execute sSql

													    sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
															    "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,SRCTYPECODE) VALUES " &_
															    "(" & iLedEntNo & "," & Pack(sOrgCode) & "," & sitmCode & "," & iClass & "," &_
															    "'"& sTransType &"',NULL,CONVERT(DATETIME," & pack(sReceivedOn) & ",103)," & iQty & "," & iValue & "," & iSrcTypeCode & " )"
															    'objTxt.write sSql & vbCrLf & vbCrLf
															   ' Response.Write sSql & vbCrLf & vbCrLf
														    con.Execute sSql
													    end if
													    iLCtr = iLCtr + 1
												    next
											    next
										    next
								    end if'if strcomp(sLoc,StorageNode.getAttribute("STORE"))=0 and strcomp(sLoc,StorageNode.getAttribute("BIN"))=0 and strcomp(iClass,StorageNode.getAttribute("CLASSIFICATION"))=0 and strcomp(sOrgCode,StorageNode.getAttribute("UNIT"))=0 then
							    next
						next

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				arrFin = split(GetFinancialYear(sMonYr),":")
				sFinFrom = arrFin(0)
				sFinTo = arrFin(1)
				
				
				
            sSql = "Update INV_T_ITEMYEARLYSTOCK set YEAROPENINGSTOCK = "& iQty &",YEAROPENINGVALUE = "& iValue &","&_
	               " YEARCLOSINGSTOCK= "& iQty &"+YearReceiptQuantity-YearIssueQuantity ,"&_
	               " YEARCLOSINGVALUE= "& iValue &"+YearReceiptValue-YearIssueValue where "&_
	               " OrganisationCode = "& Pack(sOrgCode) &" and ClassificationCode = "& iClass &" and "&_
	               " ItemCode = "& sitmCode &" and FinancialYearFrom= CONVERT(DATETIME," & Pack(sFinFrom) & ",103) and"&_
	               " FinancialYearTo =CONVERT(DATETIME," & Pack(sFinTo) & ",103) "
	        'Response.write sSql & vbCrLf & vbCrLf
	        con.execute sSql

            
			Response.Write "iSerialEntry = "&iSerialEntry &"+++++++++"

			if trim(iSerialEntry)="0" or IsNull(iSerialEntry) or trim(iSerialEntry)="" then
			    iRate = 0
			    if iQty > 0 and iValue > 0 then
			    iRate = FormatNumber(cdbl(iValue)/cdbl(iQty),2,0,0,0)
			    end if

			    sSql = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE,STORAGELOCATIONNO," &_
		        "STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE," &_
		        "SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE,BARCODETAG,ATTRIBUTELIST,SRCTYPE,DATEOFRECEIPT,ITEMENTRYNO,SRCTYPECODE) VALUES " &_
		        "(" & iInvRecNo & "," & Pack(sOrgCode) & "," & sitmCode & "," & iClass& "," & sLoc & "," & sBin & ",NULL," &_
		        "NULL," & iQty & ","& iQty & ",0,NULL,NULL,NULL,NULL,NULL,NULL," & iRate & ",NULL, "&_
		        "NULL,'"& sTransType &"',CONVERT(DATETIME,"& pack(sReceivedOn) &",103)," & iItemEntNo & "," & iSrcTypeCode & " )"
	            Response.write sSql & vbCrLf & vbCrLf
	            con.Execute sSql

	        end if 'if trim(iSerialEntry)="0" or IsNull(iSerialEntry) then


			'if trim(sBin)="0" then sBin = "NULL"
			
			if trim(sBin)="NULL" then sBin="0"
			
			
			        sSql = "Update INV_T_ITEMLOCATIONSTOCK set YEAROPENINGSTOCK = "& iQty &",YEAROPENINGVALUE = "& iValue &","&_
			               " YEARCLOSINGSTOCK= "& iQty &"+YearReceiptQuantity-YearIssueQuantity ,"&_
			               " YEARCLOSINGVALUE= "& iValue &"+YearReceiptValue-YearIssueValue where "&_
			               " OrganisationCode = "& Pack(sOrgCode) &" and ClassificationCode = "& iClass &" and "&_
			               " ItemCode = "& sitmCode &" and FinancialYearFrom= CONVERT(DATETIME," & Pack(sFinFrom) & ",103) and"&_
			               " FinancialYearTo =CONVERT(DATETIME," & Pack(sFinTo) & ",103) and LocationNumber = "& sLoc &" and IsNull(BinNumber,0)="& sBin
			        Response.Write "<p>"& sSql
			        con.execute sSql
			
				next ' for each StoNode in StoDetNode.childNodes
				end if 'if StoDetNode.nodeName="STOREDET" then
			next ' for each StoDetNode in RootNode.childNodes
		next ' for j =0 to 0

		Set newElem = oDOM.createElement("ITEM")
		newElem.setAttribute "ITEMENTRYNO",iItemEntNo
		newElem.setAttribute "ITEM",sitmCode
		newElem.setAttribute "CLASS",iClass
		newElem.setAttribute "ITEMQTY",iItemTotalQty
		newElem.setAttribute "ITEMVALUE",iItemTotalValue
		newElem.setAttribute "ATTRIBUTE",sAttribute
		newElem.setAttribute "SUMQTY",iItemTotalQty

	'	StockNode.appendChild newElem
'Response.Write "iSerialEntry = "& iSerialEntry

	'Response.clear
	'	 Response.end
	
	if sStatus = "AC" then
        sSql = "UPDATE INV_M_ITEMMASTER SET ITEMACTIVE = 'Y' WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
       ' Response.Write "<p>"&sSql
        con.execute sSql
        
        sSql = "UPDATE INV_M_ITEMMASTER SET ITEMONHOLD = 0 WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
        sSql = "UPDATE INV_M_ITEMMASTER SET DeadStock='N' WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
        
        sSql = "INSERT INTO INV_T_ONHOLDDETAILS (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE,"&_
                 "ONHOLDREASON,HOLDRELEASEDON,HOLDRELEASEDBY) VALUES ("& sitmCode &","& Pack(sOrgCode) &","&_
                 ""& iClass &",NULL,CONVERT(DATETIME,GETDATE(),103),"& iitmController &")"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
    elseif sStatus = "IA" then
        sSql = "UPDATE INV_M_ITEMMASTER SET ITEMACTIVE = 'N' WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
        sSql = "UPDATE INV_M_ITEMMASTER SET ITEMONHOLD = 0 WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
        sSql = "UPDATE INV_M_ITEMMASTER SET DeadStock='N' WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
        sSql = "INSERT INTO INV_T_ONHOLDDETAILS (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE,"&_
                 "ONHOLDREASON,HOLDRELEASEDON,HOLDRELEASEDBY) VALUES ("& sitmCode &","& Pack(sOrgCode) &","&_
                 ""& iClass &",NULL,CONVERT(DATETIME,GETDATE(),103),"& iitmController &")"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
    elseif sStatus = "OH" then
        sSql = "UPDATE INV_M_ITEMMASTER SET ITEMONHOLD = 1 WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        sSql = "UPDATE INV_M_ITEMMASTER SET DeadStock='N' WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
        sSql = "INSERT INTO INV_T_ONHOLDDETAILS (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE,"&_
                 "ONHOLDREASON,HOLDRELEASEDON,HOLDRELEASEDBY) VALUES ("& sitmCode &","& Pack(sOrgCode) &","&_
                 ""& iClass &",NULL,CONVERT(DATETIME,GETDATE(),103),"& iitmController &")"
'        Response.Write "<p>"&sSql
        con.execute sSql
    elseif sStatus = "DS" then
        sSql = "UPDATE INV_M_ITEMMASTER SET DeadStock='Y' WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
    end if' if sStatus = "AC" then
    
    
    sSql = "UPDATE INV_M_ITEMMASTER SET PurTaxType="& sPurTaxType &",SalTaxType = "& sSalTaxType &" WHERE ITEMCODE ="& sitmCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
'        Response.Write "<p>"&sSql
        con.execute sSql
        
        
    sSql = "Delete from Inv_M_ItemOrgAccountHead where ItemCode = "&sItmCode &" and ClassificationCode = "& iClass &" and OrganisationCode ="& pack(sOrgCode)
    con.execute sSql
        
    Dim sName,sType,sValue,iCnt
	sExp ="//CONTROLS/ACCHEAD"
	Set Node = RootNode.selectNodes(sExp)
	if Node.length > 0 then
	    For iCnt = 0 to Node.length -1
	        sType = trim(Node.Item(iCnt).Attributes.getNamedItem("Type").value)
		    sValue = trim(Node.Item(iCnt).Attributes.getNamedItem("Value").value)
		    
		    if trim(sValue)="" or IsNull(sValue) then sValue="NULL"
		    
	        sSql = "Insert into Inv_M_ItemOrgAccountHead (ItemCode,ClassificationCode,OrganisationCode,AccountHead,AccountHeadFor,AccountHeadType)"
	        sSql = sSql &"  Values("& sitmCode &","& iClass &","& Pack(sOrgCode)&","& sValue &","& pack(sType) &",'N')"
	        con.execute sSql
		Next
	end if
	
	sSql = "Delete from INV_M_ItemMasterAttributes where ItemCode = "& sItmCode &" and ClassificationCode = "& iClass &" and OrganisationCode = "& pack(sOrgCode)
	con.execute sSql
	
	Dim iTypeHeadId,iAttId,sAttName,sAttValue
	    
	    sExp = "//ATTRIBUTE/ATTDET"
	    set ndAttNode = RootNode.selectNodes(sExp)
'	    Response.Write "<p>ndAttNode.length = "& ndAttNode.length
	    if ndAttNode.length>0 then
	        For iCnt = 0 to ndAttNode.length -1
	            iAttId = ndAttNode.Item(iCnt).Attributes.getNamedItem("AttID").value
	            iTypeHeadId = ndAttNode.Item(iCnt).Attributes.getNamedItem("Head").Value
	            sAttValue = ndAttNode.Item(iCnt).Attributes.getNamedItem("Value").value
	            
	            sSql = "Insert into INV_M_ItemMasterAttributes (ItemCode,ClassificationCode,CategoryCode,OrganisationCode,"&_
	                   " HeaderID,ItemTypeAttributeID,AttributeValue)"&_
	                   " values("& sItmCode &","& iClass &","& sCategory &","& Pack(sOrgCode) &","& iTypeHeadId &","& iAttId &","& Pack(sAttValue) &")"
	           ' Response.Write "<p>"& sSql
	            con.execute sSql
	        Next
	    end if
	    
	set Node = nothing
	Dim Objrs,rs
	Dim sPurRate,sChaPer,sChaVal,sMarPer,sMarVal,sTotPrice,sEffDate,sPurRatPer,sAsonDate,iSellno,sRecno
	Dim sQuery,sHisNo,sQry2,sPurRatePer
	
	set objrs = server.createObject("ADODB.Recordset")
	set rs = server.createObject("ADODB.Recordset")
	
	sExp = "//PRICING"
	Set Node = RootNode.selectNodes(sExp)
	if Node.length > 0 then
        sPurRate = Node.Item(0).Attributes.getNamedItem("PURRATE").value
        sPurRatePer = Node.Item(0).Attributes.getNamedItem("PURRATEPER").value
        sChaPer = Node.Item(0).Attributes.getNamedItem("CHARPER").value
        sChaVal = Node.Item(0).Attributes.getNamedItem("CHARVAL").value
        sMarPer = Node.Item(0).Attributes.getNamedItem("MARPER").value
        sMarVal = Node.Item(0).Attributes.getNamedItem("MARVAL").value
        sTotPrice = Node.Item(0).Attributes.getNamedItem("TOTPRICE").value
        sEffDate = Node.Item(0).Attributes.getNamedItem("EFFFROM").value
        
        sAsonDate = sEffDate
        
	                sQuery = "SELECT SellingPriceno FROM Sal_M_UnitPriceHdr"
	
	                Objrs.Open sQuery,Con
	                IF Not Objrs.EOF Then
		                iSellno = objrs(0)
	                Else
		                iSellno = "0"
	                End IF
	                Objrs.Close


	                IF CStr(iSellno) = "0" Then
                		
		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = "Select isNull(Max(SellingPriceno),0) + 1 From Sal_M_UnitPriceHdr"
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                IF Not Objrs.EOF Then
			                iSellno = Objrs(0)
		                End IF
		                Objrs.Close

		                    sQuery = " INSERT INTO Sal_M_UnitPriceHdr (SellingPriceno, AsonDate, CurrencyCode, "&_
				                     " PackingType, Price, UnitPrice,UpdatedBy,UpdatedOn) "&_
				                     " VALUES ("&iSellno&", Convert(datetime,'"&sAsonDate&"',103),1,0,'G','S',"&_
				                     " "& iitmController&",Convert(DateTime,GetDate(),103)) "
		                   ' Response.Write sQuery
		                    con.execute sQuery

                			sQuery = "INSERT INTO Sal_M_UnitPriceDet (SellingPriceno, OudefinitionID,Itemcode, Classificationcode,ItemRate,RatePer,MarginPercent,MarginValue,OtherPercent,OtherValue,ItemPrice,EffectiveFrom) "
			                sQuery = sQuery &"VALUES ("&iSellno&",'"& sOrgCode &"', "&sItmCode&", "&iClass&", "&sPurRate&","& sPurRatePer &","& sMarPer&","& smarValue&","& sChaPer&" ,"& sChaVal &" ,"& sTotPrice &" ,Convert(DateTime,'"& sEffDate &"',103)) "
			               ' Response.Write "<p>query="&sQuery
			                Con.Execute sQuery
			            
	                '=========================== Amendment Part Starts Here ====================================
	                Else
		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = "Select isNull(Max(HistoryNo),0) + 1 From Sal_M_HistoryUnitPriceHdr "
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                IF Not Objrs.EOF Then
			                sHisno = Objrs(0)
		                End IF
		                Objrs.Close

		                sQuery = "SELECT ItemTypeID,isNull(UpdatedBy,0),Convert(Varchar,isNull(updatedOn,''),103),SellingPriceNo FROM Sal_M_UnitPriceHdr Where SellingPriceno = "&iSellno
		                'Response.write sQuery
		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = sQuery
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                IF Not Objrs.EOF Then
			                
			                sQuery = "Select HistoryEntryNo From SAL_M_UnitPriceHdrHistory Where HistoryEntryNo="&sHisno
			                rs.Open sQuery,con
			                If Not rs.EOF Then
				                sQry2 = " Update SAL_M_UnitPriceHdrHistory Set updatedOn = Convert(datetime,'"&Objrs(2)&"',103),updatedBy = "&Objrs(1)&" ,ModifiedBy="& iitmController&",ModifiedOn=Convert(datetime,getdate(),103)  Where HistoryEntryNo = "&sHisno
			                Else
				                sQry2 = "INSERT INTO SAL_M_UnitPriceHdrHistory (HistoryEntryNo, ItemTypeID,updatedOn,updatedBy,ModifiedBy,ModifiedOn) "
				                sQry2 = sQry2 &"VALUES   ("&sHisno&", '"&Objrs(0)&"', Convert(datetime,'"&Objrs(2)&"',103),"&Objrs(1)&","& iitmController &",Convert(datetime,getdate(),103)) "
			                End IF
			                rs.Close 
			               ' Response.Write "<p>Query1="&sQry2
			                Con.Execute sQry2
		                End IF
		                Objrs.Close
                		
		                'sQry2 = "SELECT Itemcode, Classificationcode, OudefinitionID, Price FROM Sal_M_UnitPriceDet "
		                'sQry2 = sQry2 &"Where SellingPriceno = "&iSellno
                		
		                sQry2 = " Select SellingPriceNo,OudefinitionID,Itemcode,Classificationcode,ItemRate,isNull(MarginPercent,0),"&_
				                " isNull(MarginValue,0),isNull(OtherPercent,0),isNull(OtherValue,0),isNull(ItemPrice,0),Convert(Datetime,"&_
				                " isNull(EffectiveFrom,''),103),isNull(RatePer,0),isNull(RateUOM,'') From Sal_M_UnitPriceDet "&_
				                " Where SellingPriceNo = "& iSellno

		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = sQry2
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                Do While Not Objrs.EOF
			                'sQry2 = "INSERT INTO Sal_M_HistoryUnitPriceDet (HistoryNo, SellingPriceno, Itemcode, Classificationcode, OudefinitionID, Price) "
			                'sQry2 = sQry2 &"VALUES ("&sHisno&", "&iSellno&", "&Objrs(0)&", "&Objrs(1)&", '"&Objrs(2)&"', "&Objrs(3)&") "
			                sQry2 = " INSERT INTO SAL_M_UnitPriceDetHistory (HistoryEntryNo, SellingPriceno,OudefinitionID, Itemcode, Classificationcode, "&_
					                " ItemRate,MarginPercent,MarginValue,OtherPercent,OtherValue,ItemPrice,EffectiveFrom,RatePer,RateUOM) Values ( "&sHisno&", "&iSellno&","&_
					                " "&Objrs(1)&","&Objrs(2)&","&Objrs(3)&","&Objrs(4)&","&Objrs(5)&","&Objrs(6)&","&Objrs(7)&","&Objrs(8)&","&Objrs(9)&","&_
					                " "&Objrs(10)&","& objrs(11)&",'"& objrs(12)&"') "
			               ' Response.Write "<p>Query2="&sQry2
			                Con.Execute sQry2
			                Objrs.MoveNext
		                Loop
		                Objrs.Close

		                sQry2 = " UPDATE Sal_M_UnitPriceHdr SET AsonDate = Convert(datetime,'"&sAsonDate&"',103), "&_
				                " UpdatedBy = '"& iitmController &"',UpdatedOn = Convert(datetime,getDate(),103) Where SellingPriceno = "&iSellno&" "
                		
		                Con.Execute sQry2
                		
		                		
				                sQuery = " UPDATE Sal_M_UnitPriceDet SET OudefinitionID= '"& sOrgCode &"' ,ItemRate = "&sPurRate&",MarginPercent = "& sMarPer &" ,MarginValue = "& sMarVal &",OtherPercent = "& sChaPer&" , "&_
						                 " OtherValue = "& sChaVal &" ,ItemPrice = "& sTotPrice &" , EffectiveFrom = Convert(DateTime,'"& sEffDate &"',103), "&_
						                 " RatePer = "& sPurRatePer &" "&_
						                 " WHERE SellingPriceno = "&iSellno&" AND Itemcode = "&sItmcode&" AND Classificationcode = "&iClass&" "
				               ' Response.Write "<p>Qry3="&sQuery
				                'Response.Write "<br><br><br>"
				                Con.execute sQuery,sRecno
				                IF sRecno = 0 Then
					                sQuery = "INSERT INTO Sal_M_UnitPriceDet (SellingPriceno, OudefinitionID,Itemcode, Classificationcode,ItemRate,MarginPercent,MarginValue,OtherPercent,OtherValue,ItemPrice,EffectiveFrom,RatePer) "
					                sQuery = sQuery &"VALUES ("&iSellno&",'"& sOrgCode &"', "&sItmCode&", "&iClass&", "&sPurRate&", "& sMarPer &","& sMarVal &","& sChaPer &" ,"& sChaVal &" ,"& sTotPrice &" ,Convert(DateTime,'"& sEffDate &"',103),"& sPurRatePer &") "
					               ' Response.Write "<p>Squer4="&sQuery
					                Con.Execute sQuery
				                End IF
	                End IF
	end if 'if Node.length > 0 then
	    



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if con.Errors.count <> 0 then
		dim iErrCounter
		con.RollbackTrans
		for iErrCounter=0 to con.Errors.count
			Response.Write con.Errors(iErrCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else
		if objfs.FileExists(Server.MapPath("../temp/master/Attribute"&Session.SessionID&".xml")) then
			objfs.DeleteFile server.MapPath("../temp/master/Attribute"&Session.SessionID&".xml")
		end if
	'	con.RollbackTrans
	'	Response.End
	    Response.Clear 
	 	con.CommitTrans
	end if

	con.close
	set con = nothing

%>
