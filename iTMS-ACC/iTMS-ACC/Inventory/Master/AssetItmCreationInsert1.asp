<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AssetItmCreationInsert.asp
	'Module Name				:	Inventory (Asset Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 03, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	AssetItmCreationEntry.asp
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
<!-- #include File="../../include/InventoryAccountingUpdate.asp" -->
<%
	'XML DOM Variables
	dim newxml,RootNode,StorageNode,BoMNode,AssetNode,objfs,objTxt

	dim dcrs,dcrs1,sSql,sExp,iDefinedBy,adoCmd,sExp1,sExp2
	dim sClassCode,iItmCode,sOrgCode,arrStorage,sUoM,sAssetCode,sTempClass
	dim iBoMClassCode,iBoMItemCode,sType,sItmCode,sApplicableFor
	dim sValue,arrValue,sTemp,iCtr,sLoc,sBin,sMonYr,iQty,iValue,sAppli
	dim arrFin,sFinFrom,sFinTo,arrTemp,iInvRecNo,iCounter,sTempMonYr
	dim sItemType,sUoMDesc,sItmDesc,sQtyIn,iTareQty,iLot
	dim iItmRate,iCtrAsset,iCtrBoM,iVar,jVar,isize,iCount,arrTempStorage
	dim iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sSellingType,iWeight,sPackingType
	dim iQtyGross,iQtyNett,iRate,iClassCode

	' Create our DOM Document Objects
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs  = CreateObject("Scripting.FileSystemObject")

	' has to be changed
	iDefinedBy = getUserid

	newxml.async = false
	newxml.load(Request)

	Set RootNode = newxml.documentElement

	sOrgCode		= trim(RootNode.Attributes.getNamedItem("ORG").Value)
	sAssetCode		= trim(RootNode.Attributes.getNamedItem("ASSETCODE").Value)
	sItmCode		= trim(RootNode.Attributes.getNamedItem("ITMCODE").Value)
	sTempClass		= trim(RootNode.Attributes.getNamedItem("CLACODE").Value)
	sItmDesc		= trim(RootNode.Attributes.getNamedItem("DESC").Value)
	sUoM			= trim(RootNode.Attributes.getNamedItem("SUOM").Value)
	arrTempStorage	= split(trim(RootNode.Attributes.getNamedItem("STORAGE").Value),"|")
	sApplicableFor	= trim(RootNode.Attributes.getNamedItem("APPLI").Value)

	if trim(sApplicableFor) = "" then
		sApplicableFor = "NULL"
	else
		sApplicableFor = Pack(sApplicableFor)
	end if

	con.beginTrans

	Set dcrs  = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(ITEMCODE),0) + 1 FROM INV_M_ITEMMASTER WHERE (ITEMCODE = (SELECT ISNULL(MAX(ITEMCODE), 0) FROM INV_M_ITEMMASTER))"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iItmCode = trim(dcrs(0))
	end if
	dcrs.close
	
	
	arrTemp = split(sTempClass,":")
	isize = CInt(UBound(arrTemp))

	sClassCode = arrTemp(isize)
	
	sSql = "INSERT INTO INV_M_ITEMMASTER (ITEMCODE,COMPANYITEMCODE,ITEMDESCRIPTION,PURCHASEELIGIBLE,MANUFACTUREELIGIBLE, " &_
		"SALESELIGIBLE,INVENTORYELIGIBLE,STORESUOM,PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR, " &_
		"MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE,SALETOSTOREOPERATOR, " &_
		"ITEMCREATEDON,ITEMCREATEDBY,ITEMCONTROLLER,ITEMDELETED,ITEMSOURCE,STOCKNONSTOCK,ITEMACTIVE,ELIGIBLEFOR,"&_
		"ClassificationCode,OrganisationCode,RECEIPTNUMBERING,ITEMDEFINEDBY,ITEMDEFINEDON,VALUATIONMETHOD,"&_
		"ALLOWLOCATIONTRANSFER,ALLOWINTERUNITTRANSFER) VALUES " &_

		"(" & iItmCode & "," & Pack(sItmCode) & "," & Pack(sItmDesc) & ",'0','0','0','1'," & Pack(sUoM) & ",NULL,NULL,'0'," &_
		" NULL,NULL,'0',NULL,NULL,'0',CONVERT(DATETIME,GETDATE(),103)," &_
		" "& iDefinedBy & "," & iDefinedBy & ",'0','O','N','Y'," & sApplicableFor & ",'"& sClassCode&"','"&sOrgCode&"',"&_
		" 'LS'," & iDefinedBy & ",CONVERT(DATETIME,GETDATE(),103),'W',1,1)"
			
	'objTxt.Write sSql & vbCrLf & vbCrLf
	Response.Write sSql
	con.Execute sSql


	For jVar = 2 To UBound(arrTemp)
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CLASSIFICATIONCODE FROM INV_M_ITEMGROUP WHERE CLASSIFICATIONCODE = " & trim(arrTemp(jVar)) & " AND ITEMCODE = " & iItmCode & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		If dcrs.EOF Then
			sSql = "INSERT INTO INV_M_ITEMGROUP(ITEMCODE,CLASSIFICATIONCODE,LEAFNODE) VALUES (" & iItmCode & "," & trim(arrTemp(jVar)) & ",0)"
			con.Execute sSql
		end if
		dcrs.Close

		if isize = jVar then
			sSql = "UPDATE INV_M_ITEMGROUP SET LEAFNODE = 1,ITEMPATH = '" & sTempClass & "' WHERE CLASSIFICATIONCODE = " & trim(arrTemp(jVar)) & " AND ITEMCODE = (SELECT MAX(ITEMCODE) FROM INV_M_ITEMMASTER)"
			con.Execute sSql
		end if
	next

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT ITEMCODE FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
		.Source = "SELECT ITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	
	if dcrs.EOF then
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT * FROM INV_M_ITEMGROUP WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & ""
			.ActiveConnection = con
			.Open
		end with

		if not dcrs1.EOF then
			sSql = "INSERT INTO INV_M_ITEMORGNGROUP (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
				"LEAFNODE,ITEMPATH) VALUES " &_
				"(" & iItmCode & "," & Pack(sOrgCode) & "," & sClassCode & "," &_
				"" & trim(dcrs1(2)) & "," & Pack(trim(dcrs1(3))) & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs1.Close


		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PURCHASEELIGIBLE,MANUFACTUREELIGIBLE,SALESELIGIBLE,STORESUOM,ITEMCREATEDBY FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItmCode & ""
			.ActiveConnection = con
			.Open
		end with

		if not dcrs1.EOF then
			If 1 = 2 Then
				sSql = "INSERT INTO INV_M_ITEMORGMASTER (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,COMPANYITEMCODE,SHORTDESCRIPTION," &_
					"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,PURCHASEELIGIBLE,MANUFACTUREELIGIBLE, " &_
					"SALESELIGIBLE,STORESUOM,PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR, " &_
					"MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE,SALETOSTOREOPERATOR, " &_
					"ITEMONHOLD,ITEMCONTROLLER,ITEMDEFINEDON,ITEMDEFINEDBY,ITEMDELETEDBY) VALUES " &_
					"(" & iItmCode & "," & sClassCode & "," & Pack(sOrgCode) & ",NULL,NULL," &_
					"" & Pack(sItmDesc) & ",NULL," & Pack(trim(dcrs1(0))) & "," & Pack(trim(dcrs1(1))) & "," &_
					"" & Pack(trim(dcrs1(2))) & "," & Pack(trim(dcrs1(3))) & ",NULL,NULL,'0'," &_
					"NULL,NULL,'0',NULL,NULL,'0'," &_
					"0," & trim(dcrs1(4)) & ",CONVERT(DATETIME,GETDATE(),103)," & trim(dcrs1(4))& ",0)"
				'objTxt.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			End IF
			
			sSql = " UPDATE INV_M_ITEMMASTER SET PURCHASEELIGIBLE = " & Pack(trim(dcrs1(0))) & ",MANUFACTUREELIGIBLE=" & Pack(trim(dcrs1(1))) & " ,"&_
				   " SALESELIGIBLE = " & Pack(trim(dcrs1(2))) & ",STORESUOM = " & Pack(trim(dcrs1(3))) & ", "&_
				   " ITEMCONTROLLER = " & trim(dcrs1(4)) & ",ITEMDEFINEDON = CONVERT(DATETIME,GETDATE(),103),"&_
				   " ITEMDEFINEDBY = " & trim(dcrs1(4))& " WHERE ITEMCODE="&iItmCode&" AND CLASSIFICATIONCODE='"&sClassCode &"'

		end if
		dcrs1.Close
	end if
	dcrs.Close

	' Purchase and Receiving Controls Insert
	' Inventory Controls Insert
	' Storage Location Details Insert

	for iCtr =  0 to ubound(arrTempStorage) - 1
		arrStorage = split(arrTempStorage(iCtr),"-")

		sLoc = arrStorage(0)
		sBin = arrStorage(1)
		sAppli = arrStorage(2)

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT COUNT(TAGNUMBER) FROM FAR_T_ASSETDETAILS WHERE ASSETDESCID = " & sAssetCode & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iQty = dcrs(0)
		end if
		dcrs.Close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT SUM(SUPPLIERINVOICEVALUE) FROM FAR_T_ASSETSUPPLIER WHERE ASSETNUMBER IN (SELECT ASSETNUMBER FROM FAR_T_ASSETDETAILS WHERE ASSETDESCID = " & sAssetCode & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iValue = dcrs(0)
		end if
		dcrs.Close

		sQtyIn = "N"
		iTareQty = "0"
		iLot = "NULL"
	
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_RECEIPTDETAILS"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iInvRecNo = dcrs(0)
		end if
		dcrs.Close

		iItmRate = cdbl(iValue) / cdbl(iQty)

		sSql = "INSERT INTO INV_T_RECEIPTDETAILS (INVENTORYRECEIPTNO,ORGANISATIONCODE," &_
			"ITEMCODE,CLASSIFICATIONCODE,RECEIPTQUANTITY,ACCEPTQUANTITY,ITEMRATE," &_
			"QUANTITYTYPE,SRCTYPE,ACCOUNTEDON,ACCOUNTEDBY,RECEIVEDON) VALUES " &_
			"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItmCode & "," & sClassCode & "," &_
			"" & iQty & "," & iQty & "," & iItmRate & ",'N','RO'," &_
				"CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103)," & iDefinedBy & ",CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103))"
		'objTxt.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
			"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
			"(" & Pack(sOrgCode) & "," & iItmCode & "," & sClassCode & "," &_
			"'RO'," & iInvRecNo & ",CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103)," & iQty & "," & iValue & ")"
		'objTxt.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT TAGNUMBER FROM FAR_T_ASSETDETAILS WHERE ASSETDESCID = " & sAssetCode & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				iLotEntry = Pack(dcrs(0))
				iSerialEntry = "NULL"
				iQtyRecEntry = "1"
				iTareQtyEntry = "0"

				sSellingType = "NULL"
				iWeight = "NULL"
				sPackingType = "NULL"

				iQtyGross = cdbl(iQtyRecEntry) + cdbl(iTareQtyEntry)
				iQtyNett = iQtyRecEntry

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_RECEIPTLOTDETAILS"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					iSerialEntry = dcrs1(0)
				end if
				dcrs1.Close

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT SUPPLIERINVOICEVALUE FROM FAR_T_ASSETSUPPLIER WHERE ASSETNUMBER IN (SELECT ASSETNUMBER FROM FAR_T_ASSETDETAILS WHERE ASSETDESCID = " & sAssetCode & " AND TAGNUMBER = " & Pack(dcrs(0)) & ")"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					iRate = dcrs1(0)
				end if
				dcrs1.Close

				sSql = "INSERT INTO INV_T_RECEIPTLOTDETAILS (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
					"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
					"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,RATE) VALUES " &_
					"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItmCode & "," & sClassCode& "," &_
					"" & sLoc & "," & sBin & "," & iLotEntry & "," & iSerialEntry & "," &_
					"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & ",NULL," &_
					"" & sPackingType & "," & sSellingType & "," & iWeight & "," & iRate & ")"
				'objTxt.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql

			dcrs.MoveNext
			loop
		end if
		dcrs.Close

		'INV_T_LEDGERLIFO,INV_T_LEDGERFIFO,INV_T_RECEIPTDETAILS,INV_T_LEDGERWA,INV_T_RECEIPTLOTDETAILS
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_M_ITEMORGSTORAGE WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND APPLICABLEFOR = " & Pack(sAppli) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			sSql = "INSERT INTO INV_M_ITEMORGSTORAGE (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
				"APPLICABLEFOR,LOCATIONNUMBER,BINNUMBER,ALLOWTRANSFERS) VALUES " &_
				"(" & iItmCode & "," & sClassCode & "," & Pack(sOrgCode) & "," & Pack(sAppli) & "," &_
				"" & sLoc & "," & sBin & ",'1')"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())
		
		'k
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND MONTHANDYEAR = " & Pack(sMonYr) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,MONTHANDYEAR,RECEIPTQUANTITY,RECEIPTVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & sClassCode & "," & iItmCode & "," &_
				"" & sLoc & "," & sBin & "," & Pack(sMonYr) & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET RECEIPTQUANTITY = (RECEIPTQUANTITY + " & iQty & ")," &_
				"RECEIPTVALUE = (RECEIPTVALUE + " & iValue & ") WHERE ITEMCODE = " & iItmCode & " AND " &_
				"CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
				"MONTHANDYEAR = " & Pack(sMonYr) & ""
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close
		
		'N
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_LOCATIONSTOCKFIFO WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND MONTHANDYEAR = " & Pack(sMonYr) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_LOCATIONSTOCKFIFO (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,MONTHANDYEAR,RECEIPTQUANTITY,RECEIPTVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & sClassCode & "," & iItmCode & "," &_
				"" & sLoc & "," & sBin & "," & Pack(sMonYr) & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_LOCATIONSTOCKFIFO SET RECEIPTQUANTITY = (RECEIPTQUANTITY + " & iQty & ")," &_
				"RECEIPTVALUE = (RECEIPTVALUE + " & iValue & ") WHERE ITEMCODE = " & iItmCode & " AND " &_
				"CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
				"MONTHANDYEAR = " & Pack(sMonYr) & ""
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close
		
		'N
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_LOCATIONSTOCKLIFO WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND MONTHANDYEAR = " & Pack(sMonYr) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_LOCATIONSTOCKLIFO (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,MONTHANDYEAR,RECEIPTQUANTITY,RECEIPTVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & sClassCode & "," & iItmCode & "," &_
				"" & sLoc & "," & sBin & "," & Pack(sMonYr) & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_LOCATIONSTOCKLIFO SET RECEIPTQUANTITY = (RECEIPTQUANTITY + " & iQty & ")," &_
				"RECEIPTVALUE = (RECEIPTVALUE + " & iValue & ") WHERE ITEMCODE = " & iItmCode & " AND " &_
				"CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
				"MONTHANDYEAR = " & Pack(sMonYr) & ""
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close
		
		'N
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_LOCATIONSTOCKWA WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND MONTHANDYEAR = " & Pack(sMonYr) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_LOCATIONSTOCKWA (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,MONTHANDYEAR,RECEIPTQUANTITY,RECEIPTVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & sClassCode & "," & iItmCode & "," &_
				"" & sLoc & "," & sBin & "," & Pack(sMonYr) & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_LOCATIONSTOCKWA SET RECEIPTQUANTITY = (RECEIPTQUANTITY + " & iQty & ")," &_
				"RECEIPTVALUE = (RECEIPTVALUE + " & iValue & ") WHERE ITEMCODE = " & iItmCode & " AND " &_
				"CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
				"MONTHANDYEAR = " & Pack(sMonYr) & ""
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)
		
		'k
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
				"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & sClassCode & "," & iItmCode & "," &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEAROPENINGSTOCK = (YEAROPENINGSTOCK + " & iQty & ")," &_
				"YEAROPENINGVALUE = (YEAROPENINGVALUE + " & iValue & ")," &_
				"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQty & "), " &_
				"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
				"ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
				"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
				"CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close
		
		'N
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
				"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
				"YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & sClassCode & "," & iItmCode & "," &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
				"" & sLoc & "," & sBin & "," &_
				"" & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
			
			'N
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
					"(" & trim(dcrs1(0)) & "," & Pack(sOrgCode) & "," & sClassCode & "," & iItmCode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
					"" & sLoc & "," & sBin & ")"
				'objTxt.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs1.Close

		end if
		dcrs.Close
		
		'N
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_YEARLYSTOCKFIFO WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FIFOFINYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FIFOFINYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_YEARLYSTOCKFIFO (ORGANISATIONCODE,ITEMCODE," &_
				"CLASSIFICATIONCODE,FIFOFINYEARFROM,FIFOFINYEARTO,FIFOYEAROPSTOCK,FIFOYEAROPVALUE,FIFOYEARCLSGSTOCK,FIFOYEARCLSGVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & iItmCode & "," & sClassCode & "," &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_YEARLYSTOCKFIFO SET FIFOYEAROPSTOCK = (FIFOYEAROPSTOCK + " & iQty & ")," &_
				"FIFOYEAROPVALUE = (FIFOYEAROPVALUE + " & iValue & "), " &_
				"FIFOYEARCLSGSTOCK = (FIFOYEARCLSGSTOCK + " & iQty & ")," &_
				"FIFOYEARCLSGVALUE = (FIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
				"ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
				"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FIFOFINYEARFROM,103) AND " &_
				"CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FIFOFINYEARTO,103)"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close
		
		'N
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_YEARLYSTOCKLIFO WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,LIFOFINYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,LIFOFINYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_YEARLYSTOCKLIFO (ORGANISATIONCODE,ITEMCODE," &_
				"CLASSIFICATIONCODE,LIFOFINYEARFROM,LIFOFINYEARTO,LIFOYEAROPSTOCK,LIFOYEAROPVALUE,LIFOYEARCLSGSTOCK,LIFOYEARCLSGVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & iItmCode & "," & sClassCode & "," &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_YEARLYSTOCKLIFO SET LIFOYEAROPSTOCK = (LIFOYEAROPSTOCK + " & iQty & ")," &_
				"LIFOYEAROPVALUE = (LIFOYEAROPVALUE + " & iValue & ")," &_
				"LIFOYEARCLSGSTOCK = (LIFOYEARCLSGSTOCK + " & iQty & ")," &_
				"LIFOYEARCLSGVALUE = (LIFOYEARCLSGVALUE + " & iValue & ") WHERE " &_
				"ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
				"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,LIFOFINYEARFROM,103) AND " &_
				"CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,LIFOFINYEARTO,103)"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close
		
		'N
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_YEARLYSTOCKWA WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,WAFINYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,WAFINYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_YEARLYSTOCKWA (ORGANISATIONCODE,ITEMCODE," &_
				"CLASSIFICATIONCODE,WAFINYEARFROM,WAFINYEARTO,WAYEAROPSTOCK,WAYEAROPVALUE,WAYEARCLSGSTOCK,WAYEARCLSGVALUE) VALUES " &_
				"(" & Pack(sOrgCode) & "," & iItmCode & "," & sClassCode & "," &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_YEARLYSTOCKWA SET WAYEAROPSTOCK = (WAYEAROPSTOCK + " & iQty & ")," &_
				"WAYEAROPVALUE = (WAYEAROPVALUE + " & iValue & ")," &_
				"WAYEARCLSGSTOCK = (WAYEARCLSGSTOCK + " & iQty & ")," &_
				"WAYEARCLSGVALUE = (WAYEARCLSGVALUE + " & iValue & ") WHERE " &_
				"ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
				"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,WAFINYEARFROM,103) AND " &_
				"CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,WAFINYEARTO,103)"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close


	next

	sExp1 ="//root/ASSET"
	Set AssetNode = RootNode.SelectNodes(sExp1)
	for iCtrAsset = 0 to AssetNode.Length - 1
		iBoMClassCode = trim(AssetNode.Item(iCtrAsset).Attributes.getNamedItem("CLACODE").Value)

		sExp2 ="//root/ASSET [@CLACODE = "&iBoMClassCode&"]/ITMDET"
		Set BoMNode = RootNode.SelectNodes(sExp2)
		for iCtrBoM = 0 to BoMNode.Length - 1
			iBoMItemCode = trim(BoMNode.Item(iCtrBoM).Attributes.getNamedItem("ITMCODE").Value)
			iQty = trim(BoMNode.Item(iCtrBoM).Attributes.getNamedItem("QTY").Value)
			sUoM = trim(BoMNode.Item(iCtrBoM).Attributes.getNamedItem("SUOM").Value)
			sType = trim(BoMNode.Item(iCtrBoM).Attributes.getNamedItem("ITYPE").Value)
			
			'N
			sSql = "INSERT INTO INV_M_ITEMASSETBOM (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
				"BOMCLASSCODE,BOMITEMCODE,QUANTITY,TYPEOFITEM,UOMCODE) VALUES " &_
				"(" & iItmCode & "," & sClassCode & "," & Pack(sOrgCode) & "," &_
				"" & iBoMClassCode & "," & iBoMItemCode & "," & iQty & "," & Pack(sType) & "," & Pack(sUoM) & ")"
			'objTxt.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql

		next
	next

	sSql = "UPDATE FAR_T_ASSETDETAILS SET SENTTOINVENTORY = 'S' WHERE ASSETDESCID = " & trim(sAssetCode) & ""
	'objTxt.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql


	'Declaration of Objects
	Set adoCmd = Server.CreateObject("ADODB.Command")
	Set adoCmd.ActiveConnection = con

	sSql = "mapItemOrgMaster"

	adoCmd.CommandText = sSql
	adoCmd.CommandType = 4

	adoCmd.Parameters.Refresh
	adoCmd.Parameters.Item(1).Value = sOrgCode
	adoCmd.Parameters.Item(2).Value = sClassCode
	adoCmd.Parameters.Item(3).Value = iItmCode

	Set dcrs = adoCmd.Execute()

	if con.Errors.count <> 0 then
		dim iErrCounter
		con.RollbackTrans
		for iErrCounter=0 to con.Errors.count
			Response.Write con.Errors(iErrCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else

		'con.RollbackTrans
		'Response.End 
		con.CommitTrans
	end if

	con.close
	set con = nothing
%>
