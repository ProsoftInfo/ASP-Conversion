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
	'Program Name				:	ItmEditEntry.asp
	'Module Name				:	Inventory (Item Creation and Definition)
	'Author Name				:	S.MAHESHWARI
	'Created On					:	August 31, 2007
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/UoMDecimal.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/CommonFunctions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>
iTMS - Item Creation and Definition</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">

<script type="application/xml" id="Data" data-itms-xml-island>
<root/>
</script>

<script type="application/xml" id="NewRt" data-itms-xml-island>
<root/>
</script>
<script type="application/xml" id="CategoryData" data-itms-xml-island>
<Root></Root>
</script>
<%
	Dim dcrs,dcrs1,dcrs2,oDOM,oDOM1,oDOM2,Root,Root1, Node,Node1,Node2,Node3,Node4,sItemTypeId,iClass
	Dim sIType, sITypeCode, iItemCode, sCode, sDesc, sShortDesc, sAddDesc,newElem,newElem1
	Dim iDrawingNo, iCatalogNo, iMGRNo, iPageNo, iPosNo, iCommodity, iController, Flag,sCompItemCode
	Dim UomFlag, sStoreUom, sPurUom, sManUom, sSalUom, iPurToStoreRate, iPurToStoreOperator
	Dim iManToStoreRate, iManToStoreOperator, iSaleToStoreRate, iSaleToStoreOperator
	Dim sRecptNo,sRecptRout,sAccountType,sModvat,sBoM,iReLevel,iReQty,iEcoQty,sUnitID,sCatCode,sClassName
	Dim sClassCode,sLoc,sBin,sAllowTrans,sLocName,sBinName,iNoOfBins
	Dim Lrec,LSRec,SRec,NRec,DU,Stk,InsD,InsS,InsSD,Lf,Ff,WA,iQty,iValue,sFor,sOpt
	Dim iLotNo,iSerNo,iLQtyGrs,iLQtyNet,iStorLocNo,iStorBinNo,iPackNo,iPackCode,iSellNo,iWgtPerSelForm
	Dim iSellForm,iRate,iQtyIssued,iStage,iAttribList,iCtr,dDate,Arr,dMonYear
	Dim iAttrib,sArr,sTemp,i,StoreRt,iAltGrs,iAltNet,iAltUOM,iReOrdLev,iReOrdQty,iEcoOrdQty
	Dim iBOMApplicable,BOMNode,sItemName,sQuery,sCategoryName,sCategoryCode
	Dim oDomReplicate,RootReplicate,ReplicateElem
	Dim sItemActive,sItemHold,sDeadStock,sStatus,sModVatEligible,sPurTaxType,sSalTaxType,sBlowUpImage,sSql,ndAccHead
	Dim iLotDetValue,iLotValue
	Dim bTareEligible,bSubLevel
	

	Dim arrFin,sFinFrom,sFinTo,sTempMonYr,sMonYr,sCheckFinYear,sFinPeriod,ndAtt,ndAttDet,sAttList,sLot


	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")

	set oDOM = server.CreateObject("Microsoft.XMLDOM")
	set oDOM1 = server.CreateObject("Microsoft.XMLDOM")
	set oDOM2 = server.CreateObject("Microsoft.XMLDOM")
	set oDomReplicate = Server.CreateObject("Microsoft.XMLDOM")
	Response.Write "<font color=red>"



	set Root = oDOM.createElement("ROOT")
	oDOM.appendChild Root

	set StoreRt = oDOM2.createElement("ROOT")
	oDOM2.appendChild StoreRt
	
	set RootReplicate = oDomReplicate.createElement("Output")
	oDomReplicate.appendChild RootReplicate

	sTempMonYr = mid(FormatDate(date),4,2)

	sMonYr = sTempMonYr&Year(date())

	'arrFin = split(GetFinancialYear(sMonYr),":")
	sFinPeriod = Trim(Session("FinPeriod"))
	arrFin = split(sFinPeriod,":")

	sFinFrom = "01/04/"&arrFin(0)
	sFinTo = "31/03/"&arrFin(1)


	'sUnitID = trim(Request.Form("hUnitID"))
	sUnitID = Session("organizationcode")
'	sIType = trim(Request.Form("hItemTypeName"))
'	sITypeCode = trim(Request.Form("hItemTypeCode"))
	iItemCode = trim(Request("hItemCode"))
	sClassCode = trim(Request("ClassCode"))
	Flag = False
	UomFlag = False
	'Response.Write sUnitID &"***"&sITypeCode &"***"& iItemCode&"***"&sIType
	'010101***GAR***18***Garment

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMTYPEID FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode & ""
		.Source = "SELECT DISTINCT COMPANYITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode & ""
	'	Response.Write dcrs.Source
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sCode = trim(dcrs(0))
		'sITypeCode = trim(dcrs(1))
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = " SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION, ISNULL(SHORTDESCRIPTION,''), ISNULL(ADDITIONALDESCRIPTION,''), ISNULL(DRAWINGNUMBER,''), "&_
				  " ISNULL(CATALOGUENO,''), ISNULL(ITEMCONTROLLER,0), ISNULL(STORESUOM,''), ISNULL(PURCHASEUOM,''), ISNULL(MANUFACTURINGUOM,''), ISNULL(SALESUOM,''),"&_
				  " ISNULL(PURTOSTORERATE,0), ISNULL(PURTOSTOREOPERATOR,0), ISNULL(MANTOSTORERATE,0), ISNULL(MANTOSTOREOPERATOR,0),ISNULL(SALETOSTORERATE,0), "&_
				  " ISNULL(SALETOSTOREOPERATOR,0),RECEIPTNUMBERING,RECEIPTROUTING,ACCOUNTINGTYPE,CATEGORYCODE,ISNULL(ATTRIBUTELIST,0), "&_
				  " ISNULL(REORDERLEVEL,0),ISNULL(REORDERQTY,0),ISNULL(ECOORDERQTY,0),ISNULL(BOMAPPLICABILITY,0),IsNull(AllowModvatCredit,'0'),IsNull(PurTaxType,0),IsNull(SalTaxType,0) "&_
				  " FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode & " AND ORGANISATIONCODE = " & Pack(sUnitID) & ""
		'Response.Write dcrs.source
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sCompItemCode =  trim(dcrs(0))
		sDesc = Replace(trim(dcrs(1)),chr(39),"~~")
		sDesc = Replace(trim(dcrs(1)),Chr(34),"``")

		sShortDesc = trim(dcrs(2))
		sAddDesc = trim(dcrs(3))
		iDrawingNo = trim(dcrs(4))
		iCatalogNo = trim(dcrs(5))
		iController = trim(dcrs(6))
		sStoreUom = trim(dcrs(7))
		sPurUom = trim(dcrs(8))
		sManUom = trim(dcrs(9))
		sSalUom = trim(dcrs(10))
		iPurToStoreRate = trim(dcrs(11))
		iPurToStoreOperator = trim(dcrs(12))
		iManToStoreRate = trim(dcrs(13))
		iManToStoreOperator = trim(dcrs(14))
		iSaleToStoreRate = trim(dcrs(15))
		iSaleToStoreOperator  = trim(dcrs(16))
		sRecptNo       = trim(dcrs(17))
		sRecptRout     = trim(dcrs(18))
		sAccountType   = trim(dcrs(19))
		sCatCode	   =  trim(dcrs(20))
		iAttrib	       =  trim(dcrs(21))
		iReOrdLev      =  trim(dcrs(22))
		iReOrdQty      =  trim(dcrs(23))
		iEcoOrdQty	   =  trim(dcrs(24))
		iBOMApplicable =  trim(dcrs(25))
		sModVatEligible= trim(dcrs(26))
		sPurTaxType    = trim(dcrs(27)) 
		sSalTaxType    = trim(dcrs(28))

		' Response.Write "sPurUom"&iAttrib

		Set newElem = oDOM.createElement("UOMDETAILS")
		newElem.setAttribute "PUR",sPurUom
		newElem.setAttribute "MAN", sManUom
		newElem.setAttribute "SAL", sSalUom
		newElem.setAttribute "PURFAC",iPurToStoreRate
		newElem.setAttribute "PUROPE",iPurToStoreOperator
		newElem.setAttribute "SALFAC",iSaleToStoreRate
		newElem.setAttribute "SALOPE",iSaleToStoreOperator
		newElem.setAttribute "MANFAC",iManToStoreRate
		newElem.setAttribute "MANOPE",iManToStoreOperator
		Root.appendChild newElem
				
		Set ReplicateElem  = oDomReplicate.createElement("UOMDETAILS")
		ReplicateElem.setAttribute "PUR",sPurUom
		ReplicateElem.setAttribute "MAN", sManUom
		ReplicateElem.setAttribute "SAL", sSalUom
		ReplicateElem.setAttribute "PURFAC",iPurToStoreRate
		ReplicateElem.setAttribute "PUROPE",iPurToStoreOperator
		ReplicateElem.setAttribute "SALFAC",iSaleToStoreRate
		ReplicateElem.setAttribute "SALOPE",iSaleToStoreOperator
		ReplicateElem.setAttribute "MANFAC",iManToStoreRate
		ReplicateElem.setAttribute "MANOPE",iManToStoreOperator
		RootReplicate.appendChild ReplicateElem
		

		With dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = "Select UOMCode,OptionToBaseRate,OptionToBaseOperator,OptionalUoMFor from Inv_M_ItemOptionalUoM "&_
					  " where ItemCode ="& iItemCode & " And OrganisationCode ="& Pack(sUnitID)&" "
			'Response.Write dcrs1.Source
			.Open
		End With

		Do while not dcrs1.EOF
		IF trim(dcrs1(3)) = "P" then sFor = "Purchase" else sFor = "Sales"
		IF trim(dcrs1(2)) = "0" then sOpt = "*" else sOpt = "/"

			Set newElem1 = oDOM.createElement("OPUOMENTRY")
			newElem1.setAttribute "UCODE",dcrs1(0)
			newElem1.setAttribute "BRATE",dcrs1(1)
			newElem1.setAttribute "OPERATOR",dcrs1(2)
			newElem1.setAttribute "UNAME",dcrs1(0)
			newElem1.setAttribute "OPERATORTEXT",sOpt
			newElem1.setAttribute "FOR",sFor
			newElem.appendChild newElem1
			dcrs1.MoveNext
		loop
		dcrs1.Close
		
		sQuery= "Select I.ItemTypeAttributeID,I.HeaderID,ItemTypeAttributeName,ItemTypeAttributeType,ItemTypeAttributeDataLength,ItemTypeAttributeDecimal,AttributeValue from INV_M_ItemMasterAttributes I,INV_M_ItemTypeAttributes A where I.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemCode = " & iItemCode
		dcrs1.open sQuery,con
		if not dcrs1.eof then
		    set ndAtt = oDOM.createElement("ATTRIBUTE")
		    root.appendChild ndAtt
		    do while not dcrs1.eof 
		        set ndAttDet = oDOM.createElement("ATTDET")
		            ndAttDet.setAttribute "AttID",dcrs1(0)
		            ndAttDet.setAttribute "Head",dcrs1(1)
		            ndAttDet.setAttribute "AttName",dcrs1(2)
		            ndAttDet.setAttribute "Type",dcrs1(3)
		            ndAttDet.setAttribute "Length",dcrs1(4)
		            ndAttDet.setAttribute "Decimal",dcrs1(5)
		            ndAttDet.setAttribute "Value",dcrs1(6)
		            ndAtt.appendChild ndAttDet
		        dcrs1.movenext
		    loop
		end if
		dcrs1.close
		
		
		Set Node = oDOM.createElement("DETAILS")
		Node.setAttribute "ITYPE",sITypeCode
		Node.setAttribute "ICODE",iItemCode
		Node.setAttribute "COMPITEMCODE",sCode
		Node.setAttribute "DESC",sDesc
		Node.setAttribute "SHDESC",sShortDesc
		Node.setAttribute "CATALOUGE",iCatalogNo
		Node.setAttribute "DRAWVER",iDrawingNo
		Node.setAttribute "VARIANT",""
		Node.setAttribute "ADDDESC",sAddDesc
		Node.setAttribute "UOM",sStoreUom
		Node.setAttribute "CATEGORY",sCatCode
		Node.setAttribute "ATTRIBUTES",iAttrib
		Node.setAttribute "GROUP",""
		Node.setAttribute "LEVEL",""
		Node.setAttribute "UNIT",sUnitID
		Node.setAttribute "MODVAT",sModVatEligible
		Node.setAttribute "PURTAX",sPurTaxType
		Node.setAttribute "SALTAX",sSalTaxType
		Root.appendChild Node

		Set Node =  oDOM.createElement("CONTROLS")
		Node.setAttribute "RECNUM",sRecptNo
		Node.setAttribute "ROUTING",sRecptRout
		Node.setAttribute "ACCOUNTING",sAccountType
		Node.setAttribute "MODVAT","0"
		Node.setAttribute "REORDERLEVEL",iReOrdLev
		Node.setAttribute "REORDERQTY",iReOrdQty
		Node.setAttribute "ECOORDERQTY",iEcoOrdQty
		Node.setAttribute "BOMAPPLICABLE",iBOMApplicable
		Root.appendChild Node
		
		sSql = "Select I.AccountHeadFor,I.AccountHead,IsNull(A.AccountDescription,'') from Inv_M_ItemOrgAccountHead I left outer join Acc_M_GLAccountHead A On I.AccountHead = A.AccountHead where ItemCode = "& iItemCode &"  and Organisationcode = "& Pack(sUnitID)
		dcrs1.open sSql,con
		if not dcrs1.eof then
		    do while not dcrs1.eof 
		        Set ndAccHead = oDOM.createElement("ACCHEAD")
		        ndAccHead.setAttribute "Name",dcrs1(2)
		        ndAccHead.setAttribute "Type",dcrs1(0)
		        ndAccHead.setAttribute "Value",dcrs1(1)
		        Node.appendChild ndAccHead
		        dcrs1.movenext
		    loop
		end if
		dcrs1.close
		

		if trim(sPurUom) = "-" then
			iPurToStoreRate = ""
			iPurToStoreOperator = -1
		end if
		if trim(sManUom) = "-" then
			iManToStoreRate = ""
			iManToStoreOperator = -1
		end if
		if trim(sSalUom) = "-" then
			iSaleToStoreRate = ""
			iSaleToStoreOperator = -1
		end if
		If trim(sRecptNo) = "L" then LRec = "Selected" else LRec =""
		If trim(sRecptNo) = "S" then SRec = "Selected" else SRec =""
		If trim(sRecptNo) = "LS" then LSRec = "Selected" else LSRec =""
		If trim(sRecptNo) = "N" then NRec = "Selected" else NRec =""

		IF trim(sRecptRout) = "R" then DU = "Selected" else DU = ""
		IF trim(sRecptRout) = "S" then Stk = "Selected" else Stk = ""
		IF trim(sRecptRout) = "IR" then InsD = "Selected" else InsD = ""
		IF trim(sRecptRout) = "IS" then InsS = "Selected" else InsS = ""
		IF trim(sRecptRout) = "ISR" then InsSD = "Selected" else InsSD = ""

		IF trim(sAccountType) = "L" then Lf = "Selected" else Lf = ""
		IF trim(sAccountType) = "F" then Ff = "Selected" else Ff = ""
		IF trim(sAccountType) = "W" then WA = "Selected" else WA = ""

	end if
	dcrs.Close
	'Storage Stock
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
	'	.Source = " SELECT ISNULL(I.CLASSIFICATIONCODE,0),ISNULL(I.LOCATIONNUMBER,0),ISNULL(I.BINNUMBER,0),I.ALLOWTRANSFERS,S.LOCATIONNAME,B.BINNAME FROM INV_M_ITEMSTORAGE AS I,"&_
	'			  " INV_M_STORAGE AS S,INV_M_STOREBINDETAILS AS B WHERE I.ITEMCODE = " & iItemCode & " AND I.ORGANISATIONCODE = " & Pack(sUnitID) & " AND S.LOCATIONNUMBER = I.LOCATIONNUMBER "&_
	'			  " AND S.OUDEFINITIONID = I.ORGANISATIONCODE AND ISNULL(I.BINNUMBER,0) = ISNULL(B.BINNUMBER,0)"
		.Source = " SELECT ISNULL(CLASSIFICATIONCODE,0),ISNULL(LOCATIONNUMBER,0),ISNULL(BINNUMBER,0),ALLOWTRANSFERS FROM INV_M_ITEMSTORAGE WHERE ITEMCODE = " & iItemCode & " AND "&_
				  " ORGANISATIONCODE = " & Pack(sUnitID) & " "
		.ActiveConnection = con
		.Open
	end with
	'Response.Write dcrs.source

	Do while Not dcrs.EOF
		sClassCode  = dcrs(0)
		sLoc	    = dcrs(1)
		sBin	    = dcrs(2)
		sAllowTrans = dcrs(3)

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = "Select GroupName,isNull(GroupCategory,'') from INV_M_Classification where GroupCode ='"& sClassCode &"'"' and ItemTypeID = '"& sITypeCode &"'"
			'Response.Write dcrs1.Source
			.Open
		end with
		if not dcrs1.EOF then
			sClassName = dcrs1(0)
			sCategoryCode = sCatCode 
			
			Set Node = oDOM2.createElement("CLASSIFICATION")
			Node.setAttribute "CODE",sClassCode
			Node.setAttribute "NAME",dcrs1(0)
			Node.setAttribute "CATEGORY",sCatCode 
			Root.appendChild Node
			
			Set ReplicateElem = oDomReplicate.createElement("CLASSIFICATION")
			ReplicateElem.setAttribute "CODE",sClassCode
			ReplicateElem.setAttribute "NAME",dcrs1(0)
			ReplicateElem.setAttribute "CATEGORY",sCatCode
			RootReplicate.appendChild ReplicateElem
			
			
		end if
		dcrs1.Close

		with dcrs1
		    .CursorLocation = 3
		    .CursorType = 3
		    .ActiveConnection = con
		    .Source = "Select ItemCode,ClassificationCode,OrganisationCode,BOMClassificationCode,BOMItemCode,Quantity,UOM,Type,Consumable from INV_M_ItemMasterBOM where ItemCode = "& iItemCode &" and ClassificationCode = "& sClassCode
		    'Response.Write dcrs1.Source
		    .Open
		end with
		if not dcrs1.EOF then
		    Set BOMNode = oDOM2.createElement("BOM")
		    Root.appendChild BOMNode
		    do while not dcrs1.eof
		        sQuery = "Select ItemDescription from VWITem where ItemCode = "&dcrs1(4) &" and ClassificationCode = "& dcrs1(3)
		        dcrs2.Open sQuery,con
		        if not dcrs2.EOF then
		            sItemName = trim(dcrs2(0))
		        end if
		        dcrs2.Close
		            Set Node = oDOM2.createElement("Item")
				    Node.setAttribute "ItemCode",dcrs1(4)
				    Node.setAttribute "ClassCode",dcrs1(3)
				    Node.setAttribute "ItemName",sItemName
				    Node.setAttribute "Qty",dcrs1(5)
				    Node.setAttribute "UoM",dcrs1(6)
				    Node.setAttribute "Type",dcrs1(7)

				    if triM(dcrs1(7))="F" then
				        Node.setAttribute "TypeName","Final Component"
				    else
				        Node.setAttribute "TypeName","Assembly"
				    end if

				    Node.setAttribute "Consumable",dcrs1(8)
				    BOMNode.appendChild Node
		        dcrs1.MoveNext
		    loop
		end if
		dcrs1.Close

		With dcrs1
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = " Select LocationName,NumberOfBins from INV_M_STORAGE where OUDefinitionID = "& Pack(sUnitID)&" "&_
					  " And LocationNumber ="& sLoc &" "
		'Response.Write dcrs1.Source
 			.Open
		End With
		IF Not dcrs1.EOF then
			sLocName = dcrs1(0)
			iNoOfBins = dcrs1(1)
		End IF
		dcrs1.Close

		IF trim(iNoOfBins) <> "0" then
			With dcrs2
				.ActiveConnection = con
				.CursorLocation = 3
				.CursorType = 3
				.Source = " Select BinName from INV_M_STOREBINDETAILS where OUDefinitionID = "& Pack(sUnitID)&" "&_
						  "And LocationNumber ="& sLoc &" "
 				.Open
			End With
			IF Not dcrs2.EOF then
				sBinName = dcrs2(0)
			Else
				sBinName =	""
			End IF
			dcrs2.Close
		Else
			sBinName =	""
		End IF

	With dcrs2
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = " Select Distinct S.YearOpeningStock,S.YearOpeningValue,isNull(Convert(Varchar,L.DateOfReceipt,103),''),isNull(L.Stage,''), "&_
					  " isNull(L.AltGross,0),isNull(L.AltNett,0) from Inv_T_ItemLocationStock as S,"&_
					  "	INV_T_LocationLOT as L where S.OrganisationCode = "& Pack(sUnitID)&" And S.ItemCode = "& iItemCode &" "&_
					  " And S.ClassificationCode = "& sClassCode &" And  S.LocationNumber ="& sLoc &" And  S.ItemCode = L.ItemCode  "&_
					  " and L.StorageLocationNo=S.LocationNumber and L.SrcType='RO'"&_
					  " and CONVERT(DATETIME,L.DateOfReceipt,103) BETWEEN CONVERT(DATETIME,"& Pack(sFinFrom) &",103) AND "&_
					  " CONVERT(DATETIME,"& Pack(sFinTo) &",103) "
		'Response.Write dcrs2.Source
			.Open
		End With
		'Response.Write dcrs2.Source
		IF dcrs2.EOF then
			iQty		= 0
			iValue		= 0
		End IF

		Do while Not dcrs2.EOF
			iQty		= dcrs2(0)
			iValue		= dcrs2(1)
			dDate		= dcrs2(2)
			iStage		= dcrs2(3)
			iAltGrs		= dcrs2(4)
			iAltNet		= dcrs2(5)
			'iAltUOM
			IF trim(dDate) <> "" then
				Arr = Split(dDate,"/")
				dMonYear =  Arr(1) & Arr(2)
				'Response.Write "date="&dMonYear
			Else
				dMonYear =  ""
			End IF
			'Response.Write iQty
			'Response.Write sITypeCode
			Set Node =  oDOM2.createElement("STOREDET")
			Node.setAttribute "ITEM",iItemCode
			Node.setAttribute "CLASS",sClassCode
			Node.setAttribute "UNITSTORE",sUnitID &"-"& sLoc &"-"& sBin &"-"& iQty &"-"& iValue &"-"& dMonYear
			If Not sBinName = "" then
				Node.setAttribute "STORE",sLocName &"--"& sBinName
			Else
				Node.setAttribute "STORE",sLocName
			End IF
			Node.setAttribute "BIN",sBin
			Node.setAttribute "ALLTRANS",sAllowTrans
			StoreRt.appendChild Node

			Set Node1 =  oDOM2.createElement("STORAGE")
			'Response.Write "sLoc = "& sLoc
			Node1.setAttribute "STORE",sLoc
			Node1.setAttribute "BIN", trim(sBin)
			Node1.setAttribute "MONTHYEAR",dMonYear
			Node1.setAttribute "QTY", iQty
			Node1.setAttribute "STORAGEVALUE", iValue
			Node1.setAttribute "CLASSIFICATION", sClassCode
			Node1.setAttribute "UNIT", sUnitID

			'IF sITypeCode = "YRN" then
			'	Node1.setAttribute "STAGE",iStage
			'	Node1.setAttribute "ALTGROSS",iAltGrs
			'	Node1.setAttribute "ALTNETT",iAltNet
				'Node1.setAttribute "ALTUOM",iAltUOM
			'End IF
			Node.appendChild Node1

			With dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = " Select IsNull(LotNumber,0),IsNull(AttributeList,0),SUM(LotQuantityNett),PackingCode from Inv_T_LocationLot where ItemCode ="& iItemCode & " And ClassificationCode = "& sClassCode &" And OrganisationCode ="& Pack(sUnitID)&" And SrcType = 'RO' and LotQuantityNett > 0 Group By LotNumber,AttributeList,PackingCode "
				.Open
			End With
			iCtr = 1
			if not dcrs1.eof then
			Do while not dcrs1.EOF
			    sAttList = dcrs1(1)
			    sLot = dcrs1(0)
			    if trim(sAttList)="0" then sAttList=""
			    if trim(sLot)="0" then sLot=""
			    if trim(dcrs1(3))<>"" then
			        sQuery = "Select IsNull(TarePerPack,'N') from APP_M_PackingType where PackingCode = "& dcrs1(3)
			        rsTemp.open sQuery,con
			        if not rsTemp.eof then
			            bTareEligible = rsTemp(0)
			        end if
			        rsTemp.close
    			    
			        sQuery = "Select * from APP_M_PackingTypeSubLevel where PackingCode = "& dcrs1(3)
			        rsTemp.open sQuery,con
			        if not rsTemp.eof then
			            bSubLevel = "Y"
			        else
			            bSubLevel = "N"
			        end if
			        rsTemp.close
			    end if
			    
			    'iLotValue = Round(cdbl(dcrs1(4))*cdbl(dcrs1(2)))
			    
			    Set Node2 =  oDOM2.createElement("LotSerial")
					Node2.setAttribute "QTYIN","N"
					Node2.setAttribute "TARE","0"
					Node2.setAttribute "LOT",sLot
					Node2.setAttribute "SERIALFROM",""
					Node2.setAttribute "SERIALTO",""
					Node2.setAttribute "TAREWEIGHT",""
					Node2.setAttribute "QTY",dcrs1(2)
					Node2.setAttribute "COUNTER",iCtr
					Node2.setAttribute "STAGE",""
					Node2.setAttribute "ALTGROSS",""
					Node2.setAttribute "ALTNETT",""
					Node2.setAttribute "ALTUOM",""
					Node2.setAttribute "IVALUE",iValue
					Node2.setAttribute "AUTOGEN",""
					Node2.setAttribute "ATTLIST",sAttList
					Node2.setAttribute "TAREELIGIBLE",bTareEligible
					Node2.setAttribute "SUBLEVEL",bSubLevel
					Node1.Appendchild Node2
					
				sQuery = " Select ClassificationCode,IsNull(LotNumber,0),IsNull(SerialNumber,0),LotQuantityGross,LotQuantityNett,StorageLocationNo, "&_
						 " IsNull(StorageBinNumber,0),IsNull(PackingNumber,0),IsNull(PackingCode,0),IsNull(SellingNumber,0),IsNull(WeightPerSellingForm,0), "&_
						 " IsNull(SellingForm,0),Rate,QuantityIssued,IsNull(Stage,0),IsNull(AttributeList,0),isNull(Convert(Varchar,DateOfReceipt,103),'') "&_
						 " from Inv_T_LocationLot where ItemCode ="& iItemCode & " And ClassificationCode = "& sClassCode &" And OrganisationCode ="& Pack(sUnitID)&" And SrcType = 'RO' and LotQuantityNett > 0 "
						 if trim(dcrs1(0))<>"" and trim(dcrs1(0))<>"0" then
						    sQuery = sQuery &" and LotNumber = '"& dcrs1(0) &"'"
						 end if
						 
						 if trim(dcrs1(1))<>"" and trim(dcrs1(1))<>"0" then
						    sQuery = sQuery &" and AttributeList = '"& dcrs1(1) &"'"
						 end if
						 'Response.Write sQuery
				rsTemp.open sQuery,con
				if not rsTemp.eof then
				    do while not rsTemp.eof
				        iClass			= rsTemp(0)
				        iLotNo			= rsTemp(1)
				        iSerNo			= rsTemp(2)
				        iLQtyGrs		= rsTemp(3)
				        iLQtyNet		= rsTemp(4)
				        iStorLocNo		= rsTemp(5)
				        iStorBinNo		= rsTemp(6)
				        iPackNo			= rsTemp(7)
				        iPackCode		= rsTemp(8)
				        iSellNo			= rsTemp(9)
				        iWgtPerSelForm	= rsTemp(10)
				        iSellForm		= rsTemp(11)
				        iRate			= rsTemp(12)
				        iQtyIssued		= rsTemp(13)
				        iStage			= rsTemp(14)
				        iAttribList		= rsTemp(15)
				        dDate			= rsTemp(16)
				        
				        'Response.Write   iRate &"<BR>"
				        IF iSerNo = "0" then iSerNo = "1"
				        iLotDetValue =round(cdbl(iRate) * cdbl(iLQtyGrs))
				        
				        if trim(iAttribList)="0" then iAttribList=""
				        if trim(iLotNo)="0" then iLotNo=""

				        Set Node3 =  oDOM2.createElement("LotSerialDetails")
				        Node3.setAttribute "LOTSERIAL",iSerNo
				        Node3.setAttribute "QTYREC", iLQtyNet
				        Node3.setAttribute "TAREREC","0"
				        Node3.setAttribute "SELLINGTYPE",""
				        Node3.setAttribute "WEIGHTSTYPE",iWgtPerSelForm
				        Node3.setAttribute "PACKINGTYPE",iPackCode
				        Node3.setAttribute "LOT",iLotNo
				        Node3.setAttribute "SELLINGFORM",iSellForm
				        Node3.setAttribute "PACKNUMBER",iPackNo
				        Node3.setAttribute "IVALUE",iLotDetValue
				        Node3.setAttribute "ATTRIBUTELIST",iAttribList
				        Node3.setAttribute "NOOFCONE",""
				        Node3.setAttribute "SUBLEVELID",""
				        Node3.setAttribute "AUTOGENDET","N"
				        Node2.Appendchild Node3
				        rsTemp.movenext
				    loop
				end if
				rsTemp.close
				iCtr = iCtr + 1
				dcrs1.MoveNext
			loop
			end if'if not dcrs1.eof then
			dcrs1.Close

		dcrs2.MoveNext
	loop
	dcrs2.Close
	dcrs.MoveNext
loop
dcrs.Close

if trim(sClassCode)="" then
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT TOP 1 ClassificationCode FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode & " AND ORGANISATIONCODE = " & Pack(sUnitID)
		.ActiveConnection = con
		.Open
	end with
	if not dcrs.EOF then sClassCode = trim(dcrs(0))
	dcrs.Close
end if
if trim(sClassCode)="" then sClassCode = "0"


sQuery= "Select ITEMACTIVE,ITEMONHOLD,isNull(DeadStock,'N') from INV_M_ITEMMASTER where ItemCode="& iItemCode &" and ClassificationCode="& sClassCode &" and OrganisationCode = "& Pack(sUnitID)
'Response.Write sQuery
dcrs.Open sQuery,con
if not dcrs.EOF then
    sItemActive = dcrs(0)
    sItemHold = dcrs(1)
    sDeadStock = dcrs(2)
end if
dcrs.Close

if sDeadStock = "Y" then
    sStatus = "DS"
else
    if trim(sItemActive)="Y" and trim(sItemHold) = "0" then
        sStatus = "AC"
    elseif trim(sItemActive)="Y" and trim(sItemHold) = "1" then
        sStatus = "OH"
    elseif trim(sItemActive)="N" and trim(sItemHold) = "1" then
        sStatus = "OH"
    elseif trim(sItemActive)="N" and trim(sItemHold) = "0" then
        sStatus = "IA"
    end if
end if


if trim(sStatus)="" then sStatus = "AC"
if trim(sCategoryCode)="" then sCategoryCode = sCatCode

sQuery = "Select CategoryCode,CategoryName from Inv_M_ClassificationCategory where CategoryCode in('"& sCategoryCode &"')"
'Response.Write sQuery
dcrs.Open sQuery,con
if not dcrs.EOF then
    sCategoryName = Trim(dcrs(1))
end if
dcrs.Close 

	'Checking whether the Item Exists in ItemOrgMaster
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode
		.ActiveConnection = con
		.Open
	end with
'	Response.Write dcrs.Source
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then Flag = True
	dcrs.close
	Flag = false
	''''''''''''''''''''''''
	'Inv_M_ItemOptionalUoM '
	''''''''''''''''''''''''
	'Checking whether the Item Exists in ItemYearlySTock
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE YEARRECEIPTQUANTITY > 0 AND YEARISSUEQUANTITY > 0 AND ITEMCODE = " & iItemCode
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then UomFlag = True
	dcrs.close
oDOM2.save server.MapPath("../Temp/Master/StoreDetails"&Session.SessionID&".Xml")
oDOM.save server.MapPath("../Temp/Master/ItemEditStock"&Session.SessionID&".Xml")

oDomReplicate.save server.MapPath("../Temp/Master/ItemReplicate"&Session.SessionID&".Xml")
%>
<%
Dim  ARoot

	set ARoot = oDOM1.createElement("ROOT")
	oDOM1.appendChild ARoot

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT ITEMTYPEID, ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES ORDER BY 2"
		.Source = "SELECT ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES ORDER BY 2"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	do while Not dcrs.EOF
		set Node = oDOM1.createElement("ATTRIBUTES")
		'Node.setAttribute "ITEMTYPEID",trim(dcrs(0))
		Node.setAttribute "ATTRID",trim(dcrs(0))
		Node.setAttribute "ATTRNAME",trim(dcrs(1))

		ARoot.appendChild Node
		dcrs.movenext
	loop
	dcrs.Close

	oDOM1.save server.MapPath("../Temp/Master/Attribute"&Session.SessionID&".Xml")
%>
<script type="application/xml" id="AttrData" data-itms-xml-island data-src="<%="../Temp/Master/Attribute"&Session.SessionID&".Xml"%>"></script>
<script type="application/xml" id="OutData" data-itms-xml-island data-src="<%="../Temp/Master/ItemEditStock"&Session.SessionID&".Xml"%>">
<Output/>
</script>
<script type="application/xml" id="StoreData" data-itms-xml-island data-src="<%="../Temp/Master/StoreDetails"&Session.SessionID&".Xml"%>"></script>

<script type="application/xml" id="ReplicateData" data-itms-xml-island data-src="<%="../Temp/Master/ItemReplicate"&Session.SessionID&".Xml"%>"></script>

<script type="application/xml" id="ItemAttData" data-itms-xml-island>
    <Root>
        <%
            sQuery = "Select M.ItemTypeAttributeID,M.HeaderID,ItemTypeAttributeName,ItemTypeAttributeType,ItemTypeAttributeDataLength,ItemTypeAttributeDecimal,AttributeValue from INV_M_ItemMasterAttributes M,INV_M_ItemTypeAttributes T where M.ItemTypeAttributeID = T.ItemTypeAttributeID and M.ItemCode = "& iItemCode
            dcrs.open sQuery,con
            if not dcrs.eof then
                do while not dcrs.eof 
                    %>
                        <Attribute AttID="<%=dcrs(0)%>" Head="<%=dcrs(1)%>" AttName="<%=dcrs(2)%>" Type="<%=dcrs(3)%>" Length="<%=dcrs(4)%>" Decimal="<%=dcrs(5)%>" Value="<%=dcrs(6)%>"></Attribute>
                    <%
                    dcrs.movenext
                loop
            end if
            dcrs.close
        %>
    </Root>
</script>
<script type="application/xml" id="HeadData" data-itms-xml-island>
    <Root>
        <%
            sQuery = "Select H.HeaderID,ItemTypeHeaderName from Inv_M_ItemTypeHeader H,Inv_M_ItemTypeAttributes A where H.HeaderID = A.HeaderID Group By H.HeaderID,ItemTypeHeaderName Order by H.headerId "
            dcrs.open sQuery,con
            if not dcrs.eof then
                do while not dcrs.eof 
                        %>
                            <Header ID="<%=dcrs(0)%>" Name="<%=dcrs(1)%>"></Header>
                        <%
                    dcrs.movenext
                loop
            end if
            dcrs.close
        %>
    </Root>
</script>
<script type="application/xml" id="OptionData" data-itms-xml-island>
    <Root>
        <%
            sQuery ="Select A.ItemTypeAttributeID,ItemTypeAttributeName,OptionValue,OptionName from Inv_M_ItemTypeAttributes A,INV_M_ItemTypeOptions O where A.ItemTypeAttributeID= O.ItemTypeAttributeID "
            dcrs.open sQuery,con
            if not dcrs.eof then
                do while not dcrs.eof 
                    %>
                        <Option AttID="<%=dcrs(0)%>" AttName="<%=dcrs(1)%>" ID="<%=dcrs(2)%>" Name="<%=dcrs(3)%>"></Option>
                    <%
                    dcrs.movenext
                loop
            end if
            dcrs.close
        %>
    </Root>
</script>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmCreate.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemCreationDefinition.js"></SCRIPT>

</head>
<%

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if Not dcrs.EOF then
		sCheckFinYear = "1"
	else
		sCheckFinYear = "0"
	end if
	dcrs.Close

	arrFin = split(sFinFrom,"/")
	sFinFrom = arrFin(1)&arrFin(2)
	erase arrFin
	arrFin = split(sFinTo,"/")
	sFinTo = arrFin(1)&arrFin(2)

	sFinTo = sMonYr



%>


<BODY leftMargin=0 topMargin=0 onLoad="GetStockDet();Init()">
<form method="POST" name="formname" action="" TARGET="bodyFrame" >

<input type=hidden name="hRecNum" value="">
<input type="hidden" name="hGroup">
<input type="hidden" name="hLevel">
<input type="hidden" name="hArr" value="<%=sArr%>">
<input type="hidden" name="hArrList" value="<%=iAttribList%>">
<input type="hidden" name="hFinYear" value="<%=sCheckFinYear%>">
<input type="hidden" name="selUnit" value="">
<input type=hidden name="hUnitID" value="<%=sUnitID%>">
<input type=hidden name="hItemType" value="<%=sIType%>">
<input type=hidden name="hItemTypeCode" value="<%=sITypeCode%>">
<input type=hidden name="hItemCode" value="<%=iItemCode%>">
<input type=hidden name="hFlag" value="<%=Flag%>">
<input type=hidden name="hUomFlag" value="<%=UomFlag%>">
<input type=hidden name="hItmComCode" value="<%=sCode%>">
<input type=hidden name="hItemDesc" value="<%=sDesc%>">
<input type=hidden name="hClassCode" value="<%=sClassCode%>">

<input type=hidden name="hReplicateItem" value="N">
<input type="hidden" name="hPurTaxType" value="<%=sPurTaxType%>">
<input type="hidden" name="hSalTaxType" value="<%=sSalTaxType%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Item Creation and Definition
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack"></td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
				    <tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack"></td>
								</tr>
								<tr>
									<td align="center" width="5"></td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;</td>
															<td class="GroupTitle" width="55">
																<p align="center">Details
															</td>
															<td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>
													</table>
												</td>
											</tr>

											<tr>
												<td class="GroupTable"><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="MiddlePack"></td>
														</tr>

														<tr>
															<td class="FieldCellSub">
																<table border="0" cellspacing="0" cellpadding="0">
																    <tr>
																        <td valign="top" >
																            <table border="0" cellpadding="0" cellspacing="0">
																                <tr>
																                    <td class="FieldCell" >Item Type</td>
																		            <td class="FieldCellSub" >
																			            <select size="1" name="selIType" class="FormElem" onChange="LetIType(this);GetAttr();ChangeLabel(this)">
																				            <option value="select">Select</option>
																			            <%	'Calling the Function which populates the Item Type list
																				            'populateItemType
																				            'populateItemTypeSelected(sITypeCode)
																				            popItemTypesNew()
																			            %>
																			            </select>
																		            </td>
																                </tr>
																                <tr>
																                    <td class="FieldCell" >Item Code</td>
																		            <td class="FieldCellSub" >
																			            <input type="text" name="txtitmCode" value="<%=sCompItemCode%>" size="19" maxlength=15 class="Formelem"  onblur="CheckAvailability(this,'ItemCode')">&nbsp;&nbsp;
																			            <input type="button" value="Code Create" name="btnYrnCode" class="AddButton" onClick="CreateItemCode(this.value)" disabled>
																			            <!--<input type="button" value="Existing" class="AddButtonX" onClick="DisplayItemCode()" id=button1 name=button1>-->
																		            </td>
																                </tr>
																                <tr>
																                    <td class="FieldCell">Description</td>
																		            <td class="FieldCellSub" >
																			            <input type="text" name="txtItmDesc" Value="<%=sDesc%>" size="40" maxlength="60" class="Formelem"  onblur="CheckAvailability(this,'ItemName')">
																		            </td>
																                </tr>
																                <tr>
																		            <td class="FieldCell">Add. Description</td>
																		            <td class="FieldCellSub">
																		                <input type=text name="txtItmAddDesc" size="40" maxlength=60 class="Formelem">
																		                <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" alt="Manage Item Spec." onClick="ItemSpecPop()">
																		            </td>
																	            </tr>
																            </table>
																        </td>
																        <td valign="top">
																            <table cellpadding="0" cellspacing="0" border="0">
																                <tr>
																		            <td class="FieldCellSub">Category</td>
																		            <td class="FieldCellSub">
																		                <input type="hidden" name="hCategory" value="<%=sCategoryCode%>">
																			            <span id="spanCategory" class="DataOnly"><%=sCategoryName%></span>
																			            <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="EditClass()" alt="Click to Edit the Classification Code">
																		            </td>
																	            </tr>

																	            <tr>
																		            <td class="FieldCellSub">Classification</td>
																		            <td class="FieldCellSub">
																			            <input type="text" name="txtClass" class="Formelem" value="<%Response.Write sClassName%>" disabled>
																		            </td>
																	            </tr>

																	            <tr>
																		            <td class="FieldCellSub">Variant Code</td>
																		            <td class="FieldCellSub">
																			            <input type="text" name="txtVariant" size="12" maxlength=10 class="Formelem">
																		            </td>
																	            </tr>
																	            <tr>
																	                <td class="FieldCellsub">Stores UoM</td>
																		            <td class="FieldCellSub">
																			            <select size="1" name="selUoMStores" class="FormElem">
																				            <option value="select">Select</option>
																				            <%	'Calling the Function which populates the UoM list
																					            'populateUoM
																					            populateUoMSelected(sStoreUom)
																				            %>
																			            </select>
																			            <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="UoMDetails()" alt="Alternate UoM and other UoM's">
																		            </td>
																	            </tr>
																            </table>
																        </td>
																        <td valign="top" rowspan="2">
																            <table cellpadding="0" cellspacing="0" border="0">
																                <tr>
																                    <td class="FieldCell" rowspan="2" colspan="2" valign="top">
																		                <div class="frmBody" id="frm4">
																			                <table border="0" cellspacing="1" class="BodyTable" width="100%" style="height: 81px">
																				                <tr>
																				                    <td class="ExcelHeaderCell" align="center">
																				                        <a href="#" class="ExcelDisplayLink" alt="Click here to upload image" style="cursor:hand;" onclick="UploadImage()">Add</a>&nbsp;
																				                        <a href="#" class="ExcelDisplayLink" alt="Click here to Delete image" style="cursor:hand;" onclick="DeleteImage()">Remove</a>
																				                        
																				                    </td>
																				                </tr>
																				                <tr>
																					                <td class="FieldCellSub">
																					                       <%
																	                                        sQuery = "Select ItemBlowUpPic from INV_M_ItemMaster where ItemCode = "& iItemCode &" and ItemBlowUpPic is not null"
																	                                        dcrs.open sQuery,con
																	                                        if not dcrs.eof then
                																	                            
																	                                        %>
																	                                            <img src="../../Common/GetItemImageFromDB.asp?ID=<%=iItemCode%>" width="150" height="170" border="1" />
																	                                        <%
																	                                        else 
																	                                        %>
																	                                            <img src="../../assets/images/NoImage.gif" width="150" height="170" border="1"/>
																	                                        <%
																	                                        end if
																	                                        dcrs.close
																	                                    %>
																					                </td>
																				                </tr>
																				            </table>
																			            </div>
																		            </td>
																                </tr>
																            </table>
																        </td>
																    </tr>
																	<tr>
																		<td class="FieldCell" colspan="2">
																			<div class="frmBody" id="frm1" style="width:550; height:100;">
																				<table border="0" cellspacing="1" class="BodyTable" width="100%">
																					<tr>
																						<td class="ExcelDisplayCell" align="center" style="height: 14px">Attributes  [<a href="#" onclick="ManageAttribute()">Manage</a>] </td>
																					</tr>

																					<tr>
																						<td class=FieldCellSub>
																							<table id="tblAttribute" cellpadding="0" cellspacing="0" border="0" width="100%">
																							    
																							</table>
																						</td>
																					</tr>

																				</table>
																			</div>
                                                                        </td>
																	</tr>
																	<tr>
																		
																	</tr>
																	<!--tr>
																		<td class="FieldCell" style="width: 82px">Desc. Comprises</td>
																		<td class="FieldCellSub" colspan=4>
																			<input type="checkbox" value="I" name="chkcomIC" class="Formelem"> Item Code
																			<input type="checkbox" value="I" name="chkcomIC" class="Formelem"> Item Description
																			<input type="checkbox" value="I" name="chkcomIC" class="Formelem"> Attributes
																		</td>
																	</tr>
																	<tr>
																		<td class="FieldCell" style="width: 82px">Delimited by</td>
																		<td class="FieldCellSub" colspan=4>
																			<input type="text" name="txtdelimiter" size=2 maxlength=1 class="Formelem"> Delimiter mentioned is used to separate each of above options selected
																		</td>
																	</tr-->
																</table>
															</td>
														</tr>
													</table>
                                                  </center>
												</td>
											</tr>
										</table>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;</td>
															<td class="GroupTitle" width="55">
																<p align="center">Controls
															</td>
															<td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>
													</table>
												</td>
												<td width="15">
												</td>
												<td>
												    <table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;</td>
															<td class="GroupTitle" width="110">
																<p align="center">Replenishment
															</td>
															<td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>
													</table>
												</td>
											</tr>

											<tr>
												<td class="GroupTable">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class="MiddlePack"></td>
														</tr>

														<tr>
															<td class="FieldCellSub">
																<table border="0" cellspacing="0" cellpadding="0">
																	<tr>
																		<td class="FieldCell">Receipt Numbering</td>
																		<td class="FieldCellSub">
																			<select size="1" name="selRecNum" class="FormElem" onChange="CheckNoSeries(this)">
																				<option value="select">Select</option>
																				<option value="L" <%=LRec%>>Lot</option>
																				<option value="S" <%=SRec%>>Serial</option>
																				<option value="LS" <%=LSRec%>>Lot / Serial</option>
																				<option value="N" <%=NRec%>>None</option>
																			</select>
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Receipt Routing</td>
																		<td class="FieldCell">
																			<select size="1" name="selRecRout" class="FormElem">
																				<option value="select">Select</option>
																				<option value="S" <%=Stk%>>Stock</option>
																				<option value="IS" <%=InsS%>>Inspection / Stock</option>
																			</select>
																		</td>
																	</tr>

																	<tr>
																		<td class="FieldCell">Accounting Type</td>
																		<td class="FieldCellSub">
																			<select size="1" name="selAcc" class="FormElem">
																				<option value="select">Select</option>
																				<option value="L" <%=Lf%>>LIFO</option>
																				<option value="F" <%=Ff%>>FIFO</option>
																				<option value="W" <%=WA%>>Weighted Average</option>
																			</select>
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Modvat Eligibility</td>
																		<td class="FieldCell">
																			<input type="radio" value="1" name="radMod" class="FormElem" <%if sModVatEligible=1 then response.write "checked" %>>   Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																			<input type="radio" value="0" name="radMod" class="FormElem" <%if sModVatEligible=0 then response.write "checked" %>>   No
																		</td>
																	</tr>

																	<tr>
																		<td class="FieldCell">Account Head</td>
																		<td class="FieldCellSub">
																		    <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="SelectAccHead()" alt="Account Head">
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">BoM Applicability</td>
																		<td class="FieldCell">
																			<input type="radio" value="1" name="radBoM" class="FormElem" <%If iBOMApplicable = 1 then Response.Write "checked"%>>   Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																			<input type="radio" value="0" name="radBoM" class="FormElem" <%If iBOMApplicable = 0 then Response.Write "checked"%>>   No
																			<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="GetDetails()" alt="Select BoM">
																		</td>
																	</tr>
																	
																	<tr>
																		<td class="FieldCell">Opening Acc Head</td>
																		<td class="FieldCellSub">
																		    <span id="spanOpenAccHead" class="DataOnly">&nbsp;
																		    </span>
																		    <input type="hidden" name="hOAH" value="0" />
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Pur. Tax Type </td>
																		<td class="FieldCell">
																			<select name="SelPurTaxType" class="FormElem">
																			    <option value="0">Select</option>
																			    <% 
																			        populatePurTaxType(sPurTaxType)
																			    %>
																			</select>
																		</td>
																	</tr>
																	<tr>
																		<td class="FieldCell">Closing Acc Head</td>
																		<td class="FieldCellSub">
																		    <span id="spanCloseAccHead" class="DataOnly">&nbsp;
																		    </span>
																		    <input type="hidden" name="hCAH" value="0" />
																		</td>
																		<td class="FieldCellSub"></td>
																		<td class="FieldCellSub">Sales Tax Type</td>
																		<td class="FieldCell">
																			<select name="SelSalTaxType" class="FormElem">
																			    <option value="0">Select</option>
																			    <%
																			        populateSalTaxType(sSalTaxType)
																			    %>
																			</select>
																		</td>
																	</tr>

																</table>
															</td>
														</tr>

													</table>
												</td>
												<td width="15">
												</td>
												<td class="GroupTable">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class="MiddlePack"></td>
														</tr>

														<tr>
															<td class="FieldCellSub">
																<table border="0" cellspacing="0" cellpadding="0">
																	<tr>
																		<td class="FieldCell">Reorder Level</td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtReLvl" value="<%=iReOrdLev%>" style="text-align:right" size="10" maxlength=10 class="FormElem" onkeypress="DoKeyPress('Y',10,3)">
																		</td>
																	</tr>
																	<tr>
																		<td class="FieldCell">Reorder Quantity</td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtReQty" value="<%=iReOrdQty%>" style="text-align:right" size="10" maxlength=10 class="FormElem" onkeypress="DoKeyPress('Y',10,3)">
																		</td>
																	</tr>
																	<tr>
																		<td class="FieldCell">Economic Order Quantity</td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtEcQty" size="10" value="<%=iEcoOrdQty%>" style="text-align:right" maxlength=10 class="FormElem" onkeypress="DoKeyPress('Y',10,3)">
																		</td>
																	</tr>
																</table>
															</td>
														</tr>

													</table>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>
								
								<!--<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;</td>
															<td class="GroupTitle" width="90">
																<p align="center">Other Details
															</td>
															<td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>

													</table>
												</td>
											</tr>

											<tr>
												<td class="GroupTable">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class="MiddlePack"></td>
														</tr>

														<tr>
															<td class="FieldCellSub">
																<table border="0" cellspacing="0" cellpadding="0">
																	<tr>
																		<td class="FieldCell"><span id="idCat">Catalogue No.</span></td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtItmCat" value="<%=iCatalogNo%>" size="20" maxlength=30 class="Formelem">
																		</td>
																		<td class="FieldCell"><span id="idDrw">Draw. Ver</span></td>
																		<td class="FieldCellSub">
																			<input type="text" name="txtItmDrw" value="<%=iDrawingNo%>" size="20" maxlength=20 class="Formelem">
																		</td>
																		<td class="FieldCellSub" colspan="3"></td>
																	</tr>
																</table>
															</td>
														</tr>

													</table>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

 								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>-->

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;
															</td>
															<td class="GroupTitle" width="98">
																<p align="center">Opening Stock
															</td>
														    <td class="GroupTitleRight">
															    <p align="left">&nbsp;
														    </td>
													</tr>

												</table>
											</td>
										</tr>
										<tr>
											<td class="GroupTable">
											<center>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="MiddlePack">
														</td>
													</tr>
														<tr>
															<td class="FieldCellSub">Storage Location &nbsp;
 																<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" onclick="GetStore()" alt="Storage Location">
																<span class="DataOnly" id="idStore"></span>
															</td>
														</tr>

													<tr>
														<td class="FieldCellSub">
															<div class="frmBody" id="frm2" style="width: 570; height:75;">
																<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
																	<tr>
																		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
																		<td class="ExcelHeaderCell" align="center">Storage</td>
																		<td class="ExcelHeaderCell" align="center">MonthYear</td>
																		<td class="ExcelHeaderCell" align="center">Quantity</td>
																		<td class="ExcelHeaderCell" align="center">Value</td>
																		<td class="ExcelHeaderCell" align="center" width="25">Lot/Serial</td>
																	</tr>

																</table>
															</div>
														</td>
													</tr>
													</table>
												</center>

											</td>
										</tr>
										<!--<tr>
										    <td class="FieldCell"><input type=checkbox name=ChkAllUnit value="ALL" class="FormElem">&nbsp;Applicable for All Units</input></td>
										</tr>-->
										

										</table>
									</td>
									<td align="center">
									</td>
								</tr>
								
															<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%"><center>
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10">&nbsp;
															</td>
															<td class="GroupTitle" width="60">
																<p align="center">Pricing
															</td>
														</center><td class="GroupTitleRight">
															<p align="left">&nbsp;
														</td>
													</tr>

												</table>
											</td>
										</tr>
										<tr>
											<td class="GroupTable"><center>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="MiddlePack">
														</td>
													</tr>
													<tr>
														<td class="FieldCellSub">
															<div class="frmBody" id="Div1" style="width: 750; height:75;">
																<table border="0" cellspacing="1" id="Table1" class="ExcelTable" width="100%">
																    <tr>
																        <td class="ExcelHeaderCell" colspan ="2" align="center">
																            Purchase
																        </td>
																        <td class="ExcelHeaderCell" colspan ="2" align="center">
																            Charges
																        </td>
																        <td class="ExcelHeaderCell" colspan ="2" align="center">
																            Margin
																        </td>
																        <td class="ExcelHeaderCell" rowspan="2" align="center">
																            Price
																        </td>
																        <td class="ExcelHeaderCell" rowspan="2" align="center">
																            Effective from
																        </td>
																    </tr>
																	<tr>
																		<td class="ExcelHeaderCell" align="center">
																		    Rate
																		</td>
																		<td class="ExcelHeaderCell" align="center">
																		    Per Unit
																		</td>
																		<td class="ExcelHeaderCell" align="center">
																		    %
																		</td>
																		<td class="ExcelHeaderCell" align="center">
																		    Value
																		</td>
																		<td class="ExcelHeaderCell" align="center">
																		    %
																		</td>
																		<td class="ExcelHeaderCell" align="center">
																		    Value
																		</td>
																	</tr>
																	<%
																	            Dim rsTemp
																			    Dim sRate,nMarPer,nMarVal,nOthPer,nOthVal,nPrice,sRatePer,sEffFrom
																			    
																			    set rsTemp = server.createObject("ADODB.Recordset")
																			    sRate   = 0
								                                                nMarPer = 0
								                                                nMarVal = 0
								                                                nOthPer = 0
								                                                nOthVal = 0
								                                                nPrice  = 0
								                                                sRatePer = 1
								                                                
								                                                sQuery = "Select SellingPriceno,Convert(Char,AsonDate,103) From Sal_M_UnitPriceHdr "
								                                                rsTemp.Open sQuery,Con
								                                                IF Not rsTemp.EOF Then
								                                                    iSellNo = rsTemp(0)
								                                                End If
								                                                rsTemp.close
								                                                
								                                                sRate = FormatNumber(GetItemRate(sUnitID,sFinPeriod,sClassCode,iItemCode,"WA"))
								                                                
								                                                sQuery = " Select isNull(ItemRate,0),isNull(MarginPercent,0),isNull(MarginValue,0),isNull(OtherPercent,0),isNull(OtherValue,0),isNull(ItemPrice,0),convert(Varchar,EffectiveFrom,103),RatePer"&_
										                                                " From Sal_M_UnitPriceDet Where Itemcode = "&iItemCode&" and ClassificationCode = "& sClassCode &" and OudefinitionID = '"&sUnitID&"' "
								                                                'Response.Write"<textarea>"&sQuery&"</textarea>"
								                                                rsTemp.Open sQuery,Con
								                                                IF Not rsTemp.EOF Then
								                                                    if cdbl(sRate)=0 then
								                                                        sRate = rsTemp(0)
								                                                    end if
									                                                nMarPer=rsTemp(1)
									                                                nMarVal=rsTemp(2)
									                                                nOthPer=rsTemp(3)
									                                                nOthVal=rsTemp(4)
									                                                nPrice=rsTemp(5)
									                                                sEffFrom	=rsTemp(6)
									                                                sRatePer = rsTemp(7)
								                                                End IF
								                                                rsTemp.Close
																			
																			%>
																	<tr>
																		<td class="ExcelDisplayCell" align="center">
																		    <input type="text" name="txtPurRate" value="<%=sRate%>" class="formelem" size="15" style="text-align:right" onblur="AssaignValue()">
																		</td>
																		<td class="ExcelDisplayCell" align="center">
																		    <input type="text" name="txtPurRatePer" value="<%=sRatePer%>" class="formelem"  size="3" style="text-align:right" >&nbsp;
																		    <span id="spanPurUOM" class="dataonly"></span>
																		</td>
																		<td class="ExcelDisplayCell" align="center">
																		    <input type="text" name="txtChaPer" value="<%=nOthPer%>" class="formelem" size="3" style="text-align:right" onblur="CalcValue('OP')">
																		</td>
																		<td class="ExcelDisplayCell" align="center">
																		    <input type="text" name="txtChaVal" value="<%=nOthVal%>" class="formelem"  size="15" style="text-align:right" onblur="CalcValue('OV')">
																		</td>
																		<td class="ExcelDisplayCell" align="center">
																		    <input type="text" name="txtMarPer" value="<%=nMarPer%>" class="formelem" size="3" style="text-align:right" onblur="CalcValue('MP')">
																		</td>
																		<td class="ExcelDisplayCell" align="center">
																		    <input type="text" name="txtMarVal"value="<%=nMarVal%>" class="formelem"  size="15" style="text-align:right" onblur="CalcValue('MV')">
																		</td>
																		<td class="ExcelDisplayCell" align="center">
																		    <input type="text" name="txtTotPrice" value="<%=nPrice%>" class="formelem"  size="15" style="text-align:right" >
																		    <input type="hidden" name="hEffFrom" value="<%=sEffFrom%>">
																		</td>
																		<td class="ExcelDisplayCell" align="center">
																		    <%
																		        InsertDatePicker("ctlEffDate")
																		    %>
																		</td>
																	</tr>
																</table>
															</div>
														</td>
													</tr>

												</center>
												</table>
											</td>
										</tr>
										
										<tr>
							                <td valign="top" width="100%" colspan="3">
							                    <table border=0 cellpadding=0 cellspacing =0>
							                        <tr>
							                            <td class="FieldCellSub">Item Status
							                            </td>
							                            <td class="FieldCellSub">&nbsp;&nbsp;
							                                <input type=radio name=radStatus value="AC" <%if trim(sStatus)="AC" then response.write "Checked" %> >Active&nbsp;&nbsp;
							                                <input type=radio name=radStatus value="OH" <%if trim(sStatus)="OH" then response.write "Checked" %>>On Hold&nbsp;&nbsp;
							                                <input type=radio name=radStatus value="IA" <%if trim(sStatus)="IA" then response.write "Checked" %>>In Active&nbsp;&nbsp;
							                                <input type=radio name=radStatus value="NS">Not for Sale&nbsp;&nbsp;
							                                <input type=radio name=radStatus value="DS" <%if trim(sStatus)="DS" then response.write "Checked"%>>Dead Stock&nbsp;&nbsp;
							                            </td>
							                        </tr>
							                    </table>
                    		                </td>
                                        </tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td valign="top" width="100%" colspan="3">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="button" value="Replicate Other Unit" name="But_Replicate" class="ActionButtonX" onClick="Replicate_Item()">
													<input type="button" value="Save" name="B2" class="ActionButton" onClick="CheckSubmitDetails('<%=sFinFrom%>','<%=sFinTo%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
 													<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
												</td>
											</tr>

										</table>
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="BottomPack">
									</td>
								</tr>

							</table>
	</form>
</body>
</html>
<%
Function popItemTypesNew()
Dim rsTemp,ssql
set rsTemp = Server.CreateObject("ADODB.Recordset")
ssql = "Select ItemTypeID,ItemTypeDescription from INV_M_ItemTypes"
rsTemp.Open ssql,con
if not rsTemp.EOF then
    do while not rsTemp.EOF 
        Response.Write "<option value="&Trim(rsTemp(0))&">"&Trim(rsTemp(1))&"</option>"
        
        rsTemp.MoveNext 
    loop
end if
rsTemp.Close
End Function
%>

<%
Function populatePurTaxType(sPurTaxType)
Dim rsTemp,sQuery
set rsTemp = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select PurchaseType,PurchaseTypeName from APP_M_PurchaseTypes where upper(isNull(Active,'Y')) = 'Y'  ORDER BY PURCHASETYPE"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        do while not rsTemp.eof
            if trim(sPurTaxType)=rsTemp(0) then
                response.write "<option value="&trim(rsTemp(0))&" selected>"& trim(rsTemp(1)) &"</option>"
            else
                response.write "<option value="&trim(rsTemp(0))&">"& trim(rsTemp(1)) &"</option>"
            end if
            rsTemp.movenext
        loop
    end if
    rsTemp.close
End Function
%>

<%
Function populateSalTaxType(sSalTaxType)
Dim rsTemp,sQuery
set rsTemp = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select InvoiceType,InvoiceTypeName from Sal_M_InvoiceTypes where isNull(Useable,0)=1  Order by InvoiceType"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        do while not rsTemp.eof
            if trim(sSalTaxType)=rsTemp(0) then
                response.write "<option value="&trim(rsTemp(0))&" selected>"& trim(rsTemp(1)) &"</option>"
            else
                response.write "<option value="&trim(rsTemp(0))&">"& trim(rsTemp(1)) &"</option>"
            end if
            rsTemp.movenext
        loop
    end if
    rsTemp.close
End Function
%>
