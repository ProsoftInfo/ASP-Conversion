<!--#include File="NoSeries.asp" -->
<!--#include File="MRInsertCommon.asp" -->
<!--#include File="mrsIssueInsertCommon.asp" -->
<!--#include File="NoSeriesCommonFunctions.asp"-->
<!--#include File="CommonFunctions.asp" -->
<%
Function POInsert()
	dim adoCmd,sPono
	Dim POXML,IssueXML,newXML,SUBPROXML
	 
	'Declaration of Objects
	Set adoCmd = Server.CreateObject("ADODB.Command")
	Set adoCmd.ActiveConnection = con   
	Dim dcrs2,sOrgCode,sOrdFor,sOrdValDate,sApprover,sConfReq,sWithMat,iPONo
	Dim ItemNode,sItmCode,sPackAvail,sDrwVerNo,sCatalogue,sVariantCode,sMREntry,sIssueEntry
	Dim sCategory,sTempDesc,iRate,sRateUnit,iRatePer,sMRSNo,sGPEntry,iForPurNo
	dim Pdate,sQuery
	dim iItmCode,sitmShDesc,sitmDesc,sitmAddDesc,sitmUsage,iitmController,sStUoM,sPuUoM
	dim sFormCode,sItmWho,sItmActive,sItmStock,sSalEli,sManEli,sPurEli,iCtr,sWho
	dim objFSO,objTxt,iItemCount,MRxml,ndShipLoc,iShEntNo,iSItemCode,iSClassCode,iSReqBy,sSUnitCode,sSReqQty,sSchType,sSchSchType
	''Added By Ragav on July 16,2010 for Tax Percentage 
	Dim ndTax,sTaxMode,ndAddDet
	''
	Dim sPurExp,iPurNoSeriesClassCode,ndPurTemp,sRecNum,sRecRout,sAccType,sModvat
	'Dim sPuUoM,sMaUoM,sSaUoM,iStToPur,iStToMan,iStToSal,sStToPurOp,sStToManOp,sStToSalOp
	Dim sDelInst,sSchRelReq,sNoneCnt,sLSCnt,sSCnt,sItmRcptNum,sMRSEQNO
	
	
	Dim sArrStoreInfo,sLocNo,sBinNo,sPOItemCode,ndPickSubNode
	Dim ndSUBRoot,sSubExp,ndSubSelNodes
	
	

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

	' Create our DOM Document Objects
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")
	Set MRxml = Server.CreateObject("Microsoft.XMLDOM")
	
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set SUBPROXML = Server.CreateObject("Microsoft.XMLDOM")
	if objFSO.FileExists(server.MapPath("../../purchase/xmldata/SubcontractProcessFlow.xml")) then
	    SUBPROXML.load Server.MapPath("../../purchase/xmldata/SubcontractProcessFlow.xml")
	    set ndSUBRoot = SUBPROXML.documentElement
	    sSubExp = "//Root/Activity[@id=3]"
	    set ndSubSelNodes = ndSUBRoot.Selectnodes(sSubExp)
	    if ndSubSelNodes.length>0 then
	        sMRSEQNO = ndSubSelNodes.Item(0).Attributes.getNamedItem("seqno").value
	    end if 
        Response.Write "<p>sMRSEQNO = "& sMRSEQNO	    
	end if 

	iitmController = getUserid
	newxml.load server.MapPath("../temp/transaction/PO_PUR_"&Session.sessionID&".xml")
	Set RootNode = newxml.documentElement
	
	sPurExp = "//ITEMDETAILS"
	set ndPurTemp = RootNode.selectNodes(sPurExp)
	if ndPurTemp.length>0 then
	    iPurNoSeriesClassCode = ndPurTemp.Item(0).Attributes.getNamedItem("CLASSCODE").value
	end if

	
	iForPurNo = Request("ForPurNo")
	
	sExp ="//HEADER"
	Set HeaderNode = RootNode.selectSingleNode(sExp)
	sitmType = trim(HeaderNode.Attributes.getNamedItem("ITEMTYPE").Value)
	sOrgCode = trim(HeaderNode.Attributes.getNamedItem("FORUNIT").Value)
	Pdate = trim(HeaderNode.Attributes.getNamedItem("CREATEDON").Value)
	sWho = trim(HeaderNode.Attributes.getNamedItem("REF").Value)
	sOrdFor = trim(HeaderNode.Attributes.getNamedItem("PURCHASEORDERFOR").Value)
	sOrdValDate = trim(HeaderNode.Attributes.getNamedItem("ORDERVALIDTO").Value)
	sApprover = trim(HeaderNode.Attributes.getNamedItem("APPROVER").Value)
	sConfReq = trim(HeaderNode.Attributes.getNamedItem("CONREQUIRED").Value)
	sWithMat = trim(HeaderNode.Attributes.getNamedItem("WITHMAT").Value)
	
	sAppRefType = HeaderNode.Attributes.getNamedItem("AppRefType").Value
	sAppRefNo = HeaderNode.Attributes.getNamedItem("AppRefNo").Value	
	sAppRefDate = HeaderNode.Attributes.getNamedItem("AppRefDate").Value
	
	sDelInst =  HeaderNode.Attributes.getNamedItem("delinst").value
	sSchRelReq = HeaderNode.Attributes.getNamedItem("schrelreq").value
	
	
	if trim(sAppRefType)="" or IsNull(sAppRefType) then sAppRefType="0"
	if trim(sAppRefNo)="" or IsNull(sAppRefNo) then sAppRefNo="0"
	if trim(sAppRefDate)="" then sAppRefDate = date
	
	HeaderNode.setAttribute "AppRefType",sAppRefType
	HeaderNode.setAttribute "AppRefNo",sAppRefNo
	HeaderNode.setAttribute "AppRefDate",sAppRefDate
	
    sQuery = "Select isNull(Max(PurchaseOrderNo),0)+1 from PUR_T_POHeader"
    dcrs.Open sQuery,con
    if not dcrs.EOF then
        iPONo = dcrs(0)
    end if
    dcrs.Close 
    
    if Trim(sitmtype)="" or Trim(ucase(sitmType))="NULL"  or IsNull(sitmtype) then sitmtype="NULL"
    
    'Response.Write "iPONo = "& iPONo
	
	'' To clarify whether No. series to be created for BY UNIT or
	
	
	    sTempSeries = GetPurNumberSeriesCodes("4",sOrgCode,iPurNoSeriesClassCode)    
        'Response.Write "<p> TempSeries = "& sTempSeries
        sArrSeries = Split(sTempSeries,":")
        iSeriesNo = sArrSeries(0)
        iSeriesCode = sArrSeries(1)
        
        sSql = "Select GroupName from Inv_M_Classification where GroupCode = "& iPurNoSeriesClassCode 
        rsTemp.Open sSql,con 
        if not rsTemp.EOF then
            sNumClassName = trim(rsTemp(0))
        end if
        rsTemp.Close 
        

    if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
        Response.Clear 
        'Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Purchase Order - "& sNumClassName &"  Classification </H2></p>"
        Response.Write "Number Series is not defined for Purchase Order - "& sNumClassName &"  Classification"
	    Response.End 
    end if
    
    if not CheckNoSerAvilForThisYear(sOrgCode,iSeriesNo,iSeriesCode,Pdate) then
        Response.Clear 
        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Purchase Order - "& sNumClassName &"  Classification for this Year </H2></p>"
        Response.End 
    end if
    
		
	sPono = GenSeriesNumberItemWise(sOrgCode,iSeriesNo,iSeriesCode,Pdate)
	
	
	
	sExp ="//SHIPTOHEADER/LocQty"
	Set ndShipLoc = RootNode.Selectnodes(sExp)
	
	if ndShipLoc.Length > 0 then
		For iCtr = 0 to ndShipLoc.Length - 1
		    iShEntNo = trim(ndShipLoc.Item(iCtr).Attributes.getNamedItem("EntryNo").Value)
		    iSItemCode = trim(ndShipLoc.Item(iCtr).Attributes.getNamedItem("ItemCode").Value)
		    iSClassCode = trim(ndShipLoc.Item(iCtr).Attributes.getNamedItem("ClassCode").Value)
		    iSReqBy = trim(ndShipLoc.Item(iCtr).Attributes.getNamedItem("ReqBy").Value)
		    sSUnitCode = trim(ndShipLoc.Item(iCtr).Attributes.getNamedItem("Unit").Value)
		    sSReqQty = trim(ndShipLoc.Item(iCtr).Attributes.getNamedItem("ReqQty").Value)
		    sSchType = trim(ndShipLoc.Item(iCtr).Attributes.getNamedItem("SchedType").Value)
		    
		    if sConfReq = "Y" then
		        sQuery = "Insert into PUR_T_POShipTo (PurchaseOrderNo,EntryNo,ItemCode,ClassificationCode,"&_
		             " OrganisationCode,ScheduledOn,ShipToUnitCode,ShipToQty)"&_
		             " values("& iPONo &","& iShEntNo &","& iSItemCode &","&iSClassCode &","& pack(sOrgCode) &","& pack(iSReqBy) &","& pack(sSUnitCode) &","& sSReqQty &")"
		    else
		        sQuery = "Insert into PUR_T_POShipTo (PurchaseOrderNo,EntryNo,ItemCode,ClassificationCode,"&_
		             " OrganisationCode,ScheduledOn,ShipToUnitCode,ShipToQty,ShipToQtyConfirmed)"&_
		             " values("& iPONo &","& iShEntNo &","& iSItemCode &","&iSClassCode &","& pack(sOrgCode) &","& Pack(iSReqBy) &","& pack(sSUnitCode) &","& sSReqQty &","& sSReqQty &")"
		    end if
		    
		    Response.Write sQuery
		    con.execute sQuery
		    
			
		Next
	end if 'if ndShipLoc.Length > 0 then
	

	sExp ="//ITEMDETAILS "'[ @CLASSCODE = '00']"
	Set ItemNode = RootNode.Selectnodes(sExp)
	
	if ItemNode.Length > 0 then
		For iCtr = 0 to ItemNode.Length - 1
			sitmCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("TEMPICODE").Value)
			iClass  = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CLASSCODE").Value)
			sOrgCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UNIT").Value)
			sitmShDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("TEMPISHDESC").Value)
			sitmDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMNAME").Value)
			sitmAddDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("TEMPIADDDESC").Value)
			sPackAvail = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("PACKING").Value)
			sRecBy = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("REQUIREDBY").Value)
			sDrwVerNo = "NULL"
			sCatalogue = "NULL"
			sVariantCode = "NULL"
			sStUoM = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UOM").Value)
			sCategory = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CATEGORY").Value)
'			sMatItemDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMDESC").Value)			
			'sAddnDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ADDNDESCRIPTION").Value)			
			sTempDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ADDNDESCRIPTION").Value)			
			sAttributeList = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ATTRIBUTELIST").Value)			
			iRate = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("RATE").Value)
			sRateUnit = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UNITPER").Value)
			iRatePer = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UNITRATE").Value)			
		Next
	end if
	sExp ="//ITEMDETAILS [ @CLASSCODE = '00']"
	Set ItemNode = RootNode.Selectnodes(sExp)
	
	if ItemNode.Length > 0 then
		For iCtr = 0 to ItemNode.Length - 1
			sitmCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("TEMPICODE").Value)
			iClass  = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CLASSCODE").Value)
			sOrgCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UNIT").Value)
			sitmShDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("TEMPISHDESC").Value)
			sitmDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMNAME").Value)
			sitmAddDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("TEMPIADDDESC").Value)
			sPackAvail = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("PACKING").Value)
			sRecBy = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("REQUIREDBY").Value)
			sDrwVerNo = "NULL"
			sCatalogue = "NULL"
			sVariantCode = "NULL"
			sStUoM = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UOM").Value)
			sCategory = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CATEGORY").Value)
'			sMatItemDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMDESC").Value)			
			'sAddnDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ADDNDESCRIPTION").Value)	
			sTempDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ADDNDESCRIPTION").Value)			
			sAttributeList = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ATTRIBUTELIST").Value)			
			iRate = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("RATE").Value)
			sRateUnit = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UNITPER").Value)
			iRatePer = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UNITRATE").Value)			
			
			sRecNum = "LS"
			sRecRout = "S"
			sAccType = "W"
			sModvat = "0"

			sPuUoM = sStUoM
			sMaUoM = ""
			sSaUoM = ""
			iStToPur = "1"
			iStToMan = ""
			iStToSal = ""
			sStToPurOp = "0"
			sStToManOp = ""
			sStToSalOp = ""
			
			Response.Write "sAttributeList="& sAttributeList
			'Newly Added to Update Packing Details
				IF trim(sPackAvail) = "Y" then
					For each Packnode in ItemNode.childnodes
						If trim(Packnode.nodename) = trim("PACKING") then
							If Packnode.getAttribute("ItemCode") = trim(sitmCode) then
								For each subnode in Packnode.childnodes
									If trim(subnode.nodename) = trim("Pack") then
										iPackCode = subnode.getAttibute("Packingcode")
										iPackQty = subnode.getAttibute("PackingQty")
										iPackWght =  subnode.getAttibute("PackingWght")
										iSelForm = subnode.getAttibute("SellingForm")
										iFrmWght = subnode.getAttibute("FormWeight")
										iNoOfPack = subnode.getAttibute("NoOfPack")
										'Response.Write "iPackCode="& iPackCode
									End If
								Next
							End If 
						End If	
					Next
				End If 'IF trim(sPackAvail) = "Y" then	
			
		
			'sExp ="//PACKING [ @ItemCode = "&sitmCode&"]"					
			'Set PackNode = RootNode.Selectnodes(sExp)
			'if PackNode.Length > 0 then
						
			'End if
			'Response.Write "sitmCode"&<BR>

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

			if sitmType = "PLA" then
				sItmActive = "'Y'"
				sItmStock = "'N'"
			else
				sItmActive = "'Y'"
				sItmStock = "'S'"
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
				.Source = "SELECT ISNULL(MAX(ITEMCODE),0) + 1 FROM INV_M_ITEMMASTER WHERE (ITEMCODE = (SELECT ISNULL(MAX(ITEMCODE), 0) FROM INV_M_ITEMMASTER))"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			iItmCode = dcrs(0)

			dcrs.close
		    If cint(iItmCode) = 1 Then
				sSql = "INSERT INTO INV_M_ITEMMASTER (ITEMCODE,COMPANYITEMCODE,SHORTDESCRIPTION," &_
					"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,DRAWINGNUMBER,PURCHASEELIGIBLE,MANUFACTUREELIGIBLE, " &_
					"SALESELIGIBLE,INVENTORYELIGIBLE,STORESUOM,PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR, " &_
					"MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE,SALETOSTOREOPERATOR, " &_
					"ITEMCONTROLLER,ITEMCREATEDON,ITEMCREATEDBY,ITEMDELETED,ITEMSOURCE,STOCKNONSTOCK,ITEMACTIVE,CATALOGUENO, " &_
					"ITEMTYPEID,CATEGORYCODE) VALUES " &_
					"(" & iItmCode & "," & Pack(ucase(sitmCode)) & "," & sitmShDesc & "," &_
					"" & Pack(sitmDesc) & "," & sitmAddDesc & ",NULL," & sPurEli & "," & sManEli & "," &_
					"" & sSalEli & ",'1'," & Pack(sStUoM) & "," & sPuUoM & "," & iStToPur & "," & sStToPurOp & "," &_
					"" & sMaUoM & "," & iStToMan & "," & sStToManOp & "," & sSaUoM & "," &_
					"" & iStToSal & "," & sStToSalOp & "," & iitmController & ",CONVERT(DATETIME,GETDATE(),103)," &_
					"" & iitmController & ",'0','T'," & sItmStock & "," & sItmActive & "," & sCatalogue & "," & Pack(sitmType) & "," & Pack(sCategory) & ")"
				Response.Write sSql &"<BR><BR>"
				con.Execute sSql

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				iClass = "0"

				sParentCode = iClass

				sParentCode = sitmType & ":" & sCategory & ":" & sParentCode

				'Inserting Group
				sSql = "INSERT INTO INV_M_ITEMGROUP (ITEMCODE, CLASSIFICATIONCODE, LEAFNODE, " _
				& "ITEMPATH) VALUES (" & iItmCode & "," & iClass & ",1,'" & sParentCode & "')"

				Response.Write sSql & "<BR>"
				con.Execute sSql
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			else
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT COUNT(ITEMCODE)+1 FROM INV_M_ITEMMASTER WHERE LOWER(ITEMDESCRIPTION) like " & Pack(lcase(sitmDesc)&"%")
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				If Not dcrs.EOF Then
					iItemCount = dcrs(0)
					if iItemCount > 1 then sitmDesc = sitmDesc & "("&iItemCount &")"
				End If
				dcrs.Close

				sSql = "INSERT INTO INV_M_ITEMMASTER (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,COMPANYITEMCODE,SHORTDESCRIPTION," &_
					"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,DRAWINGNUMBER,PURCHASEELIGIBLE,MANUFACTUREELIGIBLE, " &_
					"SALESELIGIBLE,STORESUOM,PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR, " &_
					"MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE,SALETOSTOREOPERATOR, " &_
					"ITEMCONTROLLER,ITEMDELETED,ITEMSOURCE,STOCKNONSTOCK,ITEMACTIVE,CATALOGUENO, " &_
					"ITEMTYPEID,CATEGORYCODE) VALUES " &_
					"(" & iItmCode & ","& iClass &"," & Pack(sOrgCode) & "," & Pack(ucase(sitmCode)) & "," & sitmShDesc & "," &_
					"" & Pack(sitmDesc) & "," & sitmAddDesc & ",NULL," & sPurEli & "," & sManEli & "," &_
					"" & sSalEli & "," & Pack(sStUoM) & "," & sPuUoM & "," & iStToPur & "," & sStToPurOp & "," &_
					"" & sMaUoM & "," & iStToMan & "," & sStToManOp & "," & sSaUoM & "," & iStToSal & "," & sStToSalOp & "," &_
					"" & iitmController & ",'0','T'," & sItmStock & "," & sItmActive & "," & sCatalogue & "," & Pack(sitmType) & "," & Pack(sCategory) & ")"
				Response.Write sSql &"<BR><BR>"
				con.Execute sSql

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				iClass = "0"

				sParentCode = iClass

				sParentCode = sitmType & ":" & sCategory & ":" & sParentCode

				'Inserting Group
				sSql = "INSERT INTO INV_M_ITEMGROUP (ITEMCODE, CLASSIFICATIONCODE, LEAFNODE, " _
				& "ITEMPATH) VALUES (" & iItmCode & "," & iClass & ",1,'" & sParentCode & "')"

				Response.Write sSql & "<BR>"
				'con.Execute sSql

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			end if

			ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMCODE").Value = iItmCode
			ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMNAME").Value = sitmDesc

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT ITEMCODE FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
				.Source = "SELECT ITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if dcrs.EOF then
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT * FROM INV_M_ITEMGROUP WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & ""
					.ActiveConnection = con
					.Open
				end with

				if not dcrs1.EOF then
					sSql = "INSERT INTO INV_M_ITEMORGNGROUP (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
						"LEAFNODE,ITEMPATH) VALUES " &_
						"(" & iItmCode & "," & Pack(sOrgCode) & "," & iClass & "," &_
						"" & trim(dcrs1(2)) & "," & Pack(trim(dcrs1(3))) & ")"
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
				dcrs1.Close

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT * FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItmCode & ""
					.ActiveConnection = con
					.Open
				end with

				if not dcrs1.EOF then
					if IsNull(trim(dcrs1(2))) or trim(dcrs1(2)) = "" then
						sItmShDesc = "NULL"
					else
						sitmShDesc = Pack(trim(dcrs1(2)))
					end if

					if IsNull(trim(dcrs1(4))) or trim(dcrs1(4)) = "" then
						sitmAddDesc = "NULL"
					else
						sitmAddDesc = Pack(trim(dcrs1(4)))
					end if

					if IsNull(trim(dcrs1(5))) or trim(dcrs1(5)) = "" then
						sDrwVerNo = "NULL"
					else
						sDrwVerNo = Pack(trim(dcrs1(5)))
					end if

					if IsNull(trim(dcrs1(32))) or trim(dcrs1(32)) = "" then
						sCatalogue = "NULL"
					else
						sCatalogue = Pack(trim(dcrs1(32)))
					end if

					if IsNull(trim(dcrs1(11))) or trim(dcrs1(11)) = "" then
						sPurUoM = "NULL"
					else
						sPurUoM = Pack(trim(dcrs1(11)))
					end if

					if IsNull(trim(dcrs1(14))) or trim(dcrs1(14)) = "" then
						sManUoM = "NULL"
					else
						sManUoM = Pack(trim(dcrs1(14)))
					end if

					if IsNull(trim(dcrs1(17))) or trim(dcrs1(17)) = "" then
						sSalUoM = "NULL"
					else
						sSalUoM = Pack(trim(dcrs1(17)))
					end if

					if IsNull(trim(dcrs1(12))) or trim(dcrs1(12)) = "" then
						iPurRate = "NULL"
					else
						iPurRate = trim(dcrs1(12))
					end if

					if IsNull(trim(dcrs1(15))) or trim(dcrs1(15)) = "" then
						iManRate = "NULL"
					else
						iManRate = trim(dcrs1(15))
					end if

					if IsNull(trim(dcrs1(18))) or trim(dcrs1(18)) = "" then
						iSalRate = "NULL"
					else
						iSalRate = trim(dcrs1(18))
					end if

					sSql = "INSERT INTO INV_M_ITEMORGMASTER (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,COMPANYITEMCODE,SHORTDESCRIPTION," &_
						"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,DRAWINGNUMBER,PURCHASEELIGIBLE,MANUFACTUREELIGIBLE, " &_
						"SALESELIGIBLE,STORESUOM,PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR, " &_
						"MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE,SALETOSTOREOPERATOR, " &_
						"ALLOWMODVATCREDIT,RECEIPTNUMBERING,RECEIPTROUTING,VALUATIONMETHOD, " &_
						"ITEMONHOLD,ITEMCONTROLLER,ITEMDEFINEDON,ITEMDEFINEDBY,ITEMDELETEDBY,CATALOGUENO) VALUES " &_
						"(" & iItmCode & "," & iClass & "," & Pack(sOrgCode) & "," & sVariantCode & "," & sItmShDesc & "," &_
						"" & Pack(trim(dcrs1(3))) & "," & sItmAddDesc & "," & sDrwVerNo & "," & Pack(trim(dcrs1(6))) & "," & Pack(trim(dcrs1(7))) & "," &_
						"" & Pack(trim(dcrs1(8))) & "," & Pack(trim(dcrs1(10))) & "," & sPurUoM & "," & iPurRate & "," & Pack(trim(dcrs1(13))) & "," &_
						"" & sManUoM & "," & iManRate & "," & Pack(trim(dcrs1(16))) & "," & sSalUoM & "," &_
						"" & iSalRate & "," & Pack(trim(dcrs1(19))) & "," & Pack(sModvat) & "," & Pack(sRecNum) & "," &_
						"" & Pack(sRecRout) & "," & Pack(sAccType) & ",0," & trim(dcrs1(20)) & "," &_
						"CONVERT(DATETIME,GETDATE(),103)," & Pack(trim(dcrs1(20)))& ",0," &_
						"" & sCatalogue & ")"
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
				dcrs1.Close
			end if
			dcrs.Close

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_ORGSLBINDETAILS WHERE OUDEFINITIONID = " & Pack(sOrgCode) & ""
				.Source = "SELECT LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM Inv_M_StoreBinDetails WHERE OUDEFINITIONID = " & Pack(sOrgCode) & ""
				
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			do while not dcrs1.EOF

				sLoc = trim(dcrs1(0))
				sBin = trim(dcrs1(1))

				if len(Month(date())) = 1 then
					sTempMonYr = "0"&Month(date())
				else
					sTempMonYr = Month(date())
				end if
				sMonYr = sTempMonYr&Year(date())

				iQty = "0"
				iValue = "0"

				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT APPLICABLEFOR FROM INV_M_ORGSTORAGE WHERE OUDEFINITIONID = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & ""
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
					objTxt.write sSql & vbCrLf & vbCrLf
					Response.Write sSql & vbCrLf
					con.Execute sSql
				end if
				dcrs.Close

				if sBin = "0" then sBin = "NULL"

				sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
					"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
					"(" & Pack(sOrgCode) & "," & iItmCode & "," & iClass & "," &_
					"'RO',NULL,CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103)," & iQty & "," & iValue & ")"
				objTxt.write sSql & vbCrLf & vbCrLf
				Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND MONTHANDYEAR = " & Pack(sMonYr) & ""
					.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME,FINANCIALYEARFROM,103) = CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME,FINANCIALYEARTO,103) = CONVERT(DATETIME," & Pack(sFinTo) & ",103)"
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing

				if dcrs.EOF then
					sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
						"LOCATIONNUMBER,BINNUMBER,FINANCIALYEARFROM,FINANCIALYEARTO,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE) VALUES " &_
						"(" & Pack(sOrgCode) & "," & iClass & "," & iItmCode & "," & sLoc & "," & sBin & ", " &_
						"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & iQty & "," & iValue & ")"
					objTxt.write sSql & vbCrLf & vbCrLf
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				else
					sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQty & ")," &_
						"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ") WHERE ITEMCODE = " & iItmCode & " AND " &_
						"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
						"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
						"CONVERT(DATETIME,FINANCIALYEARFROM,103) = CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND "&_
						"CONVERT(DATETIME,FINANCIALYEARTO,103) = CONVERT(DATETIME," & Pack(sFinTo) & ",103)  "
					objTxt.write sSql & vbCrLf & vbCrLf
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
				dcrs.Close

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				arrFin = split(GetFinancialYear(sMonYr),":")
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
					objTxt.write sSql & vbCrLf & vbCrLf
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
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
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
				dcrs.Close

				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
					.ActiveConnection = con
					.Open
				end with

				set dcrs.ActiveConnection = nothing

				if dcrs.EOF then
					sSql = "INSERT INTO INV_T_ITEMLOCYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
						"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
						"YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
						"(" & Pack(sOrgCode) & "," & iClass & "," & iItmCode & "," &_
						"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
						"" & sLoc & "," & sBin & "," &_
						"" & iQty & "," & iValue & "," & iQty & "," & iValue & ")"
					objTxt.write sSql & vbCrLf & vbCrLf
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql

					with dcrs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(MAX(STOCKNO)+1,1) FROM INV_M_STOCKSTATUS"
						.ActiveConnection = con
						.Open
					end with
					set dcrs2.ActiveConnection = nothing

					if not dcrs2.EOF then
						sSql = "INSERT INTO INV_M_STOCKSTATUS (STOCKNO,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
							"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER) VALUES " &_
							"(" & trim(dcrs2(0)) & "," & Pack(sOrgCode) & "," & iClass & "," & iItmCode & "," &_
							"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
							"" & sLoc & "," & sBin & ")"
						objTxt.write sSql & vbCrLf & vbCrLf
						 Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
					end if
					dcrs2.Close

				end if
				dcrs.Close

			dcrs1.MoveNext
			loop
			dcrs1.Close
		next '	For iCtr = 0 to ItemNode.Length - 1
end if ' if ItemNode.Length > 0 then

'Added by Ragav on July 16,2010 for assigning zero to Tax mode not equal to "P" then
'Begin
sExp ="//PURACC/TaxDetails/Tax"
	Set ndTax = RootNode.Selectnodes(sExp)
	
	if ndTax.Length > 0 then
		For iCtr = 0 to ndTax.Length - 1
			sTaxMode= trim(ndTax.Item(iCtr).Attributes.getNamedItem("TaxMode").Value)
			if sTaxMode<>"P" then
			  ndTax.Item(iCtr).Attributes.getNamedItem("TaxValue").Value =0
			end if ' if sTaxMode<>"P" then
		Next
	end if
'End 

' ItemAttributes  ,	ITEMATTRIBUTES Varchar(30) '@ITEMATTRIBUTE 

	newxml.Save server.MapPath("../temp/transaction/PO1.xml")
	
'	Response.Write newxml.xml
	
	'Response.Write "iRatePer="&iRatePer 
	 'Response.Write "OK"
	
	'Response.end
	
	if trim(iForPurNo)<>"" then
	    ssql = " Update ForPurchaseOrder set PONumber = "& iPONo &" where ForPOEntryNo = "& iForPurNo
	    con.execute ssql
	elseif Trim(iForOrderNo)<>"" then
		ssql = " Update ForPurchaseOrder set PONumber = "& iPONo &" where ForPOEntryNo = "& iForOrderNo
	    con.execute ssql
	end if
	
	adoCmd.CommandText = "Proc_POCreation"	
	adoCmd.CommandType = 4
	'sPono = "PO00001"
	adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(newxml.xml),newxml.xml)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@ProdOrder",129,1,30,sPono)
	adoCmd.Execute() 
	
	
	Dim ndPORoot,ndPOChild,ndIssRoot,ndIssChild,ndIssPick
	set POXML = Server.CreateObject("Microsoft.XMLDOM")
    set IssueXML = Server.CreateObject("Microsoft.XMLDOM")
    POXML.load newxml
    set ndPORoot = POXML.documentElement
	    
	set adoCmd = nothing
	if trim(sSchRelReq)="N" then
	    if trim(sOrdFor)="S" and trim(sAppRefType)<>"12" then
	        sNoneCnt = 0
	        sLSCnt = 0
	        sSCnt = 0
    	    
	        sSql = "Select isNull(AutomaticIssueEntry,'N') from VW_APP_ApplicationSetup where ApplicationCode = 2 and ReferenceCodeNo = 4"
	        dcrs.Open sSql,con
	        if not dcrs.EOF then
	            sIssueEntry = trim(dcrs(0))
	        end if 
	        dcrs.Close
    	    
	        sSql = "Select isNull(AutomaticGatePassEntry,'N') from VW_APP_ApplicationSetup where ApplicationCode = 2 and ReferenceCodeNo = 4"
	        dcrs.Open sSql,con
	        if not dcrs.EOF then
	            sGPEntry = trim(dcrs(0))
	        end if 
	        dcrs.Close
	       ' Response.Write " sIssueEntry = "& sIssueEntry 
	       ' Response.Write " sItmCode = "& sItmCode 
	       ' Response.Write " sClass = "& iClass 
	            if sIssueEntry="Y" and (Trim(iClass)<>"0" and trim(iClass)<>"")then
	                ''Issue Entry 
	                set ndIssRoot = IssueXML.createElement("ISSTYPE")
	                IssueXML.appendChild ndIssRoot
	                For each ndPOChild in ndPORoot.ChildNodes
	                    if ndPOChild.nodeName="HEADER" then
	                        sitmType = ndPOChild.getAttribute("ITEMTYPE")
	                        iCreatedBy = ndPOChild.getAttribute("CREATEDBY")
	                        sWithMat = ndPOChild.getAttribute("WITHMAT")
	                        dCreatedOn = ndPOChild.getAttribute("CREATEDON")
	                        ndIssRoot.setAttribute "ISSTYPE","F"
	                        ndIssRoot.setAttribute "ISSTOTYPE","Party"
	                        ndIssRoot.setAttribute "ISSTOCODE",ndPOChild.getAttribute("SUPPAGENT")
	                        ndIssRoot.setAttribute "ISSTOSUBCODE",""
	                        ndIssRoot.setAttribute "POConfirm","N"
	                        ndIssRoot.setAttribute "SInvConfirm","N"
	                        ndIssRoot.setAttribute "Invoice","A"
	                        if trim(sGPEntry)="Y" then
	                            ndIssRoot.setAttribute "GPConfirm","Y"
	                        else
	                            ndIssRoot.setAttribute "GPConfirm","N"
	                        end if
	                        ndIssRoot.setAttribute "ProConfirm","N"
	                        ndIssRoot.setAttribute "MCallFrom","PO"
	                        ndIssRoot.setAttribute "RedirectTo","POList.asp"
	                        ndIssRoot.setAttribute "AppRefType","21"
	                        ndIssRoot.setAttribute "AppRefNo",iPONo
	                        ndIssRoot.setAttribute "AppRefDate",Pdate
	                        ndIssRoot.setAttribute "AppRefDate",Pdate
	                        ndIssRoot.setAttribute "ConsumptionAccHead",""
	                        ndIssRoot.setAttribute "IssueToCode",""
	                        ndIssRoot.setAttribute "PickPackFlag","N"
	                        ndIssRoot.setAttribute "IssFrom","PA"
	                        ndIssRoot.setAttribute "Returnable","N"
	                        ndIssRoot.setAttribute "ReturnItem","S"
	                        ndIssRoot.setAttribute "TYPE","SER"
    	                    
	                        exit for
	                    end if 'if ndPOChild.nodeName="Header" then
	                Next
	                For each ndPOChild in ndPORoot.ChildNodes
	                    if  ndPOChild.nodename="ITEMDETAILS" then
	                        sPOItemCode = ndPOChild.getAttribute("ITEMCODE")
	                        sItmRcptNum = GetItemRcptNum(sPOItemCOde)
	                        if trim(sItmRcptNum)="N" then
	                            sNoneCnt = sNoneCnt + 1
	                        elseif trim(sItmRcptNum)="LS" then
	                            sLSCnt = sLSCnt + 1
	                        elseif trim(sItmRcptNum)="S" then
	                            sSCnt = sSCnt + 1
	                        end if
    	                    
	                        set ndIssChild= IssueXML.createElement("ITEM")
	                        ndIssChild.setAttribute "ENTRYNO",ndPOChild.getAttribute("ENTRYNO")
	                        ndIssChild.setAttribute "ITMCODE",sPOItemCode
	                        ndIssChild.setAttribute "CLACODE",ndPOChild.getAttribute("CLASSCODE")
	                        ndIssChild.setAttribute "ITMNAME",ndPOChild.getAttribute("ITEMDESC")
	                        ndIssChild.setAttribute "SSTORE",""
	                        ndIssChild.setAttribute "REQQTY","0"
	                        ndIssChild.setAttribute "REQBY",""
	                        ndIssChild.setAttribute "REMARKS",""
	                        ndIssChild.setAttribute "ITEMTYPE",sitmType 
	                        ndIssChild.setAttribute "ISSUEDATE",dCreatedOn 
	                        ndIssChild.setAttribute "ISSQTY",ndPOChild.getAttribute("QTY")
	                        ndIssChild.setAttribute "TRAQTY","0"
	                        ndIssChild.setAttribute "PRQTY","0"
	                        ndIssChild.setAttribute "IVALUE","1"
	                        ndIssChild.setAttribute "ORGCODE",sOrgCode
	                        ndIssChild.setAttribute "MRSNO",""
	                        ndIssChild.setAttribute "MRSDATE",""
	                        ndIssChild.setAttribute "ATTRIBUTELIST",""
	                        ndIssChild.setAttribute "CREATEDBY",iCreatedBy
	                        ndIssChild.setAttribute "CREATEDON",dCreatedOn 
	                        ndIssChild.setAttribute "RETURNABLE","N"
	                        ndIssChild.setAttribute "RefNo",iPONo 
	                        ndIssChild.setAttribute "ONLYLOT",""
	                        ndIssChild.setAttribute "RETURNITEM","S"
	                        ndIssChild.setAttribute "SCProcess",""
						    ndIssChild.setAttribute "MatRecdAsItem",""
						    ndIssChild.setAttribute "MatRecdAsCode",""
						    ndIssChild.setAttribute "Instruct",""
						    ndIssChild.setAttribute "LabourCharge",""
						    ndIssChild.setAttribute "MatType",""
	                        ndIssRoot.appendChild ndIssChild 
            	            
	                        set ndIssPick = IssueXML.createElement("Pick")
	                        ndIssPick.setAttribute "TOT",ndPOChild.getAttribute("QTY")
	                        ndIssPick.setAttribute "NoofPack",""
	                        ndIssChild.appendChild ndIssPick 
    	                    
	                        sArrStoreInfo = split(GetStoreInfo(sPOItemCode),":")
                            sLocNo = sArrStoreInfo(0)
                            sBinNo = sArrStoreInfo(1)
    	                    
	                        if trim(sItmRcptNum)="LS" or trim(sItmRcptNum)="S" then
	                            set ndPickSubNode  = IssueXML.createElement("PICK")
	                                ndPickSubNode.setAttribute "LOC",sLocNo
	                                ndPickSubNode.setAttribute "BIN",sBinNo
	                                ndPickSubNode.setAttribute "LOTNO","N/A"
	                                ndPickSubNode.setAttribute "INVRECNO",""
	                                ndPickSubNode.setAttribute "QTYISS",ndPOChild.getAttribute("QTY")
	                                ndPickSubNode.setAttribute "Count","1"
	                                ndPickSubNode.setAttribute "NoofPack","0" 
	                            ndIssPick.appendChild ndPickSubNode
	                        else
	                            set ndPickSubNode  = IssueXML.createElement("STORE")
	                                ndPickSubNode.setAttribute "LOC",sLocNo
	                                ndPickSubNode.setAttribute "BIN",sBinNo
	                                ndPickSubNode.setAttribute "LOTNO","N/A"
	                                ndPickSubNode.setAttribute "INVRECNO",""
	                                ndPickSubNode.setAttribute "QTYISS",ndPOChild.getAttribute("QTY")
	                                ndPickSubNode.setAttribute "Count","1"
	                                ndPickSubNode.setAttribute "NoofPack","0" 
	                            ndIssPick.appendChild ndPickSubNode
	                        end if
    	                    
	                    end if ' if ndPOChild.nodeName="Header" then
	                Next
    	            
	                Response.write "<p>sNoneCnt="& sNoneCnt
	                Response.write "<p>sLSCnt="& sLSCnt
	                Response.write "<p>sSCnt="& sSCnt
    	            
	                if sLSCnt>0 or sSCnt >0 then
	                    ndIssRoot.setAttribute "ISSTYPE","M"
	                    ndIssRoot.setAttribute "PickPackFlag","L"
	                end if
            	    
	                IssueXML.save Server.MapPath("../temp/transaction/mrsIssueData"& Session.SessionID &".xml")
	                MrsIssueInsert
	            end if'if trim(dcrs(0))="Y" then
	    elseif trim(sOrdFor)="C" and trim(sAppRefType)<>"12" and sMRSEQNO = 2 then
	        sSql = "Select isNull(AutomaticMREntry,'N') from VW_APP_ApplicationSetup where ApplicationCode = 2 and ReferenceCodeNo = 4"
	        dcrs.Open sSql,con
	        if not dcrs.EOF then
	            sMREntry = trim(dcrs(0))
	        end if 
	        dcrs.Close()
    	    
	            if trim(sMREntry)="Y"  and (Trim(iClass)<>"0" and trim(iClass)<>"") then
	                set ndIssRoot = IssueXML.createElement("Root")
	                IssueXML.appendChild ndIssRoot 
   	                For each ndPOChild in ndPORoot.ChildNodes
	                    if ndPOChild.nodeName="HEADER" then
	                        sitmType = ndPOChild.getAttribute("ITEMTYPE")
	                        iCreatedBy = ndPOChild.getAttribute("CREATEDBY")
	                        dCreatedOn = ndPOChild.getAttribute("CREATEDON")
	                            set ndIssChild = IssueXML.createElement("HEADER")
	                            ndIssChild.setAttribute "FORUNIT",sOrgCode 
	                            ndIssChild.setAttribute "CREATEDON",dCreatedOn
	                            ndIssChild.setAttribute "TYPE","1"
	                            ndIssChild.setAttribute "USAGE","SUB"
	                            ndIssChild.setAttribute "REMARKS",""
	                            ndIssChild.setAttribute "APPROVER","IM"
	                            ndIssChild.setAttribute "CREATEDBY",Session("userid")
	                            ndIssChild.setAttribute "RECEIPTNO",""
	                            ndIssChild.setAttribute "COSTCENTER","select"
	                            ndIssChild.setAttribute "ITEMTYPE",sitmType 
	                            ndIssChild.setAttribute "ACCHEAD",""
	                            ndIssChild.setAttribute "REFTYPE",""
	                            ndIssChild.setAttribute "CallFrom","PO"
                                ndIssChild.setAttribute "RedirectTo","POList.asp"
                                ndIssChild.setAttribute "AppRefType","22"
                                ndIssChild.setAttribute "AppRefNo",iPONo
                                ndIssChild.setAttribute "AppRefDate",Pdate
                                ndIssChild.setAttribute "ImmediateApprover","Y"
                                ndIssChild.setAttribute "MRNo",""
                                ndIssChild.setAttribute "RequestedByUnit",""
                                ndIssChild.setAttribute "ISSTOTYPE","PARTY"
                                ndIssChild.setAttribute "ISSTOCODE",ndPOChild.getAttribute("SUPPAGENT")
                                ndIssChild.setAttribute "ISSTOSUBCODE",""
                                ndIssChild.setAttribute "ISSUETYPECODE","SUB"
	                            ndIssRoot.appendChild ndIssChild 
	                        exit for
	                    end if 'if ndPOChild.nodeName="Header" then
	                Next
	                For each ndPOChild in ndPORoot.ChildNodes
	                    if  ndPOChild.nodename="ITEMDETAILS" then
	                        set ndIssChild= IssueXML.createElement("ITEMDETAILS")
	                        ndIssChild.setAttribute "ENTRYNO",ndPOChild.getAttribute("ENTRYNO")
	                        ndIssChild.setAttribute "ITEMCODE",ndPOChild.getAttribute("ITEMCODE")
	                        ndIssChild.setAttribute "CLASSCODE",ndPOChild.getAttribute("CLASSCODE")
	                        ndIssChild.setAttribute "UNIT",sOrgCode
	                        ndIssChild.setAttribute "ITEMNAME",ndPOChild.getAttribute("ITEMDESC")
	                        ndIssChild.setAttribute "UOM",ndPOChild.getAttribute("UOM")
	                        ndIssChild.setAttribute "DECIMAL","Y"
	                        ndIssChild.setAttribute "DISPLAYED","Y"
	                        ndIssChild.setAttribute "QTY",ndPOChild.getAttribute("QTY")
	                        ndIssChild.setAttribute "REQUIREDBY","I"
	                        ndIssChild.setAttribute "REQUIREDVALUE",""
	                        ndIssChild.setAttribute "ATTRIBUTELIST",""
	                        ndIssChild.setAttribute "REMARKS",""
	                        ndIssChild.setAttribute "RefNo",iPONo
	                        ndIssRoot.appendChild ndIssChild 
            	            
	                        set ndIssPick = IssueXML.createElement("Schedule")
	                        ndIssPick.setAttribute "STYPE","I"
	                        ndIssPick.setAttribute "SVALUE",dCreatedOn 
	                        ndIssPick.setAttribute "ITEMCODE",ndPOChild.getAttribute("ITEMCODE")
	                        ndIssPick.setAttribute "CLASSCODE",ndPOChild.getAttribute("CLASSCODE")
	                        ndIssPick.setAttribute "SCHENTRYNO","1"
	                        ndIssChild.appendChild ndIssPick 
	                    end if ' if ndPOChild.nodeName="Header" then
	                Next
	                IssueXML.save Server.MapPath("../temp/transaction/MRS"& Session.SessionID &".xml")
    	            
                    sMRSNo =  MRInsert()	            
                    
	             end if ' if trim(dcrs(0))="Y" then
	    elseif trim(sOrdFor)="C" and trim(sAppRefType)<>"12" and sMRSEQNO = 0 then
	        sNoneCnt = 0
	        sLSCnt = 0
	        sSCnt = 0
    	    
	        sSql = "Select isNull(AutomaticIssueEntry,'N') from VW_APP_ApplicationSetup where ApplicationCode = 2 and ReferenceCodeNo = 4"
	        dcrs.Open sSql,con
	        if not dcrs.EOF then
	            sIssueEntry = trim(dcrs(0))
	        end if 
	        dcrs.Close
    	 
	       ' Response.Write " sIssueEntry = "& sIssueEntry 
	       ' Response.Write " sItmCode = "& sItmCode 
	       ' Response.Write " sClass = "& iClass 
	            if sIssueEntry="Y" and (Trim(iClass)<>"0" and trim(iClass)<>"")then
	                ''Issue Entry 
	                set ndIssRoot = IssueXML.createElement("ISSTYPE")
	                IssueXML.appendChild ndIssRoot
	                For each ndPOChild in ndPORoot.ChildNodes
	                    if ndPOChild.nodeName="HEADER" then
	                        sitmType = ndPOChild.getAttribute("ITEMTYPE")
	                        iCreatedBy = ndPOChild.getAttribute("CREATEDBY")
	                        sWithMat = ndPOChild.getAttribute("WITHMAT")
	                        dCreatedOn = ndPOChild.getAttribute("CREATEDON")
	                        ndIssRoot.setAttribute "ISSTYPE","F"
	                        ndIssRoot.setAttribute "ISSTOTYPE","Party"
	                        ndIssRoot.setAttribute "ISSTOCODE",ndPOChild.getAttribute("SUPPAGENT")
	                        ndIssRoot.setAttribute "ISSTOSUBCODE",""
	                        ndIssRoot.setAttribute "POConfirm","N"
	                        ndIssRoot.setAttribute "SInvConfirm","N"
	                        ndIssRoot.setAttribute "Invoice","A"
	                        ndIssRoot.setAttribute "GPConfirm","N"
	                        ndIssRoot.setAttribute "ProConfirm","N"
	                        ndIssRoot.setAttribute "MCallFrom","PO"
	                        ndIssRoot.setAttribute "RedirectTo","POList.asp"
	                        ndIssRoot.setAttribute "AppRefType","22"
	                        ndIssRoot.setAttribute "AppRefNo",iPONo
	                        ndIssRoot.setAttribute "AppRefDate",Pdate
	                        ndIssRoot.setAttribute "AppRefDate",Pdate
	                        ndIssRoot.setAttribute "ConsumptionAccHead",""
	                        ndIssRoot.setAttribute "IssueToCode",""
	                        ndIssRoot.setAttribute "PickPackFlag","N"
	                        ndIssRoot.setAttribute "IssFrom","PA"
	                        ndIssRoot.setAttribute "Returnable","N"
	                        ndIssRoot.setAttribute "ReturnItem","S"
	                        ndIssRoot.setAttribute "TYPE","SUB"
    	                    
	                        exit for
	                    end if 'if ndPOChild.nodeName="Header" then
	                Next
	                For each ndPOChild in ndPORoot.ChildNodes
	                    if  ndPOChild.nodename="ITEMDETAILS" then
	                        sPOItemCode = ndPOChild.getAttribute("ITEMCODE")
	                        sItmRcptNum = GetItemRcptNum(sPOItemCOde)
	                        if trim(sItmRcptNum)="N" then
	                            sNoneCnt = sNoneCnt + 1
	                        elseif trim(sItmRcptNum)="LS" then
	                            sLSCnt = sLSCnt + 1
	                        elseif trim(sItmRcptNum)="S" then
	                            sSCnt = sSCnt + 1
	                        end if
    	                    
	                        set ndIssChild= IssueXML.createElement("ITEM")
	                        ndIssChild.setAttribute "ENTRYNO",ndPOChild.getAttribute("ENTRYNO")
	                        ndIssChild.setAttribute "ITMCODE",sPOItemCode
	                        ndIssChild.setAttribute "CLACODE",ndPOChild.getAttribute("CLASSCODE")
	                        ndIssChild.setAttribute "ITMNAME",ndPOChild.getAttribute("ITEMDESC")
	                        ndIssChild.setAttribute "SSTORE",""
	                        ndIssChild.setAttribute "REQQTY","0"
	                        ndIssChild.setAttribute "REQBY",""
	                        ndIssChild.setAttribute "REMARKS",""
	                        ndIssChild.setAttribute "ITEMTYPE",sitmType 
	                        ndIssChild.setAttribute "ISSUEDATE",dCreatedOn 
	                        ndIssChild.setAttribute "ISSQTY",ndPOChild.getAttribute("QTY")
	                        ndIssChild.setAttribute "TRAQTY","0"
	                        ndIssChild.setAttribute "PRQTY","0"
	                        ndIssChild.setAttribute "IVALUE","1"
	                        ndIssChild.setAttribute "ORGCODE",sOrgCode
	                        ndIssChild.setAttribute "MRSNO",""
	                        ndIssChild.setAttribute "MRSDATE",""
	                        ndIssChild.setAttribute "ATTRIBUTELIST",""
	                        ndIssChild.setAttribute "CREATEDBY",iCreatedBy
	                        ndIssChild.setAttribute "CREATEDON",dCreatedOn 
	                        ndIssChild.setAttribute "RETURNABLE",ndPOChild.getAttribute("RETURNABLE")
	                        ndIssChild.setAttribute "RefNo",iPONo 
	                        ndIssChild.setAttribute "ONLYLOT",""
	                        ndIssChild.setAttribute "RETURNITEM",ndPOChild.getAttribute("RETURNITEM")
	                        ndIssChild.setAttribute "SCProcess",""
						    ndIssChild.setAttribute "MatRecdAsItem",""
						    ndIssChild.setAttribute "MatRecdAsCode",""
						    ndIssChild.setAttribute "Instruct",""
						    ndIssChild.setAttribute "LabourCharge",""
						    ndIssChild.setAttribute "MatType",""
	                        ndIssRoot.appendChild ndIssChild 
            	            
	                        set ndIssPick = IssueXML.createElement("Pick")
	                        ndIssPick.setAttribute "TOT",ndPOChild.getAttribute("QTY")
	                        ndIssPick.setAttribute "NoofPack",""
	                        ndIssChild.appendChild ndIssPick 
    	                    
	                        sArrStoreInfo = split(GetStoreInfo(sPOItemCode),":")
                            sLocNo = sArrStoreInfo(0)
                            sBinNo = sArrStoreInfo(1)
    	                    
	                        if trim(sItmRcptNum)="LS" or trim(sItmRcptNum)="S" then
	                            set ndPickSubNode  = IssueXML.createElement("PICK")
	                                ndPickSubNode.setAttribute "LOC",sLocNo
	                                ndPickSubNode.setAttribute "BIN",sBinNo
	                                ndPickSubNode.setAttribute "LOTNO","N/A"
	                                ndPickSubNode.setAttribute "INVRECNO",""
	                                ndPickSubNode.setAttribute "QTYISS",ndPOChild.getAttribute("QTY")
	                                ndPickSubNode.setAttribute "Count","1"
	                                ndPickSubNode.setAttribute "NoofPack","0" 
	                            ndIssPick.appendChild ndPickSubNode
	                        else
	                            set ndPickSubNode  = IssueXML.createElement("STORE")
	                                ndPickSubNode.setAttribute "LOC",sLocNo
	                                ndPickSubNode.setAttribute "BIN",sBinNo
	                                ndPickSubNode.setAttribute "LOTNO","N/A"
	                                ndPickSubNode.setAttribute "INVRECNO",""
	                                ndPickSubNode.setAttribute "QTYISS",ndPOChild.getAttribute("QTY")
	                                ndPickSubNode.setAttribute "Count","1"
	                                ndPickSubNode.setAttribute "NoofPack","0" 
	                            ndIssPick.appendChild ndPickSubNode
	                        end if
    	                    
	                    end if ' if ndPOChild.nodeName="Header" then
	                Next
	                
	                For each ndPOChild in ndPORoot.ChildNodes
	                    if ndPOChild.nodename="SubContract" then
	                        ndIssRoot.appendChild ndPOChild
	                    end if
	                Next
    	            
	                Response.write "<p>sNoneCnt="& sNoneCnt
	                Response.write "<p>sLSCnt="& sLSCnt
	                Response.write "<p>sSCnt="& sSCnt
    	            
	                if sLSCnt>0 or sSCnt >0 then
	                    ndIssRoot.setAttribute "ISSTYPE","M"
	                    ndIssRoot.setAttribute "PickPackFlag","L"
	                end if
            	    
	                IssueXML.save Server.MapPath("../temp/transaction/mrsIssueData"& Session.SessionID &".xml")
	                MrsIssueInsert
	            end if'if trim(dcrs(0))="Y" then
	    end if ' if trim(sOrdFor)="S" then
	end if 'if trim(sSchRelReq)="N" then
End Function 'Function POInsert()
%>
