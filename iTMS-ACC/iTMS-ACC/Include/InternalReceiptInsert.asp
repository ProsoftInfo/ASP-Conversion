<!--#include File="NoSeries.asp" -->
<!--#include File="PackingCodeSeries.asp" -->
<!--#include File="NoSeriesCommonFunctions.asp" -->
<%
	'XML DOM Variables
	Dim oDOM1

	' Create our DOM Document Objects
	Set oDOM1 = Server.CreateObject("Microsoft.XMLDOM")
	
	dim newxml,ndRoot,ndHeader,LotNode,SerialNode
	dim newxml1,ndRoot1,ndHeader1,LotNode1,SerialNode1
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")
	Set newxml1 = Server.CreateObject("Microsoft.XMLDOM")
	
	dim rsObj,rsObj1,sSQLQuery,iAppCode,iDeptLineNo,rsTempRcpt
	dim iMRSNumber,iICode,sClassCode,sBinNo,sLocCode,iQty,iLot,iSerialNo,dIssDate
	dim sOrgCode,iCreateby,sDept,sSrc,iIssNo,iIntRecNo,iTempMRSNo,iQtyRet,iQtySerial
	dim sSourceRefType, sSourceRefNo,iDINo,sRemark,sTypeRcpt,sTraType,sItemType,sPackWho
	Dim sExpression, iQtyIssued, iQtyReturned, iQtyReturn, iTotReturn
	Dim iIssueEntryNo, WorkCenterNode1, i,WCODE,MachineCenterNode1, j, MCODE
	Dim AutoGen, TypeofStock
	Dim sFinPeriod, sFinPeriodFrom, sFinPeriodTo, sFinFromDt, sFinToDt,sTempMonthYear,arrFinPeriod,sMonthYr
	Dim dtCurrDate
	Dim ndDet,ndItemDet,sSubLevelID,sNoofCone,sAutoAccountingPack,sTempSeriesRcpt,sArrSeriesRcpt,sNumClassNameRcpt
	Dim sAppRefTypeRcpt,sAppRefnoRcpt,sAppRefDateRcpt,sReceviedOn,sByProduct,sUseExistingPackNum,sAutoAccount
	dim iInvRecNo,iStockQuality,sRcptNum
	Dim ndIntStorage,ndIntLotSer,ndIntLotSerDet,sIntLotNo,sIntPackNum,iIntSerNo,iIntPackCode,sIntAttList
	Dim sIntStkQty,iIntStore,iIntBin,iIntRate,iIntStoreVal,iIntStoreQty,iIntQtyTare,iIntQtyNett,sIntPackStatus,sInvStatus
	Dim nExistItemQty,nExistItemVal,iRcptSerialEntry,sCallFrom
	
Function CreateInternalReceipt(CurrDate)
	
	dtCurrDate = CurrDate
	
	sFinPeriod = Session("FinPeriod")
	sFinFromDt = "01/04/" & Mid(sFinPeriod,1,4)
	sFinToDt = "31/03/" & Mid(sFinPeriod,6,4)
	
    newxml1.Load server.MapPath("../temp/transaction/IssRet"&Session.SessionID&".xml")
	newxml.async = false
    newxml.load Server.MapPath("../temp/transaction/ReceiptLotData"&Session.SessionID&".xml")'	
    
	iCreateby = getUserid
	iAppCode = "6"
	AutoGen = ""
	
	Set rsObj = Server.CreateObject("ADODB.RecordSet")
	Set rsObj1 = Server.CreateObject("ADODB.RecordSet")
	set rsTempRcpt = Server.CreateObject("ADODB.Recordset")

	Set ndRoot = newxml.documentElement
	Set ndRoot1 = newxml1.documentElement	
	sDept = trim(ndRoot.Attributes.getNamedItem("DEPT").Value)
	sSrc = trim(ndRoot.Attributes.getNamedItem("SOURCE").Value)
	sOrgCode = trim(ndRoot.Attributes.getNamedItem("ORGCODE").Value)
	'sTypeRcpt = trim(ndRoot.Attributes.getNamedItem("sTypeRcpt").Value)
	'sItemType = trim(ndRoot.Attributes.getNamedItem("ITEMTYPE").Value)
	'sPackWho = trim(ndRoot.Attributes.getNamedItem("PACKNUM").Value)
	sSourceRefType = trim(ndRoot.Attributes.getNamedItem("SRCREFTYPE").Value)
	sSourceRefNo = trim(ndRoot.Attributes.getNamedItem("SRCREFNO").Value)	
	TypeofStock =  trim(ndRoot.Attributes.getNamedItem("RCPTNUMBERINV").Value)	
	sAppRefTypeRcpt = trim(ndRoot.Attributes.getNamedItem("APPREFTYPE").value)
	sAppRefnoRcpt = trim(ndRoot.Attributes.getNamedItem("APPREFNO").value)
	sAppRefDateRcpt = trim(ndRoot.Attributes.getNamedItem("APPREFDATE").value)
	sReceviedOn = trim(ndRoot.Attributes.getNamedItem("RCVDON").value)
	sAutoAccount = trim(ndRoot.Attributes.getNamedItem("AUTOACCOUNT").value)
	
	if trim(sAppRefTypeRcpt)="" or isNull(sAppRefTypeRcpt) or trim(sAppRefTypeRcpt)="N" then sAppRefTypeRcpt = 0
	if trim(sAppRefnoRcpt)="" or IsNull(sAppRefTypeRcpt) then sAppRefnoRcpt = "NULL"
	if trim(sAppRefnoRcpt)<>"NULL"  then pack(sAppRefnoRcpt)
	if trim(sAppRefDateRcpt)="" or isNull(sAppRefDateRcpt) then sAppRefDateRcpt = "NULL"
	if trim(sAppRefDateRcpt)<>"NULL" then sAppRefDateRcpt = pack(sAppRefDateRcpt)
	
	if trim(sUseExistingPackNum)="" or IsNull(sUseExistingPackNum) then sUseExistingPackNum="N"
	
	Response.write "<p>sUseExistingPackNum="& sUseExistingPackNum
	
	if sTypeRcpt = "T" then
		sTraType = "'RB'"
	else
		sTraType = "'RR'"
	end if
	IF sTypeRcpt = "P" Then sTraType = "'RF'"
	
	if ndRoot.HaschildNodes() then
		with rsObj
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(INTERNALRECEIPTNO)+1,1) FROM APP_T_INTERNALRECEIPTHEADER"
			.ActiveConnection = con
			.Open
		end with
		'Response.Write "<p>"& rsObj.Source 
		set rsObj.ActiveConnection = nothing

		if not rsObj.EOF then
			iIntRecNo = trim(rsObj(0))
		end if
		rsObj.Close
		
		sSQLQuery  = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
		rsObj.Open sSQLQuery,con	
		if not rsObj.EOF then
			iInvRecNo = rsObj(0)
		end if
		rsObj.Close
		
		
		dIssDate = dtCurrDate
'		Response.Write " dISSDATE="& dtCurrDate
		if sSrc = "N" then
			sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTHEADER (INTERNALRECEIPTNO,APPLICATIONCODE," &_
				"ORGANISATIONCODE,CREATEDFROMDEPT,REFTYPE,CREATEDON,CREATEDBY,STATUS,USAGETYPE,USAGESRCREFNO,USAGESRCREFTYPE,APPREFTYPE,APPREFNO,APPREFDATE,RECEIVEDON) VALUES " &_
				"(" & cDbl(iIntRecNo) & "," & iAppCode & "," & Pack(sOrgCode) & "," & Pack(sDept) & "," & Pack(sSrc) & "," &_
				"CONVERT(DATETIME," & Pack(dtCurrDate) & ",103)," & iCreateby & ",'N','" & sTypeRcpt & "','" & sSourceRefNo & "','" & sSourceRefType & "',"& sAppRefTypeRcpt &","&sAppRefnoRcpt&","& sAppRefDateRcpt &",Convert(datetime,"& pack(sReceviedOn) &",103))"
			Response.Write "<p>"& sSQLQuery
			con.Execute sSQLQuery
			'Response.Write ndRoot.XML
			For Each ndDet In ndRoot.childNodes
			    if ndDet.nodeName="Details" then
			        for each ndItemDet in ndDet.childNodes
			            if ndItemDet.nodeName="ItemDetail" then
			               ' For Each LotNode in ndHeader.ChildNodes
					       '     dIssDate = trim(LotNode.Attributes.Item(1).nodeValue)
				           ' Next
				            'Response.Write dIssDate
				            sClassCode = trim(ndItemDet.Attributes.getNamedItem("CLACODE").Value)
				            iICode = trim(ndItemDet.Attributes.getNamedItem("ItemCode").Value)
				            iQty = trim(ndItemDet.Attributes.getNamedItem("QTY").Value)
				            iTempMRSNo = trim(ndItemDet.Attributes.getNamedItem("MRSNO").Value)
				            iIssNo = trim(ndItemDet.Attributes.getNamedItem("ISSNO").Value)
				            sByProduct = trim(ndItemDet.Attributes.getNamedItem("BYPRODUCT").value)
				            sRcptNum = trim(ndItemDet.Attributes.getNamedItem("RECEIPTNUM").value)

				            If iIssNo <> "N" Then
					            with rsObj
						            .CursorLocation = 3
						            .CursorType = 3
						            .Source = "SELECT IssueEntryNo FROM INV_T_DEPARTMENTSTOCK WHERE ISSUEENTRYNO = '" & iIssNo & "'"
						            .ActiveConnection = con
						            .Open
					            end with
        					
					            set rsObj.ActiveConnection = nothing

					            if not rsObj.EOF then
						            iIssueEntryNo = trim(rsObj(0))
					            end if
					            rsObj.Close
				            End If
				            if iTempMRSNo = "N" then
					            iMRSNumber = "NULL"
				            else
					            iMRSNumber = iTempMRSNo
				            end if
				            
				            if trim(sRcptNum)="N" then
				                 for each ndIntStorage in ndItemDet.childNodes
				                    if trim(ndIntStorage.nodeName)="STORAGE" then
				                        iIntStore = ndIntStorage.getAttribute("STORE")
				                        iIntBin = ndIntStorage.getAttribute("BIN")
				                        iIntStoreVal = ndIntStorage.getAttribute("STORAGEVALUE")
				                        iIntStoreQty = ndIntStorage.getAttribute("QTY")
				                        if cdbl(iIntStoreVal)>0 then
				                            iIntRate = cdbl(iIntStoreVal)/cdbl(iIntStoreQty)
				                        end if
				                        sIntStkQty = ndIntStorage.getAttribute("SQ")
				                        if Trim(sIntStkQty)="" or IsNull(sIntStkQty) then sIntStkQty="0"
				                        if trim(iIntRate) = "" or IsNull(iIntRate) then iIntRate = "0"
				                        
				                        sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
					                        "CLASSIFICATIONCODE,ITEMCODE,QUANTITYRETURN,PRODUCTTYPE,StockQuality,STORE,BIN,RATE) VALUES " &_
					                        "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," &_
					                        "" & iQty & ","& pack(sByProduct) &","& sIntStkQty &","& iIntStore &","& iIntBin &","& iIntRate &")"
				                        Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                        con.Execute sSQLQuery
				                    end if 'if trim(ndIntStorage.nodeName)="STORAGE" then
				                next
				            else
				                for each ndIntStorage in ndItemDet.childNodes
				                    if trim(ndIntStorage.nodeName)="STORAGE" then
				                        iIntStore = ndIntStorage.getAttribute("STORE")
				                        iIntBin = ndIntStorage.getAttribute("BIN")
				                        iIntStoreVal = ndIntStorage.getAttribute("STORAGEVALUE")
				                        iIntStoreQty = ndIntStorage.getAttribute("QTY")
				                        if cdbl(iIntStoreVal)>0 and iIntStoreQty>0 then
				                            iIntRate = cdbl(iIntStoreVal)/cdbl(iIntStoreQty)
				                        else
				                            iIntRate = 0
				                        end if
				                        
				                        for each ndIntLotSer in ndIntStorage.childNodes
				                            if trim(ndIntLotSer.nodeName)="LotSerial" then
				                                for each ndIntLotSerDet in ndIntLotSer.childNodes
				                                    iQtyReturn = ndIntLotSerDet.getAttribute("QTYREC")
				                                    sIntLotNo = ndIntLotSerDet.getAttribute("LOT")
				                                    sIntPackNum = ndIntLotSerDet.getAttribute("PACKNUMBER")
				                                    iIntSerNo = ndIntLotSerDet.getAttribute("LOTSERIAL")
				                                    iIntPackCode =  ndIntLotSerDet.getAttribute("PACKINGTYPE")
				                                    sIntAttList = ndIntLotSerDet.getAttribute("ATTRIBUTELIST")
				                                    sIntStkQty = ndIntLotSerDet.getAttribute("SQ")
				                                    iIntQtyTare = ndIntLotSerDet.getAttribute("TAREREC")
				                                    
				                                    iIntQtyNett = iQtyReturn-iIntQtyTare
				                                    'Response.Write "<P>sIntPackNum="& sIntPackNum 
				                                    
				                                    if trim(sIntLotNo)="" or IsNull(sIntLotNo) then sIntLotNo="NULL"
				                                    if trim(sIntLotNo)<>"NULL" then sIntLotNo=Pack(sIntLotNo)
				                                    if trim(sIntAttList)="" or IsNull(sIntAttList) then sIntAttList="NULL"
				                                    if trim(sIntAttList)<>"NULL" then sIntAttList= pack(sIntAttList)
				                                    if trim(iIssNo)="N" or trim(iIssNo)="" or IsNull(iIssNo) then iIssNo = "NULL"
				                                    
				                                    if trim(sIntStkQty)="" or IsNull(sIntStkQty) then sIntStkQty ="NULL"
				                                    
				                                    sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
									                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYRETURN,PackingNum,AttributeList,ProductType,StockQuality,STORE,BIN,RATE,PackingCode,GrossQuantityReturn) VALUES " &_
									                "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
									                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & sIntLotNo & "," & iIntSerNo & "," & iIntQtyNett & "," & sIntPackNum & ","& sIntAttList &","& pack(sByProduct) &","& sIntStkQty &","& iIntStore &","& iIntBin &","& iIntRate &","& iIntPackCode &","& iQtyReturn &")"
									                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                                    con.Execute sSQLQuery
				                                next
				                            end if
				                        next
				                    end if
				                next
				            end if
				            if trim(sAutoAccount)="Y" then
				                InsertLotDetails(iICode)
				            end if 
				        end if 'if ndItemDet.nodeName="ItemDetail" then
				    next
				end if 'if ndDet.nodeName="Details" then
			next
			'Calling Function to insert Lot/Serial Details
			
			
			
			
		else 'MRS / Direct Issue Based
		
			sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTHEADER (INTERNALRECEIPTNO,APPLICATIONCODE," &_
				"ORGANISATIONCODE,CREATEDFROMDEPT,REFTYPE,CREATEDON,CREATEDBY,STATUS,RECEIVEDON) VALUES " &_
				"(" & iIntRecNo & "," & iAppCode & "," & Pack(sOrgCode) & "," & Pack(sDept) & "," & Pack(sSrc) & "," &_
				"CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iCreateby & ",'N',Convert(datetime,"& pack(sReceivedOn) &",103))"
			Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
			con.Execute sSQLQuery
			
			For Each ndDet In ndRoot.childNodes
			    if ndDet.nodeName="Details" then
			        for each ndItemDet in ndDet.childNodes
			            if ndItemDet.nodeName="ItemDetail" then
		    	                sClassCode = trim(ndItemDet.Attributes.getNamedItem("CLACODE").Value)
				                iICode = trim(ndItemDet.Attributes.getNamedItem("ItemCode").Value)
				                iQty = trim(ndItemDet.Attributes.getNamedItem("QTY").Value)
				                iMRSNumber = trim(ndItemDet.Attributes.getNamedItem("MRSNO").Value)
				                iDINo = trim(ndItemDet.Attributes.getNamedItem("DINO").Value)
				                iIssNo = trim(ndItemDet.Attributes.getNamedItem("ISSNO").Value)
				                sRemark = trim(ndItemDet.Attributes.getNamedItem("REMARKS").Value)
				                with rsObj
					                .CursorLocation = 3
					                .CursorType = 3
					                .Source = "SELECT IssueEntryNo FROM INV_T_DEPARTMENTSTOCK WHERE ISSUENO = '" & iIssNo & "'"
					                .ActiveConnection = con
					                .Open
				                end with
				                'Response.Write rsObj.Source 
				                set rsObj.ActiveConnection = nothing

				                if not rsObj.EOF then
					                iIssueEntryNo = trim(rsObj(0))
				                end if
				                rsObj.Close
				                'Response.Write iIssueEntryNo 
				                 ' This is added for the purpose of to add WorkCenterWise details
				                sExpression ="//ItemDetail [ @CLACODE = "&sClassCode&" and @ItemCode = "&iICode&"]/AddDet/WorkCenter"
				                Set WorkCenterNode1 = ndRoot1.Selectnodes(sExpression)
				                'Response.Write WorkCenterNode1.Length
				                If WorkCenterNode1.Length > 0 Then
				                For i = 0 to WorkCenterNode1.Length-1
					                WCODE = trim(WorkCenterNode1.Item(i).Attributes.getNamedItem("WCODE").Value)		
					                'Response.Write WCODE					
					                sExpression ="//ItemDetail [ @CLACODE = "&sClassCode&" and @ItemCode = "&iICode&"]/AddDet/WorkCenter [ @WCODE = '"&WCODE&"']/MachineCenter"
					                Set MachineCenterNode1 = ndRoot1.Selectnodes(sExpression) 
					                'Response.Write MachineCenterNode1.Length & sExpression
					                For j = 0 to MachineCenterNode1.Length-1
						                MCODE =  trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("MCODE").Value)							
						                iQtyIssued = trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("QTY").Value)												
						                iQtyReturned = trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("QTYRETURNED").Value)												
						                iQtyReturn = trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("QTYRETURN").Value)																	
						                iTotReturn = iQtyReturned + cDbl(iQtyReturn)																		
            								
						                If cDbl(iQtyReturn) > 0 Then
							                sSQLQuery = "INSERT INTO APP_T_InternalReceiptAddnDet (InternalReceiptNo,ClassificationCode,ItemCode,WorkCenterCode,MachineCenterCode,QuantityReturn) VALUES " &_
							                "(" & iIntRecNo & "," & sClassCode & ", " & iICode & ",'" & WCODE & "', '" & MCODE & "'," & iQtyReturn & ")"
							                Response.Write "<p>"& sSQLQuery
							                con.Execute sSQLQuery	
            									
							                sSQLQuery = "UPDATE inv_t_issuedeptwiseBreakUp SET QUANTITYRETURNED = " &_
								                "(ISNULL(QUANTITYRETURNED,0) + " & iQtyReturn & ") WHERE ISSUEENTRYNO = " & iIssueEntryNo & " AND " &_
								                " ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
								                "CLASSIFICATIONCODE = " & sClassCode & " AND ITEMCODE = " & iICode & " AND WORKCENTERCODE = '" & WCODE & "' AND MACHINECENTERCODE = '" & MCODE & "'"									
							                Response.Write "<p>"& sSQLQuery
							                con.Execute sSQLQuery																			
						                End If
					                Next
				                Next 			
            					
				                End If

		                '------------------------------------------------------------------------------------
            					
            					
            					
            					
            					
            					
            					
				                if Trim(iMRSNumber) = "" then iMRSNumber = "NULL"
				                if Trim(iDINo) = "" then iDINo = "NULL"
				                if sRemark = "" then
					                sRemark = "NULL"
				                else
					                sRemark = Pack(sRemark)
				                end if

				                if Trim(iMRSNumber) <> "NULL" then
					                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCK SET QUANTITYRETURNED = " &_
						                "(ISNULL(QUANTITYRETURNED,0) + " & iQty & ") WHERE ISSUENO = " & iIssNo & " AND " &_
						                "MRSNUMBER = " & iMRSNumber & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
						                "CLASSIFICATIONCODE = " & sClassCode & " AND ITEMCODE = " & iICode & ""
				                else
					                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCK SET QUANTITYRETURNED = " &_
						                "(ISNULL(QUANTITYRETURNED,0) + " & iQty & ") WHERE ISSUENO = " & iIssNo & " AND " &_
						                "DINUMBER = " & iDINo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
						                "CLASSIFICATIONCODE = " & sClassCode & " AND ITEMCODE = " & iICode & ""
				                end if
            										
				                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                con.Execute sSQLQuery

				                For Each LotNode in ndItemDet.ChildNodes
					                iLot = trim(LotNode.Attributes.Item(0).nodeValue)
					                dIssDate = trim(LotNode.Attributes.Item(1).nodeValue)
					                iQtyRet = trim(LotNode.Attributes.Item(3).nodeValue)
					                'iDeptLineNo = trim(LotNode.Attributes.Item(4).nodeValue)



					                if iLot = "0" or iLot = "" or iLot = "NULL" then
						                iLot = "NULL"
					                else
						                iLot = Pack(iLot)
					                end if

					                if not iQtyRet = "0" then
						                if LotNode.hasChildNodes then
							                For Each SerialNode in LotNode.ChildNodes
								                iSerialNo = trim(SerialNode.Attributes.Item(0).nodeValue)
								                iQtySerial = trim(SerialNode.Attributes.Item(1).nodeValue)

								                if not iQtySerial = "0" then
									                if Trim(iMRSNumber) <> "NULL" then
										                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
											                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYRETURN,REMARKS) VALUES " &_
											                "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
											                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iLot & "," & iSerialNo & "," & iQtySerial & "," & sRemark & ")"
									                else
										                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,DINUMBER," &_
											                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYRETURN,REMARKS) VALUES " &_
											                "(" & iIntRecNo & "," & iDINo & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
											                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iLot & "," & iSerialNo & "," & iQtySerial & "," & sRemark & ")"
									                end if										
									                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
									                con.Execute sSQLQuery

									                if iLot = "NULL" then
										                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
											                "(ISNULL(QUANTITYRETURNED,0) + " & iQtySerial & ") WHERE ISSUENO = " & iIssNo & " AND " &_
											                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
											                "LOTNO IS NULL AND SERIALNO = " & iSerialNo '& " AND LINENUMBER = " & iDeptLineNo & ""
									                else
										                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
											                "(ISNULL(QUANTITYRETURNED,0) + " & iQtySerial & ") WHERE ISSUENO = " & iIssNo & " AND " &_
											                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
											                "LOTNO = " & iLot & " AND SERIALNO = " & iSerialNo '& " AND LINENUMBER = " & iDeptLineNo & ""
									                end if
									                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
									                con.Execute sSQLQuery
								                end if
							                next
						                else
							                if Trim(iMRSNumber) <> "NULL" then
								                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
									                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,QUANTITYRETURN,REMARKS) VALUES " &_
									                "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
									                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iQtyRet & "," & sRemark  & ")"
							                else
								                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,DINUMBER," &_
									                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,QUANTITYRETURN,REMARKS) VALUES " &_
									                "(" & iIntRecNo & "," & iDINo & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
									                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iQtyRet & "," & sRemark  & ")"
							                end if								
							                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
							                con.Execute sSQLQuery

							                if iLot = "NULL" then
								                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
									                "(ISNULL(QUANTITYRETURNED,0) + " & iQtyRet & ") WHERE ISSUENO = " & iIssNo & " AND " &_
									                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
									                "LOTNO IS NULL AND (SERIALNO IS NULL OR SERIALNO = 0)" ' AND LINENUMBER = " & iDeptLineNo & ""
							                else
								                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
									                "(ISNULL(QUANTITYRETURNED,0) + " & iQtyRet & ") WHERE ISSUENO = " & iIssNo & " AND " &_
									                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
									                "LOTNO = " & iLot & " AND (SERIALNO IS NULL OR SERIALNO = 0)" ' AND LINENUMBER = " & iDeptLineNo & ""
							                end if
							                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
							                con.Execute sSQLQuery
						                end if
					                end if
				                next
				        ' end if for Item Node check
				        end if
				    next
				end if
			next
		end if
	end if
End Function
%>

<%	'Function to insert the Lot/Serial Details for None Cases
	Function InsertLotDetails(sSendItemCode)
		dim ItemNode,StorageNode,LotNode,LotSerialNode,objfs,FabNode,FabDetNode,OptionalUoMNode
		dim iNoofPacks, PackNode
		dim sTemp,iCtr,sLocCode,sBinNo,sMonthYr,iQty,iValue,sAppli
		dim arrFinPeriod,arrTemp,iCounter,sTempMonthYear
		dim sQtyIn,iTareQty,iLot,iSerialFrom,iSerialTo,sSellingType,iWeight,sPackingType,sPackingNumber
		dim iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sTareIn,iQtyGross,iQtyNett
		dim iSeriesNo,iSeriesCode,sProductCode,iItmRate,sItmShDesc,sItmAddDesc
		dim iLotCtr,sSellingFormType,sStage,iCodeLen,iCodeSize,sForm
		dim sDrwVerNo,sCatalogue,iCommodity,iLotQty,iLotValue,iEntryCtr,sReceivedOn
		dim iFabCtr,iPieceNo,iPieceQty,arrFab,iAltGross,iAltNett,sAltUoM,sCheck,iRate
		dim sOpUoMCode,iOpUoMFactor,sOpUoMOperator,sRecNum,iStorageQty
		dim objFSO,objTxt,sExpression,sExpression1,iDefinedBy,arrDate
		Dim iSNo, iSCode, sPackCode ,ItemNodeDet,sLotAttribute
		
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Transaction/ReceiptLotData.txt"))

		iDefinedBy = getUserid
		
		iNoofPacks = "Null"
		
		with rsObj
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(COMPANYITEMCODE,'') FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iICode & ""
			Response.Write rsObj.Source 
			.ActiveConnection = con
			.Open
		end with
		set rsObj.ActiveConnection = nothing
		if not rsObj.EOF then
			sProductCode = trim(rsObj(0))
		else
			sProductCode = ""
		end if
		rsObj.Close

		with rsObj
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT RECEIPTNUMBERING FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
			.ActiveConnection = con
			.Open
		end with
		set rsObj.ActiveConnection = nothing
		if not rsObj.EOF then
			sRecNum = trim(rsObj(0))
		end if
		rsObj.Close
		
		' Storage Location Details Insert
		
		sExpression = "//ROOT/Details/ItemDetail[@ItemCode="&sSendItemCode&"]"
		set ItemNodeDet = ndRoot.Selectnodes(sExpression)
		if ItemNodeDet.length>0 then
		    TypeOfStock = ItemNodeDet.Item(0).Attributes.getNamedItem("RECEIPTNUM").value
		end if
		
		'This Part Added For the purpose of to add No of Packs in Ledger Entries
		sExpression ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE/LotSerial/LotSerialDetails"
		Set PackNode = ndRoot.Selectnodes(sExpression)
		If PackNode.Length > 0 Then
			iNoofPacks = PackNode.Length
		End If
		
		'Response.Write "<p>No of Packs = "& iNoofPacks
		
		
			
		
		iCtr = 0
		sExpression ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE"
		Set StorageNode = ndRoot.Selectnodes(sExpression)
		
		for iCtr =  0 to StorageNode.length - 1
		
			sLocCode = StorageNode.Item(iCtr).Attributes.getNamedItem("STORE").Value
			sBinNo = StorageNode.Item(iCtr).Attributes.getNamedItem("BIN").Value
			sAppli = StorageNode.Item(iCtr).Attributes.getNamedItem("APPLICABLE").Value
			sMonthYr = StorageNode.Item(iCtr).Attributes.getNamedItem("MONTHYEAR").Value
			iStorageQty = StorageNode.Item(iCtr).Attributes.getNamedItem("QTY").Value
			iValue = Round(Cdbl(StorageNode.Item(iCtr).Attributes.getNamedItem("STORAGEVALUE").Value),2)
			 Response.Write " iValue = "& iValue
Response.Write "<p>iQty = "& iQty 
			'sReceivedOn = "01/"&left(sMonthYr,2)&"/"&right(sMonthYr,4)
			sReceivedOn = sMonthYr
			if Trim(sReceivedOn)="" or IsNull(sReceivedOn) then sReceivedOn =  FormatDate(date)
			arrDate = split(sReceivedOn,"/")
			Response.Write "<p>sReceivedOn = "& sReceivedOn 
			sMonthYr = arrDate(1)&arrDate(2)
			'Response.Write sMonthYr
			dIssDate = sReceivedOn 
			if sBinNo = "0" then sBinNo = "NULL"
			
			
			
			
			
			if sRecNum = "N" then
			    if sBinNo = "0" then sBinNo = "NULL"
			    if trim(iStockQuality)="" or IsNull(iStockQuality) then iStockQuality = "0"
			    
			    sSQLQuery = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
							"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTQUANTITYGROSS," &_
							"LOTQUANTITYNETT,DateOfReceipt,SrcType,LotNumber,SerialNumber,StockQuality) VALUES " &_
							"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iICode & "," & sClassCode& "," &_
							"" & sLocCode & "," & sBinNo & "," & iStorageQty & "," & iStorageQty & ",Convert(datetime,'"& dIssDate &"',103),"& sTraType &",NULL,NULL,"& iStockQuality &")"
						Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
						con.Execute sSQLQuery
			end if
			
			'if sBinNo = "NULL" or IsNull(sBinNo) then  sBinNo = 0 

			sExpression1 ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE [ @STORE = """&sLocCode&""" and @BIN = """&sBinNo&"""]/LotSerial"
			Response.Write "<p>"&sExpression1 
			Set LotNode = ndRoot.Selectnodes(sExpression1)
			Response.Write "<p>LotNode.Length="& LotNode.length

			if LotNode.Length > 0 then
			
				For iLotCtr = 0 to LotNode.Length - 1

					iCounter = 0

					sQtyIn = LotNode.Item(iLotCtr).Attributes.getNamedItem("QTYIN").Value
					iTareQty = LotNode.Item(iLotCtr).Attributes.getNamedItem("TARE").Value
					iLot = LotNode.Item(iLotCtr).Attributes.getNamedItem("LOT").Value
					iSerialFrom = LotNode.Item(iLotCtr).Attributes.getNamedItem("SERIALFROM").Value
					iSerialTo = LotNode.Item(iLotCtr).Attributes.getNamedItem("SERIALTO").Value
					sTareIn = LotNode.Item(iLotCtr).Attributes.getNamedItem("TAREWEIGHT").Value
					iLotQty = LotNode.Item(iLotCtr).Attributes.getNamedItem("QTY").Value
					iEntryCtr = LotNode.Item(iLotCtr).Attributes.getNamedItem("COUNTER").Value
					sStage = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("STAGE").Value)
					AutoGen = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("AUTOGEN").Value)
					'Response.Write iSerialFrom & "<" & iSerialTo 

					
					iAltGross = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTGROSS").Value)
					iAltNett = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTNETT").Value)
					sAltUoM = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTUOM").Value)
					'iLotValue = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("IVALUE").Value)
					
					if trim(iLotValue) = ""  or IsNull(iLotValue) then iLotValue = "0"
					
					if iLotValue>0 then									
					    iRate = FormatNumber(cdbl(iLotValue) / cdbl(iLotQty),4,,,0)
					else
					    iRate = 0
					end if
					
					if sStage = "select" then
						sStage = "NULL"
					else
						sStage = Pack(sStage)
					end if
					
					if iTareQty > 0 then
					    iQty = iLotQty - iTareQty
					else
					    iQty = iLotQty
					end if
				
					'Response.Write "iQty = "& iQty 
					
					if iTareQty = "" then iTareQty = "0"
					if iAltGross = "" then iAltGross = "0"
					if iAltNett = "" then iAltNett = "0"
					if sAltUoM = "" or lcase(sAltUoM) = "select" then
						sAltUoM = "NULL"
					else
						sAltUoM = Pack(sAltUoM)
					end if
					
					sSQLQuery = "Select IsNull(AutomaticAccounting,'N') from APP_M_ApplicationSetup where ApplicationCode = 6 and ReferenceCodeNo = 4"
						
                        rsObj.Open sSQLQuery,con
                        if not rsObj.EOF then
                            sAutoAccountingPack = rsObj(0)
                        else
                            sAutoAccountingPack = "N"
                        end if
                        rsObj.Close 


			'------------------------------------
					Response.Write "AutoGen = "&  AutoGen
					Response.write "TypeofStock="& TypeOfStock
					''added by ragav on aug 13,2013 for MovItem,Phyadj,MergeItem
					'If AutoGen = "AUTO" and (TypeofStock = "LS" or TypeofStock = "L") Then
					if trim(sCallFrom)="MOV" or Trim(sCallFrom)="MERGE" or Trim(sCallFrom)="PA" then
					    if Trim(AutoGen)="AUTO" then
					        sPackWho = "N"
					    else
					        sPackWho = ""
					    end if 
					else
					    sPackWho = ""
					end if 
					
					If AutoGen = "AUTO" and (TypeofStock = "LS" or TypeofStock = "L") and sAutoAccountingPack="N" and sUseExistingPackNum="N" Then
					
						with rsObj
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ITEMTYPEID FROM VWITEM WHERE ORGANISATIONCODE = '" & sOrgCode & "' AND ITEMCODE = '" & iICode & "'"
							.ActiveConnection = con
							.Open
						end with
						'Response.Write rsObj.Source 
						set rsObj.ActiveConnection = nothing	
						If Not rsObj.EOF Then
							sItemType = rsObj(0)
						End If
						rsObj.Close 
					'	with rsObj
					'		.CursorLocation = 3
					'		.CursorType = 3
					'		.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'LO' AND ORGANISATIONCODE = '" & sOrgCode & "'"
					'		.ActiveConnection = con
					'		.Open
					'	end with
					'	set rsObj.ActiveConnection = nothing
					'	'Response.Write rsObj.Source 
					'	if not rsObj.EOF then
					'		iSNo = trim(rsObj(0))
					'		iSCode = trim(rsObj(1))
					'		'Response.Write sOrgCode& "," & iSNo& ","& iSCode & "," &sFinFromDt 
					'        iLotEntry = GenSeriesNumber(sOrgCode,iSNo,iSCode,sFinFromDt)	
					'		iLotEntry = Pack(iLotEntry) 	
					'		
					'	end if
					'   rsObj.close
						
						Response.Write "<p> clsscode = "& sClassCode
						sTempSeriesRcpt = GetInvNumberSeriesCodes("LO",sOrgCode,sClassCode)
	                    sArrSeriesRcpt = Split(sTempSeriesRcpt,":")
	                    iSNo = sArrSeriesRcpt(0)
	                    iSCode = sArrSeriesRcpt(1)
                	    
	                    sSQLQuery = "Select GroupName from INV_M_Classification where GroupCode = "& sClassCode
	                    rsTempRcpt.Open sSQLQuery,con
	                    if not rsTempRcpt.EOF then
	                        sNumClassNameRcpt = Trim(rsTempRcpt(0))
	                    end if
	                    rsTempRcpt.Close 
	                    Response.write "sFinToDt="& sFinToDt
	                    
	                    Response.write "iSNo="& iSno
	                    Response.write "iSNo="& iSCode
	                    
	                    
                    	
	                    if Trim(iSNo)="0" and Trim(iSCode)="0" then
	                        Response.Clear 
	                        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Lot Entry - "& sNumClassNameRcpt &"  Classification </H2></p>"
	                        Response.End 
	                    end if
						
						if not CheckNoSerAvilForThisYear(sOrgCode,iSNo,iSCode,sFinToDt) then
                            Response.Clear 
                            Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Lot Entry - "& sNumClassName &"  Classification for this Year </H2></p>"
                            Response.End 
                        end if
                        
                    
                        iLotEntry = GenSeriesNumber(sOrgCode,iSNo,iSCode,sFinToDt)	
						iLotEntry = Pack(iLotEntry) 	
					else
					    iLotEntry = iLot
						
					End If
					If TypeofStock = "S" Then iLotEntry = "Null"
					Response.Write ":Lot No :"& iLotEntry &":"
			'---------------------------
					
					if iValue > 0 then
					    iItmRate = cdbl(iValue) / cdbl(iQty)
					else
					    iItmRate = 0
					end if
					
					'iLotValue = cdbl(iItmRate) * cdbl(iLotQty)
					'Response.Write dIssDate
				'	sSQLQuery = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE," &_
				'		"ITEMCODE,CLASSIFICATIONCODE,RECEIPTQUANTITY,ACCEPTQUANTITY,ITEMRATE," &_
				'		"QUANTITYTYPE,WEIGHTTYPE,SRCTYPE,ACCOUNTEDON,ACCOUNTEDBY,RECEIVEDON," &_
				'		"ALTGROSS,ALTNETT,ALTUOM) VALUES " &_
				'		"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iICode & "," & sClassCode & "," &_
				'		"" & iQty & "," & iQty & "," & iRate & "," & Pack(sQtyIn) & "," &_
				'		"" & Pack(sTareIn) & "," &sTraType & ",CONVERT(DATETIME," & Pack(dIssDate) & ",103)," &_
				'		"" & iDefinedBy & ",CONVERT(DATETIME," & Pack(sReceivedOn) & ",103)," &_
				'		"" & iAltGross & "," & iAltNett & "," & sAltUoM & ")"
				'	Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				'	con.Execute sSQLQuery
				
				iQty = 0
				iNoofPacks = 0
				iLotValue = 0
				Response.Write "<p>ItemRate = "& iItmRate 
				sExpression1 ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE [ @STORE = "&sLocCode&" and @BIN = '"&sBinNo&"']/LotSerial [ @LOT = '"&iLot&"' and @COUNTER = "&iEntryCtr&"]/LotSerialDetails"
					Response.Write "<p>"& sExpression1 
					Set LotSerialNode = ndRoot.Selectnodes(sExpression1)
					

					For iCounter = 0 to LotSerialNode.Length - 1
					    iQty = CDbl(iQty) + (CDbl(LotSerialNode.Item(iCounter).Attributes.getNamedItem("QTYREC").value)-cdbl(LotSerialNode.Item(iCounter).Attributes.getNamedItem("TAREREC").value))
					    iNoofPacks = cdbl(iNoofPacks) + 1
					    iLotValue = CDbl(iLotValue)+ CDbl(LotSerialNode.Item(iCounter).Attributes.getNamedItem("IVALUE").value)
					Next

					sSQLQuery = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
						"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NOOFPACKS) VALUES " &_
						"(" & Pack(sOrgCode) & "," & iICode & "," & sClassCode & "," &_
						"" &sTraType & "," & iInvRecNo & ",CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iQty & "," & iValue & "," & iNoofPacks & ")"
					Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
					con.Execute sSQLQuery

					sExpression1 ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE [ @STORE = "&sLocCode&" and @BIN = '"&sBinNo&"']/LotSerial [ @LOT = '"&iLot&"' and @COUNTER = "&iEntryCtr&"]/LotSerialDetails"
					Response.Write "<p>"& sExpression1 
					Set LotSerialNode = ndRoot.Selectnodes(sExpression1)
					

					For iCounter = 0 to LotSerialNode.Length - 1
						If AutoGen <> "AUTO" Then iLotEntry = iLot
						iSerialEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value)
						iRcptSerialEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value)
						iQtyRecEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("QTYREC").Value)
						iTareQtyEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("TAREREC").Value)

						sSellingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGTYPE").Value)
						iWeight = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("WEIGHTSTYPE").Value)
						sPackingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKINGTYPE").Value)

						sSellingFormType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGFORM").Value)
						response.Write "<p>Pack Who + "& sPackWho 
						if sPackWho <> "N" then
							sPackingNumber = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKNUMBER").Value)
						elseif sPackWho = "N" then
							sPackingNumber = PackingCodeSeries("4", "4", sProductCode, iICode, sClassCode, sOrgCode, sPackingType, sItemType, "I")
						end if
						'Response.Write "<p> Packing NUmber = "& sPackingNumber 
						sSubLevelID = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SUBLEVELID").Value)
						sNoofCone = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("NOOFCONE").Value)
						
						sLotAttribute = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").Value)
						'Response.Write "<p> sTraType  ="& sTraType
						iStockQuality =trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SQ").value)
						
						if trim(sLotAttribute)="" or IsNull(sLotAttribute) then sLotAttribute="NULL"
						if trim(sLotAttribute)<>"NULL" then sLotAttribute=Pack(sLotAttribute)
						
						sSQLQuery = "Select IsNull(AutomaticAccounting,'N') from APP_M_ApplicationSetup where ApplicationCode = 6 and ReferenceCodeNo = 4"
						
                        rsObj.Open sSQLQuery,con
                        if not rsObj.EOF then
                            sAutoAccountingPack = rsObj(0)
                        else
                            sAutoAccountingPack = "N"
                        end if
                        rsObj.Close 
						'  To Create Pack Series Code  ************************
						Response.Write "<p>sTraType = "& sTraType 
						'Response.Write "<p>AutoAccounting Pack = "& sAutoAccountingPack 
						''added by ragav on aug 13,2013 for MovItem,Phyadj,MergeItem
						'If ((sTraType = "'RB'" or sTraType = "'RF'") or  sRecNum="S") and trim(sAutoAccountingPack) = "N" Then
						If ((sTraType = "'RB'" or sTraType = "'RF'") or  sRecNum="S") and trim(sAutoAccountingPack) = "N" and sUseExistingPackNum="N" Then
						
						
						Response.Write "<P>Welcome to Packing Number Series"
						
						      sTempSeriesRcpt =  GetPRDNumberSeriesCodes(sOrgCode,sClassCode)
						      sArrSeriesRcpt = Split(sTempSeriesRcpt,":")
						      iSNo = sArrSeriesRcpt(0)
						      iSCode = sArrSeriesRcpt(1)
						  
						    sSQLQuery = "Select GroupName from INV_M_Classification where GroupCode = "& sClassCode
	                        rsTempRcpt.Open sSQLQuery,con
	                        if not rsTempRcpt.EOF then
	                            sNumClassNameRcpt = Trim(rsTempRcpt(0))
	                        end if
	                        rsTempRcpt.Close 
	                    
	                    response.write "<p>Fin To Date = "& sFinTODt
                    	
	                    if Trim(iSNo)="0" and Trim(iSCode)="0" then
	                        Response.Clear 
	                        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Packing Number - "& sNumClassNameRcpt &"  Classification </H2></p>"
	                        Response.End 
	                        
	                    end if
	                    
	                    if not CheckNoSerAvilForThisYear(sOrgCode,iSNo,iSCode,sFinToDt) then
                            Response.Clear 
                            Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Packing Number - "& sNumClassName &"  Classification for this Year </H2></p>"
                            Response.End 
                        end if
                        
                        sPackingNumber = GenSeriesNumber(sOrgCode,iSNo,iSCode,sFinToDt)										 			
                        
							
							Response.Write "Packing Number = "& sPackingNumber 
						End If
						'*****************************************************						
						
						
						
						
						if sRecNum = "S" then iLotValue = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value)
						
						if sSellingType = "select" then
							sSellingType = "NULL"
							iWeight = "NULL"
						end if

						if sPackingType = "select" then
							sPackingType = "NULL"
						end if

						if sSellingFormType = "select" then
							sSellingFormType = "NULL"
						end if

						if iLotEntry <> "N/A" and iSerialEntry = "0" then iSerialEntry = "NULL"
						if iLotEntry = "N/A" then
							iLotEntry = "NULL"
						else
							If AutoGen <> "AUTO" Then iLotEntry = Pack(iLotEntry)
						end if

						iQtyGross = cDbl(iQtyRecEntry)
						iQtyNett = cdbl(iQtyRecEntry) - cdbl(iTareQtyEntry)

						if sItemType = "FAB" then iQtyNett = iQtyRecEntry

						If AutoGen = "AUTO" Then 

							with rsObj
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LocationLOT"
								.ActiveConnection = con
								.Open
							end with
							set rsObj.ActiveConnection = nothing

							if not rsObj.EOF then
								iSerialEntry = rsObj(0)
							end if
							rsObj.Close
						
						End If

						if sRecNum = "S" then iRate = FormatNumber(cdbl(iLotValue) / cdbl(iQtyNett),4,,,0)
						if Trim(sBinNo)="" or IsNull(sBinNo) or trim(sBinNo)="NULL" then sBinNo = "0"
						if Trim(sSellingType)="" or IsNull(sSellingType) then sSellingType="NULL"
						if Trim(sSellingFormType)="" or IsNull(sSellingFormType) then sSellingFormType = "NULL"
						if Trim(sSubLevelID)="" or IsNull(sSubLevelID) then sSubLevelID = "0"
						if Trim(sNoofCone)="" or IsNull(sNoofCone) then sNoofCone = "0"
						if trim(iLotEntry)="" or IsNull(iLotEntry) then iLotEntry = "NULL"
						
						Response.write "<p>Lot NUmber="& iLotEntry
						
						if trim(iStockQuality)="" or IsNull(iStockQuality) then iStockQuality="NULL"
						
						If iWeight = "" Then iWeight = 0		
						if sBinNo = "0" then sBinNo = "NULL"				
						sSQLQuery = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
							"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
							"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE,PACKINGSUBLEVELID,PACKINGSUBLEVELQTY,PACKINGSUBLEVELUNITQTY,DateOfReceipt,SrcType,AttributeList,StockQuality) VALUES " &_
							"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iICode & "," & sClassCode& "," &_
							"" & sLocCode & "," & sBinNo & "," & iLotEntry & "," & iSerialEntry & "," &_
							"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & "," & Pack(sPackingNumber) & "," &_
							"" & sPackingType & "," & sSellingType & "," & iWeight & "," & sSellingFormType & "," & sStage & "," & iRate & ","& sSubLevelID &","& sNoofCone &","& iWeight &",Convert(datetime,'"& dIssDate &"',103),"& sTraType &","& sLotAttribute &","& iStockQuality &")"
						Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
						con.Execute sSQLQuery
						
						
						sSQLQuery = "Update APP_T_InternalReceiptDetails set PackingNum = "& Pack(sPackingNumber)
						sSQLQuery = sSQLQuery &" where InternalReceiptNo = "& iIntRecNo &" and ClassificationCode = "& sClassCode &" and ItemCode = "& iICode
						if Trim(sLotAttribute)<>""  and trim (sLotAttribute)<>"NULL" and trim (sLotAttribute)<>"'NULL'" then
						    sSQLQuery = sSQLQuery & " AttributeList = " & sLotAttribute
						end if 
						if Trim(iLotEntry)<>"" and ucase(Trim(iLotEntry))<>"NULL" then
						 sSQLQuery = sSQLQuery &" and LotNo = "& iLotEntry
						end if 
						sSQLQuery = sSQLQuery & " and SerialNo = "& iRcptSerialEntry
						Response.Write "<p>"&sSQLQuery 
						con.execute sSQLQuery

						iPieceNo = 0
						if LotSerialNode.Item(iCounter).HaschildNodes() then
							set FabNode = LotSerialNode.Item(iCounter).childNodes(0)
							iAltGross = ""
							iAltNett = ""

							if FabNode.HaschildNodes() then
								arrFab = split(trim(FabNode.Attributes.getNamedItem("QUANTITYIN").Value),":")
								sCheck = trim(FabNode.Attributes.getNamedItem("CHECK").Value)
								iAltGross = trim(FabNode.Attributes.getNamedItem("ALTGROSS").Value)
								iAltNett = trim(FabNode.Attributes.getNamedItem("ALTNETT").Value)

								sQtyIn = arrFab(0)
								iQtyRecEntry = cdbl(arrFab(1))

								if sQtyIn = "N" then
									iQtyGross = iQtyRecEntry + cdbl(iTareQty)
								else
									iQtyGross = iQtyRecEntry
								end if
								
								if sCheck = "Y" then
									sSQLQuery = "UPDATE INV_T_LOCATIONLOT SET LOTQUANTITYGROSS = " & iQtyGross & "," &_
										"LOTQUANTITYTARE = " & iTareQty & " WHERE " &_
										"SERIALNUMBER = " & iSerialEntry & ""
									'objTxt.write sSQLQuery & vbCrLf & vbCrLf
									'Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
									'con.Execute sSQLQuery

									if iAltGross = "" then iAltGross = "0"
									if iAltNett = "" then iAltNett = "0"

									sSQLQuery = "UPDATE INV_T_LOCATIONLOT SET ALTGROSS = " & iAltGross & "," &_
										"ALTNETT = " & iAltNett & " WHERE " &_
										"SERIALNUMBER = " & iSerialEntry & ""
									Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
									con.Execute sSQLQuery

								end if
														
								iAltGross = ""
								iAltNett = ""
								For Each FabDetNode In FabNode.childNodes
									'iPieceNo = trim(FabDetNode.Attributes.getNamedItem("PIECENO").Value)
									iPieceQty = trim(FabDetNode.Attributes.getNamedItem("QUANTITY").Value)
									iAltGross = trim(FabDetNode.Attributes.getNamedItem("ALTGROSS").Value)
									iAltNett = trim(FabDetNode.Attributes.getNamedItem("ALTNETT").Value)

									if iAltGross = "" then iAltGross = "0"
									if iAltNett = "" then iAltNett = "0"
									
									if cdbl(iPieceQty) > 0 then
										iPieceNo = iPieceNo + 1
									'	sSQLQuery = "INSERT INTO INV_T_RECEIPTFABDETAILS (SERIALNUMBER,PIECENO,QUANTITY,ALTGROSS,ALTNETT) VALUES " &_
									'		"(" & iSerialEntry & "," & iPieceNo & "," & iPieceQty & "," & iAltGross & "," & iAltNett & ")"
									'	objTxt.write sSQLQuery & vbCrLf & vbCrLf
									'	'Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
									'	con.Execute sSQLQuery
									end if
								next
							end if
						end if
					next
				next
			else

				sSQLQuery = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
					"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NOOFPACKS) VALUES " &_
					"(" & Pack(sOrgCode) & "," & iICode & "," & sClassCode & "," &_
					"" &sTraType & ","& iInvRecNo &",CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iStorageQty & "," & iValue & "," & iNoofPacks & ")"
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery

			end if
			
			Response.Write "<p> iValue = "& iValue

			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM INV_M_ITEMSTORAGE WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND APPLICABLEFOR = " & Pack(sAppli) & " AND LOCATIONNUMBER = " & sLocCode & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL)"
				.ActiveConnection = con
				.Open
			end with
			Response.Write "<p>"&rsObj.Source 
			set rsObj.ActiveConnection = nothing
			if rsObj.EOF then
				sSQLQuery = "INSERT INTO INV_M_ITEMSTORAGE (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
					"APPLICABLEFOR,LOCATIONNUMBER,BINNUMBER,ALLOWTRANSFERS) VALUES " &_
					"(" & iICode & "," & sClassCode & "," & Pack(sOrgCode) & "," & Pack(sAppli) & "," &_
					"" & sLocCode & "," & sBinNo & ",'1')"
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			end if
			rsObj.Close

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			arrFinPeriod = split(GetFinancialYear(sMonthYr),":")
			sFinFromDt = arrFinPeriod(0)
			sFinToDt = arrFinPeriod(1)

			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				.ActiveConnection = con
				.Open
			end with
			set rsObj.ActiveConnection = nothing
			if rsObj.EOF then
				sSQLQuery = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
					"(" & Pack(sOrgCode) & "," & sClassCode & "," & iICode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103),CONVERT(DATETIME," & Pack(sFinToDt) & ",103)," & iStorageQty & "," & iValue & "," & iStorageQty & "," & iValue & ")"
				'objTxt.write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			else
				sSQLQuery = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iStorageQty & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
					"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iStorageQty & "), " &_
					"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
					"ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				objTxt.write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			end if
			rsObj.Close

			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLocCode & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				.ActiveConnection = con
				.Open
			end with

			set rsObj.ActiveConnection = nothing

			if rsObj.EOF then
				sSQLQuery = "INSERT INTO Inv_T_ItemLocationStock (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
					"YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
					"(" & Pack(sOrgCode) & "," & sClassCode & "," & iICode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103),CONVERT(DATETIME," & Pack(sFinToDt) & ",103)," &_
					"" & sLocCode & "," & sBinNo & "," &_
					"" & iStorageQty & "," & iValue & "," & iStorageQty & "," & iValue & ")"
		    	Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			else
				sSQLQuery = "UPDATE Inv_T_ItemLocationStock SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iStorageQty & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
					"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iStorageQty & "), " &_
					"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
					"ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"LOCATIONNUMBER = " & sLocCode & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinToDt) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				'objTxt.Write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"&sSQLQuery
				con.Execute sSQLQuery
			end if
			rsObj.Close
		next
		
		sSQLQuery = "Update App_T_InternalReceiptHeader set Status = 'Y', InvRecNo="& iInvRecNo &" where InternalReceiptNo = "& iIntRecNo
		Response.write "<p>"& sSQLQuery
		con.Execute sSQLQuery
			
End Function
%>
<%
Function UpdateInternalReceipt(CurrDate,InternalReceiptNo)
	
	dtCurrDate = CurrDate
	
	sFinPeriod = Session("FinPeriod")
	sFinFromDt = "01/04/" & Mid(sFinPeriod,1,4)
	sFinToDt = "31/03/" & Mid(sFinPeriod,6,4)
	
    newxml1.Load server.MapPath("../temp/transaction/IssRet"&Session.SessionID&".xml")
	newxml.async = false
    newxml.load Server.MapPath("../temp/transaction/ReceiptLotDataEdit"&Session.SessionID&".xml")'	
    
	iCreateby = getUserid
	iAppCode = "6"
	AutoGen = ""
	
	Set rsObj = Server.CreateObject("ADODB.RecordSet")
	Set rsObj1 = Server.CreateObject("ADODB.RecordSet")
	set rsTempRcpt = Server.CreateObject("ADODB.Recordset")

	Set ndRoot = newxml.documentElement
	Set ndRoot1 = newxml1.documentElement	
	sDept = trim(ndRoot.Attributes.getNamedItem("DEPT").Value)
	sSrc = trim(ndRoot.Attributes.getNamedItem("SOURCE").Value)
	sOrgCode = trim(ndRoot.Attributes.getNamedItem("ORGCODE").Value)
	'sTypeRcpt = trim(ndRoot.Attributes.getNamedItem("sTypeRcpt").Value)
	'sItemType = trim(ndRoot.Attributes.getNamedItem("ITEMTYPE").Value)
	'sPackWho = trim(ndRoot.Attributes.getNamedItem("PACKNUM").Value)
	sSourceRefType = trim(ndRoot.Attributes.getNamedItem("SRCREFTYPE").Value)
	sSourceRefNo = trim(ndRoot.Attributes.getNamedItem("SRCREFNO").Value)	
	TypeofStock =  trim(ndRoot.Attributes.getNamedItem("RCPTNUMBERINV").Value)	
	sAppRefTypeRcpt = trim(ndRoot.Attributes.getNamedItem("APPREFTYPE").value)
	sAppRefnoRcpt = trim(ndRoot.Attributes.getNamedItem("APPREFNO").value)
	sAppRefDateRcpt = trim(ndRoot.Attributes.getNamedItem("APPREFDATE").value)
	sReceviedOn = trim(ndRoot.Attributes.getNamedItem("RCVDON").value)
	sAutoAccount = trim(ndRoot.Attributes.getNamedItem("AUTOACCOUNT").value)
	
	if trim(sAppRefTypeRcpt)="" or isNull(sAppRefTypeRcpt) or trim(sAppRefTypeRcpt)="N" then sAppRefTypeRcpt = 0
	if trim(sAppRefnoRcpt)="" or IsNull(sAppRefTypeRcpt) then sAppRefnoRcpt = "NULL"
	if trim(sAppRefnoRcpt)<>"NULL"  then pack(sAppRefnoRcpt)
	if trim(sAppRefDateRcpt)="" or isNull(sAppRefDateRcpt) then sAppRefDateRcpt = "NULL"
	if trim(sAppRefDateRcpt)<>"NULL" then sAppRefDateRcpt = pack(sAppRefDateRcpt)
	
	if trim(sUseExistingPackNum)="" or IsNull(sUseExistingPackNum) then sUseExistingPackNum="N"
	
	Response.write "<p>sUseExistingPackNum="& sUseExistingPackNum
	
	if sTypeRcpt = "T" then
		sTraType = "'RB'"
	else
		sTraType = "'RR'"
	end if
	IF sTypeRcpt = "P" Then sTraType = "'RF'"
	iIntRecNo = InternalReceiptNo
	if ndRoot.HaschildNodes() then
		with rsObj
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT IsNull(InvRecNo,0) from APP_T_INTERNALRECEIPTHEADER where INTERNALRECEIPTNO ="& iIntRecNo 
			.ActiveConnection = con
			.Open
		end with
		Response.Write "<p>"& rsObj.Source 
		set rsObj.ActiveConnection = nothing

		if not rsObj.EOF then
			iInvRecNo = trim(rsObj(0))
		end if
		rsObj.Close
		
		dIssDate = dtCurrDate
'		Response.Write " dISSDATE="& dtCurrDate
		if sSrc = "N" then
			For Each ndDet In ndRoot.childNodes
			    if ndDet.nodeName="Details" then
			        for each ndItemDet in ndDet.childNodes
			            if ndItemDet.nodeName="ItemDetail" then
			                sClassCode = trim(ndItemDet.Attributes.getNamedItem("CLACODE").Value)
				            iICode = trim(ndItemDet.Attributes.getNamedItem("ItemCode").Value)
				            iQty = trim(ndItemDet.Attributes.getNamedItem("QTY").Value)
				            iTempMRSNo = trim(ndItemDet.Attributes.getNamedItem("MRSNO").Value)
				            iIssNo = trim(ndItemDet.Attributes.getNamedItem("ISSNO").Value)
				            sByProduct = trim(ndItemDet.Attributes.getNamedItem("BYPRODUCT").value)
				            sRcptNum = trim(ndItemDet.Attributes.getNamedItem("RECEIPTNUM").value)

				           
				            if iTempMRSNo = "N" then
					            iMRSNumber = "NULL"
				            else
					            iMRSNumber = iTempMRSNo
				            end if
				            
				            if trim(sRcptNum)="N" then
				                 for each ndIntStorage in ndItemDet.childNodes
				                    if trim(ndIntStorage.nodeName)="STORAGE" then
				                        iIntStore = ndIntStorage.getAttribute("STORE")
				                        iIntBin = ndIntStorage.getAttribute("BIN")
				                        iIntStoreVal = ndIntStorage.getAttribute("STORAGEVALUE")
				                        iIntStoreQty = ndIntStorage.getAttribute("QTY")
				                        if cdbl(iIntStoreVal)>0 then
				                            iIntRate = cdbl(iIntStoreVal)/cdbl(iIntStoreQty)
				                        end if
				                        sIntStkQty = ndIntStorage.getAttribute("SQ")
				                        
				                        sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
					                        "CLASSIFICATIONCODE,ITEMCODE,QUANTITYRETURN,PRODUCTTYPE,StockQuality,STORE,BIN,RATE) VALUES " &_
					                        "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," &_
					                        "" & iQty & ","& pack(sByProduct) &","& sIntStkQty &","& iIntStore &","& iIntBin &","& iIntRate &")"
				                        Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                        con.Execute sSQLQuery
				                    end if 'if trim(ndIntStorage.nodeName)="STORAGE" then
				                next
				            else
				                for each ndIntStorage in ndItemDet.childNodes
				                    if trim(ndIntStorage.nodeName)="STORAGE" then
				                        iIntStore = ndIntStorage.getAttribute("STORE")
				                        iIntBin = ndIntStorage.getAttribute("BIN")
				                        iIntStoreVal = ndIntStorage.getAttribute("STORAGEVALUE")
				                        iIntStoreQty = ndIntStorage.getAttribute("QTY")
				                        if cdbl(iIntStoreVal)>0 and iIntStoreQty>0 then
				                            iIntRate = cdbl(iIntStoreVal)/cdbl(iIntStoreQty)
				                        else
				                            iIntRate = 0
				                        end if
				                        
				                        for each ndIntLotSer in ndIntStorage.childNodes
				                            if trim(ndIntLotSer.nodeName)="LotSerial" then
				                                for each ndIntLotSerDet in ndIntLotSer.childNodes
				                                    iQtyReturn = ndIntLotSerDet.getAttribute("QTYREC")
				                                    sIntLotNo = ndIntLotSerDet.getAttribute("LOT")
				                                    sIntPackNum = ndIntLotSerDet.getAttribute("PACKNUMBER")
				                                    iIntSerNo = ndIntLotSerDet.getAttribute("LOTSERIAL")
				                                    iIntPackCode =  ndIntLotSerDet.getAttribute("PACKINGTYPE")
				                                    sIntAttList = ndIntLotSerDet.getAttribute("ATTRIBUTELIST")
				                                    sIntStkQty = ndIntLotSerDet.getAttribute("SQ")
				                                    iIntQtyTare = ndIntLotSerDet.getAttribute("TAREREC")
				                                    sIntPackStatus = ndIntLotSerDet.getAttribute("STATUS")
				                                    
				                                    iIntQtyNett = iQtyReturn-iIntQtyTare
				                                    
				                                    
				                                    if trim(sIntLotNo)="" or IsNull(sIntLotNo) then sIntLotNo="NULL"
				                                    if trim(sIntLotNo)<>"NULL" then sIntLotNo=Pack(sIntLotNo)
				                                    if trim(sIntAttList)="" or IsNull(sIntAttList) then sIntAttList="NULL"
				                                    if trim(sIntAttList)<>"NULL" then sIntAttList= pack(sIntAttList)
				                                    if trim(iIssNo)="N" or trim(iIssNo)="" or IsNull(iIssNo) then iIssNo = "NULL"
				                                    
				                                    if trim(sIntStkQty)="" or IsNull(sIntStkQty) then sIntStkQty ="NULL"
				                                    
				                                    if trim(sIntPackStatus) = "N" then
				                                        sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
									                    "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYRETURN,PackingNum,AttributeList,ProductType,StockQuality,STORE,BIN,RATE,PackingCode,GrossQuantityReturn) VALUES " &_
									                    "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
									                    "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & sIntLotNo & "," & iIntSerNo & "," & iIntQtyNett & "," & sIntPackNum & ","& sIntAttList &","& pack(sByProduct) &","& sIntStkQty &","& iIntStore &","& iIntBin &","& iIntRate &","& iIntPackCode &","& iQtyReturn &")"
									                    Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                                        con.Execute sSQLQuery
									                elseif Trim(sIntPackStatus)="U" or Trim(sIntPackStatus)="O" then
									                    sSQLQuery = "Update APP_T_INTERNALRECEIPTDETAILS set MRSNUMBER=" & iMRSNumber  & ","
									                    sSQLQuery = sSQLQuery &" ISSUENO=" & iIssNo & ",ISSUEDATE=CONVERT(DATETIME," & Pack(dIssDate) & ",103),"
									                    sSQLQuery = sSQLQuery &" QUANTITYRETURN=" & iIntQtyNett & ","
									                    sSQLQuery = sSQLQuery &" ProductType="& pack(sByProduct) &",StockQuality="& sIntStkQty &","
									                    sSQLQuery = sSQLQuery &" STORE="& iIntStore &",BIN="& iIntBin &",RATE="& iIntRate &",PackingCode="& iIntPackCode &",GrossQuantityReturn="& iQtyReturn&"," 
									                    sSQLQuery = sSQLQuery &" PackingNum=" & sIntPackNum 
									                    sSQLQuery = sSQLQuery &" where INTERNALRECEIPTNO = "& iIntRecNo 
									                    sSQLQuery = sSQLQuery & " and CLASSIFICATIONCODE=" & sClassCode &" and ITEMCODE=" & iICode &" and SERIALNO=" & iIntSerNo 
									                    if trim(sIntLotNo)<>"" and trim(sIntLotNo)<>"NULL" then
									                        sSQLQuery =sSQLQuery &" and LOTNO=" & sIntLotNo 
									                    end if 
									                    if trim(sIntAttList)<>"" and trim(sIntAttList)<>"NULL" then
									                        sSQLQuery = sSQLQuery &" and AttributeList="& sIntAttList 
									                    end if 
									                    
									                    
									                    Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                                        con.Execute sSQLQuery
				                                    elseif Trim(sIntPackStatus)="D" then
				                                        sSQLQuery = "Delete from APP_T_INTERNALRECEIPTDETAILS where INTERNALRECEIPTNO=" & iIntRecNo & " and ITEMCODE="& iICode &"  and SERIALNO=" & iIntSerNo  
				                                        if trim(sIntLotNo)<>"" and trim(sIntLotNo)<>"NULL" then
									                        sSQLQuery =sSQLQuery &" and LOTNO=" & sIntLotNo 
									                    end if 
									                    if trim(sIntAttList)<>"" and trim(sIntAttList)<>"NULL" then
									                        sSQLQuery = sSQLQuery &" and AttributeList="& sIntAttList 
									                    end if 
				                                        Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                                        con.Execute sSQLQuery
				                                    end if
				                                next
				                            end if
				                        next
				                    end if
				                next
				            end if
				            Response.Write "<p>Auto Accounting  = "& sAutoAccount 
				            if trim(sAutoAccount)="Y" then
				                UpdateLotDetails(iICode)
				            elseif Trim(iInvRecNo)<>"" then
				                UpdateLotDetails(iICode)
				            end if 
				        end if 'if ndItemDet.nodeName="ItemDetail" then
				    next
				end if 'if ndDet.nodeName="Details" then
			next
			'Calling Function to insert Lot/Serial Details
			
			
			
			
		else 'MRS / Direct Issue Based
		
			sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTHEADER (INTERNALRECEIPTNO,APPLICATIONCODE," &_
				"ORGANISATIONCODE,CREATEDFROMDEPT,REFTYPE,CREATEDON,CREATEDBY,STATUS,RECEIVEDON) VALUES " &_
				"(" & iIntRecNo & "," & iAppCode & "," & Pack(sOrgCode) & "," & Pack(sDept) & "," & Pack(sSrc) & "," &_
				"CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iCreateby & ",'N',Convert(datetime,"& pack(sReceivedOn) &",103))"
			Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
			con.Execute sSQLQuery
			
			For Each ndDet In ndRoot.childNodes
			    if ndDet.nodeName="Details" then
			        for each ndItemDet in ndDet.childNodes
			            if ndItemDet.nodeName="ItemDetail" then
		    	                sClassCode = trim(ndItemDet.Attributes.getNamedItem("CLACODE").Value)
				                iICode = trim(ndItemDet.Attributes.getNamedItem("ItemCode").Value)
				                iQty = trim(ndItemDet.Attributes.getNamedItem("QTY").Value)
				                iMRSNumber = trim(ndItemDet.Attributes.getNamedItem("MRSNO").Value)
				                iDINo = trim(ndItemDet.Attributes.getNamedItem("DINO").Value)
				                iIssNo = trim(ndItemDet.Attributes.getNamedItem("ISSNO").Value)
				                sRemark = trim(ndItemDet.Attributes.getNamedItem("REMARKS").Value)
				                with rsObj
					                .CursorLocation = 3
					                .CursorType = 3
					                .Source = "SELECT IssueEntryNo FROM INV_T_DEPARTMENTSTOCK WHERE ISSUENO = '" & iIssNo & "'"
					                .ActiveConnection = con
					                .Open
				                end with
				                'Response.Write rsObj.Source 
				                set rsObj.ActiveConnection = nothing

				                if not rsObj.EOF then
					                iIssueEntryNo = trim(rsObj(0))
				                end if
				                rsObj.Close
				                'Response.Write iIssueEntryNo 
				                 ' This is added for the purpose of to add WorkCenterWise details
				                sExpression ="//ItemDetail [ @CLACODE = "&sClassCode&" and @ItemCode = "&iICode&"]/AddDet/WorkCenter"
				                Set WorkCenterNode1 = ndRoot1.Selectnodes(sExpression)
				                'Response.Write WorkCenterNode1.Length
				                If WorkCenterNode1.Length > 0 Then
				                For i = 0 to WorkCenterNode1.Length-1
					                WCODE = trim(WorkCenterNode1.Item(i).Attributes.getNamedItem("WCODE").Value)		
					                'Response.Write WCODE					
					                sExpression ="//ItemDetail [ @CLACODE = "&sClassCode&" and @ItemCode = "&iICode&"]/AddDet/WorkCenter [ @WCODE = '"&WCODE&"']/MachineCenter"
					                Set MachineCenterNode1 = ndRoot1.Selectnodes(sExpression) 
					                'Response.Write MachineCenterNode1.Length & sExpression
					                For j = 0 to MachineCenterNode1.Length-1
						                MCODE =  trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("MCODE").Value)							
						                iQtyIssued = trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("QTY").Value)												
						                iQtyReturned = trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("QTYRETURNED").Value)												
						                iQtyReturn = trim(MachineCenterNode1.Item(j).Attributes.getNamedItem("QTYRETURN").Value)																	
						                iTotReturn = iQtyReturned + cDbl(iQtyReturn)																		
            								
						                If cDbl(iQtyReturn) > 0 Then
							                sSQLQuery = "INSERT INTO APP_T_InternalReceiptAddnDet (InternalReceiptNo,ClassificationCode,ItemCode,WorkCenterCode,MachineCenterCode,QuantityReturn) VALUES " &_
							                "(" & iIntRecNo & "," & sClassCode & ", " & iICode & ",'" & WCODE & "', '" & MCODE & "'," & iQtyReturn & ")"
							                Response.Write "<p>"& sSQLQuery
							                con.Execute sSQLQuery	
            									
							                sSQLQuery = "UPDATE inv_t_issuedeptwiseBreakUp SET QUANTITYRETURNED = " &_
								                "(ISNULL(QUANTITYRETURNED,0) + " & iQtyReturn & ") WHERE ISSUEENTRYNO = " & iIssueEntryNo & " AND " &_
								                " ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
								                "CLASSIFICATIONCODE = " & sClassCode & " AND ITEMCODE = " & iICode & " AND WORKCENTERCODE = '" & WCODE & "' AND MACHINECENTERCODE = '" & MCODE & "'"									
							                Response.Write "<p>"& sSQLQuery
							                con.Execute sSQLQuery																			
						                End If
					                Next
				                Next 			
            					
				                End If

		                '------------------------------------------------------------------------------------
            					
            					
            					
            					
            					
            					
            					
				                if Trim(iMRSNumber) = "" then iMRSNumber = "NULL"
				                if Trim(iDINo) = "" then iDINo = "NULL"
				                if sRemark = "" then
					                sRemark = "NULL"
				                else
					                sRemark = Pack(sRemark)
				                end if

				                if Trim(iMRSNumber) <> "NULL" then
					                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCK SET QUANTITYRETURNED = " &_
						                "(ISNULL(QUANTITYRETURNED,0) + " & iQty & ") WHERE ISSUENO = " & iIssNo & " AND " &_
						                "MRSNUMBER = " & iMRSNumber & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
						                "CLASSIFICATIONCODE = " & sClassCode & " AND ITEMCODE = " & iICode & ""
				                else
					                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCK SET QUANTITYRETURNED = " &_
						                "(ISNULL(QUANTITYRETURNED,0) + " & iQty & ") WHERE ISSUENO = " & iIssNo & " AND " &_
						                "DINUMBER = " & iDINo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
						                "CLASSIFICATIONCODE = " & sClassCode & " AND ITEMCODE = " & iICode & ""
				                end if
            										
				                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				                con.Execute sSQLQuery

				                For Each LotNode in ndItemDet.ChildNodes
					                iLot = trim(LotNode.Attributes.Item(0).nodeValue)
					                dIssDate = trim(LotNode.Attributes.Item(1).nodeValue)
					                iQtyRet = trim(LotNode.Attributes.Item(3).nodeValue)
					                'iDeptLineNo = trim(LotNode.Attributes.Item(4).nodeValue)



					                if iLot = "0" or iLot = "" or iLot = "NULL" then
						                iLot = "NULL"
					                else
						                iLot = Pack(iLot)
					                end if

					                if not iQtyRet = "0" then
						                if LotNode.hasChildNodes then
							                For Each SerialNode in LotNode.ChildNodes
								                iSerialNo = trim(SerialNode.Attributes.Item(0).nodeValue)
								                iQtySerial = trim(SerialNode.Attributes.Item(1).nodeValue)

								                if not iQtySerial = "0" then
									                if Trim(iMRSNumber) <> "NULL" then
										                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
											                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYRETURN,REMARKS) VALUES " &_
											                "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
											                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iLot & "," & iSerialNo & "," & iQtySerial & "," & sRemark & ")"
									                else
										                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,DINUMBER," &_
											                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYRETURN,REMARKS) VALUES " &_
											                "(" & iIntRecNo & "," & iDINo & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
											                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iLot & "," & iSerialNo & "," & iQtySerial & "," & sRemark & ")"
									                end if										
									                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
									                con.Execute sSQLQuery

									                if iLot = "NULL" then
										                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
											                "(ISNULL(QUANTITYRETURNED,0) + " & iQtySerial & ") WHERE ISSUENO = " & iIssNo & " AND " &_
											                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
											                "LOTNO IS NULL AND SERIALNO = " & iSerialNo '& " AND LINENUMBER = " & iDeptLineNo & ""
									                else
										                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
											                "(ISNULL(QUANTITYRETURNED,0) + " & iQtySerial & ") WHERE ISSUENO = " & iIssNo & " AND " &_
											                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
											                "LOTNO = " & iLot & " AND SERIALNO = " & iSerialNo '& " AND LINENUMBER = " & iDeptLineNo & ""
									                end if
									                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
									                con.Execute sSQLQuery
								                end if
							                next
						                else
							                if Trim(iMRSNumber) <> "NULL" then
								                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
									                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,QUANTITYRETURN,REMARKS) VALUES " &_
									                "(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
									                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iQtyRet & "," & sRemark  & ")"
							                else
								                sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,DINUMBER," &_
									                "CLASSIFICATIONCODE,ITEMCODE,ISSUENO,ISSUEDATE,QUANTITYRETURN,REMARKS) VALUES " &_
									                "(" & iIntRecNo & "," & iDINo & "," & sClassCode & "," & iICode & "," & iIssNo & "," &_
									                "CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iQtyRet & "," & sRemark  & ")"
							                end if								
							                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
							                con.Execute sSQLQuery

							                if iLot = "NULL" then
								                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
									                "(ISNULL(QUANTITYRETURNED,0) + " & iQtyRet & ") WHERE ISSUENO = " & iIssNo & " AND " &_
									                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
									                "LOTNO IS NULL AND (SERIALNO IS NULL OR SERIALNO = 0)" ' AND LINENUMBER = " & iDeptLineNo & ""
							                else
								                sSQLQuery = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYRETURNED = " &_
									                "(ISNULL(QUANTITYRETURNED,0) + " & iQtyRet & ") WHERE ISSUENO = " & iIssNo & " AND " &_
									                "CONVERT(datetime,ISSUEDATE,103) = CONVERT(datetime," & Pack(dIssDate) & ",103) AND " &_
									                "LOTNO = " & iLot & " AND (SERIALNO IS NULL OR SERIALNO = 0)" ' AND LINENUMBER = " & iDeptLineNo & ""
							                end if
							                Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
							                con.Execute sSQLQuery
						                end if
					                end if
				                next
				        ' end if for Item Node check
				        end if
				    next
				end if
			next
		end if
	end if
End Function
%>
<%	'Function to insert the Lot/Serial Details for None Cases
	Function UpdateLotDetails(sSendItemCode)
		dim ItemNode,StorageNode,LotNode,LotSerialNode,objfs,FabNode,FabDetNode,OptionalUoMNode
		dim iNoofPacks, PackNode
		dim sTemp,iCtr,sLocCode,sBinNo,sMonthYr,iQty,iValue,sAppli
		dim arrFinPeriod,arrTemp,iCounter,sTempMonthYear
		dim sQtyIn,iTareQty,iLot,iSerialFrom,iSerialTo,sSellingType,iWeight,sPackingType,sPackingNumber
		dim iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sTareIn,iQtyGross,iQtyNett
		dim iSeriesNo,iSeriesCode,sProductCode,iItmRate,sItmShDesc,sItmAddDesc
		dim iLotCtr,sSellingFormType,sStage,iCodeLen,iCodeSize,sForm
		dim sDrwVerNo,sCatalogue,iCommodity,iLotQty,iLotValue,iEntryCtr,sReceivedOn
		dim iFabCtr,iPieceNo,iPieceQty,arrFab,iAltGross,iAltNett,sAltUoM,sCheck,iRate
		dim sOpUoMCode,iOpUoMFactor,sOpUoMOperator,sRecNum,iStorageQty
		dim objFSO,objTxt,sExpression,sExpression1,iDefinedBy,arrDate
		Dim iSNo, iSCode, sPackCode ,ItemNodeDet,sLotAttribute
		
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Transaction/ReceiptLotData.txt"))

		iDefinedBy = getUserid
		
		iNoofPacks = "Null"
		
		with rsObj
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(COMPANYITEMCODE,'') FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iICode & ""
			Response.Write rsObj.Source 
			.ActiveConnection = con
			.Open
		end with
		set rsObj.ActiveConnection = nothing
		if not rsObj.EOF then
			sProductCode = trim(rsObj(0))
		else
			sProductCode = ""
		end if
		rsObj.Close

		with rsObj
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT RECEIPTNUMBERING FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
			.ActiveConnection = con
			.Open
		end with
		set rsObj.ActiveConnection = nothing
		if not rsObj.EOF then
			sRecNum = trim(rsObj(0))
		end if
		rsObj.Close
		
		' Storage Location Details Insert
		
		sExpression = "//ROOT/Details/ItemDetail[@ItemCode="&sSendItemCode&"]"
		set ItemNodeDet = ndRoot.Selectnodes(sExpression)
		if ItemNodeDet.length>0 then
		    TypeOfStock = ItemNodeDet.Item(0).Attributes.getNamedItem("RECEIPTNUM").value
		end if
		
		'This Part Added For the purpose of to add No of Packs in Ledger Entries
		sExpression ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE/LotSerial/LotSerialDetails[@STATUS!='D']"
		Set PackNode = ndRoot.Selectnodes(sExpression)
		If PackNode.Length > 0 Then
			iNoofPacks = PackNode.Length
		End If
		
		Response.Write "<p>No of Packs = "& iNoofPacks
		
		
			
		
		iCtr = 0
		sExpression ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE"
		Set StorageNode = ndRoot.Selectnodes(sExpression)
		Response.Write "<p>No of Packs = "& iNoofPacks
		for iCtr =  0 to StorageNode.length - 1
		
			sLocCode = StorageNode.Item(iCtr).Attributes.getNamedItem("STORE").Value
			sBinNo = StorageNode.Item(iCtr).Attributes.getNamedItem("BIN").Value
			sAppli = StorageNode.Item(iCtr).Attributes.getNamedItem("APPLICABLE").Value
			sMonthYr = StorageNode.Item(iCtr).Attributes.getNamedItem("MONTHYEAR").Value
			iStorageQty = StorageNode.Item(iCtr).Attributes.getNamedItem("QTY").Value
			iValue = Round(Cdbl(StorageNode.Item(iCtr).Attributes.getNamedItem("STORAGEVALUE").Value),2)
			Response.Write " iValue = "& iValue
            Response.Write "<p>iQty = "& iQty 
			'sReceivedOn = "01/"&left(sMonthYr,2)&"/"&right(sMonthYr,4)
			sReceivedOn = sMonthYr
			if Trim(sReceivedOn)="" or IsNull(sReceivedOn) then sReceivedOn =  FormatDate(date)
			arrDate = split(sReceivedOn,"/")
			Response.Write "<p>sReceivedOn = "& sReceivedOn 
			sMonthYr = arrDate(1)&arrDate(2)
			'Response.Write sMonthYr
			dIssDate = sReceivedOn 
			if sBinNo = "0" then sBinNo = "NULL"
			
			if sRecNum = "N" then
			    if sBinNo = "0" then sBinNo = "NULL"
			    sSQLQuery = "Update INV_T_LOCATIONLOT ORGANISATIONCODE=" & Pack(sOrgCode) & ",STORAGELOCATIONNO=" & sLocCode & "," 
			    sSQLQuery = sSQLQuery & " STORAGEBINNUMBER=" & sBinNo & ",LOTQUANTITYGROSS=" & iStorageQty & ","
				sSQLQuery = sSQLQuery & " LOTQUANTITYNETT=" & iStorageQty & ",DateOfReceipt=Convert(datetime,'"& dIssDate &"',103),"
				sSQLQuery = sSQLQuery & " SrcType="& sTraType &",LotNumber=NULL,SerialNumber=NULL,iStockQuality"& iStockQuality 
				sSQLQuery = sSQLQuery & " where INVENTORYRECEIPTNO = "& iInvRecNo &" and ITEMCODE=" & iICode 
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			end if
			
			sSQLQuery = "Select TRANSACTQUANTITY,TRANSACTVALUE from INV_T_ITEMLEDGER where ORGANISATIONCODE=" & Pack(sOrgCode) & " and "
					sSQLQuery = sSQLQuery & " ITEMCODE=" & iICode & " and CLASSIFICATIONCODE =" & sClassCode & " and TRANSACTIONTYPE=" &sTraType & " and TRANSACTIONNO =" & iInvRecNo 
					rsTempRcpt.open sSQLQuery,con
					if not rsTempRcpt.eof then
					    nExistItemQty = rsTempRcpt(0)
					    nExistItemVal = rsTempRcpt(1)
					end if 
					rsTempRcpt.close
					
					Response.Write "<p> Exit Item Qty = "& nExistItemQty
					Response.Write "<p> Exit Item Value = "& nExistItemVal
					
					
					sExpression1 ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE [ @STORE = "&sLocCode&" and @BIN = '"&sBinNo&"']/LotSerial/LotSerialDetails[@STATUS!='D']"
					Response.Write "<p>"& sExpression1 
					Set LotSerialNode = ndRoot.Selectnodes(sExpression1)
					Response.Write "<p>LotSerialNode = "& LotSerialNode.length

					For iCounter = 0 to LotSerialNode.Length - 1
					    iQty = CDbl(iQty) + (CDbl(LotSerialNode.Item(iCounter).Attributes.getNamedItem("QTYREC").value)-cdbl(LotSerialNode.Item(iCounter).Attributes.getNamedItem("TAREREC").value))
					    iNoofPacks = cdbl(LotSerialNode.length)
					    iLotValue = CDbl(iLotValue)+ CDbl(LotSerialNode.Item(iCounter).Attributes.getNamedItem("IVALUE").value)
					Next
					
					

					sSQLQuery = "Update INV_T_ITEMLEDGER set TRANSACTIONDATE=CONVERT(DATETIME," & Pack(dIssDate) & ",103),TRANSACTQUANTITY="& iQty & ",TRANSACTVALUE=" & iLotValue & ",NOOFPACKS=" & iNoofPacks 
					sSQLQuery = sSQLQuery & " where ORGANISATIONCODE=" & Pack(sOrgCode) & " and ITEMCODE=" & iICode & " and CLASSIFICATIONCODE =" & sClassCode & " and TRANSACTIONTYPE=" &sTraType & " and TRANSACTIONNO =" & iInvRecNo 
					Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
					con.Execute sSQLQuery
					
			
			'if sBinNo = "NULL" or IsNull(sBinNo) then  sBinNo = 0 

			sExpression1 ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE [ @STORE = """&sLocCode&""" and @BIN = """&sBinNo&"""]/LotSerial"
			Response.Write "<p>"&sExpression1 
			Set LotNode = ndRoot.Selectnodes(sExpression1)
			Response.Write "<p>LotNode.Length="& LotNode.length

			if LotNode.Length > 0 then
			
				For iLotCtr = 0 to LotNode.Length - 1

					iCounter = 0

					sQtyIn = LotNode.Item(iLotCtr).Attributes.getNamedItem("QTYIN").Value
					iTareQty = LotNode.Item(iLotCtr).Attributes.getNamedItem("TARE").Value
					iLot = LotNode.Item(iLotCtr).Attributes.getNamedItem("LOT").Value
					iSerialFrom = LotNode.Item(iLotCtr).Attributes.getNamedItem("SERIALFROM").Value
					iSerialTo = LotNode.Item(iLotCtr).Attributes.getNamedItem("SERIALTO").Value
					sTareIn = LotNode.Item(iLotCtr).Attributes.getNamedItem("TAREWEIGHT").Value
					iLotQty = LotNode.Item(iLotCtr).Attributes.getNamedItem("QTY").Value
					iEntryCtr = LotNode.Item(iLotCtr).Attributes.getNamedItem("COUNTER").Value
					sStage = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("STAGE").Value)
					AutoGen = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("AUTOGEN").Value)
					'Response.Write iSerialFrom & "<" & iSerialTo 

					
					iAltGross = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTGROSS").Value)
					iAltNett = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTNETT").Value)
					sAltUoM = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("ALTUOM").Value)
					iLotValue = trim(LotNode.Item(iLotCtr).Attributes.getNamedItem("IVALUE").Value)
					
					if trim(iLotValue) = ""  or IsNull(iLotValue) then iLotValue = "0"
					
					if iLotValue>0 then									
					    iRate = FormatNumber(cdbl(iLotValue) / cdbl(iLotQty),4,,,0)
					else
					    iRate = 0
					end if
					
					iRate = Round(iRate)
					
					if sStage = "select" then
						sStage = "NULL"
					else
						sStage = Pack(sStage)
					end if
					
					if iTareQty > 0 then
					    iQty = iLotQty - iTareQty
					else
					    iQty = iLotQty
					end if
				
					'Response.Write "iQty = "& iQty 
					
					if iTareQty = "" then iTareQty = "0"
					if iAltGross = "" then iAltGross = "0"
					if iAltNett = "" then iAltNett = "0"
					if sAltUoM = "" or lcase(sAltUoM) = "select" then
						sAltUoM = "NULL"
					else
						sAltUoM = Pack(sAltUoM)
					end if


			'------------------------------------
					Response.Write "AutoGen = "&  AutoGen
					Response.write "TypeofStock="& TypeOfStock
					''added by ragav on aug 13,2013 for MovItem,Phyadj,MergeItem
					'If AutoGen = "AUTO" and (TypeofStock = "LS" or TypeofStock = "L") Then
					if Trim(iInvRecNo)<>"" then
					    sAutoAccountingPack = "Y"
					else
					sSQLQuery = "Select IsNull(AutomaticAccounting,'N') from APP_M_ApplicationSetup where ApplicationCode = 6 and ReferenceCodeNo = 4"
						
                        rsObj.Open sSQLQuery,con
                        if not rsObj.EOF then
                            sAutoAccountingPack = rsObj(0)
                        else
                            sAutoAccountingPack = "N"
                        end if
                        rsObj.Close 
                    end if
                        
					If AutoGen = "AUTO" and (TypeofStock = "LS" or TypeofStock = "L") and sAutoAccountingPack="N" and sUseExistingPackNum="N" Then
					
						with rsObj
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ITEMTYPEID FROM VWITEM WHERE ORGANISATIONCODE = '" & sOrgCode & "' AND ITEMCODE = '" & iICode & "'"
							.ActiveConnection = con
							.Open
						end with
						'Response.Write rsObj.Source 
						set rsObj.ActiveConnection = nothing	
						If Not rsObj.EOF Then
							sItemType = rsObj(0)
						End If
						rsObj.Close 
											
						Response.Write "<p> clsscode = "& sClassCode
						sTempSeriesRcpt = GetInvNumberSeriesCodes("LO",sOrgCode,sClassCode)
	                    sArrSeriesRcpt = Split(sTempSeriesRcpt,":")
	                    iSNo = sArrSeriesRcpt(0)
	                    iSCode = sArrSeriesRcpt(1)
                	    
	                    sSQLQuery = "Select GroupName from INV_M_Classification where GroupCode = "& sClassCode
	                    rsTempRcpt.Open sSQLQuery,con
	                    if not rsTempRcpt.EOF then
	                        sNumClassNameRcpt = Trim(rsTempRcpt(0))
	                    end if
	                    rsTempRcpt.Close 
	                    Response.write "sFinToDt="& sFinToDt
	                    
	                    Response.write "iSNo="& iSno
	                    Response.write "iSNo="& iSCode
	                    
	                    
                    	
	                    if Trim(iSNo)="0" and Trim(iSCode)="0" then
	                        Response.Clear 
	                        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Lot Entry - "& sNumClassNameRcpt &"  Classification </H2></p>"
	                        Response.End 
	                    end if
						
						if not CheckNoSerAvilForThisYear(sOrgCode,iSNo,iSCode,sFinToDt) then
                            Response.Clear 
                            Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Lot Entry - "& sNumClassName &"  Classification for this Year </H2></p>"
                            Response.End 
                        end if
                        
                    
                        iLotEntry = GenSeriesNumber(sOrgCode,iSNo,iSCode,sFinToDt)	
						iLotEntry = Pack(iLotEntry) 	
					else
					    iLotEntry = iLot
						
					End If
					If TypeofStock = "S" Then iLotEntry = "Null"
					Response.Write ":Lot No :"& iLotEntry &":"
			'---------------------------
					
					if iValue > 0 then
					    iItmRate = cdbl(iValue) / cdbl(iQty)
					else
					    iItmRate = 0
					end if
				
				iQty = 0
				iNoofPacks = 0
				iLotValue = 0
				Response.Write "<p>ItemRate = "& iItmRate 
				

					sExpression1 ="//ROOT/Details/ItemDetail[@ItemCode="& sSendItemCode &"]/STORAGE [ @STORE = "&sLocCode&" and @BIN = '"&sBinNo&"']/LotSerial [ @LOT = '"&iLot&"' and @COUNTER = "&iEntryCtr&"]/LotSerialDetails"
					Response.Write "<p>"& sExpression1 
					Set LotSerialNode = ndRoot.Selectnodes(sExpression1)
					

					For iCounter = 0 to LotSerialNode.Length - 1
						If AutoGen <> "AUTO" Then iLotEntry = iLot
						iSerialEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value)
						iQtyRecEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("QTYREC").Value)
						iTareQtyEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("TAREREC").Value)

						sSellingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGTYPE").Value)
						iWeight = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("WEIGHTSTYPE").Value)
						sPackingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKINGTYPE").Value)

						sSellingFormType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGFORM").Value)
						response.Write "<p>Pack Who + "& sPackWho 
						if sPackWho <> "N" then
							sPackingNumber = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKNUMBER").Value)
						elseif sPackWho = "N" then
							sPackingNumber = PackingCodeSeries("4", "4", sProductCode, iICode, sClassCode, sOrgCode, sPackingType, sItemType, "I")
						end if
						'Response.Write "<p> Packing NUmber = "& sPackingNumber 
						sSubLevelID = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SUBLEVELID").Value)
						sNoofCone = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("NOOFCONE").Value)
						
						sLotAttribute = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").Value)
						'Response.Write "<p> sTraType  ="& sTraType
						iStockQuality =trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SQ").value)
						sInvStatus = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("STATUS").value)
						
						if trim(sLotAttribute)="" or IsNull(sLotAttribute) then sLotAttribute="NULL"
						if trim(sLotAttribute)<>"NULL" then sLotAttribute=Pack(sLotAttribute)
						
						if Trim(iInvRecNo)<>"" then
					        sAutoAccountingPack = "Y"
					    else
						    sSQLQuery = "Select IsNull(AutomaticAccounting,'N') from APP_M_ApplicationSetup where ApplicationCode = 6 and ReferenceCodeNo = 4"
    						
                            rsObj.Open sSQLQuery,con
                            if not rsObj.EOF then
                                sAutoAccountingPack = rsObj(0)
                            else
                                sAutoAccountingPack = "N"
                            end if
                            rsObj.Close 
                        end if
						'  To Create Pack Series Code  ************************
						Response.Write "<p>sTraType = "& sTraType 
						Response.Write "<p>AutoAccounting Pack = "& sAutoAccountingPack 
						''added by ragav on aug 13,2013 for MovItem,Phyadj,MergeItem
						'If ((sTraType = "'RB'" or sTraType = "'RF'") or  sRecNum="S") and trim(sAutoAccountingPack) = "N" Then
						If ((sTraType = "'RB'" or sTraType = "'RF'") or  sRecNum="S") and trim(sAutoAccountingPack) = "N" and sUseExistingPackNum="N" Then
						
						
						Response.Write "<P>Welcome to Packing Number Series"
						
						      sTempSeriesRcpt =  GetPRDNumberSeriesCodes(sOrgCode,sClassCode)
						      sArrSeriesRcpt = Split(sTempSeriesRcpt,":")
						      iSNo = sArrSeriesRcpt(0)
						      iSCode = sArrSeriesRcpt(1)
						  
						    sSQLQuery = "Select GroupName from INV_M_Classification where GroupCode = "& sClassCode
	                        rsTempRcpt.Open sSQLQuery,con
	                        if not rsTempRcpt.EOF then
	                            sNumClassNameRcpt = Trim(rsTempRcpt(0))
	                        end if
	                        rsTempRcpt.Close 
	                    
	                    response.write "<p>Fin To Date = "& sFinTODt
                    	
	                    if Trim(iSNo)="0" and Trim(iSCode)="0" then
	                        Response.Clear 
	                        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Packing Number - "& sNumClassNameRcpt &"  Classification </H2></p>"
	                        Response.End 
	                        
	                    end if
	                    
	                    if not CheckNoSerAvilForThisYear(sOrgCode,iSNo,iSCode,sFinToDt) then
                            Response.Clear 
                            Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Packing Number - "& sNumClassName &"  Classification for this Year </H2></p>"
                            Response.End 
                        end if
                        
                        sPackingNumber = GenSeriesNumber(sOrgCode,iSNo,iSCode,sFinToDt)										 			
                        
							
							Response.Write "Packing Number = "& sPackingNumber 
						End If
						'*****************************************************						
						
						
						
						
						if sRecNum = "S" then iLotValue = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("IVALUE").Value)
						
						if sSellingType = "select" then
							sSellingType = "NULL"
							iWeight = "NULL"
						end if

						if sPackingType = "select" then
							sPackingType = "NULL"
						end if

						if sSellingFormType = "select" then
							sSellingFormType = "NULL"
						end if

						if iLotEntry <> "N/A" and iSerialEntry = "0" then iSerialEntry = "NULL"
						if iLotEntry = "N/A" then
							iLotEntry = "NULL"
						else
							If AutoGen <> "AUTO" Then iLotEntry = Pack(iLotEntry)
						end if

						iQtyGross = cDbl(iQtyRecEntry)
						iQtyNett = cdbl(iQtyRecEntry) - cdbl(iTareQtyEntry)

						if sItemType = "FAB" then iQtyNett = iQtyRecEntry

						If AutoGen = "AUTO" Then 

							with rsObj
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LocationLOT"
								.ActiveConnection = con
								.Open
							end with
							set rsObj.ActiveConnection = nothing

							if not rsObj.EOF then
								iSerialEntry = rsObj(0)
							end if
							rsObj.Close
						
						End If
						
						Response.Write "<P>AUTOGEN = "& AutoGen 

						if sRecNum = "S" then iRate = FormatNumber(cdbl(iLotValue) / cdbl(iQtyNett),4,,,0)
						if Trim(sBinNo)="" or IsNull(sBinNo) or trim(sBinNo)="NULL" then sBinNo = "0"
						if Trim(sSellingType)="" or IsNull(sSellingType) then sSellingType="NULL"
						if Trim(sSellingFormType)="" or IsNull(sSellingFormType) then sSellingFormType = "NULL"
						if Trim(sSubLevelID)="" or IsNull(sSubLevelID) then sSubLevelID = "0"
						if Trim(sNoofCone)="" or IsNull(sNoofCone) then sNoofCone = "0"
						if trim(iLotEntry)="" or IsNull(iLotEntry) then iLotEntry = "NULL"
						
						Response.write "<p>Lot NUmber="& iLotEntry
						
						Response.Write "<p>sInvStatus = " & sInvStatus 
						
						if trim(iStockQuality)="" or IsNull(iStockQuality) then iStockQuality="NULL"
						
						If iWeight = "" Then iWeight = 0		
						if sBinNo = "0" then sBinNo = "NULL"				
						if Trim(sInvStatus)="U" or Trim(sInvStatus)="O" then
						    sSQLQuery = "Update INV_T_LOCATIONLOT set ORGANISATIONCODE=" & Pack(sOrgCode) & "," 
							sSQLQuery =sSQLQuery & " STORAGELOCATIONNO= "& sLocCode & ",STORAGEBINNUMBER=" & sBinNo & ","
							sSQLQuery =sSQLQuery & " LOTQUANTITYGROSS=" & iQtyGross & ",LOTQUANTITYNETT=" & iQtyNett & ","
						    sSQLQuery =sSQLQuery & " LOTQUANTITYTARE=" & iTareQtyEntry & ",PACKINGCODE=" & sPackingType & ",SELLINGNUMBER=" & sSellingType & ","
						    sSQLQuery =sSQLQuery & " WEIGHTPERSELLINGFORM=" & iWeight & ",SELLINGFORM=" & sSellingFormType & ",STAGE=" & sStage & ","
						    sSQLQuery =sSQLQuery & " RATE=" & iRate & ",PACKINGSUBLEVELID="& sSubLevelID &",PACKINGSUBLEVELQTY="& sNoofCone &","
						    sSQLQuery =sSQLQuery & " PACKINGSUBLEVELUNITQTY="& iWeight &",DateOfReceipt=Convert(datetime,'"& dIssDate &"',103),"
						    sSQLQuery =sSQLQuery & " SrcType="& sTraType &",StockQuality="& iStockQuality 
						    sSQLQuery =sSQLQuery & " where INVENTORYRECEIPTNO=" & iInvRecNo & " and ITEMCODE=" & iICode & " and PACKINGNUMBER = "& Pack(sPackingNumber)
						    if trim(iLotEntry)<>"''" and Trim(iLotEntry)<>"NULL" then
						        sSQLQuery = sSQLQuery &" and LOTNUMBER=" & iLotEntry 
						    end if
						    if Trim(sLotAttribute)<>"" and Trim(sLotAttribute)<>"NULL" then
						        sSQLQuery = sSQLQuery &" and AttributeList="& sLotAttribute
						    end if
						    Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
						    con.Execute sSQLQuery
						elseif Trim(sInvStatus)="N" then
						    sSQLQuery = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
							    "CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
							    "LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE,PACKINGSUBLEVELID,PACKINGSUBLEVELQTY,PACKINGSUBLEVELUNITQTY,DateOfReceipt,SrcType,AttributeList,StockQuality) VALUES " &_
							    "(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iICode & "," & sClassCode& "," &_
							    "" & sLocCode & "," & sBinNo & "," & iLotEntry & "," & iSerialEntry & "," &_
							    "" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & "," & Pack(sPackingNumber) & "," &_
							    "" & sPackingType & "," & sSellingType & "," & iWeight & "," & sSellingFormType & "," & sStage & "," & iRate & ","& sSubLevelID &","& sNoofCone &","& iWeight &",Convert(datetime,'"& dIssDate &"',103),"& sTraType &","& sLotAttribute &","& iStockQuality &")"
						    Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
						    con.Execute sSQLQuery
						elseif Trim(sInvStatus)="D" then
						    sSQLQuery = "Delete from INV_T_LOCATIONLOT where INVENTORYRECEIPTNO=" & iInvRecNo & " and ITEMCODE=" & iICode & " and PACKINGNUMBER = "& Pack(sPackingNumber)
						    if trim(iLotEntry)<>"''" and trim(iLotEntry)<>"NULL" then
						        sSQLQuery = sSQLQuery &" and LOTNUMBER=" & iLotEntry 
						    end if
						     if Trim(sLotAttribute)<>"" and trim(sLotAttribute)<>"NULL" then
						        sSQLQuery = sSQLQuery &" and AttributeList="& sLotAttribute
						    end if
						    Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
						    con.Execute sSQLQuery
						end if 
						

						iPieceNo = 0
						if LotSerialNode.Item(iCounter).HaschildNodes() then
							set FabNode = LotSerialNode.Item(iCounter).childNodes(0)
							iAltGross = ""
							iAltNett = ""

							if FabNode.HaschildNodes() then
								arrFab = split(trim(FabNode.Attributes.getNamedItem("QUANTITYIN").Value),":")
								sCheck = trim(FabNode.Attributes.getNamedItem("CHECK").Value)
								iAltGross = trim(FabNode.Attributes.getNamedItem("ALTGROSS").Value)
								iAltNett = trim(FabNode.Attributes.getNamedItem("ALTNETT").Value)

								sQtyIn = arrFab(0)
								iQtyRecEntry = cdbl(arrFab(1))

								if sQtyIn = "N" then
									iQtyGross = iQtyRecEntry + cdbl(iTareQty)
								else
									iQtyGross = iQtyRecEntry
								end if
								
								if sCheck = "Y" then
									sSQLQuery = "UPDATE INV_T_LOCATIONLOT SET LOTQUANTITYGROSS = " & iQtyGross & "," &_
										"LOTQUANTITYTARE = " & iTareQty & " WHERE " &_
										"SERIALNUMBER = " & iSerialEntry & ""
									'objTxt.write sSQLQuery & vbCrLf & vbCrLf
									'Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
									'con.Execute sSQLQuery

									if iAltGross = "" then iAltGross = "0"
									if iAltNett = "" then iAltNett = "0"

									sSQLQuery = "UPDATE INV_T_LOCATIONLOT SET ALTGROSS = " & iAltGross & "," &_
										"ALTNETT = " & iAltNett & " WHERE " &_
										"SERIALNUMBER = " & iSerialEntry & ""
									Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
									con.Execute sSQLQuery

								end if
														
								iAltGross = ""
								iAltNett = ""
								For Each FabDetNode In FabNode.childNodes
									'iPieceNo = trim(FabDetNode.Attributes.getNamedItem("PIECENO").Value)
									iPieceQty = trim(FabDetNode.Attributes.getNamedItem("QUANTITY").Value)
									iAltGross = trim(FabDetNode.Attributes.getNamedItem("ALTGROSS").Value)
									iAltNett = trim(FabDetNode.Attributes.getNamedItem("ALTNETT").Value)

									if iAltGross = "" then iAltGross = "0"
									if iAltNett = "" then iAltNett = "0"
									
									if cdbl(iPieceQty) > 0 then
										iPieceNo = iPieceNo + 1
									'	sSQLQuery = "INSERT INTO INV_T_RECEIPTFABDETAILS (SERIALNUMBER,PIECENO,QUANTITY,ALTGROSS,ALTNETT) VALUES " &_
									'		"(" & iSerialEntry & "," & iPieceNo & "," & iPieceQty & "," & iAltGross & "," & iAltNett & ")"
									'	objTxt.write sSQLQuery & vbCrLf & vbCrLf
									'	'Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
									'	con.Execute sSQLQuery
									end if
								next
							end if
						end if
					next
				next
			else

			'	sSQLQuery = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
			'		"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NOOFPACKS) VALUES " &_
			'		"(" & Pack(sOrgCode) & "," & iICode & "," & sClassCode & "," &_
			'		"" &sTraType & ",NULL,CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iStorageQty & "," & iValue & "," & iNoofPacks & ")"
			'	Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
			'	con.Execute sSQLQuery
				
				sSQLQuery = "Update INV_T_ITEMLEDGER set TRANSACTIONDATE=CONVERT(DATETIME," & Pack(dIssDate) & ",103),TRANSACTQUANTITY="& iStorageQty & ",TRANSACTVALUE=" & iValue & ",NOOFPACKS=" & iNoofPacks 
					sSQLQuery = sSQLQuery & " where ORGANISATIONCODE=" & Pack(sOrgCode) & " and ITEMCODE=" & iICode & " and CLASSIFICATIONCODE" & sClassCode & " and TRANSACTIONTYPE=" &sTraType & " and TRANSACTIONNO =" & iInvRecNo 
					Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
					con.Execute sSQLQuery

			end if
			
			Response.Write "<p> iValue = "& iValue

			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM INV_M_ITEMSTORAGE WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND APPLICABLEFOR = " & Pack(sAppli) & " AND LOCATIONNUMBER = " & sLocCode & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL)"
				.ActiveConnection = con
				.Open
			end with
			Response.Write "<p>"&rsObj.Source 
			set rsObj.ActiveConnection = nothing
			if rsObj.EOF then
				sSQLQuery = "INSERT INTO INV_M_ITEMSTORAGE (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE," &_
					"APPLICABLEFOR,LOCATIONNUMBER,BINNUMBER,ALLOWTRANSFERS) VALUES " &_
					"(" & iICode & "," & sClassCode & "," & Pack(sOrgCode) & "," & Pack(sAppli) & "," &_
					"" & sLocCode & "," & sBinNo & ",'1')"
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			end if
			rsObj.Close

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			arrFinPeriod = split(GetFinancialYear(sMonthYr),":")
			sFinFromDt = arrFinPeriod(0)
			sFinToDt = arrFinPeriod(1)

			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				.ActiveConnection = con
				.Open
			end with
			set rsObj.ActiveConnection = nothing
			if rsObj.EOF then
				sSQLQuery = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
					"(" & Pack(sOrgCode) & "," & sClassCode & "," & iICode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103),CONVERT(DATETIME," & Pack(sFinToDt) & ",103)," & iStorageQty & "," & iValue & "," & iStorageQty & "," & iValue & ")"
				'objTxt.write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			else
			
			    sSQLQuery = "SELECT YEARRECEIPTVALUE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			    rsTempRcpt.open sSQLQuery,con
			    if not rsTempRcpt.eof then
			        Response.Write "<P>rcpt value = "&rsTempRcpt(0)
			        Response.Write "<p>Exist Rcpt Value = "& nExistItemVal 
			        if cdbl(rsTempRcpt(0))>cdbl(nExistItemVal) then
			            Response.Write "<P>Yes"
			        end if 
			    end if
			    rsTempRcpt.close
			    
			    sSQLQuery = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY - " & nExistItemQty & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE - " & nExistItemVal & ")," &_
					"YEARCLOSINGSTOCK = (YEAROPENINGSTOCK + YEARRECEIPTQUANTITY- YEARISSUEQUANTITY), " &_
					"YEARCLOSINGVALUE = (YEAROPENINGVALUE + YEARRECEIPTVALUE- YEARISSUEVALUE) WHERE " &_
					"ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				objTxt.write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			    
				sSQLQuery = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iStorageQty & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
					"YEARCLOSINGSTOCK = (YEAROPENINGSTOCK + " & iStorageQty & "), " &_
					"YEARCLOSINGVALUE = (YEAROPENINGVALUE + " & iValue & ") WHERE " &_
					"ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				objTxt.write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			end if
			rsObj.Close

			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLocCode & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinToDt) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				.ActiveConnection = con
				.Open
			end with

			set rsObj.ActiveConnection = nothing

			if rsObj.EOF then
				sSQLQuery = "INSERT INTO Inv_T_ItemLocationStock (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
					"YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
					"(" & Pack(sOrgCode) & "," & sClassCode & "," & iICode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103),CONVERT(DATETIME," & Pack(sFinToDt) & ",103)," &_
					"" & sLocCode & "," & sBinNo & "," &_
					"" & iStorageQty & "," & iValue & "," & iStorageQty & "," & iValue & ")"
		    	Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
				con.Execute sSQLQuery
			else
			
			sSQLQuery = "UPDATE Inv_T_ItemLocationStock SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY - " & nExistItemQty & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE - " & nExistItemVal & ")," &_
					"YEARCLOSINGSTOCK = (YEAROPENINGSTOCK + YEARRECEIPTQUANTITY- YEARISSUEQUANTITY), " &_
					"YEARCLOSINGVALUE = (YEAROPENINGVALUE + YEARRECEIPTVALUE- YEARISSUEVALUE) WHERE " &_
					"ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"LOCATIONNUMBER = " & sLocCode & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinToDt) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				'objTxt.Write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"&sSQLQuery
				con.Execute sSQLQuery
				
				sSQLQuery = "UPDATE Inv_T_ItemLocationStock SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iStorageQty & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
					"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iStorageQty & "), " &_
					"YEARCLOSINGVALUE = (YEAROPENINGVALUE + " & iValue & ") WHERE " &_
					"ITEMCODE = " & iICode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"LOCATIONNUMBER = " & sLocCode & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
					"CONVERT(DATETIME," & Pack(sFinFromDt) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinToDt) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				'objTxt.Write sSQLQuery & vbCrLf & vbCrLf
				Response.Write "<p>"&sSQLQuery
				con.Execute sSQLQuery
			end if
			rsObj.Close
		next
		
		sSQLQuery = "Update App_T_InternalReceiptHeader set Status = 'Y', InvRecNo="& iInvRecNo &" where InternalReceiptNo = "& iIntRecNo
		Response.write "<p>"& sSQLQuery
		con.Execute sSQLQuery
			
	End Function
%>