<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	IssReturnDetInsert.asp
	'Module Name				:	Inventory
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Jun 24,2011
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include File="../../include/NoSeries.asp" -->
<!--#include File="../../include/PackingCodeSeries.asp" -->
<%
    Dim objDOM,rsIssRet,ndIRRoot,ndIRChild,ndIRLotSer,ndIRLotSerDet,ndStorage
    Dim iRetItemCode,iRetClassCode,iRetIssueNo,iRetItemQty,iAttID
    Dim sQuery,sRetSerialNo,sRcptNumbering,sAttID
    
    dim oDOM1,newxml,ndRoot,ndHeader,LotNode,SerialNode
	dim newxml1,ndRoot1,ndHeader1,LotNode1,SerialNode1
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")
	Set newxml1 = Server.CreateObject("Microsoft.XMLDOM")
	
	dim rsObj,rsObj1,sSQLQuery,iAppCode,iDeptLineNo
	dim iMRSNumber,iICode,sClassCode,sBinNo,sLocCode,iQty,iLot,iSerialNo,dIssDate
	dim sOrgCode,iCreateby,sDept,sSrc,iIssNo,iIntRecNo,iTempMRSNo,iQtyRet,iQtySerial
	dim sSourceRefType, sSourceRefNo,iDINo,sRemark,sType,sTraType,sItemType,sPackWho
	Dim sExpression, iQtyIssued, iQtyReturned, iQtyReturn, iTotReturn
	Dim iIssueEntryNo, WorkCenterNode1, i,WCODE,MachineCenterNode1, j, MCODE
	Dim AutoGen, TypeofStock
	Dim sFinPeriod, sFinPeriodFrom, sFinPeriodTo, sFinFromDt, sFinToDt,sTempMonthYear,arrFinPeriod,sMonthYr
	Dim dtCurrDate
    
    
    set objDOM = Server.CreateObject("Microsoft.XMLDOM")
    Set oDOM1 = Server.CreateObject("Microsoft.XMLDOM")
    set rsIssRet = Server.CreateObject("ADODB.Recordset")
    objDOM.load Server.MapPath("../temp/transaction/IssueReturn"&Session.SessionID&".xml")'	
    
    iRetIssueNo= Request.QueryString("IssueNo")
        
    con.begintrans
    
	dtCurrDate = date()
	
	sFinPeriod = Session("FinPeriod")
	sFinFromDt = "01/04/" & Mid(sFinPeriod,1,4)
	sFinToDt = "31/03/" & Mid(sFinPeriod,6,4)
	
    newxml.load Server.MapPath("../temp/transaction/IssueReturn"&Session.SessionID&".xml")'	
    
	iCreateby = getUserid
	iAppCode = "6"
	AutoGen = ""
	
	Set rsObj = Server.CreateObject("ADODB.RecordSet")
	Set rsObj1 = Server.CreateObject("ADODB.RecordSet")

	Set ndRoot = newxml.documentElement
	Set ndRoot1 = newxml1.documentElement	
	sDept = trim(ndRoot.Attributes.getNamedItem("DEPT").Value)
	sSrc = trim(ndRoot.Attributes.getNamedItem("SOURCE").Value)
	sOrgCode = trim(ndRoot.Attributes.getNamedItem("ORGCODE").Value)
	sType = trim(ndRoot.Attributes.getNamedItem("STYPE").Value)
	sItemType = trim(ndRoot.Attributes.getNamedItem("ITEMTYPE").Value)
	sPackWho = trim(ndRoot.Attributes.getNamedItem("PACKNUM").Value)
	sSourceRefType = trim(ndRoot.Attributes.getNamedItem("SRCREFTYPE").Value)
	sSourceRefNo = trim(ndRoot.Attributes.getNamedItem("SRCREFNO").Value)	
	TypeofStock =  trim(ndRoot.Attributes.getNamedItem("RCPTNUMBERINV").Value)	
	if sType = "T" then
		sTraType = "'RB'"
	else
		sTraType = "'RR'"
	end if
	IF sType = "P" Then sTraType = "'RF'"
	
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
			For Each ndHeader In ndRoot.childNodes
				if StrComp(ndHeader.nodeName,"ITEM") = 0 then
					For Each LotNode in ndHeader.ChildNodes
						dIssDate = trim(LotNode.Attributes.Item(1).nodeValue)
					Next
				End if
			Next	
'			Response.Write "dIssDate = "& dIssDate
		If dIssDate = "" Then 
		
			If Len(Day(date())) = 1 Then
				dIssDate = "0"&Day(date())&"/"
			Else
				dIssDate = Day(date())&"/"
			End If					
			If Len(Month(date())) = 1 Then
				dIssDate = dIssDate & "0"&Month(date())&"/"
			Else
				dIssDate = dIssDate & Month(date())&"/"
			End If		
			dIssDate =dIssDate & Year(date())				
		End If	
		dIssDate = dtCurrDate
'		Response.Write " dISSDATE="& dtCurrDate
		if sSrc = "N" then
			sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTHEADER (INTERNALRECEIPTNO,APPLICATIONCODE," &_
				"ORGANISATIONCODE,CREATEDFROMDEPT,REFTYPE,CREATEDON,CREATEDBY,STATUS,USAGETYPE,USAGESRCREFNO,USAGESRCREFTYPE) VALUES " &_
				"(" & cDbl(iIntRecNo) & "," & iAppCode & "," & Pack(sOrgCode) & "," & Pack(sDept) & "," & Pack(sSrc) & "," &_
				"CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iCreateby & ",'Y','" & sType & "','" & sSourceRefNo & "','" & sSourceRefType & "')"
			Response.Write "<p>"& sSQLQuery
			con.Execute sSQLQuery
			'Response.Write ndRoot.XML
			For Each ndHeader In ndRoot.childNodes
				if StrComp(ndHeader.nodeName,"ITEM") = 0 then
					For Each LotNode in ndHeader.ChildNodes
						dIssDate = trim(LotNode.Attributes.Item(1).nodeValue)
					Next
					'Response.Write dIssDate
					sClassCode = trim(ndHeader.Attributes.getNamedItem("CLACODE").Value)
					iICode = trim(ndHeader.Attributes.getNamedItem("ITMCODE").Value)
					iQty = trim(ndHeader.Attributes.getNamedItem("QTY").Value)
					iTempMRSNo = trim(ndHeader.Attributes.getNamedItem("MRSNO").Value)
					iIssNo = trim(ndHeader.Attributes.getNamedItem("ISSNO").Value)
					iAttID = trim(ndHeader.Attributes.getNamedItem("ATTID").Value)
					sAttID = iAttID

					If iIssNo <> "N" Then
						iIssueEntryNo = iIssNo
    				End If
					if iTempMRSNo = "N" then
						iMRSNumber = "NULL"
					else
						iMRSNumber = iTempMRSNo
					end if
					
					if trim(iAttID)="" or IsNull(iAttID) or trim(iAttID)="0" then iAttID="NULL"
					if iAttID <>"NULL" then iAttID = Pack(iAttID)
					

					sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
						"CLASSIFICATIONCODE,ITEMCODE,QUANTITYRETURN,AttributeList) VALUES " &_
						"(" & iIntRecNo & "," & iMRSNumber & "," & sClassCode & "," & iICode & "," &_
						"" & iQty & ","& iAttID &")"
					Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
					con.Execute sSQLQuery
				' end if for Item Node check
				    InsertLotDetails iICode,sClassCode,sAttID
				end if
			next
			'Calling Function to insert Lot/Serial Details
		else 'MRS / Direct Issue Based
		
			sSQLQuery = "INSERT INTO APP_T_INTERNALRECEIPTHEADER (INTERNALRECEIPTNO,APPLICATIONCODE," &_
				"ORGANISATIONCODE,CREATEDFROMDEPT,REFTYPE,CREATEDON,CREATEDBY,STATUS) VALUES " &_
				"(" & iIntRecNo & "," & iAppCode & "," & Pack(sOrgCode) & "," & Pack(sDept) & "," & Pack(sSrc) & "," &_
				"CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iCreateby & ",'N')"
			Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf
			con.Execute sSQLQuery
			
			For Each ndHeader In ndRoot.childNodes
				if StrComp(ndHeader.nodeName,"ITEM") = 0 then
					sClassCode = trim(ndHeader.Attributes.getNamedItem("CLACODE").Value)
					iICode = trim(ndHeader.Attributes.getNamedItem("ITMCODE").Value)
					iQty = trim(ndHeader.Attributes.getNamedItem("QTY").Value)
					iMRSNumber = trim(ndHeader.Attributes.getNamedItem("MRSNO").Value)
					iDINo = trim(ndHeader.Attributes.getNamedItem("DINO").Value)
					iIssNo = trim(ndHeader.Attributes.getNamedItem("ISSNO").Value)
					sRemark = trim(ndHeader.Attributes.getNamedItem("REMARKS").Value)
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
					sExpression ="//ITEM [ @CLACODE = "&sClassCode&" and @ITMCODE = "&iICode&"]/AddDet/WorkCenter"
					Set WorkCenterNode1 = ndRoot1.Selectnodes(sExpression)
					'Response.Write WorkCenterNode1.Length
					If WorkCenterNode1.Length > 0 Then
					For i = 0 to WorkCenterNode1.Length-1
						WCODE = trim(WorkCenterNode1.Item(i).Attributes.getNamedItem("WCODE").Value)		
						'Response.Write WCODE					
						sExpression ="//ITEM [ @CLACODE = "&sClassCode&" and @ITMCODE = "&iICode&"]/AddDet/WorkCenter [ @WCODE = '"&WCODE&"']/MachineCenter"
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

					For Each LotNode in ndHeader.ChildNodes
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
	end if

          
        set ndIRRoot = objDOM.documentElement
        if ndIRRoot.hasChildNodes() then
            sRcptNumbering = ndIRRoot.getAttribute("RCPTNUMBERINV")
            Response.Write "<p>sRcptNumbering = "& sRcptNumbering
            For Each ndIRChild in ndIRRoot.childNodes
                if ndIRChild.nodeName="ITEM" then
                    iRetItemCode = ndIRChild.getAttribute("ITMCODE")
                    iRetClassCode = ndIRChild.getAttribute("CLACODE")
                    For Each ndStorage in ndIRChild.childNodes
                        For each ndIRLotSer in ndStorage.childNodes
                            if ndIRLotSer.hasChildNodes() then
                                For Each ndIRLotSerDet in ndIRLotSer.childNodes
                                    if ndIRLotSerDet.getAttribute("QTYREC")<>"" and ndIRLotSerDet.getAttribute("QTYREC")<>"0" then
                                        sRetSerialNo = ndIRLotSerDet.getAttribute("LOTSERIAL")
                                        iRetItemQty = ndIRLotSerDet.getAttribute("QTYREC")
                                        sQuery = "Update INV_T_MaterialIssueDetails set QuantityReturned ="& iRetItemQty  &" where ItemCode = "&iRetItemCode&" and ClassificationCode = " &iRetClassCode&" and IssueEntryNo ="& iRetIssueNo &" and SerialNo = "& sRetSerialNo 
                                        Response.Write "<p>"& sQuery
                                        con.execute sQuery
                                    end if
                                Next
                            else
                                if Trim(sRcptNumbering)="N" then
                                    iRetItemQty= ndIRLotSer.getAttribute("QTY")
                                    sQuery = "Update INV_T_MaterialIssueDetails set QuantityReturned ="& iRetItemQty  &" where ItemCode = "&iRetItemCode&" and ClassificationCode = " &iRetClassCode&" and IssueEntryNo ="& iRetIssueNo 
                                    
                                    Response.Write "<p>"& sQuery
                                    con.execute sQuery
                                end if
                            end if
                        Next
                   Next 'For Each ndStorage in ndIRChild.childNodes
                end if
            Next
        end if
        
        
    if con.Errors.count <> 0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
	else
	'	con.RollbackTrans
	'	Response.End
	   Response.Clear
	   con.CommitTrans
	   Response.Redirect "ISSUEMGMT.ASP?ACTN=L"
	end if 'if con.Errors.count <> 0 then
%>



<%	'Function to insert the Lot/Serial Details for None Cases
	Function InsertLotDetails(sITMCode,sCLACode,sAttributeID)
		dim ItemNode,StorageNode,LotNode,LotSerialNode,objfs,FabNode,FabDetNode,OptionalUoMNode
		dim iNoofPacks, PackNode
		dim sTemp,iCtr,sLocCode,sBinNo,sMonthYr,iQty,iValue,sAppli
		dim arrFinPeriod,arrTemp,iInvRecNo,iCounter,sTempMonthYear
		dim sQtyIn,iTareQty,iLot,iSerialFrom,iSerialTo,sSellingType,iWeight,sPackingType,sPackingNumber
		dim iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sTareIn,iQtyGross,iQtyNett
		dim iSeriesNo,iSeriesCode,sProductCode,iItmRate,sItmShDesc,sItmAddDesc
		dim iLotCtr,sSellingFormType,sStage,iCodeLen,iCodeSize,sForm
		dim sDrwVerNo,sCatalogue,iCommodity,iLotQty,iLotValue,iEntryCtr,sReceivedOn
		dim iFabCtr,iPieceNo,iPieceQty,arrFab,iAltGross,iAltNett,sAltUoM,sCheck,iRate
		dim sOpUoMCode,iOpUoMFactor,sOpUoMOperator,sRecNum,iStorageQty
		dim objFSO,objTxt,sExpression,sExpression1,iDefinedBy,arrDate
		Dim iSNo, iSCode, sPackCode 
		
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		iDefinedBy = getUserid
		
		iNoofPacks = "Null"
		
		with rsObj
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(COMPANYITEMCODE,'') FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iICode & ""
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
		
		
		'This Part Added For the purpose of to add No of Packs in Ledger Entries
		sExpression ="//ROOT/ITEM[@ITMCODE="& sITMCode &" and @CLACODE="& sCLACode &" and @ATTID='"& sAttributeID &"']/STORAGE/LotSerial/LotSerialDetails"
		Set PackNode = ndRoot.Selectnodes(sExpression)
		If PackNode.Length > 0 Then
			iNoofPacks = PackNode.Length
		End If
		
		iCtr = 0
		sExpression ="//ROOT/ITEM[@ITMCODE="& sITMCode &" and @CLACODE="& sCLACode &" and @ATTID='"& sAttributeID &"']/STORAGE"
		Response.Write "<p>sExpression = "& sExpression 
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
			arrDate = split(sReceivedOn,"/")
			sMonthYr = arrDate(1)&arrDate(2)
			'Response.Write sMonthYr
			dIssDate = sReceivedOn 
			if sBinNo = "0" then sBinNo = "NULL"
			
			sSQLQuery  = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
			rsObj.Open sSQLQuery,con	
			if not rsObj.EOF then
				iInvRecNo = rsObj(0)
			end if
			rsObj.Close
			
			if sRecNum = "N" then
			    sSQLQuery = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
							"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTQUANTITYGROSS," &_
							"LOTQUANTITYNETT) VALUES " &_
							"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iICode & "," & sClassCode& "," &_
							"" & sLocCode & "," & sBinNo & "," & iStorageQty & "," & iStorageQty & ")"
						Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
						con.Execute sSQLQuery
			end if
			
			sQuery = " Update APP_T_InternalReceiptHeader set InvRecNo ="& iInvRecNo &" where InternalreceiptNo = "& iIntRecNo
			Response.Write "<p>"& sQuery
			con.execute sQuery
			
			

			sExpression1 ="//ROOT/ITEM[@ITMCODE="& sITMCode &" and @CLACODE="& sCLACode &" and @ATTID='"& sAttributeID &"']/STORAGE [ @STORE = "&sLocCode&" and @BIN = '"&sBinNo&"']/LotSerial[@QTY>0]"
			Response.Write "<p>"&sExpression1 
			Set LotNode = ndRoot.Selectnodes(sExpression1)
			Response.Write "<p>LotNode.Length  = "&  LotNode.length
			if sBinNo = "NULL" or IsNull(sBinNo) then  sBinNo = 0 

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
				'	Response.Write AutoGen 
					If AutoGen = "AUTO" and (TypeofStock = "LS" or TypeofStock = "L") Then
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
						with rsObj
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'LO' AND ITEMTYPE = '" & sItemType & "' AND ORGANISATIONCODE = '" & sOrgCode & "'"
							.ActiveConnection = con
							.Open
						end with
						set rsObj.ActiveConnection = nothing
						'Response.Write rsObj.Source 
						if not rsObj.EOF then
							iSNo = trim(rsObj(0))
							iSCode = trim(rsObj(1))
							'Response.Write sOrgCode& "," & iSNo& ","& iSCode & "," &sFinFromDt 
							iLotEntry = GenSeriesNumber(sOrgCode,iSNo,iSCode,sFinFromDt)	
							iLotEntry = Pack(iLotEntry) 	
						end if
						rsObj.close
					End If
					If TypeofStock = "S" Then iLotEntry = "Null"
					'Response.Write ":Lot No :"& iLotEntry &":"
			'---------------------------
					with rsObj
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
						.ActiveConnection = con
						.Open
					end with
					set rsObj.ActiveConnection = nothing
					'Response.Write rsObj.Source 
					if not rsObj.EOF then
						iInvRecNo = rsObj(0)
					end if
					rsObj.Close
					
					sQuery = " Update APP_T_InternalReceiptHeader set InvRecNo ="& iInvRecNo &" where InternalreceiptNo = "& iIntRecNo
					Response.Write "<p>"& sQuery
					con.execute sQuery


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

					sSQLQuery = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
						"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NOOFPACKS,AttributeList) VALUES " &_
						"(" & Pack(sOrgCode) & "," & iICode & "," & sClassCode & "," &_
						"" &sTraType & "," & iInvRecNo & ",CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iQty & "," & iLotValue & "," & iNoofPacks & ","& iAttID &")"
					Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
					con.Execute sSQLQuery


					if sBinNo = "0" or IsNull(sBinNo) then  sBinNo ="NULL"
					sExpression1 ="//ROOT/ITEM[@ITMCODE="& sITMCode &" and @CLACODE="& sCLACode &" and @ATTID='"& sAttributeID &"']/STORAGE [ @STORE = "&sLocCode&" and @BIN = '"&sBinNo&"']/LotSerial [ @LOT = '"&iLot&"' and @COUNTER = "&iEntryCtr&"]/LotSerialDetails[@QTYREC>0]"
					Response.Write "<p>"& sExpression1 
					Set LotSerialNode = ndRoot.Selectnodes(sExpression1)
					if sBinNo = "NULL" or IsNull(sBinNo) then  sBinNo = 0 

					For iCounter = 0 to LotSerialNode.Length - 1
						If AutoGen <> "AUTO" Then iLotEntry = iLot
						iSerialEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("LOTSERIAL").Value)
						iQtyRecEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("QTYREC").Value)
						iTareQtyEntry = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("TAREREC").Value)

						sSellingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGTYPE").Value)
						iWeight = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("WEIGHTSTYPE").Value)
						sPackingType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKINGTYPE").Value)

						sSellingFormType = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("SELLINGFORM").Value)
						if sPackWho <> "N" then
							sPackingNumber = trim(LotSerialNode.Item(iCounter).Attributes.getNamedItem("PACKNUMBER").Value)
						elseif sPackWho = "N" then
							sPackingNumber = PackingCodeSeries("4", "4", sProductCode, iICode, sClassCode, sOrgCode, sPackingType, sItemType, "I")
						end if
						
						'  To Create Pack Series Code  ************************
						If sTraType = "'RB'" or sTraType = "'RF'" Then
							with rsObj
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT SERIESNO,SERIESCODE FROM PRD_M_PackingNumberSeries WHERE ITEMTYPEID = '" & sItemType & "' AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
								.ActiveConnection = con
								.Open
							end with
							Response.Write rsObj.Source 
							set rsObj.ActiveConnection = nothing
							if Not rsObj.EOF then
								iSNo = rsObj(0)
								iSCode = rsObj(1)
							end if
							rsObj.close
							sPackingNumber = GenSeriesNumber(sOrgCode,iSNo,iSCode,dIssDate)										 			
							'Response.Write sPackingNumber 
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
						
						If iWeight = "" Then iWeight = 0						
						sSQLQuery = "INSERT INTO INV_T_LOCATIONLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
							"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER,LOTQUANTITYGROSS," &_
							"LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE) VALUES " &_
							"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iICode & "," & sClassCode& "," &_
							"" & sLocCode & "," & sBinNo & "," & iLotEntry & "," & iSerialEntry & "," &_
							"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & "," & Pack(sPackingNumber) & "," &_
							"" & sPackingType & "," & sSellingType & "," & iWeight & "," & sSellingFormType & "," & sStage & "," & iRate & ")"
						Response.Write "<p>"& sSQLQuery & vbCrLf & vbCrLf & vbCrLf
						con.Execute sSQLQuery

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
					"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NOOFPACKS,AttributeList) VALUES " &_
					"(" & Pack(sOrgCode) & "," & iICode & "," & sClassCode & "," &_
					"" &sTraType & ",NULL,CONVERT(DATETIME," & Pack(dIssDate) & ",103)," & iStorageQty & "," & iValue & "," & iNoofPacks & ","& iAttID &")"
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
	End Function
%>
  