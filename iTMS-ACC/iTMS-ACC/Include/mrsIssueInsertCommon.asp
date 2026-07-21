<%
	'Program Name				:	mrsIssueInsertCommon.asp
	'Module Name				:	Include
	'Author Name				:	RAGAVENDRAN R
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
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

<!--#include file="MarkDetailsNew.asp"-->
<!--#include file="mrsStatus.asp"-->
<!--#include file="NoSeries.asp"-->
<!--#include file="NoSeriesCommonFunctions.asp"-->
<%
'XML DOM Variables
    Dim oDOM,Root,objfs,RootO,subContRoot,subContItemNode,subContScheduleNode
    dim OutData,subContDOM,rsTemp,OutDataMix
    dim Issxml,RootNode,HeaderNode,PickNode,PickDetNode,ndMixData,ndMix
    dim ScheduleNode,ScheduleDetNode,SerialHNode,SerialDetNode,sAttList
    dim WorkCenterNode, MachineCenterNode,sCallFrom ,sItemRefNo,sArrAppRefNo,sArrTemp
    dim dcrs,dcrs1,sSql,bFlag,nTransactQty
    dim iMRSNo,iItemCode,iClass,arrStore,sBin,sLoc,iReqQty,iIssQty,iEntNo,iIssEntNo,sAttributeList
    dim iTraQty,iPrQty,iValue,sOrgID,iIssuedBy,dMRSDate,sMonYr,sMethod,iRecQty
    dim arrFin,sFinFrom,sFinTo,sTempMonYr,iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue
    dim iWMQty,iWMRecQty,iWMIssQty,iTempWMQty,sIssType,sDeptNo,sRem
    dim iPickNo,sPickLoc,sPickBin,sPickLot,sPickQty,sSTOrgID,iSTQty,sTotPickQty
    dim iSchQty,iSchNo,iSchPickQty,iLineNo,sSchOn,sReqValue,sTempDate,iStockNo
    dim sExp,sSchType,sSchValue,iSerial,iSerialQty,iArrCtr,iTotVal
    dim dicSerial,arrSerial,arrSerialQty,iTransferNo,iIssueNo,sRecBy,sPartyCode
    dim iLedIssueNo, sItmType,iSeriesNo,iSeriesCode,iGenCode,iLedgEntNo,iItemQty
    Set dicSerial = Server.CreateObject("Scripting.Dictionary")
    Dim iIsQty,iTqty,IssDate,iForOrderNo
    Dim IssToType,IssToCode,IssToSubCode
    Dim iDeptEntryNo ,iCreatedBy,dCreatedOn
    Dim subPickNode,ItemDetNode,AddDetNode,sUoM,sInvTypeName
    Dim sWCode,sMCode,iMCQty,sMCName,HNode,sPurType,sPOConfirm,sSqlTemp,sTotQty
    Dim sSInvType,sSSalType,sInvNo,iNoofCases,sInvItemQty,sSalUoM,sSalInvConfirm
    Dim sDISCode,iMixCode,iMixQty
    Dim sTransPort,sTakenBy,sDeliveryBy,sRemarks,sDCNO,sGatePassNo,sGatePassEntryNo,sGatePassConfirm
    Dim sSelectedInvoice,sProformaConfirm,sSubConProfoma,sArrList,sAttID
    Dim sModuleCallFrom,sRedirectTo,sAppRefType,sAppRefNo,sAppRefDate
    Dim iLotNo,iQtyIss,SerNode,SerDetNode,iSerNo,iSerQty,iIssVal,iIssAccHead,iItmRate,ObjFs1
    Dim dtDate
    Dim RootTemp,Tempnode,sSelectedWC,nWCLocNo,nWCBinNo,dRSet,rsNew,sMonthYear,sMixCode
    Dim sStockEntryExistForCurrentWC,sCurrentMonthFirstDate,sLastEntryMonthYear
    Dim sListOfMonthYearForWIPStockEntryNotExist,sTempMonth,sFirstDate,sSqlCmd,sTransactUOM
    Dim objSubDOM,subAddRoot,ndAddDet,sAutoConsumption,sConAccHead,sConEntryNo,sConLineNo,sEntryNo,sIssueToCode,sIssuePOSID
    
    Dim sIssueExp,ndIssueTempNode,iNumIssueClassCode,sTempSeries,sArrSeries,sNumClassName,sSALPOSID
    Dim sSInvTypeName,sSSalTypeName,sSSALPOSIDName,sPickPackFlag,sOnlyLotFlag,sTempLotNo,sTempAttList,sIssFrom
    Dim iScheduleNo,dSchedule,iScheduleQty,sReturnable,sReturnItem,sType,sIReturnable,sIReturnItem
    Dim sProcessID,sRetItemCode,sRetClassCode,sInstruct,sLabCharge,sMatType
    Dim sHardWaste,sInvWaste,iReturnItemEntryNo,SubNode,sReturnItemCode,sReturnClassCode
    Dim iIssInvRecNo,ndItemNode,iCnt 
    Dim ArrMonth,nTempCtr
    Dim subContProcessHead,subContProcessItem

    set dRSet = Server.CreateObject("ADODB.Recordset")
    dTDate = IssDate
    Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
    Set objfs = CreateObject("Scripting.FileSystemObject")
    set ObjFs1 = CreateObject("Scripting.FileSystemObject")
    set rsTemp = Server.CreateObject("ADODB.RecordSet")
    set rsNew = Server.CreateObject("ADODB.RecordSet")


    set IssXML = Server.CreateObject("Microsoft.XMLDOM")

    set OutData = Server.CreateObject("Microsoft.XMLDOM")
    set OutDataMix = Server.CreateObject("Microsoft.XMLDOM")
    set subContDOM = server.CreateObject("Microsoft.XMLDOM")
    set objSubDOM = Server.CreateObject("Microsoft.XMLDOM")
    

	sCallFrom = Request.QueryString("CallFrom")

    Set dcrs = Server.CreateObject("ADODB.RecordSet")
    Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
    ''Response.Write Server.MapPath("../Temp/Transaction/mrsIssueData"&Session.SessionID&".xml")

    IssXML.load(Server.MapPath("../Temp/Transaction/mrsIssueData"&Session.SessionID&".xml"))
    

Function MrsIssueInsert()
    Set RootO = OutData.createElement("Root")
    OutData.appendChild RootO
	
	'First time XML is Not Loaded.So,XML Load is Put inside the Function
	IssXML.load(Server.MapPath("../Temp/Transaction/mrsIssueData"&Session.SessionID&".xml"))
	
	set subContRoot = subContDOM.createElement("ROOT")
    subContDOM.appendChild subContRoot

if ObjFs1.FileExists(Server.MapPath("../Temp/Transaction/SubContract"&Session.SessionID&".xml")) then
	objSubDOM.load(Server.MapPath("../Temp/Transaction/SubContract"&Session.SessionID&".xml"))
	
	set subAddRoot = objSubDOM.documentElement

    if subAddRoot.hasChildNodes() then
        for each ndAddDet in subAddRoot.childNodes
            subContRoot.appendChild ndAddDet
        next
    end if
end if  'if ObjFs1.FileExists(Server.MapPath("../Temp/Transaction/SubContract"&Session.SessionID&".xml"))
    
    ' Create our DOM Document Objects
    iDeptEntryNo = 0
    iIsQty = 0
    iTqty = 0
    iIssuedBy = getUserid
    bFlag = true


    Set RootNode = IssXML.documentElement
	'Response.Write RootNode.xml
	
	sIssueExp = "//ISSTYPE/ITEM"
	set ndIssueTempNode = RootNode.selectNodes(sIssueExp)
	if ndIssueTempNode.Length>0 then
	    iNumIssueClassCode = ndIssueTempNode.Item(0).Attributes.getNamedItem("CLACODE").Value
	    'Response.Write "<p> ClassCode = "& iNumIssueClassCode
	end if
	
    sIssType = trim(RootNode.Attributes.getNamedItem("ISSTYPE").Value)
    IssToType =trim(RootNode.Attributes.getNamedItem("ISSTOTYPE").Value)
    IssToCode =trim(RootNode.Attributes.getNamedItem("ISSTOCODE").Value)
    IssToSubCode=trim(RootNode.Attributes.getNamedItem("ISSTOSUBCODE").Value)
    sPOConfirm = trim(RootNode.Attributes.getNamedItem("POConfirm").Value)
    sSalInvConfirm = trim(RootNode.Attributes.getNamedItem("SInvConfirm").Value)
    sSelectedInvoice = trim(RootNode.Attributes.getNamedItem("Invoice").Value)
    sGatePassConfirm = trim(RootNode.Attributes.getNamedItem("GPConfirm").Value)
    sProformaConfirm = trim(RootNode.Attributes.getNamedItem("ProConfirm").Value)
    sModuleCallFrom = trim(RootNode.Attributes.getNamedItem("MCallFrom").Value)
    sRedirectTo = trim(RootNode.Attributes.getNamedItem("RedirectTo").Value)
    sAppRefType = trim(RootNode.Attributes.getNamedItem("AppRefType").Value)
    sAppRefNo = trim(RootNode.Attributes.getNamedItem("AppRefNo").Value)
    sAppRefDate = trim(RootNode.Attributes.getNamedItem("AppRefDate").Value)
    sConAccHead = trim(RootNode.Attributes.getNamedItem("ConsumptionAccHead").Value)
    sIssueToCode = trim(RootNode.Attributes.getNamedItem("IssueToCode").Value)
    sPickPackFlag = trim(RootNode.Attributes.getNamedItem("PickPackFlag").Value)
    sIssFrom = trim(RootNode.Attributes.getNamedItem("IssFrom").value)
    sReturnable = trim(RootNode.Attributes.getNamedItem("Returnable").value)
    sReturnItem =  trim(RootNode.Attributes.getNamedItem("ReturnItem").value)
    sType = trim(RootNode.Attributes.getNamedItem("TYPE").value)
    
    
			    
			    
    
    if trim(sPickPackFlag)="" or isnull(sPickPackFlag) then sPickPackFlag = "L"
    
    Response.Write "<P>IssToType ="&  IssToType
    Response.Write "<P>IssToCode ="&  IssToCode
    
    if ucase(IssToType)=ucase("PARTY") then
        sPartyCode = IssToCode
        Response.Write "<p>PartyCode = "& sPartyCode
    end if 
    
    
    if trim(IssToType)="" or isNull(IssToType) then IssToType = "NULL"
    if trim(IssToCode)="" or isNull(IssToCode) then IssToCode="NULL"
    if trim(IssToSubCode)="" or isNull(IssToSubCode) then IssToSubCode="NULL"
    
    if trim(IssToType)<>"" and trim(IssToType)<>"NULL" then IssToType = Pack(IssToType)
    'if trim(IssToCode)<>"" and trim(IssToCode)<>"NULL" then IssToCode = Pack(IssToCode)
    if trim(IssToSubCode)<>"" and trim(IssToSubCode)<>"NULL" then IssToSubCode = Pack(IssToSubCode)
    
    
    ''Response.Write "sAppRefDate = "& sAppRefDate
    
    if Trim(sConAccHead)="" or IsNull(sConAccHead) then sConAccHead ="NULL"

    if trim(sAppRefType)="" or trim(sAppRefType)="Select"  then sAppRefType ="NULL"
    if trim(sAppRefNo)="" then sAppRefNo ="NULL"
    if trim(sAppRefNo)<>"NULL" then sAppRefNo = Pack(sAppRefNo)
    if trim(sAppRefDate)="" then sAppRefDate =FormatDate(Date)

    sGatePassEntryNo = 0

    if trim(IssToCode)="SUB" then
	    sSubConProfoma = "Y"
    end if
    
    sSql = "Select isNull(AutomaticConsumptionEntry,'N') from APP_M_ApplicationSetup"
    dcrs1.open sSql,con
    if not dcrs1.eof then
        sAutoConsumption = dcrs1(0)
    end if
    dcrs1.close 
    
  

    if trim(sPartyCode)="" or IsNull(sPartyCode) then sPartyCode="NULL"
    ''Response.Write RootNode.xml
    ''Response.Write vbCrLf & vbCrLf
    if RootNode.HaschildNodes() then
    For Each HeaderNode In RootNode.childNodes
	    if StrComp(HeaderNode.nodeName,"ITEM") = 0 then
		    sOrgID = trim(HeaderNode.Attributes.getNamedItem("ORGCODE").Value)
		    iMRSNo = trim(HeaderNode.Attributes.getNamedItem("MRSNO").Value)
		    sRecBy  = trim(HeaderNode.Attributes.getNamedItem("REQBY").Value)
		    sRem	= trim(HeaderNode.Attributes.getNamedItem("REMARKS").Value)
		    sItmType = trim(HeaderNode.Attributes.getNamedItem("ITEMTYPE").Value)
		    iItemQty =  trim(HeaderNode.Attributes.getNamedItem("ISSQTY").Value)
		    if iItemQty="" or IsNull(iItemQty) then iItemQty = "0"
		    iIsQty = iIsQty + cDbl(iItemQty)
		    iTqty = iTqty + cdbl(HeaderNode.Attributes.getNamedItem("TRAQTY").Value)
		    IssDate = HeaderNode.Attributes.getNamedItem("ISSUEDATE").Value
		    iCreatedBy = HeaderNode.Attributes.getNamedItem("CREATEDBY").Value
		    dCreatedOn = HeaderNode.Attributes.getNamedItem("CREATEDON").Value
		    sAttList = trim(HeaderNode.Attributes.getNamedItem("ATTRIBUTELIST").Value)
		    ''Response.Write "sAttList ="& sAttList 
		    if trim(sAttList)<>"" then
				sArrTemp=split(sAttList,":")
				sArrList = split(sArrTemp(0),"#")
					if sArrList(0)<>"0" then
						if UBound(sArrList)>0 then
							if trim(sArrList(1))<>"0" and trim(sArrList(1))<>"" then
								sAttID =  sArrList(1)
							end if
						else
							if trim(sArrList(0))<>"0" and trim(sArrList(0))<>"" then
								sAttID = sArrList(0)
							end if
						end if
					end if
			end if'if trim(sAttList)<>"" then
			''Response.Write "sAttID = "& sAttID 

            If trim(sAttID)="" or IsNull(sAttID) then sAttID = "NULL"

		    for each PickNode in HeaderNode.childNodes
			    if strcomp(PickNode.nodeName,"Pick")=0 then
				    iNoofCases = PickNode.getAttribute("NoofPack")
			    end if
		    next

		   ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
			    if trim(sType)="SUB" then

				    set subContItemNode = subContDOM.createElement("ITEMDETAILS")
					    subContItemNode.setAttribute "ENTRYNO",trim(HeaderNode.Attributes.getNamedItem("ENTRYNO").Value)
					    subContItemNode.setAttribute "ITEMCODE",trim(HeaderNode.Attributes.getNamedItem("ITMCODE").Value)
					    subContItemNode.setAttribute "CLASSCODE",trim(HeaderNode.Attributes.getNamedItem("CLACODE").Value)
					    subContItemNode.setAttribute "UNIT",trim(HeaderNode.Attributes.getNamedItem("ORGCODE").Value)
					    subContItemNode.setAttribute "ITEMNAME",trim(HeaderNode.Attributes.getNamedItem("ITMNAME").Value)

						    with dcrs1
							    .CursorLocation = 3
							    .CursorType = 3
							    .ActiveConnection = con
							    .Source = "Select PurchaseUoM from INV_M_ItemMaster where ItemCode = "&trim(HeaderNode.Attributes.getNamedItem("ITMCODE").Value) & " and ClassificationCode = "& trim(HeaderNode.Attributes.getNamedItem("CLACODE").Value)
							    .open
						    end with
						    if not dcrs1.eof then
						    sUoM = dcrs1(0)
						    end if
						    dcrs1.close

					    if trim(sUoM)="" or IsNull(sUoM) then
						    subContItemNode.setAttribute "UOM",""
					    else
						    subContItemNode.setAttribute "UOM",sUOM
					    end if


					    subContItemNode.setAttribute "DECIMAL","Y"
					    subContItemNode.setAttribute "DISPALYED","Y"
					    subContItemNode.setAttribute "PACKING","N"
					    subContItemNode.setAttribute "QTY",trim(HeaderNode.Attributes.getNamedItem("ISSQTY").Value)
					    subContItemNode.setAttribute "RATE","0"
					    subContItemNode.setAttribute "UNITPER","select"
					    subContItemNode.setAttribute "UNITRATE","0"
					    subContItemNode.setAttribute "DISCOUNT","0"
					    subContItemNode.setAttribute "VALUE","0"
					    subContItemNode.setAttribute "REQUIREDBY","I"
					    subContItemNode.setAttribute "REQUIREDVALUE",""
					    subContItemNode.setAttribute "TEMPICODE",""
					    subContItemNode.setAttribute "TEMPISHDESC",""
					    subContItemNode.setAttribute "TEMPIADDDESC",""
					    subContItemNode.setAttribute "APPCODE",""
					    subContItemNode.setAttribute "MODCODE",""
					    subContItemNode.setAttribute "CREATIONSTAGE",""
					    subContItemNode.setAttribute "CATEGORY",""
					    subContItemNode.setAttribute "ITEMDESC",""
					    subContItemNode.setAttribute "ATTRIBUTELIST",sAttID
					    subContItemNode.setAttribute "ADDNDESCRIPTION",""
					    subContItemNode.setAttribute "ItemRate","0"
					    subContItemNode.setAttribute "MarketRate","0"
					    subContItemNode.setAttribute "RETURNABLE",trim(HeaderNode.Attributes.getNamedItem("RETURNABLE").Value)
					    subContItemNode.setAttribute "RETURNITEM",trim(HeaderNode.Attributes.getNamedItem("RETURNITEM").Value)
				    subContRoot.appendChild subContItemNode

				    set subContScheduleNode = subContDOM.createElement("Schedule")
				    subContScheduleNode.setAttribute "STYPE","I"
				    subContScheduleNode.setAttribute "SVALUE",trim(HeaderNode.getAttribute("CREATEDON"))
				    subContScheduleNode.setAttribute "ITEMCODE",trim(HeaderNode.getAttribute("ITMCODE"))
				    subContScheduleNode.setAttribute "CLASSCODE",trim(HeaderNode.getAttribute("CLACODE"))
				    subContScheduleNode.setAttribute "ENTRYNO",trim(HeaderNode.getAttribute("ENTRYNO"))
				    subContItemNode.appendChild subContScheduleNode
			    end if ' if IssToCode ="SUB" then
		   ' end if ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then

	    elseif strcomp(HeaderNode.nodeName,"PURACC")=0 then
		    sPurType = HeaderNode.getAttribute ("PurType")
	    elseif strcomp(HeaderNode.nodeName,"SALINV")=0 then
		    sSInvType = HeaderNode.getAttribute("InvType")
		    sSSalType = HeaderNode.getAttribute("SalType")
		    sSALPOSID = HeaderNode.getAttribute("POS")
		    sSInvTypeName = HeaderNode.getAttribute("InvTypeName")
		    sSSalTypeName = HeaderNode.getAttribute("SalTypeName")
		    sSSALPOSIDName = HeaderNode.getAttribute("POSName")
	    elseif strcomp(HeaderNode.nodeName,"SERVICES")=0 then
		    sTransPort = HeaderNode.getAttribute("Transport")
		    sTakenBy   = HeaderNode.getAttribute("TakenBy")
		    sDeliveryBy= HeaderNode.getAttribute("DelivertyBy")
		    sRemarks   = HeaderNode.getAttribute("Remarks")
	    End If
    Next
    
    if sSInvType ="CB" then sSInvType ="X"
    if sSInvType ="NEB" then sSInvType ="Y"
    if sSInvType ="EB" then sSInvType ="Z"

    if trim(sTransPort)="" or IsNull(sTransPort) then
	    sTransPort ="NULL"
    else
	    sTransPort =pack(sTransPort)
    end if
    if trim(sTakenBy)="" or IsNull(sTakenBy) then
	    sTakenBy = "NULL"
    else
	    sTakenBy = pack(sTakenBy)
    end if

    if trim(sDeliveryBy)="" or IsNull(sDeliveryBy) then
	    sDeliveryBy = "NULL"
    else
	    sDeliveryBy = pack(sDeliveryBy)
    end if

    if trim(sRemarks)="" or IsNull(sRemarks) then
	    sRemarks = "NULL"
    else
	    sRemarks = pack(sRemarks)
    end if
    
       If Cdbl(iIsQty+iTqty) > 0 and sIssType = "F" Then
		    with dcrs
			    .CursorLocation = 3
			    .CursorType = 3
			    '.Source = "SELECT ISNULL(MAX(ISSUENO)+1,1) FROM INV_T_DEPARTMENTSTOCK "
			    .source = "Select IsNull(max(IssueEntryNo),0)+1 from INV_T_MaterialIssueHeader"
			    .ActiveConnection = con
			    .Open
		    end with
		    set dcrs.ActiveConnection = nothing

		    if not dcrs.EOF then
			    iLedIssueNo = dcrs(0)
		    end if
		    dcrs.Close
	    End If

    if trim(sPartyCode)="" then sPartyCode="NULL"
   ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
	    if trim(sType)="SUB" then
		    set subContItemNode = subContDOM.createElement("HEADER")
				    subContItemNode.setAttribute "FORUNIT",sOrgID
				    subContItemNode.setAttribute "CREATEDON",dCreatedOn
				    subContItemNode.setAttribute "ITEMTYPE",sItmType
				    subContItemNode.setAttribute "ORDERTO","01"
				    subContItemNode.setAttribute "SUPPAGENT",sPartyCode
				    subContItemNode.setAttribute "TYPEOFPURCHASE",sPurType
				    subContItemNode.setAttribute "SHIPTOLOC","S"
				    subContItemNode.setAttribute "ORDERVALIDTO",dCreatedOn
				    subContItemNode.setAttribute "REMARKS",""
				    subContItemNode.setAttribute "CONREQUIRED","N"
				    subContItemNode.setAttribute "APPROVER","0"
				    subContItemNode.setAttribute "CREATEDBY",iCreatedBy
				    subContItemNode.setAttribute "ORDERVALUE","0"
				    subContItemNode.setAttribute "REF",""
				    subContItemNode.setAttribute "PURCHASEORDERFOR","C"
				    subContItemNode.setAttribute "WITHMAT","Y"
				    subContItemNode.setAttribute "AppRefType","12"
				    subContItemNode.setAttribute "AppRefNo",iLedIssueNo
				    subContItemNode.setAttribute "AppRefDate",IssDate
				    subContItemNode.setAttribute "delinst",""
				    subContItemNode.setAttribute "schrelreq","N"
			    subContRoot.appendChild subContItemNode

			    if RootNode.HaschildNodes() then
				    For Each HeaderNode In RootNode.childNodes
					    if strcomp(HeaderNode.nodeName,"PURACC")=0 then
						    subContRoot.appendChild HeaderNode
					    End If
				    Next
			    end if
	    end if ' if Trim(IssToCode)="SUB" then

				    if trim(IssToCode) = "DIS" or trim(sSubConProfoma)="Y" and 1 = 2 then
				    
				        sTempSeries = GetSalNumberSeriesCodes("DIS",sOrgID,iNumIssueClassCode)
				        sArrSeries = Split(sTempSeries,":")
				        iSeriesNo = sArrSeries(0)
				        iSeriesCode = sArrSeries(1)
				        
					    
					    if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
					        sDISCode = ""
					        sSql = "Select GroupName from INV_M_Classification where GroupCode = "& iNumIssueClassCode
	                        dcrs.Open sSql,con
	                        if not dcrs.EOF then
	                            sNumClassName = Trim(dcrs(0))
	                        end if
	                        dcrs.Close 
                    	
	                        Response.Clear 
	                        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Despatch Instruction Slip - "& sNumClassName &"  Classification</H2></p>"
					        Response.End 
					        
					    end if
					    
					    if not CheckNoSerAvilForThisYear(sOrgID,iSeriesNo,iSeriesCode,IssDate) then
                            Response.Clear 
                            Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Despatch Instruction Slip - "& sNumClassName &"  Classification for this Year </H2></p>"
                            Response.End 
                        end if

                        sDISCode = GenSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,IssDate)
					    
					    'Response.Write "<p>sDIsCode = "& sDISCode
				    end if ' if trim(IssToCode) = "DIS" then
    'end if ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then

			 
			    
			    sTempSeries = GetInvNumberSeriesCodes("MR",sOrgID,iNumIssueClassCode)
                sArrSeries = Split(sTempSeries,":")
                iSeriesNo = sArrSeries(0)
                iSeriesCode = sArrSeries(1)
                'Response.Write "<p>iNumIssueClassCode = "& iNumIssueClassCode
	            if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
            	
	                sSql = "Select GroupName from INV_M_Classification where GroupCode = "& iNumIssueClassCode
	                dcrs.Open sSql,con
	                if not dcrs.EOF then
	                    sNumClassName = Trim(dcrs(0))
	                end if
	                dcrs.Close 
            	
	                Response.Clear 
	                Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Material Requisition - "& sNumClassName &"  Classification</H2></p>"
	                Response.End 
	            
	            end if
	            
	            if not CheckNoSerAvilForThisYear(sOrgID,iSeriesNo,iSeriesCode,IssDate) then
                    Response.Clear 
                    Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Material Requisition - "& sNumClassName &"  Classification for this Year </H2></p>"
                    Response.End 
                end if
                
                
		            iGenCode=GenSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,IssDate)				   
	
			    sSql = " "

		    If  sIssType = "M" Then
				    with dcrs
					    .CursorLocation = 3
					    .CursorType = 3
					    '.Source = "SELECT ISNULL(MAX(ISSUENO)+1,1) FROM INV_T_DEPARTMENTSTOCK "
					    .source = "Select IsNull(max(IssueEntryNo),0)+1 from INV_T_MaterialIssueHeader"
					    .ActiveConnection = con
					    .Open
				    end with
				    set dcrs.ActiveConnection = nothing

				    if not dcrs.EOF then
					    iLedIssueNo = dcrs(0)
				    end if
				    dcrs.Close
		    End If'If  sIssType = "M" Then
		    
		    'Response.Write "<P>NumClass Code = "& iNumIssueClassCode
		    
		    sTempSeries = GetInvNumberSeriesCodes("IS",sOrgID,iNumIssueClassCode)
		    sArrSeries = Split(sTempSeries,":")
		    iSeriesNo = sArrSeries(0)
		    iSeriesCode = sArrSeries(1)
		    
		    if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
        	
                sSql = "Select GroupName from INV_M_Classification where GroupCode = "& iNumIssueClassCode
                dcrs.Open sSql,con
                if not dcrs.EOF then
                    sNumClassName = Trim(dcrs(0))
                end if
                dcrs.Close 
        	
                Response.Clear 
                Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Issue Number - "& sNumClassName &"  Classification</H2></p>"
                Response.End 
            end if
            
            if not CheckNoSerAvilForThisYear(sOrgID,iSeriesNo,iSeriesCode,IssDate) then
                Response.Clear 
                Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Issue Number - "& sNumClassName &"  Classification for this Year </H2></p>"
                Response.End 
            end if
                        
                        iGenCode=GenSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,IssDate)				   
                        
	            
		    sSql = " "

		    'End
		    
		    if trim(iGenCode)="" or IsNull(iGenCode) then iGenCode = "NULL"
		    if trim(iGenCode)<>"NULL" then iGenCode = Pack(iGenCode)

	    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    if trim(iNoofCases)="" or isnull(iNoofCases) then iNoofCases = 0
	    'Added by Maheshwari on Oct 31st 2007
	    with dcrs
		    .CursorLocation = 3
		    .CursorType = 3
		    '.Source = "SELECT ISNULL(MAX(ISSUENO)+1,1) FROM INV_T_DEPARTMENTSTOCK "
		    .source = "Select IsNull(max(IssueEntryNo),0)+1 from INV_T_MaterialIssueHeader"
		    .ActiveConnection = con
		    .Open
	    end with
	    set dcrs.ActiveConnection = nothing

	    if not dcrs.EOF then
		    iLedIssueNo = dcrs(0)
	    end if
	    dcrs.Close
	   ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR"  then
		    'added by Ragav on March 25,2010 for  ForInvocie_HEADER Table usage="DIS"
		    if trim(IssToCode)="DIS" or trim(sSubConProfoma)="Y" and 1 = 2 then
			    'if trim(sSalInvConfirm)="N" then

				    sSql = "Select isNull(max(ForInvoiceNo),0)+1 from FORINVOICE_HEADER"

				    rsTemp.Open sSql,con
				    if not rsTemp.EOF then
					    sInvNo = rsTemp(0)
				    end if
				    rsTemp.Close

				    sSql = " Select IssueEntryNo,ForInvoiceNo from INV_T_ReferenceNumbers where IssueEntryNo = "& iLedIssueNo
				    rsTemp.Open sSql,con
				    if not rsTemp.EOF then
					    sSql = "Update INV_T_ReferenceNumbers set ForInvoiceNo = "& sInvNo &" where IssueEntryNo = "& iLedIssueNo
				    else
					    sSql = "Insert into INV_T_ReferenceNumbers (IssueEntryNo,ForInvoiceNo) values("& iLedIssueNo &","& sInvNo  &")"
				    end if
				    rsTemp.Close

				    con.execute(sSql)

				    if trim(sDISCode)="" then
					    sDISCode = "NULL"
				    else
					    sDISCode = pack(sDISCode)
				    end if
				    
				    ''Response.Write"<p>sSelectedInvoice="&sSelectedInvoice

				    if sSelectedInvoice = "A" then

					    sSql = "Insert into FORINVOICE_HEADER (ForInvoiceNo,InvoicedForUnit,ReferenceName,ReferenceNo,PartyType,"&_
						       "PartySubType,PartyCode,TypeofSale,TypeofInvoice,GrossWeight,NettWeight,NoofCases,CreatedBy,"&_
						       "CreatedOn,DISCode) values ("&sInvNo&",'"& sOrgID &"','DIS','"&iLedIssueNo&"',NULL,"&_
						       "NULL,"& sPartyCode &","& sSSalType &",'"& sSInvType &"',"& iIsQty &","& (cdbl(iIsQty) -cdbl(iTqty)) &","&iNoofCases&","& iCreatedBy &","&_
						       "Convert(datetime,'"& dCreatedOn &"',103),"& sDISCode &")"
				    else
					    sSql = "Insert into FORINVOICE_HEADER (ForInvoiceNo,InvoicedForUnit,ReferenceName,ReferenceNo,PartyType,"&_
						       "PartySubType,PartyCode,TypeofSale,TypeofInvoice,GrossWeight,NettWeight,NoofCases,CreatedBy,"&_
						       "CreatedOn,DISCode,ProformaInvoice) values ("&sInvNo&",'"& sOrgID &"','DIS','"&iLedIssueNo&"',NULL,"&_
						       "NULL,"& sPartyCode &","& sSSalType &",'"& sSInvType &"',"& iIsQty &","& (cdbl(iIsQty) -cdbl(iTqty)) &","&iNoofCases&","& iCreatedBy &","&_
						       "Convert(datetime,'"& dCreatedOn &"',103),"& sDISCode &",'Y')"
				    end if

				    	'Response.Write "<p>"&sSql

				    con.execute sSql

			    'end if'if trim(sSalInvConfirm)="N" then
		    end if ' 	if trim(IssToCode)="DIS" then

		    if trim(sType)="SER" then '' For  Services Case
			    sSql = "Select isnull(max(GatePassNo),0)+1 from FORGATEPASSHEADER"
			    rsTemp.Open sSql,con
			    if not rsTemp.EOF then
				    sGatePassNo =  rsTemp(0)
			    end if
			    rsTemp.Close

			    sSql = "Insert into FORGATEPASSHEADER (GatePassNo,ReferenceName,ReferenceNo,OrganisationCode,InvoiceType,PartyCode,ApplicationCode,"&_
				       "MarkedOn,Remarks,Status,GeneratedOn,DCCode,NoofPacks,PackingType,Transport,TakenBy,DeliveryBy,RefType,AppRefType,AppRefNo,AppRefDate)"&_
				       "values("& sGatePassNo &",'ISSUE',"&iLedIssueNo &","&pack(sOrgID)&",'V',"&IssToCode&",4,"&_
				       "Convert(datetime,'"&FormatDate(date)&"',103),"&sRemarks&",'N',Convert(datetime,'"&Formatdate(date)&"',103),NULL,"&iNoofCases&",NULL,"& sTransPort &","& sTakenBy &","& sDeliveryBy &",NULL,12,"& pack(iLedIssueNo) &","& pack(IssDate) &")"

			    'Response.Write "<p>"&sSql
			    con.execute sSql
		    end if

		    'End ragav
	    'end if ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then

	    IF trim(iMRSNo) <> "" then
		    sSql = "INSERT INTO INV_T_MaterialIssueHeader (IssueEntryNo,OrganisationCode,IssueEntryCode,ReferenceNo,ReferenceType," &_
			       "MaterialReceivedBy,Remarks,IssueDate,IssuedBy,CreatedOn,IssuedToCode,IssuedToSubCode,IssuedToType,IssueType,AppRefType,AppRefNo,AppRefDate,MarkPackFlag,IssueFrom,Returnable,ReturnItem,IssueTypeCode) VALUES " &_
			       "(" & iLedIssueNo & "," & Pack(sOrgID) & "," & iGenCode & "," & iMRSNo & ",'M' ," &_
			       "" & Pack(sRecBy) & "," & Pack(sRem) & ",Convert(DateTime," & Pack(IssDate) & ",103)," &_
			       ""&iIssuedBy&",Convert(datetime,'"& IssDate &"',103),'"& IssToCode &"',"& IssToSubCode &","& IssToType &",'"& sIssType &"',"& sAppRefType &","& sAppRefNo &",'"&sAppRefDate&"',"&Pack(sPickPackFlag)&","& pack(sIssFrom) &","& pack(sReturnable) &","& pack(sReturnItem) &","& pack(sType) &")"
		     'Response.Write "<p>"&sSql & vbCrLf
		    con.execute sSql
	    Else
		    sSql = "INSERT INTO INV_T_MaterialIssueHeader (IssueEntryNo,OrganisationCode,IssueEntryCode,ReferenceType," &_
			       "MaterialReceivedBy,Remarks,IssueDate,IssuedBy,CreatedOn,IssuedToCode,IssuedToSubCode,IssuedToType,IssueType,AppRefType,AppRefNo,AppRefDate,MarkPackFlag,IssueFrom,Returnable,ReturnItem,IssueTypeCode) VALUES " &_
			       "(" & iLedIssueNo & "," & Pack(sOrgID) & "," & iGenCode & ",'D' ," & Pack(sRecBy) & "," &_
			       "" & Pack(sRem) & ",Convert(DateTime," & Pack(IssDate) & ",103),"&iIssuedBy&"," &_
			       " Convert(datetime,'"& IssDate &"',103),'"& IssToCode &"',"& IssToSubCode &","& IssToType &",'"& sIssType &"',"& sAppRefType &","& sAppRefNo &",'"&sAppRefDate&"',"&Pack(sPickPackFlag)&","& pack(sIssFrom) &","& pack(sReturnable) &","& pack(sReturnItem) &","& pack(sType) &")"
		     Response.Write "<p>"&sSql & vbCrLf
		    con.execute sSql
	    End IF
	    '**************
	    
	    sSql = "Select isNull(Max(ConsumptionNo)+1,1) from INV_T_MaterialConsumption"
	    rsTemp.open sSql,con
	    if not rsTemp.eof then
	        sConEntryNo = rsTemp(0)
	    end if
	    rsTemp.close 
	    
	    if RootNode.haschildNodes() then
	        for each HeaderNode in RootNode.childNodes
	            if trim(HeaderNode.nodeName)="SubContract" then
	                sProcessID = triM(HeaderNode.Attributes.getNamedItem("SCProcess").Value)
                    sInstruct = triM(HeaderNode.Attributes.getNamedItem("Instruct").Value)
                    sLabCharge = triM(HeaderNode.Attributes.getNamedItem("LabourCharge").Value)
                    sHardWaste = triM(HeaderNode.Attributes.getNamedItem("HardWaste").Value)
                    sInvWaste = triM(HeaderNode.Attributes.getNamedItem("InvWaste").Value)
                    
                    if Trim(sType)="SUB" then
                        set subContProcessHead = subContDOM.createElement("SubContract")
                            subContProcessHead.setAttribute "SCProcess",sProcessID
                            subContProcessHead.setAttribute "Instruct",sInstruct
                            subContProcessHead.setAttribute "LabourCharge",sLabCharge 
                            subContProcessHead.setAttribute "Currency","1"
                            subContProcessHead.setAttribute "HardWaste",sHardWaste
                            subContProcessHead.setAttribute "InvWaste",sInvWaste
                            subContProcessHead.setAttribute "ProcessName",""
                        subContRoot.appendChild subContProcessHead
                    end if 
                    
                    if trim(sProcessID)="" or IsNull(sProcessID) then sProcessID = "NULL"
			        if trim(sInstruct)="" or IsNull(sInstruct) then sInstruct = "NULL"
			        if trim(sLabCharge)="" or IsNull(sLabCharge) then sLabCharge = "NULL"
			        if trim(sHardWaste)="" or IsNull(sHardWaste) then sHardWaste = "NULL"
			        if trim(sInvWaste)="" or IsNull(sInvWaste) then sInvWaste = "NULL"
			        if trim(sInstruct)<>"NULL" then sInstruct = pack(sInstruct)
			    
                    iReturnItemEntryNo = 0
                    sSql = "Update INV_T_MaterialIssueHeader set SCProcess = "& sProcessID &",Instruct="& sInstruct &",LabourCharge="& sLabCharge &",HardWaste="& sHardWaste &",IntWaste="& sInvWaste &" where IssueEntryNo="&iLedIssueNo
                    Response.Write "<p>"&sSql & vbCrLf
		            con.execute sSql
		            for each SubNode in HeaderNode.childNodes
		                if trim(SubNode.nodeName)="Details" then
		                    iReturnItemEntryNo = iReturnItemEntryNo +1
		                    sReturnItemCode = SubNode.getAttribute("MatRecdAsItem")
		                    sReturnClassCode = SubNode.getAttribute("MatRecdAsCode")
		                    sReturnItem = SubNode.getAttribute("MatRecdAsDescr")
		                    
		                    if Trim(sType)="SUB" then
		                        set subContProcessItem = subContDOM.createElement("SubContract")
                                    subContProcessItem.setAttribute "MatRecdAsItem",sReturnItemCode
                                    subContProcessItem.setAttribute "MatRecdAsCode",sReturnClassCode
                                    subContProcessItem.setAttribute "MatRecdAsDescr",sReturnItem 
                                subContProcessHead.appendChild subContProcessItem
		                    end if 
		                    
		                    if trim(sReturnItemCode)<>"" then
		                        sSql = "Insert into INV_T_MaterialIssueReturnItem (IssueEntryNo,ItemCode,ClassificationCode,ItemDesc,ItemEntryNo,OrganisationCode,Returnable,ReturnItem) values ("& iLedIssueNo &","& sReturnItemCode &","& sReturnClassCode &","& pack(sReturnItem) &","& iReturnItemEntryNo &","& Pack(sOrgID) &",'Y','D')"
		                        Response.Write "<p>"&sSql & vbCrLf 
		                        con.execute sSql
		                    end if 'if trim(sReturnItemCode)<>"" then
		                    
		                end if
		            next
	            end if
	        next
	    end if 
	    
	    
    '''''''''''''''''''''''''''''''''''''
sOnlyLotFlag=""
	    For Each HeaderNode In RootNode.childNodes
		    if StrComp(HeaderNode.nodeName,"ITEM") = 0 then
			    iEntNo = trim(HeaderNode.Attributes.getNamedItem("ENTRYNO").Value)
			    iClass = trim(HeaderNode.Attributes.getNamedItem("CLACODE").Value)
			    iItemCode = trim(HeaderNode.Attributes.getNamedItem("ITMCODE").Value)
			    dMRSDate =  trim(HeaderNode.Attributes.getNamedItem("MRSDATE").Value)
			    arrStore = split(trim(HeaderNode.Attributes.getNamedItem("SSTORE").Value),":")
			    sAttributeList =  trim(HeaderNode.Attributes.getNamedItem("ATTRIBUTELIST").Value)
			    sInvItemQty = trim(HeaderNode.Attributes.getNamedItem("ISSQTY").Value)
			    sItemRefNo  = trim(HeaderNode.Attributes.getNamedItem("RefNo").Value)
			    sOnlyLotFlag = trim(HeaderNode.Attributes.getNamedItem("ONLYLOT").Value)
			    if trim(sInvItemQty)="" or IsNull(sInvItemQty) then sInvItemQty ="0"
			    
			    if (trim(sOnlyLotFlag)="NULL" or isNull(sOnlyLotFlag) or trim(sOnlyLotFlag)="") and sPickPackFlag="N" then sOnlyLotFlag="P"
			    
			    if trim(sOnlyLotFlag)="" or isNull(sOnlyLotFlag) then sOnlyLotFlag = "NULL"
			    if trim(sOnlyLotFlag)<>"NULL" then sOnlyLotFlag = pack(sOnlyLotFlag)
			    sIReturnable = trim(HeaderNode.Attributes.getNamedItem("RETURNABLE").Value)
			    sIReturnItem = trim(HeaderNode.Attributes.getNamedItem("RETURNITEM").Value)
			    sMatType = trim(HeaderNode.Attributes.getNamedItem("MatType").Value)
			    
			   
			    if trim(sMatType)="" or IsNull(sMatType) then sMatType = "NULL"
			    
			    if trim(sMatType)<>"NULL" then sMatType = pack(sMatType)
			    
			    Response.write "<p>OnlyLotPackFlag = "& sOnlyLotFlag
			    Response.write "<p>MarkFor Pick="& sPickPackFlag

		if trim(sAttributeList)<>"" then
    			sArrTemp=split(sAttributeList,":")
		        sArrList = split(sArrTemp(0),"#")
		        if sArrList(0)<>"0" then
			        if UBound(sArrList)>0 then
				        if trim(sArrList(1))<>"0" and trim(sArrList(1))<>"" then
					        sAttID =  sArrList(1)
				        end if
				    else
						if trim(sArrList(0))<>"0" and trim(sArrList(0))<>"" then
							sAttID = sArrList(0)
						end if
			        end if
			    end if
		end if'if trim(sAttributeList)<>"" then

                if trim(sAttID)="" or IsNull(sAttID) then sAttID = "NULL"
                
                'added by Ragav on March 25,2010 for  ForInvocie_Details Table usage="DIS"
			    sSql = "Select isNUll(SalesUoM,StoresUoM) from INV_M_ITEMMASTER where ItemCode = "& iItemCode  &" and ClassificationCode = "&iClass
			    rsTemp.Open sSql,con
			    if not rsTemp.EOF then
				    sSalUoM = rsTemp(0)
			    end if
			    rsTemp.Close
			    
			    sSql = "Select isNull(Max(LineNumber)+1,1) from INV_T_MaterialConsumption where ConsumptionNo = "& sConEntryNo
	            rsTemp.open sSql,con
	            if not rsTemp.eof then
	                sConLineNo = rsTemp(0)
	            end if
	            rsTemp.close 
                
                
                if sAutoConsumption = "Y" then
            	    sSql = " Insert into INV_T_MaterialConsumption (ConsumptionNo,LineNumber,IssueEntryNo,ConsumedByDept,"&_
            	           " OrganisationCode,ItemCode,ClassificationCode,QuantityConsumed,QuantityUOM,ApplicationCode,Remark, "&_
                           " EnteredOn,EnteredBy,AttributeList) values ("& sConEntryNo &","& sConLineNo &","& iLedIssueNo &",NULL,'"& sOrgID &"',"&_
                           " "& iItemCode &","& iClass &","& sInvItemQty &",'"& sSalUoM &"',4,'',"&_
                           " Convert(datetime,getDate(),103),"&iIssuedBy &","&sAttID&")"
                    Response.Write "<p> Auto Consumption  header =  "& sSql &"<p>"
                    con.execute sSql 
                end if


			    
			   ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
				    if (trim(IssToCode)="DIS" or trim(sSubConProfoma)="Y") and 1 = 2 then ' For despatch case
					    'if trim(sSalInvConfirm)="N" then
						    sSql = "Insert into FORINVOICE_DETAILS (ForInvoiceNo,ItemCode,ClassificationCode,QuantityForInvoice,"&_
								    "InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemAttributes) values("& sInvNo &","& iItemCode &","&_
								    ""& iClass &","& sInvItemQty &","& Pack(sSalUoM) &",0,0,0,0,"& sAttID &")"
						    	Response.Write "<p>"&sSql
						    con.execute sSql
					    'end if ' if trim(sSalInvConfirm)="N" then
				    elseif trim(sType)="SER" then 'for ServicesCase
					    sGatePassEntryNo = sGatePassEntryNo + 1
						    sSql  = "Insert into FORGATEPASSDETAILS (GatePassNo,EntryNo,ItemCode,ClassificationCode,Quantity,"&_
								    "InvoicedUOM,Description,MaterialRcvd,NoofPacks,PackingType,ItemValue,ItemAttributes) values("& sGatePassNo &","&_
								    ""&sGatePassEntryNo&","&iItemCode&","&iClass&","&sInvItemQty&","&Pack(sSalUoM)&",NULL,'N',NULL,NULL,NULL,"& sAttID &")"
						    	Response.Write "<p>"&sSql
						    con.execute sSql
				    end if ' if trim(IssToCode)="DIS" then
				    'End Ragav command
			'    end if ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then


			    bFlag = true
			    IF dMRSDate = "" then dMRSDate = IssDate
			    sTempMonYr = mid(dMRSDate,4,2)
			    sMonYr = sTempMonYr&Year(dMRSDate)

			    'arrFin = split(GetFinancialYear(sMonYr),":")
			    'sFinFrom = arrFin(0)
			    'sFinTo = arrFin(1)
			    arrFin = split(session("Finperiod"),":")
			    sFinFrom = "01/04/"&arrFin(0)
			    sFinTo = "31/03/"&arrFin(1)
			    'Response.Write "sFinFrom="&sFinFrom &"  "&sFinTo
			    'sLoc = arrStore(0)
			    'sBin = arrStore(1)
'			    'Response.Write "sLoc======="&sLoc
			    'IF sLoc = "0" then
				    with dcrs
					    .CursorLocation = 3
					    .CursorType = 3
					    .Source = "SELECT ISNULL(LOCATIONNUMBER,0),ISNULL(BINNUMBER,0) FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
					    .ActiveConnection = con
					    .Open
				    end with
				    set dcrs.ActiveConnection = nothing
'				    'Response.Write dcrs.Source
				    if not dcrs.EOF then
					    sLoc = dcrs(0)
					    sBin = dcrs(1)
				    end if
				    dcrs.close
'				    'Response.Write "sLoc ="&sLoc & "***"& sBin & vbCrLf& vbCrLf
			    'End IF
			    if sBin <> "N" then
				    sBin = sBin
			    else
				    sBin = "NULL"
			    end if

			    iReqQty = cdbl(HeaderNode.Attributes.getNamedItem("REQQTY").Value)
			    iIssQty = trim(HeaderNode.Attributes.getNamedItem("ISSQTY").Value)
			    iTraQty = cdbl(HeaderNode.Attributes.getNamedItem("TRAQTY").Value)
			    iPrQty = trim(HeaderNode.Attributes.getNamedItem("PRQTY").Value)
			    iValue = trim(HeaderNode.Attributes.getNamedItem("IVALUE").Value)
			    dMRSDate = trim(HeaderNode.Attributes.getNamedItem("MRSDATE").Value)
			    if iIssQty = "" or iIssQty = "0" then
				    iIssQty = cdbl("0")
			    else
				    iIssQty = cdbl(iIssQty)
			    end if

			    if iPrQty = "" or iPrQty = "0" then
				    iPrQty = cdbl("0")
			    else
				    iPrQty = cdbl(iPrQty)
			    end if

			    with rsTemp
				    .CursorLocation = 3
				    .CursorType = 3
				    .ActiveConnection = con
				    .Source = "Select isNull(YearClosingValue,0),isNull(YearClosingStock,0) from INV_T_ITEMLOCATIONSTOCK where itemcode = "&iItemCode
		    	Response.Write "<p>"&rsTemp.Source
				    .Open
			    end with
			    if not rsTemp.EOF then
				    if trim(rsTemp(0))<>"0" and trim(rsTemp(1))<>"0" then
			    '	'Response.Write " Closing Value = "& rsTemp(0) &"  cdbl(Closing Value) = "& cdbl(rsTemp(0))
			    '	'Response.Write "  Closing Stock = "& rsTemp(1) & " cdbl(Closing Stock) = "& cdbl(rsTemp(1))
					    iValue = cdbl(rsTemp(0))/cdbl(rsTemp(1))
				    '	'Response.Write "Hello ** "
				    end if
			    end if
			    rsTemp.Close
			    ''Response.Write "iValue = "& iValue
			    if trim(iValue)="" or IsNull(iValue) then iValue = "0"
			    iValue = Round(iValue,2)

			   ' if trim(sAppRefType)="14" then ' Mix Code
			   Response.Write "<p>sItemRefNo="& sItemRefNo 
			        sArrAppRefNo = split(sItemRefNo,",")
			       ' if UBound(sArrAppRefNo) = 0 then
			            if trim(iMRSNo)<>"" then
                            sSql = "Insert into Inv_T_MaterialIssueAdditionalDetails (IssueEntryNo,MRSNumber,OrganisationCode,ClassificationCode,ItemCode,MixCode,QuantityIssued)"&_
                                   " values ("& iLedIssueNo &","& iMRSNo &",'"& sOrgID &"',"& iClass &","& iItemCode  &","& Pack(sItemRefNo) &","& iIssQty  &")"
                            'Response.Write "<p>"&sSql
                            con.Execute sSql
                        else
                            sSql = "Insert into Inv_T_MaterialIssueAdditionalDetails (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,MixCode,QuantityIssued)"&_
                                   " values ("& iLedIssueNo &",'"& sOrgID &"',"& iClass &","& iItemCode  &","& Pack(sItemRefNo) &","& iIssQty &")"
                            'Response.Write "<p>"&sSql
                            con.Execute sSql
                        end if
                        Response.Write "<p>"& sSql 

                        nTransactQty = iIssQty
                        sTransactUOM = sSalUoM
                '        'Response.Clear
                        ''Added by Ragav on Mar 04,2011
                        Response.Write "<p> nTransasctQty = "& nTransactQty
                            if CDbl(nTransactQty) > 0 then
				                ' WIP Stock Table updation - start
				                'if trim(IssToCode) = "PRD" and 1=2 then
				                if trim(IssToCode) = "PRD" then
				                
					                  '  if  ObjFs1.FileExists( server.MapPath("/Production/xmldata/DailyMixing.xml")) then
                					  '      OutDataMix.Load server.MapPath("/Production/xmldata/DailyMixing.xml")
							          '      Set RootTemp = OutDataMix.documentElement

							          '      sExp = "//MixingData[@Unit='"& sOrgID &"']"
							          '      Set Tempnode = RootTemp.SelectNodes(sExp)
                					  '      if Tempnode.length > 0 then
								      '              sSelectedWC		= Tempnode.item(0).GetAttribute("MixingWC")
								      '      end if 'if Tempnode.length > 0 then
						              '  end if
						              
						              sSelectedWC = replace(IssToSubCode,"'","")

					                nWCLocNo  = 0
					                nWCBinNo = 0

					                sSqlTemp = "select LocationNumber,BinNumber from PRD_M_Workcenter where WorkCenterCode = '" & sSelectedWC & "'"
					                Response.Write "<p>"&  sSqlTemp

					                with dRSet
						                .ActiveConnection = con
						                .CursorLocation = 3
						                .CursorType = 3
						                .Source = sSqlTemp
						                .Open
					                end with

					                set dRSet.ActiveConnection = nothing

					                if not dRSet.EOF then
						                nWCLocNo	= dRSet(0)
						                nWCBinNo	= dRSet(1)
					                end if
					                dRSet.Close

					                sMonthYear = FindMonthYear(IssDate)
					                Response.Write "<p>sMonthYear = "& sMonthYear
					                sMixCode = sItemRefNo


					                sStockEntryExistForCurrentWC = "N"
					                '----------- to carry previous month cl.stk to current month - start
					                sSqlTemp = "Select isNull(InputStock,0) from PRD_T_WIPProcessStockYRLY  " & _
							                " where WorkCenterCode = '" & sSelectedWC & "' and OrganisationCode = '" & sOrgID  & "'" &_
							                " and MonthYear='" & sMonthYear & "' and WasteCode is NULL " &_
							                " and LocationNumber = " & nWCLocNo & " and BinNumber = " & nWCBinNo & "" &_
							                " and MixCode='" & sMixCode & "'"
							         Response.Write "<p>"&sSqlTemp
					                with dRSet
						                .ActiveConnection = con
						                .CursorLocation = 3
						                .CursorType = 3
						                .Source = sSqlTemp
						                .Open
					                end with

					                set dRSet.ActiveConnection = nothing

					                if not dRSet.EOF then
						                sStockEntryExistForCurrentWC = "Y"
					                end if
					                dRSet.Close


					                ''Response.Write "<p> " & sStockEntryExistForCurrentWC

					                '----------- to carry previous month cl.stk to current month - end
					                if sStockEntryExistForCurrentWC = "N" then

                					    sCurrentMonthFirstDate = FindFirstDate(sMonthYear)

						                Response.Write "<P> sCurrentMonthFirstDate = " & sCurrentMonthFirstDate



						                sLastEntryMonthYear = ""

						                'checking any entry is exist for selected mix code
						                sSql = "Select MixCode from PRD_T_WIPStockLedger where WorkCenterCode='" & sSelectedWC & "' and convert(datetime,TransactionDate,103) <= convert(datetime,'" & dTDate & "',103) and MixCode='" & sMixCode & "'"
'						                'Response.Write "<p> " & sSql
						                With rsNew
						 	                .CursorLocation = 3
							                .CursorType = 3
							                .Source = sSql
							                .ActiveConnection = con
							                .Open
						                End with

						                Set rsNew.Activeconnection = Nothing

						                If Not rsNew.EOF Then

							                rsNew.close

							                sSqlCmd  = "Select convert(datetime,max(TransactionDate),103) from PRD_T_WIPStockLedger where WorkCenterCode='" & sSelectedWC & "' and convert(datetime,TransactionDate,103) <= convert(datetime,'" & dtDate & "',103) and MixCode='" & sMixCode & "'"
'							                'Response.Write "<p> " & sql
							                With rsNew
							 	                .CursorLocation = 3
								                .CursorType = 3
								                .Source = sSqlCmd
								                .ActiveConnection = con
								                .Open
							                End with

							                Set rsNew.Activeconnection = Nothing

							                If Not rsNew.EOF Then
								                sLastEntryMonthYear = FindMonthYear(rsNew(0))
								                Response.Write "<p>"&sLastEntryMonthYear 
							                end if
						                Else
							                'if no entry is exist for selected mixcode then put entry for current month
							                ' so generate month from current month -1

							                sLastEntryMonthYear = FindMonthYear(cdate(sCurrentMonthFirstDate) -1) ' here  ( -1 ) give you last month last date
							                Response.Write "<p>"&sLastEntryMonthYear 
							                
						                End If
						                rsNew.close

						                ''Response.Write "<p> sLastEntryMonthYear = " & sLastEntryMonthYear
						                if trim(sLastEntryMonthYear) <> "" then

							                'finding next month before loop
							                if right("0" & trim( CInt(left(sLastEntryMonthYear,2)) + 1),2) = "13" then
								                sTempMonth = "01" & trim(cint(right(sLastEntryMonthYear,4)) + 1)
							                else
								                sTempMonth = right("0" & trim( CInt(left(sLastEntryMonthYear,2)) + 1),2) & right(sLastEntryMonthYear,4)
							                end if

							                ''Response.Write "<p>  sTempMonth = " & sTempMonth


							                do while true

								                sFirstDate =  FindFirstDate(sTempMonth)

								                ''Response.Write "sFirstDate = "& sFirstDate

								                sListOfMonthYearForWIPStockEntryNotExist = sListOfMonthYearForWIPStockEntryNotExist & "," & FindMonthYear(sFirstDate)

								                ''Response.Write "<p>new =" & sListOfMonthYearForWIPStockEntryNotExist

								                'finding next month inside loop
								                if right("0" & trim( CInt(left(sTempMonth,2)) + 1),2) = "13" then
									                sTempMonth = "01" & trim(cint(right(sTempMonth,4)) + 1)
								                else
									                sTempMonth = right("0" & trim( CInt(left(sTempMonth,2)) + 1),2) & right(sTempMonth,4)
								                end if

								                ''Response.Write "<p> sTempMonth = "& sTempMonth

								                sFirstDate =  FindFirstDate(sTempMonth)
								                ''Response.Write "sFirstDate = "& sFirstDate

								                sSql = "Select top 1 *  from PRD_T_WIPStockLedger where convert(datetime,'" & sFirstDate & "',103) <= convert(datetime,'" & sCurrentMonthFirstDate & "',103)"

								                ''Response.Write "<p> " & sSql
								                rsNew.open sSql,con
								                If Not rsNew.EOF Then
								                else
									                rsNew.close
									                exit do
								                End If
								                rsNew.close

							                loop

							                if trim(sListOfMonthYearForWIPStockEntryNotExist) <> "" then
								                sListOfMonthYearForWIPStockEntryNotExist = mid(sListOfMonthYearForWIPStockEntryNotExist,2)
							                end if
							                ''Response.Write "<p>new =" & sListOfMonthYearForWIPStockEntryNotExist


							                '////////
							                sSql   = "Select isNull(ClosingStock,0),isNull(ClosingStockValue,0),isNull(ClosingStockUOM,'') from PRD_T_WIPProcessStockYRLY " &_
									                " where isNull(MonthYear,'') = '" & sLastEntryMonthYear & "' and WorkCenterCode='" & sSelectedWC   &"'" & _
									                " and OrganisationCode ='" & sOrgID & "' and MixCode='" & sMixCode  & "'" & _
									                " and ProductCode is Null and WasteCode is Null"
						                Response.Write "<p> sql = " & sSql
							                With rsNew
							 	                .CursorLocation = 3
								                .CursorType = 3
								                .Source = sSql
								                .ActiveConnection = con
								                .Open
							                End with
							                Set rsNew.Activeconnection = Nothing
							                If Not rsNew.EOF Then
								                'if cdbl(rsNew(0)) > 0 then
                                                Response.Write "<p>"& sListOfMonthYearForWIPStockEntryNotExist

									                ArrMonth = Split(sListOfMonthYearForWIPStockEntryNotExist,",")

									                For nTempCtr = LBound(ArrMonth) to UBound(ArrMonth)

										                sTempMonth =  ArrMonth(nTempCtr)
										                sCurrentMonthFirstDate = FindFirstDate(sTempMonth)

										                'opening stock entry in PRD_T_WIPStockLedger for current wc
										                sSqlCmd = " Insert Into PRD_T_WIPStockLedger(OrganisationCode,WorkCenterCode,ClassificationCode,ItemCode,ProductCode,TransactionType,TransactionDate,ShiftCode,TransactQuantity,TransactValue,TransactUOM,MixCode,WasteCode,AppRefType,AppRefNo,AppRefDate) " &_
													                " Values('" & sOrgID & "','" & sSelectedWC & "',NULL,NULL,NULL,'O',convert(datetime,'" & sCurrentMonthFirstDate & "',103),NULL," & rsNew(0) & "," & rsNew(1) & ",'" & rsNew(2) & "','" & sMixCode & "',NULL,12,"& iLedIssueNo &",Convert(datetime,'"& IssDate &"',103))"

										                Response.Write "<p> For Next WC--Ledger Entry = " & sSqlCmd
										                con.execute(sSqlCmd)


										                If sStockEntryExistForCurrentWC = "N" then
											                 'opening stock entry in PRD_T_WIPProcessStockYRLY for current wc
											                sSqlCmd = " Insert Into PRD_T_WIPProcessStockYRLY (WorkCenterCode,OrganisationCode,LocationNumber,BinNumber,MonthYear,OpeningStock,OpeningStockUOM,OpeningStockValue,MixCode,WasteCode,ClassificationCode,ItemCode,ProductCode) " & _
													                " Values('" & sSelectedWC  & "','" & sOrgID  & "'," & nWCLocNo & "," & nWCBinNo & ",'" & sTempMonth & "'," & rsNew(0) & ",'" & rsNew(2) & "'," & rsNew(1) & ",'" & sMixCode & "',NULL,NULL,NULL,NULL)"
										                else
											                sSqlCmd = " Update PRD_T_WIPProcessStockYRLY set OpeningStock=" & rsNew(0) & ", OpeningStockValue=" & rsNew(1) & " " & _
													                " where WorkCenterCode='" & sSelectedWC  & "' and OrganisationCode='" & sOrgID & "' and LocationNumber=" & sLocNo & " " & _
													                " and BinNumber=" & sBinNO & " and MonthYear='" & sTempMonth & "' and MixCode='" & sMixCode & "' and WasteCode is NULL and ProductCode is NULL"
										                End If

										                ''Response.Write "<p> For Next WC-- Stock Entry = " & sSqlCmd
										                con.execute(sSqlCmd)

										                sSqlCmd = "Update PRD_T_WIPProcessStockYRLY set ClosingStock=(isNull(OpeningStock,0)+isNull(InputStock,0))-isNull(OutputStock,0),ClosingStockValue=(isNull(OpeningStockValue,0)+isNull(InputStockValue,0)) - isNull(OutputStockValue,0) " &_
												                " where isNull(MonthYear,'') = '" & sTempMonth & "' and WorkCenterCode='" & sSelectedWC   &"' and OrganisationCode ='" & sOrgID & "'"
										                Response.Write "<p>" & sSqlCmd
										                Con.Execute(sSqlCmd)

									                Next


								                'end if 'if cdbl(rsNew(0)) > 0 then
							                End If
							                rsNew.close

						                end if 'if trim(sLastEntryMonthYear) <> "" then

					                end if 'if sStockEntryExistForCurrentWC = "N" then

					                sSqlCmd = "Insert Into PRD_T_WIPStockLedger (OrganisationCode,WorkCenterCode,ClassificationCode,ItemCode,ProductCode,TransactionType,TransactionDate,ShiftCode,TransactQuantity,TransactUOM,MixCode,WasteCode,AppRefType,AppRefNo,AppRefDate) "  & _
						                "Values('" & sOrgID & "','" & sSelectedWC  & "'," & iClass & "," & iItemCode & ",NULL,'I',Convert(DateTime,'" & IssDate & "',103),NULL," & nTransactQty & ",'" & sTransactUOM & "','" & sMixCode & "',NULL,12,"& iLedIssueNo &",Convert(datetime,'"& IssDate &"',103))"
					                Response.Write "<p> sSqlCmd = " & sSqlCmd
					                con.execute sSqlCmd


					                sSqlTemp = "Select isNull(InputStock,0) from PRD_T_WIPProcessStockYRLY  " & _
							                " where WorkCenterCode = '" & sSelectedWC & "' and OrganisationCode = '" & sOrgID  & "'" &_
							                " and MonthYear='" & sMonthYear & "' and WasteCode is NULL " &_
							                " and LocationNumber = " & nWCLocNo & " and BinNumber = " & nWCBinNo & "" &_
							                " and MixCode='" & sMixCode & "'"
'							        'Response.Write sSqlTemp
					                with dRSet
						                .ActiveConnection = con
						                .CursorLocation = 3
						                .CursorType = 3
						                .Source = sSqlTemp
						                .Open
					                end with

					                set dRSet.ActiveConnection = nothing

					                if dRSet.EOF then
						                sSqlCmd = "Insert Into PRD_T_WIPProcessStockYRLY (WorkCenterCode,OrganisationCode,ClassificationCode,ItemCode,ProductCode,LocationNumber,BinNumber,MonthYear,InputStock,InputStockUOM,InputStockValue,MixCode,WasteCode) " &_
									                "Values('" & sSelectedWC & "','" & sOrgID  & "'," & iClass & "," & iItemCode & ",NULL," & nWCLocNo  & "," & nWCBinNo & ",'" & sMonthYear & "'," & nTransactQty & ",'" & sTransactUOM  & "',0,'" & sMixCode & "',NULL)"
					                else
						                sSqlCmd = " Update PRD_T_WIPProcessStockYRLY Set InputStock = isNull(InputStock,0) + " & nTransactQty & ", InputStockValue = isNull(InputStockValue,0) + 0 " & _
							                " where WorkCenterCode = '" & sSelectedWC & "' and OrganisationCode = '" & sOrgID  & "'" &_
							                " and MonthYear='" & sMonthYear & "' and WasteCode is NULL " &_
							                " and LocationNumber = " & nWCLocNo & " and BinNumber = " & nWCBinNo & "" &_
							                " and MixCode='" & sMixCode & "'"
					                end if
					                dRSet.Close
					                Response.Write "<p>" & sSqlCmd
					                con.execute(sSqlCmd)

					                sSqlCmd = "Update PRD_T_WIPProcessStockYRLY set ClosingStock=(isNull(OpeningStock,0)+isNull(InputStock,0))-isNull(OutputStock,0) " &_
							                 " where isNull(MonthYear,'') = '" & sMonthYear & "' and WorkCenterCode='" & sSelectedWC   &"' and OrganisationCode ='" & sOrgID & "'"
			                    Response.Write "<p>" & sSqlCmd
					                Con.Execute(sSqlCmd)

					                '----
				                end if 'if IssToCode ="PRD" then
				                ' WIP Stock Table updation - end
			                end if 'if CDbl(nTransactQty) > 0 then

                        ''end added by ragav

                   ' else

'                        if HeaderNode.hasChildNodes() then
 '                           For each ndMixData in HeaderNode.childNodes
  '                              if ndMixData.nodeName = "MixData" then
   '                                 For each ndMix in ndMixData.childNodes
    '                                    if ndMix.nodeName="Mix" then
     '                                       iMixQty = ndMix.getAttribute("Qty")
      '                                      iMixCode = ndMix.getAttribute("Code")
       '                                     if trim(iMRSNo)<>"" then
        '                                        sSql = "Insert into Inv_T_MaterialIssueAdditionalDetails (IssueEntryNo,MRSNumber,OrganisationCode,ClassificationCode,ItemCode,MixCode,QuantityIssued)"&_
         '                                              " values ("& iLedIssueNo &","& iMRSNo &",'"& sOrgID &"',"& iClass &","& iItemCode  &","& iMixCode &","& iMixQty  &")"
'         '                                       'Response.Write "<p>"&sSql
           '                                     con.Execute sSql
            '                                else
             '                                   sSql = "Insert into Inv_T_MaterialIssueAdditionalDetails (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,MixCode,QuantityIssued)"&_
              '                                         " values ("& iLedIssueNo &",'"& sOrgID &"',"& iClass &","& iItemCode  &","& iMixCode &","& iMixQty &")"
'              '                                  'Response.Write "<p>"&sSql
                '                                con.Execute sSql
                 '                           end if
                  '                      end if
                   '                 next
                    '            end if
                    '        next
                    '    end if
                   ' end if
                'end if
Response.Write "<p>Hai"



		    '	'Response.Write " Iitem Value = "& iValue

			    ' check for total qty > 0
			    if (iIssQty + iPrQty + iTraQty) > 0 then
			    
			    'Response.Write " <p> Isstype  = "& sIssType 

                iScheduleNo = 0
			    iIssAccHead = ""
			    if trim(sIssType)="F" then
			    if HeaderNode.HaschildNodes() then
				    For Each PickNode In HeaderNode.childNodes
					    if StrComp(PickNode.nodeName,"Pick") = 0 then

						    sTotPickQty = cdbl(trim(PickNode.Attributes.Item(0).nodeValue))
						    if PickNode.HaschildNodes() then
							    For Each PickDetNode In PickNode.childNodes
								    if StrComp(PickDetNode.nodeName,"PICK") = 0 then

								    ''Response.Write PickDetNode.Attributes.Item(2).nodeValue
									    sLoc	  = PickDetNode.Attributes.Item(0).nodeValue
									    sBin	  = PickDetNode.Attributes.Item(1).nodeValue
									    iLotNo	  = PickDetNode.Attributes.Item(2).nodeValue
									    iIssInvRecNo = PickDetNode.Attributes.Item(3).nodeValue
									    iQtyIss   = PickDetNode.Attributes.Item(4).nodeValue
									    if trim(iLotNo)="N/A" then iLotNo=""
									    if PickDetNode.HaschildNodes() then
										    For Each SerNode in PickDetNode.childnodes
											    if StrComp(SerNode.NodeName,"SERIALHEADER") = 0 then
												    For Each SerDetNode in SerNode.childnodes
													    if StrComp(SerDetNode.NodeName,"SERIALDETAILS") = 0 then
														    iSerNo  =  SerDetNode.Attributes.Item(0).nodeValue
														    iSerQty =  SerDetNode.Attributes.Item(1).nodeValue

														    with dcrs
															    .cursorLocation = 3
															    .cursorType = 3
															    .ActiveConnection = con
															    if trim(iSerNo)="NULL" then
																    .source = "Select InventoryReceiptNo from VW_ITEMLOCATIONLOT_STOCK where SerialNumber is null and ItemCode ="& iItemCode
															    else
																    .source = "Select InventoryReceiptNo from VW_ITEMLOCATIONLOT_STOCK where SerialNumber ="& iSerNo
															    end if
													    	Response.Write dcrs.source
															    .open
														    end with
														    if not dcrs.eof then
															    iIssInvRecNo = dcrs(0)
														    end if
														    dcrs.close

														    Dim iPackCode,iPackNo,iLotGrossQty,iLotNettQty,WPerSellForm,SellNumber,NoofSellForm
														    if cdbl(iSerQty)>0 then
															    if (trim(IssToCode)="DIS" or trim(sSubConProfoma)="Y") and 1 = 2 then
																    'if trim(sSalInvConfirm)="N" then

																	    sSql = "Select isNull(PackingCode,NULL),isNull(PackingNumber,NULL),isNull(LotQuantityGross,NULL),isNull(LotQuantityNett,NULL),isNull(WeightPerSellingForm,NULL),"&_
																	    "isNull(SellingNumber,NULL) from INV_T_LocationLot where SerialNumber = "&iSerNo
																	    Response.Write sSql

																	    rsTemp.Open sSql,con

																	    if not rsTemp.EOF then
																		    iPackCode=rsTemp(0)
																		    iPackNo =rsTemp(1)
																		    iLotGrossQty=rsTemp(2)
																		    iLotNettQty = rsTemp(3)
																		    WPerSellForm= rsTemp(4)
																		    SellNumber= rsTemp(5)

																		    if isnull(iPackCode) then iPackCode ="NULL"
																		    if IsNull(iPackNo) then iPackNo ="NULL"
																		    if IsNull(iLotGrossQty) or iLotGrossQty="" then iLotGrossQty ="0"
																		    if IsNull(iLotNettQty) or iLotNettQty then iLotNettQty = "0"
																		    if IsNull(WPerSellForm) then WPerSellForm="0"
																		    if isnull(SellNumber) then SellNumber="NULL"
'																		    'Response.Write "WPerSellForm = "& WPerSellForm
																		    
																			    if cdbl(WPerSellForm)>0 then
																				    NoofSellForm = cdbl(iLotNettQty)/cdbl(WPerSellForm)
																			    else
																				    NoofSellForm = 0
																			    end if
																			    if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
																				    sSql = "Insert into FORINVOICE_PACKDETAILS (ForInvoiceNo,ItemCode,ClassificationCode,PackingCode,"&_
																					    "PackNumber,PackGrossWeight,PackNettWeight,PackNoofsellingForm,WeightPerSellingForm,"&_
																					    "SellingNumber,InventoryReceiptSerialNo) values("& sInvNo &","& iItemCode &","&iClass&","&iPackCode&","&_
																					    "'"&iPackNo&"',"&iLotGrossQty&","& iLotNettQty &","& NoofSellForm &","& WPerSellForm &","&SellNumber&","& iSerNo &")"
																				    Response.Write "<p>"&sSql
																				    Con.execute sSql
																			    end if ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
																	    end if
																	    rsTemp.Close
																    'end if 'if trim(sSalInvConfirm)="N" then
															    end if ' if trim(IssToCode)="DIS" then
														    end if 'if cdbl(iSerQty)>0 then


													    'Response.Write "<p>iLotNo = "& iLotNo
														    IF iLotNo <> "" and iLotNo <> "NULL" and iLotNo <> "0" then

															    with dcrs
																    .CursorLocation = 3
																    .CursorType = 3
																    .Source = "SELECT ISNULL(RATE,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE INVENTORYRECEIPTNO = " & iIssInvRecNo & " AND ISNULL(ITEMENTRYNO,0) = " & iEntNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (LOTNUMBER = '" & iLotNo & "' or LotNumber is Null)"
'															    	Response.Write "<p>"&"Lot ="& dcrs.source
																    .ActiveConnection = con
																    .Open
															    end with
															    set dcrs.ActiveConnection = nothing
''															    Response.Write dcrs.Source &vbCrLf
															    if not dcrs.EOF then
																    iItmRate = cdbl(dcrs(0))
															    end if
															    dcrs.Close
															    if trim(iSerNo)="NULL" then iSerNo="0"
														    	'Response.Write " if Lot <> "" then iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo
														    'Response.Write "<p> First Updation = " 
														    'Response.write "<p> Ser Qty = " & iSerQty
														    
															    if cdbl(iSerQty)>0 then
															        if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
															            'Response.write "<p>First UpdateLocation Lot"
																        UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
																    end if
															    end if

															    with dcrs
																    .CursorLocation = 3
																    .CursorType = 3
																    .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
																    Response.Write dcrs.source
																    .ActiveConnection = con
																    .Open
															    end with
															    set dcrs.ActiveConnection = nothing

															    if not dcrs.EOF then
															    '	Response.Write " Issue Value ="& dcrs(1)
															    '	Response.Write " Issue Value ="&  dcrs(0)
															    '	Response.Write " Qty = "& iSerQty
															    '	Response.Write " iIss Value = "& iIssVal
															    	if cdbl(dcrs(1))>0 then
															            iIssVal = iSerQty * Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
														            else
														                iIssVal = 0
														            end if
															    end if
															    dcrs.Close
    															if trim(iIssVal)="" then
															        iIssVal = iSerQty * iItmRate
															    end if  'if trim(iIssVal)="" then
															    
															    'Response.Write "<p>IssValue = "& iIssVal
															    'Response.Write iLotNo &" ** "&iSerNo  &" ** "&	iSerQty
															    

															    IF cdbl(iSerQty)>0 then
															    
															    sTempLotNo = iLotNo
															    if trim(sTempLotNo)="" or isNull(sTempLotNo) then sTempLotNo = "NULL"
															    if trim(sTempLotNo)<>"NULL"  then sTempLotNo = pack(sTempLotNo)
															    
															    
															    
																    
																    if sAutoConsumption = "Y" then
																        sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
																		        " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,QuantityConsumed,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
																		        "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
																		        " " & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iSerQty &","& iSerQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
															    			    'Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
																		        Con.Execute sSql
															        else
															            sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
																		        " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
																		        "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
																		        " " & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iSerQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
															    			    'Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
																		        Con.Execute sSql
															        end if

		 													    End IF
		 													    Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
		 													    
		 													    ''Added by Ragav on Oct 20,2012 for POSSTOCKDETAILS Update on POS IssueCase
		 													    if trim(IssToCode)="POS" and 1 = 2 then
		 													        sSql = "Select InventoryReceiptNo,isNull(PackingNumber,0),isNull(PackingCode,0),isNull(PackingSubLevelID,0),isNull(PackingSubLevelQty,0),isNull(PackingSubLevelUnitQty,0),isNull(SellingNumber,0),"
		 													        sSql = sSql &" isNull(WeightPerSellingForm,0),isNull(SellingForm,0),isNull(Rate,0),AttributeList,Convert(varchar,DateOfReceipt,103),SrcTypeCode from INV_T_LocationLot where SerialNumber = "& iSerNo
		 													        dcrs.open sSql,con
		 													        sSql = ""
		 													        if not dcrs.eof then
		 													            sTempAttList = dcrs(10)
		 													            if trim(sTempAttList)="" or IsNull(sTempAttList) then sTempAttList = "NULL"
		 													            sSql = " Insert into INV_T_POSStockDetails (POSInvReceiptNo,OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,PackingNumber,PackingCode,"
                                                                        sSql = sSql &" PackingSubLevelID,PackingSubLevelQty,PackingSubLevelUnitQty,SellingNumber,WeightPerSellingForm,SellingForm,"
                                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,SrcTypeCode,POSID,CreatedBy,CreatedOn)"
                                                                        sSql = sSql &" Values ("& dcrs(0) &","& Pack(sOrgID) &"," & iItemCode & ","& iClass &","& iEntNo &","
                                                                        sSql = sSql &""& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iSerQty &",0,"& dcrs(1) &","& dcrs(2) &","
                                                                        sSql = sSql &""& dcrs(3)&","& dcrs(4) &","& dcrs(5) &","& dcrs(6) &","& dcrs(7) &","& dcrs(8) &","
                                                                        sSql = sSql &""& dcrs(9) &","& sTempAttList &",Convert(datetime,'"& dcrs(11) &"',103),"& dcrs(12) &","& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
		 													        else
		 													            sSql = " Insert into INV_T_POSStockDetails (OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,"
                                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,POSID,CreatedBy,CreatedOn)"
                                                                        sSql = sSql &" Values ("& pack(sOrgID) &","& iItemCode &","& iClass &","& iEntNo &","
                                                                        sSql = sSql &" "& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iSerQty &",0,"
                                                                        sSql = sSql &" "&iItmRate&","& sAttID &",Convert(datetime,'"& date() &"',103),"& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
		 													        end if
		 													        if trim(sSql)<>"" then
		 													            Response.write "<p>"& sSql
		 													            con.execute sSql
		 													        end if
		 													        dcrs.close
		 													    end if
		 													    ''end
		 													    
		 													    
		 													    if sAutoConsumption = "Y" then
		 													        sSql = "Insert into INV_T_MaterialConsumptionDetail (LineNumber,ConsumptionNo,IssueEntryNo,"&_
		 													               " IssueDate,LotNo,SerialNo,QuantityConsumed,ConsumptionACHead,AttributeList) "&_
		 													               " values("& sConLineNo &","& sConEntryNo &","& iLedIssueNo &",Convert(datetime,getDate(),103),"&_
		 													               " '"& iLotNo &"',"& iSerNo &","& iSerQty &","& sConAccHead &","& sAttID &")"
                                                                    Response.Write "<p> Auto Consumption  header =  "& sSql &"<p>"
                                                                    con.execute sSql 
                                                                end if


														    Else
															    with dcrs
																    .CursorLocation = 3
																    .CursorType = 3
																    .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) and LocationNumber = "& sLoc
							    									Response.Write dcrs.source
																    .ActiveConnection = con
																    .Open
															    end with
															    set dcrs.ActiveConnection = nothing

															    if not dcrs.EOF then
															    'Response.write dcrs.source
															        if cdbl(dcrs(0))>0 then
																        iIssVal = iSerQty * Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
																    end if
															    end if
															    dcrs.Close
															   ' Response.Write "iIssVal = " & iIssVal 
															    IF iSerQty >0  then
																    ''Response.Write iLotNo &" ** "&iSerNo  &" ** "&	iSerQty
																    
															    sTempLotNo = iLotNo
															    if trim(sTempLotNo)="" or isNull(sTempLotNo) then sTempLotNo = "NULL"
															    if trim(sTempLotNo)<>"NULL"  then sTempLotNo = pack(sTempLotNo)
															    if Trim(iIssVal)="" or IsNull(iIssVal) then iIssVal = "0"
															    
																    if sAutoConsumption = "Y" then
																        sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
																		        " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,QuantityConsumed,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
																		        "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
																		        " " & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iSerQty &","& iSerQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
																		         Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
																		        Con.Execute sSql
															        else
															            sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
																		        " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
																		        "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
																		        " " & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iSerQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
																		         Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
																		        Con.Execute sSql
															        end if
                                                                            'Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
		 													    End IF

		 													    ''Added By Ragav on Apr 01 ,2010
		 													    if trim(iSerNo)="NULL" then iSerNo="0"
															    'Response.Write " if Lot <> "" then iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo
															    'Response.Write "<p>Second Updation = "
															    if cdbl(iSerQty)>0 then
															        if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
															        'Response.write "<p>Second UpdateLocation Lot"
																        UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
																    end if
															    end if
															    
															    
															    ''Added by Ragav on Oct 20,2012 for POSSTOCKDETAILS Update on POS IssueCase
		 													    if trim(IssToCode)="POS" and 1 = 2 then
		 													        sSql = "Select InventoryReceiptNo,isNull(PackingNumber,0),isNull(PackingCode,0),isNull(PackingSubLevelID,0),isNull(PackingSubLevelQty,0),isNull(PackingSubLevelUnitQty,0),isNull(SellingNumber,0),"
		 													        sSql = sSql &" isNull(WeightPerSellingForm,0),isNull(SellingForm,0),isNull(Rate,0),AttributeList,Convert(varchar,DateOfReceipt,103),SrcTypeCode from INV_T_LocationLot where SerialNumber = "& iSerNo
		 													        dcrs.open sSql,con
		 													        sSql = ""
		 													        if not dcrs.eof then
		 													            
		 													            'Response.write "<p>Attribute List = "& dcrs(10)
		 													            sTempAttList = dcrs(10)
		 													            if trim(sTempAttList)="" or IsNull(sTempAttList) then sTempAttList = "NULL"
		 													            
		 													            sSql = " Insert into INV_T_POSStockDetails (POSInvReceiptNo,OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,PackingNumber,PackingCode,"
                                                                        sSql = sSql &" PackingSubLevelID,PackingSubLevelQty,PackingSubLevelUnitQty,SellingNumber,WeightPerSellingForm,SellingForm,"
                                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,SrcTypeCode,POSID,CreatedBy,CreatedOn)"
                                                                        sSql = sSql &" Values ("& dcrs(0) &","& Pack(sOrgID) &"," & iItemCode & ","& iClass &","& iEntNo &","
                                                                        sSql = sSql &""& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iSerQty &",0,"& dcrs(1) &","& dcrs(2) &","
                                                                        sSql = sSql &""& dcrs(3)&","& dcrs(4) &","& dcrs(5) &","& dcrs(6) &","& dcrs(7) &","& dcrs(8) &","
                                                                        sSql = sSql &""& dcrs(9) &","& sTempAttList &",Convert(datetime,'"& dcrs(11) &"',103),"& dcrs(12) &","& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
		 													        else
		 													            sSql = " Insert into INV_T_POSStockDetails (OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,"
                                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,POSID,CreatedBy,CreatedOn)"
                                                                        sSql = sSql &" Values ("& pack(sOrgID) &","& iItemCode &","& iClass &","& iEntNo &","
                                                                        sSql = sSql &" "& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iSerQty &",0,"
                                                                        sSql = sSql &" "&iItmRate&","& sAttID &",Convert(datetime,'"& date() &"',103),"& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
		 													        end if
		 													        if trim(sSql)<>"" then
		 													            Response.write "<p>"& sSql
		 													            con.execute sSql
		 													        end if
		 													        dcrs.close
		 													    end if
		 													    ''end
															    

														    End if ' IF iLotNo <> "" and iLotNo <> "0" then


													    end if 'if StrComp(SerDetNode.NodeName,"SERIALDETAILS") = 0 then
												    Next ' For Each SerDetNode in SerNode.childnodes
											    end if 'if StrComp(SerNode.NodeName,"SERIALHEADER") = 0 then
										    Next 'For Each SerNode in PickDetNode.childnodes
										else
										    'Response.Write "<p>Hello Store"
									        
									          if iQtyIss <>"0" then
									            'Response.Write "<p> Loc = "& sLoc
									            'Response.Write "<p> Bin = "& sBin
    									        
										            with dcrs
												        .CursorLocation = 3
												        .CursorType = 3
												        .Source = "SELECT ISNULL(MAX(ISSUEENTRYNO)+1,1) FROM INV_T_MATERIALISSUEDETAILS"
												        ''Response.Write dcrs.source
												        .ActiveConnection = con
												        .Open
											        end with
											        if not dcrs.EOF then
												        iIssEntNo = dcrs(0)
											        END IF
											        dcrs.close
											        IF ISNULL(sLoc) OR sLoc="" THEN sLoc = "NULL"
											        IF IsNull(sBin) OR sBin="" THEN sBin = "NULL"
											        if iLotNo="" or isnull(iLotNo) then iLotNo ="NULL"
											        if iSerNo ="" or isnull(iSerNo) then iSerNo ="NULL"


											        with dcrs
												        .CursorLocation = 3
												        .CursorType = 3
												        .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
												        Response.Write "<p>"&dcrs.source
												        .ActiveConnection = con
												        .Open
											        end with
											        set dcrs.ActiveConnection = nothing

											        if not dcrs.EOF then
												        if cdbl(dcrs(0))>cdbl(0) then
												            iItmRate = Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
													        iIssVal = iSerQty * iItmRate
												        end if
											        end if
											        
											        dcrs.Close
											        
											        if trim(iLotNo)="N/A" then iLotNo="NULL"
		 									        
											        if IsNull(iIssVal) or trim(iIssVal)="" then iIssVal = 0
											        
											        sTempLotNo = iLotNo
												    if trim(sTempLotNo)="" or isNull(sTempLotNo) then sTempLotNo = "NULL"
												    if trim(sTempLotNo)<>"NULL"  then sTempLotNo = pack(sTempLotNo)
											        
										                if sAutoConsumption ="Y" then
									    	                sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,QuantityConsumed,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
													                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													                'Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                Con.Execute sSql
    														        
												        else
												            sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
													                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
'													                Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                Con.Execute sSql
												        end if
    														    Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
		 									        ''Added By Ragav on Apr 01 ,2010

		 									       ' if trim(iLotNo)="NULL" then iLotNo="0"
		 									       ' if trim(iSerNo)="NULL" then iSerNo="0"
										    	'    Response.Write "<p> if not isNull(Lot) then iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo &" sLoc  = "& sLoc
											    '    if cdbl(iIssQty)>0 then
												'       UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
											    '    end if
											    
											    if trim(sAttID)="" or isNull(sAttID) then sAttID="NULL"
											    ''Added by Ragav on Oct 20,2012 for POSSTOCKDETAILS Update on POS IssueCase
											    if trim(IssToCode)="POS" and 1 = 2 then
											        sSql = "Select InventoryReceiptNo,isNull(PackingNumber,0),isNull(PackingCode,0),isNull(PackingSubLevelID,0),isNull(PackingSubLevelQty,0),isNull(PackingSubLevelUnitQty,0),isNull(SellingNumber,0),"
											        sSql = sSql &" isNull(WeightPerSellingForm,0),isNull(SellingForm,0),isNull(Rate,0),AttributeList,Convert(varchar,DateOfReceipt,103),SrcTypeCode from INV_T_LocationLot where SerialNumber = "& iSerNo
											        dcrs.open sSql,con
											        sSql = ""
											        if not dcrs.eof then
											            
											            sTempAttList = dcrs(10)
		 												if trim(sTempAttList)="" or IsNull(sTempAttList) then sTempAttList = "NULL"
		 													            
											            sSql = " Insert into INV_T_POSStockDetails (POSInvReceiptNo,OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,PackingNumber,PackingCode,"
                                                        sSql = sSql &" PackingSubLevelID,PackingSubLevelQty,PackingSubLevelUnitQty,SellingNumber,WeightPerSellingForm,SellingForm,"
                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,SrcTypeCode,POSID,CreatedBy,CreatedOn)"
                                                        sSql = sSql &" Values ("& dcrs(0) &","& Pack(sOrgID) &"," & iItemCode & ","& iClass &","& iEntNo &","
                                                        sSql = sSql &""& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iIssQty &",0,"& dcrs(1) &","& dcrs(2) &","
                                                        sSql = sSql &""& dcrs(3)&","& dcrs(4) &","& dcrs(5) &","& dcrs(6) &","& dcrs(7) &","& dcrs(8) &","
                                                        sSql = sSql &""& dcrs(9) &","& sTempAttList &",Convert(datetime,'"& dcrs(11) &"',103),"& dcrs(12) &","& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
											        else
											            sSql = " Insert into INV_T_POSStockDetails (OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,"
                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,POSID,CreatedBy,CreatedOn)"
                                                        sSql = sSql &" Values ("& pack(sOrgID) &","& iItemCode &","& iClass &","& iEntNo &","
                                                        sSql = sSql &" "& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iIssQty &",0,"
                                                        sSql = sSql &" "&iItmRate&","& sAttID &",Convert(datetime,'"& date() &"',103),"& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
											        end if
											        if trim(sSql)<>"" then
											            Response.write "<p>"& sSql
											            con.execute sSql
											        end if
											        dcrs.close
											    end if
											    ''end
											    
									        end if 'if trim(iQtyIss)<>"0" then
									    end if 'if PickDetNode.HaschildNodes() then
									    
									elseif StrComp(PickDetNode.nodeName,"STORE") = 0 then    
									        Response.Write "<p>Hello Store"
									        
									        sLoc	  = PickDetNode.Attributes.Item(0).nodeValue
									        sBin	  = PickDetNode.Attributes.Item(1).nodeValue
									        iLotNo	  = PickDetNode.Attributes.Item(2).nodeValue
									        iIssInvRecNo = PickDetNode.Attributes.Item(3).nodeValue
									        iQtyIss   = PickDetNode.Attributes.Item(4).nodeValue
									        if iQtyIss <>"0" then
									            'Response.Write "<p> Loc = "& sLoc
									            'Response.Write "<p> Bin = "& sBin
    									        
										            with dcrs
												        .CursorLocation = 3
												        .CursorType = 3
												        .Source = "SELECT ISNULL(MAX(ISSUEENTRYNO)+1,1) FROM INV_T_MATERIALISSUEDETAILS"
												        ''Response.Write dcrs.source
												        .ActiveConnection = con
												        .Open
											        end with
											        if not dcrs.EOF then
												        iIssEntNo = dcrs(0)
											        END IF
											        dcrs.close
											        IF ISNULL(sLoc) OR sLoc="" THEN sLoc = "NULL"
											        IF IsNull(sBin) OR sBin="" THEN sBin = "NULL"
											        if iLotNo="" or isnull(iLotNo) then iLotNo ="NULL"
											        if iSerNo ="" or isnull(iSerNo) then iSerNo ="NULL"


											        with dcrs
												        .CursorLocation = 3
												        .CursorType = 3
												        .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
												        Response.Write "<p>"&dcrs.source
												        .ActiveConnection = con
												        .Open
											        end with
											        set dcrs.ActiveConnection = nothing

											        if not dcrs.EOF then
												        if cdbl(dcrs(0))>cdbl(0) then
													        iIssVal = iSerQty * Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
												        end if
											        end if
											        
											        dcrs.Close
											        
											        if trim(iLotNo)="N/A" then iLotNo="NULL"
		 									        
											        if IsNull(iIssVal) or trim(iIssVal)="" then iIssVal = 0
											        
											        sTempLotNo = iLotNo
												    if trim(sTempLotNo)="" or isNull(sTempLotNo) then sTempLotNo = "NULL"
												    if trim(sTempLotNo)<>"NULL"  then sTempLotNo = pack(sTempLotNo)
											        
										                if sAutoConsumption ="Y" then
									    	                sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,QuantityConsumed,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
													                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													                'Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                Con.Execute sSql
    														        
												        else
												            sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
													                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													                Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                Con.Execute sSql
												        end if
												        Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
    														    
		 									        ''Added By Ragav on Apr 01 ,2010

		 									        if trim(iLotNo)="NULL" then iLotNo="0"
		 									        if trim(iSerNo)="NULL" then iSerNo="0"
										    	'    Response.Write "<p> if not isNull(Lot) then iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo &" sLoc  = "& sLoc
											    '    if cdbl(iIssQty)>0 then
												'       UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
											    '    end if
											    
											    
											    ''Added by Ragav on Oct 20,2012 for POSSTOCKDETAILS Update on POS IssueCase
											    if trim(IssToCode)="POS" and 1 = 2 then
											        sSql = "Select InventoryReceiptNo,isNull(PackingNumber,0),isNull(PackingCode,0),isNull(PackingSubLevelID,0),isNull(PackingSubLevelQty,0),isNull(PackingSubLevelUnitQty,0),isNull(SellingNumber,0),"
											        sSql = sSql &" isNull(WeightPerSellingForm,0),isNull(SellingForm,0),isNull(Rate,0),AttributeList,Convert(varchar,DateOfReceipt,103),SrcTypeCode from INV_T_LocationLot where SerialNumber = "& iSerNo
											        dcrs.open sSql,con
											        sSql = ""
											        if not dcrs.eof then
											            sTempAttList = dcrs(10)
		 												if trim(sTempAttList)="" or IsNull(sTempAttList) then sTempAttList = "NULL"
		 													            
											            sSql = " Insert into INV_T_POSStockDetails (POSInvReceiptNo,OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,PackingNumber,PackingCode,"
                                                        sSql = sSql &" PackingSubLevelID,PackingSubLevelQty,PackingSubLevelUnitQty,SellingNumber,WeightPerSellingForm,SellingForm,"
                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,SrcTypeCode,POSID,CreatedBy,CreatedOn)"
                                                        sSql = sSql &" Values ("& dcrs(0) &","& Pack(sOrgID) &"," & iItemCode & ","& iClass &","& iEntNo &","
                                                        sSql = sSql &""& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iIssQty &",0,"& dcrs(1) &","& dcrs(2) &","
                                                        sSql = sSql &""& dcrs(3)&","& dcrs(4) &","& dcrs(5) &","& dcrs(6) &","& dcrs(7) &","& dcrs(8) &","
                                                        sSql = sSql &""& dcrs(9) &","& sTempAttList &",Convert(datetime,'"& dcrs(11) &"',103),"& dcrs(12) &","& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
											        else
											            sSql = " Insert into INV_T_POSStockDetails (OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,"
                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,POSID,CreatedBy,CreatedOn)"
                                                        sSql = sSql &" Values ("& pack(sOrgID) &","& iItemCode &","& iClass &","& iEntNo &","
                                                        sSql = sSql &" "& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iIssQty &",0,"
                                                        sSql = sSql &" "&iItmRate&","& sAttID &",Convert(datetime,'"& date() &"',103),"& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
											        end if
											        if trim(sSql)<>"" then
											            Response.write "<p>"& sSql
											            con.execute sSql
											        end if
											        dcrs.close
											    end if
											    ''end
											    
									        end if 'if trim(iQtyIss)<>"0" then
									end if ' if StrComp(PickDetNode..nodeName,"PICK") = 0 then
							    Next 'For Each PickDetNode In PickNode.childNodes
						    ELSE 'if PickNode.HaschildNodes() then
							    			'Response.Write "HEello"
										    with dcrs
												    .CursorLocation = 3
												    .CursorType = 3
												    .Source = "SELECT ISNULL(MAX(ISSUEENTRYNO)+1,1) FROM INV_T_MATERIALISSUEDETAILS"
												    ''Response.Write dcrs.source
												    .ActiveConnection = con
												    .Open
											    end with
											    if not dcrs.EOF then
												    iIssEntNo = dcrs(0)
											    END IF
											    dcrs.close
											    IF ISNULL(sLoc) OR sLoc="" THEN sLoc = "NULL"
											    IF IsNull(sBin) OR sBin="" THEN sBin = "NULL"
											    if iLotNo="" or isnull(iLotNo) then iLotNo ="NULL"
											    if iSerNo ="" or isnull(iSerNo) then iSerNo ="NULL"


											    with dcrs
												    .CursorLocation = 3
												    .CursorType = 3
												    .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
												    Response.Write "<p>"&dcrs.source
												    .ActiveConnection = con
												    .Open
											    end with
											    set dcrs.ActiveConnection = nothing

											    if not dcrs.EOF then
												    if cdbl(dcrs(0))>cdbl(0) then
													    iIssVal = iSerQty * Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
												    end if
											    end if
											    dcrs.Close


											    if IsNull(iIssVal) or trim(iIssVal)="" then iIssVal = 0



										    '	IF cdbl(iIssVal) <> cdbl(0) then
												    ''Response.Write iLotNo &" ** "&iSerNo  &" ** "&	iSerQty
												    
												    sTempLotNo = iLotNo
												    if trim(sTempLotNo)="" or isNull(sTempLotNo) then sTempLotNo = "NULL"
												    if trim(sTempLotNo)<>"NULL"  then sTempLotNo = pack(sTempLotNo)
												    
												    if sAutoConsumption = "Y" then
												        sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
														        " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,QuantityConsumed,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
														        "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
														        "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
														        Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
														        Con.Execute sSql
												    else
												            sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
														        " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
														        "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
														        "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
														        Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
														        Con.Execute sSql
												    end if
												    'Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf

		 								    '	End IF
		 								    

		 									    ''Added By Ragav on Apr 01 ,2010

		 									    if trim(iLotNo)="NULL" then iLotNo="0"
		 									    if trim(iSerNo)="NULL" then iSerNo="0"
										    '	'Response.Write " if Lot <> "" then iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo
										    'response.write "<p>Third Updation = "
											    if cdbl(iIssQty)>0 then
											        if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
											        Response.write "<p>3rd UpdateLocation Lot"
												        UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
												    end if' if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
											    end if
											    
											    ''Added by Ragav on Oct 20,2012 for POSSTOCKDETAILS Update on POS IssueCase
											    if trim(IssToCode)="POS" and 1 = 2 then
											        sSql = "Select InventoryReceiptNo,isNull(PackingNumber,0),isNull(PackingCode,0),isNull(PackingSubLevelID,0),isNull(PackingSubLevelQty,0),isNull(PackingSubLevelUnitQty,0),isNull(SellingNumber,0),"
											        sSql = sSql &" isNull(WeightPerSellingForm,0),isNull(SellingForm,0),isNull(Rate,0),AttributeList,Convert(varchar,DateOfReceipt,103),SrcTypeCode from INV_T_LocationLot where SerialNumber = "& iSerNo
											        dcrs.open sSql,con
											        sSql = ""
											        if not dcrs.eof then
											            sTempAttList = dcrs(10)
		 												if trim(sTempAttList)="" or IsNull(sTempAttList) then sTempAttList = "NULL"
		 												
											            sSql = " Insert into INV_T_POSStockDetails (POSInvReceiptNo,OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,PackingNumber,PackingCode,"
                                                        sSql = sSql &" PackingSubLevelID,PackingSubLevelQty,PackingSubLevelUnitQty,SellingNumber,WeightPerSellingForm,SellingForm,"
                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,SrcTypeCode,POSID,CreatedBy,CreatedOn)"
                                                        sSql = sSql &" Values ("& dcrs(0) &","& Pack(sOrgID) &"," & iItemCode & ","& iClass &","& iEntNo &","
                                                        sSql = sSql &""& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iIssQty &",0,"& dcrs(1) &","& dcrs(2) &","
                                                        sSql = sSql &""& dcrs(3)&","& dcrs(4) &","& dcrs(5) &","& dcrs(6) &","& dcrs(7) &","& dcrs(8) &","
                                                        sSql = sSql &""& dcrs(9) &","& sTempAttList &",Convert(datetime,'"& dcrs(11) &"',103),"& dcrs(12) &","& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
											        else
											            sSql = " Insert into INV_T_POSStockDetails (OrganisationCode,ItemCode,ClassificationCode,ItemEntryNo, "
                                                        sSql = sSql &" StorageLocationNo,StorageBinNumber,LotNumber,SerialNumber,NetQuantity,QuantityIssued,"
                                                        sSql = sSql &" ItemRate,AttributeList,DateOfReceipt,POSID,CreatedBy,CreatedOn)"
                                                        sSql = sSql &" Values ("& pack(sOrgID) &","& iItemCode &","& iClass &","& iEntNo &","
                                                        sSql = sSql &" "& sLoc &","& sBin &","& sTempLotNo &","& iSerNo &","& iIssQty &",0,"
                                                        sSql = sSql &" "&iItmRate&","& sAttID &",Convert(datetime,'"& date() &"',103),"& sIssuePOSID &","& iIssuedBy &",Convert(datetime,'"& date() &"',103))"
											        end if
											        if trim(sSql)<>"" then
											            Response.write "<p>"& sSql
											            con.execute sSql
											        end if
											        dcrs.close
											    end if
											    ''end

						    end if 'if PickNode.HaschildNodes() then
					    elseif strcomp(PickNode.nodeName,"AddDet")=0 then
					        sEntryNo = 0
						    set AddDetNode = PickNode
						    for each WorkCenterNode in AddDetNode.childnodes
							    sWCode = WorkCenterNode.getAttribute("WCODE")
							    for each MachineCenterNode in WorkCenterNode.childNodes
								    sMCode = MachineCenterNode.getAttribute("MCODE")
								    iMCQty = MachineCenterNode.getAttribute("QTY")
								    if sMCode ="select" or sMCode ="" or IsNull(sMCode) then sMCode ="NULL"
								    if sMCode <>"NULL" then sMCode = pack(sMCode)
								    if sWCode<>"" then sWCode = Pack(sWCode)
								    sSql = "Insert into Inv_T_MaterialIssueAdditionalDetails (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,WorkCenterCode,MachineCenterCode,QuantityIssued)"&_
									    " values("& iLedIssueNo  &","&pack(sOrgID)&","&iClass&","&iItemCode&","& sWCode &","& sMCode &","& iMCQty &")"
								    	Response.Write "<p>"&sSql
								    con.execute sSql
								    
								    if sAutoConsumption = "Y" then
								        sEntryNo = sEntryNo + 1
								        sSql = " Insert into INV_T_MaterialConsumptionIssueAddnDet (LineNumber,ConsumptionNo,EntryNo,"&_
								               " WorkCenterCode,MachineCenterCode,MixCode,QuantityIssued,QuantityConsumed) values "&_
								               " ("& sConLineNo &","& sConEntryNo &","& sEntryNo &","& sWCode &","& sMCode &",NULL,"& iMCQty &","& iMCQty  &")"
								        Response.Write "<p> Auto Consumption  header =  "& sSql &"<p>"
                                        con.execute sSql 
                                    end if
								    
							    next
						    next
					    end if 'if StrComp(PickNode.nodeName,"Pick") = 0 then
				    Next 'For Each PickNode In HeaderNode.childNodes
			    end if 	'if HeaderNode.HaschildNodes() then
			  else ' if trim(sIssType)="M" then
			  
			    response.write "<p>Welcome to Marked"
			            if HeaderNode.HaschildNodes() then
        				    For Each PickNode In HeaderNode.childNodes
		        			    if StrComp(PickNode.nodeName,"Pick") = 0 then
        						    sTotPickQty = cdbl(trim(PickNode.Attributes.Item(0).nodeValue))
		            				    if PickNode.HaschildNodes() then
					            		    For Each PickDetNode In PickNode.childNodes
								                if StrComp(PickDetNode.nodeName,"PICK") = 0 then
								                ''Response.Write PickDetNode.Attributes.Item(2).nodeValue
									                sLoc	  = PickDetNode.Attributes.Item(0).nodeValue
									                sBin	  = PickDetNode.Attributes.Item(1).nodeValue
									                iLotNo	  = PickDetNode.Attributes.Item(2).nodeValue
									                iIssInvRecNo = PickDetNode.Attributes.Item(3).nodeValue
									                iQtyIss   = PickDetNode.Attributes.Item(4).nodeValue
    									            
									                if trim(iLotNo)="N/A" then iLotNo="NULL"
									                if PickDetNode.hasChildNodes() then
									                
									               	    for each SerNode in PickDetNode.childNodes
														    if SerNode.nodeName="SERIALHEADER" then
															    for each SerDetNode in SerNode.childNodes
																    if SerDetNode.nodeName="SERIALDETAILS" then
																	    iSerNo  =  SerDetNode.Attributes.Item(0).nodeValue
																	    iSerQty =  SerDetNode.Attributes.Item(1).nodeValue
																	    
																	    'Response.write "<p>iSer No = "& iSerNo
    																	
																	    with dcrs
																	        .CursorLocation = 3
																	        .CursorType = 3
																	        .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
																	        Response.Write "<p>"&dcrs.source
																	        .ActiveConnection = con
																	        .Open
																	    end with
																	    set dcrs.ActiveConnection = nothing

																	    if not dcrs.EOF then
																	        if cdbl(dcrs(0))>cdbl(0) then
																		        iIssVal = iSerQty  * Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
																	        end if
																	    end if
																	    dcrs.Close
																	    
																	    'Response.write "<p>IssValue = "&iIssVal
    									            
									            					   ' IF cdbl(iIssVal) <> cdbl(0) then
																	        ''Response.Write iLotNo &" ** "&iSerNo  &" ** "&	iSerQty
																	        if trim(iLotNo)<>"NULL" then
																	            sSql =	" INSERT INTO INV_T_MaterialIssuedForPick (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                                " SerialNo,LocationNumber,BinNumber,QuantityForPick,ItemEntryNo,ItemAttributes,FlagLotOrPack,Returnable,ReturnItem,MaterialType) VALUES "&_
													                                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                                "'" & iLotNo & "'," & iSerNo & "," & sLoc & "," & sBin & "," & iSerQty &"," & iEntNo & "," & sAttID & ","& sOnlyLotFlag &",'"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													                                Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                                Con.Execute sSql
													                        else
													                            sSql =	" INSERT INTO INV_T_MaterialIssuedForPick (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                                " SerialNo,LocationNumber,BinNumber,QuantityForPick,ItemEntryNo,ItemAttributes,FlagLotOrPack,Returnable,ReturnItem,MaterialType) VALUES "&_
													                                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                                "" & iLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iSerQty &"," & iEntNo & "," & sAttID & ","& sOnlyLotFlag &",'"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													                                Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                                Con.Execute sSql
													                        end if'if trim(iLotNo)<>"NULL" then
 																	   ' End IF
																    end if
															    next
														    end if
													    next
													else ' if not PickDetNode.hasChildNodes() then
													    'if trim(iSerNo)="" or isNull(iSerNo) then iSerNo = "NULL"
													    iSerNo = "NULL"
													    if cdbl(iQtyIss)>0 then
													        if trim(iLotNo)<>"NULL" then
												                sSql =	" INSERT INTO INV_T_MaterialIssuedForPick (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
												                                " SerialNo,LocationNumber,BinNumber,QuantityForPick,ItemEntryNo,ItemAttributes,FlagLotOrPack,Returnable,ReturnItem,MaterialType) VALUES "&_
												                                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
												                                "'" & iLotNo & "'," & iSerNo & "," & sLoc & "," & sBin & "," & iQtyIss &"," & iEntNo & "," & sAttID & ","& sOnlyLotFlag &",'"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
												                                Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
												                                Con.Execute sSql
												            else
												                sSql =	" INSERT INTO INV_T_MaterialIssuedForPick (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
												                                " SerialNo,LocationNumber,BinNumber,QuantityForPick,ItemEntryNo,ItemAttributes,FlagLotOrPack,Returnable,ReturnItem,MaterialType) VALUES "&_
												                                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
												                                "" & iLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iQtyIss &"," & iEntNo & "," & sAttID & ","& sOnlyLotFlag &",'"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
												                                Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
												                                Con.Execute sSql
												            end if
													    end if'if cdbl(iQtyIss)>0 then
									                end if 'if PickDetNode.hasChildNodes() then
									        elseif StrComp(PickDetNode.nodeName,"STORE") = 0 then    
									                    'Response.Write "<p>Hello Store"
            									        
									                    sLoc	  = PickDetNode.Attributes.Item(0).nodeValue
									                    sBin	  = PickDetNode.Attributes.Item(1).nodeValue
									                    iLotNo	  = PickDetNode.Attributes.Item(2).nodeValue
									                    iIssInvRecNo = PickDetNode.Attributes.Item(3).nodeValue
									                    iQtyIss   = PickDetNode.Attributes.Item(4).nodeValue
									                    if iQtyIss <>"0" then
									                        'Response.Write "<p> Loc = "& sLoc
									                        'Response.Write "<p> Bin = "& sBin
                									        
										                        with dcrs
												                    .CursorLocation = 3
												                    .CursorType = 3
												                    .Source = "SELECT ISNULL(MAX(ISSUEENTRYNO)+1,1) FROM INV_T_MaterialIssuedForPick"
												                    Response.Write dcrs.source
												                    .ActiveConnection = con
												                    .Open
											                    end with
											                    if not dcrs.EOF then
												                    iIssEntNo = dcrs(0)
											                    END IF
											                    dcrs.close
											                    IF ISNULL(sLoc) OR sLoc="" THEN sLoc = "NULL"
											                    IF IsNull(sBin) OR sBin="" THEN sBin = "NULL"
											                    if iLotNo="" or isnull(iLotNo) then iLotNo ="NULL"
											                    if iSerNo ="" or isnull(iSerNo) then iSerNo ="NULL"


											                    with dcrs
												                    .CursorLocation = 3
												                    .CursorType = 3
												                    .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
												                    Response.Write "<p>"&dcrs.source
												                    .ActiveConnection = con
												                    .Open
											                    end with
											                    set dcrs.ActiveConnection = nothing

											                    if not dcrs.EOF then
												                    if cdbl(dcrs(0))>cdbl(0) then
													                    iIssVal = iSerQty * Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
												                    end if
											                    end if
            											        
											                    dcrs.Close
            											        
											                    if trim(iLotNo)="N/A" then iLotNo="NULL"
            		 									        
											                    if IsNull(iIssVal) or trim(iIssVal)="" then iIssVal = 0
            											        
										                                sSql =	" INSERT INTO INV_T_MaterialIssuedForPick (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                            " SerialNo,LocationNumber,BinNumber,QuantityForPick,ItemEntryNo,ItemAttributes,FlagLotOrPack,Returnable,ReturnItem,MaterialType) VALUES "&_
													                            "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                            "" & iLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssQty &"," & iEntNo & "," & sAttID & ","& sOnlyLotFlag &",'"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													                            Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                            Con.Execute sSql
                														    
		 									                    ''Added By Ragav on Apr 01 ,2010

		 									                    if trim(iLotNo)="NULL" then iLotNo="0"
		 									                    if trim(iSerNo)="NULL" then iSerNo="0"
										    	            '    Response.Write "<p> if not isNull(Lot) then iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo &" sLoc  = "& sLoc
											                '    if cdbl(iIssQty)>0 then
												            '       UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
											                '    end if
									                    end if 'if trim(iQtyIss)<>"0" then
									        elseif StrComp(PickDetNode.nodeName,"PickSchedule") = 0 then    
									            iScheduleNo = iScheduleNo + 1
									            dSchedule = PickDetNode.getAttribute("Date")
									            iScheduleQty = PickDetNode.getAttribute("Qty")
        									    
									                sSql = "Insert into Inv_T_IssueForPickSchedule (IssueEntryNo,ItemEntryNo,ItemCode,"
									                sSql = sSql & " ClassificationCode,ItemAttributes,ScheduleNo,ScheduledOn,ScheduledQty)"
									                sSql = sSql & " values("& iLedIssueNo &","& iEntNo &","& iItemCode &","& iClass &","
									                sSql = sSql & ""& sAttID &","& iScheduleNo &",Convert(datetime,'"& dSchedule &"',103),"& iScheduleQty &")"
									                response.write "<p>"& sSql
									                con.execute sSql
									        
								    
								            end if ' if StrComp(PickDetNode..nodeName,"PICK") = 0 then 
					    		        Next 'For Each PickDetNode In PickNode.childNodes
						        ELSE 'if PickNode.HaschildNodes() then
						        'Response.Write " ///////PICK NOt Have Child /////"
										    with dcrs
												    .CursorLocation = 3
												    .CursorType = 3
												    .Source = "SELECT ISNULL(MAX(ISSUEENTRYNO)+1,1) FROM INV_T_MaterialIssuedForPick"
												    Response.Write dcrs.source
												    .ActiveConnection = con
												    .Open
											    end with
											    if not dcrs.EOF then
												    iIssEntNo = dcrs(0)
											    END IF
											    dcrs.close
											    IF ISNULL(sLoc) OR sLoc="" THEN sLoc = "NULL"
											    IF IsNull(sBin) OR sBin="" THEN sBin = "NULL"
											    if iLotNo="" or isnull(iLotNo) then iLotNo ="NULL"
											    if iSerNo ="" or isnull(iSerNo) then iSerNo ="NULL"
											    


											    with dcrs
												    .CursorLocation = 3
												    .CursorType = 3
												    .Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
												    Response.Write "<p>"&dcrs.source
												    .ActiveConnection = con
												    .Open
											    end with
											    set dcrs.ActiveConnection = nothing

											    if not dcrs.EOF then
												    if cdbl(dcrs(0))>cdbl(0) then
													    iIssVal = iSerQty * Round((cdbl(dcrs(1)) / cdbl(dcrs(0))),2)
												    end if
											    end if
											    dcrs.Close


											    if IsNull(iIssVal) or trim(iIssVal)="" then iIssVal = 0



										    '	IF cdbl(iIssVal) <> cdbl(0) then
												    ''Response.Write iLotNo &" ** "&iSerNo  &" ** "&	iSerQty
												 											    
												     sSql =	" INSERT INTO INV_T_MaterialIssuedForPick (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
								                            " SerialNo,LocationNumber,BinNumber,QuantityForPick,ItemEntryNo,ItemAttributes,FlagLotOrPack,Returnable,ReturnItem,MaterialType) VALUES "&_
								                            "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
								                            "" & iLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssQty &"," & iEntNo & "," & sAttID & ","& sOnlyLotFlag &",'"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
								                            Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
								                            Con.Execute sSql

		 								    '	End IF

		 									    ''Added By Ragav on Apr 01 ,2010

		 									    if trim(iLotNo)="NULL" then iLotNo="0"
		 									    if trim(iSerNo)="NULL" then iSerNo="0"
										    '	'Response.Write " if Lot <> "" then iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo
										    response.write "<p>Fourth Updation = "
											'    if cdbl(iIssQty)>0 then
											'	    UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
											 '   end if
						        end if 'if PickNode.HaschildNodes() then
					        end if 'if StrComp(PickNode.nodeName,"Pick") = 0 then
				        Next 'For Each PickNode In HeaderNode.childNodes
			        end if 	'if HeaderNode.HaschildNodes() then
			  end if ' if trim(sIssType)="F" then
			   
			    ''Response.End

			    ''''''''''''''''''''''''''''''''''''''''''''''''
			    IF iMRSNo <> "" then
				    with dcrs
					    .CursorLocation = 3
					    .CursorType = 3
					    .Source = "SELECT IssToCode FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRSNo & ""
					    '.Source = "SELECT DEPTNO FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRSNo & ""
					    .ActiveConnection = con
					    .Open
				    end with
				    set dcrs.ActiveConnection = nothing
			    	''Response.Write "<p>"&dcrs.source & vbCrLf & vbCrLf
				    if not dcrs.EOF then
					    sDeptNo = trim(dcrs(0))
				    end if
				    dcrs.Close
				    ''Response.Write "*****"& sTotPickQty &"****" & vbCrLf
			     End IF 'IF iMRSNo <> "" then

			    if iPrQty > 0 then
				    IF iMRSNo <> "" then
					    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYTOPURCHASE = (ISNULL(QUANTITYTOPURCHASE,0) + " & iPrQty & ") " &_
							    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
							    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
					     Response.Write "<p>"&sSql & vbCrLf & vbCrLf
					    con.execute sSql
				    Else
					    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYTOPURCHASE = (ISNULL(QUANTITYTOPURCHASE,0) + " & iPrQty & ") " &_
							    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
							    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
					     Response.Write "<p>"&sSql & vbCrLf & vbCrLf
					    con.execute sSql

				    End IF '	IF iMRSNo <> "" then
				    sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&" and @ENTRYNO = "&iEntNo&"]/PRSchedule"
				    Set ScheduleNode = RootNode.Selectnodes(sExp)

				    if ScheduleNode.Length > 0 then

					    sSchType = trim(ScheduleNode.Item(0).Attributes.Item(0).nodeValue)
					    sSchValue = trim(ScheduleNode.Item(0).Attributes.Item(1).nodeValue)
					    if sSchType = "S" then
						    For Each ScheduleDetNode In ScheduleNode.Item(0).childNodes
							    iSchNo = trim(ScheduleDetNode.Attributes.Item(0).nodeValue)
							    sSchValue = trim(ScheduleDetNode.Attributes.Item(1).nodeValue)
							    iSchQty = trim(ScheduleDetNode.Attributes.Item(2).nodeValue)
							    sSchType = trim(ScheduleDetNode.Attributes.Item(3).nodeValue)

							    if iSchQty = "" then iSchQty = "0"

							    if cdbl(iSchQty) > 0 then
								    with dcrs
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT SCHEDULENO FROM INV_T_MRSITEMPRSTSCHEDULES WHERE SCHEDULENO = " & iSchNo & " AND MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND SCHEDULEDON = " & Pack(sSchValue) & " AND SCHEDULETYPE = " & Pack(sSchType) & " AND STPR = 'PR'"
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs.ActiveConnection = nothing
								    if dcrs.EOF then
									    sSql = "INSERT INTO INV_T_MRSITEMPRSTSCHEDULES (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
										    "ITEMCODE,SCHEDULENO,SCHEDULETYPE,SCHEDULEDON,SCHEDULEDQTY,STPR) VALUES " &_
										    "(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
										    "" & iSchNo & "," & Pack(sSchType) & "," & Pack(sSchValue) & "," & iSchQty & ",'PR')"
									     Response.Write "<p>"&sSql & vbCrLf & vbCrLf
									    con.Execute sSql
								    else
									    sSql = "UPDATE INV_T_MRSITEMPRSTSCHEDULES SET SCHEDULEDQTY = (ISNULL(SCHEDULEDQTY,0) + " & iSchQty & ") WHERE " &_
										    "MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND SCHEDULENO = " & iSchNo & " " &_
										    "AND SCHEDULEDON = " & Pack(sSchValue) & " AND SCHEDULETYPE = " & Pack(sSchType) & " AND STPR = 'PR'"
								    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
									    con.Execute sSql
								    end if
								    dcrs.Close
							    end if
						    next
					    else
						    with dcrs
							    .CursorLocation = 3
							    .CursorType = 3
							    .Source = "SELECT SCHEDULENO FROM INV_T_MRSITEMPRSTSCHEDULES WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND SCHEDULEDON = " & Pack(sSchValue) & " AND STPR = 'PR'"
							    .ActiveConnection = con
							    .Open
						    end with
						    set dcrs.ActiveConnection = nothing
						    if dcrs.EOF then
							    iSchNo = "1"
							    sSql = "INSERT INTO INV_T_MRSITEMPRSTSCHEDULES (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
								    "ITEMCODE,SCHEDULENO,SCHEDULETYPE,SCHEDULEDON,SCHEDULEDQTY,STPR) VALUES " &_
								    "(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
								    "" & iSchNo & "," & Pack(sSchType) & "," & Pack(sSchValue) & "," & iPrQty & ",'PR')"
						    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
							    con.Execute sSql
						    else
							    iSchNo = dcrs(0)
							    sSql = "UPDATE INV_T_MRSITEMPRSTSCHEDULES SET SCHEDULEDQTY = (ISNULL(SCHEDULEDQTY,0) + " & iPrQty & ") WHERE " &_
								    "MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND " &_
								    "SCHEDULENO = " & iSchNo & " AND SCHEDULEDON = " & Pack(sSchValue) & " AND STPR = 'PR'"
						    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
							    con.Execute sSql
						    end if
						    dcrs.Close
					    end if
				    end if

			    end if
			    'Response.Clear 
			    'Response.Write "<p> iTraQty = "& iTraQty 
			    
			   ' sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick"
				'    set ndItemNode = RootNode.selectnodes(sExp)
				 '   if ndItemNode.length>0 then
				  '      For iCnt = 0 to ndItemnode.length - 1
				        with dcrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = "SELECT ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
			                .ActiveConnection = con
			                .Open
		                end with
		                if not dcrs.EOF then
			                iLedgEntNo = dcrs(0)
		                end if
		                dcrs.close

				        
				            sSql = " INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
					            "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,SENTTOACCOUNTS,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
					            "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
					            "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(IssDate) & ",103)," & iIssQty & "," & Round(iValue) & ",'T',"& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
						    	Response.Write "<p>This is out side area "& sSql & vbCrLf & vbCrLf
				            con.execute sSql
				   '     Next
				   ' end if
				    
			    
			    
			    if trim(sIssType)="F" then
			    
		            sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick"
				    Set PickNode = RootNode.Selectnodes(sExp)
				    
	    			sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick/PICK"
				    Set subPickNode = RootNode.Selectnodes(sExp)
				    
				    
				    'response.write "<p> subPickNode.length = "& subPickNode.length

    			    if subPickNode.Length > 0 then
				    'Response.Write "<P>****Ragav***"
					    if not cdbl(PickNode.Item(0).Attributes.getNamedItem("TOT").Value) = 0 then
						    iTotVal =   PickNode.Item(0).Attributes.getNamedItem("TOT").Value
						    For Each PickDetNode In PickNode.Item(0).childNodes
						    
						    'Response.Write "<p> PickDetNode.nodeName="& PickDetNode.nodeName
						        if PickDetNode.nodeName = "PICK" then
							        sPickLoc = trim(PickDetNode.Attributes.getNamedItem("LOC").Value)
							        sPickBin = trim(PickDetNode.Attributes.getNamedItem("BIN").Value)
							        sPickLot = trim(PickDetNode.Attributes.getNamedItem("LOTNO").Value)
							        iIssInvRecNo = trim(PickDetNode.Attributes.getNamedItem("INVRECNO").Value)
							        sPickQty = trim(PickDetNode.Attributes.getNamedItem("QTYISS").Value)

						        '	'Response.Write "sPick Lot = "& sPickLot

							        if ucase(sPickLot) = "N/A" then
								        sPickLot = "NULL"
							        else
								        sPickLot = Pack(sPickLot)
							        end if
							        ''Response.Write ">>>>>>"&sPickBin&"<<<<<"&vbCrLf
							        'If sPickBin = 0
							        if sPickQty	= "" then sPickQty = iTotVal
							        ''Response.Write ">>>>>>"&sPickQty
							        'if cdbl(sPickQty) > 0 then
							        if PickDetNode.hasChildNodes() then
							        ''Response.Write "1"
								        if iIssInvRecNo = "" or IsNull(iIssInvRecNo) then iIssInvRecNo = "NULL"
								        if sPickBin = "0" then sPickBin = "NULL"

								        if IsArray(arrSerial) then
									        erase arrSerial
									        erase arrSerialQty
									        dicSerial.RemoveAll
								        end if
								        
								        

								        For Each SerialHNode In PickDetNode.childNodes
									        For Each SerialDetNode In SerialHNode.childNodes
										        iSerial = trim(SerialDetNode.Attributes.getNamedItem("SERIALNO").Value)
										        iSerialQty = trim(SerialDetNode.Attributes.getNamedItem("QTY").Value)

										        if iSerialQty = "" then iSerialQty = "0"

										        if cdbl(iSerialQty) > 0 then
											        dicSerial.Add iSerial,iSerialQty

										        end if
									        next
								        next 'For Each SerialHNode In PickDetNode.childNodes
								        ''Response.Write dicSerial.Count & "S"
								        if dicSerial.Count > 0 then
									        arrSerial = dicSerial.Keys
									        arrSerialQty = dicSerial.Items

									        'for iArrCtr = 0 to dicSerial.Count - 1
									        '	'Response.Write  arrSerial(iArrCtr) & " >>>>>>> " & arrSerialQty(iArrCtr) & vbCrLf
									        'next

									        ' Function call with serial details
									        
									        'Response.Write "Mark Pick Insert"
									        
									        'Response.Write "<p>Mark Pick Insert First Updation "

									        MarkPickInsert iClass,iItemCode,sOrgID,sPickLoc,sPickBin,sPickLot,sPickQty,iIssInvRecNo,sDeptNo,arrSerial,arrSerialQty,IssDate
								        else

                                         'Response.Write "<p>Mark Pick Insert Second Updation "
									        ' Function call without serial details
									        MarkPickInsert iClass,iItemCode,sOrgID,sPickLoc,sPickBin,sPickLot,sPickQty,iIssInvRecNo,sDeptNo,"NO","A",IssDate
								        end if 'if dicSerial.Count > 0 then

        ''''''''''''''''''''''''''''''''''''''''''''''''Status Updation'''''''''''''''''''''''''''''''''''''''''''''
								        ' Function Call to Update the Line Status of an MR for Inventory Application
								        MRLineStatusUpdate "Issue","Create",iMRSNo,iItemCode,iClass,iEntNo,sOrgID,"4","F","","0"

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

							        'end if 'if cdbl(sPickQty) > 0 then
							        end if'if PickDetNode.hasChildNodes() then
						
							   end if ' if PickDetNode.nodeName = "PICK" then

						    next  'For Each PickDetNode In PickNode.Item(0).childNodes

					    end if 'if not cdbl(PickNode.Item(0).Attributes.getNamedItem("TOT").Value) = 0 then

				    end if 'if subPickNode.Length > 0 then        
			        
			    else 'if trim(sIssType)="M" then
			        
		            sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick"
				    Set PickNode = RootNode.Selectnodes(sExp)
				    
    				sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick/PICK"
				    Set subPickNode = RootNode.Selectnodes(sExp)
				    
				    
				    'response.write "<p> subPickNode.length = "& subPickNode.length

    			    if subPickNode.Length > 0 then
				    'Response.Write "<P>****Ragav***"
					    if not cdbl(PickNode.Item(0).Attributes.getNamedItem("TOT").Value) = 0 then
						    iTotVal =   PickNode.Item(0).Attributes.getNamedItem("TOT").Value
						    For Each PickDetNode In PickNode.Item(0).childNodes
						    
						    'Response.Write "<p> PickDetNode.nodeName="& PickDetNode.nodeName
						        if PickDetNode.nodeName = "PICK" then
							        sPickLoc = trim(PickDetNode.Attributes.getNamedItem("LOC").Value)
							        sPickBin = trim(PickDetNode.Attributes.getNamedItem("BIN").Value)
							        sPickLot = trim(PickDetNode.Attributes.getNamedItem("LOTNO").Value)
							        iIssInvRecNo = trim(PickDetNode.Attributes.getNamedItem("INVRECNO").Value)
							        sPickQty = trim(PickDetNode.Attributes.getNamedItem("QTYISS").Value)

						        
							        if ucase(sPickLot) = "N/A" then
								        sPickLot = "NULL"
							        else
								        sPickLot = Pack(sPickLot)
							        end if
							    
							        if sPickQty	= "" then sPickQty = iTotVal
							    
							        if PickDetNode.hasChildNodes() then
							    
								        if iIssInvRecNo = "" or IsNull(iIssInvRecNo) then iIssInvRecNo = "NULL"
								        if sPickBin = "0" then sPickBin = "NULL"

								        if IsArray(arrSerial) then
									        erase arrSerial
									        erase arrSerialQty
									        dicSerial.RemoveAll
								        end if
								        
								        

								        For Each SerialHNode In PickDetNode.childNodes
									        For Each SerialDetNode In SerialHNode.childNodes
										        iSerial = trim(SerialDetNode.Attributes.getNamedItem("SERIALNO").Value)
										        iSerialQty = trim(SerialDetNode.Attributes.getNamedItem("QTY").Value)

										        if iSerialQty = "" then iSerialQty = "0"

										        if cdbl(iSerialQty) > 0 then
											        dicSerial.Add iSerial,iSerialQty

										        end if
									        next
								        next 'For Each SerialHNode In PickDetNode.childNodes
								        
								       
								        if dicSerial.Count > 0 then
									        arrSerial = dicSerial.Keys
									        arrSerialQty = dicSerial.Items

									        ' Function call with serial details
									        
									        'Response.Write "Mark Pick Insert"
									        
									        'Response.Write "<p>Mark Pick Insert First Updation "

									        MarkPickInsert iClass,iItemCode,sOrgID,sPickLoc,sPickBin,sPickLot,sPickQty,iIssInvRecNo,sDeptNo,arrSerial,arrSerialQty,IssDate
								        else

                                         'Response.Write "<p>Mark Pick Insert Second Updation "
									        ' Function call without serial details
									        MarkPickInsert iClass,iItemCode,sOrgID,sPickLoc,sPickBin,sPickLot,sPickQty,iIssInvRecNo,sDeptNo,"NO","A",IssDate
								        end if 'if dicSerial.Count > 0 then
								        
								        ' Function Call to Update the Line Status of an MR for Inventory Application
								        MRLineStatusUpdate "Issue","Create",iMRSNo,iItemCode,iClass,iEntNo,sOrgID,"4","F","","0"
								        
								    else 'if not PickDetNode.hasChildNodes() then
								        if cdbl(sPickQty)>0 then
								            'Response.Write "<p>Mark Pick Insert Thired Updation "
									            ' Function call without serial details
									         MarkPickInsert iClass,iItemCode,sOrgID,sPickLoc,sPickBin,sPickLot,sPickQty,iIssInvRecNo,sDeptNo,"NO","A",IssDate
    									     
									         ' Function Call to Update the Line Status of an MR for Inventory Application
								            MRLineStatusUpdate "Issue","Create",iMRSNo,iItemCode,iClass,iEntNo,sOrgID,"4","F","","0"
								        end if 'if cdbl(sPickQty)>0 then
								     
							        end if'if PickDetNode.hasChildNodes() then
						
							   end if ' if PickDetNode.nodeName = "PICK" then

						    next  'For Each PickDetNode In PickNode.Item(0).childNodes

					    end if 'if not cdbl(PickNode.Item(0).Attributes.getNamedItem("TOT").Value) = 0 then

				    end if 'if subPickNode.Length > 0 then
			        
			    end if' if trim(sIssType)="F" then
			    
			        
				    
				    sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick/STORE"
				    
				    'Response.Write "<p>sExp ="& sExp
				    Set subPickNode = RootNode.Selectnodes(sExp)

    				'Response.Write "PickNode.Length="&subPickNode.Length& vbCrLf
				    if subPickNode.Length > 0 then
				        if not cdbl(PickNode.Item(0).Attributes.getNamedItem("TOT").Value) = 0 then
						    iTotVal =   PickNode.Item(0).Attributes.getNamedItem("TOT").Value
						    For Each PickDetNode In PickNode.Item(0).childNodes
						    
						        if PickDetNode.nodeName="STORE" then
						    
						    'Response.Write "<p> PickDetNode.nodeName="& PickDetNode.nodeName
						            sPickLoc = trim(PickDetNode.Attributes.getNamedItem("LOC").Value)
							        sPickBin = trim(PickDetNode.Attributes.getNamedItem("BIN").Value)
							        sPickLot = trim(PickDetNode.Attributes.getNamedItem("LOTNO").Value)
							        iIssInvRecNo = trim(PickDetNode.Attributes.getNamedItem("INVRECNO").Value)
							        sPickQty = trim(PickDetNode.Attributes.getNamedItem("QTYISS").Value)

						            if ucase(sPickLot) = "N/A" then
								        sPickLot = "NULL"
							        else
								        sPickLot = Pack(sPickLot)
							        end if
							        'Response.Write ">>>>>>"&sPickLoc &"<<<<<"&vbCrLf
							        'Response.Write ">>>>>>"&sPickBin &"<<<<<"&vbCrLf
							        'If sPickBin = 0
							        if sPickQty	= "" then sPickQty = iTotVal
							        ''Response.Write ">>>>>>"&sPickQty
							        if cdbl(sPickQty) > 0 then
							        ''Response.Write "1"
								        if iIssInvRecNo = "" or IsNull(iIssInvRecNo) then iIssInvRecNo = "NULL"
								        
								        if sPickBin = "0" then sPickBin = "NULL"
								        
								        'Response.Write "<p> One  = "
								        
									        ' Function call without serial details
									        MarkInsert iClass,iItemCode,iEntNo,sPickLoc,sPickBin,sPickQty,sPickQty,iValue,sOrgID,iMRSNo,dMRSDate,sDeptNo
									        
									        
        ''''''''''''''''''''''''''''''''''''''''''''''''Status Updation'''''''''''''''''''''''''''''''''''''''''''''
								        ' Function Call to Update the Line Status of an MR for Inventory Application
								        MRLineStatusUpdate "Issue","Create",iMRSNo,iItemCode,iClass,iEntNo,sOrgID,"4","F","","0"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
							        end if 'if cdbl(sPickQty) > 0 then
							    end if 'if PickDetNode.nodeName="STORE" then
						    next  'For Each PickDetNode In PickNode.Item(0).childNodes

					    end if 'if not cdbl(PickNode.Item(0).Attributes.getNamedItem("TOT").Value) = 0 then

				    end if 'if subPickNode.Length > 0 then
				    

			    if iTraQty > 0 then

	    '''''''''''''''''''''''''''''''''' Stock Transfer Start '''''''''''''''''''''''''''''''
				    if HeaderNode.HaschildNodes() then
					    For Each PickNode In HeaderNode.childNodes
						    if StrComp(PickNode.nodeName,"STDETAILS") = 0 then
							    sSTOrgID = trim(PickNode.Attributes.Item(0).nodeValue)
							    iSTQty = trim(PickNode.Attributes.Item(1).nodeValue)

							    if cdbl(iSTQty) > 0 then
								    sSql = "INSERT INTO INV_T_MRSSTOCKTRANSFER (STOCKTRANSFERNO,MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
									    "ITEMCODE,TRANSFERFROMUNIT,QUANTITYREQUESTEDTR) VALUES " &_
									    "(" & iTransferNo & "," & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
									    "" & Pack(sSTOrgID) & "," & iSTQty & ")"
							    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
								    con.Execute sSql
							    end if
						    end if
					    next
				    ' end if for Header Node has Child Nodes
				    end if
	    '''''''''''''''''''''''''''''''''''End Stock Transfer ''''''''''''''''''''''''''''''''''''''

				    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYFORTRANSFER = (ISNULL(QUANTITYFORTRANSFER,0) + " & iTraQty & ") " &_
						    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
						    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) =  " & iEntNo & " "
				    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
					    con.Execute sSql

				    sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/STSchedule"
				    Set ScheduleNode = RootNode.Selectnodes(sExp)

				    if ScheduleNode.Length > 0 then

					    sSchType = trim(ScheduleNode.Item(0).Attributes.Item(0).nodeValue)
					    sSchValue = trim(ScheduleNode.Item(0).Attributes.Item(1).nodeValue)
					    if sSchType = "S" then
						    For Each ScheduleDetNode In ScheduleNode.Item(0).childNodes
							    iSchNo = trim(ScheduleDetNode.Attributes.Item(0).nodeValue)
							    sSchValue = trim(ScheduleDetNode.Attributes.Item(1).nodeValue)
							    iSchQty = trim(ScheduleDetNode.Attributes.Item(2).nodeValue)
							    sSchType = trim(ScheduleDetNode.Attributes.Item(3).nodeValue)

							    if iSchQty = "" then iSchQty = "0"

							    if cdbl(iSchQty) > 0 then
								    with dcrs
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT SCHEDULENO FROM INV_T_MRSITEMPRSTSCHEDULES WHERE SCHEDULENO = " & iSchNo & " AND MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND SCHEDULEDON = " & Pack(sSchValue) & " AND SCHEDULETYPE = " & Pack(sSchType) & " AND STPR = 'ST'"
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs.ActiveConnection = nothing
								    if dcrs.EOF then
									    sSql = "INSERT INTO INV_T_MRSITEMPRSTSCHEDULES (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
										    "ITEMCODE,SCHEDULENO,SCHEDULETYPE,SCHEDULEDON,SCHEDULEDQTY,STPR) VALUES " &_
										    "(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
										    "" & iSchNo & "," & Pack(sSchType) & "," & Pack(sSchValue) & "," & iSchQty & ",'ST')"
								    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
									    con.Execute sSql
								    else
									    sSql = "UPDATE INV_T_MRSITEMPRSTSCHEDULES SET SCHEDULEDQTY = (ISNULL(SCHEDULEDQTY,0) + " & iSchQty & ") WHERE " &_
										    "MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND SCHEDULENO = " & iSchNo & " " &_
										    "AND SCHEDULEDON = " & Pack(sSchValue) & " AND SCHEDULETYPE = " & Pack(sSchType) & " AND STPR = 'ST'"
								    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
									    con.Execute sSql
								    end if
								    dcrs.Close
							    end if
						    next
					    else
						    with dcrs
							    .CursorLocation = 3
							    .CursorType = 3
							    .Source = "SELECT SCHEDULENO FROM INV_T_MRSITEMPRSTSCHEDULES WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND SCHEDULEDON = " & Pack(sSchValue) & " AND STPR = 'PR'"
							    .ActiveConnection = con
							    .Open
						    end with
						    set dcrs.ActiveConnection = nothing
						    if dcrs.EOF then
							    iSchNo = "1"
							    sSql = "INSERT INTO INV_T_MRSITEMPRSTSCHEDULES (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
								    "ITEMCODE,SCHEDULENO,SCHEDULETYPE,SCHEDULEDON,SCHEDULEDQTY,STPR) VALUES " &_
								    "(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
								    "" & iSchNo & "," & Pack(sSchType) & "," & Pack(sSchValue) & "," & iPrQty & ",'ST')"
				    			Response.Write "<p>"&sSql & vbCrLf & vbCrLf
							    con.Execute sSql
						    else
							    iSchNo = dcrs(0)
							    sSql = "UPDATE INV_T_MRSITEMPRSTSCHEDULES SET SCHEDULEDQTY = (ISNULL(SCHEDULEDQTY,0) + " & iPrQty & ") WHERE " &_
								    "MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND " &_
								    "SCHEDULENO = " & iSchNo & " AND SCHEDULEDON = " & Pack(sSchValue) & " AND STPR = 'ST'"
				    			Response.Write "<p>"&sSql & vbCrLf & vbCrLf
							    con.Execute sSql
						    end if
						    dcrs.Close
					    end if
				    end if
			    end if

			    'if issue type is Firm


			    if sIssType = "F" then
    			'Response.Write "sIssType="&sIssType
    			
    			sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick"
				    Set PickNode = RootNode.Selectnodes(sExp)
					
					If PickNode.length > 0 Then	
						sTotPickQty = cdbl(trim(PickNode.Item(0).Attributes.getNamedItem("TOT").Value))

						if sTotPickQty > 0 then
						    if trim(iMRSNo)<>"" then
							    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & sTotPickQty & ") " &_
								    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
						    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
							    con.execute sSql
						    end if
						end if
					End If
					
		
	    '''''''''''''''''''''''''''''''''' Schedule Entry Start '''''''''''''''''''''''''''''''
				    if HeaderNode.HaschildNodes() then
					    For Each PickNode In HeaderNode.childNodes
						    if StrComp(PickNode.nodeName,"SCHEDULE") = 0 then
							    iSchQty = trim(PickNode.Attributes.Item(0).nodeValue)

							    if cdbl(iSchQty) > 0 then
								    For Each PickDetNode In PickNode.childNodes
									    iSchNo = trim(PickDetNode.Attributes.Item(0).nodeValue)
									    iSchPickQty = trim(PickDetNode.Attributes.Item(2).nodeValue)

									    if cdbl(iSchPickQty) > 0 then
										    sSql = "UPDATE INV_T_MRSITEMSCHEDULES SET ISSUEDQTY = (ISNULL(ISSUEDQTY,0) + " & iSchPickQty & ") " &_
											    " WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
											    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND SCHEDULENO = " & iSchNo & ""
									    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
										    con.Execute sSql
									    end if
								    next
							    end if
						    end if
					    next
				    ' end if for Header Node has Child Nodes
				    end if
	    '''''''''''''''''''''''''''''''''''End Schedule Entry ''''''''''''''''''''''''''''''''''''''

			    ' if issue type is marked
			    elseif sIssType = "M" then
		    ''Response.Write "*********** if issue type is marked ***********"
				    if HeaderNode.HaschildNodes() then
					    For Each PickNode In HeaderNode.childNodes
						    if StrComp(PickNode.nodeName,"Pick") = 0 then

							    sTotPickQty = cdbl(trim(PickNode.Attributes.Item(0).nodeValue))

							    if sTotPickQty > 0 then
								    bFlag = false

    ''''''''''''''''''''''''''''''''''''''''''''''
								    if trim(iMRSNo)<>"" then
									    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & sTotPickQty & ") " &_
										    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
										    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
								    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    end if
							    end if
						    end if

						    ' insert Schedule values if it has been marked
						    if StrComp(PickNode.nodeName,"SCHEDULE") = 0 then
							    iSchQty = trim(PickNode.Attributes.Item(0).nodeValue)

							    if not cdbl(iSchQty) = 0 then
								    For Each PickDetNode In PickNode.childNodes
									    iSchNo = trim(PickDetNode.Attributes.Item(0).nodeValue)
									    sSchOn = trim(PickDetNode.Attributes.Item(1).nodeValue)
									    iSchPickQty = trim(PickDetNode.Attributes.Item(2).nodeValue)

									    if cdbl(iSchPickQty) > 0 then
										    sSql = "INSERT INTO INV_T_ISSUEPICKSCHEDULE (LINENUMBER,PICKNUMBER,SCHEDULENO," &_
											    "SCHEDULEDON,SCHEDULEDQTY) VALUES " &_
											    "(" & iLineNo & "," & iPickNo & "," & iSchNo & "," & Pack(sSchOn) & "," &_
											    "" & iSchPickQty & ")"
									    	''Response.Write "<p>"&sSql & vbCrLf & vbCrLf
										    'con.Execute sSql
									    end if
								    next
							    end if
						    end if

					    next
				    ' end if for Header Node has Child Nodes
				    end if

	    '''''''''''''''''''''''''''''''''' Stock Transfer Start '''''''''''''''''''''''''''''''
				    if HeaderNode.HaschildNodes() then
					    For Each PickNode In HeaderNode.childNodes
						    if StrComp(PickNode.nodeName,"STDETAILS") = 0 then
							    sSTOrgID = trim(PickNode.Attributes.Item(0).nodeValue)
							    iSTQty = trim(PickNode.Attributes.Item(1).nodeValue)

							    if cdbl(iSTQty) > 0 then

								    with dcrs
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT ISNULL(MAX(STOCKTRANSFERNO)+1,1) FROM INV_T_MRSSTOCKTRANSFER WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND TRANSFERFROMUNIT = " & Pack(sSTOrgID) & ""
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs.ActiveConnection = nothing

								    if not dcrs.EOF then
									    with dcrs1
										    .CursorLocation = 3
										    .CursorType = 3
										    .Source = "SELECT STOCKTRANSFERNO FROM INV_T_MRSSTOCKTRANSFER WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND TRANSFERFROMUNIT = " & Pack(sSTOrgID) & " AND STACTION IS NULL"
										    .ActiveConnection = con
										    .Open
									    end with
									    set dcrs1.ActiveConnection = nothing


									    if dcrs1.EOF then
										    sSql = "INSERT INTO INV_T_MRSSTOCKTRANSFER (STOCKTRANSFERNO,MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
											    "ITEMCODE,TRANSFERFROMUNIT,QUANTITYREQUESTEDTR) VALUES " &_
											    "(" & trim(dcrs(0)) & "," & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
											    "" & Pack(sSTOrgID) & "," & iSTQty & ")"

									    	''Response.Write "<p>"&sSql & vbCrLf & vbCrLf
										    'con.Execute sSql
									    else
										    sSql = "UPDATE INV_T_MRSSTOCKTRANSFER SET QUANTITYREQUESTEDTR = (QUANTITYREQUESTEDTR + " & iSTQty & ")" &_
											    " WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
											    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND TRANSFERFROMUNIT = " & Pack(sSTOrgID) & " AND STACTION IS NULL"
									    	''Response.Write "<p>"&sSql & vbCrLf & vbCrLf
										    'con.Execute sSql
									    end if
									    dcrs1.Close
								    end if
								    dcrs.Close

							    end if
						    end if

					    next
				    ' end if for Header Node has Child Nodes
				    end if
	    '''''''''''''''''''''''''''''''''''End Stock Transfer ''''''''''''''''''''''''''''''''''''''

	    '''''''''''''''''''''''''''''''''' Schedule Entry Start '''''''''''''''''''''''''''''''
				    if HeaderNode.HaschildNodes() then
					    For Each PickNode In HeaderNode.childNodes
						    if StrComp(PickNode.nodeName,"SCHEDULE") = 0 then
							    iSchQty = trim(PickNode.Attributes.Item(0).nodeValue)

							    if cdbl(iSchQty) > 0 then
								    For Each PickDetNode In PickNode.childNodes
									    iSchNo = trim(PickDetNode.Attributes.Item(0).nodeValue)
									    iSchPickQty = trim(PickDetNode.Attributes.Item(2).nodeValue)

									    if cdbl(iSchPickQty) > 0 then
										    sSql = "UPDATE INV_T_MRSITEMSCHEDULES SET STATUS = 'MK',MARKEDQTY = (ISNULL(MARKEDQTY,0) + " & iSchPickQty & ") " &_
											    " WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
											    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND SCHEDULENO = " & iSchNo & ""
									    	''Response.Write "<p>"&sSql & vbCrLf & vbCrLf
										    con.Execute sSql
									    end if
								    next
							    end if
						    end if
					    next
				    ' end if for Header Node has Child Nodes
				    end if
	    '''''''''''''''''''''''''''''''''''End Schedule Entry ''''''''''''''''''''''''''''''''''''''

	    ''''''''''''''''''''''''' Not Picked but Issue Type is Marked ''''''''''''''''''''''''''''''
				    if bFlag then
					    ' check for Issue Type whether it is Marked
					    sPickQty = iIssQty
					end if

				    ''''''''''''''''''''''''''''''''''''Stock Status Updation'''''''''''''''''''''''''''''''''''''''''''''''''''''
					    ' Function Call to Update the Header Status of an MR
					    if trim(iMRSNo)<>"" then
						    MRStatusUpdate "Issue","Mark",iMRSNo,iItemCode,iClass,iEntNo,sOrgID
					    end if ' if iMRSNo<>"" then
				    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			    ' end if for Issue type check
			    end if
			    ' end if for Total qty > 0 check
			    end if
		    ' end if for Item Node check
		    end if
	    next

    ''''''''''''''''''''''''''''''''''''Stock Status Updation'''''''''''''''''''''''''''''''''''''''''''''''''''''
	    ' Function Call to Update the Header Status of an MR
	    if trim(iMRSNo)<>"" then
		    MRStatusUpdate "Issue","Create",iMRSNo,iItemCode,iClass,iEntNo,sOrgID
	    end if ' if trim(iMRSNo)<>"" then
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    end if

   ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
	    if trim(IssToCode)="SUB" and 1 = 2 then
		   ' if Trim(sPOConfirm)="N" then

			    sSqlTemp = "select isnull(Max(ForPOEntryNo),0)+1 from ForPurchaseOrder"
								    with rsTemp
									    .ActiveConnection = con
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = sSqlTemp
									    .Open
								    end with

								    set rsTemp.ActiveConnection = nothing

								    if not rsTemp.EOF then
									    iForOrderNo = rsTemp(0)
								    End If
								    rsTemp.Close

			    sSql = "Insert into ForPurchaseOrder(ForPOEntryNo,SendPOTo,ItemCode,ClassCode,QuantityToOrder,OrderedQty,OrganisationCode,MarkedForPOOn,MarkedForPOBy,RefTypeCode,RefNumber) " & _
					    " values ("&iForOrderNo&"," & sPartyCode  & "," & iItemCode & "," & iClass & "," & iIsQty & ",0,'" & sOrgID & "',Convert(datetime,'"&dCreatedOn&"',103),"& iCreatedBy &",12,"& iLedIssueNo &")"
				    	Response.Write "<p>"&sSql
					    con.execute sSql
		   ' end if

	    end if ' if trim(IssToCode)="SUB" then
	    if (trim(sType))="SER" then
		    'if trim(sGatePassConfirm)="Y" then
		    '	if trim(IssToCode)="SER" or trim(IssToCode)="JWK" then
		    
		    
		    sTempSeries = GetInvNumberSeriesCodes("DC",sOrgID,iNumIssueClassCode)
            sArrSeries = Split(sTempSeries,":")
            iSeriesNo = sArrSeries(0)
            iSeriesCode = sArrSeries(1)
        	
	        if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
        	
	            sSql = "Select GroupName from INV_M_Classification where GroupCode = "& iNumIssueClassCode
	            dcrs.Open sSql,con
	            if not dcrs.EOF then
	                sNumClassName = Trim(dcrs(0))
	            end if
	            dcrs.Close 
        	
	            sDCNo = "NULL"
	            Response.Clear 
	            Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Gate Pass - "& sNumClassName &"  Classification</H2></p>"
	            Response.End 
	        end if
	        
	        if not CheckNoSerAvilForThisYear(sOrgID,iSeriesNo,iSeriesCode,IssDate) then
                Response.Clear 
                Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Gate Pass - "& sNumClassName &"  Classification for this Year </H2></p>"
                Response.End 
            end if
            
            sDCNo = GenSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,IssDate)
		    sDCNo = Pack(sDCNo)
		    	    
			if Trim(sDCNO)="" or IsNull(sDCNO) then sDCNo = "NULL"
				    
				    
		    '	end if 'if trim(IssToCode)="SER" then

			    sSql = "UPDATE FORGATEPASSHEADER SET Status = 'Y',DCCode="&sDCNO&" WHERE GatePassNo ="&sGatePassNo
		    	Response.Write "<p>"&sSql
			    CON.EXECUTE sSql

		    'end if 'if trim(sGatePassConfirm)="Y" then
	    end if 'if trim(IssToCode)="SER" or trim(IssToCode)="JWK" then
    'end if ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then

    ''Response.Write Server.MapPath("/Purchase/temp/transaction/PO_PUR_"&Session.SessionID&".xml")
    subContDOM.save(Server.MapPath("../temp/transaction/PO_PUR_"&Session.SessionID&".xml"))
    ''Response.Clear
    ''Response.End
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	    if objfs.FileExists(server.MapPath("../temp/transaction/MRISSUEDETAILS"&Session.SessionID&".xml")) then
		    objfs.DeleteFile server.MapPath("../temp/transaction/MRISSUEDETAILS"&Session.SessionID&".xml")
	    end if
	    if objfs.FileExists(server.MapPath("../temp/transaction/MRSPICKISSUE"&Session.SessionID&".xml")) then
		    objfs.DeleteFile server.MapPath("../temp/transaction/MRSPICKISSUE"&Session.SessionID&".xml")
	    end if
	    if objfs.FileExists(server.MapPath("../temp/transaction/MRSIssue"&Session.SessionID&".xml")) then
		    objfs.DeleteFile server.MapPath("../temp/transaction/MRSIssue"&Session.SessionID&".xml")
	    end if

'	    if objfs.FileExists(server.MapPath("../temp/transaction/mrsIssueData"&Session.SessionID&".xml")) then
'		    objfs.DeleteFile server.MapPath("../temp/transaction/mrsIssueData"&Session.SessionID&".xml")
'	    end if

 End Function 'Function MrsIssueInsert()
    'IF iMRSNo <> "" then
    '	'Response.Write Cdbl(iLedIssueNo)&"."&CDbl(iMRSNo)
    'Else
    '	'Response.Write Cdbl(iLedIssueNo)&".0"
    'End IF



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Function for Inserting the Marked but not picked
    Function MarkInsert(iClass,iItemCode,iEntNo,sLoc,sBin,iReqQty,iIssQty,iValue,sOrgID,iMRSNo,dMRSDate,sDeptNo)
    ''Response.Write "<br>MarkInsert <br>"
    dim dcrs,dcrs1,dcrs2,sSql
    dim sMonYr,sMethod
    dim arrFin,sFinFrom,sFinTo,sTempMonYr,iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue
    dim iWMQty,iWMRecQty,iWMIssQty,iTempWMQty,sUoM,iLineNo
    
    if sBin = "NULL" or IsNull(sBin) or sBin ="" then sBin = "0"

    iReqQty = cdbl(iReqQty)
    iIssQty = cdbl(iIssQty)
    ''Response.Write iReqQty
    'iRecQty = 0
    Set dcrs = Server.CreateObject("ADODB.RecordSet")
    Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
    Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
			    If dMRSDate = "" then dMRSDate = IssDate
			    sTempMonYr = mid(dMRSDate,4,2)
			    sMonYr = sTempMonYr&Year(dMRSDate)

			    'arrFin = split(GetFinancialYear(sMonYr),":")
			    'sFinFrom = arrFin(0)
			    'sFinTo = arrFin(1)
			    arrFin = split(session("Finperiod"),":")
			    sFinFrom = "01/04/"&arrFin(0)
			    sFinTo = "31/03/"&arrFin(1)
			    with dcrs
				    .CursorLocation = 3
				    .CursorType = 3
				    '.Source = "SELECT ISNULL(YEARRECEIPTQUANTITY,0),YEAROPENINGSTOCK,YEARISSUEQUANTITY,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    IF sLoc <> 0 then  'Added on 03 Nov 2007 by Maheshwari
					    .Source = "SELECT ISNULL(YEARRECEIPTQUANTITY,0),ISNULL(YEAROPENINGSTOCK,0),YEARISSUEQUANTITY,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    Else
					    .Source = "SELECT ISNULL(YEARRECEIPTQUANTITY,0),ISNULL(YEAROPENINGSTOCK,0),YEARISSUEQUANTITY,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    End IF
				    .ActiveConnection = con
				    .Open
			    end with
		    	'Response.Write dcrs.source
			    set dcrs.ActiveConnection = nothing

			    if not dcrs.EOF then
				    iRecQty = cdbl(dcrs(0))
				    iYrOpStock = cdbl(dcrs(1))
				    iYrIssQty = cdbl(dcrs(2))
				    iYrCloQty = cdbl(dcrs(3))
				    iYrCloValue = cdbl(dcrs(4))
			    end if
			    dcrs.Close
			    
			    'Response.Write "<p>iYrCloQty  ="& iYrCloQty 
			    
'			    Response.Write "iRecQty = "& iRecQty 
		    
    ''***************************************************************************''''''''
			    ' check for Receipt Quantity so for received



			    if iRecQty = 0 then

				    ' check for Year Opening and Issue Stock, if Issue Stock is there then
				    ' Issue from the Stock available
				    if iYrOpStock > iYrIssQty then
					    ' check for Issue Qty greater than Available Qty
					    if iIssQty > (iYrOpStock - iYrIssQty) then
						    ' to PR Qty
						    iPrQty = iIssQty - (iYrOpStock - iYrIssQty)
						    ' Issue the remaining Qty
						    iIssQty = iYrOpStock - iYrIssQty
					    else
						    iIssQty = iIssQty
					    end if
					    iValue = iIssQty * round((cdbl(iYrCloValue) / cdbl(iYrCloQty)),2)
				    elseif iYrCloQty > 0 then
					    ' check for Issue Qty greater than Closing Qty
					    if iIssQty > iYrCloQty then
						    ' to PR Qty
						    iPrQty = iIssQty - iYrCloQty
						    ' Issue the remaining Qty
						    iIssQty = iYrCloQty
					    else
						    iIssQty = iIssQty
					    end if
					    iValue = iIssQty * round((cdbl(iYrCloValue) / cdbl(iYrCloQty)),2)
				    end if
				    'iValue = iIssQty
				    ''Response.Write "iValue="&iValue & vbCrLf
				    if 1 =1 then
					    ' insert for requested Department other than Inventory
					    if not sDeptNo = "DIS" then
						    if 1 = 2 then
							    with dcrs1
								    .CursorLocation = 3
								    .CursorType = 3
								    .Source = "SELECT ISSUENO FROM INV_T_DEPARTMENTSTOCK WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND ISSUENO = " & iIssueNo & ""
								    .ActiveConnection = con
								    .Open
							    end with
							    set dcrs1.ActiveConnection = nothing

							    if dcrs1.EOF then

								    with dcrs2
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs2.ActiveConnection = nothing

								    if not dcrs2.EOF then
									    sUoM = trim(dcrs2(0))
								    end if
								    dcrs2.Close

								    sSql = "INSERT INTO INV_T_DEPARTMENTSTOCK (ISSUENO,DEPTNO,MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
									    "ITEMCODE,QUANTITYISSUED,QUANTITYUOM,ISSUEENTRYNO) VALUES " &_
									    "(" & iIssueNo & "," & Pack(sDeptNo) & "," & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
									    "" & iIssQty & "," & Pack(sUoM) & "," & iLedIssueNo & ")"
'								    Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql

								    with dcrs2
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT ISNULL(MAX(LINENUMBER)+1,1) FROM INV_T_DEPARTMENTSTOCKISSUEDETAILS"
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs2.ActiveConnection = nothing

								    if not dcrs2.EOF then
									    iLineNo = trim(dcrs2(0))
								    end if
								    dcrs2.Close

								    sSql = "INSERT INTO INV_T_DEPARTMENTSTOCKISSUEDETAILS (LINENUMBER,ISSUENO,ISSUEDATE,QUANTITYISSUED,LOCATIONNUMBER,BINNUMBER,ISSUEVALUE,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE) VALUES " &_
									    "(" & iLineNo & "," & iIssueNo & ",CONVERT(DATETIME," & Pack(IssDate) & ",103)," & iIssQty & "," & sLoc & "," & sBin & "," & Round(iValue) & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ")"
'								    Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
							    else
								    sSql = "UPDATE INV_T_DEPARTMENTSTOCK SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ")" &_
									    " WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
									    "CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & " AND ISSUENO = " & iIssueNo & ""
'								    Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql

								    with dcrs2
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT ISNULL(MAX(LINENUMBER)+1,1) FROM INV_T_DEPARTMENTSTOCKISSUEDETAILS"
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs2.ActiveConnection = nothing

								    if not dcrs2.EOF then
									    iLineNo = trim(dcrs2(0))
								    end if
								    dcrs2.Close

								    sSql = "INSERT INTO INV_T_DEPARTMENTSTOCKISSUEDETAILS (LINENUMBER,ISSUENO,ISSUEDATE,QUANTITYISSUED,LOCATIONNUMBER,BINNUMBER,ISSUEVALUE,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE) VALUES " &_
									    "(" & iLineNo & "," & trim(dcrs1(0)) & ",CONVERT(DATETIME," & Pack(IssDate) & ",103)," & iIssQty & "," & sLoc & "," & sBin & "," & Round(iValue) & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ")"
'								    Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
							    end if
							    dcrs1.Close

						    end if

					    ' end if for check of dept
					    end if
					    with dcrs2
						    .CursorLocation = 3
						    .CursorType = 3
						    .Source = "SELECT ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
						    .ActiveConnection = con
						    .Open
					    end with
					    if not dcrs2.EOF then
						    iLedgEntNo = dcrs2(0)
					    end if
					    dcrs2.close
					    if cdbl(iValue)<>cdbl(0) then
					    sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
						    "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,SENTTOACCOUNTS,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
						    "("& iLedgEntNo  &"," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
						    "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(dMRSDate) & ",103)," & iIssQty & "," & Round(iValue) & ",'T',"& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
				    	Response.Write "<p>1="&sSql & vbCrLf & vbCrLf
					    con.execute sSql
					    end if
	    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

					    'sTempMonYr = mid(dMRSDate,4,2)
					    'sMonYr = sTempMonYr&Year(dMRSDate)

					    with dcrs
						    .CursorLocation = 3
						    .CursorType = 3
						    '.Source = "SELECT ITEMCODE FROM  INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
						    .Source = "SELECT ITEMCODE,IsNull(YearClosingValue,0) FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
						    .ActiveConnection = con
						    .Open
					    end with
					    set dcrs.ActiveConnection = nothing

					    if dcrs.EOF then
						    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
							    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
							    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iIssQty & "," &_
							    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
					    	Response.Write sSql & vbCrLf & vbCrLf
						    con.execute sSql
					    else
					        if cdbl(dcrs(1))>=Round(iValue) then
						        sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
							        "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIssQty & "),"&_
							        "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
							        "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
							        "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
							        "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
							        "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
						        Response.Write " <p>FIRST = "& sSql & vbCrLf & vbCrLf
						    else
						        sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
							        "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIssQty & "),"&_
							        "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
							        "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
							        "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
							        "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
							        "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
						        Response.Write " <p>FIRST = "& sSql & vbCrLf & vbCrLf
						    end if
						    con.execute sSql
					    end if
					    dcrs.Close

	    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

					    with dcrs
						    .CursorLocation = 3
						    .CursorType = 3
						    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
						    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					    	Response.Write "<p>"&dcrs.source
						    .ActiveConnection = con
						    .Open
					    end with

					    set dcrs.ActiveConnection = nothing


					    if not dcrs.EOF then
					    
					        if cdbl(dcrs(0)) >= round(iValue) then
					         sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						            "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
						            "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
						            "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
						            "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						            "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						            "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    	        Response.Write "<p>"&sSql & vbCrLf & vbCrLf
					            
						    else
							    sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						            "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
						            "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
						            "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
						            "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						            "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						            "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    	        Response.Write "<p>"&sSql & vbCrLf & vbCrLf
						    end if
						    con.execute sSql
					    end if
					    dcrs.Close
					    
					    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                        sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                        sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                        Response.Write "<p>"& sSql 
	                    con.execute sSql 

					   


				    end if

				    if iMRSNo <>"" then
					    with dcrs
						    .CursorLocation = 3
						    .CursorType = 3
						    .Source = "SELECT ISNULL(LOTNUMBER,0),ISNULL(SERIALNUMBER,0) FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iMRSNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND ISNULL(ITEMENTRYNO,0) = " & iEntNo & " "
					    '	'Response.Write dcrs.source
						    .ActiveConnection = con
						    .open
					    end with
					    set dcrs.ActiveConnection = nothing

					    if not dcrs.EOF then
						    iLotNo = dcrs(0)
						    iSerNo  = dcrs(1)
					    end if
					    dcrs.close
				    end if 'if iMRSNo <>"" then
					    ''Response.Write "************************ MARK INSERT ************************************"
					    IF trim(iLotNo) = "NULL" then iLotNo = 0
					    if trim(iSerNo)="NULL" then iSerNo="0"
    '					'Response.Clear
				    '	'Response.Write " if RecQty=0 iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo


					    if trim(iLotNo) = "0" and trim(iSerNo) = "0" then
					        'Response.Write "<p>Welcome to Location Lot Updation"
					        if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
					            Response.write "<p>4th UpdateLocation Lot"
						        UpdateLocLot iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
						    end if 'if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
						    'sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iIssQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iIssQty & ") WHERE INVENTORYRECEIPTNO = " & iMRSNo & " AND"&_
						    '		"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
						    '		"ISNULL(ITEMENTRYNO,0) = " & iEntNo & " AND STORAGELOCATIONNO = " & sLoc & " AND (ISNULL(STORAGEBINNUMBER,NULL) = " & sBin & " OR ISNULL(STORAGEBINNUMBER,0) = " & sBin & ") "
						    ''Response.Write "lot= Series"& sSql & vbcrlf & vbcrlf
						    'Con.Execute sSql
					    end if

				    '''***************************************************************************''''''''
			    else

		    '		'Response.Write "iRecQty ="& iRecQty

				  '  with dcrs
				  ' 	.CursorLocation = 3
				  '  	.CursorType = 3
				  '  	.Source = "SELECT AccountingType FROM INV_M_ITEMMaster WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
				  '  	.ActiveConnection = con
				  '  	.Open
				  '  end with
				  '  set dcrs.ActiveConnection = nothing

				  '  if not dcrs.EOF then
				  '  	sMethod = dcrs(0)
				  '  end if
				  '  dcrs.Close
			   ' Response.Write "sMethod="& sMethod

			    sMethod = "W" ' Weighted Average
		    '	'Response.Write " iIsQty = "& iIsQty  & "   "

				    if sMethod = "F" then
					    ' check for Issue Qty greater than Available Qty
					    if (iYrOpStock - iYrIssQty) > 0  then
						    ' to check for Year to date Stock tables
						    iWMQty = iIssQty - (iYrOpStock - iYrIssQty)

						    ' Issue the remaining Qty
						    iIssQty = iYrOpStock - iYrIssQty

						    iValue = iIssQty * round((cdbl(iYrCloValue) / cdbl(iYrCloQty)),2)
					    with dcrs2
						    .CursorLocation = 3
						    .CursorType = 3
						    .Source = "SELECT ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
						    .ActiveConnection = con
						    .Open
					    end with
					    if not dcrs2.EOF then
						    iLedgEntNo = dcrs2(0)
					    end if
					    dcrs2.close
					    if cdbl(iValue)<>cdbl(0) then
						    sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
							    "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
							    "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
							    "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(dMRSDate) & ",103)," & iIssQty & "," & Round(iValue) & ","& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
					    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
						    con.execute sSql
					    end if


		    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

						    sTempMonYr = mid(dMRSDate,4,2)
						    sMonYr = sTempMonYr&Year(dMRSDate)

						    with dcrs
							    .CursorLocation = 3
							    .CursorType = 3
							    'Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
							    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
							    .ActiveConnection = con
							    .Open
						    end with
						    set dcrs.ActiveConnection = nothing

						    if dcrs.EOF then
							    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
								    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
								    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iIssQty & "," &_
								    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
						    	Response.Write sSql & vbCrLf & vbCrLf
							    con.execute sSql
						    else
							    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
							    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
								    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
								    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
								    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
								    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"

							    Response.Write "<p>Second = "& sSql & vbCrLf & vbCrLf
							    con.execute sSql
						    end if
						    dcrs.Close




		    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						    sTempMonYr = mid(dMRSDate,4,2)
						    sMonYr = sTempMonYr&Year(dMRSDate)

						    'arrFin = split(GetFinancialYear(sMonYr),":")
						    'sFinFrom = arrFin(0)
						    'sFinTo = arrFin(1)
						    arrFin = split(session("Finperiod"),":")
						    sFinFrom = "01/04/"&arrFin(0)
						    sFinTo = "31/03/"&arrFin(1)
						    with dcrs
							    .CursorLocation = 3
							    .CursorType = 3
							    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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
							    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
							    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
							    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
							    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
							    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
							    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					    	Response.Write sSql & vbCrLf & vbCrLf
						    con.execute sSql
						    
						    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                            sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                            sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                            Response.Write "<p>"& sSql 
	                        con.execute sSql 

					    end if

					    ' check for stock in Stock Table
					    if (iWMQty > 0) or ((iYrOpStock - iYrIssQty) <= 0) then

						    if ((iYrOpStock - iYrIssQty) <= 0) then
							    iWMQty = iWMQty + iIssQty
						    end if

						    with dcrs1
							    .CursorLocation = 3
							    .CursorType = 3
							    .Source = "SELECT ISNULL(QUANTITYRECEIVED,0),RECEIPTVALUE,QUANTITYISSUED,INVENTORYRECEIPTNO FROM INV_T_YTDSTOCKFIFO WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY INVENTORYRECEIPTNO"
							    .ActiveConnection = con
							    .Open
						    end with
						    set dcrs1.ActiveConnection = nothing

						    if not dcrs1.EOF then
							    do while not dcrs1.EOF
								    iWMRecQty = cdbl(trim(dcrs1(0)))
								    iWMIssQty = cdbl(trim(dcrs1(2)))

								    iTempWMQty = iWMQty

								    ' Check for Issue Quanity
								    if iWMQty  <= 0 then exit do

								    ' check for Receipt and Issue Quantity for that Location and Bin
								    if iWMRecQty > iWMIssQty then
									    if iWMQty > (iWMRecQty - iWMIssQty) then
										    ' Issue the remaining Qty
										    iWMQty = iWMRecQty - iWMIssQty
									    else
										    iWMQty = iWMQty
									    end if
									    iValue = iWMQty * (cdbl(trim(dcrs1(1))) / iWMRecQty)


					    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

									    sTempMonYr = mid(dMRSDate,4,2)
									    sMonYr = sTempMonYr&Year(dMRSDate)

									    with dcrs
										    .CursorLocation = 3
										    .CursorType = 3
										    '.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
										    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
										    .ActiveConnection = con
										    .Open
									    end with
									    set dcrs.ActiveConnection = nothing

									    if dcrs.EOF then
										    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
											    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
											    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iWMQty & "," &_
											    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
									    	Response.Write sSql & vbCrLf & vbCrLf
										    con.execute sSql
									    else
										    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iWMQty & ")," &_
										    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
											    "YEARRESERVED = (YEARRESERVED + " & iWMQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
											    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
											    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
											    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
											    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
									    	Response.Write "<p>Thired = "& sSql & vbCrLf & vbCrLf
										    con.execute sSql
									    end if
									    dcrs.Close


					    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

									    sTempMonYr = mid(dMRSDate,4,2)
									    sMonYr = sTempMonYr&Year(dMRSDate)

									    'arrFin = split(GetFinancialYear(sMonYr),":")
									    'sFinFrom = arrFin(0)
									    'sFinTo = arrFin(1)
									    arrFin = split(session("Finperiod"),":")
									    sFinFrom = "01/04/"&arrFin(0)
									    sFinTo = "31/03/"&arrFin(1)

									    sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iWMQty & ")," &_
										    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iWMQty & "), " &_
										    "YEARRESERVED = (YEARRESERVED + " & iWMQty & ") WHERE " &_
										    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
										    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
										    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
								        Response.Write sSql & vbCrLf & vbCrLf
									    con.execute sSql
									    
									    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                                        sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                                        sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                                        Response.Write "<p>"& sSql 
	                                    con.execute sSql 


									    ''Response.Write "Issue Ori " & iTempWMQty & vbCrLf & vbCrLf
									    '''Response.Write "Issued " & iWMQty & vbCrLf & vbCrLf
									    iWMQty = iTempWMQty - iWMQty
									    '''Response.Write iWMQty & vbCrLf & vbCrLf
								    end if
							    dcrs1.MoveNext
							    loop
						    end if
						    dcrs1.Close
					    end if

				    ' check for LIFO method
				    elseif sMethod = "L" then
					    iWMQty = iIssQty
					    with dcrs1
						    .CursorLocation = 3
						    .CursorType = 3
						    '.Source = "SELECT ISNULL(QUANTITYRECEIVED,0),RECEIPTVALUE,QUANTITYISSUED,INVENTORYRECEIPTNO FROM INV_T_YTDSTOCKLIFO WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY INVENTORYRECEIPTNO DESC"
						    .Source = "Select LotQuantityNett,Rate,QuantityIssued,InventoryReceiptNo from INV_T_LocationLot WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND StorageLocationNo = " & sLoc & " AND (StorageBinNumber = " & sBin & " OR StorageBinNumber IS NULL) ORDER BY INVENTORYRECEIPTNO DESC"
						    Response.Write dcrs1.source
						    .ActiveConnection = con
						    .Open
					    end with
					    set dcrs1.ActiveConnection = nothing

					    if not dcrs1.EOF then
						    do while not dcrs1.EOF
							    iWMRecQty = cdbl(trim(dcrs1(0)))
							    iWMIssQty = cdbl(trim(dcrs1(2)))

							    iTempWMQty = iWMQty

							    ' Check for Issue Quanity
							    if iWMQty  <= 0 then exit do

							    ' check for Receipt and Issue Quantity for that Location and Bin
							    if iWMRecQty > iWMIssQty then
								    if iWMQty > (iWMRecQty - iWMIssQty) then
									    ' Issue the remaining Qty
									    iWMQty = iWMRecQty - iWMIssQty
								    else
									    iWMQty = iWMQty
								    end if
								    'iValue = iWMQty * (cdbl(trim(dcrs1(1))) / iWMRecQty)
								    iValue = iWMQty * Trim(dcrs1(1))
								    
								    with dcrs2
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
									    .ActiveConnection = con
									    .Open
								    end with
								    if not dcrs2.EOF then
									    iLedgEntNo = dcrs2(0)
								    end if
								    dcrs2.close

								    sSql = "UPDATE Inv_T_LocationLot SET QUANTITYISSUED = (QUANTITYISSUED + " & iWMQty & ")" &_
									    " WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
									    " ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
									    " StorageLocationNo = " & sLoc & " AND (StorageBinNumber= " & sBin & " OR StorageBinNumber IS NULL) AND INVENTORYRECEIPTNO = " & trim(dcrs1(3)) & ""
								    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
								    con.execute sSql

								    if cdbl(iValue)<>cdbl(0) then
									    sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
										    "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
										    "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
										    "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(dMRSDate) & ",103)," & iWMQty & "," & Round(iValue) & ","& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
								    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    end if



				    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

								    sTempMonYr = mid(dMRSDate,4,2)
								    sMonYr = sTempMonYr&Year(dMRSDate)

								    with dcrs
									    .CursorLocation = 3
									    .CursorType = 3
									    '.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
									    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs.ActiveConnection = nothing

								    if dcrs.EOF then
									    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
										    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
										    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iWMQty & "," &_
										    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
								    	Response.Write sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    else
									    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iWMQty & ")," &_
									    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
										    "YEARRESERVED = (YEARRESERVED + " & iWMQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
										    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
										    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
										    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
								    	Response.Write "<p>Fourth = "& sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    end if
								    dcrs.Close

				    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

								    sTempMonYr = mid(dMRSDate,4,2)
								    sMonYr = sTempMonYr&Year(dMRSDate)

								    'arrFin = split(GetFinancialYear(sMonYr),":")
								    'sFinFrom = arrFin(0)
								    'sFinTo = arrFin(1)
								    arrFin = split(session("Finperiod"),":")
								    sFinFrom = "01/04/"&arrFin(0)
								    sFinTo = "31/03/"&arrFin(1)
								    with dcrs
									    .CursorLocation = 3
									    .CursorType = 3
									    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
									    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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

								    sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iWMQty & ")," &_
									    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iWMQty & "), " &_
									    "YEARRESERVED = (YEARRESERVED + " & iWMQty & ") WHERE " &_
									    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
									    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
									    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
									    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							    	Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
								    
								    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                                    sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                                    sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                                    Response.Write "<p>"& sSql 
	                                con.execute sSql 

								    ''Response.Write "Issue Ori " & iTempWMQty & vbCrLf & vbCrLf
								    '''Response.Write "Issued " & iWMQty & vbCrLf & vbCrLf
								    iWMQty = iTempWMQty - iWMQty
								    '''Response.Write iWMQty & vbCrLf & vbCrLf
							    end if
						    dcrs1.MoveNext
						    loop

						    '''Response.Write " remain " & iWMQty & vbCrLf & vbCrLf

						    if iWMQty > 0 then
							    'should go for PR
							    '''Response.Write "reduce from Stock " & iWMQty & vbCrLf & vbCrLf

							    ' check for Issue Qty greater than Available Qty
							    if (iYrOpStock - iYrIssQty) > 0  then
								    ' to check for Year to date Stock tables
								    iWMQty = iWMQty - (iYrOpStock - iYrIssQty)

								    ' Issue the remaining Qty
								    iIssQty = iYrOpStock - iYrIssQty

								    iValue = iIssQty * (iYrCloValue / iYrCloQty)
								    with dcrs2
									    .CursorLocation = 3
									    .CursorType = 3
									    .Source = "SELECT ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
									    .ActiveConnection = con
									    .Open
								    end with
								    if not dcrs2.EOF then
									    iLedgEntNo = dcrs2(0)
								    end if
								    dcrs2.close

								    if cdbl(iValue)<>cdbl(0) then

								    sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
									    "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
									    "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
									    "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(dMRSDate) & ",103)," & iIssQty & "," & Round(iValue) & ","& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
								    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
								    con.execute sSql
								    end if



				    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

								    sTempMonYr = mid(dMRSDate,4,2)
								    sMonYr = sTempMonYr&Year(dMRSDate)

								    with dcrs
									    .CursorLocation = 3
									    .CursorType = 3
									    '.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
									    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
									    .ActiveConnection = con
									    .Open
								    end with
								    set dcrs.ActiveConnection = nothing

								    if dcrs.EOF then
									    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
										    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
										    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iIssQty & "," &_
										    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
								    	Response.Write sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    else
									    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
									    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
										    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
										    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
										    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
										    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
									    Response.Write "<p>Fifth = "& sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    end if
								    dcrs.Close
				    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
								    sTempMonYr = mid(dMRSDate,4,2)
								    sMonYr = sTempMonYr&Year(dMRSDate)

								    'arrFin = split(GetFinancialYear(sMonYr),":")
								    'sFinFrom = arrFin(0)
								    'sFinTo = arrFin(1)
								    arrFin = split(session("Finperiod"),":")
								    sFinFrom = "01/04/"&arrFin(0)
								    sFinTo = "31/03/"&arrFin(1)
								    with dcrs
									    .CursorLocation = 3
									    .CursorType = 3
									    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
									    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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
									    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
									    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
									    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
									    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
									    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
									    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
								    Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
								    
								    
								    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                                    sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                                    sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                                    Response.Write "<p>"& sSql 
	                                con.execute sSql 

							    end if
						    end if
					    ' no Quantity exists check in stock table
					    else

						    ' check for Issue Qty greater than Available Qty
						    if (iYrOpStock - iYrIssQty) > 0  then
							    ' to check for Year to date Stock tables
							    iWMQty = iIssQty - (iYrOpStock - iYrIssQty)

							    ' Issue the remaining Qty
							    iIssQty = iYrOpStock - iYrIssQty

							    iValue = iIssQty * (iYrCloValue / iYrCloQty)


			    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

							    sTempMonYr = mid(dMRSDate,4,2)
							    sMonYr = sTempMonYr&Year(dMRSDate)

							    with dcrs
								    .CursorLocation = 3
								    .CursorType = 3
								    '.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
								    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
								    .ActiveConnection = con
								    .Open
							    end with
							    set dcrs.ActiveConnection = nothing

							    if dcrs.EOF then
								    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
									    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
									    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & Pack(sMonYr) & "," & iIssQty & "," &_
									    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
							    	Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
							    else
								    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
								    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
									    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
									    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
									    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
									    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
									    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
								    Response.Write "<p>Sixth = "&sSql & vbCrLf & vbCrLf
								    con.execute sSql
							    end if
							    dcrs.Close



			    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
							    sTempMonYr = mid(dMRSDate,4,2)
							    sMonYr = sTempMonYr&Year(dMRSDate)

							    'arrFin = split(GetFinancialYear(sMonYr),":")
							    'sFinFrom = arrFin(0)
							    'sFinTo = arrFin(1)
							    arrFin = split(session("Finperiod"),":")
							    sFinFrom = "01/04/"&arrFin(0)
							    sFinTo = "31/03/"&arrFin(1)

							    with dcrs
								    .CursorLocation = 3
								    .CursorType = 3
								    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
								    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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
								    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
								    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
								    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
								    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
								    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							    Response.Write sSql & vbCrLf & vbCrLf
							    con.execute sSql
							    
							    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                                sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                                sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                                Response.Write "<p>"& sSql 
	                            con.execute sSql 

						    end if

					    end if
					    dcrs1.Close

				    ' check for WA method
				    elseif sMethod = "W" then
				    
'				        Response.Write "<p>Method = w"
'                        Response.Write "iYrCloQty = "& iYrCloQty
					    ' check for Closing Qty greater than 0
					    if iYrCloQty > 0  then
					        
						    if iYrCloQty > iIssQty then
							    iIssQty = iIssQty
						    else
							    ' to PR
							    iWMQty = iIssQty - iYrCloQty

							    ' Issue the remaining Qty
							    iIssQty = iYrCloQty
						    end if

						    ''Response.Write "iYrCloValue = "&iYrCloValue
						    ''Response.Write "iYrCloQty = "& iYrCloQty

						    iValue = iIssQty * round((cdbl(iYrCloValue) / cdbl(iYrCloQty)),2)
						    
					'	    with dcrs2
					'		    .CursorLocation = 3
					'		    .CursorType = 3
					'		    .Source = "SELECT ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
					'		    .ActiveConnection = con
					'		    .Open
					'	    end with
					'	    if not dcrs2.EOF then
					'		    iLedgEntNo = dcrs2(0)
					'	    end if
					'	    dcrs2.close
						    
						    
						    
						    ''added the condition by ragav on Sep 12,2012 for Pick Later case stock updation stoped other case it will happen
				            ''begin 
				            
				          '  if cdbl(iValue)<>cdbl(0) then
					       '     sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
						  '          "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
						  '          "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
						  '          "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(dMRSDate) & ",103)," & iIssQty & "," & Round(iValue) & ","& iNoofCases &","&sAttID&","& sPartyCode &",'"& IssToCode &"')"
				    	  '      Response.Write sSql & vbCrLf & vbCrLf
					      '      con.execute sSql
					      '  end if
				            
						   ' if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
						   '     if cdbl(iValue)<>cdbl(0) then
'
'						            sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
'							            "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
'							            "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
'							            "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(dMRSDate) & ",103)," & iIssQty & "," & Round(iValue) & ","& iNoofCases &","&sAttID&","& sPartyCode &",'"& IssToCode &"')"
'					    	           	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
'						            con.execute sSql
'						        end if
						   ' end if' if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
						    'end



		    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

						    sTempMonYr = mid(dMRSDate,4,2)
						    sMonYr = sTempMonYr&Year(dMRSDate)
					    '	'Response.Write "*****iIsQty = "& iIsQty & " ****"


						    with dcrs
							    .CursorLocation = 3
							    .CursorType = 3
							    '.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
							    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
							   ' Response.Write "<p>"& dcrs.source
							    .ActiveConnection = con
							    .Open
						    end with
						    set dcrs.ActiveConnection = nothing

						    if dcrs.EOF then
							    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
								    "LOCATIONNUMBER,BINNUMBER,MONTHANDYEAR,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
								    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iIssQty & "," &_
								    "" & sLoc & "," & sBin & "," & iIssQty & "," & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
						    	Response.Write sSql & vbCrLf & vbCrLf
							    con.execute sSql
						    else
							    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
							    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIssQty & "), "&_
								    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
								    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
								    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
								    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
							    Response.Write "<p>Seventh = "& sSql & vbCrLf & vbCrLf
							    con.execute sSql
						    end if
						    dcrs.Close





		    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						    sTempMonYr = mid(dMRSDate,4,2)
						    sMonYr = sTempMonYr&Year(dMRSDate)

						    'arrFin = split(GetFinancialYear(sMonYr),":")
						    'sFinFrom = arrFin(0)
						    'sFinTo = arrFin(1)
						    arrFin = split(session("Finperiod"),":")
						    sFinFrom = "01/04/"&arrFin(0)
						    sFinTo = "31/03/"&arrFin(1)
						    with dcrs
							    .CursorLocation = 3
							    .CursorType = 3
							    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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
							    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
							    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
							    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
							    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
							    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
							    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
						    Response.Write "<p>"&sSql & vbCrLf & vbCrLf
						    con.execute sSql
						    
						    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                            sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                            sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                            Response.Write "<p>"& sSql 
	                        con.execute sSql 

					    end if

				    ' end if for VALUATION METHOD check
				    end if

				    ''Added By Ragav on Apr 01 ,2010
				    IF trim(iLotNo) = "NULL" or Trim(iLotNo)="N/A" then iLotNo = 0
				    if trim(iSerNo)="NULL" then iSerNo="0"
    '					'Response.Clear
                Response.Write "<p>SerNo = "& iSerNo 
					    Response.Write "<p> if RecQty<>0 iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo
					    
				    if trim(iLotNo) = "0" and trim(iSerNo) = "0" then
'				        Response.Write "<P>Welcome to Update Location Lot <p>Fifth Updation = "
				        
				        ''added the condition by ragav on Sep 12,2012 for Pick Later case stock updation stoped other case it will happen
				        ''begin 
				        'UpdateLocLot iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
				        if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
				        Response.write "<p>5th UpdateLocation Lot"
					        UpdateLocLot iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
					    end if' if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
					    'end '
				    end if

    '''***************************************************************************''''''''
			    ' end if for receipt qty check
			    end if
    end function

    ' Function to Insert Marked and Picked Values
    Function MarkPickInsert(iClass,iItemCode,sOrgID,sLoc,sBin,sPickLot,iIssQtyP,iIssInvRecNo,sDeptNo,arrSerialP,arrSerialQtyP,iTransDate)
    ''Response.Write "MarkPickInsert" &"<br>"
    ' Declaration of variables
    Dim dcrs,dcrs1,dcrs2,dcrs3
    dim iSerialNo,iSerialQty,iLineNo,sUoM,iItmRate,sMonYr
    dim arrFin,sFinFrom,sFinTo,sTempMonYr,iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue,iPrQty
    dim iWMQty,iWMRecQty,iWMIssQty,iTempWMQty,sIssType,iTempTotal,iTemp,iIssQty,iCtr
    dim sWCCode, sMCCode,iMCQty, iQtyIssued, iEntryNo, sWCCIDE, sExp4, sExp5, WCNode, MCNode
    dim iMCounter, iWCounter,iTempVal,iClosingVal,iClosingStk,sTempSerialNo,iCnt,iTempIssQty
    iCtr = 0
    iIssQty = cdbl(iIssQtyP)

    iTemp = iIssQty
    iDeptEntryNo = iDeptEntryNo + 1
    'Declaration of Objects
    arrFin = split(session("Finperiod"),":")
    sFinFrom = "01/04/"&arrFin(0)
    sFinTo = "31/03/"&arrFin(1)

    Set dcrs = Server.CreateObject("ADODB.RecordSet")
    Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
    Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
    Set dcrs3 = Server.CreateObject("ADODB.RecordSet")

    ''Response.Write "iTransDate="&iTransDate &vbCrLf
    If sBin = "NULL" then sBin = 0

		    'To Add WorkCenter Details
		    sExp4 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItemCode&"]/AddDet/WorkCenter"
		    Set WCNode = RootNode.Selectnodes(sExp4)
		    For iWCounter = 0 to WCNode.Length - 1
			    sWCCode = trim(WCNode.Item(iWCounter).Attributes.getNamedItem("WCODE").Value)

			    sExp5 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItemCode&"]/AddDet/WorkCenter [ @WCODE = '"&sWCCode&"']/MachineCenter"
			    Set MCNode = RootNode.Selectnodes(sExp5)
			    if MCNode.length > 0 then
				    For iMCounter = 0 to MCNode.Length - 1
					    sMCCode = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("MCODE").Value)
					    iMCQty = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("QTY").Value)

					    if sMCCode = "select" then
						    sMCCode = "NULL"
					    else
						    sMCCode = Pack(sMCCode)
					    end if

					    sSql = "INSERT INTO inv_t_issuedeptwiseBreakUp (IssueEntryNo,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
						    "WORKCENTERCODE,MACHINECENTERCODE,QUANTITYISSUED,ENTRYNO) VALUES " &_
						    "(" & iLedIssueNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
						    "" & Pack(sWCCode) & "," & sMCCode & "," & iMCQty & "," & iDeptEntryNo & " )"
			    '		'Response.Write sSql & vbCrLf & vbCrLf
					    con.Execute sSql
					    iDeptEntryNo = iDeptEntryNo + 1
				    next
			    else

				    sSql = "INSERT INTO inv_t_issuedeptwiseBreakUp (IssueEntryNo,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
					    "WORKCENTERCODE,MACHINECENTERCODE,QUANTITYISSUED,ENTRYNO) VALUES " &_
					    "(" & iLedIssueNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
					    "" & Pack(sWCCode) & "," & sMCCode & "," & iMCQty & "," & iDeptEntryNo & " )"
			    '	'Response.Write sSql & vbCrLf & vbCrLf
				    con.Execute sSql
				    iDeptEntryNo = iDeptEntryNo + 1
			    end if
		    next



	    if IsArray(arrSerialP)  then

		  

				    with dcrs
					    .CursorLocation = 3
					    .CursorType = 3
					    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0),ISNULL(YEARCLOSINGSTOCK,0) FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					    .ActiveConnection = con
					    .Open
				    end with
		    			Response.Write dcrs.source & vbCrLf
				    set dcrs.ActiveConnection = nothing

				    if not dcrs.EOF then

		    '			'Response.Write "Closing Value = "& dcrs(0)
		    '			'Response.Write "Closing Qty = "& dcrs(1)

					    if cdbl(dcrs(0))>0  and cdbl(dcrs(1))>0 then
						    iValue = iIssQty * round((cdbl(dcrs(0))/cdbl(dcrs(1))),2)
					    end if

				    end if
				    dcrs.Close



                ''added the condition by ragav on Sep 12,2012 for Pick Later case stock updation stoped other case it will happen
				''begin 
			     '   if cdbl(iValue)<>cdbl(0) then
'
'				        sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
'					        "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,SENTTOACCOUNTS,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
'					        "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
'					        "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(iTransDate) & ",103)," & iIssQty & "," & Round(iValue) & ",'T',"& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
'				        Response.Write " Sql =  "& sSql & vbCrLf
'				        con.execute sSql
'			        end if
			       ' if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
			    '    
			     '       if cdbl(iValue)<>cdbl(0) then
'
'				            sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
'					            "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,SENTTOACCOUNTS,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
'					            "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
'					            "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(iTransDate) & ",103)," & iIssQty & "," & Round(iValue) & ",'T',"& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
'						    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
'				            con.execute sSql
'			            end if
			       ' end if'if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
			        ''end 

	    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				    If dMRSDate = "" then dMRSDate = IssDate
				    sTempMonYr = mid(dMRSDate,4,2)
				    sMonYr = sTempMonYr&Year(dMRSDate)
				    'arrFin = split(GetFinancialYear(sMonYr),":")
				    'sFinFrom = arrFin(0)
				    'sFinTo = arrFin(1)
				    arrFin = split(session("Finperiod"),":")
				    sFinFrom = "01/04/"&arrFin(0)
				    sFinTo = "31/03/"&arrFin(1)
    ''Response.Clear
				    with dcrs
					    .CursorLocation = 3
					    .CursorType = 3
					    .Source = "SELECT ITEMCODE,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) "
					    .ActiveConnection = con
					    .Open
				    end with
				    set dcrs.ActiveConnection = nothing

				    if dcrs.EOF then

					    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
						    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
						    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iIssQty & "," &_
						    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
				    	Response.Write sSql & vbCrLf
					    con.execute sSql
				    else

					    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
						    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
						    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
						    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103) "

					    Response.Write "<p>Eigth="& sSql & vbCrLf
					    con.execute sSql
					    
					    sSql = " Update INV_T_ITEMLOCATIONSTOCK set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                        sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                        sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103) and LocationNumber = "& sLoc &" and (BINNUMBER = "& sBin &" OR BINNUMBER IS NULL)"
                        Response.Write "<p><p>Eigth="& sSql  & vbCrLf 
	                    con.execute sSql 
					    
				    end if
				    dcrs.Close
    ''Response.End

	    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				    sTempMonYr = mid(dMRSDate,4,2)
				    sMonYr = sTempMonYr&Year(dMRSDate)


		    '		'Response.Write vbCrLf & "ClosVal="& iIssQty


				    sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty  & ")," &_
					    "YEARCLOSINGSTOCK = (YEAROPENINGSTOCK + YearReceiptQuantity - YEARISSUEQUANTITY), " &_
					    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
					    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
					    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
					    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    Response.Write "<p>Chk="&sSql & vbCrLf
				    con.execute sSql
				    
				    
				    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                    sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                    sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                    Response.Write "<p>"& sSql 
	                con.execute sSql 
	                
				    If iMRSNo <> "" then
					    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
							    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
							    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & "  AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
					    con.execute sSql
				    Else
					    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
							    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
							    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
					    con.execute sSql
				    End IF
		     Response.Write sSql & vbCrLf
	    '''***************************************************************************''''''''
	    else
		    If dMRSDate = "" then dMRSDate = IssDate
		    sTempMonYr = mid(dMRSDate,4,2)
		    sMonYr = sTempMonYr&Year(dMRSDate)

		    'arrFin = split(GetFinancialYear(sMonYr),":")
		    'sFinFrom = arrFin(0)
		    'sFinTo = arrFin(1)
		    arrFin = split(session("Finperiod"),":")
		    sFinFrom = "01/04/"&arrFin(0)
		    sFinTo = "31/03/"&arrFin(1)

		    with dcrs
			    .CursorLocation = 3
			    .CursorType = 3
			    .Source = "SELECT ISNULL(YEARRECEIPTQUANTITY,0),YEAROPENINGSTOCK,YEARISSUEQUANTITY,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			    .ActiveConnection = con
			    .Open
		    end with
		    set dcrs.ActiveConnection = nothing

		    ''Response.Write dcrs.source &"<BR>"

		    if not dcrs.EOF then
			    iRecQty = cdbl(dcrs(0))
			    iYrOpStock = cdbl(dcrs(1))
			    iYrIssQty = cdbl(dcrs(2))
			    iYrCloQty = cdbl(dcrs(3))
			    iYrCloValue = cdbl(dcrs(4))
		    end if
		    dcrs.Close
		    ''Response.Write "<P>iIssQty="&iIssQty

		    ''Response.Write iIssQty & " *  " & " ( " & iYrCloValue  & " / " & iYrCloQty  & " ) "
		    If iYrCloQty <> 0 then 	iValue = iIssQty * Round((cdbl(iYrCloValue) / cdbl(iYrCloQty)),2)

			     ''Response.Write "%%%%%% iValue = "& iValue &"%%%%%%<BR>"

			    with dcrs2
				    .CursorLocation = 3
				    .CursorType = 3
				    .Source = "SELECT ISNULL(MAX(LEDGERENTRYNO)+1,1) FROM INV_T_ITEMLEDGER"
				    .ActiveConnection = con
				    .Open
			    end with
			    if not dcrs2.EOF then
				    iLedgEntNo = dcrs2(0)
			    end if
			    dcrs2.close
			    
			    ''added the condition by ragav on Sep 12,2012 for Pick Later case stock updation stoped other case it will happen
				''begin 
		         '   if cdbl(iValue)<>cdbl(0) then'
			     '       sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
				 '           "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,SENTTOACCOUNTS,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
				 '           "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
				 '           "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(iTransDate) & ",103)," & iIssQty & "," & Round(iValue) & ",'T',"& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
		    	 '       Response.Write vbCrLf &"INV_T_ITEMLEDGER="&sSql & vbCrLf & vbCrLf
			     '       con.execute sSql
		         '   end if
		          
		         ' if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
		         '       if cdbl(iValue)<>cdbl(0) then'
			      '          sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
				   '             "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,SENTTOACCOUNTS,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
				    '            "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
				    '            "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(iTransDate) & ",103)," & iIssQty & "," & Round(iValue) & ",'T',"& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
		    	    '        Response.Write vbCrLf &"INV_T_ITEMLEDGER="&sSql & vbCrLf & vbCrLf
			        '        con.execute sSql
		            '    end if
		         ' end if 'if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
		          ''end
	    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			    with dcrs2
				    .CursorLocation = 3
				    .CursorType = 3
				    .Source = "SELECT  YEARCLOSINGVALUE,YEARCLOSINGSTOCK FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    .ActiveConnection = con
				    .Open
			    end with
			    if not dcrs2.EOF then
				    iClosingVal = dcrs2(0)
				    iClosingStk = dcrs2(1)
			    end if
			    dcrs2.close


			    with dcrs
				    .CursorLocation = 3
				    .CursorType = 3
				    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    .ActiveConnection = con
				    .Open
			    end with
			    set dcrs.ActiveConnection = nothing



			    if dcrs.EOF then
				    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
					    "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
					    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iIssQty & "," &_
					    "" & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
			    	Response.Write sSql & vbCrLf & vbCrLf
				    con.execute sSql
			    else
			    	'Response.Write "<p>"&iClosingVal  &"  >  "& iValue &"  and  "& iClosingStk &"  >  "& iIssQty &"<BR><BR>"
				    if cdbl(iClosingVal) > cdbl(Round(iValue,2)) and cdbl(iClosingStk)  >  cdbl(iIssQty)  then
					    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
						    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND "&_
						    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
						    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    else
					    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						    "YEARCLOSINGSTOCK = (" & iIssQty & " - YEARCLOSINGSTOCK), " &_
						    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND "&_
						    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
						    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    end if
				    Response.Write "<p> Ningth = "&sSql & vbCrLf & vbCrLf
				    con.execute sSql
			    end if
			    dcrs.Close

	    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			    with dcrs
				    .CursorLocation = 3
				    .CursorType = 3
				    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    .ActiveConnection = con
				    .Open
			    end with
			    set dcrs.ActiveConnection = nothing
			    if not dcrs.EOF then
				    if cdbl(dcrs(0)) < Round(iValue) then
					    iValue = cdbl(dcrs(0))
				    else
					    iValue = iValue
				    end if
			    end if
			    dcrs.Close
			    if cdbl(iClosingVal) > cdbl(Round(iValue,2)) and cdbl(iClosingStk)  >  cdbl(iIssQty)  then
				    sSql = " UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
					       " YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
					       " YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
					       " ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
					       " ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
					       " CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					       " CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
				    con.execute sSql
			    else
				    sSql = " UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
					       "YEARCLOSINGSTOCK = (" & iIssQty & " - YEARCLOSINGSTOCK), " &_
					       " YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
					       " ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
					       " ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
					       " CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					       " CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
				    con.execute sSql
			    end if
			    
			    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
                sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
                sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                Response.Write "<p>"& sSql 
	            con.execute sSql 
    '					'Response.Write "END IF"
				 '   If iMRSNo <> "" then
				'	    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
				'			    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
				'			    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & "  AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
				'	    con.execute sSql
				 '   Else
				'	    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
				'			    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
				'			    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
				'	    con.execute sSql
				 '   End IF
				  '   Response.Write "<p>"& sSql & vbCrLf & vbCrLf

	    end if ' if IsArray(arrSerialP) then
	''Response.Write "END IF"
	'end if 'if 1 = 2 then
	
	'Response.Write UBound(arrSerialP)

	    if IsArray(arrSerialP) then

		    for icnt = 0 to UBound(arrSerialP)
			    sTempSerialNo = sTempSerialNo & ","& arrSerialP(icnt)
		    next
			    sTempSerialNo = mid(sTempSerialNo,2)
		    ''Response.Clear
			    with dcrs
				    .CursorLocation = 3
				    .CursorType = 3
				    .Source = "SELECT ISNULL(LOTNUMBER,0),ISNULL(SERIALNUMBER,0),LotQuantityNett FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " and SERIALNUMBER in ("& sTempSerialNo &") "
			    '	'Response.Write dcrs.source
				    .ActiveConnection = con
				    .open
			    end with
			    set dcrs.ActiveConnection = nothing

			    if not dcrs.EOF then
				    iTempIssQty = iIssQty
				    do while not dcrs.eof
					    if cdbl(iTempIssQty)<>0 and not cdbl(iTempIssQty) < 0 then
						    iLotNo =  dcrs(0)
						    iSerNo  = dcrs(1)
						    IF trim(iLotNo) = "NULL" then iLotNo = "0"
						    if trim(iSerNo)="NULL" then iSerNo="0"

'					    Response.Write " Mark Pick insert iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo
'					    response.write "<p>Sixth Updation = "
					    
					    
					        ''added the condition by ragav on Sep 12,2012 for Pick Later case stock updation stoped other case it will happen
				            ''begin 
			                'UpdateLocLot iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iTempIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
			                    if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
			                        UpdateLocLot iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iTempIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
			                    end if 'if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
			                ''end
			                
						    iTempIssQty = cdbl(iTempIssQty) - cdbl(dcrs(2))
					    end if ' if trim(iTempIssQty)<>"0" then
					    dcrs.movenext
				    loop
			    end if
			    dcrs.close
			    ''Response.Write iLotNo
			    ''Response.Write "************************"&iLotNo &" MARK PICK INSERT ************************************"


	    else
		    iLotNo = sPickLot
		    iSerNo = ""
		    with dcrs
			    .CursorLocation = 3
			    .CursorType = 3
			    .Source = "SELECT ISNULL(LOTNUMBER,0),ISNULL(SERIALNUMBER,0),LotQuantityNett FROM INV_T_LOCATIONLOT WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) &" and LotNumber = "& iLotNo &" and (LotQuantityNett - QuantityIssued) > 0"
			    ''Response.Write dcrs.source
			    .ActiveConnection = con
			    .open
		    end with
		    set dcrs.ActiveConnection = nothing

		    if not dcrs.EOF then
			    iTempIssQty = iIssQty
			    do while not dcrs.eof
				    if cdbl(iTempIssQty)<>0 and not cdbl(iTempIssQty) < 0 then
					    iLotNo = dcrs(0)
					    iSerNo  = dcrs(1)

					    IF trim(iLotNo) = "NULL" then iLotNo = "0"
					    if trim(iSerNo)="NULL" then iSerNo="0"

					    ''Response.Write " Mark Pick insert iLotNo ="& iLotNo  & " iSeriesNo ="& iSerNo
'					    response.write "<p>Seventh Updation = "
'					    Response.write "<p>Serial Number = "& iSerNo
					    
					     ''added the condition by ragav on Sep 12,2012 for Pick Later case stock updation stoped other case it will happen
				            ''begin 
			                'UpdateLocLot iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
			                    if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
'			                        Response.write "<p>Welcome to Lot Updation"
			                        UpdateLocLot iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
			                    end if 'if trim(sPickPackFlag)<>"L" or trim(sIssType)="F" then
			                ''end
					    iTempIssQty = cdbl(iTempIssQty) - cdbl(dcrs(2))
				    end if
				    dcrs.movenext
			    loop
		    end if
		    dcrs.close
		    ''Response.Write iLotNo

		    ''Response.Write "************************"&iLotNo &" MARK PICK INSERT ************************************"

		    'if trim(iLotNo)  = "0" and trim(iSerNo) = "0" then

		    'end if
	    end if 'if IsArray(arrSerialP) then


		    if trim(iMRSNo) <> "" then
			    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
				    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
				    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
		    else
			    sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
				    "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
				    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
		    end if
	    	''Response.Write sSql & vbCrLf & vbCrLf
		    con.execute sSql
	    '	'Response.Write iIssueNo
	    '''***************************************************************************''''''''
	    ' end if for receipt qty check




	    ''Response.Write Cdbl(iLedIssueNo)&"."&CDbl(iMRSNo)
    end function
%>

<%
'-------------------------------------------------------------------------------------------
Function FindMonthYear(dtTansDate)

Dim rsTemp,sQuery,rsNew,sRetValue

Set rsTemp  = Server.CreateObject("ADODB.RecordSet")
Set rsNew	= Server.CreateObject("ADODB.RecordSet")
	sQuery = "Select Month(convert(datetime,'" & dtTansDate & "',103))"
	''Response.Write "<p>"& sQuery
	with rsTemp
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.Open
	end with

	set rsTemp.ActiveConnection = nothing

	if rsTemp(0) >= 10 then

		sQuery = "Select Cast(Month(convert(datetime,'" & dtTansDate & "',103)) As Varchar) + Cast(Year(convert(datetime,'" & dtTansDate & "',103)) As Varchar)"
		''Response.Write "<p>"& sQuery
		with rsNew
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		end with

		set rsNew.ActiveConnection = nothing
		sRetValue = rsNew(0)
		rsNew.Close
	else
		sQuery = "Select '0'+Cast(Month(convert(datetime,'" & dtTansDate & "',103)) As varchar) + Cast(Year(convert(datetime,'" & dtTansDate & "',103)) As Varchar)"
		''Response.Write "<p>"& sQuery
		with rsNew
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		end with

		set rsNew.ActiveConnection = nothing
		sRetValue = rsNew(0)
		rsNew.Close
	end if
	rsTemp.Close

	FindMonthYear = sRetValue
End Function
'-------------------------------------------------------------------------------------------
Function FindFirstDate(sPassedMonthYear)

Dim rsTemp,sQuery,dtFirstDate

Set rsTemp  = Server.CreateObject("ADODB.RecordSet")

    ''Response.Write "<p> " & sPassedMonthYear


	sQuery  = "Select convert(varchar,'01" & "/" & left(sPassedMonthYear,2) & "/" & right(sPassedMonthYear,4) & "',103)"
	''Response.Write "<p> " & sQuery

	with rsTemp
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.Open
	end with

	set rsTemp.ActiveConnection = nothing

	if not rsTemp.Eof then
		dtFirstDate = rsTemp(0)
	end if
	rsTemp.close

	FindFirstDate = dtFirstDate

	''Response.Write "<p> " & dtFirstDate
	''Response.End
End Function
'-------------------------------------------------------------------------------------------
%>
