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
    Dim iInvRecNo,nExistTransQty,nExistTransVal,iTotLocQtyIssued
    Dim sRcptNumber
    
    Dim bEligible

    set dRSet = Server.CreateObject("ADODB.Recordset")
    dTDate = FormatDate(date())
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
    

    Set dcrs = Server.CreateObject("ADODB.RecordSet")
    Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
    
    

Function MrsIssueUpdate(iLedIssueNo)
'Response.Clear 
    Set RootO = OutData.createElement("Root")
    OutData.appendChild RootO
	
	'First time XML is Not Loaded.So,XML Load is Put inside the Function
	if ObjFs1.FileExists(server.MapPath("../Temp/Transaction/mrsIssueDataEdit"&Session.SessionID&".xml")) then
	IssXML.load(Server.MapPath("../Temp/Transaction/mrsIssueDataEdit"&Session.SessionID&".xml"))
	end if 'if ObjFs1.FileExists(server.MapPath("../Temp/Transaction/mrsIssueDataEdit"&Session.SessionID&".xml")) then
	
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
    
	
	sIssueExp = "//ISSTYPE/ITEM"
	set ndIssueTempNode = RootNode.selectNodes(sIssueExp)
	if ndIssueTempNode.Length>0 then
	    iNumIssueClassCode = ndIssueTempNode.Item(0).Attributes.getNamedItem("CLACODE").Value
	    Response.Write "<p> ClassCode = "& iNumIssueClassCode
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
    
    if trim(IssToType)="" or isNull(IssToType) then IssToType = "NULL"
    if trim(IssToCode)="" or isNull(IssToCode) then IssToCode="NULL"
    if trim(IssToSubCode)="" or isNull(IssToSubCode) then IssToSubCode="NULL"
    
    if trim(IssToType)<>"" and trim(IssToType)<>"NULL" then IssToType = Pack(IssToType)
    if trim(IssToSubCode)<>"" and trim(IssToSubCode)<>"NULL" then IssToSubCode = Pack(IssToSubCode)
    
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
		    'Response.Write "sAttList ="& sAttList 
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
		
            If trim(sAttID)="" or IsNull(sAttID) then sAttID = "NULL"

		    for each PickNode in HeaderNode.childNodes
			    if strcomp(PickNode.nodeName,"Pick")=0 then
				    iNoofCases = PickNode.getAttribute("NoofPack")
			    end if
		    next
			    
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

    if trim(sPartyCode)="" then sPartyCode="NULL"
			    
		    sSql = " "

		    'End
		    
		    if trim(iGenCode)="" or IsNull(iGenCode) then iGenCode = "NULL"
		    if trim(iGenCode)<>"NULL" then iGenCode = Pack(iGenCode)

	    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    if trim(iNoofCases)="" or isnull(iNoofCases) then iNoofCases = 0
	    
		    if trim(sType)="SER" then '' For  Services Case
		    
		        sSql = "Select GatePassNo from FORGATEPASSHEADER where AppRefType = 12 and AppRefNo = "& iLedIssueNo 
			    rsTemp.Open sSql,con
			    if not rsTemp.EOF then
				    sGatePassNo =  rsTemp(0)
			    end if
			    rsTemp.Close
		    end if
		    
		    sSql = "Select SerialNo,LocationNumber,QuantityIssued,ItemCode,BinNumber,LotNo,ClassificationCode from INV_T_Materialissuedetails where IssueEntryNo = "& iLedIssueNo 
            Response.Write "<p>"& sSql
            With dcrs 
                .CursorLocation = 3
                .CursorType = 3
                .ActiveConnection = con
                .Source = sSql 
                .Open 
            End With
            if not dcrs.eof then
                do while not dcrs.EOF 
                    
                    sSql = "Select ReceiptNumbering from INV_M_ItemMaster where ItemCode = "& dcrs(3)
                    dcrs1.open sSql,con
                    if not dcrs1.eof then
                        sRcptNumber = dcrs1(0)
                    end if 
                    dcrs1.close 
                    if trim(sRcptNumber)="N" then
                        RestoreLocLot dcrs(3),dcrs(6),sOrgID,dcrs(1),dcrs(4),dcrs(2),dcrs(0),dcrs(5)
                    else
                        sSql = "Update INV_T_LocationLot set QuantityIssued = QuantityIssued - "& dcrs(2) &" where SerialNumber ="& dcrs(0) &" and StorageLocationNo = "& dcrs(1) &" and ItemCode = "& dcrs(3)
                        Response.Write "<p>"& sSql
                        con.execute sSql
                    end if 
                    dcrs.MoveNext 
                loop
            end if 
            dcrs.close 
		    
		    
		    sSql = "Delete from INV_T_MaterialConsumption where IssueEntryNo = "& iLedIssueNo 
		    Response.Write "<p>"& sSql 
		    con.execute sSql 
		    
            sSql = "Delete from INV_T_MaterialConsumptionDetail where IssueEntryNo = "& iLedIssueNo 
            Response.Write "<p>"& sSql 
		    con.execute sSql 
		    
		    sSql = "Delete from Inv_T_MaterialIssueAdditionalDetails where IssueEntryNo = "& iLedIssueNo 
		    Response.Write "<p>"& sSql 
		    con.execute sSql 
		    
		    sSql = "Delete from INV_T_MaterialIssueReturnItem where IssueEntryNo = "& iLedIssueNo 
		    Response.Write "<p>"&sSql
		    con.execute sSql
		    
		    sSql = "Delete from INV_T_MaterialIssueDetails where IssueEntryNo = "& iLedIssueNo 
		    Response.Write "<p>"&sSql 
		    con.execute sSql 
		    
		    sSql = "Delete from INV_T_MaterialIssuedForPick where IssueEntryNo = "& iLedIssueNo 
		    Response.Write "<p>"&sSql 
		    con.execute sSql 
		    
		    sSql = "Select IssueEntryCode from INV_T_MaterialIssueHeader where IssueEntryno = "& iLedIssueNo 
		    rsTemp.open sSql,con
		    if not rsTemp.eof then
		        iGenCode = rsTemp(0)
		    end if
		    rsTemp.close 
		    
		    sSql = "Delete from INV_T_MaterialIssueHeader where IssueEntryNo = "& iLedIssueNo 
		    Response.Write "<p>"&sSql 
		    con.execute sSql 
		    
		    sSql = "INSERT INTO INV_T_MaterialIssueHeader (IssueEntryNo,OrganisationCode,IssueEntryCode,ReferenceType," &_
			       "MaterialReceivedBy,Remarks,IssueDate,IssuedBy,CreatedOn,IssuedToCode,IssuedToSubCode,IssuedToType,IssueType,AppRefType,AppRefNo,AppRefDate,MarkPackFlag,IssueFrom,Returnable,ReturnItem,IssueTypeCode) VALUES " &_
			       "(" & iLedIssueNo & "," & Pack(sOrgID) & "," & pack(iGenCode) & ",'D' ," & Pack(sRecBy) & "," &_
			       "" & Pack(sRem) & ",Convert(DateTime," & Pack(IssDate) & ",103),"&iIssuedBy&"," &_
			       " Convert(datetime,'"& IssDate &"',103),'"& IssToCode &"',"& IssToSubCode &","& IssToType &",'"& sIssType &"',"& sAppRefType &","& sAppRefNo &",'"&sAppRefDate&"',"&Pack(sPickPackFlag)&","& pack(sIssFrom) &","& pack(sReturnable) &","& pack(sReturnItem) &","& pack(sType) &")"
		     Response.Write "<p>"&sSql & vbCrLf
		    con.execute sSql
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


            Response.Write "<p>IssDate = "& IssDate
			    bFlag = true
			    IF dMRSDate = "" then dMRSDate = IssDate
			    sTempMonYr = mid(dMRSDate,4,2)
			    sMonYr = sTempMonYr&Year(dMRSDate)

			    arrFin = split(session("Finperiod"),":")
			    sFinFrom = "01/04/"&arrFin(0)
			    sFinTo = "31/03/"&arrFin(1)
				    with dcrs
					    .CursorLocation = 3
					    .CursorType = 3
					    .Source = "SELECT ISNULL(LOCATIONNUMBER,0),ISNULL(BINNUMBER,0) FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
					    .ActiveConnection = con
					    .Open
				    end with
				    set dcrs.ActiveConnection = nothing
				    if not dcrs.EOF then
					    sLoc = dcrs(0)
					    sBin = dcrs(1)
				    end if
				    dcrs.close
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
				    .Open
			    end with
			    if not rsTemp.EOF then
			        Response.Write "<p>Value = "& rsTemp(0) &"<p>Stock="& rsTemp(1)
				    if trim(rsTemp(0))<>"0" and trim(rsTemp(1))<>"0" then
					    iValue = cdbl(rsTemp(0))/cdbl(rsTemp(1))
				    end if
			    end if
			    rsTemp.Close
			    Response.Write "<p>Value = "& iValue
			    if trim(iValue)="" or IsNull(iValue) then iValue = "0"
			    iValue = Round(iValue,2)
			    
			    sSql = "Select SUM(TransactQuantity),SUM(TransactValue) from INV_T_ItemLedger where TransactionNo = "& iLedIssueNo &" and TransactionType = 'I' and ItemCode = "& iItemCode 
			    if Trim(sAttID)<>"" and Trim(sAttID)<>"NULL" then 
			    sSql = sSql &" and AttributeList = "& sAttID
			    end if 
			    Response.Write "<p>"& sSql
			    dcrs.Open sSql,con
			    if not dcrs.EOF then
			        nExistTransQty = dcrs(0)
			        nExistTransVal = dcrs(1)
			    end if
			    dcrs.Close 
			    
			    Response.Write "<p> nExistTransQty = "& nExistTransQty
			    Response.Write "<p> nExistTransVal = "& nExistTransVal
			    
			    sSql = "Select IsNull(YearIssueValue,0) from INV_T_ItemYearlyStock where ITemCode = "& iItemCode  &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
			    dcrs.Open sSql,con
			    if not dcrs.EOF then
			        if cdbl(dcrs(0))<cdbl(nExistTransVal) then
			            nExistTransVal = 0
			        end if
			    end if
			    dcrs.Close 
			    Response.Write "<P>iValue = "& iValue
			    sSql = "Select IsNull(YearClosingValue,0),IsNull(YearClosingStock,0),IsNull(YearReceiptQuantity,0),IsNull(YearReceiptValue,0),IsNull(YearOpeningStock,0),IsNull(YearOpeningValue,0) from INV_T_ItemYearlyStock where "
			    sSql = sSql & " ITemCode = "& iItemCode  &" and FinancialYearFrom = Convert(datetime,'"&  sFinFrom &"',103)"
			    Response.Write "<p>"& sSql 
			    dcrs.Open sSql,con
			    if not dcrs.EOF then
			        if cdbl(dcrs(1))>0 then
			        iValue = cdbl(dcrs(0))/cdbl(dcrs(1))
			        elseif cdbl(dcrs(2))>0 then
			        iValue = cdbl(dcrs(3))/cdbl(dcrs(2))
			        elseif cdbl(dcrs(4))>0 then
			        iValue = cdbl(dcrs(5))/cdbl(dcrs(4))
			        end if 
			        Response.Write "<P>iValue = "& iValue
			    end if
			    dcrs.Close 
			    iValue = Round(iValue,2)
			    Response.Write "<P>iValue = "& iValue
			    
			    sSql = "Update INV_T_ItemYearlyStock set YearIssueQuantity=YearIssueQuantity-"& nExistTransQty &",YearIssueValue=YearIssueValue-"& nExistTransVal &" where ItemCode = "& iItemCode 
			    sSql = sSql &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
			    Response.Write "<p>"& sSql 
			    con.execute sSql
			    
			    sSql = "Update INV_T_ItemYearlyStock set YearIssueQuantity=YearIssueQuantity+"& iIssQty &",YearIssueValue=YearIssueValue+"& cdbl(iValue)*cdbl(iIssQty) 
	            sSql = sSql & " where ItemCode = "& iItemCode &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
	            Response.Write "<p>"& sSql 
	            con.execute sSql 

                sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity"
                sSql = sSql & " where ItemCode = "& iItemCode 
                sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                Response.Write "<p>"& sSql 
	            con.execute sSql 
	            
	            sSql = "Select YearClosingStock,YearClosingValue from INV_T_ItemYearlyStock where ItemCode = "& iItemCode 
	            sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
	            dcrs.open sSql,con
	            if not dcrs.eof then
	                Response.Write "<p>Year Closing Stock = "& dcrs(0)
	                Response.Write "<p>Year Closing value = "& dcrs(1)
	            end if 
	            dcrs.close
	            
	            sSql = "Select LocationNumber from INV_M_ItemStorage where ItemCode = "& iItemCode 
	            dcrs.Open sSql,con
	            if not dcrs.EOF then
	                sLoc = dcrs(0)
	            end if
	            dcrs.Close 
	            
	            sSql = "Select IsNull(YearIssueValue,0) from INV_T_ItemLocationStock where ItemCode ="& iItemCode &" and LocationNumber ="& sLoc &" and FinancialYearFrom =Convert(datetime,'"& sFinFrom &"',103)"
	            dcrs.open sSql,con
	            if not dcrs.eof then
	                if cdbl(nExistTransVal)>cdbl(dcrs(0)) then
	                    nExistTransVal = dcrs(0)
	                end if 
	            else
	                nExistTransVal = 0
	            end if 
	            dcrs.close 
	            
	            sSql = "Update INV_T_ItemLocationStock set YearIssueQuantity=YearIssueQuantity-"& nExistTransQty &",YearIssueValue=YearIssueValue-"& nExistTransVal &"  where ItemCode = "& iItemCode 
			    sSql = sSql &" and LocationNumber = "& sLoc &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
			    Response.Write "<p>"& sSql 
			    con.execute sSql
			  

                sSql = " Update INV_T_ItemLocationStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity"
                sSql = sSql & "  where ItemCode = "& iItemCode 
                sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103) and LocationNumber = "& sLoc 
                Response.Write "<p>"& sSql 
	            con.execute sSql 
	            
	            sSql = " Select YearClosingStock,YearClosingValue from INV_T_ItemLocationStock  where ItemCode = "& iItemCode 
                sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103) and LocationNumber = "& sLoc 
                dcrs.open sSql,con
	            if not dcrs.eof then
	                Response.Write "<p>Location Closing Stock = "& dcrs(0)
	                Response.Write "<p>Location Closing Value = "& dcrs(1)
	            end if 
	            dcrs.close
			    
			    sSql = "Update INV_T_ItemLedger set TransactQuantity = "& iIssQty &",TransactValue="& cdbl(iValue)*cdbl(iIssQty) &" where TransactionNo = "& iLedIssueNo &" and TransactionType= 'I' and ItemCode ="&iItemCode 
			    if Trim(sAttributeList)<>"" then
			        sSql = sSql &" and AttributeList = "& sAttributeList
			    end if 
			    Response.Write "<p>"&sSql 
			    con.execute sSql 
			    
			    Response.Write " <p> Isstype  = "& sIssType 

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
									    sLoc	  = PickDetNode.Attributes.Item(0).nodeValue
									    sBin	  = PickDetNode.Attributes.Item(1).nodeValue
									    iLotNo	  = PickDetNode.Attributes.Item(2).nodeValue
									    iInvRecNo = PickDetNode.Attributes.Item(3).nodeValue
									    iQtyIss   = PickDetNode.Attributes.Item(4).nodeValue
									  
									    
									    
									    if trim(iLotNo)="N/A" then iLotNo=""
									    if PickDetNode.HaschildNodes() then
									        Response.Write "<p>Welcome to PICK NODE"
										    For Each SerNode in PickDetNode.childnodes
										    iTotLocQtyIssued = 0
										    Response.Write "<p>SerNode = "& SerNode.nodeName
											    if StrComp(SerNode.NodeName,"SERIALHEADER") = 0 then
											        if SerNode.hasChildNodes() then
											    Response.Write "<p>Welcome to SerialHeader"
												    For Each SerDetNode in SerNode.childnodes
													    if StrComp(SerDetNode.NodeName,"SERIALDETAILS") = 0 then
														    iSerNo  =  SerDetNode.Attributes.Item(0).nodeValue
														    iSerQty =  SerDetNode.Attributes.Item(1).nodeValue
														    
														    if Trim(sAppRefType)="18" then
														        sSql = "Select * from Sal_T_InvPackDetails where SaleTransactionNo = "& sAppRefNo &" and InventoryReceiptSerialNo in ("& iSerNo &")"
														        dcrs.Open sSql,con
														        if not dcrs.EOF then
														            bEligible = true
														        else
														            bEligible = false
														        end if
														        dcrs.Close 
														    else
														        bEligible = true
														    end if 
														    Response.Write "<p> Eligible = "&bEligible
														    
														    if bEligible then
														    
														    iTotLocQtyIssued = CDbl(iTotLocQtyIssued)+CDbl(iSerQty)

														        with dcrs
															        .cursorLocation = 3
															        .cursorType = 3
															        .ActiveConnection = con
															        if trim(iSerNo)="NULL" then
																        .source = "Select InventoryReceiptNo from VW_ITEMLOCATIONLOT_STOCK where SerialNumber is null and ItemCode ="& iItemCode
															        else
																        .source = "Select InventoryReceiptNo from VW_ITEMLOCATIONLOT_STOCK where SerialNumber ="& iSerNo
															        end if
														    	    Response.Write "<p>"& dcrs.source
															        .open
														        end with
														        if not dcrs.eof then
															        iInvRecNo = dcrs(0)
														        end if
														        dcrs.close

														        Dim iPackCode,iPackNo,iLotGrossQty,iLotNettQty,WPerSellForm,SellNumber,NoofSellForm
    														    
															        with dcrs
																        .CursorLocation = 3
																        .CursorType = 3
																        .Source = "SELECT ISNULL(RATE,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ISNULL(ITEMENTRYNO,0) = " & iEntNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (LOTNUMBER = '" & iLotNo & "' or LotNumber is Null)"
															    	    Response.Write "<p>"&"Lot ="& dcrs.source
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
    														    	
															        if cdbl(iSerQty)>0 then
															            if (trim(sPickPackFlag)<>"L" and trim(sOnlyLotFlag)="'P'") or trim(sIssType)="F" then
															                'Response.write "<p>First UpdateLocation Lot"
																            UpdateLocLot  iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerNo,iLotNo
																        end if
															        end if

															        sSql = "Select IsNull(YearClosingValue,0),IsNull(YearClosingStock,0),IsNull(YearReceiptQuantity,0),IsNull(YearReceiptValue,0),IsNull(YearOpeningStock,0),IsNull(YearOpeningValue,0) from INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
                                                                    Response.Write "<p>"& sSql 
                                                                    dcrs.Open sSql,con
                                                                    if not dcrs.EOF then
                                                                        if cdbl(dcrs(1))>0 then
                                                                            iIssVal = cdbl(iSerQty) * (cdbl(dcrs(0))/cdbl(dcrs(1)))
                                                                        elseif cdbl(dcrs(2))>0 then
                                                                            iIssVal = cdbl(iSerQty) * (cdbl(dcrs(3))/cdbl(dcrs(2)))
                                                                        elseif cdbl(dcrs(4))>0 then
                                                                            iIssVal = cdbl(iSerQty) * (cdbl(dcrs(5))/cdbl(dcrs(4)))
                                                                        else 
                                                                            iIssVal = 0
                                                                        end if 
                                                                        iIssVal = Round(iIssVal,2)
                                                                        Response.Write "<P>iValue = "& iValue
                                                                    end if
                                                                    dcrs.Close 
															        
    															    if trim(iIssVal)="" then
															            iIssVal = iSerQty * iItmRate
															        end if  'if trim(iIssVal)="" then
    															
															        IF cdbl(iSerQty)>0 then
    															    
															        sTempLotNo = iLotNo
															        if trim(sTempLotNo)="" or isNull(sTempLotNo) then sTempLotNo = "NULL"
															        if trim(sTempLotNo)<>"NULL"  then sTempLotNo = pack(sTempLotNo)
    															  	    Response.Write "<p>Pack Details"
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

		 													        End IF
    		 													    
    		 													    
		 													        if sAutoConsumption = "Y" then
		 													            sSql = "Insert into INV_T_MaterialConsumptionDetail (LineNumber,ConsumptionNo,IssueEntryNo,"&_
		 													                   " IssueDate,LotNo,SerialNo,QuantityConsumed,ConsumptionACHead,AttributeList) "&_
		 													                   " values("& sConLineNo &","& sConEntryNo &","& iLedIssueNo &",Convert(datetime,getDate(),103),"&_
		 													                   " '"& iLotNo &"',"& iSerNo &","& iSerQty &","& sConAccHead &","& sAttID &")"
                                                                        Response.Write "<p> Auto Consumption  header =  "& sSql &"<p>"
                                                                        con.execute sSql 
                                                                    end if
                                                                
                                                                end if 'if bEligible then

													    end if 'if StrComp(SerDetNode.NodeName,"SERIALDETAILS") = 0 then
													    
												    Next ' For Each SerDetNode in SerNode.childnodes
												    
												    Response.Write "<P>iValue = "& iValue
												   '     sSql = "Select IsNull(YearClosingValue,0),IsNull(YearClosingStock,0),IsNull(YearReceiptQuantity,0),IsNull(YearReceiptValue,0),IsNull(YearOpeningStock,0),IsNull(YearOpeningValue,0) from INV_T_ItemLocationStock where "
			                                        '    sSql = sSql & " ITemCode = "& iItemCode  &" and FinancialYearFrom = Convert(datetime,'"&  sFinFrom &"',103) and LocationNumber =" & sLoc
			                                         '   Response.Write "<p>"& sSql 
			                                         '   dcrs.Open sSql,con
			                                         '   if not dcrs.EOF then
			                                         '   Response.Write "<P>Closing Stock = "& dcrs(1)
			                                         '       if cdbl(dcrs(1))>0 then
			                                         '       iValue = cdbl(dcrs(0))/cdbl(dcrs(1))
			                                         '       Response.Write "<P>Step1"
			                                         '       elseif cdbl(dcrs(2))>0 then
			                                         '       iValue = cdbl(dcrs(3))/cdbl(dcrs(2))
			                                         '       Response.Write "<P>Step1"
			                                          '      elseif cdbl(dcrs(4))>0 then
			                                          '      iValue = cdbl(dcrs(5))/cdbl(dcrs(4))
			                                          '      Response.Write "<P>Step1"
			                                          '      end if 
			                                          '      Response.Write "<P>iValue = "& iValue
			                                          '  end if
			                                          '  dcrs.Close 
			                                           ' iValue = Round(iValue,2)
			                                            
												    
												    'iValue = 1
												        sSql = "Update INV_T_ItemLocationStock set YearIssueQuantity=YearIssueQuantity+"& iTotLocQtyIssued &",YearIssueValue=YearIssueValue+"& cdbl(iValue)*cdbl(iTotLocQtyIssued) 
									                    sSql = sSql & " where ItemCode = "& iItemCode &" and LocationNumber = "& sLoc &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
									                    Response.Write "<p>"& sSql 
									                    con.execute sSql 
									                    
									                    sSql = "Select IsNull(YearIssueValue,0),IsNull(YearIssueQuantity,0),IsNull(YearReceiptQuantity,0),IsNull(YearReceiptValue,0),IsNull(YearOpeningStock,0),IsNull(YearOpeningValue,0) from INV_T_ItemLocationStock where "
									                    sSql = sSql & " ITemCode = "& iItemCode &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom  &"',103) and LocationNumber =" & sLoc 
									                    Response.Write "<p>"& sSql 
									                    dcrs.open sSql,con
									                    if not dcrs.eof then
									                        Response.Write "<P>IsNull(YearIssueValue,0)="& dcrs(0)
									                        Response.Write "<P>IsNull(YearIssueStock,0)="& dcrs(1)
									                        Response.Write "<P>IsNull(YearReceiptQuantity,0)="& dcrs(2)
									                        Response.Write "<P>IsNull(YearReceiptValue,0)="& dcrs(3)
									                        Response.Write "<P>IsNull(YearOpeningStock,0)="& dcrs(4)
									                        Response.Write "<P>IsNull(YearOpeningValue,0)="& dcrs(5)
									                    end if
									                    dcrs.close
									                   
									                    
									                    
												        
                                                        sSql = " Update INV_T_ItemLocationStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                                                        sSql = sSql & " where ItemCode = "& iItemCode 
                                                        sSql = sSql & " and LocationNumber = "&sLoc &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                                                        Response.Write "<p>"& sSql 
									                    con.execute sSql 
												    end if 'if SerNode.hasChildNodes() then
											    end if 'if StrComp(SerNode.NodeName,"SERIALHEADER") = 0 then
										    Next 'For Each SerNode in PickDetNode.childnodes
									    end if 'if PickDetNode.HaschildNodes() then
									elseif StrComp(PickDetNode.nodeName,"STORE") = 0 then    
									        'Response.Write "<p>Hello Store"
									        
									        sLoc	  = PickDetNode.Attributes.Item(0).nodeValue
									        sBin	  = PickDetNode.Attributes.Item(1).nodeValue
									        iLotNo	  = PickDetNode.Attributes.Item(2).nodeValue
									        iInvRecNo = PickDetNode.Attributes.Item(3).nodeValue
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
												        ''Response.Write "<p>"&dcrs.source
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
											        Response.Write "<p>Store  Details"
										                if sAutoConsumption ="Y" then
									    	                sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,QuantityConsumed,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
													                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													             '   Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                Con.Execute sSql
    														        
												        else
												            sSql =	" INSERT INTO INV_T_MATERIALISSUEDETAILS (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,LotNo, "&_
													                " SerialNo,LocationNumber,BinNumber,IssueValue,QuantityIssued,ItemEntryNo,ItemAttributes,QuantityUOM,Returnable,ReturnItem,MaterialType) VALUES "&_
													                "(" & iLedIssueNo  & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & ","&_
													                "" & sTempLotNo & "," & iSerNo & "," & sLoc & "," & sBin & "," & iIssVal  & "," & iIssQty &"," & iEntNo & "," & sAttID & ",'"& trim(sSalUoM) &"','"&sIReturnable&"','"&sIReturnItem&"',"&sMatType&") "
													              '  Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
													                Con.Execute sSql
												        end if
												        Response.Write  "<p>"&sSql  & vbCrLf & vbCrLf
    														    
		 									        ''Added By Ragav on Apr 01 ,2010
		 									        
		 									        sSql = "Update INV_T_ItemLocationStock set YearIssueQuantity=YearIssueQuantity+"& iIssQty &",YearIssueValue=YearIssueValue+"& cdbl(iValue)*cdbl(iIssQty) 
									                    sSql = sSql & " where ItemCode = "& iItemCode &" and LocationNumber = "& sLoc &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
									                    Response.Write "<p>"& sSql 
									                    con.execute sSql 
									                    
									                    sSql = "Select IsNull(YearIssueValue,0),IsNull(YearIssueQuantity,0),IsNull(YearReceiptQuantity,0),IsNull(YearReceiptValue,0),IsNull(YearOpeningStock,0),IsNull(YearOpeningValue,0) from INV_T_ItemLocationStock where "
									                    sSql = sSql & " ITemCode = "& iItemCode &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom  &"',103) and LocationNumber =" & sLoc 
									                    Response.Write "<p>"& sSql 
									                    dcrs.open sSql,con
									                    if not dcrs.eof then
									                        Response.Write "<P>IsNull(YearIssueValue,0)="& dcrs(0)
									                        Response.Write "<P>IsNull(YearIssueStock,0)="& dcrs(1)
									                        Response.Write "<P>IsNull(YearReceiptQuantity,0)="& dcrs(2)
									                        Response.Write "<P>IsNull(YearReceiptValue,0)="& dcrs(3)
									                        Response.Write "<P>IsNull(YearOpeningStock,0)="& dcrs(4)
									                        Response.Write "<P>IsNull(YearOpeningValue,0)="& dcrs(5)
									                    end if
									                    dcrs.close
									                   
									                    
									                    
												        
                                                        sSql = " Update INV_T_ItemLocationStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                                                        sSql = sSql & " where ItemCode = "& iItemCode 
                                                        sSql = sSql & " and LocationNumber = "&sLoc &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                                                        Response.Write "<p>"& sSql 
									                    con.execute sSql 

		 									        if trim(iLotNo)="NULL" then iLotNo="0"
		 									        if trim(iSerNo)="NULL" then iSerNo="0"
										    end if 'if trim(iQtyIss)<>"0" then
									end if ' if StrComp(PickDetNode..nodeName,"PICK") = 0 then
							    Next 'For Each PickDetNode In PickNode.childNodes
						    end if 'if PickNode.HaschildNodes() then
						    
						    sSql = "Select YearClosingStock,YearClosingValue from INV_T_ItemYearlyStock where ItemCode = "& iItemCode 
	                        sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
	                        dcrs.open sSql,con
	                        if not dcrs.eof then
	                            Response.Write "<p>Year Closing Stock = "& dcrs(0)
	                            Response.Write "<p>Year Closing Stock = "& dcrs(1)
	                        end if 
	                        dcrs.close
	                        
	                        sSql = "Select YearClosingStock,YearClosingValue from INV_T_ItemLocationStock where ItemCode = "& iItemCode 
                            sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                            dcrs.open sSql,con
                            if not dcrs.eof then
                                Response.Write "<p>Location Closing Stock = "& dcrs(0)
                                Response.Write "<p>Location Closing Value = "& dcrs(1)
                            end if 
                            dcrs.close
						    
					        	    sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItemCode&"  and @ENTRYNO = "&iEntNo&"]/Pick/STORE"
				    
				                    'Response.Write "<p>sExp ="& sExp
				                    Set subPickNode = RootNode.Selectnodes(sExp)

    				                'Response.Write "PickNode.Length="&subPickNode.Length& vbCrLf
				                    if subPickNode.Length > 0 then
				                    
				                        if not cdbl(PickNode.getAttribute("TOT")) = 0 then
						                    iTotVal =   PickNode.getAttribute("TOT")
						                    For Each PickDetNode In PickNode.childNodes
                						    
						                        if PickDetNode.nodeName="STORE" then
                						    
						                    Response.Write "<p> PickDetNode.nodeName="& PickDetNode.nodeName
						                            sPickLoc = trim(PickDetNode.Attributes.getNamedItem("LOC").Value)
							                        sPickBin = trim(PickDetNode.Attributes.getNamedItem("BIN").Value)
							                        sPickLot = trim(PickDetNode.Attributes.getNamedItem("LOTNO").Value)
							                        iInvRecNo = trim(PickDetNode.Attributes.getNamedItem("INVRECNO").Value)
							                        sPickQty = trim(PickDetNode.Attributes.getNamedItem("QTYISS").Value)

						                            if ucase(sPickLot) = "N/A" then
								                        sPickLot = "NULL"
							                        else
								                        sPickLot = Pack(sPickLot)
							                        end if
							                        if sPickQty	= "" then sPickQty = iTotVal
							                        if cdbl(sPickQty) > 0 then
							                            if iInvRecNo = "" or IsNull(iInvRecNo) then iInvRecNo = "NULL"
                								        
								                        if sPickBin = "0" then sPickBin = "NULL"
                								        
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
				                    
				                     sSql = "Select YearClosingStock,YearClosingValue from INV_T_ItemYearlyStock where ItemCode = "& iItemCode 
	                        sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
	                        dcrs.open sSql,con
	                        if not dcrs.eof then
	                            Response.Write "<p>Year Closing Stock = "& dcrs(0)
	                            Response.Write "<p>Year Closing Stock = "& dcrs(1)
	                        end if 
	                        dcrs.close
	                        
	                        sSql = "Select YearClosingStock,YearClosingValue from INV_T_ItemLocationStock where ItemCode = "& iItemCode 
                            sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                            dcrs.open sSql,con
                            if not dcrs.eof then
                                Response.Write "<p>Location Closing Stock = "& dcrs(0)
                                Response.Write "<p>Location Closing Value = "& dcrs(1)
                            end if 
                            dcrs.close
				                    
					    end if 'if StrComp(PickNode.nodeName,"Pick") = 0 then
				    Next 'For Each PickNode In HeaderNode.childNodes
			    end if 	'if HeaderNode.HaschildNodes() then
			  end if ' if trim(sIssType)="F" then
		    end if 
		next
	end if 'if RootNode.hasChildNodes() then
				    
    subContDOM.save(Server.MapPath("../temp/transaction/PO_PUR_"&Session.SessionID&".xml"))
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 End Function 'Function MrsIssueInsert()
    
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

<%
Function RestoreLocLot(iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iSerialNo,iLotNo)
	dim dcrs,dcrs1,sSql,iInvRecNo,iQtyIss,iLotNetQty,iCtr,iTempQty,iChkIssQty,sAccountType
	
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	sSql = "Select IsNull(AccountingType,'W')  from Inv_M_ItemMaster where ItemCode in ("& iItemCode &")"
	dcrs.open sSql,con
	if not dcrs.eof then
	    sAccountType = dcrs(0)
	end if
	dcrs.close
	
	if Trim(sAccountType)="L" then
	        
	        if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
	        
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo &" and LotNumber = "& Pack(iLotNo) &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO "
	        elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and ( trim(iLotNo)="0" or trim(iLotNo)="") then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO "
	        elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID)  &" and LotNumber = "& Pack(iLotNo)  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO "
        				
	        else
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
					        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" " &_
					        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO "
	        end if
	elseif Trim(sAccountType)="F" or Trim(sAccountType)="W" then
	
	        if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
	        
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo &" and LotNumber = "& Pack(iLotNo) &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO desc" 
	        elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and ( trim(iLotNo)="0" or trim(iLotNo)="") then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO desc"
	        elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID)  &" and LotNumber = "& Pack(iLotNo)  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO desc"
        				
	        else
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
					        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" " &_
					        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO desc"
	        end if
	end if ' if Trim(sAccountType)="L" then
	
	
	Response.Write "<p> First = "&   sSql
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	'Response.Write "INV_T_LOCATIONLOT="& sSql & vbcrlf & vbcrlf 
	set dcrs.ActiveConnection = nothing
	iCtr = 1
	iChkIssQty = iIssQty 
	IF NOT dcrs.EOF THEN
		while not dcrs.EOF
		    
	'IF not  dcrs.EOF then
				iQtyIss = dcrs(0)
			'	Response.Write "iQtyIss = "& iQtyIss
				'iInvRecNo = dcrs(1)
				IF cdbl(iQtyIss) = cdbl(0) then	
				    if Trim(sAccountType)="L" then
				        if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc "
							
						elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="0" or trim(iLotNo)="") then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc "
							
						elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode  &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc "
							
						else
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc "
						end if
				    elseif Trim(sAccountType)="F" or Trim(sAccountType)="W" then	
						if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo "
							
						elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="0" or trim(iLotNo)="") then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo "
							
						elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode  &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo "
							
						else
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo "
						end if
					end if
						Response.Write "<p> ->SeconD = "&  sSql
								with dcrs1
									.CursorLocation = 3
									.CursorType = 3
									.Source = sSql
									.ActiveConnection = con
									.Open
								end with		
								if not dcrs1.EOF then
									while not dcrs1.EOF 
										iLotNetQty = dcrs1(0)
										iInvRecNo  = dcrs1(1)
										iCtr = iCtr + 1
									'	 Response.Write "<BR><BR>iChkIssQty="&iChkIssQty&"<BR><BR>"
										IF cdbl(iChkIssQty) <> 0 then 
									'	Response.Write "<P>"& iLotNetQty &"   <   "& iChkIssQty &"<BR><BR>"
											
											If cdbl(iLotNetQty) < cdbl(iChkIssQty) then 	
												iTempQty = iIssQty
												
												'iIssQty = iLotNetQty
												if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iLotNetQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber = "& iSerialNo & " and LotNumber ="& Pack(iLotNo) &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="0" or trim(iLotNo)="" ) then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iLotNetQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber = "& iSerialNo &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"" and trim(iLotNo)<>"0" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iLotNetQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and LotNumber ="& Pack(iLotNo) &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												else
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iLotNetQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) "
												end if

												'Response.Write "<p> lot= "& sSql & vbcrlf & vbcrlf 
												Con.Execute sSql			
												iChkIssQty = cdbl(iChkIssQty) - cdbl(iLotNetQty)
											Else
												
												if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"" and trim(iLotNo)<>"0" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iChkIssQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber ="& iSerialNo & " and LotNumber = "& Pack(iLotNo) &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
															
												elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="" or trim(iLotNo)="0") then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iChkIssQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber ="& iSerialNo &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
															
												elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"" and trim(iLotNo)<>"0" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iChkIssQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and LotNumber ="& Pack(iLotNo)  &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												else
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED -  "& iChkIssQty &") WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) "
												end if

												Response.Write "<p> lot1= "& sSql & vbcrlf & vbcrlf 
												Con.Execute sSql
												iChkIssQty = 0 'cint(iChkIssQty) - cint(iLotNetQty)
												
											End IF 'If cint(iLotNetQty) < cint(iIssQty) then 	
															
											iTempQty = cdbl(iTempQty) - cdbl(iIssQty)
										 End If
											'Response.Write "<BR><p>"&iChkIssQty&"   =   " &iTempQty&"<BR><BR>"
										dcrs1.MoveNext 				
									wend
								end if
								dcrs1.Close
					End IF
	
		
			dcrs.movenext
		wend
	End IF ' IF NOT DCRS.EOF THEN
	dcrs.Close
	
End Function

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
'				    	Response.Write "<p>1="&sSql & vbCrLf & vbCrLf
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
'					    	Response.Write sSql & vbCrLf & vbCrLf
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
						        'Response.Write " <p>FIRST = "& sSql & vbCrLf & vbCrLf
						    else
						        sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
							        "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIssQty & "),"&_
							        "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
							        "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
							        "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
							        "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
							        "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
						        'Response.Write " <p>FIRST = "& sSql & vbCrLf & vbCrLf
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
				    	       ' Response.Write "<p>"&sSql & vbCrLf & vbCrLf
					            
						    else
							    sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						            "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
						            "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
						            "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
						            "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						            "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						            "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				    	       ' Response.Write "<p>"&sSql & vbCrLf & vbCrLf
						    end if
						    con.execute sSql
					    end if
					    dcrs.Close
					    
					    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                        sSql = sSql & " where ItemCode = "& iItemCode 
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
'					    	Response.Write "<p>"&sSql & vbCrLf & vbCrLf
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
'						    	Response.Write sSql & vbCrLf & vbCrLf
							    con.execute sSql
						    else
							    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
							    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
								    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
								    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
								    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
								    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
								    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"

'							    Response.Write "<p>Second = "& sSql & vbCrLf & vbCrLf
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
'					    	Response.Write sSql & vbCrLf & vbCrLf
						    con.execute sSql
						    
						    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                            sSql = sSql & " where ItemCode = "& iItemCode 
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
'									    	Response.Write sSql & vbCrLf & vbCrLf
										    con.execute sSql
									    else
										    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iWMQty & ")," &_
										    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
											    "YEARRESERVED = (YEARRESERVED + " & iWMQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
											    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
											    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
											    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
											    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
'									    	Response.Write "<p>Thired = "& sSql & vbCrLf & vbCrLf
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
'								        Response.Write sSql & vbCrLf & vbCrLf
									    con.execute sSql
									    
									    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                                        sSql = sSql & "  where ItemCode = "& iItemCode 
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
						   ' Response.Write dcrs1.source
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
'								    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
								    con.execute sSql

								    if cdbl(iValue)<>cdbl(0) then
									    sSql = "INSERT INTO INV_T_ITEMLEDGER (LEDGERENTRYNO,ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
										    "TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NoOfPacks,ATTRIBUTELIST,PartyCode,IssueToRcptFrom) VALUES " &_
										    "(" & iLedgEntNo & "," & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
										    "'I'," & iLedIssueNo & ",CONVERT(DATETIME," & Pack(dMRSDate) & ",103)," & iWMQty & "," & Round(iValue) & ","& iNoofCases &","& sAttID &","& sPartyCode &",'"& IssToCode &"')"
'								    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
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
'								    	Response.Write sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    else
									    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iWMQty & ")," &_
									    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
										    "YEARRESERVED = (YEARRESERVED + " & iWMQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
										    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
										    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
										    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
'								    	Response.Write "<p>Fourth = "& sSql & vbCrLf & vbCrLf
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
'							    	Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
								    
								    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                                    sSql = sSql & "  where ItemCode = "& iItemCode 
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
'								    	Response.Write "<p>"& sSql & vbCrLf & vbCrLf
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
'								    	Response.Write sSql & vbCrLf & vbCrLf
									    con.execute sSql
								    else
									    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
									    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
										    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
										    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
										    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
										    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
'									    Response.Write "<p>Fifth = "& sSql & vbCrLf & vbCrLf
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
'								    Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
								    
								    
								    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                                    sSql = sSql & " where ItemCode = "& iItemCode 
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
'							    	Response.Write sSql & vbCrLf & vbCrLf
								    con.execute sSql
							    else
								    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
								    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIsQty & "),"&_
									    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
									    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
									    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
									    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
									    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
'								    Response.Write "<p>Sixth = "&sSql & vbCrLf & vbCrLf
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
'							    Response.Write sSql & vbCrLf & vbCrLf
							    con.execute sSql
							    
							    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                                sSql = sSql & " where ItemCode = "& iItemCode 
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


					'	    with dcrs
					'		    .CursorLocation = 3
					'		    .CursorType = 3
					'		    '.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
					'		    .Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
					'		   ' Response.Write "<p>"& dcrs.source
					'		    .ActiveConnection = con
					'		    .Open
					'	    end with
					'	    set dcrs.ActiveConnection = nothing
'
'						    if dcrs.EOF then
'							    sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
'								    "LOCATIONNUMBER,BINNUMBER,MONTHANDYEAR,YEARISSUEQUANTITY,YEARISSUEVALUE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
'								    "(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," & sLoc & "," & sBin & "," & iIssQty & "," &_
'								    "" & sLoc & "," & sBin & "," & iIssQty & "," & Round(iValue) & ",CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103),"& CInt(iIssQty) * -1 &","& CDbl(iValue) * -1 &")"
''						    	Response.Write sSql & vbCrLf & vbCrLf
'							    con.execute sSql
'						    else
'							    sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
'							    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - "& iIssQty & "), "&_
'								    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
'								    "CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
'								    "LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
'								    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND "&_
'								    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
'							    Response.Write "<p>Seventh = "& sSql & vbCrLf & vbCrLf
'							    con.execute sSql
'						    end if
'						    dcrs.Close
'




		    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						    sTempMonYr = mid(dMRSDate,4,2)
						    sMonYr = sTempMonYr&Year(dMRSDate)

						    'arrFin = split(GetFinancialYear(sMonYr),":")
						    'sFinFrom = arrFin(0)
						    'sFinTo = arrFin(1)
						    arrFin = split(session("Finperiod"),":")
						    sFinFrom = "01/04/"&arrFin(0)
						    sFinTo = "31/03/"&arrFin(1)
					'	    with dcrs
					'		    .CursorLocation = 3
					'		    .CursorType = 3
					'		    '.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					'		    .Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					'		    .ActiveConnection = con
					'		    .Open
					'	    end with
					'	    set dcrs.ActiveConnection = nothing
'
'						    if not dcrs.EOF then
'							    if cdbl(dcrs(0)) < iValue then
'								    iValue = cdbl(dcrs(0))
'							    else
'								    iValue = iValue
'							    end if
'						    end if
'						    dcrs.Close
'
'						    sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
'							    "YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
'							    "YEARRESERVED = (YEARRESERVED + " & iIssQty & ") WHERE " &_
'							    "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
'							    "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
'							    "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
'							    "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
'						    Response.Write "<p>"&sSql & vbCrLf & vbCrLf
'						    con.execute sSql
'						    
'						    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity,"
 '                           sSql = sSql & " YearClosingValue=YearOpeningValue+YearReceiptValue-YearIssueValue  where ItemCode = "& iItemCode 
  '                          sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
   '                         Response.Write "<p>"& sSql 
	'                        con.execute sSql 
'
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
    
 %>
 
 