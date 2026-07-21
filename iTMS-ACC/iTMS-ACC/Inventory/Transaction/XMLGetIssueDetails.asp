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
	'Program Name				:	XMLGetIssueDetails.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	
	'Created On					:	
	'Modified BY				:	RAGAVENDRAN R
	'Modified On				:	April 17,2014
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	MaterialIssueEntry.asp
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
<!-- #include File="../../include/purpopulate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->

<%
    Dim OutDataXML,rsHead,rsItem,rsPick,rsStore,rsSerial
    Dim ndRoot,ndItem,ndPick,ndPickNode,ndSerHdr,ndSerDet,ndSubcontract,ndDetails
    Dim sQuery
    Dim sIssEntNo,sIssType,sIssFrom,sIssToType,sIssToCode,sIssToSubCode,sPackPackFlag,sReturnable,sReturnItem
    Dim sAppRefType,sAppRefNo,sAppRefDate,sType,sIssueDate,sOrgCode,sItemDesc,sRcptNum
    Dim sLoc,sBin,sLotNo
    
    Dim iItemEntNo,iItemCode,iClassCode,iIssQty,iIssValue,iCount
    
    set OutDataXML = Server.CreateObject("Microsoft.XMLDOM")
    set rsHead = Server.CreateObject("ADODB.Recordset")
    set rsItem = Server.CreateObject("ADODB.Recordset")
    set rsPick = Server.CreateObject("ADODB.Recordset")
    set rsStore = Server.CreateObject("ADODB.Recordset")
    set rsSerial = server.CreateObject("ADODB.Recordset")
    
    sIssEntNo = Request("IssEntNo")
    
    sQuery = "Select IssueType,IssueFrom,IssuedToType,IssuedToCode,IsNull(IssuedToSubCode,''),IsNull(MarkPackFlag,'N'),"
    sQuery = sQuery &" IsNull(Returnable,'N'),IsNull(ReturnItem,'S'),IsNull(AppRefType,''),IsNull(AppRefNo,0),IsNull(AppRefDate,getdate()),"
    sQuery = sQuery &" IssueTypeCode,Convert(varchar,IssueDate,103),OrganisationCode from INV_T_MaterialIssueHeader where IssueEntryNo = "& sIssEntNo
    'Response.Write sQuery & vbCrLf 
    rsHead.Open sQuery,con
    if not rsHead.EOF then
        sIssType = rsHead(0)
        sIssFrom = rsHead(1)
        sIssToType = rsHead(2)
        sIssToCode = rsHead(3)
        sIssToSubCode = rsHead(4)
        sPackPackFlag = rsHead(5)
        sReturnable = rsHead(6)
        sReturnItem = rsHead(7)
        sAppRefType = rsHead(8)
        sAppRefNo = rsHead(9)
        sAppRefDate = rsHead(10)
        sType = rsHead(11)
        sIssueDate = rsHead(12)
        sOrgCode = rsHead(13)
    end if 
    rsHead.Close 
    
    set ndRoot = OutDataXML.createElement("ISSTYPE")
    ndRoot.setAttribute "ISSTYPE",sIssType
    ndRoot.setAttribute "ISSTOTYPE",sIssToType
    ndRoot.setAttribute "ISSTOCODE",sIssToCode
    ndRoot.setAttribute "ISSTOSUBCODE",sIssToSubCode
    ndRoot.setAttribute "POConfirm","N"
    ndRoot.setAttribute "SInvConfirm","N"
    ndRoot.setAttribute "Invoice","A"
    ndRoot.setAttribute "GPConfirm","N"
    ndRoot.setAttribute "ProConfirm","N"
    ndRoot.setAttribute "MCallFrom","MRIssue"
    ndRoot.setAttribute "RedirectTo","ISSUEMGMT.ASP"
    ndRoot.setAttribute "AppRefType",sAppRefType
    ndRoot.setAttribute "AppRefNo",sAppRefNo
    ndRoot.setAttribute "AppRefDate",sAppRefDate
    ndRoot.setAttribute "ConsumptionAccHead",""
    ndRoot.setAttribute "IssueToCode",sIssToCode
    ndRoot.setAttribute "PickPackFlag",sPackPackFlag
    ndRoot.setAttribute "IssFrom",sIssFrom
    ndRoot.setAttribute "Returnable",sReturnable
    ndRoot.setAttribute "ReturnItem",sReturnItem
    ndRoot.setAttribute "TYPE",sType
    OutDataXML.appendChild ndRoot
    iItemEntNo = 0
    iCount = 0
    sQuery = "Select D.ItemCode,D.ClassificationCode,V.ItemDescription,SUM(D.QuantityIssued),"
    sQuery = sQuery &" V.StoresUOM,IsNull(D.Returnable,'N'),IsNull(D.ReturnItem,'S'),IsNull(D.ItemAttributes,'') from INV_T_MaterialIssueDetails D join VWITEM V on  D.ItemCode = "
    sQuery = sQuery &" V.ItemCode where IssueEntryNo = "& sIssEntNo 
    sQuery = sQuery &" Group By D.ItemCode,D.ClassificationCode,V.ItemDescription,V.StoresUOM,D.Returnable,D.ReturnItem,D.ItemAttributes"
    rsItem.Open sQuery,con
    if not rsItem.EOF then
        do while not rsItem.EOF 
            iItemEntNo = iItemEntNo + 1
             iCount = 0
            iItemCode = rsItem(0)
            iClassCode = rsItem(1)
            sItemDesc = rsItem(2)
            iIssQty = rsItem(3)
            iIssValue = rsItem(4)
            set ndItem = OutDataXML.createElement("ITEMDETAILS")
                ndRoot.appendChild ndItem
                ndItem.setAttribute "ENTRYNO",iItemEntNo
                ndItem.setAttribute "ITEMCODE",iItemCode
                ndItem.setAttribute "CLASSCODE",iClassCode
                ndItem.setAttribute "UNIT",sOrgCode
                ndItem.setAttribute "ITEMNAME",sItemDesc
                ndItem.setAttribute "UOM",rsItem(4)
                ndItem.setAttribute "DECIMAL",""
                ndItem.setAttribute "DISPALYED","N"
                ndItem.setAttribute "QTY",iIssQty
                ndItem.setAttribute "REQUIREDBY","" 
                ndItem.setAttribute "REQUIREDVALUE",""
                ndItem.setAttribute "ATTRIBUTELIST",rsItem(7)
                ndItem.setAttribute "RefNo",""
                ndItem.setAttribute "ReqQty",iIssQty
                ndItem.setAttribute "ONLYLOT","" 
                ndItem.setAttribute "RETURNABLE",rsItem(5)
                ndItem.setAttribute "RETURNITEM",rsItem(6)
                
                sQuery = "Select SUM(QuantityIssued),count(*) from INV_T_MaterialIssueDetails where IssueEntryNo = "& sIssEntNo &" and ItemCode = "& iItemCode 
                if Trim(rsItem(7))<>"" then
                    sQuery = sQuery & " and ItemAttributes = "& rsItem(7)
                end if 
                rsPick.Open sQuery,con
                if not rsPick.EOF then
                    set ndPick = OutDataXML.createElement("Pick")
                    ndPick.setAttribute "TOT",rsPick(0)
                    ndPick.setAttribute "NoofPack",rsPick(1)
                    ndItem.appendChild ndPick
                    
                    sRcptNum = GetItemRcptNum(iItemCode)
                    if Trim(sRcptNum)="N" then
                        sQuery = "Select LocationNumber,IsNull(BinNumber,0),SUM(QuantityIssued),IsNull(LotNo,''),count(*) from INV_T_MaterialIssueDetails "
                        sQuery = sQuery & " where IssueEntryNo = "& sIssEntNo &" and ItemCode = "& iItemCode &" Group by LocationNumber,BinNumber,LotNo"
                        rsStore.Open sQuery,con
                        if not rsStore.EOF then
                            do while not rsStore.EOF 
                                sLoc = rsStore(0)
                                sBin = rsStore(1)
                                sLotNo = rsStore(3)
                                iCount = iCount + 1
                                if Trim(sLotNo)="" or IsNull(sLotNo) then sLotNo = "N/A"
                                set ndPickNode = OutDataXML.createElement("STORE")
                                    ndPickNode.setAttribute "LOC",sLoc
                                    ndPickNode.setAttribute "BIN",sBin
                                    ndPickNode.setAttribute "LOTNO",sLotNo
                                    ndPickNode.setAttribute "INVRECNO",""
                                    ndPickNode.setAttribute "QTYISS",rsStore(2)
                                    ndPickNode.setAttribute "NoofPack",rsStore(4)
                                    ndPickNode.setAttribute "Count",iCount
                                ndPick.appendChild ndPickNode
                                rsStore.MoveNext 
                            loop
                        end if
                        rsStore.Close 
                    else
                        sQuery = "Select LocationNumber,IsNull(BinNumber,0),SUM(QuantityIssued),IsNull(LotNo,''),count(*) from INV_T_MaterialIssueDetails "
                        sQuery = sQuery & " where IssueEntryNo = "& sIssEntNo &" and ItemCode = "& iItemCode 
                        if Trim(rsItem(7))<>"" then
                            sQuery = sQuery & " and ItemAttributes = "& rsItem(7)
                        end if 
                        sQuery = sQuery &" Group by LocationNumber,BinNumber,LotNo"
                        Response.Write vbCrLf + sQuery + vbCrLf 
                        
                        rsStore.Open sQuery,con
                        if not rsStore.EOF then
                            do while not rsStore.EOF 
                                sLoc = rsStore(0)
                                sBin = rsStore(1)
                                sLotNo = rsStore(3)
                                iCount = iCount + 1
                                if Trim(sLotNo)="" or IsNull(sLotNo) then sLotNo = "N/A"
                                set ndPickNode = OutDataXML.createElement("PICK")
                                    ndPickNode.setAttribute "LOC",sLoc
                                    ndPickNode.setAttribute "BIN",sBin
                                    ndPickNode.setAttribute "LOTNO",sLotNo
                                    ndPickNode.setAttribute "INVRECNO",""
                                    ndPickNode.setAttribute "QTYISS",rsStore(2)
                                    ndPickNode.setAttribute "NoofPack",rsStore(4)
                                    ndPickNode.setAttribute "Count",iCount
                                ndPick.appendChild ndPickNode
                                
                                sQuery = "Select D.SerialNo,D.QuantityIssued,L.PackingNumber from INV_T_MaterialIssueDetails D join INV_T_LocationLot L on D.SerialNo=L.SerialNumber and D.ItemCode = L.ItemCode "
                                sQuery = sQuery & " where D.IssueEntryNo = "& sIssEntNo &" and D.ItemCode = "& iItemCode 
                                if Trim(sLotNo)<>"" and Trim(sLotNo)<>"N/A" then
                                    sQuery = sQuery & " and L.LotNumber = "& pack(sLotNo)
                                end if 
                                if Trim(rsItem(7))<>"" then
                                    sQuery = sQuery & " and L.AttributeList =D.ItemAttributes and L.AttributeList = "& rsItem(7)
                                end if 
                                Response.Write sQuery + vbCrLf 
                                rsSerial.Open sQuery,con
                                if not rsSerial.EOF then
                                    do while not rsSerial.EOF 
                                        set ndSerDet = OutDataXML.createElement("Selection")
                                        ndSerDet.setAttribute "SerialNo",rsSerial(0)
                                        ndSerDet.setAttribute "StockQty",rsSerial(1)
                                        ndSerDet.setAttribute "Qty",rsSerial(1)
                                        ndSerDet.setAttribute "YesNo","Y"
                                        ndSerDet.setAttribute "PackNo",rsSerial(2)
                                        ndPickNode.appendChild ndSerDet
                                        rsSerial.MoveNext 
                                    loop
                                end if 
                                rsSerial.Close 
                                
                                set ndSerHdr = OutDataXML.createElement("SERIALHEADER")
                                ndPickNode.appendChild ndSerHdr 
                                
                                sQuery = "Select SerialNo,QuantityIssued from INV_T_MaterialIssueDetails "
                                sQuery = sQuery & " where IssueEntryNo = "& sIssEntNo &" and ItemCode = "& iItemCode 
                                if Trim(sLotNo)<>"" and Trim(sLotNo)<>"N/A" then
                                    sQuery = sQuery &" and LotNo = "& pack(sLotNo)
                                end if 
                                rsSerial.Open sQuery,con
                                if not rsSerial.EOF then
                                    do while not rsSerial.EOF 
                                        set ndSerDet = OutDataXML.createElement("SERIALDETAILS")
                                            ndSerDet.setAttribute "SERIALNO",rsSerial(0)
                                            ndSerDet.setAttribute "QTY",rsSerial(1)
                                        ndSerHdr.appendChild ndSerDet 
                                        rsSerial.MoveNext 
                                    loop
                                end if
                                rsSerial.Close 
                                
                                rsStore.MoveNext 
                            loop
                        end if
                        rsStore.Close 
                    end if 
                end if
                rsPick.Close 
            rsItem.MoveNext 
        loop
    end if 
    rsItem.Close 
    
    
    sQuery =" Select SCProcess,Instruct,LabourCharge,HardWaste,IntWaste,SubconProcessName from INV_T_Materialissueheader H join App_M_SubContractProcess P on H.SCProcess=P.SubconProcessID where IssueEntryNo = "& sIssEntNo 
    
    rsItem.Open sQuery,con
    if not rsItem.EOF then
        set ndSubcontract = OutDataXML.createElement("SubContract")
        ndSubcontract.setAttribute "SCProcess",rsItem(0)
        ndSubcontract.setAttribute "Instruct",rsItem(1)
        ndSubcontract.setAttribute "LabourCharge",rsItem(2)
        ndSubcontract.setAttribute "Currency","INR"
        ndSubcontract.setAttribute "HardWaste",rsItem(3)
        ndSubcontract.setAttribute "InvWaste",rsItem(4)
        ndSubcontract.setAttribute "ProcessName",rsItem(5)
        ndRoot.appendChild ndSubcontract 
        
        sQuery = "Select ItemCode,ClassificationCode,ItemDesc from INV_T_MaterialIssueReturnItem where IssueEntryNo = "& sIssEntNo 
        rsStore.Open sQuery,con
        if not rsStore.EOF then
            do while not rsStore.EOF 
                set ndDetails = OutDataXML.createElement("Details")
                ndDetails.setAttribute "MatRecdAsItem",rsStore(0)
                ndDetails.setAttribute "MatRecdAsCode",rsStore(1)
                ndDetails.setAttribute "MatRecdAsDescr",rsStore(2)
                ndSubcontract.appendChild ndDetails
                rsStore.MoveNext 
            loop
        end if
        rsStore.Close 
    end if
    rsItem.Close 
    OutDataXML.save server.MapPath("../Temp/Transaction/MaterialIssueEditPopulate.xml")
    Response.Clear 
	Response.ContentType="text/xml"
	Response.Write OutDataXML.xml
%>