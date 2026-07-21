<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PopIntRcptData.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	Ragavendran R
	'Created On					:	Mar 24,2014
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->

<%
Dim objDom,rsTemp,rsHead,rsDet,rsStore,rsLotSer,rsLotSerdet
Dim ndRoot,ndDet,ndItemDet,ndStore,ndLotser,ndLotserDet
Dim sAutomaticAccount,sQuery,sRcvdOn,sAppRefType,sAppRefNo,sAppRefDate,RcptNo,sByProduct,sLot,sAttList,sIntRcptBin
Dim sItemAttList,sInvRecNo
Dim iEntryNo,iItemRate,iCounter,iLotser
Set objDom = Server.CreateObject("Microsoft.XMLDOM")
set rsTemp = Server.CreateObject("ADODB.Recordset")
set rsHead = Server.CreateObject("ADODB.Recordset")
set rsDet = Server.CreateObject("ADODB.Recordset")
set rsStore = Server.CreateObject("ADODB.Recordset")
set rsLotSer = Server.CreateObject("ADODB.Recordset")
set rsLotSerdet = Server.CreateObject("ADODB.Recordset")

RcptNo = Request("RcptNo")

sQuery = "Select AutomaticAccounting from APP_M_ApplicationSetup where ApplicationCode= 6 and ReferenceCodeNo = 4"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    sAutomaticAccount = rsTemp(0)
end if
rsTemp.Close 

set ndRoot = objDom.createElement("ROOT")
objDom.appendChild ndRoot

if trim(RcptNo)<>"" then
    sQuery = "Select CreatedFromDept,RefType,OrganisationCode,IsNull(AppRefType,''),IsNull(AppRefNo,''),IsNull(AppRefDate,CreatedOn),CreatedOn,IsNull(InvRecNo,0) from APP_T_InternalReceiptHeader where InternalReceiptNo = " & RcptNo
    rsHead.Open sQuery,con
    if not rsHead.EOF then
      sAppRefType = rsHead(3)
      sAppRefNo = rsHead(4)
      sAppRefDate = rsHead(5)
      sRcvdOn = rsHead(6)
      sInvRecNo = rsHead(7)
      ndRoot.setAttribute "DEPT",rsHead(0)
      ndRoot.setAttribute "SOURCE",rsHead(1)
      ndRoot.setAttribute "ORGCODE",rsHead(2)
      ndRoot.setAttribute "STYPE",""
      ndRoot.setAttribute "ITEMTYPE",""
      ndRoot.setAttribute "PACKNUM",""
      ndRoot.setAttribute "SRCREFTYPE","N"
      ndRoot.setAttribute "SRCREFNO",""
      ndRoot.setAttribute "RCPTNUMBERINV",""
      ndRoot.setAttribute "APPREFTYPE",sAppRefType
      ndRoot.setAttribute "APPREFNO",sAppRefNo
      ndRoot.setAttribute "APPREFDATE",sAppRefDate
      ndRoot.setAttribute "RCVDON",sRcvdOn
      ndRoot.setAttribute "AUTOACCOUNT",sAutomaticAccount
    end if 
    rsHead.Close 

            iEntryNo = 0
            set ndDet = objDom.createElement("Details")
            ndRoot.appendChild ndDet     
            sQuery = "Select D.ItemCode,D.ClassificationCode,D.MRSNumber,D.IssueNo,V.OrganisationCode,V.Itemdescription,"
            sQuery = sQuery &" V.StoresUOM,V.ReceiptNumbering,IsNull(ProductType,'P'),IsNull(SUM(QuantityReturn),0)"
            sQuery = sQuery &" from APP_T_InternalReceiptDetails D left join VWItem V  on D.ItemCode = V.ItemCode where "
            sQuery = sQuery &" InternalReceiptNo = "& RcptNo &" Group By D.ItemCode,D.ClassificationCode,D.MRSNumber,D.IssueNo,"
            sQuery = sQuery &" V.OrganisationCode,V.Itemdescription,V.StoresUOM,V.ReceiptNumbering,ProductType"
            rsDet.Open sQuery,con
            if not rsDet.EOF then
                do while not rsDet.EOF 
                    iEntryNo = iEntryNo + 1
                        
                    sByProduct = rsDet(8)
                    'sItemAttList = rsDet(7)
                    set ndItemDet  = objDom.createElement("ItemDetail")
                        ndItemDet.setAttribute "ItemCode",rsDet(0)
                        ndItemDet.setAttribute "CLACODE",rsDet(1)
                        ndItemDet.setAttribute "QTY",rsDet(9)
                        ndItemDet.setAttribute "MRSNO","N"
                        ndItemDet.setAttribute "ISSNO","N"
                        ndItemDet.setAttribute "ENTRYNO",iEntryNo
                        ndItemDet.setAttribute "UNIT",rsDet(4)
                        ndItemDet.setAttribute "ITEMNAME",rsDet(5)
                        ndItemDet.setAttribute "UOM",rsDet(6)
                        ndItemDet.setAttribute "ATTRIBUTELIST",""
                        ndItemDet.setAttribute "RefNo",""
                        ndItemDet.setAttribute "ReqQty",""
                        ndItemDet.setAttribute "RECEIPTNUM",rsDet(7)
                        ndItemDet.setAttribute "BYPRODUCT",sByProduct
                        ndDet.appendChild ndItemDet
                        
                        if Trim(sInvRecNo)<>"0" and Trim(sInvRecNo)<>"" then
                        
                            sQuery= "Select TransactValue/TransactQuantity from INV_T_ITEMLEDGER Where TransactionNo in (Select InvRecNo from APP_T_InternalReceiptheader "
                            sQuery = sQuery & " where InternalReceiptNo = "& RcptNo &") and ItemCode = "& rsDet(0)
                            rsStore.Open sQuery,con
                            if not rsStore.EOF then
                                iItemRate = rsStore(0)
                            end if
                            rsStore.Close 
                            ndItemDet.setAttribute "ITEMRATE",iItemRate
                            
                            
                            sQuery = "Select IsNull(StorageLocationNo,0),IsNull(StorageBinNumber,0),SUM(LotQuantityNett),ItemCode,ClassificationCode,"
                            sQuery = sQuery & " OrganisationCode from INV_T_LocationLot where InventoryReceiptNo in (Select InvRecNo from APP_T_InternalReceiptheader "
                            sQuery = sQuery & " where InternalReceiptNo = "& RcptNo &")  and ItemCode = "& rsDet(0) 
                            if Trim(sItemAttList)<>"" then
                                sQuery = sQuery & " and AttributeList = "& sItemAttList
                            end if
                            sQuery = sQuery &" Group By StorageLocationNo,StorageBinNumber,ItemCode,ClassificationCode,OrganisationCode"
                            rsStore.open sQuery,con
                            if not rsStore.EOF then
                                do while not rsStore.EOF 
                                    sIntRcptBin= rsStore(1)
                                    if trim(sIntRcptBin)="" or Trim(sIntRcptBin)="0" or IsNull(sIntRcptBin) then sIntRcptBin="NULL"
                                    set ndStore = objDom.createElement("STORAGE")
                                        ndStore.setAttribute "STORE",rsStore(0)
                                        ndStore.setAttribute "BIN",sIntRcptBin
                                        ndStore.setAttribute "APPLICABLE","IN"
                                        ndStore.setAttribute "MONTHYEAR",""
                                        ndStore.setAttribute "QTY",rsStore(2)
                                        ndStore.setAttribute "STORAGEVALUE",cdbl(rsStore(2))*cdbl(iItemRate)
                                        ndStore.setAttribute "CLASSIFICATION",""
                                        ndStore.setAttribute "UNIT",rsDet(4)
                                    ndItemDet.appendChild ndStore
                                    
                                    
                                    
                                    iCounter = 0
                                    sQuery = "Select SUM(LotQuantityTare),IsNull(LotNumber,''),IsNull(AttributeList,''),SUM(LotQuantityGross),SUM(LotQuantityNett)*Rate,ItemCode,ClassificationCode from INV_T_LocationLot "
                                    sQuery = sQuery & " where InventoryReceiptNo in (Select InvRecNo from APP_T_InternalReceiptheader where  InternalReceiptNo = "& RcptNo &")  and StorageLocationNo = "& rsStore(0) &" and ItemCode = "& rsDet(0) 
                                    if Trim(sItemAttList)<>"" then
                                        sQuery = sQuery & " and AttributeList = "& sItemAttList
                                    end if
                                    sQuery = sQuery & " Group By LotNumber,AttributeList,Rate,ItemCode,ClassificationCode"
                                    'Response.Write sQuery&vbCrLf 
                                    rsLotSer.Open sQuery,con
                                    if not rsLotSer.EOF then
                                        do while not rsLotSer.EOF 
                                        iCounter = iCounter + 1
                                        sLot = rsLotSer(1)
                                        sAttList = rsLotSer(2)
                                        set ndLotSer = objDom.createElement("LotSerial")
                                            ndLotSer.setAttribute "QTYIN","N"
                                            ndLotSer.setAttribute "TARE",rsLotSer(0)
                                            ndLotSer.setAttribute "LOT",sLot
                                            ndLotSer.setAttribute "SERIALFROM",""
                                            ndLotSer.setAttribute "SERIALTO",""
                                            ndLotSer.setAttribute "TAREWEIGHT","U"
                                            ndLotSer.setAttribute "ATTLIST",sAttList
                                            ndLotSer.setAttribute "QTY",rsLotSer(3)
                                            ndLotSer.setAttribute "COUNTER",iCounter
                                            ndLotSer.setAttribute "STAGE","N"
                                            ndLotSer.setAttribute "ALTGROSS",""
                                            ndLotSer.setAttribute "ALTNETT",""
                                            ndLotSer.setAttribute "ALTUOM",""
                                            ndLotSer.setAttribute "IVALUE",rsLotSer(4)
                                            ndLotSer.setAttribute "AUTOGEN",""
                                            ndLotSer.setAttribute "TAREELIGIBLE","Y"
                                            ndLotSer.setAttribute "SUBLEVEL",""
                                            
                                            ndStore.appendChild ndLotSer
                                            
                                            
                                            iLotSer =0
                                            sQuery="Select LotQuantityGross,LotQuantityTare,IsNull(SellingNumber,0),WeightPerSellingForm,PackingCode,IsNull(LotNumber,''),"
                                            sQuery = sQuery & " IsNull(SellingForm,0),PackingNumber,Rate,IsNull(AttributeList,''),StockQuality,LotQuantityNett from INV_T_LocationLot where "
                                            sQuery = sQuery & " InventoryReceiptNo in (Select InvRecNo from APP_T_InternalReceiptheader where "
                                            sQuery = sQuery & " InternalReceiptNo = "& RcptNo &")  and StorageLocationNo = "& rsStore(0) &" and ItemCode = "& rsDet(0) &""
                                            
                                            if Trim(sLot)<>"" then
                                                sQuery = sQuery & " and LotNumber = "& sLot
                                            end if 
                                            if trim(sAttList)<>"" then
                                                sQuery = sQuery & " and AttributeList= "& sAttList
                                            end if
                                            sQuery = sQuery &" Order By cast(PackingNumber as numeric)"
                                            'Response.Write sQuery&vbCrLf 
                                            rsLotSerDet.Open sQuery,con
                                            if not rsLotSerDet.EOF then
                                                do while not rsLotSerDet.EOF 
                                                iLotSer=iLotSer+1
                                                set ndLotSerDet = objDom.createElement("LotSerialDetails")
                                                    ndLotSerDet.setAttribute "LOTSERIAL",iLotSer
                                                    ndLotSerDet.setAttribute "QTYREC",rsLotSerDet(0)
                                                    ndLotSerDet.setAttribute "TAREREC",rsLotSerDet(1)
                                                    ndLotSerDet.setAttribute "SELLINGTYPE",rsLotSerDet(2)
                                                    ndLotSerDet.setAttribute "WEIGHTSTYPE",rsLotSerDet(3)
                                                    ndLotSerDet.setAttribute "PACKINGTYPE",rsLotSerDet(4)
                                                    ndLotSerDet.setAttribute "LOT",rsLotSerDet(5)
                                                    ndLotSerDet.setAttribute "SELLINGFORM",rsLotSerDet(6)
                                                    ndLotSerDet.setAttribute "PACKNUMBER",rsLotSerDet(7)
                                                    ndLotSerDet.setAttribute "IVALUE",cdbl(rsLotSerDet(11))*cdbl(iItemRate)
                                                    ndLotSerDet.setAttribute "ATTRIBUTELIST",rsLotSerDet(9)
                                                    ndLotSerDet.setAttribute "NOOFCONE",""
                                                    ndLotSerDet.setAttribute "SUBLEVELID",""
                                                    ndLotSerDet.setAttribute "SQ",rsLotSerDet(10)
                                                    ndLotSerDet.setAttribute "STATUS","O" ' New-N,Old-O,Update-U,Delete -D
                                                    ndLotSer.appendchild ndLotSerDet 
                                                    rsLotSerDet.MoveNext 
                                                loop
                                            end if
                                            rsLotSerDet.Close 
                                            rsLotSer.MoveNext 
                                        loop
                                    end if
                                    rsLotSer.Close 
                                    rsStore.MoveNext 
                                loop
                            end if
                            rsStore.Close 
                        else
                            sQuery= "Select Rate from APP_T_InternalReceiptDetails Where InternalReceiptNo = "& RcptNo &" and ItemCode = "& rsDet(0)
                            rsStore.Open sQuery,con
                            if not rsStore.EOF then
                                iItemRate = rsStore(0)
                            end if
                            rsStore.Close 
                            ndItemDet.setAttribute "ITEMRATE",iItemRate
                            
                            
                            sQuery = "Select IsNull(Store,0),IsNull(Bin,0),SUM(QuantityReturn),ItemCode,ClassificationCode, "
                            sQuery = sQuery &" OrganisationCode from APP_T_InternalReceiptheader H join APP_T_InternalReceiptDetails D "
                            sQuery = sQuery & "  on H.InternalReceiptNo = D.InternalReceiptNo  where H.InternalReceiptNo = "& RcptNo &" and ItemCode = "& rsDet(0)
                            if Trim(sItemAttList)<>"" then
                                sQuery = sQuery & " and AttributeList = "& sItemAttList
                            end if
                            sQuery = sQuery & " Group By Store,Bin,ItemCode,ClassificationCode,OrganisationCode"
                            rsStore.open sQuery,con
                            if not rsStore.EOF then
                                do while not rsStore.EOF 
                                    sIntRcptBin= rsStore(1)
                                    if trim(sIntRcptBin)="" or Trim(sIntRcptBin)="0" or IsNull(sIntRcptBin) then sIntRcptBin="NULL"
                                    set ndStore = objDom.createElement("STORAGE")
                                        ndStore.setAttribute "STORE",rsStore(0)
                                        ndStore.setAttribute "BIN",sIntRcptBin
                                        ndStore.setAttribute "APPLICABLE","IN"
                                        ndStore.setAttribute "MONTHYEAR",""
                                        ndStore.setAttribute "QTY",rsStore(2)
                                        ndStore.setAttribute "STORAGEVALUE",cdbl(rsStore(2))*cdbl(iItemRate)
                                        ndStore.setAttribute "CLASSIFICATION",""
                                        ndStore.setAttribute "UNIT",rsDet(4)
                                    ndItemDet.appendChild ndStore
                                    
                                    
                                    
                                    iCounter = 0
                                    sQuery= "Select SUM(GrossQuantityReturn-QuantityReturn),IsNull(LotNo,''),IsNull(AttributeList,''),"
                                    sQuery = sQuery &" SUM(GrossQuantityReturn),SUM(QuantityReturn)*Rate,ItemCode,ClassificationCode from "
                                    sQuery = sQuery &" APP_T_InternalReceiptDetails where InternalReceiptNo = "& RcptNo 
                                    sQuery = sQuery &" and Store = "& rsStore(0) &" and ItemCode = "& rsDet(0)
                                    if Trim(sItemAttList)<>"" then
                                        sQuery = sQuery & " and AttributeList = "& sItemAttList
                                    end if
                                    sQuery = sQuery &"  Group By LotNo,AttributeList,Rate,ItemCode,ClassificationCode "
                                    'Response.Write sQuery&vbCrLf 
                                    rsLotSer.Open sQuery,con
                                    if not rsLotSer.EOF then
                                        do while not rsLotSer.EOF 
                                        iCounter = iCounter + 1
                                        sLot = rsLotSer(1)
                                        sAttList = rsLotSer(2)
                                        set ndLotSer = objDom.createElement("LotSerial")
                                            ndLotSer.setAttribute "QTYIN","N"
                                            ndLotSer.setAttribute "TARE",rsLotSer(0)
                                            ndLotSer.setAttribute "LOT",sLot
                                            ndLotSer.setAttribute "SERIALFROM",""
                                            ndLotSer.setAttribute "SERIALTO",""
                                            ndLotSer.setAttribute "TAREWEIGHT","U"
                                            ndLotSer.setAttribute "ATTLIST",sAttList
                                            ndLotSer.setAttribute "QTY",rsLotSer(3)
                                            ndLotSer.setAttribute "COUNTER",iCounter
                                            ndLotSer.setAttribute "STAGE","N"
                                            ndLotSer.setAttribute "ALTGROSS",""
                                            ndLotSer.setAttribute "ALTNETT",""
                                            ndLotSer.setAttribute "ALTUOM",""
                                            ndLotSer.setAttribute "IVALUE",rsLotSer(4)
                                            ndLotSer.setAttribute "AUTOGEN",""
                                            ndLotSer.setAttribute "TAREELIGIBLE","Y"
                                            ndLotSer.setAttribute "SUBLEVEL",""
                                            
                                            ndStore.appendChild ndLotSer
                                            
                                            
                                            iLotSer =0
                                            sQuery= "Select GrossQuantityReturn,GrossQuantityReturn-QuantityReturn,IsNull(SellingNumber,0),WeightPerSellingForm,PackingCode,"
                                            sQuery = sQuery & " IsNull(LotNo,''),IsNull(PackingForm,0),PackingNum,Rate,IsNull(AttributeList,''),StockQuality,QuantityReturn  "
                                            sQuery = sQuery & " from APP_T_InternalReceiptDetails  where  InternalReceiptNo = "& RcptNo &"  and Store = "& rsStore(0)
                                            sQuery= sQuery &" and ItemCode ="& rsDet(0)
                                            
                                            if Trim(sLot)<>"" then
                                                sQuery = sQuery & " and LotNo = "& sLot
                                            end if 
                                            if trim(sAttList)<>"" then
                                                sQuery = sQuery & " and AttributeList= "& sAttList
                                            end if
                                            sQuery = sQuery &" Order By cast(PackingNum as numeric)"
                                            
                                            'Response.Write sQuery&vbCrLf 
                                            rsLotSerDet.Open sQuery,con
                                            if not rsLotSerDet.EOF then
                                                do while not rsLotSerDet.EOF 
                                                iLotSer=iLotSer+1
                                                set ndLotSerDet = objDom.createElement("LotSerialDetails")
                                                    ndLotSerDet.setAttribute "LOTSERIAL",iLotSer
                                                    ndLotSerDet.setAttribute "QTYREC",rsLotSerDet(0)
                                                    ndLotSerDet.setAttribute "TAREREC",rsLotSerDet(1)
                                                    ndLotSerDet.setAttribute "SELLINGTYPE",rsLotSerDet(2)
                                                    ndLotSerDet.setAttribute "WEIGHTSTYPE",rsLotSerDet(3)
                                                    ndLotSerDet.setAttribute "PACKINGTYPE",rsLotSerDet(4)
                                                    ndLotSerDet.setAttribute "LOT",rsLotSerDet(5)
                                                    ndLotSerDet.setAttribute "SELLINGFORM",rsLotSerDet(6)
                                                    ndLotSerDet.setAttribute "PACKNUMBER",rsLotSerDet(7)
                                                    ndLotSerDet.setAttribute "IVALUE",cdbl(rsLotSerDet(11))*cdbl(iItemRate)
                                                    ndLotSerDet.setAttribute "ATTRIBUTELIST",rsLotSerDet(9)
                                                    ndLotSerDet.setAttribute "NOOFCONE",""
                                                    ndLotSerDet.setAttribute "SUBLEVELID",""
                                                    ndLotSerDet.setAttribute "SQ",rsLotSerDet(10)
                                                    ndLotSerDet.setAttribute "STATUS","O" ' New-N,Old-O,Update-U,Delete -D
                                                    ndLotSer.appendchild ndLotSerDet 
                                                    rsLotSerDet.MoveNext 
                                                loop
                                            end if
                                            rsLotSerDet.Close 
                                            rsLotSer.MoveNext 
                                        loop
                                    end if
                                    rsLotSer.Close 
                                    rsStore.MoveNext 
                                loop
                            end if
                            rsStore.Close 
                        end if 'if Trim(sInvRecNo)<>"0" and Trim(sInvRecNo)<>"" then
                    rsDet.MoveNext 
                loop
            end if
            rsDet.Close 
            
                    
                    
                  
        
end if 'if trim(RcptNo)<>"" then	


    Response.ContentType = "text/xml"
	Response.Write objDom.xml

  
%>