<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	newreceiptInsert.asp
	'Module Name				:	Inventory (Internal Receipt Accounting)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 21, 2003
	'Modified By				:	KUMAR K A
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptEntry.asp
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
<%
Dim oDOM
Dim dcrs,dcrs1,dcrs2,dcrs3,adoCmd
Dim iAccountedBy,iRecNo,iInvRecNo,iItemEntryNo,iStore,iBin,iQty,iRate,iPackNum,iPackType,iSerNo,iStockQuality
Dim arrFin,sFinFrom,sFinTo,sQuery,dRcvd,sOrgID,sAttribute,iQtyGross,iQtyTare
Dim ndRoot,ndItem,ndStorage,ndLot
Dim sLotNo

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
Set dcrs3 = Server.CreateObject("ADODB.RecordSet")

iAccountedBy = getUserid

sFinFrom = Request("hFinFrom")
sFinTo = Request("hFinTo")
dRcvd = Request("hDate")
sOrgID = Request("hOrgID")

iRecNo = request.QueryString("RecNo")

Response.write "<font color=red>"

sQuery = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
Response.write "<p>"&sQuery
dcrs.open sQuery,con
if not dcrs.eof then
	iInvRecNo = dcrs(0)
end if
dcrs.Close
			
con.beginTrans
set ndRoot = oDOM.createElement("STOCK")
ndRoot.setAttribute "TRANSDATE",dRcvd
ndRoot.setAttribute "FINFROMDATE",sFinFrom
ndRoot.setAttribute "FINTODATE",sFinTo
ndRoot.setAttribute "UNIT",sOrgID
ndRoot.setAttribute "SRCTYPE","RR"
ndRoot.setAttribute "TRANSACTIONTYPE","RR"
ndRoot.setAttribute "REFTYPE",""
ndRoot.setAttribute "RECEIPTFOR",""
ndRoot.setAttribute "RECNO",iRecNo
ndRoot.setAttribute "RECEIVEDON",dRcvd
oDOM.appendChild ndRoot

sQuery = "Select ItemCode,ClassificationCode,SUM(Quantityreturn) Quantity,IsNull(Rate,1) Rate,CreatedBy,"
sQuery = sQuery &" Convert(varchar,CreatedOn,103),IsNull(StockQuality,0) from APP_T_InternalReceiptHeader H join "
sQuery = sQuery &" APP_T_InternalReceiptDetails D on H.InternalReceiptNo = D.InternalReceiptNo where "
sQuery = sQuery &" H.InternalReceiptNo = "& iRecNo &" Group By ItemCode,ClassificationCode,Rate,CreatedBy,CreatedOn,StockQuality"
With dcrs
    .CursorLocation = 3
    .CursorType = 3
    .Source = sQuery
    .ActiveConnection = con
    .Open
End With
if not dcrs.eof then
    iItemEntryNo =0
    do while not dcrs.eof 
        iItemEntryNo = iItemEntryNo + 1
            set ndItem = oDOM.createElement("ITEM")
                ndItem.setAttribute "ITEMENTRYNO",iItemEntryNo
                ndItem.setAttribute "ITEM",dcrs(0)
                ndItem.setAttribute "CLASS",dcrs(1)
                ndItem.setAttribute "ITEMQTY",dcrs(2)
                ndItem.setAttribute "ITEMVALUE",cdbl(dcrs(2))*cdbl(dcrs(3))
                ndItem.setAttribute "ATTRIBUTE",""
                ndItem.setAttribute "SUMQTY",dcrs(2)
                ndItem.setAttribute "CREATEDBY",dcrs(4)
                ndItem.setAttribute "CREATEDON",dcrs(5)
                ndItem.setAttribute "STOCKQUALITY",dcrs(6)
                sQuery = "Select IsNull(Store,0),IsNull(Bin,0) BIN,SUM(QuantityReturn) Quantity,IsNull(Rate,1) Rate from "
                sQuery = sQuery &" APP_T_InternalReceiptDetails where InternalReceiptNo = "& iRecNo &" and ItemCode ="& dcrs(0)
                sQuery = sQuery &" Group By Store,Bin,Rate "
                dcrs1.open sQuery,con
                if not dcrs1.eof then
                    do while not dcrs1.eof 
                        iStore = dcrs1(0)
                        iBin = dcrs1(1)
                        if trim(iBin)="" or IsNull(iBin) then iBin = "0"
                        set ndStorage = oDOM.createElement("STORAGE")
                            ndStorage.setAttribute "STOENTRYNO",iItemEntryNo
                            ndStorage.setAttribute "ITEM",dcrs(0)
                            ndStorage.setAttribute "CLASS",dcrs(1)
                            ndStorage.setAttribute "STORE",iStore
                            ndStorage.setAttribute "BIN",iBin
                            ndStorage.setAttribute "STOREQTY",dcrs1(2)
                            ndStorage.setAttribute "STOREVALUE",cdbl(dcrs1(2))*cdbl(dcrs1(3))
                            ndStorage.setAttribute "DATERECEIVED",dRcvd
                        ndItem.appendChild ndStorage
                        
                        sQuery = "Select LotNo,QuantityReturn,IsNull(Rate,1) Rate,PackingNum,PackingCode,AttributeList,SerialNo,"
                        sQuery = sQuery & " StockQuality,GrossQuantityReturn from APP_T_InternalReceiptDetails where InternalReceiptNo = "& iRecNo &" and ItemCode ="& dcrs(0) &" and Store = "& dcrs1(0) &" and SerialNo is Not null"
                        With dcrs2
                            .CursorLocation = 3
                            .CursorType =  3
                            .Source = sQuery
                            .ActiveConnection = con
                            .Open 
                        End With
                        if not dcrs2.eof then
                            do while not dcrs2.eof
                                sLotNo = dcrs2(0)
                                iQty = dcrs2(1)
                                iRate = dcrs2(2)
                                iPackNum = dcrs2(3)
                                iPackType = dcrs2(4)
                                sAttribute = dcrs2(5)
                                iSerNo = dcrs2(6)
                                iStockQuality=dcrs2(7)
                                iQtyGross = dcrs2(8)
                                if trim(sLotNo)="" or IsNull(sLotNo) then sLotNo = ""
                                if trim(sAttribute)="" or IsNull(sAttribute) then sAttribute=""
                                if trim(iPackNum)="" or IsNull(iPackNum) then iPackNum=""
                                if trim(iPackType)="" or IsNull(iPackType) then iPackType=""
                                if trim(iSerNo)="" or IsNull(iSerNo) then iSerNo = ""
                            
                                set ndLot = oDOM.createElement("LOT")
                                    ndLot.setAttribute "LOTENTRYNO",iItemEntryNo
                                    ndLot.setAttribute "ITEM",dcrs(0)
                                    ndLot.setAttribute "CLASS",dcrs(1)
                                    ndLot.setAttribute "STORE",iStore
                                    ndLot.setAttribute "BIN",iBin
                                    ndLot.setAttribute "LOT",sLotNo
                                    ndLot.setAttribute "QTY",iQty
                                    ndLot.setAttribute "RATE",iRate
                                    ndLot.setAttribute "GROSSQTY",iQtyGross
                                    ndLot.setAttribute "PACKINGNUMBER",iPackNum
                                    ndLot.setAttribute "PACKINGCODE",iPackType
                                    ndLot.setAttribute "SELLINGNUMBER","0"
                                    ndLot.setAttribute "WEIGHTPERSELLINGFORM","0"
                                    ndLot.setAttribute "SELLINGFORM","0"
                                    ndLot.setAttribute "STAGE","0"
                                    ndLot.setAttribute "ATTRIBUTE",sAttribute
                                    ndLot.setAttribute "SERIALNO",iSerNo
                                    ndLot.setAttribute "SQ",iStockQuality
                                ndStorage.appendChild ndLot
                                dcrs2.movenext
                            loop
                        end if
                        dcrs2.close
                        dcrs1.movenext
                    loop
                end if
                dcrs1.close
            ndRoot.appendChild ndItem
        dcrs.movenext
    loop
end if
dcrs.close

oDOM.save(server.mappath("../temp/transaction/IntRecAccount"&session.sessionID&".xml"))
Set adoCmd = Server.CreateObject("ADODB.Command")
    Set adoCmd.ActiveConnection = con

    adoCmd.CommandText = "StockUpdation"
    adoCmd.CommandType = 4
    adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(oDOM.xml),oDOM.xml)
    adoCmd.Execute()

sQuery = "Update APP_T_InternalReceiptHeader set InvRecNo = "& iInvRecNo &" where InternalReceiptNo = "& iRecNo
response.write "<p>"& sQuery
con.execute sQuery

	
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
	con.RollbackTrans
	Response.End 
    Response.clear
	con.CommitTrans
end if

con.close
set con = nothing
Response.redirect "MaterialReceipts.asp?OptType=I"
%>
