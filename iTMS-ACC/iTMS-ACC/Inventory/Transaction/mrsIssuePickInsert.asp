<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	mrsIssuePickInsert.asp
	'Module Name				:	Inventory (MRS Issue Pick Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 27, 2003
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	APRIL 28 2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsIssuePickDetailsEntry.asp
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
<!--#include file="../../include/mrsStatus.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
<%
Dim newxml,subContDOM,dcrs,dcrs1
Dim ndRoot,subContRoot,ndItem,ndPickDet,ndPick,ndSerialHead,ndSerialDet
Dim sTempValues,Values,sOrgID,sQuery,sIssForType
Dim sPOConfirm,sSalInvConfirm,sSelectedInvoice,sGatePassConfirm,sProformaConfirm,sGatePassEntryNo
Dim sSubConProforma,sLotNo,sLocNo,sBinNo
Dim iIssuedBy,iIssueEntryNo,iTotPickedQty,issForCode,iParCode,iCreatedBy,iPickNo
Dim iItemEntNo,iClassCode,iItemCode,iAttID,iSerialNo,iSerialQty,iPickEntryNO,iRate,iValue
Dim iPickQty,iLotQty,iLotNetQty,iInvRecptNo,sTempLotQty,sChkLotQty
Dim dCreatedOn,sFinFrom,sFinTo,sFinPeriod,sArrPeriod,sSelLotNo
Dim bSerialFlag,sMarkPackFlag,sLotOrPackFlag,sScheduleNo,sPickDate
Set newxml = Server.CreateObject("Microsoft.XMLDOM")
Set subContDOM = Server.CreateObject("Microsoft.XMLDOM")

set dcrs = server.createObject("ADODB.Recordset")
set dcrs1 = server.createObject("ADODB.Recordset")

sFinPeriod = session("FinPeriod")
sArrPeriod = split(sFinPeriod,":")
sFinFrom = "01/04/"&sArrPeriod(0)
sFinTo = "31/03/"& sArrPeriod(1)

newxml.async = false
 '   newxml.load(Request)
  '  newxml.save Server.MapPath("../Temp/Transaction/Test.xml")
  
    newxml.load(server.mappath("../temp/transaction/IssuePick_"&Session.sessionID&".xml"))

iIssuedBy = getUserid
sScheduleNo = Request("ScheduleNo")
if trim(sScheduleNo)="" or isNull(sScheduleNo) then sScheduleNo = ""


Set ndRoot = newxml.documentElement
set subContRoot = subContDOM.createElement("ROOT")
subContDOM.appendChild subContRoot 

iIssueEntryNo = trim(ndRoot.getAttribute("ISSUENO"))
sOrgID = trim(ndRoot.getAttribute("UNIT"))
iTotPickedQty = trim(ndRoot.getAttribute("TOT"))

sPOConfirm = trim(ndRoot.getAttribute("POConfirm"))
sSalInvConfirm = trim(ndRoot.getAttribute("SInvConfirm"))
sSelectedInvoice = trim(ndRoot.getAttribute("Invoice"))
sGatePassConfirm = trim(ndRoot.getAttribute("GPConfirm"))
sProformaConfirm = trim(ndRoot.getAttribute("ProConfirm"))
sGatePassEntryNo = 0

sPickDate = trim(ndRoot.getAttribute("PickDate"))



newxml.Save server.MapPath("../temp/transaction/"&iIssueEntryNo&".xml")
	
sQuery = "Select IssuedToCode,IssuedToType,Convert(varchar,IssueDate,103),IssuedBy from Inv_T_MaterialIssueHeader where IssueEntryNo = "& iIssueEntryNo
dcrs.open sQuery,con
if not dcrs.EOF then
	issForCode = trim(dcrs(0))
	sIssForType = trim(dcrs(1))
	dCreatedOn = trim(dcrs(2))
	iCreatedBy = trim(dcrs(3))
end if
dcrs.Close

Response.Write "issForCode = "& issForCode
if trim(issForCode)="SUB" then
	if trim(sProformaConfirm)="Y" then
		sSubConProforma="Y"
	else
		sSubConProforma = "N"
	end if
else
	sSubConProforma="N"
end if

con.beginTrans

sQuery = "Select isNull(Max(PickNumber),0)+1 from INV_T_IssuePick"
dcrs.open sQuery,con
if not dcrs.eof then
    iPickNo = dcrs(0)
end if
dcrs.close
sQuery = "Select MarkPackFlag from INV_T_MaterialIssueHeader where IssueEntryNo = "& iIssueEntryNo
dcrs.open sQuery,con
if not dcrs.eof then
    sMarkPackFlag =trim(dcrs(0))
end if
dcrs.close

sQuery = "Insert into INV_T_IssuePick (PickNumber,OrganisationCode,IssueEntryNo,Pickedon,PickedBy) "
sQuery = sQuery & "Values("& iPickNo &","&pack(sOrgID)&","& iIssueEntryNo &",Convert(datetime,"& pack(sPickDate) &",103),"&iIssuedBy &")"
Response.write "<p>"& sQuery
con.execute sQuery


if ndRoot.hasChildNodes() then
    for each ndItem in ndRoot.childNodes
        if trim(ndItem.nodeName)="ITM" then
            iItemEntNo = ndItem.getAttribute("ItemEntNo")
            iClassCode = ndItem.getAttribute("CLACODE")
            iItemCode = ndItem.getAttribute("ITMCODE")
            iAttID = ndItem.getAttribute("ATTID")
            
            if trim(iAttID)="" or isNull(iAttID) then iAttID = "NULL"
            
            for each ndPickDet in ndItem.childNodes
                if trim(ndPickDet.nodeName)="PickDet" then
                    for each ndPick in ndPickDet.childNodes
                        if trim(ndPick.nodeName)="PICK" then
                            sLocNo = ndPick.getAttribute("LOC")
                            sBinNo =ndPick.getAttribute("BIN")
                            sLotNo =  ndPick.getAttribute("LOTNO")
                            iPickQty =  ndPick.getAttribute("ISSQTY")
                            if trim(sLotNo)="" or isNull(sLotNo) then sLotNo = "NULL"
                            if trim(sLotNo)<>"NULL"  then sLotNo = pack(sLotNo)
                            if ndPick.hasChildNodes() then
                                for each ndSerialHead in ndPick.childNodes
                                    if trim(ndSerialHead.nodeName)="SERIALHEADER" then
                                        for each ndSerialDet in ndSerialHead.childNodes
                                            if trim(ndSerialDet.nodeName)="SERIALDETAILS" then
                                                iSerialNo=ndSerialDet.getAttribute("SERIALNO")
                                                iSerialQty = ndSerialDet.getAttribute("QTY")
                                                sSelLotNo =  ndSerialDet.getAttribute("LOTNO")
                                                if trim(sSelLotNo)="" or isNull(sSelLotNo) then sSelLotNo = "NULL"
                                                if trim(sSelLotNo)<>"NULL"  then sSelLotNo = pack(sSelLotNo)
                                                
                                                if cdbl(iSerialQty)>0 then
                                                    sQuery = "Select isNull(max(EntryNumber),0)+ 1 from INV_T_IssuePickDetails "
                                                    dcrs.open sQuery,con
                                                    if not dcrs.eof then
                                                        iPickEntryNO = dcrs(0)
                                                    end if
                                                    dcrs.close
                                                    
                                                    sQuery = "Insert into INV_T_IssuePickDetails (PickNumber,EntryNumber,ClassificationCode,ItemCode,LotNumber,"
                                                    sQuery = sQuery &" SerialNo,QuantityPicked,ItemAttributes)"
                                                    sQuery = sQuery &" Values("& iPickNo &","&iPickEntryNO&","&iClassCode&","& iItemCode &","& sSelLotNo &","
                                                    sQuery = sQuery &" "& iSerialNo &","& iSerialQty &","&iATTID&")"
                                                    Response.write "<p>"&sQuery
                                                    con.execute sQuery
                                                    
                                                    sQuery = "Select * from INV_T_MaterialIssuedForPick where IssueEntryNo = "& iIssueEntryNo & " and ItemCode =  "& iItemCode
                                                    sQuery = sQuery & "  and ClassificationCode = "& iClassCode &" and LocationNumber = "& sLocNo &" and BinNumber = "& sBinNo &" and SerialNo ="& iSerialNo
                                                    
                                                    if trim(sLotNo)<>"NULL" then
                                                        sQuery = sQuery & " and LotNo = "& sLotNo
                                                    end if
                                                    
                                                    if trim(iATTID)<>"" and trim(iAttID)<>"NULL" then
                                                        sQuery = sQuery & " and ItemAttributes = "& iATTID
                                                    end if
                                                    Response.write "<p>"& sQuery
                                                    dcrs.open sQuery,con
                                                    if not dcrs.eof then
                                                        bSerialFlag =  true
                                                    else
                                                        bSerialFlag =false
                                                    end if 'if not dcrs.eof then
                                                    dcrs.close
                                                    
                                                    if bSerialFlag then
                                                        sQuery = "Update INV_T_MaterialIssuedForPick set QuantityPicked = IsNull(QuantityPicked,0) + "& iSerialQty &" where IssueEntryNo = "& iIssueEntryNo
                                                        sQuery = sQuery &" and ItemCode = "& iItemCode &" and ClassificationCode = "& iClassCode &" and LocationNumber = "& sLocNo &" and BinNumber = "& sBinNo
                                                        sQuery = sQuery &" and SerialNo ="&iSerialNo  &" and ItemEntryNo = "& iItemEntNo
                                                        if trim(sLotNo)<>"NULL" then
                                                            sQuery = sQuery & " and LotNo = "& sLotNo
                                                        end if
                                                        
                                                        if trim(iATTID)<>"" and trim(iAttID)<>"NULL" then
                                                            sQuery = sQuery & " and ItemAttributes = "& iATTID
                                                        end if
                                                        Response.write "<p>"& sQuery
                                                        con.execute sQuery
                                                        
                                                        if trim(sScheduleNo)<>"" then
                                                            sQuery = "Update Inv_T_IssueForPickSchedule set PickedQty = IsNull(PickedQty,0) + "& iSerialQty  &" where ScheduleNo ="& sScheduleNo &" and IssueEntryNo ="& iIssueEntryNo
                                                            Response.write "<p>"& sQuery
                                                            con.execute sQuery
                                                        end if
                                                        
                                                        sQuery = "Select FlagLotOrPack from INV_T_MaterialIssuedForPick where IssueEntryNo = "& iIssueEntryNo & " and ItemCode =  "& iItemCode
                                                        sQuery = sQuery & "  and ClassificationCode = "& iClassCode &" and LocationNumber = "& sLocNo &" and BinNumber = "& sBinNo &" and SerialNo ="& iSerialNo &" and ItemEntryNo = "& iItemEntNo
                                                        if trim(sLotNo)<>"NULL" then
                                                            sQuery = sQuery & " and LotNo = "& sLotNo
                                                        end if
                                                        
                                                        if trim(iATTID)<>"" and trim(iAttID)<>"NULL" then
                                                            sQuery = sQuery & " and ItemAttributes = "& iATTID
                                                        end if
                                                        Response.write "<p>"& sQuery
                                                        dcrs.open sQuery,con
                                                        if not dcrs.eof then
                                                            sLotOrPackFlag = trim(dcrs(0))
                                                        end if
                                                        dcrs.close
                                                        
                                                    else
                                                        sQuery = "Update INV_T_MaterialIssuedForPick set QuantityPicked = IsNull(QuantityPicked,0) + "& iSerialQty &" where IssueEntryNo = "& iIssueEntryNo
                                                        sQuery = sQuery &" and ItemCode = "& iItemCode &" and ClassificationCode = "& iClassCode &" and LocationNumber = "& sLocNo &" and BinNumber = "& sBinNo 
                                                        sQuery = sQuery &" and ItemEntryNo = "& iItemEntNo
                                                        if trim(sLotNo)<>"NULL" then
                                                            sQuery = sQuery & " and LotNo = "& sLotNo
                                                        end if
                                                        if trim(iATTID)<>"" and trim(iAttID)<>"NULL" then
                                                            sQuery = sQuery & " and ItemAttributes = "& iATTID
                                                        end if
                                                        Response.write "<p>"& sQuery
                                                        con.execute sQuery
                                                        
                                                        if trim(sScheduleNo)<>"" then
                                                            sQuery = "Update Inv_T_IssueForPickSchedule set PickedQty = IsNull(PickedQty,0) + "& iSerialQty  &" where ScheduleNo ="& sScheduleNo &" and IssueEntryNo ="& iIssueEntryNo
                                                            Response.write "<p>"& sQuery
                                                            con.execute sQuery
                                                        end if 
                                                        
                                                        
                                                        sQuery = "Select FlagLotOrPack from INV_T_MaterialIssuedForPick where IssueEntryNo = "& iIssueEntryNo & " and ItemCode =  "& iItemCode
                                                        sQuery = sQuery & "  and ClassificationCode = "& iClassCode &" and LocationNumber = "& sLocNo &" and BinNumber = "& sBinNo &" and ItemEntryNo = "& iItemEntNo
                                                        if trim(sLotNo)<>"NULL" then
                                                            sQuery = sQuery & " and LotNo = "& sLotNo
                                                        end if
                                                        
                                                        if trim(iATTID)<>"" and trim(iAttID)<>"NULL" then
                                                            sQuery = sQuery & " and ItemAttributes = "& iATTID
                                                        end if
                                                        Response.write "<p>"& sQuery
                                                        dcrs.open sQuery,con
                                                        if not dcrs.eof then
                                                            sLotOrPackFlag = trim(dcrs(0))
                                                        end if
                                                        dcrs.close
                                                        
                                                    end if
                                                    
                                                   ' if trim(sLotNo)="NULL" then 
                                                   '     sQuery = "Select isNull(LotNumber,'') from INV_T_LocationLot where SerialNumber="& iSerialNo
                                                   '     dcrs.open sQuery,con
                                                   '     if not dcrs.eof then
                                                   '        sLotNo = trim(dcrs(0))
                                                   '        if trim(sLotNo)="" or isNull(sLotNo) then sLotNo = "NULL"
                                                   '        if trim(sLotNo)<>"NULL" then sLotNo = pack(sLotNo)
                                                   '     end if
                                                   '     dcrs.close
                                                   ' end if
                                                    
                                                    sQuery = "SELECT ISNULL(RATE,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE SERIALNUMBER = " & iSerialNo & ""
                                                    dcrs.open sQuery,con
                                                    if not dcrs.EOF then
												        iRate = cdbl(dcrs(0))
												    else
												        iRate = 0
											        end if
											        dcrs.Close
											        iValue = cdbl(iSerialQty )*cdbl(iRate)
                                                    
                                                    sQuery = "Insert into INV_T_MaterialIssueDetails (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,ItemEntryNo,"
                                                    sQuery = sQuery & "ItemAttributes, LotNo,SerialNo,LocationNumber,BinNumber,QuantityIssued)"
                                                    sQuery = sQuery & " Values("& iIssueEntryNo &","&pack(sOrgID)&","& iClassCode &","& iItemCode &","& iItemEntNo &","
                                                    sQuery = sQuery & ""& iATTID &","& sSelLotNo &","& iSerialNo &","& sLocNo &","& sBinNo &","& iSerialQty &")"
                                                    Response.write "<p>"&sQuery
                                                    con.execute sQuery
                                                    
                                                    if trim(sMarkPackFlag)="L" or trim(sLotOrPackFlag)="L" then
                                                        sQuery = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iSerialQty & ") "
											            sQuery = sQuery & "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE =" & Pack(sOrgID) & " AND "
											            sQuery = sQuery & "SERIALNUMBER = " & iSerialNo & " AND STORAGELOCATIONNO = " & sLocNo & " AND (ISNULL(STORAGEBINNUMBER,NULL) = " & sBinNo & " OR "
											            sQuery = sQuery & "ISNULL(STORAGEBINNUMBER,0) = " & sBinNo & ") "
							            	            Response.Write "<p>"&sQuery
										                con.Execute sQuery
                                                    end if 'if trim(sMarkPackFlag)="L" then
                                                    
                                                    sQuery = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) 
												    sQuery = sQuery & " AND LOCATIONNUMBER = " & sLocNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " 
												    sQuery = sQuery & " CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
												    Response.write "<p>"&sQuery
												    with dcrs
												        .CursorLocation = 3
												        .CursorType = 3
												        .ActiveConnection = con
												        .Source = sQuery
												        .Open
												    end with
												    if dcrs.EOF then
											            sQuery = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
											            "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE) VALUES " &_
													            "(" & Pack(sOrgID) & "," & iClassCode & "," & iItemCode & "," &_
													            "" & sLocNo & "," & sBinNo & "," & iSerialQty & ","& iValue &")"
													    Response.write "<p>"&sQuery
												        con.Execute sQuery
											        else
												        sQuery = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRESERVED = (YEARRESERVED - " & iSerialQty & ") WHERE " &_
													        "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
												            "LOCATIONNUMBER = " & sLocNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
													        "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
													        "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
													    Response.write "<p>"&sQuery
												        con.Execute sQuery
											        end if
											        dcrs.Close
    											
											        sQuery = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRESERVED = (YEARRESERVED - " & iSerialQty & ") " &_
												        "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
												        "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
												        "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
												        "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
											        Response.write "<p>"&sQuery
											        con.Execute sQuery
                                                    
                                                end if'if cdbl(iSerialQty)>0 then 
                                                
                                            end if 'if trim(ndSerialDet.nodeName)="SERIALDETATILES" then
                                        next
                                    end if 'if trim(ndSerialHead.nodeName)="SERIALHEADER" then
                                next
                            else 'if not ndPick.hasChildNodes() then
                                Response.write "<p>Welcome to none case"
                                
                                sQuery = "Select isNull(max(EntryNumber),0)+ 1 from INV_T_IssuePickDetails "
                                dcrs.open sQuery,con
                                if not dcrs.eof then
                                    iPickEntryNO = dcrs(0)
                                end if
                                dcrs.close
                                
                                sQuery = "Insert into INV_T_IssuePickDetails (PickNumber,EntryNumber,ClassificationCode,ItemCode,LotNumber,"
                                sQuery = sQuery &" SerialNo,QuantityPicked,ItemAttributes)"
                                sQuery = sQuery &" Values("& iPickNo &","&iPickEntryNO&","&iClassCode&","& iItemCode &","& sLotNo &",NULL,"& iPickQty &",NULL)"
                                Response.write "<p>"&sQuery
                                con.execute sQuery
                                
                                sQuery = "Update INV_T_MaterialIssuedForPick set QuantityPicked = IsNull(QuantityPicked,0) + "& iPickQty &" where IssueEntryNo = "& iIssueEntryNo
                                sQuery = sQuery &" and ItemCode = "& iItemCode &" and ClassificationCode = "& iClassCode &" and LocationNumber = "& sLocNo &" and BinNumber = "& sBinNo &" and ItemEntryNo = "& iItemEntNo
                                Response.write "<p>"& sQuery
                                con.execute sQuery
                                
                                if trim(sScheduleNo)<>"" then
                                    sQuery = "Update Inv_T_IssueForPickSchedule set PickedQty = IsNull(PickedQty,0) + "& iPickQty  &" where ScheduleNo ="& sScheduleNo &" and IssueEntryNo ="& iIssueEntryNo
                                    Response.write "<p>"& sQuery
                                    con.execute sQuery
                                end if
                                
                                
                                sQuery = "Select FlagLotOrPack from INV_T_MaterialIssuedForPick where IssueEntryNo = "& iIssueEntryNo & " and ItemCode =  "& iItemCode
                                sQuery = sQuery & "  and ClassificationCode = "& iClassCode &" and LocationNumber = "& sLocNo &" and BinNumber = "& sBinNo &" and ItemEntryNo = "& iItemEntNo
                                if trim(sLotNo)<>"NULL" then
                                    sQuery = sQuery & " and LotNo = "& sLotNo
                                end if
                                
                                if trim(iATTID)<>"" and trim(iAttID)<>"NULL" then
                                    sQuery = sQuery & " and ItemAttributes = "& iATTID
                                end if
                                Response.write "<p>"& sQuery
                                dcrs.open sQuery,con
                                if not dcrs.eof then
                                    sLotOrPackFlag = trim(dcrs(0))
                                end if
                                dcrs.close
                                
                                 sQuery = "Insert into INV_T_MaterialIssueDetails (IssueEntryNo,OrganisationCode,ClassificationCode,ItemCode,ItemEntryNo,"
                                sQuery = sQuery & "ItemAttributes, LotNo,SerialNo,LocationNumber,BinNumber,QuantityIssued)"
                                sQuery = sQuery & " Values("& iIssueEntryNo &","&pack(sOrgID)&","& iClassCode &","& iItemCode &","& iItemEntNo &","
                                sQuery = sQuery & ""& iATTID &","& sLotNo &",NULL,"& sLocNo &","& sBinNo &","& iPickQty &")"
                                Response.write "<p>"&sQuery
                                con.execute sQuery
                                
                                
                                if trim(sMarkPackFlag)="L" or trim(sLotOrPackFlag)="L" then
                                sChkLotQty = iPickQty
                                    sQuery = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "
                                    sQuery = sQuery & " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClassCode & " and OrganisationCode = " & pack(sOrgID) 
                                    sQuery = sQuery & " and StorageLocationNo = "& sLocNo &" and (StorageBinNUmber = "& sBinNo &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO"
                                    Response.write "<p>"& sQuery
                                    With dcrs
                                        .CursorLocation = 3
                                        .CursorType = 3
                                        .ActiveConnection = Con
                                        .Source = sQuery
                                        .Open
                                    End With
                                    if not dcrs.eof then
                                        do while not dcrs.eof 
                                            iLotQty = dcrs(0)
                                                if cdbl(iLotQty)>0 then
                                                    sQuery = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode
                                                    sQuery = sQuery &" and StorageLocationNo = "& sLocNo &" and (StorageBinNUmber = "& sBinNo &" or StorageBinNumber is Null) Order by InventoryReceiptNo"
                                                    Response.write "<p>"& sQuery
                                                    With dcrs1
                                                        .CursorLocation = 3
                                                        .CursorType = 3
                                                        .ActiveConnection = Con
                                                        .Source = sQuery
                                                        .Open
                                                    End With
                                                    if not dcrs1.eof then
                                                        do while not dcrs1.eof
                                                            iLotNetQty = dcrs1(0)
                                                            iInvRecptNo = dcrs1(1)
                                                            if cdbl(sChkLotQty)<> 0 then
                                                                if cint(iLotNetQty)<cint(sChkLotQty) then
                                                                    sTempLotQty = iPickQty
                                                                    
                                                                    sQuery = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iLotNetQty &")"
															        sQuery = sQuery & " WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "
															        sQuery = sQuery & " INVENTORYRECEIPTNO = "& iInvRecptNo & " and StorageLocationNo = "& sLocNo &" and (StorageBinNumber = "& sBinNo &" or StorageBinNumber is Null) "
															        Response.write "<p>"&sQuery
															        con.execute sQuery
															        sChkLotQty = cint(sChkLotQty) - cint(iLotNetQty)
                                                                else
                                                                    sQuery =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& sChkLotQty &")"
															        sQuery = sQuery & " WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "
															        sQuery = sQuery & " INVENTORYRECEIPTNO = "& iInvRecptNo & " and StorageLocationNo = "& sLocNo &" and (StorageBinNUmber = "& sBinNo &" or StorageBinNumber is Null) "    
															        Response.write "<p>"&sQuery
															        con.execute sQuery
															        sChkLotQty = 0
                                                                end if
                                                                sTempLotQty = cint(sTempLotQty) - cint(iPickQty)
                                                            end if 'if cdbl(sChkLotQty)<> 0 then
                                                            dcrs1.movenext
                                                        loop    
                                                    end if
                                                    dcrs1.close
                                                end if
                                            dcrs.movenext
                                        loop
                                    end if
                                    dcrs.close
                                end if 'if trim(sMarkPackFlag)="L" then
                                
                               
                                
                                sQuery = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) 
						        sQuery = sQuery & " AND LOCATIONNUMBER = " & sLocNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " 
						        sQuery = sQuery & " CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
						        Response.write "<p>"&sQuery
						        with dcrs
						            .CursorLocation = 3
						            .CursorType = 3
						            .ActiveConnection = con
						            .Source = sQuery
						            .Open
						        end with
						        if dcrs.EOF then
					                sQuery = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
					                "LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE) VALUES " &_
							                "(" & Pack(sOrgID) & "," & iClassCode & "," & iItemCode & "," &_
							                "" & sLocNo & "," & sBinNo & "," & iPickQty & ","& iValue &")"
							        Response.write "<p>"&sQuery
						            con.Execute sQuery
					            else
						            sQuery = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRESERVED = (YEARRESERVED - " & iPickQty & ") WHERE " &_
							            "ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						                "LOCATIONNUMBER = " & sLocNo & " AND (BINNUMBER = " & sBinNo & " OR BINNUMBER IS NULL) AND " &_
							            "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
							            "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
							        Response.write "<p>"&sQuery
						            con.Execute sQuery
					            end if
					            dcrs.Close
							
					            sQuery = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRESERVED = (YEARRESERVED - " & iPickQty & ") " &_
						            "WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
						            "ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
						            "CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						            "CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					            Response.write "<p>"&sQuery
					            con.Execute sQuery
                                
                            end if'if ndPick.hasChildNodes() then
                        end if'if trim(ndPick.nodeName)="PICK" then
                    next
                end if 'if trim(ndPickDet.nodeName)="PickDet" then
            next
        end if 'if trim(ndItem.nodeName)="ITM" then
    next
end if 'if ndRoot.hasChildNodes() then

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	Response.End 
    Response.Clear 
    con.CommitTrans
    
subContDOM.save(Server.MapPath("/Purchase/temp/transaction/PO_PUR_"&Session.SessionID&".xml"))	

Response.redirect "MRSISSUEPICKLIST.ASP"
''Dont Delete or Block this Below Print It is the Agrgument form Direct Invoice Creation
	''Added by Ragav April 28,2010
	if trim(issForCode)="SUB" then
		Response.Write "@ForSUB:"&sInvNo&":"&iIssueEntryNo
	elseif trim(issForCode)="DIS" then
		Response.Write "@ForInvNo:"&sInvNo&":"&iIssueEntryNo
	elseif trim(issForCode)="SER" then
		Response.Write "@ForGPNo:"&sGatePassNo 
	end if
'Response.End 
end if

con.close
set con = nothing
%>

