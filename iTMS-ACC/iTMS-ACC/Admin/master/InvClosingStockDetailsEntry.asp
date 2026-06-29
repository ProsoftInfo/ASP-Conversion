<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	InvClosingStockDetailsEntry.asp
	'Module Name				:	Inventory Closing Stock
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 13, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	InvClosingStockDetailsInsert.asp
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
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
'XML DOM Variables
dim YrStockXML,YrLocStockXML,RootNode,ItemNode,xmlhttp,sQuery,sResult,sSql,YrLocLotXML
Dim ndRootLot,ndItemLot

dim dcrs,dcrs1,sUnit,sUnitName,dPreFinStartDate,dPreFinEndDate,dCurFinStartDate,dCurFinEndDate,rsTemp
dim iItemCode,iClassCode,iLoc,iBin,iYrOpStock,iYrOpValue,iYrClStock,iYrClValue, iNoOfPacks,iYrRQty,iYrRVal,iYrIQty,iYrIVal
dim iRecQty,iRecValue,iIssQty,iIssValue,iClStock,iClValue, sItemTypeID,iCreadBy,sArrCategory,sCategoryCode,sSubCategory,sClassification
Dim iItemEntryNo,iLocNo,iBinNo,sLotNumber,iSerialNumber,iInventoryReceiptNo,iSerNumber,iLotNett
Dim iSumOfNettQty,iYrStClQty
Dim sRcptNumber
Response.Write "<font color=#000000>"
'Response.Write Request.QueryString

sUnit = trim(Request("UnitCode"))
sUnitName = trim(Request("UnitName"))
dCurFinStartDate = trim(Request("CurrFromDate"))
dCurFinEndDate = trim(Request("CurrToDate"))
dPreFinStartDate = trim(Request("PrevFromDate"))
dPreFinEndDate = trim(Request("PrevToDate"))
sArrCategory = split(trim(Request("CategoryCodes")),":")
sCategoryCode = sArrCategory(0)
sSubCategory = sArrCategory(1)
sClassification =  sArrCategory(2)

iCreadBy = Session("userid")
'Response.Write "<br> sUnit = "& sunit
'Response.Write "<br> sUnitName = "& sUnitName
'Response.Write "<br> dCurFinStartDate = "& dCurFinStartDate
'Response.Write "<br> dCurFinEndDate = "&dCurFinEndDate
'Response.Write "<br> dPreFinStartDate = "&dPreFinStartDate
'Response.Write "<br> dPreFinEndDate "& dPreFinEndDate
'Response.Write "<br> sItemTypeID "& sItemTypeID


'dPreFinStartDate = "01/04/2004"
'dPreFinEndDate = "31/03/2005"

'iPrePeriodFrom = right(dPreFinStartDate,4)&mid(dPreFinStartDate,4,2)
'iPrePeriodTo = right(dPreFinEndDate,4)&mid(dPreFinEndDate,4,2)

' Create our DOM Document Objects
Set YrStockXML = Server.CreateObject("Microsoft.XMLDOM")
Set YrLocStockXML = Server.CreateObject("Microsoft.XMLDOM")
Set YrLocLotXML = Server.CreateObject("Microsoft.XMLDOM")

set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set rsTemp = Server.CreateObject("ADODB.Recordset")

YrStockXML.async=false
YrLocStockXML.async=false
YrLocLotXML.async = false

con.beginTrans

sQuery = "SELECT ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARISSUEQUANTITY,YEARISSUEVALUE FROM INV_T_ITEMYEARLYSTOCK WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " and Itemcode In ("

if trim(sClassification)<>"" and (not IsNull(sClassification)) then
    sQuery = sQuery & " Select ItemCode from VWItemClassCatForYearEndClosing where GroupCode = "& sClassification &" and CategoryCode = '" & sCategoryCode &"'"
elseif Trim(sSubCategory)<>"" and (not IsNull(sSubCategory)) then
    sQuery = sQuery & " Select ItemCode from VWItemClassCatForYearEndClosing where GroupCode<>ParentGroup and ParentGroup = "& sSubCategory &" and CategoryCode = '" & sCategoryCode &"'"
elseif Trim(sCategoryCode)<>"" and (not IsNull(sCategoryCode)) then
    sQuery = sQuery & " Select ItemCode from VWItemClassCatForYearEndClosing where CategoryCode = '" & sCategoryCode &"'"
end if

sQuery = sQuery  & ") ORDER BY 1,2,3 FOR XML AUTO"
Response.write "<p>"&Squery
'Response.end
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

do while not dcrs.EOF
	sResult = sResult & dcrs(0)
dcrs.MoveNext
loop
dcrs.Close

sResult = "<ROOT>" & sResult & "</ROOT>"

YrStockXML.loadXML sResult
set RootNode = YrStockXML.documentElement

for each ItemNode in RootNode.ChildNodes
	 iClassCode = trim(ItemNode.attributes.getNamedItem("CLASSIFICATIONCODE").value)
	 iItemCode = trim(ItemNode.attributes.getNamedItem("ITEMCODE").value)
	 iYrOpStock = trim(ItemNode.attributes.getNamedItem("YEAROPENINGSTOCK").value)
	 iYrOpValue = trim(ItemNode.attributes.getNamedItem("YEAROPENINGVALUE").value)
	 iYrRQty= trim(ItemNode.attributes.getNamedItem("YEARRECEIPTQUANTITY").value)
	 iYrRVal= trim(ItemNode.attributes.getNamedItem("YEARRECEIPTVALUE").value)
	 iYrIQty= trim(ItemNode.attributes.getNamedItem("YEARISSUEQUANTITY").value)
	 iYrIVal= trim(ItemNode.attributes.getNamedItem("YEARISSUEVALUE").value)
	 iYrClStock = Trim(ItemNode.attributes.getNamedItem("YEARCLOSINGSTOCK").value)
	 iYrClValue = Trim(ItemNode.attributes.getNamedItem("YEARCLOSINGSTOCK").value)
	 
'	 iYrClStock = cdbl(iYrOpStock)+cdbl(iYrRQty)-cdbl(iYrIQty)
'	 iYrClValue = cdbl(iYrOpValue)+CDbl(iYrRVal)-CDbl(iYrIVal)
'	 
'	 if cdbl(iYrClValue)<0 then
'	    iYrClValue = cdbl(iYrClValue)*-1
'	 end if
'	 
'	 sSql = "Update INV_T_ITEMYEARLYSTOCK set YEARCLOSINGSTOCK="& iYrClStock &",YEARCLOSINGVALUE="& iYrClValue  &"  WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " and Itemcode In ("& iItemCode &")"
'	 Response.Write "<p>"& sSql
'	 con.execute sSql
	 
	 sSql = "Select ReceiptNumbering from Inv_M_ItemMaster where ItemCode = "& iItemCode 
	 With rsTemp
	    .CursorLocation = 3
	    .CursorType =3
	    .Source = sSql
	    .ActiveConnection = con
	    .Open 
	 end with
	 set rsTemp.ActiveConnection =  nothing
	 if not rsTemp.eof then
	    sRcptNumber = rsTemp(0)
	 end if
	 rsTemp.Close 

	sQuery = "SELECT ORGANISATIONCODE,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARISSUEQUANTITY,YEARISSUEVALUE,YEAROPENINGSTOCK,YEAROPENINGVALUE FROM INV_T_ITEMYEARLYSTOCK WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dCurFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dCurFinEndDate) & ""
	Response.write "<p>"&sQuery
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iRecQty = trim(dcrs(5))
		iRecValue = trim(dcrs(6))
		iIssQty = trim(dcrs(7))
		iIssValue = trim(dcrs(8))

		iClStock = (cdbl(iYrClStock) + cdbl(iRecQty)) - cdbl(iIssQty)
		iClValue = (cdbl(iYrClValue) + cdbl(iRecValue)) - cdbl(iIssValue)

		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEAROPENINGSTOCK = " & iYrClStock & "," &_
			"YEAROPENINGVALUE = " & iYrClValue & ",YEARCLOSINGSTOCK = " & iClStock & ", " &_
			"YEARCLOSINGVALUE = " & iClValue & " WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
			"ORGANISATIONCODE = " & Pack(sUnit) & " AND " &_
			"CONVERT(DATETIME," & Pack(dCurFinStartDate) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(dCurFinEndDate) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write "<p>"&sSql & "<BR>"
		con.Execute sSql

		sSql = "Select Count(1) from Inv_T_LocationLot where OrganisationCode = " & Pack(sUnit) & " and classificationcode = " & iClassCode & " and ItemCode = " & iItemCode & " and SerialNumber IS NOT NULL and (LotQuantityNett - QuantityIssued) > 0"
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					iNoOfPacks = dcrs1(0)
				Else
					iNoOfPacks = 0
				End If

		dcrs1.Close

		sSql = "Update INV_T_ItemLedger Set TransactQuantity = (TransactQuantity + " & iYrClStock & "), TransactValue = (TransactValue + " & iYrClValue & "), NoOfPacks = (NoOfPacks + " & iNoOfPacks & ") where OrganisationCode = " & Pack(sUnit) & " and ClassificationCode = " & iClassCode & " and ItemCode = " & iItemCode & " and TransactionType = 'RO' and  Convert(datetime,TransactionDate,103) = Convert(datetime," & Pack(dCurFinStartDate) & ",103)"
		Response.Write "<P>"&sSql & "<BR>"
		con.Execute sSql

	else
		sSql = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
			"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
			"(" & Pack(sUnit) & "," & iClassCode & "," & iItemCode & "," &_
			"CONVERT(DATETIME," & Pack(dCurFinStartDate) & ",103),CONVERT(DATETIME," & Pack(dCurFinEndDate) & ",103)," &_
			"" & iYrClStock & "," & iYrClValue & "," & iYrClStock & "," & iYrClValue & ")"
		Response.Write "<P>"&sSql & "<BR>"
		con.Execute sSql

		sSql = "Select Count(1) from Inv_T_LocationLot where OrganisationCode = " & Pack(sUnit) & " and classificationcode = " & iClassCode & " and ItemCode = " & iItemCode & " and SerialNumber IS NOT NULL and (LotQuantityNett - QuantityIssued) > 0"
		Response.write "<p>"&sSql
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		if not dcrs1.EOF then
			iNoOfPacks = dcrs1(0)
		Else
			iNoOfPacks = 0
		End If

		dcrs1.Close
		sSql = "Insert Into INV_T_ItemLedger (OrganisationCode,ItemCode, ClassificationCode, TransactionType, TransactionDate, TransactQuantity, TransactValue, SentToAccounts,NoOfPacks) Values(" & Pack(sUnit) & "," & iItemCode & "," & iClassCode & ",'RO',Convert(datetime," & Pack(dCurFinStartDate) & ",103)," & iYrClStock & "," & iYrClValue & ",'T'," & iNoOfPacks & ")"
		Response.write "<p>"&sSql
		con.Execute sSql
		
		if Trim(sRcptNumber)<>"N" then
		    sResult = ""
		    sSql = "Select isNull(InventoryReceiptNo,0) InventoryReceiptNo,isNull(ItemEntryNo,0) ItemEntryNo,IsNull(StorageLocationNo,0) StorageLocationNo,IsNull(StorageBinNumber,0) StorageBinNumber,IsNull(LotNumber,'') LotNumber,IsNull(SerialNumber,0) SerialNumber,IsNull(LotQuantityNett,0) LotQuantityNett  from Inv_T_LocationLot where OrganisationCode = " & Pack(sUnit) & " and classificationcode = " & iClassCode & " and ItemCode = " & iItemCode & " and (LotQuantityNett - QuantityIssued) > 0 ORDER BY 1 FOR XML AUTO "
		    Response.write "<p>"&sSql
		    with rsTemp
		        .CursorLocation = 3
		        .CursorType =3
		        .Source = sSql
		        .ActiveConnection = con
		        .Open 
		    end with 
		    set rsTemp.ActiveConnection =  nothing
		    if not rsTemp.EOF then
		        do while not rsTemp.EOF 
		        sResult = sResult & rsTemp(0)
		        rsTemp.MoveNext 
		        loop
		    end if
		    rsTemp.Close 
		    sResult = "<ROOT>"& sResult &"</ROOT>"
		    YrLocLotXML.loadXML sResult 
		    YrLocLotXML.save server.MapPath("../Temp/LocationLot_"&iItemCode&".xml")
    		
		    set ndRootLot = YrLocLotXML.documentElement
		    if ndRootLot.hasChildNodes then
		        for each ndItemLot in ndRootLot.childNodes
		            iInventoryReceiptNo = ndItemLot.getAttribute("InventoryReceiptNo")
		            iItemEntryNo = ndItemLot.getAttribute("ItemEntryNo")
		            iLocNo = ndItemLot.getAttribute("StorageLocationNo")
		            iBinNo = ndItemLot.getAttribute("StorageBinNumber")
		            sLotNumber = ndItemLot.getAttribute("LotNumber")
	                iSerNumber  = ndItemLot.getAttribute("SerialNumber")
	                iLotNett = ndItemLot.getAttribute("LotQuantityNett")
    		    
		            if Trim(iItemEntryNo)="0" or IsNull(iItemEntryNo) then iItemEntryNo = "NULL"
		            if Trim(iLocNo)="0" or IsNull(iLocNo) then iLocNo = "NULL"
		            if Trim(iBinNo)="0" or IsNull(iBinNo) then iBinNo = "NULL"
		            if Trim(sLotNumber)="" or IsNull(sLotNumber) then sLotNumber = "NULL"
		            if Trim(sLotNumber)<>"NULL" then sLotNumber = Pack(sLotNumber)
		            if Trim(iSerNumber)="0" or IsNull(iSerNumber) then iSerNumber ="NULL"
		            if Trim(iLotNett)="0" or IsNull(iLotNett) then iLotNett = "NULL"
        		        
		            sSql = " Insert into INV_T_ItemLedgerOpeningPacks (OrganisationCode,ItemCode,ClassificationCode,"&_
                            "TransactionType,TransactionDate,InventoryReceiptNo,ItemEntryNo,StorageLocationNo,StorageBinNumber,"&_
                            "LotNumber,SerialNumber,LotQuantityNett) values("& pack(sUnit) &","& iItemCode &","& iClassCode &","&_
                            "'RO',Convert(datetime,"&Pack(dCurFinStartDate)&",103),"& iInventoryReceiptNo &","& iItemEntryNo &","& iLocNo &","&_
                            " "& iBinNo &","& sLotNumber &","& iSerNumber &","& iLotNett&")"
                            Response.Write "<p>"& sSql
                            con.Execute sSql
    		        
		        next
		    end if 'if ndRootLot.hasChildNodes then
		end if 'if Trim(sRcptNumber)<>"N" then
    end if
	dcrs.Close
next
YrStockXML.save server.MapPath("../Temp/StockData.xml")

sResult = ""
'sQuery = "SELECT ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE,RTRIM(CONVERT(CHAR,FINANCIALYEARFROM,103)) AS FINANCIALYEARFROM,RTRIM(CONVERT(CHAR,FINANCIALYEARTO,103)) AS FINANCIALYEARTO,LOCATIONNUMBER,ISNULL(BINNUMBER,0) AS BINNUMBER,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARISSUEQUANTITY,YEARISSUEVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " ORDER BY 1,2,3,4,5 FOR XML AUTO"
sQuery = "SELECT ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE,LOCATIONNUMBER,ISNULL(BINNUMBER,0) AS BINNUMBER,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARISSUEQUANTITY,YEARISSUEVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " and Itemcode In("

if trim(sClassification)<>"" and (not IsNull(sClassification)) then
    sQuery = sQuery & " Select ItemCode from VWItemClassCatForYearEndClosing where GroupCode = "& sClassification &" and CategoryCode = '" & sCategoryCode &"'"
elseif Trim(sSubCategory)<>"" and (not IsNull(sSubCategory)) then
    sQuery = sQuery & " Select ItemCode from VWItemClassCatForYearEndClosing where GroupCode<>ParentGroup and ParentGroup = "& sSubCategory &" and CategoryCode = '" & sCategoryCode &"'"
elseif Trim(sCategoryCode)<>"" and (not IsNull(sCategoryCode)) then
    sQuery = sQuery & " Select ItemCode from VWItemClassCatForYearEndClosing where CategoryCode = '" & sCategoryCode &"'"
end if

sQuery = sQuery & ")  ORDER BY 1,2,3,4,5 FOR XML AUTO"

Response.write "<p>"&sQuery
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
'dcrs.Open sQuery,con
do while not dcrs.EOF
	sResult = sResult & dcrs(0)
dcrs.MoveNext
loop
dcrs.Close

sResult = "<ROOT>" & sResult & "</ROOT>"

YrLocStockXML.loadXML sResult
set RootNode = YrLocStockXML.documentElement

for each ItemNode in RootNode.ChildNodes
	 iClassCode = trim(ItemNode.attributes.getNamedItem("CLASSIFICATIONCODE").value)
	 iItemCode = trim(ItemNode.attributes.getNamedItem("ITEMCODE").value)
	 iLoc = trim(ItemNode.attributes.getNamedItem("LOCATIONNUMBER").value)
	 iBin = trim(ItemNode.attributes.getNamedItem("BINNUMBER").value)
	 iYrOpStock = trim(ItemNode.attributes.getNamedItem("YEAROPENINGSTOCK").value)
	 iYrOpValue = trim(ItemNode.attributes.getNamedItem("YEAROPENINGVALUE").value)
	 iYrRQty= trim(ItemNode.attributes.getNamedItem("YEARRECEIPTQUANTITY").value)
	 iYrRVal= trim(ItemNode.attributes.getNamedItem("YEARRECEIPTVALUE").value)
	 iYrIQty= trim(ItemNode.attributes.getNamedItem("YEARISSUEQUANTITY").value)
	 iYrIVal= trim(ItemNode.attributes.getNamedItem("YEARISSUEVALUE").value)
	 iYrClStock = Trim(ItemNode.attributes.getNamedItem("YEARCLOSINGSTOCK").value)
	 iYrClValue = Trim(ItemNode.attributes.getNamedItem("YEARCLOSINGSTOCK").value)
	 
	 Response.Write "<P>iYrOpStock = "& iYrOpStock 
	 Response.Write "<P>iYrRQty = "& iYrRQty 
	 Response.Write "<P>iYrIQty = "& iYrIQty 
	 
	 if Trim(iLoc)="0" or IsNull(iLoc) then iLoc = "NULL"
	 if Trim(iBin)="0" or IsNull(iBin) then iBin = "NULL"
	 
         
         'iYrClStock = cdbl(iYrOpStock)+cdbl(iYrRQty)-cdbl(iYrIQty)
         'iYrClValue = cdbl(iYrOpValue)+CDbl(iYrRVal)-CDbl(iYrIVal)
         
         'if cdbl(iYrClValue)<0 then
         '   iYrClValue = cdbl(iYrClValue)*-1
         'end if
         
         'sSql = "Select YearClosingStock from INV_T_ItemYearlyStock WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " and Itemcode = "& iItemCode 
         'dcrs.Open sSql,con
         'if not dcrs.EOF then
         '   iYrStClQty = dcrs(0)
         'end if
         'dcrs.Close 
         
         'sSql = "Select IsNull(SUM(LotQuantityNett),0) from INV_T_ItemLedgerOpeningPacks where ItemCode = "& iItemCode &" and TransactionDate =Convert(datetime,'"& dCurFinStartDate &"',103) and TransactionType ='RO'"
         'dcrs.Open sSql,con
         'if not dcrs.EOF then
         '   iSumOfNettQty = dcrs(0)
         'end if
      '   dcrs.Close 
         
     '    if CDbl(iYrStClQty)=cdbl(iSumOfNettQty) then
    '        Response.Write "<p>Stock are equal"
   '      end if
         
  '       if iBin = "NULL" then
  '          sSql = "Select IsNull(SUM(LotQuantityNett),0) from INV_T_ItemLedgerOpeningPacks where ItemCode = "& iItemCode &" and TransactionDate =Convert(datetime,'"& dCurFinStartDate &"',103) and TransactionType ='RO' and StorageLocationNo = "& iLoc &" and StorageBinNumber is null"
 '        else
'            sSql = "Select IsNull(SUM(LotQuantityNett),0) from INV_T_ItemLedgerOpeningPacks where ItemCode = "& iItemCode &" and TransactionDate =Convert(datetime,'"& dCurFinStartDate &"',103) and TransactionType ='RO' and StorageLocationNo = "& iLoc &" and StorageBinNumber = "& iBin 
         'end if 
         'Response.Write "<p>"& sSql
         'dcrs.Open sSql,con
         'if not dcrs.EOF then
          '  iYrClStock = dcrs(0)
         'end if
         'dcrs.Close 
         
         
         'if iBin ="NULL" then
        '    sSql = "Update INV_T_ITEMLOCATIONSTOCK set YEARCLOSINGSTOCK="& iYrClStock &",YEARCLOSINGVALUE="& iYrClValue  &"  WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " and Itemcode = "& iItemCode &" AND LOCATIONNUMBER = " & iLoc & " AND BINNUMBER IS NULL"
       '  else 
      '      sSql = "Update INV_T_ITEMLOCATIONSTOCK set YEARCLOSINGSTOCK="& iYrClStock &",YEARCLOSINGVALUE="& iYrClValue  &"  WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " and Itemcode = "& iItemCode &" AND LOCATIONNUMBER = " & iLoc & " AND BINNUMBER = " & iBin 
     '    end if 
    '     Response.Write "<p>"& sSql
    '     con.execute sSql

	if iBin = "0" then iBin = "NULL"

	if iBin = "NULL" then
		sQuery = "SELECT ORGANISATIONCODE,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARISSUEQUANTITY,YEARISSUEVALUE,YEAROPENINGSTOCK,YEAROPENINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND LOCATIONNUMBER = " & iLoc & " AND BINNUMBER IS NULL AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dCurFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dCurFinEndDate) & ""
	else
		sQuery = "SELECT ORGANISATIONCODE,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARISSUEQUANTITY,YEARISSUEVALUE,YEAROPENINGSTOCK,YEAROPENINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND LOCATIONNUMBER = " & iLoc & " AND BINNUMBER = " & iBin & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dCurFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dCurFinEndDate) & ""
	end if
	Response.write "<p>"&sQuery
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iRecQty = trim(dcrs(5))
		iRecValue = trim(dcrs(6))
		iIssQty = trim(dcrs(7))
		iIssValue = trim(dcrs(8))
		
		

		iClStock = (cdbl(iYrClStock) + cdbl(iRecQty)) - cdbl(iIssQty)
		iClValue = (cdbl(iYrClValue) + cdbl(iRecValue)) - cdbl(iIssValue)

		sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEAROPENINGSTOCK = " & iYrClStock & "," &_
			"YEAROPENINGVALUE = " & iYrClValue & ",YEARCLOSINGSTOCK = " & iClStock & ", " &_
			"YEARCLOSINGVALUE = " & iClValue & " WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
			"ORGANISATIONCODE = " & Pack(sUnit) & " AND " &_
			"LOCATIONNUMBER = " & iLoc & " AND (BINNUMBER = " & iBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(dCurFinStartDate) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(dCurFinEndDate) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
		Response.Write "<p>"& sSql & "<BR>"
		con.Execute sSql
	else
	
	'Response.Write "<p>iRecQty = "& iRecQty &" <p>IssueQty = "& iIssQty 
		sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
			"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
			"YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
			"(" & Pack(sUnit) & "," & iClassCode & "," & iItemCode & "," &_
			"CONVERT(DATETIME," & Pack(dCurFinStartDate) & ",103),CONVERT(DATETIME," & Pack(dCurFinEndDate) & ",103)," &_
			"" & iLoc & "," & iBin & "," &_
			"" & iYrClStock & "," & iYrClValue & "," & iYrClStock & "," & iYrClValue & ")"
		Response.Write "<p>"& sSql & "<BR>"
		con.Execute sSql
		
	
    ''Blocked by ragav on April 01,2011

	'	with dcrs1
	'		.CursorLocation = 3
	'		.CursorType = 3
	'		.Source = "SELECT ISNULL(MAX(STOCKNO)+1,1) FROM INV_M_STOCKSTATUS"
	'		.ActiveConnection = con
	'		.Open
	'	end with
	'	set dcrs1.ActiveConnection = nothing
'
'		if not dcrs1.EOF then
'			sSql = "INSERT INTO INV_M_STOCKSTATUS (STOCKNO,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
'				"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER) VALUES " &_
'				"(" & trim(dcrs1(0)) & "," & Pack(sUnit) & "," & iClassCode & "," & iItemCode & "," &_
'				"CONVERT(DATETIME," & Pack(dCurFinStartDate) & ",103),CONVERT(DATETIME," & Pack(dCurFinEndDate) & ",103)," &_
'				"" & iLoc & "," & iBin & ")"
'			Response.Write sSql & "<BR>"
'			con.Execute sSql
'		end if
'		dcrs1.Close

    ''end of ragav blocked

	end if
	dcrs.Close

next
YrLocStockXML.save server.MapPath("../Temp/LocStockData.xml")

'Insert to MS_StockClosing table
    if Trim(sCategoryCode)="" or IsNull(sCategoryCode) then sCategoryCode = "NULL"
    if Trim(sCategoryCode)<>"NULL" then sCategoryCode = Pack(sCategoryCode)
    if Trim(sSubCategory)="" or IsNull(sSubCategory) then sSubCategory = "NULL"
    if Trim(sClassification)="" or IsNull(sClassification) then sClassification ="NULL"
    

    sQuery = "Insert into MS_StockClosing (FromPeriod,ToPeriod,OUDefinitionID,Transferred,"&_
             "TransferredBy,TransferredOn,CategoryCode,SubCategory,Classification) values(Convert(datetime,'"& dPreFinStartDate &"',103),"&_
             "Convert(datetime,'"& dPreFinEndDate &"',103),'"& sUnit &"','Y',"& iCreadBy &",Convert(datetime,getDate(),103),"& sCategoryCode &","& sSubCategory &","& sClassification &")"
    Response.Write "<p>"& sQuery
    con.execute sQuery
    


if con.Errors.count <> 0 then
	dim iErrCounter
	con.RollbackTrans
	for iErrCounter=0 to con.Errors.count
		Response.Write con.Errors(iErrCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	Response.End
	Response.Clear
	con.CommitTrans
end if

con.close
set con = nothing



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sQuery = "SELECT ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE,RTRIM(CONVERT(CHAR,FINANCIALYEARFROM,103)) AS FINANCIALYEARFROM,RTRIM(CONVERT(CHAR,FINANCIALYEARTO,103)) AS FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER,YEAROPENINGSTOCK,YEAROPENINGVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCATIONSTOCK WHERE ORGANISATIONCODE = " & Pack(sUnit) & " AND CONVERT(CHAR,FINANCIALYEARFROM,103) = " & Pack(dPreFinStartDate) & " AND CONVERT(CHAR,FINANCIALYEARTO,103) = " & Pack(dPreFinEndDate) & " ORDER BY 1,2,3,4,5 FOR XML AUTO&root=ROOT"
'Response.Write sQuery
'xmlhttp.Open "POST","http://192.168.1.1/dev_itms?sql="&sQuery, false
'xmlhttp.send sQuery

'newxml.load xmlhttp.responseXML
'newxml.save server.MapPath("../Temp/StockData.xml")
'Response.End
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%>
<SCRIPT LANGUAGE=vbscript>
	alert("Closing Stock has been carry forwarded to the Current Financial Year")
	window.location.href = "CloseEntry.asp?Frm=IS"
</SCRIPT>
