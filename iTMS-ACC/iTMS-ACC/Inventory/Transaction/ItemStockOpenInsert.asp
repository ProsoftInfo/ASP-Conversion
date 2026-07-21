
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItemStockOpenInsert.asp
	'Module Name				:	Inventory (Stock Opening - Insert)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Oct 20,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
    Dim ObjDOM,rsTemp
    Dim ndRoot,ndItem,ndLoc
    Dim sItemCode,sClassCode,sOrgCode,sLocNo,sBinNo,sTotChageStock
    Dim sOpRate,sStkChange,sStkVal,sFinPeriod,sFinFrom,sFinTo
    Dim sQuery,sOpeningValue,sTotItemQty,sTotItemValue
    Dim iInvRecNo
    
    set ObjDOM = Server.CreateObject("Microsoft.XMLDOM")
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    ObjDOM.load Server.MapPath("../temp/transaction/ItemOpenStockChange_"& Session.SessionID&".xml")
    
    sFinPeriod = split(Session("FinPeriod"),":")
    sFinFrom = "01/04/"& sFinPeriod(0)
    sFinTo = "31/03/"& sFinPeriod(1)
    
    con.begintrans
    set ndRoot = ObjDOM.documentElement
    if ndRoot.hasChildNodes() then
        for each ndItem in ndRoot.childNodes
            if ndItem.nodeName="Item" then
                sTotItemQty = 0
                sTotItemValue = 0
                sTotChageStock = 0
                sItemCode = ndItem.getAttribute("ItemCode")
                sClassCode = ndItem.getAttribute("ClassCode")
                sOrgCode = ndItem.getAttribute("OrgCode")
                if ndItem.hasChildNodes() then
                    for each ndLoc in ndItem.childNodes
                        
                        if ndLoc.nodeName="Loc" then
                            sLocNo = ndLoc.getAttribute("Loc")
                            sBinNo = ndLoc.getAttribute("Bin")
                            sTotChageStock = ndLoc.getAttribute("TotChangeQty")
                            sOpRate = ndLoc.getAttribute("Rate")
                            sStkChange = ndLoc.getAttribute("StkChange")
                            sStkVal = ndLoc.getAttribute("StkValue")
                            
                           ' if Trim(sTotChageStock)<>"0" and Trim(sTotChageStock)<>"" then
                            
                                if Trim(sOpRate)<>"0" and Trim(sOpRate)<>"" then
                                    sOpeningValue = CDbl(sTotChageStock)*CDbl(sOpRate)
                                else
                                    if sStkChange = sTotChageStock then
                                        sOpeningValue = sStkVal 
                                    else
                                        sOpeningValue = (cdbl(sStkVal)/CDbl(sStkChange))*cdbl(sTotChageStock)
                                    end if
                                    sOpRate = CDbl(sStkVal)/CDbl(sStkChange)
                                end if 'if Trim(sOpRate)<>"0" and Trim(sOpRate)<>"" then
                                
                                sTotItemQty = CDbl(sTotItemQty) + CDbl(sTotChageStock)
                                sTotItemValue = CDbl(sTotItemValue) + CDbl(sOpeningValue)
                                
                                sQuery = "Select InventoryReceiptNo from INV_T_LocationLot where ItemCode = "& sItemCode &" and OrganisationCode = '"& sOrgCode &"' "&_
                                         " and SrcType = 'RO' and StorageLocationNo = "&  sLocNo &" and (StorageBinNumber = "& sBinNo &" or StorageBinNumber is Null)  and "&_
                                         " Convert(datetime,DateOfReceipt,103)>=Convert(datetime,'"& sFinFrom &"',103) and  "&_
                                         " Convert(datetime,DateOfReceipt,103)<=Convert(datetime,'"& sFinTo &"',103) "
                                Response.Write "<p>"& sQuery
                                rsTemp.Open sQuery,con
                                if not rsTemp.EOF then
                                    iInvRecNo = rsTemp(0)
                                end if
                                rsTemp.Close 
                                
                                sQuery = "Update INV_T_LocationLot set LotQuantityGross ="& sTotChageStock &",LotQuantityNett="& sTotChageStock &",Rate = "& sOpRate &" where ItemCode = "& sItemCode  &" "&_
                                         " and OrganisationCode = '"& sOrgCode &"' and SrcType = 'RO' and StorageLocationNo = "& sLocNo &" and (StorageBinNumber = "& sBinNo &" or "&_
                                         " StorageBinNumber is Null) and Convert(datetime,DateOfReceipt,103)>=Convert(datetime,'"& sFinFrom  &"',103) and  "&_
                                         " Convert(datetime,DateOfReceipt,103)<=Convert(datetime,'"& sFinTo &"',103) "
                                Response.Write "<p>"&sQuery
                                con.execute sQuery
                                
                                sQuery = "Update INV_T_ItemLocationStock set YearOpeningStock ="& sTotChageStock &",YearOpeningValue ="& sOpeningValue &","&_
                                         " YearClosingStock = "& sTotChageStock &"+YearReceiptQuantity-YearIssueQuantity, "&_
                                         " YearClosingValue = "& sOpeningValue &"+YearReceiptvalue-YearIssueValue where ItemCode = "& sItemCode &" "&_
                                         " and LocationNumber = "& sLocNo &" and (BinNumber = "& sBinNo &" or BinNumber is Null) and Convert(datetime,FinancialYearFrom,103) = "&_
                                         " Convert(datetime,'"& sFinFrom &"',103) and Convert(datetime,FinancialYearTo,103) = Convert(datetime,'"& sFinTo &"',103) "
                                Response.Write "<p>"&sQuery
                                con.execute sQuery
                                
                                
                                
                           ' end if 'if Trim(sTotChageStock)<>"0" and Trim(sTotChageStock)<>"" then
                        end if 'if ndLoc.nodeName="Loc" then
                    next
                end if 'if ndItem.hasChildNodes() then
                
                sQuery = "Update INV_T_ItemLedger set TransactQuantity ="& sTotItemQty &",TransactValue="& sTotItemValue &" where TransactionType = 'RO' "&_
                         " and TransactionNo = "& iInvRecNo &" and ItemCode = "& sItemCode &" and Convert(datetime,TransactionDate,103)>=Convert(datetime,'"& sFinFrom &"',103) "&_
                         " and Convert(datetime,TransactionDate,103) <= Convert(datetime,'"& sFinTo &"',103) "
                Response.Write "<p>"&sQuery
                con.execute sQuery
                
                sQuery = "Update INV_T_ItemYearlyStock set YearOpeningStock ="& sTotItemQty &" ,YearOpeningValue ="& sTotItemValue &" , "&_
                         " YearClosingStock = "& sTotItemQty &"+YearReceiptQuantity-YearIssueQuantity, "&_
                         " YearClosingValue = "& sTotItemValue &"+YearReceiptvalue-YearIssueValue "&_
                         " where ItemCode = "& sItemCode &" and Convert(datetime,FinancialYearFrom,103) =  "&_
                         " Convert(datetime,'"& sFinFrom &"',103) and Convert(datetime,FinancialYearTo,103) = "&_
                         " Convert(datetime,'"& sFinTo &"',103)"
                Response.Write "<p>"&sQuery
                con.execute sQuery
            end if 'if ndItem.nodeName="Item" then
        next
    end if 'if ndRoot.hasChildNodes() then

    if con.Errors.count <> 0 then
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
    Response.Redirect "../Master/ITEMLISTENTRY.ASP?ACTN=SO"
	end if 'if con.Errors.count <> 0 then
%>
