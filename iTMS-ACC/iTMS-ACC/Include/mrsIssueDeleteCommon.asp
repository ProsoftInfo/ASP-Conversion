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
    Dim dcrs,dcrs1,dcrs2
    Dim sFinPeriod,sArrFin,sFinFrom,sFinTo,sSql,sLoc
    Dim iItemCode,nExistTransQty,nExistTransVal
    Set dcrs = Server.CreateObject("ADODB.RecordSet")
    Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
    Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
    
    

Function MrsIssueDelete(iLedIssueNo)

sFinPeriod = Session("FinPeriod")
sArrFin = Split(sFinPeriod,":")
sFinFrom = "01/04/"& sArrFin(0)
sFinTo = "31/03/"&sArrFin(1)

    	    
		    sSql = "Select ItemCode,LocationNumber from INV_T_MaterialIssueDetails where IssueEntryNo = "& iLedIssueNo 
		    sSql = sSql &" group by ItemCode,LocationNumber"
		    with dcrs1
		        .CursorLocation = 3
		        .CursorType = 3
		        .ActiveConnection = con
		        .Source = sSql 
		        .Open 
		    End With 
		    if not dcrs1.EOF then
		        do while not dcrs1.EOF
		            iItemCode = dcrs1(0)
		            sLoc = dcrs1(1)
		            
		            sSql = "Select SerialNo,LocationNumber,QuantityIssued from INV_T_Materialissuedetails where IssueEntryNo = "& iLedIssueNo &" and ItemCode = "& iItemCode 
		            Response.Write "<p>"& sSql
		            dcrs2.open sSql,con
		            if not dcrs2.eof then
		                do while not dcrs2.EOF 
		                    sSql = "Update INV_T_LocationLot set QuantityIssued = QuantityIssued - "& dcrs2(2) &" where SerialNumber ="& dcrs2(0) &" and StorageLocationNo = "& dcrs2(1)
		                    Response.Write "<p>"& sSql
		                    con.execute sSql
		                    dcrs2.MoveNext 
		                loop
		            end if 
		            dcrs2.close 
		            
		            sSql = "Select SUM(TransactQuantity),SUM(TransactValue) from INV_T_ItemLedger where TransactionNo = "& iLedIssueNo &" and TransactionType = 'I' and ItemCode = "& iItemCode 
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
    			    
			        sSql = "Update INV_T_ItemYearlyStock set YearIssueQuantity=YearIssueQuantity-"& nExistTransQty &",YearIssueValue=YearIssueValue-"& nExistTransVal &" where ItemCode = "& iItemCode 
			        sSql = sSql &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
			        Response.Write "<p>"& sSql 
			        con.execute sSql
    			    
                    sSql = " Update INV_T_ItemYearlyStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                    sSql = sSql & "   where ItemCode = "& iItemCode 
                    sSql = sSql & " and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                    Response.Write "<p>"& sSql 
	                con.execute sSql 
	                
	                
	                sSql = "Select YearClosingStock,YearClosingValue from INV_T_ItemYearlyStock where ItemCode = "& iItemCode 
	                dcrs.open sSql,con
	                if not dcrs.eof then
	                    Response.Write "<p>Year Closing Stock = "& dcrs(0)
	                    Response.Write "<p>Year Closing Stock = "& dcrs(1)
	                end if 
	                dcrs.close
	            
	                
	                sSql = "Update INV_T_ItemLocationStock set YearIssueQuantity=YearIssueQuantity-"& nExistTransQty &",YearIssueValue=YearIssueValue-"& nExistTransVal &"  where ItemCode = "& iItemCode 
				    sSql = sSql &" and LocationNumber = "& sLoc &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
				    Response.Write "<p>"& sSql 
				    con.execute sSql
				    
				    sSql = " Update INV_T_ItemLocationStock set YearClosingStock=YearOpeningStock+YearReceiptQuantity-YearIssueQuantity "
                    sSql = sSql & "  where ItemCode = "& iItemCode 
                    sSql = sSql & " and LocationNumber = "& sLoc &"  and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103)"
                    Response.Write "<p>"& sSql 
	                con.execute sSql 
	                
	                sSql = "Delete from INV_T_ItemLedger where TransactionNo = "& iLedIssueNo &" and TransactionType= 'I' and ItemCode ="&iItemCode 
			        Response.Write "<p>"&sSql 
			        con.execute sSql 
	                
		            dcrs1.MoveNext 
		        loop
		    end if 
		    dcrs1.Close 
		    
		    
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
		    
		    sSql = "Delete from INV_T_MaterialIssueHeader where IssueEntryNo = "& iLedIssueNo 
		    Response.Write "<p>"&sSql 
		    con.execute sSql 
		    
 End Function 'Function MrsIssueInsert()
    
%>

