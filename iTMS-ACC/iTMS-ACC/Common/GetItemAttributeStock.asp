<%Option Explicit%>
<%
	'Program Name				:	GetItemAttributeStock.asp
	'Module Name				:	
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim rsObj
Dim iItemCode,iAttributeID,iStock
Dim sQuery
set rsObj = Server.CreateObject("ADODB.Recordset")
iItemCode = Request.QueryString("ItemCode")
iAttributeID =  Request.QueryString("AttID")

    if Trim(iAttributeID)<>"0" and Trim(iAttributeID)<>"" then
        sQuery ="Select isNull(SUM(isNull(LotQuantityNett,0)-isNull(QuantityIssued,0)),0) from Inv_T_LocationLot "&_
                "where isNull(LotQuantityNett,0)-isNull(QuantityIssued,0)>0 and ItemCode = "& iItemCode &" and AttributeList like '%"& iAttributeID&"%'"
                'Response.Write sQuery
                rsObj.Open sQuery,con
                if not rsObj.EOF then
                    iStock = cdbl(rsObj(0))
                end if
                rsObj.Close 
                
                if iStock>0 then
                    Response.Write "Y"
                else    
                    Response.Write "N"
                end if
    else
        sQuery ="Select isNull(SUM(isNull(LotQuantityNett,0)-isNull(QuantityIssued,0)),0) from Inv_T_LocationLot "&_
                "where isNull(LotQuantityNett,0)-isNull(QuantityIssued,0)>0 and ItemCode = "& iItemCode 
               ' Response.Write sQuery
                rsObj.Open sQuery,con
                if not rsObj.EOF then
                    iStock = cdbl(rsObj(0))
                end if
                rsObj.Close 
                
                if iStock>0 then
                    Response.Write "Y"
                else    
                    Response.Write "N"
                end if
    end if 'if Trim(iAttributeID)<>"0" and Trim(iAttributeID)<>"" then
%>

