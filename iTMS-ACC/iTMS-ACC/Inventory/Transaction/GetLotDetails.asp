<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<%
	'Program Name				:	GetLotDetails.asp
	'Module Name				:	Inventory 
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Jun 06,2011
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
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
    Dim rsTemp,objDOM
    Dim sQuery,sInvRcptNo,sSerNo
    Dim ndRoot,ndLot
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    set objDOM = Server.CreateObject("Microsoft.XMLDOM")
    
    sInvRcptNo = Request.QueryString("InvNo")
    sSerNo = Request.QueryString("SerialNo")
    
    set ndRoot = objDOM.createElement("Root")
    objDOM.appendChild ndRoot
    
    if trim(sInvRcptNo)<>"" and trim(sSerNo)<>"" then
        sQuery = "Select isNull(PackingNumber,0),PackingCode,isNull(SellingNumber,0),WeightPerSellingForm, "&_
                 " isNull(SellingForm,0),isNull(AttributeList,'') from INV_T_LocationLot Where "&_
                 " InventoryReceiptNo = "& sInvRcptNo  &"  and SerialNumber = "& sSerNo
    else
        sQuery = "Select isNull(PackingNumber,0),PackingCode,isNull(SellingNumber,0),WeightPerSellingForm, "&_
                 " isNull(SellingForm,0),isNull(AttributeList,'') from INV_T_LocationLot Where "&_
                 " SerialNumber = "& sSerNo
    end if 'if trim(sInvRcptNo)<>"" and trim(iSerNo)<>"" then
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        set ndLot = objDOM.createElement("Lot")
            ndLot.setAttribute "PackNo",rsTemp(0)
            ndLot.setAttribute "PackCode",rsTemp(1)
            ndLot.setAttribute "SellNo",rsTemp(2)
            ndLot.setAttribute "WeightPerSellForm",rsTemp(3)
            ndLot.setAttribute "SellForm",rsTemp(4)
            ndLot.setAttribute "AttributeList",rsTemp(5)
            ndRoot.appendChild ndLot
    end if
    rsTemp.Close
    
    Response.ContentType = "text/xml"
    Response.Write objDOM.xml
%>