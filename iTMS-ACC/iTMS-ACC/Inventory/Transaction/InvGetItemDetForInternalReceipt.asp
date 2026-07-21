<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	InvGetItemDetForInternalReceipt.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Feb 19,2013
	'							:
	'Connects To				:	receiptInternalEntry.asp
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
	Dim	rsItem,objDOM
	Dim ndRoot,ndItem
	Dim sOrgCode,sItemType,sStock,sRefCodes,sQuery,sRefType
	Dim iEntryNo
	Set rsItem	=	Server.CreateObject("ADODB.Recordset")
	Set objDOM	=	Server.CreateObject("Microsoft.XMLDOM")
	
	sOrgCode = Request.QueryString("orgID")
	sItemType= Request.QueryString("sIType")
	sStock = Request.QueryString("Stock")
	sRefCodes = Request.QueryString("RefCodes")
	sRefCodes = Replace(sRefCodes,",","','")
	sRefType = Request.QueryString("RefType")
	set ndRoot = objDOM.createElement("Root")
	objDOM.appendChild(ndRoot)

	'Response.Write "Reftype = "& sRefType
	if sRefType = "12" then 'Material Issue
		sQuery = "Select V.CompanyItemCode,(V.ItemDescription + '--'+ isnull(V.CatalogueNo,'N/A')),"&_
			     " V.GroupName,V.StoresUOM,V.ItemCode,V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,"&_
			     " isNull(M.ItemAttributes,''),'',IssueEntryNo,SUM(QuantityIssued) from INV_T_MaterialIssueDetails M,VWITEM V where IssueEntryNo in ("& sRefCodes &")"&_
			     " and V.ItemCode=M.ITemCode and V.ClassificationCode = M.ClassificationCode and M.OrganisationCode = V.OrganisationCode"&_
			     " and M.OrganisationCode = '"& sOrgCode &"' Group By V.CompanyItemCode,V.ItemDescription,V.CatalogueNo,V.GroupName,V.StoresUOM,"&_
			     " V.ItemCode,V.ClassificationCode,V.DecimalAllowed,V.REceiptNumbering,M.ItemAttributes,M.IssueEntryNo"
    elseif sRefType = "17" then ' Production Order
        sQuery = " Select V.CompanyItemCode,V.ItemDescription,V.GroupName,V.StoresUoM,V.ItemCode,"&_
                 " V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,isNull(AttributeList,''),"&_
                 " V.ItemTypeID,ProductionOrderNo,RequiredQuantity from PRD_T_ProductDetails P,VWITEM V where V.ItemCode=P.ItemCode and "&_
                 " V.ClassificationCode = P.ClassificationCode and ProductionOrderNo in ('"& sRefCodes &"') and V.OrganisationCode = '"& sOrgCode &"'"
    elseif sRefType = "42" then ' Sales Order - Jobwork
        sQuery = "Select V.CompanyItemCode,V.ItemDescription,V.GroupName,V.StoresUOM,V.ItemCode,V.ClassificationCode, "&_
                 " V.DecimalAllowed,V.ReceiptNumbering,IsNull(S.AttributeList,''),V.ItemTypeID,SalesOrderNo,0 from "&_
                 " Sal_T_SOItemAdditionaldetail S,VWItem V where V.ItemCode = S.ItemSendAs and "&_
                 " V.ClassificationCode = S.ClassSendAs and SalesOrderNo in ('"& sRefCodes &"') and V.OrganisationCode = '"& sOrgCode &"'"
        
    end if
	'Response.Write sQuery		     
	rsItem.Open sQuery,con
	if not rsItem.EOF then
		iEntryNo = 0
		do while not rsItem.EOF 
			ndRoot.setAttribute "ItemType",rsItem(9)
			iEntryNo = iEntryNo + 1
			set ndItem = objDOM.createElement("Item")
				ndItem.setAttribute "EntryNo",iEntryNo
				ndItem.setAttribute "CompanyItemCode",rsItem(0)
				ndItem.setAttribute "ItemCode",rsItem(4)
				ndItem.setAttribute "ClassCode",rsItem(5)
				ndItem.setAttribute "ItemName",rsItem(1)
				ndItem.setAttribute "ClassName",rsItem(2)
				ndItem.setAttribute "StoresUoM",rsItem(3)
				ndItem.setAttribute "Decimal",rsItem(6)
				ndItem.setAttribute "ReceiptNum",rsItem(7)
				ndItem.setAttribute "AttributeList",rsItem(8)
				ndItem.setAttribute "No",rsItem(10)
				ndItem.setAttribute "Qty",rsItem(11)
			ndRoot.appendChild(ndItem)
			rsItem.MoveNext
		loop
	end if
	rsItem.Close 
	
	Response.ContentType = "text/xml"
	Response.Write objDOM.xml
%>
