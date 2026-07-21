<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetRefItemForIssue.asp
	'Module Name				:	Inventory (MR Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Feb 22,2011
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
	Dim sOrgCode,sItemType,sStock,sRefCodes,sQuery,sRefType,sRefNo
	Dim iEntryNo
	Set rsItem	=	Server.CreateObject("ADODB.Recordset")
	Set objDOM	=	Server.CreateObject("Microsoft.XMLDOM")
	
	sOrgCode = Session("organisationcode")
	sRefNo = Request.QueryString("RefNo")
	sRefType = Request.QueryString("RefType")
	
	if sRefType = "11" then 'Material Requisition
		sQuery = "Select V.CompanyItemCode,(V.ItemDescription + '--'+ isnull(V.CatalogueNo,'N/A')),"&_
			     " V.GroupName,V.StoresUOM,V.ItemCode,V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,"&_
			     " isNull(M.ItemAttributes,0),V.ItemTypeID from Inv_T_MRSItemDetails M,VWITEM V where MRSNumber in ("& sRefCodes &")"&_
			     " and V.ItemCode=M.ITemCode and V.ClassificationCode = M.ClassificationCode and M.OrganisationCode = V.OrganisationCode"&_
			     " and M.OrganisationCode = '"& sOrgCode &"'"
'	elseif sRefType= "12" then ' Material Issue
'		sQuery = "Select V.CompanyItemCode,(V.ItemDescription + '--'+ isnull(V.CatalogueNo,'N/A')),"&_
'			     " V.GroupName,V.StoresUOM,V.ItemCode,V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,"&_
'			     " isNull(M.ItemAttributes,0),V.ItemTypeID from INV_T_MaterialIssueDetails M,VWITEM V where IssueEntryNo in ("& sRefCodes &")"&_
'			     " and V.ItemCode =M.ItemCode and V.ClassificationCode=M.ClassificationCode and M.OrganisationCode = '"& sOrgCode &"'"&_
'			     " and M.OrganisationCode = V.OrganisationCode"
    elseif sRefType = "14" then ' Mix Code
    elseif sRefType = "15" then ' Sales Order
        sQuery = " Select V.CompanyItemCode,V.ItemDescription,V.GroupName,V.StoresUoM,V.ItemCode,"&_
                 " V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,'0',V.ItemTypeID "&_
                 " from Sal_T_OrdersDetails M,VwItem V where OrderNumber in("& sRefCodes &") and V.ItemCode = M.ItemCode"&_
                 " and V.ClassificationCode=M.ClassificationCode and V.OrganisationCode = '"& sOrgCode &"'"
    elseif sRefType = "17" then ' Production Order
        
	end if
	'Response.Write sQuery		     
	rsItem.Open sQuery,con
	if not rsItem.EOF then
		iEntryNo = 0
			set ndRoot = objDOM.createElement("Root")
			objDOM.appendChild(ndRoot)
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
			ndRoot.appendChild(ndItem)
			rsItem.MoveNext
		loop
	end if
	rsItem.Close 
	
	Response.ContentType = "text/xml"
	Response.Write objDOM.xml
%>
