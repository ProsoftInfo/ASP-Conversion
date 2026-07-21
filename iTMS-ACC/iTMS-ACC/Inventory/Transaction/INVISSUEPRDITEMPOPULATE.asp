<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	INVISSUEPRDITEMPOPULATE.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Sep 13,2010
	'							:
	'Connects To				:	DirectIssueItemEntry.asp
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
	Dim sOrgCode,sItemType,sStock,sRefCodes,sQuery
	Dim iEntryNo
	Set rsItem	=	Server.CreateObject("ADODB.Recordset")
	Set objDOM	=	Server.CreateObject("Microsoft.XMLDOM")
	
	sOrgCode = Request.QueryString("orgID")
	sItemType= Request.QueryString("sIType")
	sStock = Request.QueryString("Stock")
	sRefCodes = Request.QueryString("RefCodes")
	sRefCodes = Replace(sRefCodes,",","','")
	
	sQuery = "Select VI.CompanyItemCode,(VI.ItemDescription+ ' -- ' + isnull(DrawingNumber,'N/A') + ' -- ' + isnull(VI.CatalogueNo,'N/A')),VI.GROUPNAME,"&_
		     " VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DecimalAllowed,VI.ReceiptNumbering,isNull(AM.AttributeList,0) from VWITEM VI,"&_
		     " APP_M_MIXVARIETYQUALITYPARAMETERS AM where VI.ItemCode = AM.ItemCode and AM.CategoryCode = VI.ClassificationCode and VI.ITEMTYPEID = '"& sItemType &"'"&_
		     " AND VI.ORGANISATIONCODE ='"& sOrgCode & "' AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 and MixCode in ('"& sRefCodes &"')"
	'Response.Write sQuery	     
	rsItem.Open sQuery,con
	if not rsItem.EOF then
		iEntryNo = 0
			set ndRoot = objDOM.createElement("Root")
			ndRoot.setAttribute "ItemType",sItemType
			objDOM.appendChild(ndRoot)
		do while not rsItem.EOF 
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
