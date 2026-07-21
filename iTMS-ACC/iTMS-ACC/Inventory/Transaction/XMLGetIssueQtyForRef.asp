<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetIssueQtyForRef.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Dec 30,2011
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
	Dim sOrgCode,sItemType,sStock,sRefCodes,sQuery,sRefType,sItemCode,sClassCode
	Dim iEntryNo
	Set rsItem	=	Server.CreateObject("ADODB.Recordset")
	Set objDOM	=	Server.CreateObject("Microsoft.XMLDOM")
	
	sOrgCode = Request.QueryString("orgID")
	sRefCodes = Request.QueryString("RefCodes")
	sRefCodes = Replace(sRefCodes,",","','")
	sRefType = Request.QueryString("RefType")
	sItemCode = Request.QueryString("ItemCode")
	sClassCode = Request.QueryString("ClassCode")
	set ndRoot = objDOM.createElement("Root")
	objDOM.appendChild(ndRoot)

	'Response.Write "Reftype = "& sRefType
	if Trim(sRefType)<>"" then
	    sQuery = "Select ItemCode,ClassificationCode,isNull(SUM(QuantityIssued),0) from Inv_T_MaterialIssueDetails where IssueEntryNo in("&_
	             " Select IssueEntryNo from Inv_T_MaterialIssueHeader where OrganisationCode = '"& sOrgCode &"'"&_
	             " and AppRefType in("& sRefType &") and AppRefNo in("& sRefCodes &")) and ItemCode = "& sItemCode  &" and ClassificationCode = "& sClassCode &""&_
	             " Group by ItemCode,ClassificationCode"
    	'Response.Write sQuery
	    rsItem.Open sQuery,con
	    if not rsItem.EOF then
			    set ndItem = objDOM.createElement("Item")
			        ndItem.setAttribute "ItemCode",rsItem(0)
			        ndItem.setAttribute "ClassCode",rsItem(1)
				    ndItem.setAttribute "ToIssueQty",rsItem(2)
			    ndRoot.appendChild(ndItem)
	    end if
	    rsItem.Close 
	end if 'if Trim(sRefType)<>"" then
	
	Response.Clear 
	Response.ContentType = "text/xml"
	Response.Write objDOM.xml
%>
