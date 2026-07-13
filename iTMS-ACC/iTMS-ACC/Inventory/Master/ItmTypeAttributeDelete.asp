<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmTypeAttributeDelete.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	March 08,2010
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	
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
dim sTypeName,sQuery,sItemType,iAttrID,sClassCode

iAttrID = Request.QueryString("sValue")
sItemType = Request.QueryString("ItemType")
sClassCode = Request.QueryString("ClassCode")

sQuery  = "Delete from INV_M_ITEMTYPEOPTIONS where ItemTypeAttributeID = "& iAttrID
Response.Write "<p>"& sQuery
con.execute sQuery

sQuery ="Delete from INV_M_ITEMTYPEATTRIBUTES Where ItemTypeAttributeID = "& iAttrID
	'Response.Write sQuery & vbCrLf & vbCrLf
	con.Execute sQuery
	'Response.End 
Response.Redirect "ItmTypeAttributeEntry.asp?ItemType="&sItemType&"&ClassCode="&sClassCode 
%>
