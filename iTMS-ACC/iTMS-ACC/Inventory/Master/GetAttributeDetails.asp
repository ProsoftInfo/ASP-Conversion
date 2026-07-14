<%@ Language=VBScript %>
<% Option Explicit%>
<%
	'Program Name				:	GetAttributeDetails
	'Module Name				:	Inventory (Master)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	May 26,2010
	'Modified By				:	
	'Modified On				:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
	Dim objDom
	Dim ndRoot,ndAttributeList
	Dim rs,rsDetails
	Dim iAttributeList,sItemType,sAttributeListValue
	Dim sQuery,sReqValue,sArrValue
	
	sReqValue = Request.QueryString("sTemp")
	sArrValue = split(sReqValue,":")
	iAttributeList = sArrValue(0)
	sItemType = sArrValue(1)

	set objDom  = Server.CreateObject("Microsoft.XMLDOM")
	set rs		= Server.CreateObject("ADODB.Recordset")
	set rsDetails = Server.CreateObject("ADODB.Recordset")
	if trim(iAttributeList)<>"" then
	   ' sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where "&_
		'	     " O.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemTypeID='"&sItemType&"' and A.ItemTypeAttributeID in ("&iAttributeList&") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID "
	    sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,'' from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where "&_
			     " O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID in ("&iAttributeList&") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName "
			    ' Response.Write sQuery
    			 
	    rs.Open sQuery,con
	    if rs.EOF then
		   ' sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O "&_
			'     "where O.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemTypeID = '"& sItemType &"' and O.OptionValue = "& iAttributeList
			sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O "&_
			     "where O.ItemTypeAttributeID = A.ItemTypeAttributeID and O.OptionValue = "& iAttributeList
			     'Response.Write  sQuery
		    rsDetails.Open sQuery,con
		    if not rsDetails.EOF then
			    sAttributeListValue = trim(rsDetails(3))
		    end if
		    rsDetails.Close 
	    else
		    sAttributeListValue = trim(rs(1))
	    end if
	    rs.Close 
	end if 'if trim(iAttributeList)<>"" then

	set ndRoot = objDom.createElement("Root")
	objDom.appendChild ndRoot
	
	set ndAttributeList = objDom.createElement("Attribute")
	ndAttributeList.setAttribute "ID",iAttributeList
	ndAttributeList.setAttribute "Name",sAttributeListValue
	ndRoot.appendChild ndAttributeList
	
	Response.ContentType = "text/xml"
	Response.Write objDom.xml
%>
