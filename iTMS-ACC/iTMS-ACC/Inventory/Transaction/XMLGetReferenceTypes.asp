<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetReferenceTypes.asp
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Dec 27,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<%
	Dim objDOM,rsObj,sQuery,sRefAppCode,sRefCode,ndRoot,ndChild
	set rsObj = Server.CreateObject("ADODB.Recordset")
	set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	sRefAppCode = Request.QueryString("RefAppCode")
    sRefCode = Request.QueryString("RefCode")
	
	set ndRoot = objDOM.CreateElement("Root")
	objDOM.appendChild ndRoot
	
	sQuery = "Select ReferenceEntryNo,ReferenceName from VW_ReferenceTypes where RefCodeNo = "& sRefCode &" and RefApplicationCode = "& sRefAppCode
	rsObj.Open sQuery,con
	if not rsObj.EOF then
		do while not rsObj.EOF 
			set ndChild = objDOM.createElement("Ref")
				ndChild.setAttribute "No",rsObj(0)
				ndChild.setAttribute "Name",rsObj(1)
			ndRoot.appendChild ndChild
			rsObj.MoveNext
		loop
	end if
	rsObj.Close 
	Response.ContentType = "text/xml"
	Response.Write objDOM.xml
%>