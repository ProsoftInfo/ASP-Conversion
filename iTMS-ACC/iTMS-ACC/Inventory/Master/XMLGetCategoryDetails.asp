<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLGetCategoryDetails.asp	
	'Module Name				:	Inventory 
	'Author Name				:	Ragavendran R
	'Created On					:	Jul 22,2011
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim oDOM,rsTemp,ndRoot,ndChild
Dim sCatCode,sCatName,sCatShortName,sQuery

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set rsTemp = Server.CreateObject("ADODB.Recordset")

set ndRoot= oDOM.createElement("Root")
oDOM.appendChild ndRoot

    sQuery = "Select CategoryCode,CategoryName,CategoryShortName from INV_M_CLASSIFICATIONCATEGORY"
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        do while not rsTemp.EOF 
            sCatCode = rsTemp(0)
            sCatName = rsTemp(1)
            sCatShortName = rsTemp(2)
            set ndChild = oDOM.createElement("Category")
            ndChild.setAttribute "CATEGORYCODE",sCatCode
            ndChild.setAttribute "CATEGORYNAME",sCatName 
            ndChild.setAttribute "CATEGORYSHORTNAME",sCatShortName 
            ndRoot.appendChild ndChild
            rsTemp.MoveNext
        loop
    end if
    rsTemp.Close 
    
    Response.ContentType = "text/xml"
    Response.Write oDOM.xml
%>
