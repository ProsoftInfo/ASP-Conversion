<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XmlGetClassification.asp
	'Module Name				:	Inventory Closing Stock
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
    Dim objDOM,rsObj,sLevel,sQuery,ndRoot,ndSubLevel
    Dim sCategory,sSubCategory
    
    set objDOM =Server.CreateObject("Microsoft.XMLDOM")
    set rsObj = Server.CreateObject("ADODB.Recordset")
    
    sLevel = Request.QueryString("Level")
    sCategory = Request.QueryString("Category")
    sSubCategory = Request.QueryString("SubCategory")
    if Trim(sLevel)="1" then
        sQuery = "Select GroupName,GroupCode from Inv_M_Classification where GroupCode<>ParentGroup and ParentGroup = "& sSubCategory
    else
        sQuery = "Select GroupName,GroupCode from Inv_M_Classification where GroupCategory = '"& sCategory &"'"
    end if
    set ndRoot = objDOM.createElement("Root")
        objDOM.appendChild ndRoot
    if Trim(sQuery)<>"" then
        rsObj.Open sQuery,con
        if not rsObj.EOF then
            do while not rsObj.EOF 
                set ndSubLevel = objDOM.createElement("Level")
                    ndSubLevel.setAttribute "GroupName",rsObj(0)
                    ndSubLevel.setAttribute "GroupCode",rsObj(1)
                    ndRoot.appendChild ndSubLevel
                rsObj.MoveNext 
            loop
        end if
        rsObj.Close 
    end if
    
    Response.Clear
    Response.ContentType = "text/xml"
    Response.Write objDOM.xml
%>