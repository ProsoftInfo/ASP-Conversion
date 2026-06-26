<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetClassAttribute.asp
%>
<!-- #include File="DatabaseConnection.asp" -->
<%	
    Dim rsObj
    Dim sQuery,iClassCode
    
    iClassCode = Request("ClassCode")
    
    set rsObj = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select * from INV_M_ItemTypeAttributes where ClassificationCode = "& iClassCode
    rsObj.open sQuery,con
    if not rsObj.eof then
        Response.write "Y"
    else
        Response.write "N"
    end if
    rsObj.close
%>
