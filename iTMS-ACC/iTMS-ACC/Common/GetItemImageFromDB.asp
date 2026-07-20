<!-- #include File="../include/DatabaseConnection.asp" -->
<%
response.expires = 0
response.buffer = 0

Dim rs
Dim sItemCode 

set rs = Server.CreateObject("ADODB.Recordset")
sItemCode= Request.QueryString("ID")
rs.open "Select ItemBlowupPic from INV_M_ItemMaster where ItemCode = "& sItemCode,con
if not rs.eof then
response.contenttype="image"
response.binarywrite rs(0)    
end if
rs.close
%>
