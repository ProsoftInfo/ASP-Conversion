<%@ EnableSessionState=true%> 
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
Dim sQuery,sItemCode
sItemCode = Request.QueryString("ItemCode")
sQuery = "Update INV_M_ItemMaster set ItemThumbNailPic=NULL,ItemBlowupPic=NULL where ItemCode = "& sItemCode   
con.execute sQuery
%>
