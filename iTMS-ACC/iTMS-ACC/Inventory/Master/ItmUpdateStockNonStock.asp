<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmUpdateStockNonStock.asp
	'Module Name				:	Inventory (Item Updation)
	'Author Name				:	Ragav
	'Created On					:	May 29,2014
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<%
Dim sItemCode,sStockNonStock,sQuery
sItemCode = Request("ItemCode")
sStockNonStock = Request("Stock")
    con.begintrans	    
    
    sQuery = "Update Inv_M_ItemMaster set StockNonStock = '"& sStockNonStock &"' where ItemCode = "& sItemCode
    Response.Write "<p>"& sQuery
    con.execute sQuery

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if con.Errors.count <> 0 then
		dim iErrCounter
		con.RollbackTrans
		for iErrCounter=0 to con.Errors.count
			Response.Write con.Errors(iErrCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.End
	    Response.Clear 
	 	con.CommitTrans
	end if

	con.close
	set con = nothing
Response.Redirect "ITEMLISTENTRY.ASP?ACTN=A"
%>
