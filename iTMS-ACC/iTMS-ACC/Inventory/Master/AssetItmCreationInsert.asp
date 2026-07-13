<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AssetItmCreationInsert.asp
	'Module Name				:	Inventory (Asset Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 03, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	AssetItmCreationEntry.asp
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
<!--#include File="../../include/AsstItemCreationInsertCommon.asp" -->
<%
	Dim sTempArr,sCallFrom,nAssetCode,sSql
	
	sTempArr = Request.QueryString("sPassValue")
	sCallFrom = trim(Split(sTempArr,":")(0))
	nAssetCode = trim(Split(sTempArr,":")(1))
	
	con.begintrans
	
	'' To Call the Inventory Insert Common Function
    InvItemCreationInsert
	
	Response.Clear
	
	if con.Errors.count <> 0 then
		dim iErrCounter
		con.RollbackTrans
		for iErrCounter=0 to con.Errors.count
			Response.Write con.Errors(iErrCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else
		sSql = "UPDATE FAR_T_ASSETDETAILS SET SENTTOINVENTORY = 'S' WHERE ASSETDESCID = " & trim(nAssetCode) & ""
		con.Execute sSql
		
		con.RollbackTrans
		Response.End 
		con.CommitTrans
		
		If sCallFrom = "INV" Then
			Response.Redirect("ASSETITMLISTENTRY.ASP")
		Else
			Response.Redirect("../../Fixedassets/TRANSACTION/FIXEDASSETS.ASP")
		End If
		
	end if

	con.close
	set con = nothing
%>
