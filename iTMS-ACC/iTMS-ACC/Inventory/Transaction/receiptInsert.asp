<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	receiptInsert.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 29,2013
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptItemEntry.asp
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
<!--#include file="../../include/NoSeries.asp"-->
<!--#include file="../../include/PurchaseInvFnc.asp"-->
<!--#include file="../../include/mrsStatus.asp"-->
<!--#include file="../../include/NoSeriesCommonFunctions.asp"-->
<!--#include file="../../include/InventoryAccountingUpdate.asp"-->
<%
'XML DOM Variables
Dim Root,objfs,sSql,rsTemp

Set objfs = CreateObject("Scripting.FileSystemObject")

dim newxml,iAccountedBy,sOrgID,sReceiptFor,iRcptNo

Set newxml = Server.CreateObject("Microsoft.XMLDOM")
set rsTemp = Server.CreateObject("ADODB.Recordset")

newxml.async = false
newxml.load(Request)

newxml.save(Server.MapPath("../temp/transaction/InventoryAcc_"&Session.SessionID&".xml"))

set root = newxml.documentElement
iAccountedBy = getUserid
sReceiptFor = 22
iRcptNo =  root.getAttribute("RECNO")
sOrgID = root.getAttribute("UNIT")

con.begintrans
			
InvAccountUpdate sReceiptFor,iRcptNo,sOrgID,iAccountedBy

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	Response.End 
	
	con.CommitTrans
	Response.clear
	
	if objfs.FileExists(server.MapPath("../temp/transaction/PACKING"&Session.SessionID&".xml")) then
		objfs.DeleteFile server.MapPath("../temp/transaction/PACKING"&Session.SessionID&".xml")
	end if
	if objfs.FileExists(server.MapPath("../temp/transaction/SELLING"&Session.SessionID&".xml")) then
		objfs.DeleteFile server.MapPath("../temp/transaction/SELLING"&Session.SessionID&".xml")
	end if
	if objfs.FileExists(server.MapPath("../temp/transaction/RECEIPTEX"&Session.SessionID&".xml")) then
		objfs.DeleteFile server.MapPath("../temp/transaction/RECEIPTEX"&Session.SessionID&".xml")
	end if
	if objfs.FileExists(server.MapPath("../temp/transaction/InventoryAcc_"&Session.SessionID&".xml")) then
		objfs.DeleteFile server.MapPath("../temp/transaction/InventoryAcc_"&Session.SessionID&".xml")
	end if

end if

con.close
set con = nothing
%>
