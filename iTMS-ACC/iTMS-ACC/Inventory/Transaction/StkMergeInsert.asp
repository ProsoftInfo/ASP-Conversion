<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	StkMergeInsert.asp
	'Module Name				:	Inventory (Stock Management Stock Transfer)
	'Author Name				:	Ragavendran R
	'Created On					:	May 26, 2011
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/InternalReceiptInsert.asp"-->
<!--#include file="../../include/mrsIssueInsertCommon.asp"-->
<!--#include file="../../include/getCurrentDate.asp"-->
<%

Dim objIssueDOM,objRcptDOM,objFSO,rsMergeTemp
Dim ndIssRoot,ndRcptRoot,ndRcptItem,ndIssItem
Dim iMergedEntryNo,iMergedItemCode,iMergedClassCode,sArrFinPeriod,sFinFromDate,sFinToDate
Dim iStkCnt,iToMergeItemCode,iToMergeClassCode,sQuery
sArrFinPeriod = Split(Session("FinPeriod"),":")
sFinFromDate = "01/04/"& sArrFinPeriod(0)
sFinToDate = "31/03/"& sArrFinPeriod(1)


set objIssueDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRcptDOM = Server.CreateObject("Microsoft.XMLDOM")
set objFSO = CreateObject("Scripting.FileSystemObject")
set rsMergeTemp = Server.CreateObject("ADODB.Recordset")

if objFSO.FileExists(Server.MapPath("../temp/transaction/mrsIssueData"&Session.SessionID&".xml")) then
    objIssueDOM.load(Server.MapPath("../temp/transaction/mrsIssueData"&Session.SessionID&".xml"))
end if
if objFSO.FileExists(Server.MapPath("../temp/transaction/ReceiptLotData"&Session.SessionID&".xml")) then
    objRcptDOM.load(Server.MapPath("../temp/transaction/ReceiptLotData"&Session.SessionID&".xml"))
end if


With rsMergeTemp
    .CursorLocation = 3
    .CursorType = 3
    .Source = "Select isNull(Max(MergedEntryNo)+1,1) from INV_T_MergedItemLedger"
    .ActiveConnection = CON
    .Open 
End With
if not rsMergeTemp.EOF then
    iMergedEntryNo = rsMergeTemp(0)
end if
rsMergeTemp.Close 

con.beginTrans

set ndRcptRoot = objRcptDOM.documentElement
if ndRcptRoot.hasChildNodes() then
    sExp = "//ITEM"
    set ndRcptItem = ndRcptRoot.selectnodes(sExp)
    if ndRcptItem.length>0 then
        iMergedItemCode = ndRcptItem.Item(0).Attributes.getNamedItem("ITMCODE").value
        iMergedClassCode = ndRcptItem.Item(0).Attributes.getNamedItem("CLACODE").value
    end if
end if

set ndIssRoot = objIssueDOM.documentElement
if ndIssRoot.hasChildNodes() then
    sExp = "//ItemDetail"
    set ndIssItem = ndIssRoot.selectnodes(sExp)
    if ndIssItem.length>0 then
        For iStkCnt = 1 to ndIssItem.length
            iToMergeItemCode = ndRcptItem.Item(0).Attributes.getNamedItem("ItemCode").value
            iToMergeClassCode = ndRcptItem.Item(0).Attributes.getNamedItem("CLACODE").value
            
            sQuery ="Insert into INV_T_MergedItemLedger Select "& iMergedEntryNo &",*,"& iMergedItemCode &","& iMergedClassCode &" from "&_
            " INV_T_ITEMLEDGER Where itemCode ="& iToMergeItemCode &" and ClassificationCode = "& iToMergeClassCode &" and "&_
            " (Convert(datetime,TransactionDate,103) >= Convert(datetime,'"& sFinFromDate &"',103) and Convert(datetime,TransactionDate,103) <= Convert(datetime,'"& sFinToDate &"',103))"
            Response.Write sQuery
            con.execute squery
        next
    end if
    
end if


MrsIssueInsert

''Added by ragav on Aug 17,2013 for Use Existing Pack Number while merging item 
''begin
sUseExistingPackNum = "Y"
''end

sCallFrom = "PA"
CreateInternalReceipt(getCurrentDate())

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	response.end
	response.clear
	con.CommitTrans
end if
response.redirect "../Master/ITEMLISTENTRY.ASP?ACTN=ME"
%>