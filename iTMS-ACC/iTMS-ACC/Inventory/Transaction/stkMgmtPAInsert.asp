<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtPAInsert.asp
	'Module Name				:	Inventory (Stock Management Physical Adjustment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 28, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 22,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	stkMgmtPAEntry.asp
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
Dim objFSOIssue,objFSORcpt
con.beginTrans

set objFSOIssue = CreateObject("Scripting.FileSystemObject")
set objFSORcpt = CreateObject("Scripting.FileSystemObject")

if objFSOIssue.FileExists(Server.MapPath("../Temp/Transaction/mrsIssueData"&Session.SessionID&".xml")) then
    MrsIssueInsert
end if

if objFSORcpt.FileExists(Server.MapPath("../Temp/Transaction/ReceiptLotData"&Session.SessionID&".xml")) then
sCallFrom = "PA"
    CreateInternalReceipt(getCurrentDate())
end if 



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
	
	if objFSOIssue.FileExists(Server.MapPath("../Temp/Transaction/mrsIssueData"&Session.SessionID&".xml")) then
        objFSOIssue.DeleteFile(Server.MapPath("../Temp/Transaction/mrsIssueData"&Session.SessionID&".xml"))
    end if

    if objFSORcpt.FileExists(Server.MapPath("../Temp/Transaction/ReceiptLotData"&Session.SessionID&".xml")) then
        objFSORcpt.DeleteFile(Server.MapPath("../Temp/Transaction/ReceiptLotData"&Session.SessionID&".xml"))
    end if 

end if

con.close
set con = nothing




response.redirect "../Master/ITEMLISTENTRY.ASP?ACTN=PA"

%>

