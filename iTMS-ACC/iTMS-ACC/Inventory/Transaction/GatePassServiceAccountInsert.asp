<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GatePassServiceAccountInsert.asp
	'Module Name				:	Purchase(Transaction)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 28, 2003
	'Modified By				:	KUMAR K A
	'Modified By				:	Ragavendran R
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	GatePassServiceAccountPop.asp
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
<%

dim newxml,RootNode,ItemNode

Set newxml = Server.CreateObject("Microsoft.XMLDOM")

newxml.async = false
newxml.load(Request)

dim iGPNo,sSql, iEntryNo,dtReceivedOn
Set RootNode = newxml.documentElement
iGPNo = trim(RootNode.Attributes.getNamedItem("GPNO").Value)

con.beginTrans



if RootNode.hasChildNodes then
	for each ItemNode in RootNode.ChildNodes
		iEntryNo = trim(ItemNode.Attributes.getNamedItem("EntryNo").Value)
		dtReceivedOn = trim(ItemNode.Attributes.getNamedItem("ReceivedOn").Value)
	
		if dtReceivedOn <> "" then
			sSql = "UPDATE FORGATEPASSDETAILS SET MaterialRcvdOn = convert(datetime,'" & dtReceivedOn & "',103),MaterialRcvd = 'Y'	WHERE gatepassno = " & iGPNo & " and EntryNo = " & iEntryNo
	'		Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql 
		end if 
	next
end if 'if RootNode.hasChildNodes then


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
end if

con.close
set con = nothing

%>

