<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccGroupOrderUpdatePopup.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Dec 09,2010
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	
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
<!--#include file="../../include/sessionVerify.asp"-->
<%
dim objRs,sQuery,saGroup,saOrder,iCounter
dim sParentCode,sGroupCode,iGroupCount,sGroupOrder
dim sTempCode

Set objRs = Server.CreateObject("ADODB.RecordSet")

sParentCode=Request.Form("hParentGroupCode")
sGroupCode=Request.Form("hGroupCode")
iGroupCount=Request.Form("hGroupCount")
sGroupOrder=Request.Form("txtOrder")

saOrder=Split(sGroupOrder,",")
saGroup=Split(sGroupCode,",")

for iCounter=1 to iGroupCount
	if len(CStr(iCounter))=1 then
		sTempCode=mid(trim(saGroup(iCounter-1)),1,len(trim(saGroup(iCounter-1)))-2)&"0"&trim(saOrder(iCounter-1))
	else
		sTempCode=mid(trim(saGroup(iCounter-1)),1,len(trim(saGroup(iCounter-1)))-2)&"0"&trim(saOrder(iCounter-1))
	end if	
	
	sQuery="update Acc_M_AccountGroups  set Grouphierarchy='"&sTempCode&"'+ substring(Grouphierarchy,len('"&trim(saGroup(iCounter-1))&"')+1,len(Grouphierarchy)) "&_
	"where substring(oldhierarchy,1,len('"&trim(saGroup(iCounter-1))&"'))='"&trim(saGroup(iCounter-1))&"'"
	con.Execute sQuery
next
sQuery="update Acc_M_AccountGroups  set oldhierarchy=Grouphierarchy"
con.Execute sQuery

%>
<HTML>
<HEAD>
</HEAD>
<body>
<form name="formname">
</form>
</body>
</HTML>
<%
Response.Redirect "comAccountGroupTreePopup.asp"
%>
