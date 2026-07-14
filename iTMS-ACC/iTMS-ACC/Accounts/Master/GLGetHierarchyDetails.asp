<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLGetHierarchyDetails.asp	
	'Module Name				:	Accounts (Master)
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
Dim objRs
dim sGroupName,sGroupCode,iHierarchy,iCounter,sGroup,oDOM,ndRoot,ndChild
Dim sParentCode,sQuery

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.Recordset")

sParentCode = Request.QueryString("GCode")

Set ndRoot = oDOM.createElement("Root")
oDOM.appendChild ndRoot

	if Len(sParentCode)=2 then
		sQuery="select AccountsGroupCode,AccountsGroupName,GroupHierarchy from Acc_M_AccountGroups where AccountsParentGroup=AccountsGroupCode and GroupCategory='"&sParentCode&"' ORDER BY Grouphierarchy "
	else
		sQuery="select AccountsGroupCode,AccountsGroupName,GroupHierarchy from Acc_M_AccountGroups where  AccountsParentGroup<>AccountsGroupCode and AccountsParentGroup='"& mid(sParentCode,3)&"' ORDER BY Grouphierarchy"
	end if
	
	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	if not objRs.EOF then
		iCounter=0
		while not objRs.EOF
			sGroupCode=objRs("AccountsGroupCode")
			sGroupName=objRs("AccountsGroupName")
			iHierarchy=objRs("GroupHierarchy")
			
			iCounter=CInt(iCounter)+1
			if iCounter=1 then 
				sGroup=iHierarchy
			else
				sGroup=sGroup&","&iHierarchy
			end if	
			
			set ndChild = oDOM.createElement("AccHead")
			ndChild.setAttribute "SNo",iCounter
			ndChild.setAttribute "GName",sGroupName
			ndChild.setAttribute "Hierarchy",cint(Right(iHierarchy,2))
			ndRoot.appendChild ndChild

			objRs.MoveNext
		wend
	end if 'if not objRs.EOF then
	objRs.close
	ndRoot.setAttribute "Group",sGroup
	ndRoot.setAttribute "Counter",iCounter
	
	Response.ContentType = "text/xml"
	Response.Write oDOM.xml
%>                                            
