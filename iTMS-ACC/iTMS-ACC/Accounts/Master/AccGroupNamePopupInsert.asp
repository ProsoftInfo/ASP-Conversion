<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccGroupNamePopupInsert.asp	
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 16,2010
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
<HTML>
<head>
<base target="_self">
</head>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT>
</SCRIPT>

<%
dim dcrs,sQuery,Temp
dim sClassName,sGcode,sParentCode,arrpGroupCode,sGroupValue
dim sCategoryCode
sClassName = UCase(replace (trim(Request.Form("txtClassCreateName")),"'","''"))
sGroupValue = trim(Request.Form("GCode"))

'Response.Write "sGroupValue = "& sGroupValue

if  len(trim(sGroupValue)) >2 then
	arrpGroupCode = split(sGroupValue,":")
	sParentCode = Mid(sGroupValue,3,len(sGroupValue))
	
else
	sParentCode = "GRP"
	sCategoryCode= Mid(sGroupValue,1,2)
end if

Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sQuery = "SELECT AccountsGroupName FROM Acc_M_AccountGroups WHERE LOWER(AccountsGroupName) = '" & lcase(Trim(sClassName)) & "'"
    with dcrs
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
    If Not dcrs.EOF Then
%>
	<BODY  onLoad = "msgbox(' Account Group Name Already Exists','N')">
<%
	
	Response.End
    End If
    dcrs.Close

    If Trim(sParentCode) = "GRP" Then
			sQuery = "Select count(AccountsGroupCode)+1 from Acc_M_AccountGroups where AccountsParentGroup=AccountsGroupCode and GroupCategory ='"&sCategoryCode&"'"
			Response.Write sQuery
			dcrs.Open sQuery,con
		
			if len(dcrs(0))=1 then
				sGCode = sCategoryCode&"0"&dcrs(0)
			else
				sGCode = sCategoryCode&dcrs(0)
			end if
			
			dcrs.Close

		sQuery = "Insert into Acc_M_AccountGroups (AccountsGroupCode,AccountsGroupName,AccountsParentGroup,ChildCount,GroupCategory,GroupHierarchy,OldHierarchy) values('"&sGCode&"','"&Trim(sClassName)&"','"&sGCode&"',0,'"&sCategoryCode&"','"&sGCode&"','"&sGCode&"')"	
		Response.Write sQuery
        con.Execute sQuery
%>
	<BODY  onLoad = "msgbox('Account Group has been Created Successfully','Y')">
<%
    Else
		'*********
		sQuery = "Select childcount + 1,GroupCategory from Acc_M_AccountGroups where AccountsGroupCode = '"&sParentCode&"'"
		
		dcrs.Open sQuery,con
		
		If Not dcrs.EOF then
		  If dcrs(0) < 10 then			
			 sGcode = "0"&trim(dcrs(0))
		  Else
			sGcode = dcrs(0)
		  End if	
			sCategoryCode=dcrs(1)
		End if
			
		sGcode = trim(sParentCode)&trim(sGcode)
			
		sQuery = "Insert into Acc_M_AccountGroups (AccountsGroupCode,AccountsGroupName,AccountsParentGroup,ChildCount,GroupCategory,GroupHierarchy,OldHierarchy) values('"&sGcode&"','"&Trim(sClassName)&"','"&sParentCode&"',0,'"&sCategoryCode&"','"&sGCode&"','"&sGCode&"')"	
		con.Execute sQuery
			
		sQuery = "Update Acc_M_AccountGroups SET Childcount = childcount + 1 where AccountsGroupCode = '"&sParentCode&"' "
		con.execute sQuery
        '*********
%>
	<BODY  onLoad = "msgbox('Group has been Created Successfully','Y')">
<%
    End If


con.close
set con = nothing
Response.Redirect "comAccountGroupTreePopup.asp"
%>
</HTML>