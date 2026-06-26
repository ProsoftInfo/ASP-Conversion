<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GroupEditPopupUpdate.asp	
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
<html>
<head>
<base target="_self">
</head>
<SCRIPT LANGUAGE=javascript>
</SCRIPT>

<%
dim sQuery
dim sClassName,sGcode,sGroupFlag,sTargetName

sClassName = replace (trim(Request.Form("txtClassEditName")),"'","''")
sGcode = trim(Request.Form("GCode"))
sGroupFlag = trim(Request.Form("GroupFlag"))
select case  sGroupFlag
case "G"
	sGcode=MID (sGcode,3)
	sTargetName="AccGroupCreationMain.asp"
	sQuery = "update Acc_M_AccountGroups set AccountsGroupName='"&sClassName&"' where AccountsGroupCode='"&sGcode&"'"
case "A"
	sTargetName="ANALGroupCreationMain.asp"
	sQuery = "update Acc_M_AnalyticalGroup set AHGroupName='"&sClassName&"' where AHGroupCode='"&sGcode&"'"
case "C"
	sTargetName="CCGroupCreationMain.asp"
	sQuery = "update Acc_M_CostCenterGroup set CCGroupName='"&sClassName&"' where CCGroupCode='"&sGcode&"'"
end select 
con.Execute sQuery
%>
	<BODY BGCOLOR="#336699" onLoad = "msgbox('Group has been Updated Successfully','Y')">
	<form name="formname">
		<input type="hidden" name="hTargetPage" value="<%=sTargetName%>">
	</form>
<%
con.close
set con = nothing
Response.Redirect "comAccountGroupTreePopup.asp"
%>
</HTML>