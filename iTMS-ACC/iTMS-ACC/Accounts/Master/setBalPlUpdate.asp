<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	setBalPLUpdate.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Manohar Prabhu .R
	'Created On					:	Nov 27 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<%
	Dim sQuery,sNewVal,iAccHead,objrs,iCtr,sOrgID,sCatCode
	Set objrs = Server.CreateObject("ADODB.RecordSet")
	iCtr = 0
	
	sOrgID = Request.Form("selUnitID")
	sCatCode = Request.Form("selCategory")
	
	Con.BeginTrans
	sQuery = "Select Distinct AccountHead From Acc_R_OrgGLAccountHead Where  "&_
			 "SubString(AccountsGroupCode,1,2) = '"&sCatCode&"' and  "&_
			 "OUDefinitionID = '"&sOrgID&"' "
	With objrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set objrs.ActiveConnection = Nothing
	Do While Not objrs.EOF
		sNewVal = Request.Form("txtAcc"&Trim(objrs(0)))
		sNewVal = Replace(sNewVal,"'"," ")
		sQuery = "Update Acc_R_OrgGLAccountHead  Set AccHeadAlias = '"&sNewVal&"' Where "&_
				 "AccountHead = "&objrs(0)&" and OUDefinitionID = '"&sOrgID&"' "
				 
		Response.Write sQuery &"<br><br>"
		Con.Execute sQuery
		objrs.MoveNext
	Loop
	objrs.Close
	
	Con.CommitTrans
	'Response.End
	Response.Redirect "ConfigureBalanceSheet.ASP"
	
	
%>