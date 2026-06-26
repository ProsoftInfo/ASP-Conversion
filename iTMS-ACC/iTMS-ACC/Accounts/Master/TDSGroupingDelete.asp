<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TDSGroupingDelete.asp
	'Module Name				:	Accounts-TDS (Master Amedment)
	'Author Name				:	Kumar K.A.
	'Created On					:	January 17 2007
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	
	'Procedures/Functions Used	:
	'Internal Variables			:

	'Database					:	ITMS_Test
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
	Dim sQuery,iTdsID,sCallType,sDelFrom
	
	sDelFrom = Trim(Request.Form("hDelFrom"))
	
	If sDelFrom = "M" Then
		iTdsID = Request.QueryString("sGrpID")
	Else
		iTdsID = Request.Form("GroupName") 'Request.Form("selGPName")
		sCallType = Trim(Request.form("hRequest"))
	End IF
	
	con.BeginTrans
	
	sQuery = "Update ACC_M_TDSGroup Set Useable = 'N' Where GroupID IN ("&iTdsID&") "
	Con.Execute sQuery
	
	con.CommitTrans
	
	If sDelFrom = "M" Then
		Response.Redirect("TDSGroupingSetup.asp")
	Else
		Response.Redirect("TDSGroupingSetup.asp?CallType="&sCallType)
	End IF
	
%>