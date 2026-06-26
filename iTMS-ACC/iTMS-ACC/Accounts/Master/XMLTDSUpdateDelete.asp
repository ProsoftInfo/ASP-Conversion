<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLcomputeSave.asp
	'Module Name				:	Accounts(TDS)
	'Author Name				:	Kumar K A
	'Created On					:	Dec 28,2006
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	
	'							:
	'Connects To				:	TDSComputationDetails.ASP
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
<%

	Dim oDOM,sMod,sName,sDesgDir
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	sMod=Request("Mod")
	sName=Request("Name")
	sDesgDir = trim(Request("ToDir"))
	'Response.Write "<p> sDesgDir = " & sDesgDir 	
	oDOM.async = false
	oDOM.load(Request)
	oDOM.Save server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	set oDOM=nothing
	
%>
<%
	Dim objrs,objrs1,objrs2,sqlstring,Root,i,TempNode,sExp
	Dim id,GroupHeadID,GroupName,GroupHeadName,ComputeMode,GroupID
	Dim AcHeadCode,Herarchy,AccHeadName,sql
	set objRs  = server.CreateObject("adodb.recordset")
	Dim sOrgID,PrevGroupName,PrevGroupID
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	Con.BeginTrans
	set Root=oDOM.documentElement
	sExp = "//TDS"
	Set TempNode = Root.selectNodes(sExp)
		id = TempNode.Item(0).Attributes.getNamedItem("id").value
	If id = "1" Then
		GroupHeadID = TempNode.Item(0).Attributes.getNamedItem("GroupHeadID").value
		GroupName =  TempNode.Item(0).Attributes.getNamedItem("GroupName").value
		GroupHeadName = TempNode.Item(0).Attributes.getNamedItem("GroupHeadName").value
		ComputeMode = TempNode.Item(0).Attributes.getNamedItem("ComputeMode").value
		AcHeadCode = TempNode.Item(0).Attributes.getNamedItem("AcHeadCode").value
		Herarchy = TempNode.Item(0).Attributes.getNamedItem("Herarchy").value
		AccHeadName = TempNode.Item(0).Attributes.getNamedItem("AccHeadName").value
		GroupID = TempNode.Item(0).Attributes.getNamedItem("GroupID").value
		sql = "Update Acc_M_TDSHeadComputation Set GroupHeadName='"&GroupHeadName&"', ComputeMode='"&ComputeMode&"',AcheadCode="&AcHeadCode&",Herarchy="&Cint(Herarchy)&" where GroupID="&GroupID&" and GroupHeadID="&GroupHeadID
		con.execute sql
	End If
	If id = "2" Then
		GroupHeadID = TempNode.Item(0).Attributes.getNamedItem("GroupHeadID").value
		GroupName =  TempNode.Item(0).Attributes.getNamedItem("GroupName").value
		GroupHeadName = TempNode.Item(0).Attributes.getNamedItem("GroupHeadName").value
		ComputeMode = TempNode.Item(0).Attributes.getNamedItem("ComputeMode").value
		AcHeadCode = TempNode.Item(0).Attributes.getNamedItem("AcHeadCode").value
		Herarchy = TempNode.Item(0).Attributes.getNamedItem("Herarchy").value
		AccHeadName = TempNode.Item(0).Attributes.getNamedItem("AccHeadName").value
		GroupID = TempNode.Item(0).Attributes.getNamedItem("GroupID").value
		If GroupHeadID <>"" Then
		sql = "Delete From Acc_M_TDSHeadComputation where GroupID="&GroupID&" and GroupHeadID="&GroupHeadID
		con.execute sql
		Else
		sql = "Delete From Acc_M_TDSGroup where GroupID="&GroupID
		con.execute sql
		sql = "Delete From Acc_M_TDSHeadComputation where GroupID="&GroupID
		con.execute sql
		End If
	End If
	If id = "3" Then
		
		
		
			
	
	
	End If
	Con.CommitTrans
	'Con.RollbackTrans
	
	  
	
	
%>
