<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLTDSFormulaSave.asp
	'Module Name				:	Accounts(TDS)
	'Author Name				:	Kumar K A
	'Created On					:	Jan 24,2006
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
	Dim GroupID,TDSHeadID,Formula
	set objRs  = server.CreateObject("adodb.recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	Con.BeginTrans
	set Root=oDOM.documentElement
	sExp = "//TDS"
	Set TempNode = Root.selectNodes(sExp)
	For i = 0 to TempNode.Length-1
		GroupID =  TempNode.Item(i).Attributes.getNamedItem("GroupID").value
		TDSHeadID =  TempNode.Item(i).Attributes.getNamedItem("TDSHeadID").value
		Formula =  TempNode.Item(i).Attributes.getNamedItem("Formula").value
		If Formula <> "" Then 
			sqlstring = "Update ACC_M_TDSHeadComputation Set ComputeFormula ='"&Formula&"' where GroupID="&GroupID&" and GroupHeadID="&TDSHeadID 
			con.execute sqlstring
		End If
		
	Next
	
	Con.CommitTrans
	'Con.RollbackTrans
	
	  
	
	
%>
