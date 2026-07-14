<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
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
	Dim	GroupID,ComputeFormula,iSplt,sql,iSplt1,HeadID
	set objRs  = server.CreateObject("adodb.recordset")
	Dim sOrgID,PrevGroupName,PrevGroupID
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	Con.BeginTrans
	set Root=oDOM.documentElement
	sExp = "//Schedule"
	'AccHead = 0
	Set TempNode = Root.selectNodes(sExp)
	For i = 0 to TempNode.Length-1
		GroupID = TempNode.Item(i).Attributes.getNamedItem("GroupID").value
		HeadID = TempNode.Item(i).Attributes.getNamedItem("HeadID").value
		ComputeFormula =  TempNode.Item(i).Attributes.getNamedItem("ComputeFormula").value
		iSplt = Split(ComputeFormula,",") 
		'Response.Write HeadID
		If UBound(iSplt) > 0 Then
			iSplt1 = Split(iSplt(1),"#")  
		Else
			iSplt1 = Split(iSplt(0),"#")
		End If
		sql = "Update Acc_M_TDSHeadComputation set ComputeFormula='"&ComputeFormula&"' where GroupID="&GroupID&" and GroupHeadID="&Cint(HeadID)
		'Response.Write sql
		con.execute sql
	Next 
	Con.CommitTrans
	'Con.RollbackTrans
	
	  
	
	
%>
