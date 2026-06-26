<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLTDSSave.asp
	'Module Name				:	Accounts(TDS)
	'Author Name				:	Kumar K A
	'Created On					:	Dec 28,2006
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
	Dim HeadName,GroupID,GroupName,HeadID,AccountHead,HeadDetails,HeadCode
	Dim Hierarchy,CreatedOn,Mode
	Dim sOrgID,PrevGroupName,PrevGroupID
	
	set objRs  = server.CreateObject("adodb.recordset")
	set objrs1 = server.CreateObject("adodb.recordset")
	
	PrevGroupName = ""
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	Con.BeginTrans
	set Root=oDOM.documentElement
	sExp = "//Schedule"
	'AccHead = 0
	Set TempNode = Root.selectNodes(sExp)
	For i = 0 to TempNode.Length-1
		HeadName = TempNode.Item(i).Attributes.getNamedItem("HeadName").value
		GroupID =  TempNode.Item(i).Attributes.getNamedItem("GroupID").value
		HeadID =  TempNode.Item(i).Attributes.getNamedItem("HeadID").value
		'HeadName =  TempNode.Item(i).Attributes.getNamedItem("HeadName").value
		AccountHead =  TempNode.Item(i).Attributes.getNamedItem("AccountHead").value
		Mode =  TempNode.Item(i).Attributes.getNamedItem("Mode").value
		HeadDetails =  TempNode.Item(i).Attributes.getNamedItem("HeadDetails").value
		GroupName =  TempNode.Item(i).Attributes.getNamedItem("GroupName").value
		Hierarchy =  TempNode.Item(i).Attributes.getNamedItem("Hierarchy").value
		CreatedOn =  TempNode.Item(i).Attributes.getNamedItem("CreatedOn").value
		sOrgID =  TempNode.Item(i).Attributes.getNamedItem("sOrgID").value 
		HeadCode = TempNode.Item(i).Attributes.getNamedItem("HeadCode").value 
		'Response.Write GroupID 
		If GroupID = "A" Then			
		
			With objrs1
				.CursorLocation = 3
				.CursorType =3
				.ActiveConnection = con
				.Source = "Select GroupID from ACC_M_TDSGroup Where isNull(GroupName,'') = '"& lcase(GroupName)&"' "
				.Open 
			End With
			'Response.Write "<p>"&objrs1.Source 
			If Not objrs1.EOF Then
				GroupID = Trim(objrs1(0))
			Else
				sqlstring= "Select isNull(max(GroupID),0)+1 from ACC_M_TDSGroup"  
				With objrs 
					.CursorLocation = 3
					.CursorType =3
					.ActiveConnection = con
					.Source = sqlstring
					.Open 
				End With
				GroupID = objrs(0) 
			
				If i = 0 Then PrevGroupID = GroupID 
				objrs.Close  
				If  PrevGroupName <> GroupName Then
						sqlstring = " Insert into ACC_M_TDSGroup (GroupID,GroupName,OUDefinitionID,GroupACHeadCode)"&_   
								    " Values ("&GroupID&",'"&GroupName&"','"&sOrgID&"','')"   
						con.execute sqlstring
						PrevGroupName = GroupName 
					
				Else
				GroupID = PrevGroupID 
				End If
			End If	
			
		End IF	'If objrs1.EOF Then
		'objrs1.Close 
	
	If HeadID = "A" Then
	'Response.Write GroupID 
		sqlstring = "Select isnull(max(GroupHeadID),0)+1 from ACC_M_TDSHeadComputation where GroupID="&GroupID
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		HeadID = objrs(0)
		objrs.Close   
	
		sqlstring = "Select isnull(max(Herarchy),0)+1 from ACC_M_TDSHeadComputation where GroupID="&GroupID
		With objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring
			.Open 
		End With
		
		If Hierarchy<>"" Then Hierarchy = objrs(0) 
			objrs.Close 
			sqlstring = "Select * From ACC_M_TDSHeadComputation"
			With objrs 
				.CursorLocation = 3
				.CursorType = 3
				.LockType = 3
				.ActiveConnection = con
				.Source = sqlstring 
				.Open 
				.AddNew 
			End With
			objrs("Groupid") = GroupID 
			objrs("GroupHeadID") = HeadID  
			objrs("GroupHeadName") = HeadName 
			objrs("computeMode") = Mode
			objrs("AcHeadCode") = HeadCode 
			objrs("Herarchy") = Hierarchy
			objrs.Update 
			objrs.Close  	
		End If
		PrevGroupID = GroupID 
	Next
	
	Con.CommitTrans
	'Con.RollbackTrans
	
	  
	
	
%>
