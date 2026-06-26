<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLSchBSHeadSave.asp
	'Module Name				:	Accounts(P&L)
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
	dim objrs,objrs1,objrs2,sqlstring,Root
	Dim SubID,SubSubID
	set objRs  = server.CreateObject("adodb.recordset")
	Dim OrgId,BSHeadID,LevelID,Level1ID,Level2ID,Level1Name,Level2Name,ModeType,AccHead
	Dim AccHeadName,FinYear,ComputeMode,sExp,TempNode,SchName,sInsDate
	Dim BSSubId,BSSubSubID,DispAccDescr,SubText,SubSubText,EntryNo,Hierachy
	
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	
	set Root=oDOM.documentElement
	sExp = "//Details"
	AccHead = 0
	Set TempNode = Root.selectNodes(sExp)
	SchName = TempNode.Item(0).Attributes.getNamedItem("SchName").value
	OrgId =  TempNode.Item(0).Attributes.getNamedItem("OrgID").value
	BSHeadID =  TempNode.Item(0).Attributes.getNamedItem("SchID").value
	LevelID =  TempNode.Item(0).Attributes.getNamedItem("LevelID").value
	Level1ID =  TempNode.Item(0).Attributes.getNamedItem("Level1ID").value
	Level2ID =  TempNode.Item(0).Attributes.getNamedItem("Level2ID").value
	Level1Name =  TempNode.Item(0).Attributes.getNamedItem("Level1Name").value
	Level2Name =  TempNode.Item(0).Attributes.getNamedItem("Level2Name").value
	ModeType =  TempNode.Item(0).Attributes.getNamedItem("ModeType").value
	AccHead =  TempNode.Item(0).Attributes.getNamedItem("AccHead").value
	AccHeadName =  TempNode.Item(0).Attributes.getNamedItem("AccHeadName").value
	FinYear =  TempNode.Item(0).Attributes.getNamedItem("FinYear").value	
	ComputeMode =  TempNode.Item(0).Attributes.getNamedItem("ComputeMode").value
	Hierachy =  TempNode.Item(0).Attributes.getNamedItem("Hierachy").value
	sInsDate = TempNode.Item(0).Attributes.getNamedItem("InsDate").value
	
	SubID = Split(Level1ID,"-") 
	Set oDOM = Nothing
	If ModeType = "A" then
		DispAccDescr = "Y"
	Else
		DispAccDescr = "N"
	End If
	
	Con.BeginTrans
	SubID = Split(Level1ID,",")
	SubSubID = Split(Level2ID,",") 
	If LevelID = "0" and Level1ID<>"A" and BSHeadID <>"A" Then 
		sqlstring = "Update ACC_M_BSSetupHeads Set BSHeading='"&SchName&"' where BSHeadID='"&BSHeadID&"'"    
		con.execute sqlstring
		sqlstring = "Update Acc_M_BSSetupSubHeads Set BSSubHeadingName='"&Level1Name&"',EntryType='"&ModeType&"',ComputeMode='"&ComputeMode&"',Hierachy='"&Hierachy&"' where BSHeadID="&BSHeadID&" and BSSubID="&SubID(0)&" and BSSubSubID = 0"  
		con.execute sqlstring 
		Level1ID = SubID(0) 
		Level2ID = 0
		BSSubId = SubID(0) 
	ElseIF LevelID = "1" and Level2ID <>"A" Then 
		sqlstring = "Update ACC_M_BSSetupHeads Set BSHeading='"&SchName&"' where BSHeadID='"&BSHeadID&"'"    
		con.execute sqlstring
		sqlstring = "Update Acc_M_BSSetupSubHeads Set BSSubHeadingName='"&Level1Name&"',EntryType='N',ComputeMode='"&ComputeMode&"' where BSHeadID = '"&BSHeadID&"' and BSSubID='"&SubID(0)&"' and BSSubSubID = 0"  
		con.execute sqlstring 
		sqlstring = "Update Acc_M_BSSetupSubHeads Set BSSubHeadingName='"&Level2Name&"',EntryType='"&ModeType&"',ComputeMode='"&ComputeMode&"',Hierachy='"&Hierachy&"' where BSHeadID = '"&BSHeadID&"' and BSSubID='"&SubID(0)&"' and BSSubSubID = '"&SubSubID(1)&"'"  
		con.execute sqlstring 
		Level1ID = SubID(0)  
		Level2ID = SubSubID(1) 
		BSSubId= SubID(0)  
	Else
	If BSHeadID = "A" Then
		sqlstring = "Select IsNull(Max(BSHeadID),0)+1 From Acc_M_BSSetupHeads"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		BSHeadID = objrs(0)
		objrs.Close 
		sqlstring = "Select IsNull(Max(Hierarchy),0)+1 From Acc_M_BSSetupHeads"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close 
		sqlstring = "INSERT INTO Acc_M_BSSetupHeads (BSHeadID, BSHeading, Hierarchy, FinYear)"&_
					"VALUES("&BSHeadID&",'"&SchName&"',"&Hierachy&",'"&FinYear&"')"
		con.execute (sqlstring)
	End If
	'Response.Write NewID
	'Response.End 
	If Level1ID = "A" Then
		dim NewID
		If Level2ID = "A" and Level2Name <>"" Then 
			NewID = "N"
		Else
			NewID = ModeType 
		End If
		sqlstring = "Select IsNull(Max(Hierachy),0) + 1 From Acc_M_BSSetupSubHeads Where BSHeadID = "&BSHeadID&" and FinYear = '"&FinYear&"'"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close
	
		sqlstring = "SELECT ISNULL(MAX(BSSubID) , 0)+ 1 AS BSSubID FROM dbo.Acc_M_BSSetupSubHeads Where (BSHeadID = " & BSHeadID &" )"
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con	
			.Source = sqlstring 
			.Open 
		End With
		BSSubId = objrs(0)
		BSSubSubID = 0
		objrs.Close  
		sqlstring = "INSERT INTO Acc_M_BSSetupSubHeads (BSHeadID, BSSubID, BSSubSubID, BSSubHeadingName, Hierachy, EntryType, FinYear, ComputeMode)"&_
					"VALUES("&BSHeadID&","&BSSubId&","&BSSubSubID&",'"&Level1Name&"','"&Hierachy&"','"&NewID&"','"&FinYear&"','"&ComputeMode&"')"
		con.execute (sqlstring)
		Level1ID=BSSubId	
		SubID = Split(Level1ID,",")
		'Level2ID = BSSubSubID 
	End If
	
	If Level2ID = "A" and LevelID = "1" Then
		sqlstring = "Select Max(Hierachy) + 1 From Acc_M_BSSetupSubHeads Where BSHeadID = "&BSHeadID&" and FinYear = '"&FinYear&"'"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close
		 
		sqlstring = "SELECT ISNULL(MAX(BSSubSubID), 0) + 1 AS BSSubId FROM dbo.Acc_M_BSSetupSubHeads Where (BSHeadID = " & BSHeadID & ") AND (BSSubID = " & SubID(0) & ")"		
		'Response.Write sqlstring 
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		
		If Not objrs.EOF then Level2ID = objrs(0) Else Level2ID= BSSubSubID  
		objrs.Close 
		sqlstring = "INSERT INTO Acc_M_BSSetupSubHeads (BSHeadID, BSSubId, BSSubSubID, BSSubHeadingName, Hierachy, EntryType, FinYear, ComputeMode) " &_
					"Values("&BSHeadID&","&SubID(0)&","&Level2ID&",'"&Level2Name&"',"&Hierachy&" ,'"&ModeType&"','"&FinYear&"','"&ComputeMode&"') "
	'	Response.Write sqlstring 
		con.execute (sqlstring)		
		BSSubId = SubID(0) 
		Level1ID = SubID(0)   
	End If
	End If
	If ModeType = "A" or ModeType="D" Then
		If Level2ID = "A" Then Level2ID = 0
		sqlstring = "Select Max(Hierachy) + 1 From Acc_M_BSSetupSubHeads Where BSHeadID ='"&BSHeadID&"' and BSSubID='"&Level1ID&"' and FinYear='"&FinYear&"'"
		'Response.Write sqlstring  
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close
		sqlstring = "Delete From Acc_T_BSACDetail Where BSHeadID="&BSHeadID&" and BSSubID="&Level1ID&" and BSSubSubID="&Level2ID&" and ScheduleID=0 and ScheduleSubID=0 and ScheduleSubSubID=0" 
		con.execute sqlstring	
		
		If ModeType= "D" Then
		sqlstring = "Insert INTO Acc_T_BSACDetail (BSHeadID, BSSubID, BSSubSubID, Hierarchy, OrganisationCode, ApplicableACHeadCode, DisplayACHeadDescr, FinYear, BSSubHeadValue,ComputeMode, ScheduleID, ScheduleSubID, ScheduleSubSubID,AsOnDate) " &_
					"Values("&BSHeadID&","&Level1ID&","&Level2ID&","&Hierachy&",'"&OrgId&"',0,'N','"&FinYear&"',0,'"&ComputeMode&"',0,0,0,Convert(Datetime,'"&sInsDate&"',103)) " 
		Else
		sqlstring = "Insert INTO Acc_T_BSACDetail (BSHeadID, BSSubID, BSSubSubID, Hierarchy, OrganisationCode, ApplicableACHeadCode, DisplayACHeadDescr, FinYear, BSSubHeadValue,ComputeMode, ScheduleID, ScheduleSubID, ScheduleSubSubID,AsOnDate) " &_
					"Values("&BSHeadID&","&Level1ID&","&Level2ID&","&Hierachy&",'"&OrgId&"',"&AccHead&",'N','"&FinYear&"',0,'"&ComputeMode&"',0,0,0,Convert(Datetime,'"&sInsDate&"',103)) " 
		End If
		'Response.Write sqlstring
		con.execute (sqlstring)
		BSSubId = Level1ID 
	End If
	
	If ModeType = "S" Then
		If Level2ID = "A" Then Level2ID = 0
		'check here
		sqlstring = "Delete From Acc_T_BSACDetail Where BSHeadID="&BSHeadID&" and BSSubID="&Level1ID&" and BSSubSubID="&Level2ID&"" ' and ScheduleID='"&ScheduleID&"' and ScheduleSubID='"&ScheduleSubID&"' and ScheduleSubSubID='"&ScheduleSubSubID&"'" 
		Response.Write sqlstring
		con.execute sqlstring
		Dim Description,ScheduleID,ScheduleSubID,ScheduleSubSubID,recno
		Dim SchSubHeadValue 
		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		sName = "SchedBSBrkSubHeads"
		sMod = "Acc"
		oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
		
		set Root=oDOM.documentElement
		sExp = "//SchDetails"
		Set TempNode = Root.selectNodes(sExp)
		'Response.Write TempNode.length 
		For recno = 0 To TempNode.length-1
			Description=TempNode.Item(recno).Attributes.getNamedItem("Description").value
			ScheduleID=TempNode.Item(recno).Attributes.getNamedItem("ScheduleID").value
			ScheduleSubID=TempNode.Item(recno).Attributes.getNamedItem("ScheduleSubID").value	
			ScheduleSubSubID=TempNode.Item(recno).Attributes.getNamedItem("ScheduleSubSubID").value

			sqlstring = "SELECT ScheduleSubHeadValue FROM dbo.Acc_T_ScheduleACDetail" &_
					" WHERE (ScheduleID="&ScheduleID&") AND (ScheduleSubID="&ScheduleSubID&") AND (ScheduleSubSubID="&ScheduleSubSubID&")"
			With objrs 
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring 
				.Open 
			End With
			IF Not objrs.EOF Then
				SchSubHeadValue = objrs(0)
			Else
				SchSubHeadValue = 0
			End IF
			
			objrs.Close 
			sqlstring = "Select isnull(Max(Hierachy),0) + 1 From Acc_M_BSSetupSubHeads Where BSHeadID = "&BSHeadID&" and BSSubID = "&BSSubId&" and FinYear = '"&FinYear&"'"
			'Response.Write sqlstring
				With Objrs
					.CursorLocation = 3
					.CursorType = 3
					.ActiveConnection = con
					.Source = sqlstring 
					.Open 
				End With
			Hierachy = objrs(0)
			objrs.Close
			
			'sqlstring = "Insert INTO Acc_T_BSACDetail (BSHeadID, BSSubID, BSSubSubID, Hierarchy, OrganisationCode, ApplicableACHeadCode, DisplayACHeadDescr, FinYear, BSSubHeadValue,ComputeMode, ScheduleID, ScheduleSubID, ScheduleSubSubID,AsOnDate)" &_
			'				"Values("&BSHeadID&","&BSSubId&","&Level2ID&","&Hierachy&",'"&OrgId&"',0,'N','"&FinYear&"','"&SchSubHeadValue&"','"&ComputeMode&"','"&ScheduleID&"','"&ScheduleSubID&"','"&ScheduleSubSubID&"',Convert(Datetime,'"&sInsDate&"',103))" 
			
			sqlstring = "Insert INTO Acc_T_BSACDetail (BSHeadID, BSSubID, BSSubSubID, Hierarchy, OrganisationCode, ApplicableACHeadCode, DisplayACHeadDescr, FinYear, BSSubHeadValue,ComputeMode, ScheduleID, ScheduleSubID, ScheduleSubSubID,AsOnDate)" &_
						"Values("&BSHeadID&","&Level1ID&","& Level2ID &","&Hierachy&",'"&OrgId&"',0,'N','"&FinYear&"','"&SchSubHeadValue&"','"&ComputeMode&"','"&ScheduleID&"','"&ScheduleSubID&"','"&ScheduleSubSubID&"',Convert(Datetime,'"&sInsDate&"',103))" 
							
			Response.Write sqlstring
			con.execute (sqlstring)
		Next
	End If		

	Response.Clear 
	Con.CommitTrans
	'Con.RollbackTrans

	
%>
