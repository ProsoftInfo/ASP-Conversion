<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLSchPLHeadSave.asp
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
	Dim OrgId,PLHeadID,LevelID,Level1ID,Level2ID,Level1Name,Level2Name,ModeType,AccHead
	Dim AccHeadName,FinYear,ComputeMode,sExp,TempNode,SchName
	Dim PLSubId,PLSubSubID,DispAccDescr,SubText,SubSubText,EntryNo,Hierachy,sInsDate
	
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	
	set Root=oDOM.documentElement
	sExp = "//Details"
	AccHead = 0
	Set TempNode = Root.selectNodes(sExp)
	SchName = TempNode.Item(0).Attributes.getNamedItem("SchName").value
	OrgId =  TempNode.Item(0).Attributes.getNamedItem("OrgID").value
	PLHeadID =  TempNode.Item(0).Attributes.getNamedItem("SchID").value
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
	If LevelID = "0" and Level1ID<>"A" and PLHeadID <>"A" Then 
		sqlstring = "Update ACC_M_PLSetupHeads Set PLHeading='"&SchName&"' where PLHeadID='"&PLHeadID&"'"    
		con.execute sqlstring
		sqlstring = "Update Acc_M_PLSetupSubHeads Set PLSubHeadingName='"&Level1Name&"',EntryType='"&ModeType&"',ComputeMode='"&ComputeMode&"',Hierachy='"&Hierachy&"' where PLHeadID="&PLHeadID&" and PLSubID="&SubID(0)&" and PLSubSubID = 0"  
		con.execute sqlstring 
		Level1ID = SubID(0) 
		Level2ID = 0
		PLSubId = SubID(0) 
	ElseIF LevelID = "1" and Level2ID <>"A" Then 
		sqlstring = "Update ACC_M_PLSetupHeads Set PLHeading='"&SchName&"' where PLHeadID='"&PLHeadID&"'"    
		con.execute sqlstring
		sqlstring = "Update Acc_M_PLSetupSubHeads Set PLSubHeadingName='"&Level1Name&"',EntryType='N',ComputeMode='"&ComputeMode&"' where PLHeadID = '"&PLHeadID&"' and PLSubID='"&SubID(0)&"' and PLSubSubID = 0"  
		con.execute sqlstring 
		sqlstring = "Update Acc_M_PLSetupSubHeads Set PLSubHeadingName='"&Level2Name&"',EntryType='"&ModeType&"',ComputeMode='"&ComputeMode&"',Hierachy='"&Hierachy&"' where PLHeadID = '"&PLHeadID&"' and PLSubID='"&SubID(0)&"' and PLSubSubID = '"&SubSubID(1)&"'"  
		con.execute sqlstring 
		Level1ID = SubID(0)  
		Level2ID = SubSubID(1) 
		PLSubId= SubID(0)  
	Else
	If PLHeadID = "A" Then
		sqlstring = "Select IsNull(Max(PLHeadID),0)+1 From Acc_M_PLSetupHeads"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		PLHeadID = objrs(0)
		objrs.Close 
		sqlstring = "Select IsNull(Max(Hierarchy),0)+1 From Acc_M_PLSetupHeads"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close 
		sqlstring = "INSERT INTO Acc_M_PLSetupHeads (PLHeadID, PLHeading, Hierarchy, FinYear)"&_
					"VALUES("&PLHeadID&",'"&SchName&"',"&Hierachy&",'"&FinYear&"')"
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
		sqlstring = "Select IsNull(Max(Hierachy),0) + 1 From Acc_M_PLSetupSubHeads Where PLHeadID = "&PLHeadID&" and FinYear = '"&FinYear&"'"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close
	
		sqlstring = "SELECT ISNULL(MAX(PLSubID) , 0)+ 1 AS PLSubID FROM dbo.Acc_M_PLSetupSubHeads Where (PLHeadID = " & PLHeadID &" )"
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con	
			.Source = sqlstring 
			.Open 
		End With
		PLSubId = objrs(0)
		PLSubSubID = 0
		objrs.Close  
		sqlstring = "INSERT INTO Acc_M_PLSetupSubHeads (PLHeadID, PLSubID, PLSubSubID, PLSubHeadingName, Hierachy, EntryType, FinYear, ComputeMode)"&_
					"VALUES("&PLHeadID&","&PLSubId&","&PLSubSubID&",'"&Level1Name&"','"&Hierachy&"','"&NewID&"','"&FinYear&"','"&ComputeMode&"')"
		con.execute (sqlstring)
		Level1ID=PLSubId	
		SubID = Split(Level1ID,",")
		'Level2ID = PLSubSubID 
	End If
	
	If Level2ID = "A" and LevelID = "1" Then
		sqlstring = "Select Max(Hierachy) + 1 From Acc_M_PLSetupSubHeads Where PLHeadID = "&PLHeadID&" and FinYear = '"&FinYear&"'"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close
		 
		sqlstring = "SELECT ISNULL(MAX(PLSubSubID), 0) + 1 AS PLSubId FROM dbo.Acc_M_PLSetupSubHeads Where (PLHeadID = " & PLHeadID & ") AND (PLSubID = " & SubID(0) & ")"		
		'Response.Write sqlstring 
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		
		If Not objrs.EOF then Level2ID = objrs(0) Else Level2ID= PLSubSubID  
		objrs.Close 
		sqlstring = "INSERT INTO Acc_M_PLSetupSubHeads (PLHeadID, PLSubId, PLSubSubID, PLSubHeadingName, Hierachy, EntryType, FinYear, ComputeMode) " &_
					"Values("&PLHeadID&","&SubID(0)&","&Level2ID&",'"&Level2Name&"',"&Hierachy&" ,'"&ModeType&"','"&FinYear&"','"&ComputeMode&"') "
	'	Response.Write sqlstring 
		con.execute (sqlstring)		
		PLSubId = SubID(0) 
		Level1ID = SubID(0)   
	End If
	End If
	If ModeType = "A" or ModeType="D" Then
		If Level2ID = "A" Then Level2ID = 0
		sqlstring = "Select Max(Hierachy) + 1 From Acc_M_PLSetupSubHeads Where PLHeadID ='"&PLHeadID&"' and PLSubID='"&Level1ID&"' and FinYear='"&FinYear&"'"
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
		sqlstring = "Delete From Acc_T_PLACDetail Where PLHeadID="&PLHeadID&" and PLSubID="&Level1ID&" and PLSubSubID="&Level2ID&" and ScheduleID=0 and ScheduleSubID=0 and ScheduleSubSubID=0" 
		con.execute sqlstring	
		
		If ModeType= "D" Then
		sqlstring = "Insert INTO Acc_T_PLACDetail (PLHeadID, PLSubID, PLSubSubID, Hierarchy, OrganisationCode, ApplicableACHeadCode, DisplayACHeadDescr, FinYear, PLSubHeadValue,ComputeMode, ScheduleID, ScheduleSubID, ScheduleSubSubID,AsOnDate) " &_
					"Values("&PLHeadID&","&Level1ID&","&Level2ID&","&Hierachy&",'"&OrgId&"',0,'N','"&FinYear&"',0,'"&ComputeMode&"',0,0,0,Convert(datetime,'"&sInsDate&"',103)) " 
		Else
		sqlstring = "Insert INTO Acc_T_PLACDetail (PLHeadID, PLSubID, PLSubSubID, Hierarchy, OrganisationCode, ApplicableACHeadCode, DisplayACHeadDescr, FinYear, PLSubHeadValue,ComputeMode, ScheduleID, ScheduleSubID, ScheduleSubSubID,AsOnDate) " &_
					"Values("&PLHeadID&","&Level1ID&","&Level2ID&","&Hierachy&",'"&OrgId&"',"&AccHead&",'N','"&FinYear&"',0,'"&ComputeMode&"',0,0,0,Convert(datetime,'"&sInsDate&"',103)) " 
		End If
		'Response.Write sqlstring
		con.execute (sqlstring)
		PLSubId = Level1ID 
	End If
	
	If ModeType = "S" Then
		If Level2ID = "A" Then Level2ID = 0
		sqlstring = "Delete From Acc_T_PLACDetail Where PLHeadID="&PLHeadID&" and PLSubID="&Level1ID&" and PLSubSubID="&Level2ID&"" ' and ScheduleID='"&ScheduleID&"' and ScheduleSubID='"&ScheduleSubID&"' and ScheduleSubSubID='"&ScheduleSubSubID&"'" 
		'Response.Write sqlstring
		con.execute sqlstring
		Dim Description,ScheduleID,ScheduleSubID,ScheduleSubSubID,recno
		Dim SchSubHeadValue 
		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		sName = "SchedPLBrkSubHeads"
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
			sqlstring = "Select isnull(Max(Hierachy),0) + 1 From Acc_M_PLSetupSubHeads Where PLHeadID = "&PLHeadID&" and PLSubID = "&PLSubId&" and FinYear = '"&FinYear&"'"
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
		
			
					
			sqlstring = "Insert INTO Acc_T_PLACDetail (PLHeadID, PLSubID, PLSubSubID, Hierarchy, OrganisationCode, ApplicableACHeadCode, DisplayACHeadDescr, FinYear, PLSubHeadValue,ComputeMode, ScheduleID, ScheduleSubID, ScheduleSubSubID,AsOnDate)" &_
							"Values("&PLHeadID&","&PLSubId&","&Level2ID&","&Hierachy&",'"&OrgId&"',0,'N','"&FinYear&"','"&SchSubHeadValue&"','"&ComputeMode&"','"&ScheduleID&"','"&ScheduleSubID&"','"&ScheduleSubSubID&"',Convert(datetime,'"&sInsDate&"',103))" 
			'Response.Write sqlstring
			con.execute (sqlstring)
		Next
	End If		

	Con.CommitTrans
	'Con.RollbackTrans
	
	  
	
	
%>
