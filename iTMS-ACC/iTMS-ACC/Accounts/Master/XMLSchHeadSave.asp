<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLSchHeadSave.asp
	'Module Name				:	Purchase(Receipts)
	'Author Name				:	SRIDEVI PRIYA A
	'Created On					:	March 14,2003
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
	set objRs  = server.CreateObject("adodb.recordset")
	
	Dim OrgId,SchID,LevelID,Level1ID,Level2ID,Level1Name,Level2Name,ModeType,AccHead
	Dim AccHeadName,FinYear,ComputeMode,sExp,TempNode,splt
	Dim SchSubId,SchSubSubID,DispAccDescr,SubText,SubSubText,EntryNo,Hierachy,sInsDate
	
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	
	set Root=oDOM.documentElement
	'	sAgent = Root.Attributes.Item(0).nodeValue
	sExp = "//Details"
	AccHead = 0
	Set TempNode = Root.selectNodes(sExp)
	OrgId =  TempNode.Item(0).Attributes.getNamedItem("OrgID").value
	SchID =  TempNode.Item(0).Attributes.getNamedItem("SchID").value
	LevelID =  TempNode.Item(0).Attributes.getNamedItem("LevelID").value
	splt = split( TempNode.Item(0).Attributes.getNamedItem("Level1ID").value,"-")
	Level1ID = splt(0) 
	'Level1ID =  TempNode.Item(0).Attributes.getNamedItem("Level1ID").value
	Level2ID =  TempNode.Item(0).Attributes.getNamedItem("Level2ID").value
	Level1Name =  TempNode.Item(0).Attributes.getNamedItem("Level1Name").value
	Level2Name =  TempNode.Item(0).Attributes.getNamedItem("Level2Name").value
	ModeType =  TempNode.Item(0).Attributes.getNamedItem("ModeType").value
	AccHead =  TempNode.Item(0).Attributes.getNamedItem("AccHead").value
	AccHeadName =  TempNode.Item(0).Attributes.getNamedItem("AccHeadName").value
	FinYear =  TempNode.Item(0).Attributes.getNamedItem("FinYear").value	
	ComputeMode =  TempNode.Item(0).Attributes.getNamedItem("ComputeMode").value
	sInsDate = TempNode.Item(0).Attributes.getNamedItem("InsDate").value
	
	Set oDOM = Nothing
	If ModeType = "A" then
		DispAccDescr = "Y"
	Else
		DispAccDescr = "N"
	End If
	'****************Edition Part**************************************88
	Dim Lvl2Id,spltlev2
	
		
	If LevelID = 1 And Level2ID <> "A" Then 
		spltlev2 = split(Level2ID,"-") 
		sqlstring = "Delete From Acc_T_ScheduleACDetail where  (ScheduleID = "&SchID&") AND (ScheduleSubID = "&Level1ID&") and (schedulesubsubid = "&spltlev2(0)&")"
		con.execute sqlstring
		
		sqlstring =	"UPDATE Acc_M_SchdSetupSubHeads SET SubHeadingName = '"&replace(Level2Name,"'","''")&"', EntryType='"&ModeType&"', computemode ='"&ComputeMode&"' where (ScheduleID = "&SchID&") AND (ScheduleSubID = "&Level1ID&") and (schedulesubsubid = "&spltlev2(0)&")"		    
		con.execute sqlstring 
		sqlstring =	"UPDATE Acc_M_SchdSetupSubHeads SET SubHeadingName = '"&replace(Level1Name,"'","''")&"',computemode='"&computemode&"' where (ScheduleID = "&SchID&") AND (ScheduleSubID = "&Level1ID&") and (schedulesubsubid = 0)"		    
		con.execute sqlstring 
		If AccHeadName <>"" Then
			sqlstring = "Select Max(Hierachy) + 1 From Acc_M_SchdSetupSubHeads Where ScheduleID = "&SchID&" and ScheduleSubID = "&Level1ID&" and FinYear = '"&FinYear&"'"
			With Objrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring 
				.Open 
			End With
			Hierachy = objrs(0)
			objrs.Close
		
			sqlstring = "SELECT ISNULL(MAX(EntryNumber), 0) + 1 AS EntryNo FROM dbo.Acc_T_ScheduleACDetail"
			With objrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring 
				.Open 
			End With
			If Not objrs.EOF Then EntryNo = objrs(0)
			objrs.Close   
			sqlstring = "Insert INTO Acc_T_ScheduleACDetail (EntryNumber, ScheduleID, ScheduleSubID, ScheduleSubSubID, Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode,FinYear, ScheduleSubHeadValue, ComputeMode,AsOnDate ) " &_
						"Values("&EntryNo&","&SchID&","&Level1ID&",'"&spltlev2(0)&"',"&Hierachy&",'"&OrgId&"',0,"&AccHead&",'"&FinYear&"',0,'"&ComputeMode&"',Convert(datetime,'"&sInsDate&"',103))" 
		'	Response.Write sqlstring
			con.execute (sqlstring)
		
		
		End If
		
		
	ElseIf LevelID = 0 And Level1ID <>"A" Then  
		sqlstring =	"UPDATE Acc_M_SchdSetupSubHeads SET SubHeadingName = '"&replace(Level1Name,"'","''")&"', Entrytype='"&ModeType&"', computemode='"&ComputeMode&"' where (ScheduleID = "&SchID&") AND (ScheduleSubID = "&Level1ID&") and (schedulesubsubid = 0)"		    
		con.execute sqlstring 
		sqlstring = "Delete From Acc_T_ScheduleACDetail  where  (ScheduleID = "&SchID&") AND (ScheduleSubID = "&Level1ID&") and (schedulesubsubid = 0)"
		con.execute sqlstring
		If ModeType <>"N" Then
			sqlstring = "Select Max(Hierachy) + 1 From Acc_M_SchdSetupSubHeads Where ScheduleID = "&SchID&" and ScheduleSubID = "&Level1ID&" and FinYear = '"&FinYear&"'"
			With Objrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring 
				.Open 
			End With
			Hierachy = objrs(0)
			objrs.Close
		
			sqlstring = "SELECT ISNULL(MAX(EntryNumber), 0) + 1 AS EntryNo FROM dbo.Acc_T_ScheduleACDetail"
			With objrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring 
				.Open 
			End With
			If Not objrs.EOF Then EntryNo = objrs(0) 
			objrs.Close   
			If AccHeadName <>"" Then
			sqlstring = "Insert INTO Acc_T_ScheduleACDetail (EntryNumber, ScheduleID, ScheduleSubID, ScheduleSubSubID, Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode,FinYear, ScheduleSubHeadValue, ComputeMode,AsOnDate ) " &_
						"Values("&EntryNo&","&SchID&","&Level1ID&",0,"&Hierachy&",'"&OrgId&"',0,"&AccHead&",'"&FinYear&"',0,'"&ComputeMode&"',Convert(datetime,'"&sInsDate&"',103))" 
			Else
			sqlstring = "Insert INTO Acc_T_ScheduleACDetail (EntryNumber, ScheduleID, ScheduleSubID, ScheduleSubSubID, Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode,FinYear, ScheduleSubHeadValue, ComputeMode,AsOnDate ) " &_
						"Values("&EntryNo&","&SchID&","&Level1ID&",0,"&Hierachy&",'"&OrgId&"',0,0,'"&FinYear&"',0,'"&ComputeMode&"',Convert(datetime,'"&sInsDate&"',103))" 			
			End If
			Response.Write sqlstring
			con.execute (sqlstring)
		End If
		
	Else
	'******************Insertion Part*****************
	
	
	
	
	Con.BeginTrans
	If Level1ID = "A" Then
		Dim Discr,iMode
	   If LevelID <> 0 and Level2ID = "A" Then
		 Discr ="N"
		 iMode = "N"
	   Else 
	    Discr = DispAccDescr
	    iMode = ModeType  
	   End If	
		sqlstring = "SELECT ISNULL(MAX(ScheduleSubID) + 1, 1) AS ShSubID FROM dbo.Acc_M_SchdSetupSubHeads Where (ScheduleID = " & SchID &" )"
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con	
			.Source = sqlstring 
			.Open 
		End With
		SchSubId = objrs(0)
		SchSubSubID = 0
		objrs.Close  
		sqlstring = "INSERT INTO Acc_M_SchdSetupSubHeads (ScheduleID, ScheduleSubID, ScheduleSubSubID, SubHeadingName, DisplayACHeadDescr, Hierachy, EntryType, FinYear, ComputeMode)"&_
					"VALUES("&SchID&","&SchSubId&","&SchSubSubID&",'"&replace(Level1Name,"'","''")&"','"&Discr&"',2,'"&iMode&"','"&FinYear&"','"&ComputeMode&"')"
		con.execute (sqlstring)
		Level1ID=SchSubId			
	End If	
	If Level2ID = "A" and LevelID = "1" Then
		sqlstring = "Select Max(Hierachy) + 1 From Acc_M_SchdSetupSubHeads Where ScheduleID = "&SchID&" and FinYear = '"&FinYear&"'"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close
		sqlstring = "SELECT ISNULL(MAX(ScheduleSubSubID), 0) + 1 AS SubSubId FROM dbo.Acc_M_SchdSetupSubHeads Where (ScheduleID = " & SchID & ") AND (ScheduleSubID = " & Level1ID & ")"		
		With objrs 
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		If Not objrs.EOF then SchSubSubID = objrs(0)
		objrs.Close 
		sqlstring = "INSERT INTO Acc_M_SchdSetupSubHeads (ScheduleID, ScheduleSubID, ScheduleSubSubID, SubHeadingName, DisplayACHeadDescr, Hierachy, EntryType, FinYear, ComputeMode) " &_
					"Values("&SchID&","&Level1ID&","&SchSubSubID&",'"&replace(Level2Name,"'","''")&"','"&DispAccDescr&"',"&Hierachy&" ,'"&ModeType&"','"&FinYear&"','"&ComputeMode&"') "
					
		'Response.Write sqlstring
		con.execute (sqlstring)		
	End If
	If ModeType <> "N" Then 'and LevelID = "1" then
		sqlstring = "Delete From Acc_T_ScheduleACDetail where  (ScheduleID = "&SchID&") AND (ScheduleSubID = "&Level1ID&") and (schedulesubsubid = "&SchSubSubID&")"
		con.execute sqlstring
		sqlstring = "Select Max(Hierachy) + 1 From Acc_M_SchdSetupSubHeads Where ScheduleID = "&SchID&" and ScheduleSubID = "&Level1ID&" and FinYear = '"&FinYear&"'"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		Hierachy = objrs(0)
		objrs.Close
		
		sqlstring = "SELECT ISNULL(MAX(EntryNumber), 0) + 1 AS EntryNo FROM dbo.Acc_T_ScheduleACDetail"
		With objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring 
			.Open 
		End With
		If Not objrs.EOF Then EntryNo = objrs(0)
		objrs.Close   
		If ModeType = "A" Then
		sqlstring = "Insert INTO Acc_T_ScheduleACDetail (EntryNumber, ScheduleID, ScheduleSubID, ScheduleSubSubID, Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode,FinYear, ScheduleSubHeadValue, ComputeMode,AsOnDate ) " &_
					"Values("&EntryNo&","&SchID&","&Level1ID&","&SchSubSubID&","&Hierachy&",'"&OrgId&"',0,"&AccHead&",'"&FinYear&"',0,'"&ComputeMode&"',Convert(datetime,'"&sInsDate&"',103)) " 
		Else
		sqlstring = "Insert INTO Acc_T_ScheduleACDetail (EntryNumber, ScheduleID, ScheduleSubID, ScheduleSubSubID, Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode,FinYear, ScheduleSubHeadValue, ComputeMode,AsOnDate ) " &_
					"Values("&EntryNo&","&SchID&","&Level1ID&","&SchSubSubID&","&Hierachy&",'"&OrgId&"',0,0,'"&FinYear&"',0,'"&ComputeMode&"',Convert(datetime,'"&sInsDate&"',103)) " 
		End If
		'Response.Write sqlstring
		con.execute (sqlstring)
	End If

	Con.CommitTrans
	'Con.RollbackTrans
	End If

%>
