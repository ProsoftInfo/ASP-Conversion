<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLSchBrkHeadSave.asp
	'Module Name				:	Purchase(Receipts)
	'Author Name				:	K A Kumar
	'Created On					:	Dec 27,2006
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	
	'							:
	'Connects To				:	AddSchedBrkSubHeads.asp
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

	Dim oDOM,sMod,sName,sDesgDir,AccHead
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	sMod=Request("Mod")
	sName=Request("Name")
	sDesgDir = trim(Request("ToDir"))
	'Response.Write "<p> sDesgDir = " & sDesgDir 	
	oDOM.async = false
	oDOM.load(Request)
	oDOM.Save server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	'oDOM.Save server.MapPath("../temp/transaction/Test1.xml")
	set oDOM=nothing
	
%>

<%
	Dim objrs,objrs1,objrs2,sqlstring,Root,splt
	set objRs  = server.CreateObject("adodb.recordset")
	Dim ID,OrgID,ScheduleID,ScheduleSubID,ScheduleSubSubID,Level2ID,Level3ID
	Dim HeadName,BreakUpHeadName,BreakUpSubHead,BreakupId,Hierarchy
	Dim BreakupSubId,BreakupSubSubId,Mode,FinYear,ComputeMode
	Dim DisplayACHeadDescr,DataEntry,EntryNumber,sAccHeadID
	Dim sExp,TempNode,sInsDate
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	
	set Root=oDOM.documentElement
	sExp = "//Schedule"
	AccHead = 0
	Set TempNode = Root.selectNodes(sExp)
	
	ID =  TempNode.Item(0).Attributes.getNamedItem("ID").value
	OrgID =  TempNode.Item(0).Attributes.getNamedItem("OrgID").value
	Level2ID = TempNode.Item(0).Attributes.getNamedItem("Level2ID").value
	splt = Split(Level2ID,",")
	Level2ID = splt(0)  
	Level3ID = TempNode.Item(0).Attributes.getNamedItem("Level3ID").value
	ScheduleID =  TempNode.Item(0).Attributes.getNamedItem("ScheduleID").value
	ScheduleSubID =  TempNode.Item(0).Attributes.getNamedItem("ScheduleSubID").value
	ScheduleSubSubID =  TempNode.Item(0).Attributes.getNamedItem("ScheduleSubSubID").value
	HeadName =  TempNode.Item(0).Attributes.getNamedItem("HeadName").value
	BreakUpHeadName =  TempNode.Item(0).Attributes.getNamedItem("BreakUpHeadName").value
	BreakUpSubHead =  TempNode.Item(0).Attributes.getNamedItem("BreakUpSubHead").value
	BreakupId =  TempNode.Item(0).Attributes.getNamedItem("BreakupId").value
	splt = Split(BreakupId,",")
	BreakupId = splt(0)   
	BreakupSubId =  TempNode.Item(0).Attributes.getNamedItem("BreakupSubId").value
	BreakupSubSubId =  TempNode.Item(0).Attributes.getNamedItem("BreakupSubSubId").value
	Mode =  TempNode.Item(0).Attributes.getNamedItem("Mode").value	
	Hierarchy = TempNode.Item(0).Attributes.getNamedItem("Hierarchy").value	
	FinYear =  TempNode.Item(0).Attributes.getNamedItem("FinYear").value	
	ComputeMode =  TempNode.Item(0).Attributes.getNamedItem("ComputeMode").value
	sAccHeadID = TempNode.Item(0).Attributes.getNamedItem("AccountHeadID").value
	sInsDate = TempNode.Item(0).Attributes.getNamedItem("InsDate").value
	
	If ComputeMode= "+" Then
	ComputeMode = "+"
	Else
	ComputeMode="-"
	End If
	
	
	
	if sAccHeadID="" Then sAccHeadID = 0 
	DisplayACHeadDescr = "N"
	DataEntry = "N"
	If Mode= "A" Then
	DisplayACHeadDescr = "Y"
	ElseIf Mode = "D" Then
	DataEntry ="Y"
	End If
	Con.BeginTrans
	'------Begin---------------------Updation Part---------------------------------------------
	If ID = "2" And Level2ID<>"A" Then 
		sqlstring = "UPDATE ACC_M_SchdBreakupHeads SET BreakupHeading='"&BreakUpHeadName&"'  where BreakupID = '"&Level2ID&"'"    		
		con.execute sqlstring 
	ElseIf ID ="3" And Level3ID<>"A" Then 
		sqlstring = "UPDATE ACC_M_SchdBreakupHeads SET BreakupHeading='"&BreakUpHeadName&"'  where BreakupID = '"&Level2ID&"'"    		
		con.execute sqlstring 
		sqlstring = "UPDATE ACC_M_SchdBreakupSubHeads Set SubBreakupName='"&BreakUpSubHead&"', DisplayACHeadDescr='"&DisplayACHeadDescr&"',DataEntry='"&DataEntry&"',computemode='"&ComputeMode&"',Hierachy='"&Hierarchy&"' where BreakupID = '"&Level2ID&"' and BreakupSubID='"&BreakupSubId&"' and BreakupSubSubID='"&BreakupSubSubId&"'"     
		con.execute sqlstring 		
		If Hierarchy="" Then
		sqlstring = "Select Hierachy From ACC_M_SchdBreakupSubHeads where BreakupID = '"&Level2ID&"' and BreakupSubID='"&BreakupSubId&"' and BreakupSubSubID='"&BreakupSubSubId&"'"     
		With objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring
			.Open 
		End With
		Hierarchy = objrs(0)
		objrs.Close 
		End If
		'Response.Write BreakupId '&"-"&BreakupSubId&"-"&BreakupSubSubId   
		
		
		sqlstring = "Delete From ACC_T_SchdBreakupACDetail where BreakupID='"&Level2ID&"' and BreakupSubId='"&BreakupSubId&"' and BreakupSubSubId='"&BreakupSubSubId&"'"     
		con.execute sqlstring
	'---------End--------------------Updation Part----------------------------------------------
	Else
	
	'--------Begin-------------------Insertion Part--------------------------------------------
	
	If ID = "1" Then
	
	
	ElseIf ID = "2" and BreakUpHeadName<>""  Then
	     
		sqlstring = "Select IsNull(MAX(BreakUpID),0)+1 from Acc_m_SchdBreakUpHeads"	
		With objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring
			.Open 
		End With
		BreakupId = objrs(0)
		objrs.Close 
		sqlstring = "SELECT ISNULL(MAX(Hierarchy), 0) + 1 AS Hierarchy FROM dbo.ACC_M_SchdBreakupHeads" &_
                    " WHERE (ScheduleID = "&ScheduleID&") AND (ScheduleSubID = "&ScheduleSubID&") AND (ScheduleSubSubID = "&ScheduleSubSubID&")"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring
			.Open 
		End With
		Hierarchy = objrs(0)
		objrs.Close 
		
		sqlstring = "INSERT INTO Acc_M_SchdBreakUpHeads (ScheduleID, ScheduleSubID, ScheduleSubSubID, BreakupID, BreakupHeading, Hierarchy, FinYear, Useable)"&_
					"VALUES("&ScheduleID&","&ScheduleSubID&","&ScheduleSubSubID&",'"&BreakupId&"','"&BreakUpHeadName&"','"&Hierarchy&"','"&FinYear&"','Y')"
	   ' Response.Write sqlstring
		con.execute (sqlstring)
		
	Else
		'Adding BreakUp Head
		if BreakUpHeadName<>"" Then 
			sqlstring = "Select IsNull(MAX(BreakUpID),0)+1 from Acc_m_SchdBreakUpHeads"	
		With objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring
			.Open 
		End With
		BreakupId = objrs(0)
		objrs.Close 
		sqlstring = "SELECT ISNULL(MAX(Hierarchy), 0) + 1 AS Hierarchy FROM dbo.ACC_M_SchdBreakupHeads" &_
                    " WHERE (ScheduleID = "&ScheduleID&") AND (ScheduleSubID = "&ScheduleSubID&") AND (ScheduleSubSubID = "&ScheduleSubSubID&")"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring
			.Open 
		End With
		Hierarchy = objrs(0)
		objrs.Close 
		
		sqlstring = "INSERT INTO Acc_M_SchdBreakUpHeads (ScheduleID, ScheduleSubID, ScheduleSubSubID, BreakupID, BreakupHeading, Hierarchy, FinYear, Useable)"&_
					"VALUES("&ScheduleID&","&ScheduleSubID&","&ScheduleSubSubID&",'"&BreakupId&"','"&BreakUpHeadName&"','"&Hierarchy&"','"&FinYear&"','Y')"
		con.execute (sqlstring)
		End If
		If BreakUpSubHead <>"" Then
			sqlstring = "SELECT ISNULL(MAX(Hierachy), 0) + 1 AS Hierachy FROM dbo.ACC_M_SchdBreakupSubHeads WHERE (BreakupID = "&BreakupId&")"
			With Objrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring
				.Open 
			End With
			Hierarchy = objrs(0)		
			objrs.Close
			sqlstring = "SELECT ISNULL(MAX(BreakupSubID), 0) + 1 AS BreakupSubID FROM dbo.ACC_M_SchdBreakupSubHeads WHERE (BreakupID = "&BreakupId&")"
			With Objrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring
				.Open 
			End With
			BreakupSubId = objrs(0)		
			objrs.Close 
			sqlstring = "SELECT ISNULL(MAX(BreakupSubSubID), 0) + 1 AS BreakupSubSubID FROM dbo.ACC_M_SchdBreakupSubHeads WHERE (BreakupID = "&BreakupId&")"
			With Objrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sqlstring
				.Open 
			End With
			BreakupSubSubId = objrs(0)		
			objrs.Close 
			sqlstring = "INSERT INTO dbo.ACC_M_SchdBreakupSubHeads (BreakupID, BreakupSubID, BreakupSubSubID, SubBreakupName, DisplayACHeadDescr, Hierachy, DataEntry, FinYear, ComputeMode)"&_
						"VALUES("&BreakupId&","&BreakupSubId&","&BreakupSubSubId&",'"&BreakUpSubHead&"','"&DisplayACHeadDescr&"','"&Hierarchy&"','"&DataEntry&"','"&FinYear&"','"&ComputeMode&"')"
			'Response.Write sqlstring 
			con.execute sqlstring
	End If
	End If

	End If	  
'-------------End----------------InsertionPart------------------------------


'------------------------Insert Or Update With Account Head-----------------
		'To Add AccountHead Details
	
	If Mode <> "N" and BreakUpSubHead <>"" Then
		sqlstring = "SELECT ISNULL(MAX(EntryNumber), 0) + 1 AS EntryNumber FROM dbo.ACC_T_SchdBreakupACDetail"
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sqlstring
			.Open 
		End With
		EntryNumber = objrs(0)		
		objrs.Close 
		If BreakupSubSubId ="" Then BreakupSubSubId = 0
		If BreakupId <> 0 Then
		sqlstring = "INSERT INTO dbo.ACC_T_SchdBreakupACDetail(EntryNumber, BreakupID, BreakupSubID, BreakupSubSubID, Hierarchy, OrganisationCode,  ApplicableACHeadCode, FinYear,ScheduleSubHeadValue, ComputeMode,AsoNDate)"&_
					"VALUES("&EntryNumber&","&BreakupId&","&BreakupSubId&","&BreakupSubSubId&","&Hierarchy&",'"&OrgID&"',"&sAccHeadID&",'"&FinYear&"',0,'"&ComputeMode&"',Convert(datetime,'"&sInsDate&"',103))"
		'Response.Write sqlstring 
		con.execute sqlstring
		End If
	End If
'--------End------------Insert Or Update With Account Head-------------------	
	
	Con.CommitTrans
	'Con.RollbackTrans
	
%>
