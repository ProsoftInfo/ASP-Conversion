<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<%
	'Program Name				:	XMLShdDelete_Update.asp
	'Module Name				:	Purchase(Receipts)
	'Author Name				:	K A Kumar
	'Created On					:	Jan 03,2006
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	
	'							:
	'Connects To				:	AddSchSubHeads.asp
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

	Dim oDOM,sMod,sName,sDesgDir,AccHead,id
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	sMod=Request("Mod")
	sName=Request("Name")
	id = Request("id") 
	sDesgDir = trim(Request("ToDir"))
	'Response.Write "<p> sDesgDir = " & sDesgDir 	
	oDOM.async = false
	oDOM.load(Request)
	oDOM.Save server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	'oDOM.Save server.MapPath("../temp/transaction/Test1.xml")
	set oDOM=nothing
	
%>

<%
	Dim objrs,objrs1,objrs2,sqlstring,Root
	set objRs  = server.CreateObject("adodb.recordset")
	Dim sExp,TempNode
	Dim LevelID,iBreakID,iBreakSubID,iBreakSubSubID,sOrgID,sFinyr,sShID,AcCode,splt
	Dim sschedno,SubID,SubSubID,sNo
	Dim ScheduleID, ScheduleSubID, ScheduleSubSubID, FinYear, EntryType
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.Load server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	
	set Root=oDOM.documentElement
	sExp = "//Schedule"
	AccHead = 0
	Set TempNode = Root.selectNodes(sExp)
	'Response.Write TempNode.length
	'Response.End
	
	Con.BeginTrans	
	
	id =  TempNode.Item(0).Attributes.getNamedItem("id").value
	If id = "1" Then

		sschedno = TempNode.Item(0).Attributes.getNamedItem("sschedno").value
		SubID =  TempNode.Item(0).Attributes.getNamedItem("SubID").value
		SubSubID =  TempNode.Item(0).Attributes.getNamedItem("SubSubID").value
		sOrgID =  TempNode.Item(0).Attributes.getNamedItem("sOrgID").value
		sFinyr =  TempNode.Item(0).Attributes.getNamedItem("sFinyr").value
		
		If SubSubID = 0 Then
			sqlstring = "SELECT TOP 100 PERCENT ScheduleID, ScheduleSubID, ScheduleSubSubID, FinYear, EntryType " &_
						" FROM dbo.Acc_M_SchdSetupSubHeads WHERE (FinYear='"&sFinyr&"') AND (ScheduleSubID='"&SubID&"') AND (ScheduleSubSubID>='"&SubSubID&"') AND (ScheduleID="&sschedno&")" &_
						" ORDER BY ScheduleSubSubID DESC"
		Else
			sqlstring = "SELECT TOP 100 PERCENT ScheduleID, ScheduleSubID, ScheduleSubSubID, FinYear, EntryType " &_
						" FROM dbo.Acc_M_SchdSetupSubHeads WHERE (FinYear='"&sFinyr&"') AND (ScheduleSubID='"&SubID&"') AND (ScheduleSubSubID='"&SubSubID&"') AND (ScheduleID="&sschedno&")" &_
						" ORDER BY ScheduleSubSubID DESC"
		End If
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = Con
			.Source = sqlstring
			.Open
		End With
	
		Set Objrs.ActiveConnection = Nothing
			'Response.Write sqlstring 
			While Not objrs.EOF 
				ScheduleID = objrs(0)
				ScheduleSubID = objrs(1)
				ScheduleSubSubID = objrs(2)
				FinYear = objrs(3)
				EntryType = objrs(4)
				If EntryType = "A" Then
					sqlstring = "Delete FROM dbo.Acc_T_ScheduleACDetail WHERE (ScheduleID = "&ScheduleID&") AND (ScheduleSubID = "&ScheduleSubID&") AND (ScheduleSubSubID = "&ScheduleSubSubID&" ) AND (FinYear = '"&FinYear&"')"
					con.execute sqlstring		
				End If
					sqlstring = "Delete FROM dbo.Acc_M_SchdSetupSubHeads WHERE (ScheduleID = "&ScheduleID&") AND (ScheduleSubID = "&ScheduleSubID&") AND (ScheduleSubSubID = "&ScheduleSubSubID&" ) AND (FinYear = '"&FinYear&"')"
					con.execute sqlstring	
			objrs.MoveNext 
			Wend
			objrs.Close 
		
	
		'Con.RollbackTrans
	End If	

	If id = "2" Then			
			LevelID = TempNode.Item(0).Attributes.getNamedItem("LevelID").value
			iBreakID = TempNode.Item(0).Attributes.getNamedItem("iBreakID").value
			splt = Split(iBreakID,",")
			iBreakID = splt(0)   
			iBreakSubID = TempNode.Item(0).Attributes.getNamedItem("iBreakSubID").value
			iBreakSubSubID = TempNode.Item(0).Attributes.getNamedItem("iBreakSubSubID").value
			sOrgID = TempNode.Item(0).Attributes.getNamedItem("sOrgID").value
			sFinyr = TempNode.Item(0).Attributes.getNamedItem("sFinyr").value
			sShID = TempNode.Item(0).Attributes.getNamedItem("sShID").value
			AcCode = TempNode.Item(0).Attributes.getNamedItem("AcCode").value
			
			
			If LevelID = 1 Then
				sqlstring = "Delete from ACC_M_SchdBreakupHeads where BreakupID = '"&iBreakID&"'" 	
				con.execute sqlstring
				sqlstring = "Delete from ACC_M_SchdBreakupSubHeads where BreakupID='"&iBreakID&"'" 
				con.execute sqlstring
				sqlstring = "Delete from Acc_T_SchdBreakupAcDetail where BreakupID='"&iBreakID&"'" 
				con.execute sqlstring
			End If
			If LevelID = 2 Then
				sqlstring = "Delete from ACC_M_SchdBreakupSubHeads where BreakupID='"&iBreakID&"' and  BreakupSubID='"&iBreakSubID&"' and BreakupSubSubID='"&iBreakSubSubID&"'"   
				'Response.Write sqlstring
				con.execute sqlstring
				'Response.Write AcCode 
			If Not AcCode="" Then
				sqlstring = "Delete from Acc_T_SchdBreakupAcDetail where BreakupID='"&iBreakID&"' and  BreakupSubID='"&iBreakSubID&"' and BreakupSubSubID='"&iBreakSubSubID&"'"   
			'	Response.Write sqlstring
				con.execute sqlstring
			End If			
			End If
			
	End If
	
	
	
	If id = "3" Then
		sNo = TempNode.Item(0).Attributes.getNamedItem("sNo").value
		sOrgID = TempNode.Item(0).Attributes.getNamedItem("sOrgId").value
		sFinyr = TempNode.Item(0).Attributes.getNamedItem("sFinYr").value
		sqlstring = "Delete from Acc_T_ScheduleAcDetail where ScheduleID='"&sNo&"'"
		con.execute sqlstring 
		sqlstring = "Delete from Acc_M_SchdSetupsubheads Where ScheduleID ='"&Sno&"'" 
		con.execute sqlstring
		sqlstring = "Delete from Acc_M_SchdSetupHeads where ScheduleNumber = '"&sNo&"'" 
		con.execute sqlstring  
	End If
	
	If id = "4" Then
		sschedno = TempNode.Item(0).Attributes.getNamedItem("sschedno").value
		sFinyr = TempNode.Item(0).Attributes.getNamedItem("sFinyr").value
		SubID = TempNode.Item(0).Attributes.getNamedItem("SubID").value
		SubSubID = TempNode.Item(0).Attributes.getNamedItem("SubSubID").value
		sOrgID = TempNode.Item(0).Attributes.getNamedItem("sOrgID").value
			If SubSubID = 0 Then
				sqlstring = "delete from Acc_T_PLAcDetail where PLHeadID = '"&sschedno&"' and PLSubID='"&SubID&"' and PLSubSubID>='"&SubSubID&"'"		
				con.execute sqlstring
				sqlstring = "Delete From Acc_M_PLSetupsubHeads where PLHeadID = '"&sschedno&"' and PLSubID='"&SubID&"' and PLSubSubID>='"&SubSubID&"'"
				con.execute sqlstring
			Else
				sqlstring = "delete from Acc_T_PLAcDetail where PLHeadID = '"&sschedno&"' and PLSubID='"&SubID&"' and PLSubSubID='"&SubSubID&"'"		
				con.execute sqlstring
				sqlstring = "Delete From Acc_M_PLSetupsubHeads where PLHeadID = '"&sschedno&"' and PLSubID='"&SubID&"' and PLSubSubID='"&SubSubID&"'"
				con.execute sqlstring
			End If
	End If
	
	If id = "5" Then
		sOrgID =   TempNode.Item(0).Attributes.getNamedItem("sOrgID").value
		sschedno = TempNode.Item(0).Attributes.getNamedItem("sschedno").value
		sqlstring = "Delete from Acc_T_PLAcDetail where PLHeadId ='"&sschedno&"'" 
		con.execute sqlstring
		sqlstring = "Delete from Acc_M_PLSetupsubHeads where PLHeadId='"&sschedno&"'"
		con.execute sqlstring
		sqlstring = "Delete from Acc_M_PLSetupHeads where PLHeadID='"&sschedno&"'"   
		con.execute sqlstring
	End If
	
	If id = "6" Then
		sschedno = TempNode.Item(0).Attributes.getNamedItem("sschedno").value
		sFinyr = TempNode.Item(0).Attributes.getNamedItem("sFinyr").value
		SubID = TempNode.Item(0).Attributes.getNamedItem("SubID").value
		SubSubID = TempNode.Item(0).Attributes.getNamedItem("SubSubID").value
		sOrgID = TempNode.Item(0).Attributes.getNamedItem("sOrgID").value
			If SubSubID = 0 Then
				sqlstring = "delete from Acc_T_BSAcDetail where BSHeadID = '"&sschedno&"' and BSSubID='"&SubID&"' and BSSubSubID>='"&SubSubID&"'"		
				con.execute sqlstring
				sqlstring = "Delete From Acc_M_BSSetupsubHeads where BSHeadID = '"&sschedno&"' and BSSubID='"&SubID&"' and BSSubSubID>='"&SubSubID&"'"
				con.execute sqlstring
			Else
				sqlstring = "delete from Acc_T_BSAcDetail where BSHeadID = '"&sschedno&"' and BSSubID='"&SubID&"' and BSSubSubID='"&SubSubID&"'"		
				con.execute sqlstring
				sqlstring = "Delete From Acc_M_BSSetupsubHeads where BSHeadID = '"&sschedno&"' and BSSubID='"&SubID&"' and BSSubSubID='"&SubSubID&"'"
				con.execute sqlstring
			End If
	End If
	Con.CommitTrans
	
	If id = "7" Then
		sOrgID =   TempNode.Item(0).Attributes.getNamedItem("sOrgID").value
		sschedno = TempNode.Item(0).Attributes.getNamedItem("sschedno").value
		sqlstring = "Delete from Acc_T_BSAcDetail where BSHeadId ='"&sschedno&"'" 
		con.execute sqlstring
		sqlstring = "Delete from Acc_M_BSSetupsubHeads where BSHeadId='"&sschedno&"'"
		con.execute sqlstring
		sqlstring = "Delete from Acc_M_BSSetupHeads where BSHeadID='"&sschedno&"'"   
		con.execute sqlstring
	End If
	
%>
