<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BSUpdate.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Manohar Prabhu .R
	'Created On					:	Nov 27 2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
	Dim sQuery,sTemp,iSchID,iSchSubID,iSchSubSubID,iAccHead,sUnit,sFinYr
	Dim Objrs,dSchVal,sForDate,sCheckVal,iNewEntNo,Objrs1
	
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	sFinYr = Session("FinPeriod")
	sUnit = session("organizationcode") 'Request.Form("selUnitID")
	iSchID = Request.Form("selSCH")
	iSchID = Request.Form("selSCH")
	sForDate = Request.Form("selForMonth")
	
	'Taking the Unique SchSubID,SchSubSubID and Accounthead if any for the Selected
	'SchID. Combine SchID,SchSubIDSchSubSubID and ACcounthead and Take the Val From Previous Page
	
	sQuery = "Select BSSubID,BSSubSubID,BSSubHeadingName,EntryType From  "&_
			 "ACC_M_BSSetupSubHeads Where BSHeadID = "&iSchID&" Order By BSSubID,BSSubSubID,Hierachy "
			 
	'Response.Write sQuery
	Con.BeginTrans
	With Objrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set Objrs.ActiveConnection = Nothing
	dSchVal = 0
	Do While Not Objrs.EOF
		iSchSubID = Objrs("BSSubID")
		iSchSubSubID = Objrs("BSSubSubID")
		
		
		dSchVal = Request.Form("txtCurrVal"&iSchID&"?"&iSchSubID&"?"&iSchSubSubID)
		
		IF Len(dSchVal) = 0 Then
			dSchVal = 0
		End IF	
		
		sQuery = "Select BSHeadID From ACC_T_BSACDetail WHERE BSHeadID = "&iSchID&" AND  "&_
				 "BSSubID = "&iSchSubID&" AND BSSubSubID = "&iSchSubSubID&" AND OrganisationCode = '"&sUnit&"' "&_
				 "AND FinYear = '"&sFinYr&"' and Convert(Varchar,AsOnDate,103) = '"&sForDate&"' "
		
		'Response.Write sQuery 
		With Objrs1
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = Con
			.Source = sQuery
			.Open
		End With
		Set Objrs1.ActiveConnection = Nothing
		IF Not Objrs1.EOF Then
			sCheckVal = "U"
		Else
			sCheckVal = "I"
		End IF
		Objrs1.Close
		IF CStr(sCheckVal) = "U" Then
			'Update The Trans Table With the Match of SchID,SchSubID,SchSUbSubID and Unit and Fin Year
			sQuery = "UPDATE ACC_T_BSACDetail SET BSSubHeadValue = "&dSchVal&" "&_
					 "WHERE BSHeadID = "&iSchID&" AND BSSubID = "&iSchSubID&" AND  "&_
					 "BSSubSubID = "&iSchSubSubID&" AND OrganisationCode = '"&sUnit&"' AND  "&_
					 "FinYear = '"&sFinYr&"' "
					 
					 
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
		Else
			sQuery = "Select Hierarchy,ApplicableACHeadCode,DisplayACHeadDescr,ComputeMode, "&_
					 "isNull(AddnDescription,'NULL'), ScheduleID, ScheduleSubID, ScheduleSubSubID From  "&_
					 "ACC_T_BSACDetail WHERE BSHeadID = "&iSchID&" AND BSSubID = "&iSchSubID&" AND "&_
					 "BSSubSubID = "&iSchSubSubID&" AND OrganisationCode = '"&sUnit&"'  "&_
					 "AND FinYear = '"&sFinYr&"' "
				
			'Response.Write sQuery	 
			With Objrs1
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = Con
			.Source = sQuery
			.Open
		End With
		Set Objrs1.ActiveConnection = Nothing
		Do While Not Objrs1.EOF
			sQuery = "INSERT INTO ACC_T_BSACDetail (BSHeadID, BSSubID, BSSubSubID, Hierarchy, OrganisationCode, "&_
					 "ApplicableACHeadCode, DisplayACHeadDescr, FinYear, BSSubHeadValue,ComputeMode, AsOnDate, "&_
					 "AddnDescription, ScheduleID, ScheduleSubID, ScheduleSubSubID) "&_
					 "VALUES ("&iSchID&", "&iSchSubID&", "&iSchSubSubID&", "&Objrs1(0)&", "&_
					 "'"&sUnit&"', "&Objrs1(1)&", '"&Objrs1(2)&"', '"&sFinYr&"', "&dSchVal&", "&_
					 "'"&Objrs1(3)&"', Convert(datetime,'"&sForDate&"',103), '"&Objrs1(4)&"', "&Objrs1(5)&", "&Objrs1(6)&", "&Objrs1(7)&") "
					 
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
			Objrs1.MoveNext
		Loop
		Objrs1.Close
		End IF		 
		Objrs.MoveNext
	Loop
	Objrs.Close
	
	'Con.RollbackTrans
	
	Con.CommitTrans
	
	Response.Redirect "BSSetup.asp"
%>