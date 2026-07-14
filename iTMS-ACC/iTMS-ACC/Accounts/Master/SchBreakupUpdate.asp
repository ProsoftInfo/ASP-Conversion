<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SchBreakupUpdate.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Manohar Prabhu .R
	'Created On					:	Nov 27 2003
	'Modified By				:	Maheshwari S.
	'Modified On				:	Oct 06 2006
	'Tables Used				:	Vw_Acc_SchBreakSetup,Acc_T_SchdBreakupACDetail
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	SchBreakupSetup.asp
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
	Dim sQuery,sTemp,iBkID,iBkSubID,iBkSubSubID,iAccHead,sUnit,sSchID,sFinYr
	Dim Objrs,dSchVal,Objrs1,Objrs2,iEntNo,iNewAccHead,iAccHeadNo,iNewBkUpID
	Dim sForDate,sCheckVal,iNewEntNo
		
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set Objrs1 = Server.CreateObject("ADODB.Recordset")
	Set Objrs2 = Server.CreateObject("ADODB.Recordset")
	sFinYr = Session("FinPeriod")
	sUnit = Request.Form("selUnitID")
	sSchID = Request.Form("selSCH")
	sForDate = Request.Form("selForMonth")
	
	sQuery = "Select Distinct V.BreakupID,V.BreakupSubID,V.BreakupSubSubID,T.ApplicableACHeadCode From "&_
			 "Vw_Acc_SchBreakSetup V,Acc_T_SchdBreakupACDetail T Where V.BreakUpID = T.BreakUpID "&_
			 " and V.BreakupSubID = T.BreakupSubID and V.BreakupSubSubID = T.BreakupSubSubID "&_
			 " and T.OrganisationCode = '"&sUnit&"' and V.FinYear = T.FinYear and V.ScheduleID = "&sSchID
	'Response.Write sQuery & "<BR><BR>"	 
			 
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
		iBkID = Objrs("BreakupID")
		iBkSubID = Objrs("BreakupSubID")
		iBkSubSubID = Objrs("BreakupSubSubID")
		iAccHead = Objrs("ApplicableACHeadCode")
		dSchVal = Request.Form("txtCurrVal"&iBkID&"?"&iBkSubID&"?"&iBkSubSubID&"?"&iAccHead)
		Response.Write iBkId & " " & iBkSubID &" " & iBkSubSubID &" " & iAccHead &" " & dSchVal &"<br>"
		
		IF CStr(dSchVal) = "" Then
			dSchVal = 0
		ENd IF
		If trim(iAccHead) <> "Null" then
			sQuery = "Select BreakUpID From Acc_T_SchdBreakupACDetail Where BreakUpID = "&iBkID&"  "&_
					 "and BreakupSubID = "&iBkSubID&" and BreakupSubSubID = "&iBkSubSubID&" and OrganisationCode = '"&sUnit&"' "&_
					 "and FinYear = '"&sFinYr&"' and ApplicableACHeadCode  = "&iAccHead&" and "&_
					 "Convert(Varchar,AsOnDate,103) = '"&sForDate&"' "
			'	
			'Response.Write sQuery &"<br><br>"	 
			With Objrs1
				.ActiveConnection = Con
				.CursorLocation = 3
				.CursorType = 3
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
			
			'sCheckVal = "U"
			
			IF CStr(sCheckVal) = "U" Then
				sQuery = "Update Acc_T_SchdBreakupACDetail Set ScheduleSubHeadValue = "&dSchVal&" Where "&_
						 "BreakUpID = "&iBkID&" and BreakupSubID = "&iBkSubID&" and BreakupSubSubID = "&iBkSubSubID&" and  "&_
						 "OrganisationCode = '"&sUnit&"' and FinYear = '"&sFinYr&"'  "&_
						 "and ApplicableACHeadCode  = "&iAccHead&" and Convert(Varchar,AsOnDate,103) = '"&sForDate&"' "
				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery	
			Else
			'	Response.Write "Indise"
				sQuery = "Select isNull(Max(EntryNumber),0)+1 From Acc_T_SchdBreakupACDetail "
				With Objrs1
					.ActiveConnection = Con
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.Open 
				End With
				Set Objrs1.ActiveConnection = Nothing
				IF Not Objrs1.EOF Then
					iNewEntNo = Objrs1(0)
				End IF
				Objrs1.Close
				
				sQuery = "Select Hierarchy,isNull(ApplicableACGroupCode,'0'),ComputeMode, "&_
						 "isNull(AddnDescription,'NULL') From ACC_T_SchdBreakupACDetail Where BreakUpID = "&iBkID&" "&_
						 "and BreakupSubID = "&iBkSubID&" and BreakupSubSubID = "&iBkSubSubID&" and OrganisationCode = '"&sUnit&"' and  "&_
						 "FinYear = '"&sFinYr&"' and ApplicableACHeadCode  = "&iAccHead&" "
				With Objrs1
					.ActiveConnection = Con
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.Open 
				End With
				Set Objrs1.ActiveConnection = Nothing
				IF Not Objrs1.EOF Then
					sQuery = "INSERT INTO ACC_T_SchdBreakupACDetail(EntryNumber, BreakupID, BreakupSubID, BreakupSubSubID, "&_
							 "Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode, FinYear, "&_
							 "ScheduleSubHeadValue, ComputeMode, AsOnDate, AddnDescription) "&_
							 "VALUES ("&iNewEntNo&", "&iBkID&", "&iBkSubID&", "&iBkSubSubID&", "&Objrs1(0)&",  "&_
							 "'"&sUnit&"', '"&Objrs1(1)&"', "&iAccHead&", '"&sFinYr&"', "&dSchVal&", "&_
							 "'"&Objrs1(2)&"', Convert(datetime,'"&sForDate&"',103), '"&Objrs1(3)&"') "
				
					Response.Write sQuery &"<br><br>"
					Con.Execute sQuery				 
				End IF
				Objrs1.Close
			End IF
		End if
		
		Objrs.MoveNext
	Loop
	Objrs.Close
	
	
	'==========================================================================================
	sQuery = "Select Distinct BreakupID From Vw_Acc_SchBreakSetup Where ScheduleID = "&sSchID&" "
	
	
		With Objrs1
			.CursorLocation =3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End with
		
		Do While not Objrs1.EOF
			iNewBkUpID = Objrs1(0)
	
				sQuery = "Select EntryNumber,ApplicableACHeadCode From Acc_T_SchdBreakupACDetail  Where BreakupID = "&iNewBkUpID&" "&_
						 "and OrganisationCode = '"&sUnit&"' and FinYear = '"&sFinYr&"' "&_
						 "and Convert(Varchar,AsOnDate,103) = '"&sForDate&"' "
				Response.Write sQuery &"<br><br>"
				With Objrs2
					.CursorLocation =3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				End with
		
				iEntNo = 0
				Do While not Objrs2.EOF
					'Response.Write Objrs2("EntryNumber")
					iEntNo = Objrs2(0)
					iNewAccHead = Request.Form("hAccHead"&iEntNo)
					
					IF iNewAccHead = "" Then
						iNewAccHead = 0
					End IF
						sQuery = "UPDATE Acc_T_SchdBreakupACDetail SET ApplicableACHeadCode = "&iNewAccHead&" Where "&_
								 "EntryNumber = "&iEntNo&" "		
						Response.Write sQuery &"<Br><Br>"
					Con.Execute sQuery
					Objrs2.MoveNext
				loop
				Objrs2.close				
			Objrs1.MoveNext
		loop
		Objrs1.close
				
	'Con.RollbackTrans	
	Con.CommitTrans	
	Response.Redirect "SchBreakupSetup.asp"
%>