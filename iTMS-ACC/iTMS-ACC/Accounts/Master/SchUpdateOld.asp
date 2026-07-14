<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SchUpdateOld.asp
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
	Dim sQuery,sTemp,iSchID,iSchSubID,iSchSubSubID,iAccHead,sUnit,sFinYr,iCount,temp1
	Dim Objrs,Objrs1,Objrs2,dSchVal,iEntNo,iNewAccHead,sEntryType,sBkupID,iBkID,tBkupId
	Dim sBkTemp,iBkPara,sBkArr,Arr,iCtr,i,sBkHead,sUseable,iHrchy,iBkSchID,iBkSchSubId,iBkSchSubSubId
	Dim Array1, iNewSchId,iNewSchSubId,iNewSchSubSubId,sTempVal,k,sForTheDate
	
	
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs2 = Server.CreateObject("ADODB.RecordSet")
	sFinYr = Session("FinPeriod")
	'sUnit = Request.Form("selUnitID")
	sUnit = session("organizationcode")
	iSchID = Request.Form("selSCH")
	sEntryType = Request.Form("htype")
	sForTheDate = Trim(Request.Form("selForMonth"))
	
	'Taking SchID,SchSubID,SchSubIDSchSubSubID for the Schedule thing
	sQuery = "Select Distinct ScheduleID,ScheduleSubID,ScheduleSubSubID From Vw_Acc_SchSetup  "&_
			 "Where ScheduleID = "&iSchID&" and FinYear = '"&sFinYr&"' --and EntryType = 'S' "
			 
	'Response.Write sQuery &"<br><br>"
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
		iSchID = Objrs("ScheduleID")
		iSchSubID = Objrs("ScheduleSubID")
		iSchSubSubID = Objrs("ScheduleSubSubID")
		iAccHead = 0
		
		
		dSchVal = Request.Form("txtCurrVal"&iSchID&"?"&iSchSubID&"?"&iSchSubSubID&"?"&iAccHead)
		Response.Write " ====== "& iSchID&" " & iSchSubID &" " & iSchSubSubID &" " & iAccHead &" "&dSchVal &"<br>"
		
		IF Len(dSchVal) = 0 Then
			dSchVal = 0
		End IF	
		
		Dim sCheckVal,iNewEntNo
		
		sQuery = "Select ScheduleID From Acc_T_ScheduleACDetail WHERE ScheduleID = "&iSchID&" "&_
				 "AND ScheduleSubID = "&iSchSubID&" AND ScheduleSubSubID = "&iSchSubSubID&" AND  "&_
				 "OrganisationCode = '"&sUnit&"' "&_
				 "AND FinYear = '"&sFinYr&"' and "&_
				 "Convert(Varchar,isNull(AsOnDate,''),103) = '"&sForTheDate&"' "
				 
		Response.Write sQuery &"<br><br>"
		Objrs1.Open sQuery,Con
		IF Not Objrs1.EOF Then
			sCheckVal = "U"
		Else
			sCheckVal = "I"
		End IF
		Objrs1.Close
		
		IF CStr(sCheckVal) = "U" Then
			'Update The Trans Table With the Match of SchID,SchSubID,SchSUbSubID and Unit and Fin Year
			sQuery = "UPDATE Acc_T_ScheduleACDetail SET ScheduleSubHeadValue = "&dSchVal&" "&_
					 "WHERE ScheduleID = "&iSchID&"  "&_
					 "AND ScheduleSubID = "&iSchSubID&" AND ScheduleSubSubID = "&iSchSubSubID&" AND  "&_
					 "OrganisationCode = '"&sUnit&"' "&_
					 "AND FinYear = '"&sFinYr&"' and Convert(Varchar,AsOnDate,103) = '"&sForTheDate&"'  "
					 				 
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery						 
		Else
			sQuery = "Select Max(isNull(EntryNumber,0)) + 1 From Acc_T_ScheduleACDetail "
			Objrs1.Open sQuery,con
			iNewEntNo = Objrs1(0)
			Objrs1.Close
			
			sQuery = "Select Hierarchy,ApplicableACGroupCode,ApplicableACHeadCode,ComputeMode, "&_
					 "isNull(AddnDescription,'NULL'),BreakupID,BreakupSubID From Acc_T_ScheduleACDetail "&_
					 "Where ScheduleID = "&iSchID&" AND  "&_
					 "ScheduleSubID = "&iSchSubID&" AND ScheduleSubSubID = "&iSchSubSubID&" AND  "&_
					 "OrganisationCode = '"&sUnit&"'  "
					 
			Response.Write sQuery &"<br><br>"
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With
			Set objRs.ActiveConnection = Nothing
			IF Not Objrs1.EOF Then
				sQuery = "INSERT INTO Acc_T_ScheduleACDetail(EntryNumber, ScheduleID, ScheduleSubID, ScheduleSubSubID, "&_
						 "Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode,FinYear, "&_
						 "ScheduleSubHeadValue, ComputeMode, AsOnDate, AddnDescription, BreakupID, BreakupSubID) "&_
						 "VALUES ("&iNewEntNo&", "&iSchID&", "&iSchSubID&", "&iSchSubSubID&", "&Objrs1(0)&",  "&_
						 "'"&sUnit&"', '"&Objrs1(1)&"', "&Objrs1(2)&", '"&sFinYr&"', "&dSchVal&", "&_
						 "'"&Objrs1(3)&"', Convert(datetime,'"&sForTheDate&"',103), "&_
						 "'"&Objrs1(4)&"', "&Objrs1(5)&", "&Objrs1(6)&") "
				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery						 
			End IF
			Objrs1.Close
			
					 
		End IF	
		Objrs.MoveNext
	loop
	Objrs.Close
	'Taking the Unique SchSubID,SchSubSubID and Accounthead if any for the Selected
	'SchID. Combine SchID,SchSubIDSchSubSubID and ACcounthead and Take the Val From Previous Page
	
	sQuery = "Select Distinct V.ScheduleID,V.ScheduleSubID,V.ScheduleSubSubID,V.EntryType, "&_
			 "isNull(T.ApplicableAcHeadCode,0) ApplicableAcHeadCode From Vw_Acc_SchSetup  V,Acc_T_ScheduleACDetail T  "&_
			 "Where V.ScheduleID = T.ScheduleID and V.ScheduleSubID = T.ScheduleSubID  "&_
			 "and V.ScheduleSubSubID = T.ScheduleSubSubID and V.ScheduleID = "&iSchID&" "&_
			 "and T.OrganisationCode = '"&sUnit&"' and V.FinYear = '"&sFinYr&"' "		 
			 
	Response.Write sQuery &"<br><br>"
	
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
		iSchID = Objrs("ScheduleID")
		iSchSubID = Objrs("ScheduleSubID")
		iSchSubSubID = Objrs("ScheduleSubSubID")
		iAccHead = Objrs("ApplicableACHeadCode")
		
		
		dSchVal = Request.Form("txtCurrVal"&iSchID&"?"&iSchSubID&"?"&iSchSubSubID&"?"&iAccHead)
		Response.Write " ====== "& iSchID&" " & iSchSubID &" " & iSchSubSubID &" " & iAccHead &" "&dSchVal &"<br>"
		
		IF Len(dSchVal) = 0 Then
			dSchVal = 0
		End IF	
		
		sQuery = "Select ScheduleID From Acc_T_ScheduleACDetail WHERE ScheduleID = "&iSchID&" "&_
				 "AND ScheduleSubID = "&iSchSubID&" AND ScheduleSubSubID = "&iSchSubSubID&" AND  "&_
				 "OrganisationCode = '"&sUnit&"' "&_
				 "AND isNull(ApplicableACHeadCode,0) = "&iAccHead&" AND FinYear = '"&sFinYr&"' and "&_
				 "Convert(Varchar,isNull(AsOnDate,''),103) = '"&sForTheDate&"' "
				 
		Response.Write sQuery &"<br><br>"
		Objrs1.Open sQuery,Con
		IF Not Objrs1.EOF Then
			sCheckVal = "U"
		Else
			sCheckVal = "I"
		End IF
		Objrs1.Close
		
		IF CStr(sCheckVal) = "U" Then
			'Update The Trans Table With the Match of SchID,SchSubID,SchSUbSubID and Unit and Fin Year
			sQuery = "UPDATE Acc_T_ScheduleACDetail SET ScheduleSubHeadValue = "&dSchVal&" "&_
					 "WHERE ScheduleID = "&iSchID&"  "&_
					 "AND ScheduleSubID = "&iSchSubID&" AND ScheduleSubSubID = "&iSchSubSubID&" AND  "&_
					 "OrganisationCode = '"&sUnit&"' AND isNull(ApplicableACHeadCode,0) = "&iAccHead&"  "&_
					 "AND FinYear = '"&sFinYr&"' and Convert(Varchar,AsOnDate,103) = '"&sForTheDate&"'  "
					 					 				 
			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery						 
		Else
			sQuery = "Select Max(isNull(EntryNumber,0)) + 1 From Acc_T_ScheduleACDetail "
			Objrs1.Open sQuery,con
			iNewEntNo = Objrs1(0)
			Objrs1.Close
			
			sQuery = "Select Hierarchy,ApplicableACGroupCode,ApplicableACHeadCode,ComputeMode, "&_
					 "isNull(AddnDescription,''),BreakupID,BreakupSubID From Acc_T_ScheduleACDetail "&_
					 "Where ScheduleID = "&iSchID&" AND  "&_
					 "ScheduleSubID = "&iSchSubID&" AND ScheduleSubSubID = "&iSchSubSubID&" AND  "&_
					 "OrganisationCode = '"&sUnit&"'  "
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With
			Set objRs.ActiveConnection = Nothing
			IF Not Objrs1.EOF Then
				sQuery = "INSERT INTO Acc_T_ScheduleACDetail(EntryNumber, ScheduleID, ScheduleSubID, ScheduleSubSubID, "&_
						 "Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode,FinYear, "&_
						 "ScheduleSubHeadValue, ComputeMode, AsOnDate, AddnDescription, BreakupID, BreakupSubID) "&_
						 "VALUES ("&iNewEntNo&", "&iSchID&", "&iSchSubID&", "&iSchSubSubID&", "&Objrs1(0)&",  "&_
						 "'"&sUnit&"', '"&Objrs1(1)&"', "&Objrs1(2)&", '"&sFinYr&"', "&dSchVal&", "&_
						 "'"&Objrs1(3)&"', Convert(datetime,'"&sForTheDate&"',103), "&_
						 "'"&Objrs1(4)&"', "&Objrs1(5)&", "&Objrs1(6)&") "
				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery						 
			End IF
			Objrs1.Close
		End IF	
				
			'====================== Account Head Updation =========================================
			If trim(sEntryType) = "A" Then
				sQuery= "Select EntryNumber From Acc_T_ScheduleACDetail  Where ScheduleID =  "&iSchID&" "&_
						" and OrganisationCode = '"&sUnit&"' and FinYear = '"&sFinYr&"' "
				'Response.Write sQuery &"<Br><Br>"
				With Objrs1
					.CursorType = 3
					.CursorLocation = 3
					.Source = sQuery
					.ActiveConnection = Con
					.Open
				End With
				
				Do while not Objrs1.EOF 
					iEntNo = Objrs1(0)
					
					iNewAccHead = Request.Form("hAccHead"&iEntNo)
					'Response.Write "NewAcc=" & iNewAccHead &"<Br><Br>"
					IF CStr(iNewAccHead) <> "" Then
						sQuery = "UPDATE Acc_T_ScheduleACDetail SET ApplicableACHeadCode = "&iNewAccHead&", "&_
								 "AsOnDate = Convert(datetime,'"&sForTheDate&"',103) Where "&_
								 "EntryNumber = "&iEntNo&" "
						Response.Write sQuery &"<Br><Br>"
						Con.Execute sQuery
					End IF
					Objrs1.MoveNext
				Loop
				Objrs1.Close
			End IF 'If trim(sEntryType) = "A" Then
		Objrs.MoveNext
	Loop
	Objrs.Close
	

Dim sCheck
'Response.Clear
iCount = 0
'sQuery = "Update Acc_M_SchdBreakupheads Set Useable = 'N' where ScheduleID =  "&iSchID&" "
'Con.Execute sQuery

sQuery = "Select Distinct ScheduleSubID From Vw_Acc_SchSetup WHere ScheduleID = "&iSchID&" and EntryType = 'S' "
With Objrs2
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End with
Do While Not Objrs2.EOF 
	iCount = CDbl(iCount) + 1
	sTempVal = Request.Form("hSchHead"&iCount&Objrs2(0))
	Response.Write sTempVal &"==================== <br><br>"
	sTemp = Split(sTempVal,"~~")
	IF UBound(sTemp) > 1 Then
		sBkArr = Split(sTemp(1),",") 'For Schedule ID,ScheduleSUB ID,ScheduleSUBSUB ID,Breakup ID 
		sBkHead = Split(sTemp(0),":") 'For Breakup Name
		For iCtr = 0 To UBound(sBkHead)
			sCheck = ""
			
			sBkTemp = Split(sBkArr(iCtr),"-") 
			
			
			sQuery = "Select Count(1) From Acc_M_SchdBreakupheads WHere ScheduleID = "&sBkTemp(0)&" and  "&_
					 "ScheduleSubID = "&sBkTemp(1)&" and ScheduleSubSubID = "&sBkTemp(2)&" and BreakupID = "&sBkTemp(3)&" and  "&_
					 "FinYear = '"&sFinYr&"' "
						 
			Response.Write sQuery &"<br><br>"
			Objrs.Open sQuery,Con
			IF Not Objrs.EOF Then
				sCheck = "U"
			Else
				
				sCheck = "I"
			End IF
			Objrs.Close
			IF CStr(sCheck) = "I" Then
				sQuery = "Select isNull(Max(Hierarchy),0) From Acc_M_SchdBreakupheads WHere "&_
						 "ScheduleID = "&sBkTemp(0)&" and ScheduleSubID = "&sBkTemp(1)&" and ScheduleSubSubID = "&sBkTemp(2)&" "&_
						 "and FinYear = '"&sFinYr&"' "
				Objrs.Open sQuery,Con
				IF Not Objrs.EOF Then
					iHrchy = Objrs(0)
				End IF
				Objrs.Close
				
				iHrchy = CDbl(iHrchy) + 1
				
				sQuery = "INSERT INTO ACC_M_SchdBreakupHeads (ScheduleID, ScheduleSubID, ScheduleSubSubID, "&_
						 "BreakupID, BreakupHeading, Hierarchy, FinYear, Useable) "&_
						 "VALUES ("&sBkTemp(0)&", "&sBkTemp(1)&", "&sBkTemp(2)&", "&sBkTemp(3)&", '"&sBkHead(iCtr)&"', "&iHrchy&", '"&sFinYr&"', 'Y') "
						
				
				Response.Write sQuery &"<br><br>"	 
				Con.execute sQuery
			Elseif CStr(sCheck) = "U" Then
				sQuery = "Update Acc_M_SchdBreakupheads Set Useable = 'Y' WHere ScheduleID = "&sBkTemp(0)&" and  "&_
						 "ScheduleSubID = "&sBkTemp(1)&" and ScheduleSubSubID = "&sBkTemp(2)&" and BreakupID = "&sBkTemp(3)&" and  "&_
						 "FinYear = '"&sFinYr&"' "
				
				Response.Write sQuery &"<br><br>"		 
				Con.execute sQuery
			End IF	
			
		Next
	ENd IF
	Objrs2.MoveNext
loop
Objrs2.Close

'Response.Clear
sQuery = "Select Distinct ScheduleSubID,ScheduleSubSubID From Vw_Acc_SchSetup WHere ScheduleID = "&iSchID&" and EntryType = 'S' "
With Objrs2
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End with
Do While Not Objrs2.EOF 
	iCount = CDbl(iCount) + 1
	'Response.Write Objrs2(0)&Objrs2(0)&Objrs2(1) &"<br>"
	sTempVal = Request.Form("hSchHead"&Objrs2(0)&Objrs2(0)&Objrs2(1))
	'Response.Write " ================= " & sTempVal
	IF Cstr(sTempVal) <> "" Then
		sTemp = Split(sTempVal,"~~")
		IF UBound(sTemp) > 1 Then
			sBkArr = Split(sTemp(1),",") 'For Schedule ID,ScheduleSUB ID,ScheduleSUBSUB ID,Breakup ID 
			sBkHead = Split(sTemp(0),":") 'For Breakup Name
			For iCtr = 0 To UBound(sBkHead)
				sCheck = ""
				
				sBkTemp = Split(sBkArr(iCtr),"-") 
				sQuery = "Select Count(1) From Acc_M_SchdBreakupheads WHere ScheduleID = "&sBkTemp(0)&" and  "&_
						 "ScheduleSubID = "&sBkTemp(1)&" and ScheduleSubSubID = "&sBkTemp(2)&" and BreakupID = "&sBkTemp(3)&" and  "&_
						 "FinYear = '"&sFinYr&"' "
							 
				Response.Write sQuery &"<br><br>"
				Objrs.Open sQuery,Con
				IF Not Objrs.EOF Then
					sCheck = "U"
				Else
					
					sCheck = "I"
				End IF
				Objrs.Close
				IF CStr(sCheck) = "I" Then
					sQuery = "Select isNull(Max(Hierarchy),0) From Acc_M_SchdBreakupheads WHere "&_
							 "ScheduleID = "&sBkTemp(0)&" and ScheduleSubID = "&sBkTemp(1)&" and ScheduleSubSubID = "&sBkTemp(2)&" "&_
							 "and FinYear = '"&sFinYr&"' "
					Objrs.Open sQuery,Con
					IF Not Objrs.EOF Then
						iHrchy = Objrs(0)
					End IF
					Objrs.Close
					
					iHrchy = CDbl(iHrchy) + 1
					
					sQuery = "INSERT INTO ACC_M_SchdBreakupHeads (ScheduleID, ScheduleSubID, ScheduleSubSubID, "&_
							 "BreakupID, BreakupHeading, Hierarchy, FinYear, Useable) "&_
							 "VALUES ("&sBkTemp(0)&", "&sBkTemp(1)&", "&sBkTemp(2)&", "&sBkTemp(3)&", '"&sBkHead(iCtr)&"', "&iHrchy&", '"&sFinYr&"', 'Y') "
							
					
					Response.Write sQuery &"<br><br>"	 
					Con.execute sQuery
				Elseif CStr(sCheck) = "U" Then
					sQuery = "Update Acc_M_SchdBreakupheads Set Useable = 'Y' WHere ScheduleID = "&sBkTemp(0)&" and  "&_
							 "ScheduleSubID = "&sBkTemp(1)&" and ScheduleSubSubID = "&sBkTemp(2)&" and BreakupID = "&sBkTemp(3)&" and  "&_
							 "FinYear = '"&sFinYr&"' "
					
					Response.Write sQuery &"<br><br>"		 
					Con.execute sQuery
				End IF	
				
			Next
		End IF
	End IF
	Objrs2.MoveNext
loop
Objrs2.Close
	

'///////////////////////////// Changes Done By Manohar Prabhu on 19/10/2006///////////////////////////
	
	'Con.RollbackTrans
	Con.CommitTrans
	Response.Redirect "SchSetup.asp?selUnitId="&sUnit&"&selSch="&iSchID
%>