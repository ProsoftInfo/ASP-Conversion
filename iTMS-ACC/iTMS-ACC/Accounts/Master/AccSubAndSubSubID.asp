<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AccSubAndSubSubID.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	KUMAR K A
	'Created On					:	Dec 25, 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->

<%
Dim objRs,objRs1,sQuery,OutData,Root,newElem,sTdsElgi,sql,id
Dim sShID,sOrgID,sFinyr 
Dim sCode,sValue,bCostcenter,bAnalytical,iHeadCount,sTranFalg
Dim sShSubID,sShSubSubID,sBreakID
Dim NewCode
Set OutData = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")
sBreakID = 0
id = Request("id") 
'Response.Write id  
	Select Case id
	case "0"
		sOrgID = Request("sOrgID")
		sShID= Request("sShID")
		sFinyr = Request("sFinyr")
		sShSubID = Request("sShSubID")
		sShSubSubID= Request("sShSubSubID")
		sql = "Select Distinct BreakupHeading,BreakupID,Hierarchy From Vw_Acc_SchBreakSetup  "&_
			  "Where ScheduleID="&sShID&" and ScheduleSubID="&sShSubID&" and ScheduleSubSubID="&sShSubSubID&" "&_
			  "and FinYear = '"&sFinYr&"' Order By BreakupHeading "
			  Response.Write sql
			 ' Response.End 
			With objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			
			Set objRs.Activeconnection = Nothing
			iHeadCount=objRs.RecordCount

			If Not objRs.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs.EOF
					Set newElem = OutData.createElement("Details")
						newElem.setAttribute "BreakUpName",Objrs(0)
						newElem.setAttribute "BreakUpID",Objrs(1)&","&objRs(2)  
						newElem.setAttribute "id",id
						'newElem.text= sValue
						Root.appendChild newElem
				  	objRs.MoveNext
				Loop
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
			End If
			objRs.Close
	
	Case "1"
		Dim Acdesc,ComputeMode
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		sBreakID = Request("sBreakID")
		sql = "Select BreakupSubID,BreakupSubSubID,SubBreakupName,DisplayACHeadDescr,DataEntry,ComputeMode,Hierachy From Vw_Acc_SchBreakSetup" &_
			  "	Where BreakupID="&sBreakID&" and FinYear ='"&sFinyr&"'"
			 ' Response.Write sql
			With objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			
			Set objRs.Activeconnection = Nothing
			iHeadCount=objRs.RecordCount

			If Not objRs.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs.EOF
				IF objrs(2)<>"" Then
					If objRs(5)="+" Then ComputeMode = "+" Else ComputeMode="++" 
					Set newElem = OutData.createElement("Details")
						newElem.setAttribute "SubBreakUpName",objRs(2)
						newElem.setAttribute "BreakUpSubID",objRs(0) 
						newElem.setAttribute "BreakUpSubSubID",objRs(1) 
						newElem.setAttribute "DisplayACHeadDescr",Replace(objRs(3),"-"," ") 
						newElem.setAttribute "DataEntry",objRs(4) 
						If objrs(3)="Y" Then
							sql ="Select ApplicableAcHeadCode From Acc_T_SchdBreakUpACDetail Where BreakupID = '"&sBreakID&"' and BreakupSubID ='"&objrs(0)&"' and BreakupSubSubID='"&Objrs(1)&"'" 
							 ' Response.Write sql
							With objRs1
								.CursorLocation = 3
								.CursorType = 3
								.Source = sql
								.ActiveConnection = con
								.Open
							End with
							If not objRs1.EOF Then AcCode = objrs1(0) Else AcCode = 0
							
							objRs1.Close 
							sql = "select AccountDescription From VwOrgGLHeads Where AccountHead = '"&AcCode&"'"
							With objRs1
								.CursorLocation = 3
								.CursorType = 3
								.Source = sql
								.ActiveConnection = con
								.Open
							End with
							If Not objRs1.EOF Then Acdesc = objrs1(0) Else Acdesc = 0 
							objRs1.Close 
						End If
						newElem.setAttribute "AcCode",AcCode  
						newElem.setAttribute "Acdesc",Replace(Acdesc,"-"," ") 
						newElem.setAttribute "ComputeMode",ComputeMode   
						newElem.setAttribute "Hierarchy",objRs(6) 
						'newElem.setAttribute "id",id
						Root.appendChild newElem
				  	End If
				  	objRs.MoveNext
				Loop
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
			End If
			objRs.Close
	Case "2"
		Dim SubID,SubSubID,sschedno,EntryType,AcCode,AccHead
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		sschedno = Request("sschedno") 
		SubID = Request("SubID")
		SubSubID = Request("SubSubID") 
			
		sql = "SELECT ScheduleSubSubID, SubHeadingName, EntryType,computemode FROM dbo.Acc_M_SchdSetupSubHeads " &_
			" WHERE (ScheduleID ="&sschedno&") AND (ScheduleSubID="&SubID&") AND (FinYear='"&sFinyr&"') AND (ScheduleSubSubID<>'"&SubSubID&"')"
		Response.Write sql
			With objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			
			Set objRs.Activeconnection = Nothing
			iHeadCount=objRs.RecordCount

			If Not objRs.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs.EOF
					Set newElem = OutData.createElement("Details")
						newElem.setAttribute "ScheduleSubSubID",objRs(0)&"-"&objRs(2)&"-"&objRs(3) 
						newElem.setAttribute "SubHeadingName",objRs(1) 
					'	newElem.setAttribute "ComputeMode",objRs(3) 						
						Root.appendChild newElem
				  	objRs.MoveNext
				Loop
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
			End If
			objRs.Close
			
	Case "3"
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		sschedno = Request("sschedno") 
		SubID = Request("SubID")
		SubSubID = Request("SubSubID") 
		sql = "SELECT ISNULL(ApplicableACHeadCode,0) AS ApplicCode FROM dbo.Acc_T_ScheduleACDetail WHERE (ScheduleID="&sschedno&") AND (ScheduleSubID="&SubID&") AND (ScheduleSubSubID="&SubSubID&")" 
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			If Not objRs1.EOF Then AcCode = objRs1(0) Else AcCode = 0 
			objRs1.Close  
		sql = "Select AccountDescription from vworgglheads where accounthead='"&AcCode&"'"
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			If Not objRs1.EOF Then AccHead=objRs1(0) Else AccHead = 0
				objRs1.Close  
			Set Root = OutData.createElement("Root")
				OutData.appendChild Root
					Set newElem = OutData.createElement("Details")
						newElem.setAttribute "AcCode",AcCode 
						newElem.setAttribute "AccHead",AccHead 
						Root.appendChild newElem
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
	Case "4"
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		sschedno = Request("sschedno")		
		sql = "select PLSubID,PLSubSubID,EntryType,computemode,PLSubHeadingName,Hierachy From Acc_M_PLSetupSubHeads where PLHeadID = '"&sschedno&"' and FinYear='"&sFinyr&"' And PLSubSubID = 0"  
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			If Not objRs1.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs1.EOF
					NewCode = 0
					Set newElem = OutData.createElement("Details")
					If objrs1(2)="S" Then
						sql = "select PLSubHeadValue from Acc_T_PLAcDetail where PLHeadID='"&sschedno&"' and PLSubID='"&objrs1(0)&"' and PLSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF= True Then	NewCode = objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					If objrs1(2)="A" Then
						sql = "select ApplicableACHeadCode from Acc_T_PLAcDetail where PLHeadID='"&sschedno&"' and PLSubID='"&objrs1(0)&"' and PLSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF = True Then NewCode=objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					newElem.setAttribute "PLSubID",objRs1(0)&","&objRs1(1)&","&objRs1(2)&","&objrs1(3)&","&NewCode&","&objRs1(5)    
					newElem.setAttribute "SubHeadingName",objRs1(4) 
					Root.appendChild newElem
				  	objRs1.MoveNext
				Loop
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
			End If
			objRs1.Close
	case "5"
		Dim objhttp
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		sschedno = Request("sschedno")		
		SubID = Request("SubID")  	
		sql = "select PLSubID,PLSubSubID,EntryType,Computemode,PLSubHeadingName,Hierachy From Acc_M_PLSetupSubHeads where PLHeadID="&sschedno&" and FinYear='"&sFinyr&"' and PLSubID="&SubID&" And PLSubSubID <> 0"  
		'Response.Write sql
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			If Not objRs1.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs1.EOF
					Set newElem = OutData.createElement("Details")
					If objrs1(2)="S" Then
					sql = "select PLSubHeadValue from Acc_T_PLAcDetail where PLHeadID='"&sschedno&"' and PLSubID='"&objrs1(0)&"' and PLSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF= True Then	NewCode = objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					If objrs1(2)="A" Then
						sql = "select ApplicableACHeadCode from Acc_T_PLAcDetail where PLHeadID='"&sschedno&"' and PLSubID='"&objrs1(0)&"' and PLSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF = True Then NewCode = objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					newElem.setAttribute "PLSubID",objRs1(0)&","&objRs1(1)&","&objRs1(2)&","&objRs1(3)&","&NewCode&","&objRs1(5)   
					newElem.setAttribute "SubHeadingName",objRs1(4)  
					Root.appendChild newElem
				  	objRs1.MoveNext
				Loop
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
			End If
			objRs1.Close

		Case "6"
			sOrgID = Request("sOrgID")
			sFinyr = Request("sFinyr")
			sschedno = Request("sschedno")		
			SubID = Request("SubID")  
			SubSubID = Request("SubSubID")   
			Set Root = OutData.createElement("Root")
			OutData.appendChild Root

			sql = "Select ScheduleID,ScheduleSubID,ScheduleSubSubID from Acc_T_PLACDetail Where PLHeadID = '"&sschedno&"' and PLSubID='"&SubID&"' and PLSubSubID='"&SubSubID&"'"   
			
			With objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			While Not objrs.EOF 
			Set newElem = OutData.createElement("SchDetails")
				newElem.setAttribute "Description","ScheduleDet"
				newElem.setAttribute "ScheduleID", objRs(0) 
				newElem.setAttribute "ScheduleSubID",objRs(1)     
				newElem.setAttribute "ScheduleSubSubID",objRs(2) 
				Root.appendChild newElem
				objrs.MoveNext 
				'Response.Write objRs(2)
				'Response.End  
			Wend
			objrs.Close 
			Response.Clear
			Response.ContentType="text/xml"
			Response.Write OutData.xml
	Case "7"			
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		AcCode = Request("AcCode")
		' Response.Write AcCode&sFinyr&sOrgID    
			sql = "Select AccountDescription from vworgglheads where accounthead='"&AcCode&"'"
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			If Not objRs1.EOF Then AccHead=objRs1(0) Else AccHead = 0
				objRs1.Close  
			Set Root = OutData.createElement("Root")
				OutData.appendChild Root
					Set newElem = OutData.createElement("Details")
						newElem.setAttribute "AcCode",AcCode 
						newElem.setAttribute "AccHead",AccHead 
						Root.appendChild newElem
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml	


	Case "8"
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		sschedno = Request("sschedno")		
		sql = "select BSSubID,BSSubSubID,EntryType,computemode,BSSubHeadingName,Hierachy From Acc_M_BSSetupSubHeads where BSHeadID = '"&sschedno&"' and FinYear='"&sFinyr&"' And BSSubSubID = 0"  
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			Response.Write sOrgID&sFinyr&sschedno  
			If Not objRs1.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs1.EOF
					NewCode = 0
					Set newElem = OutData.createElement("Details")
					If objrs1(2)="S" Then
						sql = "select BSSubHeadValue from Acc_T_BSAcDetail where BSHeadID='"&sschedno&"' and BSSubID='"&objrs1(0)&"' and BSSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF= True Then	NewCode = objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					If objrs1(2)="A" Then
						sql = "select ApplicableACHeadCode from Acc_T_BSAcDetail where BSHeadID='"&sschedno&"' and BSSubID='"&objrs1(0)&"' and BSSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF = True Then NewCode=objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					newElem.setAttribute "BSSubID",objRs1(0)&","&objRs1(1)&","&objRs1(2)&","&objrs1(3)&","&NewCode&","&objRs1(5)    
					newElem.setAttribute "SubHeadingName",objRs1(4) 
					Root.appendChild newElem
				  	objRs1.MoveNext
				Loop
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
			End If
			objRs1.Close

	Case "9"
			sOrgID = Request("sOrgID")
			sFinyr = Request("sFinyr")
			sschedno = Request("sschedno")		
			SubID = Request("SubID")  
			SubSubID = Request("SubSubID")   
			Set Root = OutData.createElement("Root")
			OutData.appendChild Root
			sql = "Select ScheduleID,ScheduleSubID,ScheduleSubSubID from Acc_T_BSACDetail Where BSHeadID = '"&sschedno&"' and BSSubID='"&SubID&"' and BSSubSubID='"&SubSubID&"'"   
			With objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			While Not objrs.EOF 
			Set newElem = OutData.createElement("SchDetails")
				newElem.setAttribute "Description","ScheduleDet"
				newElem.setAttribute "ScheduleID", objRs(0) 
				newElem.setAttribute "ScheduleSubID",objRs(1)     
				newElem.setAttribute "ScheduleSubSubID",objRs(2) 
				Root.appendChild newElem
				objrs.MoveNext 
				'Response.Write objRs(2)
				'Response.End  
			Wend
			objrs.Close 
			Response.Clear
			Response.ContentType="text/xml"
			Response.Write OutData.xml


	case "10"
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		sOrgID = Request("sOrgID")
		sFinyr = Request("sFinyr")
		sschedno = Request("sschedno")		
		SubID = Request("SubID")  	
		sql = "select BSSubID,BSSubSubID,EntryType,Computemode,BSSubHeadingName,Hierachy From Acc_M_BSSetupSubHeads where BSHeadID="&sschedno&" and FinYear='"&sFinyr&"' and BSSubID="&SubID&" And BSSubSubID <> 0"  
		'Response.Write sql
			With objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			If Not objRs1.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs1.EOF
					Set newElem = OutData.createElement("Details")
					If objrs1(2)="S" Then
					sql = "select BSSubHeadValue from Acc_T_BSAcDetail where BSHeadID='"&sschedno&"' and BSSubID='"&objrs1(0)&"' and BSSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF= True Then	NewCode = objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					If objrs1(2)="A" Then
						sql = "select ApplicableACHeadCode from Acc_T_BSAcDetail where BSHeadID='"&sschedno&"' and BSSubID='"&objrs1(0)&"' and BSSubSubID='"&objrs1(1)&"'"    
					With objRs
						.CursorLocation = 3
						.CursorType = 3
						.Source = sql
						.ActiveConnection = con
						.Open
					End with
					If Not objrs.EOF = True Then NewCode = objRs(0) Else NewCode = 0
					objrs.Close 
					End If
					newElem.setAttribute "BSSubID",objRs1(0)&","&objRs1(1)&","&objRs1(2)&","&objRs1(3)&","&NewCode&","&objRs1(5)   
					newElem.setAttribute "SubHeadingName",objRs1(4)  
					Root.appendChild newElem
				  	objRs1.MoveNext
				Loop
				Response.Clear
				Response.ContentType="text/xml"
				Response.Write OutData.xml
			End If
			objRs1.Close





				
	End Select
%>

