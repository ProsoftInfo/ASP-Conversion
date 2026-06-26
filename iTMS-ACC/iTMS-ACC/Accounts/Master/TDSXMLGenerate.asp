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
	'Program Name				:	TDSXMLGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	KUMAR K A
	'Created On					:	Jan 19, 2007
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	TDSGroupingSetup.asp
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/Accpopulate.asp" -->
<%
Dim objRs,objRs1,objRs2,sql,id,OutData,Root,newElem
Dim sOrgID,GroupID,GroupHeadID,GroupHeadName,ComputeFormula
Set OutData = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")
set objRs2 = Server.CreateObject("ADODB.Recordset")
id = Request("id")
	Select Case id
	case "0"
		sOrgID = Request("sOrgID")
		sql = "Select GroupName,GroupID,GroupACHeadCode from ACC_M_TDSGroup where OUDefinitionID='"&sOrgID&"'"  
			With objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sql
				.ActiveConnection = con
				.Open
			End with
			Set objRs.Activeconnection = Nothing
			If Not objRs.EOF then
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				Do while not objRs.EOF
					Set newElem = OutData.createElement("Details")
						newElem.setAttribute "GroupName",Objrs(0)
						newElem.setAttribute "GroupID",Objrs(1)  '&","&objRs(2)  
						newElem.setAttribute "id",id
						Root.appendChild newElem
				  	objRs.MoveNext
				Loop
			End If
			objRs.Close
		Case "1"
			GroupID= Request("GroupID") 
			sql = "Select GroupHeadID,GroupHeadName,ComputeMode,AcHeadCode,Herarchy,ComputeFormula From Acc_M_TDSHeadcomputation where GroupID ="&GroupID
			With objRs 
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sql
				.Open 
			End With
				If Not objRs.EOF Then 
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
				While Not objRs.eof  
						Set newElem = OutData.createElement("Details")
						newElem.setAttribute "GroupHeadID",Objrs(0)
						newElem.setAttribute "GroupHeadName",Objrs(1)
						newElem.setAttribute "ComputeMode",Objrs(2)
						newElem.setAttribute "AcHeadCode",Objrs(3)
						newElem.setAttribute "Herarchy",Objrs(4)
						If objRs(5)<>"" Then 	
						newElem.setAttribute "Formula",Objrs(5)
						Else
						newElem.setAttribute "Formula",""
						End If
						Root.appendChild newElem
					objRs.MoveNext 
				Wend 
				objRs.Close 
				End If
				
		Case "2"
			GroupID= Request("GroupCode") 
			GroupHeadID = Request("HeadID")
			sql = "Select GroupHeadName,ComputeFormula   "&_
				  "From Acc_M_TDSHeadcomputation where GroupID = "&GroupID&" and GroupHeadID = "&GroupHeadID
			With objRs 
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sql
				.Open 
			End With
			GroupHeadName = objRs(0)
			If objrs(1)<> "" Then ComputeFormula = objrs(1) Else ComputeFormula=""
			objrs.Close 
			
			sql = "Select GroupHeadID,GroupHeadName,ComputeMode,AcHeadCode,Herarchy "&_
				  "From Acc_M_TDSHeadcomputation where GroupID = "&GroupID&" and GroupHeadID < "&GroupHeadID
			With objRs 
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sql
				.Open 
			End With
			sql = "Select GroupName from ACC_M_TDSGroup where GroupID='"&GroupID&"'"  
			With objrs1
				.CursorLocation = 3
				.CursorLocation = 3
				.ActiveConnection = con
				.Source = sql
				.Open 
			End With
				If Not objRs.EOF Then 
				Set Root = OutData.createElement("iXML")
				OutData.appendChild Root
				While Not objRs.eof  
						Set newElem = OutData.createElement("Details")
						newElem.setAttribute "GroupHeadID",Cint(Objrs(0))
						newElem.setAttribute "GroupHeadDetName",Objrs(1)
						newElem.setAttribute "ComputeMode",Objrs(2)
						newElem.setAttribute "AcHeadCode",Objrs(3)
						newElem.setAttribute "Herarchy",Objrs(4)
						newElem.setAttribute "Formula",ComputeFormula
						newElem.setAttribute "GroupName",Objrs1(0)
						newElem.setAttribute "GroupHeadName",GroupHeadName 
						Root.appendChild newElem
					objRs.MoveNext 
				Wend 
				Else
				Set Root = OutData.createElement("Root")
				OutData.appendChild Root
						Set newElem = OutData.createElement("Details")
						newElem.setAttribute "GroupHeadID",""
						newElem.setAttribute "GroupHeadDetName",""
						newElem.setAttribute "ComputeMode",""
						newElem.setAttribute "AcHeadCode",""
						newElem.setAttribute "Herarchy",""
						newElem.setAttribute "Formula",ComputeFormula
						newElem.setAttribute "GroupName",Objrs1(0)
						newElem.setAttribute "GroupHeadName",GroupHeadName 
						Root.appendChild newElem
				End If	
				objRs.Close 
				objRs1.Close 
	
			Case "3"
				GroupID= Request("GroupID") 
				GroupHeadID = Request("HeadID") 
			'	Response.Write GroupID  
			'	Response.Write GroupHeadID 
			'	Response.End 
				sql = "Select GroupHeadName  "&_
					  "From Acc_M_TDSHeadcomputation where GroupID="&GroupID&" and GroupHeadID="&GroupHeadID
			'	Response.Write sql
				With objRs 
					.CursorLocation = 3
					.CursorType = 3
					.ActiveConnection = con
					.Source = sql
					.Open 
				End With
				GroupHeadName = objRs(0)
				objrs.Close 
			
				sql = "Select GroupHeadID,GroupHeadName,ComputeMode,AcHeadCode,Herarchy,ComputeFormula  "&_
					  "From Acc_M_TDSHeadcomputation where GroupID = "&GroupID&" and GroupHeadID = "&GroupHeadID
				With objRs 
					.CursorLocation = 3
					.CursorType = 3
					.ActiveConnection = con
					.Source = sql
					.Open 
				End With
				sql = "Select GroupName from ACC_M_TDSGroup where GroupID='"&GroupID&"'"  
				With objrs1
					.CursorLocation = 3
					.CursorLocation = 3
					.ActiveConnection = con
					.Source = sql
					.Open 
				End With
					If Not objRs.EOF Then 
					Set Root = OutData.createElement("Root")
					OutData.appendChild Root
					While Not objRs.eof  
							Set newElem = OutData.createElement("Details")
							newElem.setAttribute "GroupHeadID",Cint(Objrs(0))
							newElem.setAttribute "GroupHeadDetName",Objrs(1)
							newElem.setAttribute "ComputeMode",Objrs(2)
							newElem.setAttribute "AcHeadCode",Objrs(3)
							newElem.setAttribute "Herarchy",Objrs(4)
							If objRs(5)<>"" Then 	
							newElem.setAttribute "Formula",Objrs(5)
							Else
							newElem.setAttribute "Formula",""
							End If
							newElem.setAttribute "GroupName",Objrs1(0)
							newElem.setAttribute "GroupHeadName",GroupHeadName 
							sql = "select AccountDescription From Acc_M_GLAccountHead where AccountHead='"&Objrs(3)&"'"
							With objRs2 
								.CursorLocation = 3
								.CursorType = 3
								.ActiveConnection = con
								.Source = sql
								.Open 
							End With
							If not objRs2.EOF  Then
							newElem.setAttribute "AccHeadName",objrs2(0)
							Else
							newElem.setAttribute "AccHeadName",""
							End If
							objRs2.Close 
							Root.appendChild newElem
						objRs.MoveNext 
					Wend 
					objRs.Close 
					objRs1.Close 
					End If	
			
			
			
			
			
			
			
			
			
				
				
	End Select
			

	Response.Clear
	Response.ContentType="text/xml"
	Response.Write OutData.xml



%>

