
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SchedInsert.asp
	'Module Name				:	Accounts 
	'Author Name				:	Kumar K A
	'Created On					:	Jan 05, 2007
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<%
	Dim oDOM,sMod,sName,sDesgDir

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	sMod=Request("Mod")
	sName=Request("Name")
	sDesgDir = trim(Request("ToDir"))
	
	oDOM.async = false
	oDOM.load(Request)
%>
<%
'Declaration
Dim sNo,sOrgId,sHead,sHiera,sApp,sFinyr,iSchId,PrevSchNo
Dim Root,iNo,Node1,sQry,objRs,iSchNumber,sInsDate

set objRs = Server.CreateObject("ADODB.Recordset")
	set Root = oDOM.documentElement
	If Root.HasChildNodes then
		For each Node1 in Root.childnodes
			If Node1.NodeName = "Schedule" then
				PrevSchNo = Node1.GetAttribute("PrevSchNo")
				sNo = Node1.getAttribute("SchedNo")
				sOrgId = Node1.getAttribute("OrgID")
				sHead = Node1.getAttribute("SchedHead")
				sHiera = Node1.getAttribute("SchedHiera")
				sApp =  Node1.getAttribute("SchedApp")
				sFinYr = Node1.getAttribute("SchedYear")
				sInsDate = Node1.getAttribute("InsDate")
			End If
		next
	End If

	sQry = "SELECT ScheduleNumber FROM dbo.Acc_M_SchdSetupHeads WHERE (ScheduleID<>'"&PrevSchNo&"') and (ScheduleNumber='"&sNo&"')"
	objRs.open sQry,con
		If Not objRs.EOF Then
			Response.Write "Scheule Number Already Exist...!"
			Response.End 
		End If	
		objRs.Close 
	sQry = "SELECT ScheduleNumber FROM dbo.Acc_M_SchdSetupHeads WHERE (ScheduleID<>'"&PrevSchNo&"') and (Hierarchy='"&sHiera&"')"
	objRs.open sQry,con
		If Not objRs.EOF Then
			 Response.Write "Hierarchy Already Exist...!"
			 Response.End 
		End If	
	objRs.Close 

If PrevSchNo <> 0 Then
		con.BeginTrans
		sQry = "UPDATE Acc_M_SchdSetupHeads SET ScheduleNumber = '"&sNo&"',ScheduleHeading = '"&Replace(sHead,"'","''")&"',ApplicableFor='"&sApp&"', Hierarchy = '"&sHiera&"' WHERE (ScheduleID ='"&PrevSchNo&"')"
		con.execute sqry
		If con.Errors.count <> 0 Then
			For iCounter=0 to con.Errors.count - 1
				Response.Write con.Errors(iCounter) &"<br>"
			Next
		Else
			Response.Clear
			con.CommitTrans
		End if				
Else		
	con.BeginTrans
	'To get Schedule ID
	sQry="select isnull(max(ScheduleID),0)+1 from Acc_M_SchdSetupHeads"
		objRs.open sQry,con
			iSchId=objRs(0)
		objRs.Close

		sQry = "Insert into Acc_M_SchdSetupHeads Values('"&iSchId&"','"&sOrgId&"','"&sNo&"','"&Replace(sHead,"'","''")&"','"&sHiera&"','"&sApp&"','"&sFinyr&"')"
			'Response.Write "Insert="& sQry &"<BR><BR>"
			Con.Execute sQry
			
	If con.Errors.count <> 0 Then
		con.RollbackTrans
		For iCounter=0 to con.Errors.count - 1
			Response.Write con.Errors(iCounter) &"<br>"
		Next
	Else
		Response.Clear
		con.CommitTrans
	End if		
	
End If 
%>