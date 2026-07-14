<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	RoleActivityMappingInsert.asp
	'Module Name				:	Admin (Role Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	December 10, 2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/populate.asp"-->

<%
Dim dcrs,sSql,iInternalRoleID,sRoleDesc,sPassType,objDOM,Root,iRoleID,sUser,sOrgUnit
Dim Node,SubNode,nAppCode,sAppName,nProcessCode,sProcessName,nActivityCode,sActivityName
Dim sItemTypeID,sTemp,sArr,nRoleID,nTempNo,iTemplateNo  

sTemp = Trim(Request("sPassData"))
sArr  = Split(sTemp,":")

sPassType = Trim(sArr(0))

If sPassType = "DEL" Then
	nProcessCode = sArr(1)
	nActivityCode = sArr(2)
	nAppCode = sArr(3)
	nRoleID = sArr(4)
	nTempNo = sArr(5)
Elseif sPassType = "ADD" Then
	nRoleID = sArr(1)
End IF

Set dcrs = Server.CreateObject("ADODB.RecordSet")

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
objDOM.async = False
objDOM.load(Request)

objDOM.Save server.MapPath("../xmldata/Activity.xml")
'Response.End 

'Response.Write "<p>sPassType="&sPassType

sUser = 2 ' For Test
sOrgUnit = Session("organizationcode")
'sItemTypeID = "STO"

Set Root     = objDOM.documentElement

con.beginTrans

'Response.Write "<p>sPassType"&sPassType

If sPassType = "ADDAPPUSERROLE" Then	'Call From Pop up Program
	
	'iRoleID = Root.attributes.getNamedItem("RoleID").value
	sUser   = Root.attributes.getNamedItem("UserID").value
	sItemTypeID = Root.attributes.getNamedItem("ItemType").value
	
	If Root.hasChildNodes Then
		
		For Each Node in Root.childNodes 
			If Node.nodeName = "ACTIVITY" Then
	
				nAppCode = Node.attributes.getNamedItem("APPCode").value
				nProcessCode  = Node.attributes.getNamedItem("PRCode").value
				nActivityCode = Node.attributes.getNamedItem("ACCode").value
				iRoleID = Node.attributes.getNamedItem("RoleID").value
				iTemplateNo = Node.Attributes.getNamedItem("TempNo").value
				
				'with dcrs
				'	.CursorLocation = 3
				'	.CursorType = 3
				'	.Source = "SELECT INTERNALUSERID FROM MS_USERROLES WHERE INTERNALUSERID = " & sUser & " AND ROLEID = " & iRoleID & ""
				'	.ActiveConnection = con
				'	.Open
				'end with
				'set dcrs.ActiveConnection = nothing

				'if dcrs.EOF then

				'	sSql = "INSERT INTO MS_USERROLES(INTERNALUSERID,ROLEID) VALUES " &_ 
				'		"(" & sUser & "," & iRoleID & ")"
					'Response.Write sSql & "<BR>"
				'	con.Execute sSql
				'end if
				'dcrs.Close
		
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT ROLEID FROM MS_ROLEACTIVITY WHERE ROLEID = " & iRoleID & " AND APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & " AND ACTIVITYCODE = " & nActivityCode & ""
					
					'.Source = "SELECT INTERNALUSERID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & sUser & " AND APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & " AND ACTIVITYCODE = " & nActivityCode & " "
					.Source = "SELECT INTERNALUSERID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & sUser & " AND APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & " AND ACTIVITYCODE = " & nActivityCode & " AND ROLEID = "& iRoleID &" and ActivityTemplateNo =  "& iTemplateNo
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				
				'Response.Write "<p>sq="&dcrs.Source 
				
				if dcrs.EOF then

					'sSql = "INSERT INTO MS_ROLEACTIVITY(ROLEID,APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE) VALUES " &_ 
					'	"(" & iRoleID & "," & nAppCode & "," & nProcessCode & "," & nActivityCode & ")"
					
					sSql = " INSERT INTO MS_USERACTIVITY(INTERNALUSERID,APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE,ORGANISATIONCODE,ITEMTYPEID,ROLEID,ActivityTemplateNo) VALUES " &_
						   " (" & sUser & "," & nAppCode & "," & nProcessCode & "," & nActivityCode & "," & Pack(sOrgUnit) & "," & Pack(sItemTypeID) & ","& iRoleID &","& iTemplateNo &" )"
					'Response.Write sSql & "<BR>"
					con.Execute sSql
				end if
				dcrs.Close
	
			End IF
		Next
	
	End IF	'If Root.hasChildNodes Then

Elseif sPassType = "DELAPPUSERROLE" Then
	
	'iRoleID = Root.attributes.getNamedItem("RoleID").value
	sUser   = Root.attributes.getNamedItem("UserID").value
	
	If Root.hasChildNodes Then
		
		For Each Node in Root.childNodes 
			If Node.nodeName = "ACTIVITY" Then
				
				iRoleID = Node.attributes.getNamedItem("RoleID").value
				nAppCode = Node.attributes.getNamedItem("APPCode").value
				nProcessCode  = Node.attributes.getNamedItem("PRCode").value
				nActivityCode = Node.attributes.getNamedItem("ACCode").value
				iTemplateNo = Node.attributes.getNamedItem("TempNo").value
				
				'sSql = "DELETE FROM MS_USERACTIVITY WHERE INTERNALUSERID = "& sUser&" and APPLICATIONCODE = "& nAppCode &" and PROCESSCODE = "& nProcessCode&" and ACTIVITYCODE in ("& nActivityCode &") "
				sSql = "DELETE FROM MS_USERACTIVITY WHERE INTERNALUSERID = "& sUser&" and APPLICATIONCODE = "& nAppCode &" and PROCESSCODE = "& nProcessCode&" and ACTIVITYCODE in ("& nActivityCode &") AND ROLEID = "& iRoleID&" and ActivityTemplateNo = "& iTemplateNo
				'Response.Write "<p>sql="&sSql
				con.Execute sSql
				
			End IF
		Next
	End IF
		
ElseIf sPassType = "ADD" Then						
	
	
	If Root.hasChildNodes then
		For Each Node in Root.childNodes
			If Node.NodeName = "ACTIVITYMAPPING" Then
					
				nAppCode = Node.attributes.getNamedItem("APPCODE").value
				sAppName = Node.attributes.getNamedItem("APPNAME").value
				nProcessCode = Node.attributes.getNamedItem("PROCESSCODE").value
				sProcessName = Node.attributes.getNamedItem("PROCESSNAME").value
					
				For Each SubNode in Node.childNodes
					If SubNode.nodeName= "ACTIVITY" Then

						nActivityCode = SubNode.attributes.getNamedItem("CODE").value
						sActivityName = SubNode.attributes.getNamedItem("NAME").value
						nTempNo = SubNode.Attributes.getNamedItem("TCODE").value
						
						
						with dcrs
							.CursorLocation = 3
							.CursorType = 3
							'.Source = "SELECT INTERNALUSERID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & sUser & " AND APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & " AND ACTIVITYCODE = " & nActivityCode & " "
							.Source = "SELECT Distinct RoleID From MS_ROLEACTIVITY Where RoleID = "& nRoleID &" AND APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & " AND ACTIVITYCODE = " & nActivityCode & " and ActivityTemplateNo =  "& nTempNo
							.ActiveConnection = con
							.Open
						end with
						set dcrs.ActiveConnection = nothing

						if dcrs.EOF then
							'sSql = "INSERT INTO MS_USERACTIVITY(INTERNALUSERID,APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE,ORGANISATIONCODE,ITEMTYPEID) VALUES " &_
							'	"(" & sUser & "," & nAppCode & "," & nProcessCode & "," & nActivityCode & "," & Pack(sOrgUnit) & "," & Pack(sItemTypeID) & ")"
							
							sSql = "INSERT INTO MS_ROLEACTIVITY(ROLEID,APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE,ActivityTemplateNo) VALUES " &_
								"(" & nRoleID & "," & nAppCode & "," & nProcessCode & "," & nActivityCode & ","& nTempNo &" )"
							
							'Response.Write sSql & "<BR>"
							con.Execute sSql
						end if
						dcrs.Close
							
					End IF	'If SubNode.nodeName= "ACTIVITY" Then
				Next
					
			End If
		Next
	End IF	'If Root.hasChildNodes then

Else
	'sSql = "DELETE FROM MS_USERACTIVITY WHERE APPLICATIONCODE = "& nAppCode &" and PROCESSCODE = "& nProcessCode&" and ACTIVITYCODE in ("& nActivityCode &") "
	sSql = "DELETE FROM MS_ROLEACTIVITY WHERE APPLICATIONCODE = "& nAppCode &" and PROCESSCODE = "& nProcessCode&" and ACTIVITYCODE in ("& nActivityCode &") AND ROLEID ="& nRoleID&" and ActivityTemplateNo in ( "& nTempNo &")"
	'Response.Write "<p>sql="&sSql
	con.Execute sSql
End If 'If sPassType = "ADD" Then

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
