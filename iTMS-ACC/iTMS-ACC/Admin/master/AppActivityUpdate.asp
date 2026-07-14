<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppActivityUpdate.asp
	'Module Name				:	Admin (Activity Creation)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	December 17, 2003
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
dim dcrs,sSql,objDOM,sType,Root,sTemp,sArr,Node
dim nRoleID,nAppCode,sAppName,nProcessCode,sProcessName,nOrderNo,iInternalPracticeID

sTemp = Request.QueryString("sPassData")
sArr  = Split(sTemp,":")
sType = Trim(sArr(0))
nRoleID = Trim(sArr(1))
nAppCode = Trim(sArr(2))
sAppName = Trim(sArr(3))

con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
objDOM.async = False
objDOM.load(Request)

Set Root = objDOM.documentElement

If sType = "PROCESS" Then
	sSql ="UPDATE MS_APPLICATIONS SET APPLICATIONNAME=" & Pack(sAppName) & " WHERE APPLICATIONCODE=" & nAppCode & " "
	'Response.Write sSql & "<BR>"
	con.Execute sSql
End IF

If Root.hasChildNodes Then
	For Each Node in Root.childNodes
		If Node.nodeName= "DETAILS" Then
			
			nProcessCode = Node.attributes.getNamedItem("PROCESSCODE").value
			sProcessName = Node.attributes.getNamedItem("PROCESSNAME").value
			nOrderNo     = Node.attributes.getNamedItem("ORDERNUMBER").value
			
			
			If sType = "PROCESS" Then
			
				sSql ="UPDATE Ms_ApplicationProcess SET ORDERNUMBER = "& nOrderNo &" WHERE APPLICATIONCODE=" & nAppCode & " and PROCESSCODE = "& nProcessCode &" "
			
			Elseif sType="PRACTICE" Then
			
				sSql ="UPDATE Ms_ApplicationProcess SET ORDERNUMBER = "& nOrderNo &",PROCESSNAME='"& sProcessName &"' WHERE APPLICATIONCODE=" & nAppCode & " and PROCESSCODE = "& nProcessCode &" "
			
			Elseif sType = "ADDPRACTICE" Then
				
				If sProcessName = "" Then Exit For
					
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(MAX(PROCESSCODE)+1,1) FROM MS_APPLICATIONPROCESS WHERE APPLICATIONCODE = " & nAppCode & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing

				if not dcrs.EOF then
					iInternalPracticeID = trim(dcrs(0))
				end if
				dcrs.Close
				
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT PROCESSCODE FROM MS_APPLICATIONPROCESS WHERE LOWER(PROCESSNAME) = " & Pack(lcase(sProcessName)) & " AND APPLICATIONCODE = " & nAppCode
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing

				If dcrs.EOF then
					sSql = " INSERT INTO MS_APPLICATIONPROCESS(APPLICATIONCODE,PROCESSCODE,PROCESSNAME,ORDERNUMBER) VALUES " &_ 
						   " (" & nAppCode & "," & iInternalPracticeID & "," & Pack(sProcessName ) & "," & nOrderNo & ")"
					'Response.Write sSql & "<BR>"
				Else
					sSql ="UPDATE Ms_ApplicationProcess SET ORDERNUMBER = "& nOrderNo &",PROCESSNAME='"& sProcessName &"' WHERE APPLICATIONCODE=" & nAppCode & " and PROCESSCODE = "& nProcessCode &" "
				End IF
				dcrs.Close 
				
			End IF
			'Response.Write sSql & "<BR>"
			con.Execute sSql
			
		End IF
	Next
End IF
		

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	'Response.End 
	con.CommitTrans
end if

con.close
set con = nothing
%>
