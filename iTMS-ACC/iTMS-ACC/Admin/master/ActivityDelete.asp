<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ActivityDelete.asp
	'Module Name				:	Admin (Activity Creation)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	07 January 2011
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
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "ActivityCreationMain.asp"
		}
		else {
			alert(strr);
			window.location.href = "ActivityCreationMain.asp"
		}
	}
//-->
</SCRIPT>

<%
dim dcrs,dcrs1,sSql,objDOM,sTemp,sArr,nRoleID,nProcessCode,nApplicationCode,nActivityCode

sTemp = Request.QueryString("sData")
sArr  = Split(sTemp,":")

Set dcrs   = Server.CreateObject("ADODB.RecordSet")
Set dcrs1  = Server.CreateObject("ADODB.RecordSet")

nApplicationCode = Trim(sArr(1))
nProcessCode	 = Trim(sArr(2))
nActivityCode	 = Trim(sArr(3))
	
'Response.Write "<p>sTemp="&sTemp
con.beginTrans

	'sSql = " SELECT Distinct ROLEID FROM MS_ROLEACTIVITY WHERE APPLICATIONCODE="& nApplicationCode&" AND PROCESSCODE = "& nProcessCode &" and ACTIVITYCODE="& nActivityCode&" "
	sSql = " SELECT Distinct INTERNALUSERID,ROLEID FROM Ms_UserActivity WHERE APPLICATIONCODE="& nApplicationCode&" AND PROCESSCODE = "& nProcessCode &" and ACTIVITYCODE="& nActivityCode&" "
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	
	If dcrs.EOF Then
	
	    sSql = "Select Count(*) from MS_ApplicationActivityTemplates where APPLICATIONCODE="& nApplicationCode&" AND PROCESSCODE = "& nProcessCode &" and ACTIVITYCODE="& nActivityCode&" "
	    with dcrs1
		    .CursorLocation = 3
		    .CursorType = 3
		    .Source = sSql
		    .ActiveConnection = con
		    .Open
	    end with
	    if not dcrs1.eof then
	        if cint(dcrs1(0))=1 then
	            sSql = "Delete from MS_ApplicationActivityTemplates where APPLICATIONCODE="& nApplicationCode&" AND PROCESSCODE = "& nProcessCode &" and ACTIVITYCODE="& nActivityCode&" "
	            con.execute sSql
	        end if
	    end if
	    dcrs1.close
	

		sSql = "DELETE FROM MS_APPLICATIONACTIVITY WHERE APPLICATIONCODE="& nApplicationCode&" AND PROCESSCODE = "& nProcessCode &" and ACTIVITYCODE="& nActivityCode&" "
		'Response.Write "<P>sSql="&sSql
		con.Execute sSql
		
		%>
			<BODY onLoad = "msgbox('Selected Activity has been deleted Successfully','Y')">
		<%
		else
		%>
			<BODY onLoad = "msgbox('Activity is not deleted because its mapped to Role','N')">
		<%
	
		
	End IF
	dcrs.Close 

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

'Response.Redirect("ActivityCreationMain.asp")
%>
