<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgUnitCreationInsert.asp	
	'Module Name				:	Inventory (Organization Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 12, 2002
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "OrgUnitDefinitionEntry.asp"
		}
		else if (flag == "Z") {
			alert(strr);
			window.location.href = "OrgCreationEntry.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
dim dcrs,sOrgId,sSql,inoLevel,imaxUnit
dim sorgUnitName

sorgUnitName = trim(Request.Form("txtUnitName"))

con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ORGANIZATIONID FROM DCS_ORGANIZATION"
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
set sOrgId = dcrs(0)
If Not dcrs.EOF Then
	sOrgId = sOrgId
	dcrs.Close
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(ORGANIZATIONUNITS,0) FROM DCS_ORGANIZATION"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	set inoLevel = dcrs(0)
	if not dcrs.EOF then
		inoLevel = inoLevel
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(ORGANIZATIONUNITID),0) + 1 FROM DCS_ORGANIZATIONUNITS"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	set imaxUnit = dcrs(0)
	if not dcrs.EOF then
		imaxUnit = imaxUnit
	end if
	dcrs.Close
	
	If cint(imaxUnit) <= cint(inoLevel) Then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ORGANIZATIONUNITNAME FROM DCS_ORGANIZATIONUNITS WHERE LOWER(ORGANIZATIONUNITNAME) = " & Pack(lcase(sorgUnitName)) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			sSql = "INSERT INTO DCS_ORGANIZATIONUNITS (ORGANIZATIONUNITID,ORGANIZATIONID," &_
				"ORGANIZATIONUNITNAME) VALUES " &_
				"(" & imaxUnit & "," & Pack(sOrgId) & "," & Pack(sorgUnitName) & ")"
			'Response.Write sSql & "<BR>"
			con.Execute sSql
%>
	<BODY onLoad = "msgbox('Organization Unit <%=replace(sorgUnitName,"'","\'")%> has been Created Successfully','Y')">
<%
		else
%>
	<BODY onLoad = "msgbox('Organization Unit <%=replace(sorgUnitName,"'","\'")%> already created','N')">
<%
		end if
	Else
%>
	<BODY onLoad = "msgbox('Unit cannot be created more than <%=inoLevel%> Organization level(s)','N')">
<%
	End If
Else
%>
	<BODY onLoad = "msgbox('No Organization Exists, Create Organization and then create Units','Z')">
<%
End If
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
