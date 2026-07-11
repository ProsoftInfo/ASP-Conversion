<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	UserCreationInsert.asp
	'Module Name				:	Admin (User Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 29, 2003
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
			if (confirm("Do You want to define another User")) 
				window.location.href = "UserCreationEntry.asp"
			else
				window.location.href = "../welcome_admin.asp"

		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
dim dcrs,iCount,sSql,iInternalUserID,sUserType
dim sLoginID,sPassword,sFName,sMName,sLName,sTitle,sAName,sSName,sEmployeeID,sDesig,sorgID
dim sWPhone,sWFax,sWEmail,sCell,sStreet,sCity,sState,sCountry,sPIN,sHPhone,sHFax,sHEmail,sNotes,sUserAccessMode
Dim sPartyCode
'Response.write "<font color=red>"

sLoginID = trim(Request.Form("txtLoginID"))
sPassword = trim(Request.Form("txtPassword"))
sFName = trim(Request.Form("txtFName"))
sMName = trim(Request.Form("txtMName"))
sLName = trim(Request.Form("txtLName"))
sTitle = trim(Request.Form("txtTitle"))
sAName = trim(Request.Form("txtAName"))
sSName = trim(Request.Form("txtSName"))
sEmployeeID = trim(Request.Form("txtEmployeeID"))
sDesig = trim(Request.Form("txtDesignation"))
sorgID = trim(Request.Form("selUnit"))
sWPhone = trim(Request.Form("txtWorkPhone"))
sWFax = trim(Request.Form("txtWorkFax"))
sWEmail = trim(Request.Form("txtWorkEmail"))
sCell = trim(Request.Form("txtCell"))
sStreet = trim(Request.Form("txtStreet"))
sCity = trim(Request.Form("txtCity"))
sState = trim(Request.Form("txtState"))
sCountry = trim(Request.Form("selCountry"))
sPIN = trim(Request.Form("txtPostal"))
sHPhone = trim(Request.Form("txtHomePhone"))
sHFax = trim(Request.Form("txtHomeFax"))
sHEmail = trim(Request.Form("txtHomeEmail"))
sNotes = trim(Request.Form("txtNotes"))
sUserType = Trim(Request.Form("hUserType"))
sUserAccessMode = trim(Request.Form("hUAM"))
sPartyCode = trim(Request.Form("hCode"))

if trim(sPartyCode)="" or IsNull(sPartyCode) then sPartyCode = "0"

if sCountry = "select" then
	sCountry = "NULL"
else
	sCountry = Pack(sCountry)
end if

con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ISNULL(MAX(InternalUserID)+1,1) FROM DCS_USER"
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	iInternalUserID = trim(dcrs(0))
end if
dcrs.Close

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT LOGINID FROM DCS_USER WHERE LOWER(LOGINID) = " & Pack(lcase(sLoginID)) & " OR LOWER(UserCode) = " & Pack(lcase(sEmployeeID)) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if dcrs.EOF then

	sSql = "INSERT INTO DCS_USER(InternalUserID,OUDEFINITIONID,LOGINID,PASSWORD,USERNAME,MIDDLENAME,LASTNAME,TITLE," &_
		"ALTERNATENAME,SHORTNAME,USERCODE,DESIGNATION,WORKPHONE,WORKFAX,WORKEMAIL,CELLPHONE," &_
		"HOMESTREETADDRESS,HOMECITY,HOMESTATEPROVINCE,HOMECOUNTRY,HOMEPOSTALCODE,HOMEPHONE,HOMEFAX,HOMEEMAIL,NOTES,USERTYPE,USERACCESSMODE,PartyCode) VALUES " &_ 
		"(" & iInternalUserID & "," & Pack(sorgID) & "," & Pack(sLoginID) & "," & Pack(sPassword) & ", " &_
		" " & Pack(sFName) & "," & Pack(sMName) & "," & Pack(sLName) & "," & Pack(sTitle) & ", " &_
		" " & Pack(sAName) & "," & Pack(sSName) & "," & Pack(sEmployeeID) & "," & Pack(sDesig) & ", " &_
		" " & Pack(sWPhone) & "," & Pack(sWFax) & "," & Pack(sWEmail) & "," & Pack(sCell) & ", " &_
		" " & Pack(sStreet) & "," & Pack(sCity) & "," & Pack(sState) & "," & sCountry & ", " &_
		" " & Pack(sPIN) & "," & Pack(sHPhone) & "," & Pack(sHFax) & ", " &_
		" " & Pack(sHEmail) & "," & Pack(sNotes) & "," & pack(sUserType ) & ",'"& sUserAccessMode &"',"& sPartyCode &")"
	Response.Write sSql & "<BR>"
	con.Execute sSql

%>
	<BODY onLoad = "msgbox('Login ID [<%=replace(sLoginID,"'","\'")%>] has been Created Successfully','Y')">
<%
else
%>
	<BODY onLoad = "msgbox('Login ID or Employee ID already Exists','N')">
<%
end if
dcrs.Close

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	Response.End 
	con.CommitTrans
end if

con.close
set con = nothing
%>
