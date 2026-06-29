<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AmendUserCreationInsert.asp
	'Module Name				:	Admin (User Creation Amendment)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 29, 2004
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

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag,callFrom) {
		//alert(callFrom);
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to amend another User")) 
				if (callFrom == "FM"){
					window.location.href = "AmendUserCreationEntry.asp"}
				if (callFrom == "FG"){
					window.location.href = "ApplicationUserGrid.asp"}
			else
				window.location.href = "../welcome_admin.asp"

		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
		
	function msgboxxx(strr,flag) {
		
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to amend another User")) 
				window.location.href = "AmendUserCreationEntry.asp"
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
dim dcrs,iCount,sSql,sCallFrom,sUserType
dim sLoginID,sType,sPassword,sFName,sMName,sLName,sTitle,sAName,sSName,sEmployeeID,sDesig,sorgID
dim sWPhone,sWFax,sWEmail,sCell,sStreet,sCity,sState,sCountry,sPIN,sHPhone,sHFax,sHEmail,sNotes
dim bFlag

bFlag = False

sLoginID = trim(Request.QueryString("sLoginID"))
sorgID = trim(Request.QueryString("sOrgID"))
sCallFrom = trim(Request.QueryString("sCallFrom"))
sType = trim(Request.QueryString("sType"))

''Response.Write "<p>sData="&sLoginID & "--" & sorgID & "--"& sCallFrom & "--" & sType
	con.beginTrans

	if sType = "A" then

		'sPassword = trim(Request.Form("txtPassword"))
		sFName = trim(Request.Form("txtFName"))
		sMName = trim(Request.Form("txtMName"))
		sLName = trim(Request.Form("txtLName"))
		sTitle = trim(Request.Form("txtTitle"))
		sAName = trim(Request.Form("txtAName"))
		sSName = trim(Request.Form("txtSName"))
		sEmployeeID = trim(Request.Form("txtEmployeeID"))
		sDesig = trim(Request.Form("txtDesignation"))
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
		sUserType = trim(Request.Form("hUserType"))
		
		'Response.Write "<p>sUserType="&sUserType

		if sCountry = "select" then
			sCountry = "NULL"
		else
			sCountry = Pack(sCountry)
		end if

			
		sSql = "UPDATE DCS_USER SET " _
		& "USERNAME = " & Pack(sFName) & ",MIDDLENAME = " & Pack(sMName) & ",LASTNAME = " _
		& Pack(sLName) & ",TITLE = " & Pack(sTitle) & ",ALTERNATENAME = " & Pack(sAName) _
		& ",SHORTNAME = " & Pack(sSName) & ",UserCode = " & Pack(sEmployeeID) & ",DESIGNATION = " _
		& Pack(sDesig) & ",WORKPHONE = " & Pack(sWPhone) & ",WORKFAX = " & Pack(sWFax) _
		& ",WORKEMAIL = " & Pack(sWEmail) & ",CELLPHONE = " & Pack(sCell) & ",HOMESTREETADDRESS = " _
		& Pack(sStreet) & ",HOMECITY = " & Pack(sCity) & ",HOMESTATEPROVINCE = " & Pack(sState) _
		& ",HOMECOUNTRY = " & sCountry & ",HOMEPOSTALCODE = " & Pack(sPIN) & ",HOMEPHONE = " _
		& Pack(sHPhone) & ",HOMEFAX = " & Pack(sHFax) & ",HOMEEMAIL = " & Pack(sHEmail) _
		& ",NOTES = " & Pack(sNotes) & " , USERTYPE = " & pack(sUserType) & " WHERE OUDEFINITIONID = " & Pack(sorgID) & " AND LOGINID = " _
		& pack(sLoginID)

		con.Execute sSql

	elseif sType = "D" then
		
		on error resume next 
		
	'	sSql = "DELETE FROM MS_EMPLOYEEMASTER WHERE ORGANISATIONCODE = " & Pack(sorgID) & " AND LOGINID = " _
	'	& pack(sLoginID)
	'	
	'	con.Execute sSql
		
		sSql = "DELETE FROM DCS_USER WHERE OUDEFINITIONID = " & Pack(sorgID) & " AND LOGINID = " _
		& pack(sLoginID)
		
		con.Execute sSql
	end if
	if con.Errors.count <> 0 then
		dim iCounter
		bFlag = True
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			'Response.Write con.Errors(iCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.end
		con.CommitTrans
	end if

	con.close
	set con = nothing
	
	if sType = "A" then
%>
	<BODY onLoad = "msgbox('Login ID [<%=replace(sLoginID,"'","\'")%>] has been Amended Successfully','Y','<%=sCallFrom%>')">
<%	elseif sType="D" and not bFlag	then %>
	<BODY onLoad = "msgbox('Login ID [<%=replace(sLoginID,"'","\'")%>] has been Deleted Successfully','Y','<%=sCallFrom%>')">
<%	elseif sType="D" and bFlag	then %>
	<BODY onLoad = "msgbox('Login ID [<%=replace(sLoginID,"'","\'")%>] could not be Deleted','N','<%=sCallFrom%>')">
<%	end if	%>
