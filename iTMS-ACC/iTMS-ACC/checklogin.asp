<%@ Language=VBScript %>
<%
Response.Buffer = TRUE
Response.Clear
%>

<!--#include virtual="/include/DBConnectionLogin.asp"-->
<!--#include virtual="/include/Cypher.asp"-->

<HTML><HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<Title>Verifying your Password Please Wait...!</Title>
</HEAD>

<BODY>
<%

dim loginSuccess
loginSuccess = False

' Getting the value of the loginid, and password from the previous page.
dim loginid, password,sFinYear,sOrgID, sUserType, sORGShortName,sUserAccessMode
loginid  = Request.Form("loginid")

password = Request.Form("password")
redirectioncheck = Request.Form("check")
sFinYear = Request.Form("selFinYear")
sOrgID = Request.Form("selUnit")


dim login, recordsfound
login = Lcase(loginid)
recordsfound = false

dim rs, sql
dim pass, loginname

' The Query is used to Fetch the login name and password from the database.
Set rs = Server.CreateObject("ADODB.RecordSet")

    sql = "Select OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sOrgID 
    rs.Open sql,con
    if not rs.EOF then
        sORGShortName = trim(rs(0))
    end if
    rs.Close 
    

	'sql = " SELECT LOGINPASSWORD,EMPLOYEENUMBER,EMPLOYEENAME,ORGANISATIONCODE FROM MS_EMPLOYEEMASTER WHERE LOWER(LOGINID) = '"&login&"'"
	sql = " SELECT PASSWORD,InternalUserId,UserName,OUDEFINITIONID,USERTYPE,IsNull(UserAccessMode,'I') FROM DCS_USER WHERE LOWER(LOGINID) = '"&login&"'"

	rs.Open sql,con

   ' The following loop is used to retrieve a single record from the database for the loginname.

	if not rs.EOF Then
		pass = trim(rs(0))
		iEmpNo = trim(rs(1))
		username = trim(rs(2))
		organizationcode = trim(rs(3))
		sUserType = trim(rs(4))
		sUserAccessMode = trim(rs(5))
		recordsfound = true
	end if
	rs.Close

	session("sUserName")=username
	if(recordsfound and cstr(Lcase(trim(pass))) = cstr(Lcase(trim(password)))) Then
		PutSessionAndInsertLog()
	end if

	dim loginpassword
	' ***   calls the following FUNCTION and stores the value returned to the variable loginpassword	***
	'loginpassword = Decrypt(Lcase(pass))
	loginpassword = (Lcase(pass))

	' If the decrypted password from the database matches with the password given then
	' this loop is executed and sets the loginSuccess flag.
	if(cstr(Lcase(trim(password))) = cstr(Lcase(trim(loginpassword)))) Then
		loginSuccess = true
	end if

	' If the loginSuccess is set then redirect it to the respective menu page
	if (loginSuccess) Then
		PutSessionAndInsertLog()
		Response.Redirect("userHome.asp")
	else
		Session.Contents.RemoveAll()
		Response.Redirect("itmsLogin.asp?e=Y&red=" & redirectioncheck)
 	end if

con.close
set con = nothing

' ***	Puts the value of the loginid in session and write the log message. ***

function PutSessionAndInsertLog()
	Session.Contents.RemoveAll()

	Session("loginid")          = login
	Session("userid")           = iEmpNo
	Session("username")         = username
	Session("organizationcode") = sOrgID
	Session("employeenumber")   = iEmpNo
	Session("FinPeriod")		= sFinYear
	Session("UserType")			= sUserType
	Session("OrgShortName")		= sORGShortName
	Session("useraccessmode")   = sUserAccessMode
end function
%>
</BODY>
</HTML>