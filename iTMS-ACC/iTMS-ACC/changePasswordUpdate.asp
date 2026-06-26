<%@ Language=VBScript %>
<% Response.Buffer = true %>

<!-- #include File="include/DatabaseConnection.asp"-->

<HTML>
<HEAD>
<SCRIPT LANGUAGE=javascript>
<!--
function msgbox() {
	alert("Password has been Successfully Changed!")
	window.location.href="itmslogin.asp"
}
//-->
</SCRIPT>
</HEAD>
<LINK REL="STYLESHEET" HREF="assets/styles/StandardBody.css" TYPE="text/css">
<%
	password = trim(Request.Form("oldpass"))
	newpassword = trim(Request.Form("newpass"))
	loginname = lcase(trim(session("loginid")))

	set rs = Server.CreateObject("ADODB.RecordSet")

' The following query fetches the password for the corresponding username.
	sql = " SELECT PASSWORD FROM DCS_USER WHERE LOWER(LOGINID) = '"&loginname&"'"
	rs.Open sql,con
	if not rs.EOF Then
		pass = trim(rs(0))
	end if
	rs.Close


' Changing the default password to lower case obtained from the previous page.
	oldpasswd = cstr(Lcase(trim(password)))
' Decrypting the password obtained from the database.
	dbpasswd = cstr(Lcase(trim(pass)))

	cypherpassword = Lcase(newpassword)

	query1 = " UPDATE MS_EMPLOYEEMASTER SET LOGINPASSWORD = '"& cypherpassword &"' WHERE LOWER(LOGINID) = '"& loginname&"'"
	
	if(dbpasswd <> oldpasswd) Then
		Response.Redirect("ChangePassword.asp?e=Y")
	else
		con.Execute(query1)
		query1 = " UPDATE DCS_USER SET PASSWORD = '"& cypherpassword &"' WHERE LOWER(LOGINID) = '"& loginname&"'"
		con.Execute(query1)
		status = true
	End if

con.close
set con = nothing


%>
<BODY BGCOLOR="#FFFFFF" BACKGROUND="../assets/images/BG.jpg" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<%
if status = true then
Session.Abandon

Response.Buffer = TRUE
Response.Clear
%>
    <TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30" align="center" width="100%">
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD height="100">
            </TD>
        </TR>
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD align=center><FONT SIZE="3" COLOR="red" FACE="Verdana,Tahoma,Arial,Helvetica">Password has been Successfully Changed!</FONT></TD>
        </TR>
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD align=center height="10"></TD>
        </TR>
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD align=center>
				<FONT  COLOR="#010000" SIZE="2" FACE="Verdana,Tahoma,Arial,Helvetica">Click <A href="../itmslogin.asp" target="_parent">here</A> to login again!</FONT>
            </TD>
        </TR>
    </TABLE>
<%  
end if 
%>
</BODY>
</HTML>