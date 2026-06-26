
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Change Password</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="scripts/logincheck.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<%
	check = trim(Request.QueryString("e"))
	loginname = lcase(trim(session("loginid")))

' The following logic is executed when the session has been expired.
	if (loginname="") then
		Response.Redirect ("SessionExpired.asp")
	end if

	%>

		<tr valign="TOP" align="LEFT">
			<td class="TopPack">
			</td>
		</tr>

		<tr valign="TOP" align="LEFT">
			<td >
				<p align="center" class=PageTitle ><b>Change Password</b>
			</td>
		</tr>

		<tr valign="TOP" align="LEFT">
			<td class="MiddlePack">
			</td>
		</tr>


		<tr valign="TOP" align="LEFT">
			<td >

				<form method="POST" action="changePasswordUpdate.asp" onsubmit="return(passcheck())">
				<div align="center">
					<center>
					<table border="0" cellspacing="1" class="ExcelTable" width="30%">
						<%	if check = "Y" then %>
						<tr>
							<td colspan="2">
								<p align="CENTER" ><font size="-2" face="Verdana,Tahoma,Arial,Helvetica" color="#BD1010">Please Check the Existing Password.
							</td>
						</tr>

						<%	end if %>
						<tr>
							<td class="ExcelDisplayCell">Old Password
							</td>
							<td class="ExcelInputCell">
								<input class="formelem" type="password" name="oldpass" size="20" maxlength="20">
							</td>
						</tr>

						<tr>
							<td class="ExcelDisplayCell">New Password
							</td>
							<td class="ExcelInputCell" width="10">
								<input class="formelem" type="password" name="newpass" size="20" maxlength="20">
							</td>
						</tr>

						<tr>
							<td class="ExcelDisplayCell">Confirm&nbsp; Password
							</td>
							<td class="ExcelInputCell">
								<input class="formelem" type="password" name="conpass" size="20" maxlength="20">
							</td>
						</tr>

					</center>

					<tr class="ActionCell">
						<td colspan="2"><center>
							<input type="submit" value="Change" name="B1" class="ActionButton" >
							&nbsp;<input type="reset" value="Reset" name="B2" class="ActionButton">
							</center>
						</td>
					</tr>

					</table>
				</div>
				</form>
			</td>
		</tr>

		<tr valign="TOP" align="LEFT">
			<td >
			</td>
		</tr>

</table>
</form>
</BODY>
