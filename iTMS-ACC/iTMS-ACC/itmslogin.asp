<!-- #include File="include/DBConnectionLogin.asp" -->

<%
'make sure that client is authenticated
'If Len(Trim(CStr(Request.ServerVariables("LOGON_USER")))) = 0 Then
'	Response.Status = "401 Access Denied"
'	Response.End
'End If

Response.Buffer = TRUE
Response.Clear
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<LINK REL="STYLESHEET" HREF="assets/styles/StandardBody.css" TYPE="text/css">
<TITLE>Login</TITLE>
<SCRIPT LANGUAGE="JavaScript" SRC="scripts/logincheck.js"></SCRIPT>
<SCRIPT>
<!--
function CheckLogin() {
	Pass = document.forms[0].password.value;
	if (!loginidcheck()) {
		return false;
	}
	else if (trimTrue(Pass) == "") {
		alert("Enter Password");
		document.forms[0].password.select();
		return false;
	}
	else
		return true;
}

function focusAt() {
	window.document.forms[0].elements[0].focus();
}

//-->
</SCRIPT>
</HEAD>
<%
	check = trim(Request.QueryString("e"))
%>
<BODY BGCOLOR="#336699" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 onLoad="focusAt()">
	<TABLE CELLSPACING="0" CELLPADDING="0" ALIGN="CENTER" border="0">
		<TR VALIGN=TOP ALIGN=LEFT>
			<TD HEIGHT="83px" colspan="7"><IMG SRC="../assets/images/autogen/clearpixel.gif" WIDTH=204 HEIGHT=1 BORDER=0></TD>
		</TR>
		<tr>
			<td>
				<TABLE CELLSPACING="0" CELLPADDING="0" ALIGN="CENTER" class="TabBodyWithTopLine">
					<TR VALIGN=TOP ALIGN=LEFT>
						<TD COLSPAN="7">
						  <table cellpadding="0" cellspacing="0" width="100%">
							<!--tr>
								<td>
									<table cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="GroupTitleLeft">&nbsp;
											</td>
											<td class="GroupTitle" align="center">Login
											</td>
											<td class="GroupTitleRight">&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr-->
							<tr>
								<td>
									<table cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="TopPack">
											</td>
										</tr>
										<tr>
											<td>&nbsp;
											</td>
											<td class="PageTitle">Login
											</td>
											<td>&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<!--td class="GroupTable"-->
								<td>
									<FORM NAME="Table1FORM" ACTION="checklogin.asp" METHOD=POST onSubmit="return(CheckLogin())">
										<TABLE BORDER=0 CELLSPACING="3" CELLPADDING="1">
										<%	if check = "Y" then %>
											<TR>
												<TD COLSPAN=3 ><P ALIGN=CENTER><FONT SIZE="-2" FACE="Verdana,Tahoma,Arial,Helvetica" COLOR=#BD1010>Invalid Loginid or Password.</FONT></TD>
											</TR>
										<%	end if %>
											<TR>
												<TD WIDTH="150px" class="FieldCell">Login Id</TD>
												<TD WIDTH="6px" class="FieldCell">:</TD>
												<TD WIDTH="257px"><INPUT TYPE=TEXT NAME="Loginid" VALUE="" SIZE=20 MAXLENGTH=15 class="FormElem">&nbsp;</TD>
											</TR>
											<TR>
												<TD WIDTH="150px" class="FieldCell">Password</TD>
												<TD WIDTH="6px" class="FieldCell">:</TD>
												<TD WIDTH="257px"><INPUT TYPE="password" NAME="password" VALUE="" SIZE=20 MAXLENGTH=20 class="FormElem">&nbsp;</TD>
											</TR>
											<TR>
												<TD WIDTH="150px" class="FieldCell">Organisation Units</TD>
												<TD WIDTH="6px" class="FieldCell">:</TD>
												<TD WIDTH="257px">
													<select size="1" name="selUnit" class="FormElem">
													<%
														'Calling the Function which populates the Financial Year list
														populateUnit()
													%>
													</select>
												</TD>
											</TR>
											<TR>
												<TD WIDTH="150px" class="FieldCell">Financial Year</TD>
												<TD WIDTH="6px" class="FieldCell">:</TD>
												<TD WIDTH="257px">
													<select size="1" name="selFinYear" class="FormElem">
													<%	'Calling the Function which populates the Financial Year list
														populateFinYear
													%>
													</select>
												</TD>
											</TR>
											<TR>
												<TD WIDTH="150px" class="FieldCell"></TD>
												<TD WIDTH="6px" class="FieldCell"></TD>
												<TD WIDTH="257px" class="FieldCell">
													<a href="">Forgot password?</a>
												</TD>
											</TR>
											<TR>
												<TD valign="top" colspan="3">
													<table border="0" cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="ActionCell">
																<INPUT TYPE=SUBMIT NAME="FormsButton2" class="ActionButton" VALUE="Login">&nbsp;&nbsp;
																<INPUT TYPE=RESET NAME="FormsButton3" class="ActionButton" VALUE="Reset"></TD>
															</td>
														</tr>
													</table>
											</TR>
										</TABLE>
									</FORM>
								</td>
							</tr>
						  </table>
						</TD>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
</BODY>
</HTML>

<%
	' Function which populates the Previous Financial Year Start Date list
	Function populateFinYear()
		' Declaration of variables
		Dim dcrs,sStartDate,sEndDate
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CONVERT(CHAR,FROMPERIOD,103),CONVERT(CHAR,TOPERIOD,103) FROM MS_FINANCIALPERIOD ORDER BY 1 DESC"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sStartDate = dcrs(0)
		set sEndDate = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&right(trim(sStartDate),4)&":"&right(trim(sEndDate),4)&""">"&right(trim(sStartDate),4)&"-"&right(trim(sEndDate),4)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
	'*********************************************
	Function populateUnit()
		Dim dcrs, sUnitLID,sUnitLName,sUnitSName
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,OrgUnitShortDescription FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) ORDER BY ORGANIZATIONUNITID"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUnitLID = dcrs(0)
		set sUnitLName = dcrs(1)
		set sUnitSName = dcrs(3)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sUnitLID)&""">"&trim(sUnitSName)&"</OPTION>")
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
	End Function

%>

