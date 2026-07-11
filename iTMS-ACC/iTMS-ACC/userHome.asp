<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	userHome.asp
	'Module Name				:	Menu
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 15, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Jan 05,2010
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
<!-- #include file="include/DatabaseConnection.asp" -->
<!-- #include file="include/GetOrganization.asp" -->
<%
Dim sOrgID,sOrgName,dcrs,sQuery
set dcrs = Server.CreateObject("ADODB.Recordset")
sOrgID = Session("organizationcode")
sQuery = "Select OrgUnitShortDescription from DCS_OrganizationUnitDefinitions  where OUDefinitionID = '"& sOrgID  & "'"
'Response.Write sQuery
with dcrs
	.cursorlocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.open
end with
if not dcrs.eof then
	sOrgName = dcrs(0)
end if
dcrs.close
%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<Title>Welcome</Title>
<link rel="stylesheet" href="assets/styles/Standard.css" type="text/css">
<SCRIPT SRC="scripts/checkwidth.js"></SCRIPT>
<SCRIPT SRC="scripts/redirectToPage.js"></SCRIPT>
<SCRIPT SRC="scripts/Timer.js"></SCRIPT>
</head>
<body class="MainBack" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0  onload="onLoad()">
	<FORM id="form1id" method="POST" name="form1name">
		<table border="0" id="main" cellpadding="0" cellspacing="0" width=100% height=100%>
			<TR>
				<TD VALIGN="top" align="center" height="45" >
					<table border="0" cellpadding="0" cellspacing="0" width="100%" class="MainTitle">
						<tr>
							<td height="40"><b><%=GetOrganization()%></td>
							<td style="border-right: 1px solid #31659C" align="right"></td>
							<td align="right"><a href="http://www.prosoftinfo.com" target="_blank">
								<img src="assets/images/Prosoft_Logo.gif" width="35" height="35"></a>&nbsp;
							</td>
						</tr>
					</table>
				</TD>
			</TR>
			<TR>
				<TD VALIGN="TOP">
					<table border="0" id="tblBody" cellpadding="0" cellspacing="0" width=100% HEIGHT="100%" class="MainBack">
						<TR>
							<TD  valign="top" >
								<table border="0" cellspacing="0" width="100%" class="MainToolBarTable" cellpadding="0">
									<tr>
										<td align="Center" class="MainToolBarCell"><b>&nbsp;&nbsp;<%=FormatDateTime(now,1)%>&nbsp;<span id="timer"></span>&nbsp;&nbsp;&nbsp;</b>
                                        </td>
										<td align="Center" class="MainToolBarCell">Welcome <span style="text-decoration:none;font:bold;color:black"><%=Session("username")%></span><br><span style="text-decoration:none;font:bold;color:red">&nbsp;&nbsp;Financial Year - <%=Session("FinPeriod")%>&nbsp;</span></td>
										<td align="center" class="MainToolBarCell"><span style="text-decoration:none;font:bold;color:red"><b><%=sOrgName%></b></span></td>
                                        <td width="125" align="center" class="MainToolBarCell"><A style="text-decoration: none;color:black;font:bold" HREF="changePassword.asp" target="bodyFrame">Change Password</A></td>
                                        <td width="60" align="center" class="MainToolBarCell"><A style="text-decoration: none;color:black;font:bold" HREF="itmslogout.asp">Logout</A></td>
									</tr>
								</table>
								<IFRAME NAME="bodyFrame" ID=IFrame2 FRAMEBORDER=0 SCROLLING=AUTO SRC="./welcome_Welcome.asp" NORESIZE="RESIZE" width="100%" HEIGHT="690px"></IFRAME>
							</TD>
						</TR>
					</TABLE>
				</TD>

			</TR>
		</TABLE>
	</FORM>
</body>
</html>

<%
	' Function to populate Application
	Function populateApplication()
		' Declaration of variables
		Dim dcrs,sAppDesc,sAppCode,sAppPath
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME,APPLICATIONPATH FROM MS_APPLICATIONS ORDER BY APPLICATIONCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sAppCode = dcrs(0)
		set sAppDesc = dcrs(1)
		set sAppPath = dcrs(2)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sAppPath)&""">"&trim(sClassDesc)&" -- "&trim(sItemDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>

