<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	welcome_welcome.asp
	'Module Name				:	Menu
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 15, 2003
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Welcome</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="scripts/rolloverout.js"></SCRIPT>
<Script>
function Init(sActive, sClosed) {
	var value = document.formname.hTemp.value;
	var userType = String(document.formname.hUserType.value);
	if (String(value).replace(/^\s+|\s+$/g, "") === "True") {
		alert("Financial Year Closed.Transactions Not allowed");
	}
	if (String(sActive) === "NA" && (userType === "AD" || userType === "SU")) {
		alert("Financial Year is Not Created!! ");
		location.href = "CreateFinYear.asp";
		return;
	}
	if (String(sActive) === "NA") {
		alert("Financial Year is Not Created!!, Please Contact Administrator ");
		history.back();
	}
}
</Script>
<script src="/Scripts/itms-modern-compat.js"></script>
</head>


<%
Dim dcrs,sTemp,sSysMon,sSysYr,sSysMonYr,sFinYear,sFinSt,sFinEnd,sCheckStDt,sCheckEndDt
Dim sQuery,sActive,sClosed,sCloseUnit
'Declaration of Objects
sTemp = Request.QueryString("sFinFun")
'Response.Write sTemp
sSysMon = Month(Date())
sSysYr = Year(Date())
sFinYear = Session("FinPeriod")

IF Len(sSysMon) = 1 Then
	sSysMon = "0"&sSysMon
End IF
sSysMonYr = sSysMon&sSysYr
sFinSt = Trim(Left(sFinYear,4))
sFinEnd = Trim(Left(sFinYear,4))
sFinSt = sFinSt&"04"
sFinEnd = sFinEnd&"03"
IF CInt(sSysMon) >= 4 and Cint(sSysMon) <= 12 Then
	sCheckStDt = "01/04/"&sSysYr
	sCheckEndDt = "31/03/"&Cint(sSysYr)+1
Else
	sCheckStDt = "01/04/"&Cint(sSysYr)-1
	sCheckEndDt = "31/03/"&Cint(sSysYr)
End IF

'sCheckStDt = "01/04/2007"

Set dcrs = Server.CreateObject("ADODB.RecordSet")

sQuery = "Select Active,IsNull(Closed,'Y') From Ms_FinancialPeriod Where  "&_
		 "Convert(Varchar,FromPeriod,103) = '"&sCheckStDt&"' "

 'Response.Write sQuery
dcrs.Open sQuery,con
IF Not dcrs.EOF Then
	sActive = dcrs(0)
	sClosed = dcrs(1)
Else
	sActive = "NA"
	sClosed = "NA"
End IF
dcrs.Close

'Response.Write sClosed
IF CStr(sClosed) = "N" Then
	sQuery = "Select OUDefinitionID From Ms_AuditorClosing Where Convert(Varchar,FromPeriod,103) = '"&sCheckStDt&"' "
	dcrs.Open sQuery,con
	Do While Not dcrs.EOF
		sCloseUnit = sCloseUnit &","&dcrs(0)
		dcrs.MoveNext
	Loop
	dcrs.Close
Else
	sCloseUnit = ""
End IF


Session("Active") = sActive
Session("Closed") = sClosed
Session("ClosedUnit") = sCloseUnit

Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3

	if ucase(Session("UserType")) <> "SU" or ucase(Session("UserType")) <> "AD" then
		.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME,APPLICATIONPATH FROM MS_APPLICATIONS WHERE APPLICATIONCODE IN (SELECT DISTINCT APPLICATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & trim(Session("userid")) & ") ORDER BY APPLICATIONCODE"
	else
		.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME,APPLICATIONPATH FROM MS_APPLICATIONS ORDER BY APPLICATIONCODE"
	end if
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
if not dcrs.eof then
    if dcrs.recordcount = 1 then
        if Session("UserType")="AU" then
            Response.Write "<" & "script>"&vbCrLf
            Response.Write "parent.location.href = '"& Replace(dcrs(2), "'", "\'") &"';"
            Response.Write vbCrLf & "</" & "script>"
        end if 'if Session("UserType")="AU" then
    end if
end if
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init('<%=sActive%>','<%=sClosed%>')">
<form method="post" name="formname" action="">
<input type=hidden name="hTemp" value="<%=sTemp%>">
<input type=hidden name="hlogid" value="<%=lcase(session("loginid"))%>">
<input type=hidden name="hUserType" value="<%=ucase(session("UserType"))%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="middle" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
		<p>&nbsp;
            <div align="center">
              <center>
			<TABLE id=Table16 cellSpacing="4" cellPadding=0 border=0 width="75%">
				<TR>
			<%	' Declaration of variables


				'Declaration of Objects

				if ucase(session("UserType")) = "SU" or ucase(session("UserType")) = "AD" then %>

					<TD class="ExcelHeaderCell" align="middle" height="22">
                        <A href="Admin/Index_admin.asp" target=_parent>ADMIN</A>

					</TD>

			<%
				end if
				do while not dcrs.EOF
			%>

					<TD valign="top">
						<table cellSpacing="0" cellPadding="0" border="0" width="100%">
							<tr>
								<td>
									<table cellSpacing="0" cellPadding="0" border="0" width="100%">
										<tr>
											<TD class="GroupTitleLeft">&nbsp;
											</TD>
											<TD class="GroupTitle" align="middle">
												<A href="<%=trim(dcrs(2))%>?AppCode=<%=trim(dcrs(0))%>" target=_parent><%=ucase(trim(dcrs(1)))%></A>

											</TD>
											<TD class="GroupTitleRight">&nbsp;
											</TD>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td Class="GroupTable">
									<table cellSpacing="2" cellPadding="0" border="0" width="100%">
										<tr>
											<td height="40px">
												&nbsp;
											</td>
										</tr>
										<tr>
											<td height=40>
												&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</TD>
			<%	dcrs.MoveNext
				loop
				dcrs.Close
			%>
				<!-- THE FOLLOWING HAS BEEN HARDCODED TO ADD QC LINK IN THE MENU LIST. CHANGE DONE BY SRIDHARAN. J. ON 29/12/2004 -->
				<!--TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">
                        <A href="../qualitycontrol/Index_QC.asp" target=_parent>QUALITY CONTROL</A>
					</TD></A>
				</TR-->
				<%if ucase(session("UserType")) = "SU" or ucase(session("UserType")) = "AD" then %>

						<TD class=ExcelHeaderCell align="middle"  height="22">
							<A href="CreateFinYear.asp">FINANCIAL YEAR CLOSING</A>
						</TD>

				<%End IF %>
				</TR>
				<!--<TR>
					<TD valign="top">
						<table cellSpacing="0" cellPadding="0" border="0" width="100%">
							<tr>
								<td>
									<table cellSpacing="0" cellPadding="0" border="0" width="100%">
										<tr>
											<TD class="GroupTitleLeft">&nbsp;
											</TD>
											<TD class="GroupTitle" align="middle">
												<A href="http://192.168.1.32:86/web-store/login.asp" target=_new>WEB STORE</A>
											</TD>
											<TD class="GroupTitleRight">&nbsp;
											</TD>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td Class="GroupTable">
									<table cellSpacing="2" cellPadding="0" border="0" width="100%">
										<tr>
											<td height=40>
												&nbsp;
											</td>
										</tr>
										<tr>
											<td height=40>
												&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</TD>
					<TD valign="top">
						<table cellSpacing="0" cellPadding="0" border="0" width="100%">
							<tr>
								<td>
									<table cellSpacing="0" cellPadding="0" border="0" width="100%">
										<tr>
											<TD class="GroupTitleLeft">&nbsp;
											</TD>
											<TD class="GroupTitle" align="middle">
												<A href="http://192.168.1.11:84/login.aspx" target=_new>EXTRANET</A>
											</TD>
											<TD class="GroupTitleRight">&nbsp;
											</TD>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td Class="GroupTable">
									<table cellSpacing="2" cellPadding="0" border="0" width="100%">
										<tr>
											<td height=40>
												&nbsp;
											</td>
										</tr>
										<tr>
											<td height=40>
												&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</TD>
					<TD valign="top">
						<table cellSpacing="0" cellPadding="0" border="0" width="100%">
							<tr>
								<td>
									<table cellSpacing="0" cellPadding="0" border="0" width="100%">
										<tr>
											<TD class="GroupTitleLeft">&nbsp;
											</TD>
											<TD class="GroupTitle" align="middle">
												<A href="http://192.168.1.32:9093" target=_new>EXTRANET OLD</A>
											</TD>
											<TD class="GroupTitleRight">&nbsp;
											</TD>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td Class="GroupTable">
									<table cellSpacing="2" cellPadding="0" border="0" width="100%">
										<tr>
											<td height=40>
												&nbsp;
											</td>
										</tr>
										<tr>
											<td height=40>
												&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</TD>
				</TR>-->
			</TABLE>
              </center>
            </div>
		</td>
	</tr>
</table>
</form>
</BODY>
