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
<!-- #include file="include/DatabaseConnection.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Welcome</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="scripts/rolloverout.js"></SCRIPT>
<Script>
function Init(sActive, sClosed) {
	var value = document.formname.hTemp.value;
	if (String(value).replace(/^\s+|\s+$/g, "") === "True") {
		alert("Financial Year Closed.Transactions Not allowed");
	}
	if (String(sActive) === "NA" && String(document.formname.hlogid.value) === "admin") {
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

<script src="scripts/itms-modern-compat.js"></script>
</head>

<%
' Declaration of variables
Dim dcrs,sTemp,sSysMon,sSysYr,sSysMonYr,sFinYear,sFinSt,sFinEnd,sCheckStDt,sCheckEndDt
Dim sQuery,sActive,sClosed,sActMonYr,iCloseGlPar
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
	sActMonYr = "04"&sSysYr
Else
	sCheckStDt = "01/04/"&Cint(sSysYr)-1
	sCheckEndDt = "31/03/"&Cint(sSysYr)
	sActMonYr = "04"&sSysYr-1
End IF

'sCheckStDt = "01/04/2007"


Set dcrs = Server.CreateObject("ADODB.RecordSet")

sQuery = "Select Active,Closed From Ms_FinancialPeriod Where  "&_
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

if Cstr(sActive) = "Y" Then
	sActive = "NA"
End IF




'sQuery = "Select Count(1) From Acc_T_PartyOpeningAmt Where OpeningMonthYear = '"&sActMonYr&"' "
'dcrs.Open sQuery,con
'IF Not dcrs.EOF Then
'	iCloseGlPar = dcrs(0)
'Else
'	iCloseGlPar = 0
'End IF
'dcrs.Close




%>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="post" name="formname" action="">
<input type=hidden name="hTemp" value="<%=sTemp%>">
<input type=hidden name="hlogid" value="<%=lcase(session("loginid"))%>">
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
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="50%" bordercolor="#000000">
				<TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">
						<%IF CStr(sActive) = "NA" Then %>
							<A href="../admin/Master/FinancialPeriodEntry.asp">New Financial Period </A>
						<%Else%>
							New Financial Period
						<%End IF %>
					</TD></A>
				</TR>

				<TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">

							<A href="../admin/Master/CloseEntry.asp?Frm=GL">Account Heads Closing(GL & Party)</A>

					</TD></A>
				</TR>

				<!--TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">
						<%IF CStr(sActive) <> "NA" Then %>
							<A href="../admin/Master/CloseEntry.asp?Frm=PA">Party Closing </A>
						<%Else%>
							Party Closing
						<%End IF %>
					</TD></A>
				</TR-->

				<TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">
						<A href="../admin/Master/CloseEntry.asp?Frm=IS">Item Stock Closing </A>
					</TD></A>
				</TR>

				<TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">

							<A href="../admin/Master/CloseEntry.asp?Frm=NS">No Series Transfer </A>

					</TD></A>
				</TR>

				<TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">
						<%IF CStr(sActive) <> "NA" Then %>
							<A href="../admin/Master/CloseEntry.asp?Frm=SB">Schedule, Balance Sheet and P&L Transfer </A>
						<%Else%>
							Schedule, Balance Sheet and P&L Transfer
						<%End IF %>
					</TD></A>
				</TR>

				<TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">
						<%IF CStr(sClosed) <> "NA" Then %>
							<A href="../admin/Master/CloseEntry.asp?Frm=AG">Audit Closing </A>
						<%Else%>
							Audit Closing
						<%End IF %>
					</TD></A>
				</TR>

				<!--TR>
					<TD class=ExcelHeaderCell align="middle"  height="22">
						<%IF CStr(sActive) <> "NA" Then %>
							<A href="../admin/Master/CloseEntry.asp?Frm=AP">Audit Party Closing </A>
						<%Else%>
							Party GL Closing
						<%End IF %>
					</TD></A>
				</TR-->



			</TABLE>
              </center>
            </div>
		</td>
	</tr>
</table>
</form>
</BODY>

