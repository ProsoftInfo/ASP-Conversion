<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PopMonBalance.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	May 19,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/Accpopulate.asp"-->
<%
dim sQuery,objRs
dim iAccHead,sOrgId,sToDate,sFromMonth,sYearMon,iSno
dim dOpeningAmt,sOpeningCrDr,dCrAmount,dDrAmount

set objRs  = server.CreateObject("adodb.recordset")

sOrgId=Request("orgid")
iAccHead=Request("Acchead")
sToDate=Request("TillDate")


if Len(sToDate)<6 then
	sToDate=left(sToDate,4)&"0"&Right(sToDate,1)
end if

dOpeningAmt=GetDayOpening(sOrgId,iAccHead,getFromFinDate)

sFromMonth=mid(getFromFinYear,3,4)&mid(getFromFinYear,1,2)

if dOpeningAmt< 0 then
	sOpeningCrDr="Cr"
	dOpeningAmt=dOpeningAmt*-1
else
	sOpeningCrDr="Dr"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
function PopDayBalance(sValue) {
	var parts = String(sValue || "").split("~");
	window.open("PopDayBalance.asp?orgid=" + encodeURIComponent(parts[0] || "") + "&Acchead=" + encodeURIComponent(parts[1] || "") + "&TillDate=" + encodeURIComponent(parts[2] || ""), "", "height=390,width=580,toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
}
var popDayBalance = PopDayBalance;
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Monthwise Balance
		</td>
    </tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
												<DIV class=frmBody id=frm1 style="width: 385; height:240;">
                                                <table border="0" cellspacing="1" class="ExcelTable">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center">Month</td>
                                        <td class="ExcelHeaderCell" align="center">Opening</td>
                                        <td class="ExcelHeaderCell" align="center">Receipt</td>
                                        <td class="ExcelHeaderCell" align="center">Payment</td>
                                        <td class="ExcelHeaderCell" align="center">Closing</td>
                                            </tr>
<%
	iSno=0
			sQuery="SELECT MONTHDRAMOUNT,MONTHCRAMOUNT,substring(MonthYear,3,4)+substring(MonthYear,1,2) FROM ACC_T_GLACCTRANSACTAMT "&_
				"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) > ="&sFromMonth& " and " &_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) <= "&sToDate

			with objRs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with
			Set objRs.ActiveConnection=nothing

			set dDrAmount=objRs(0)
			set dCrAmount=objRs(1)
			set sYearMon=objRs(2)
			While not objRs.EOF
				iSno=CInt(iSno)+1
%>
                                            <tr>
                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
                                        <td class="ExcelDisplayCell">
                                        <%
										dim sTemDate

											sTemDate= LastDayOfMonth(mid(sYearMon,1,4)&mid(sYearMon,5,2)) &"/"&mid(sYearMon,5,2)&"/"& mid(sYearMon,1,4)
											Response.Write "<a href=""#"" onClick=""PopDayBalance('"&sOrgId&"~"&iAccHead&"~"&sTemDate&"')"">"

											'Response.Write "<a href=""PopDayBalance.asp?OrgId="&sOrgId&"&Acchead="&iAccHead&"&TillDate="&sTemDate&""">"
											Response.Write MonthName(mid(sYearMon,5,2),true) &"-"& mid(sYearMon,1,4)&"</a>"

                                        %>
                                         </td>
                                        <td class="ExcelDisplayCell" align="right"><p align="right"><%=FormatNumber(dOpeningAmt,2,,,0)%>&nbsp;<%=sOpeningCrDr%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dDrAmount,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dCrAmount,2,,,0)%></td>
<%
	if sOpeningCrDr="Cr" then
		dOpeningAmt=dOpeningAmt*-1
	end if
	dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dDrAmount)-CDbl(dCrAmount)
	if  dOpeningAmt < 0 then
		sOpeningCrDr="Cr"
		dOpeningAmt=dOpeningAmt*-1
	else
		sOpeningCrDr="Dr"
	end if
%>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dOpeningAmt,2,,,0)%>&nbsp;<%=sOpeningCrDr%></td>
                                            </tr>
<%
					objRs.MoveNext
				wend
				objRs.Close
%>
                                                </table>
												</div>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                                <input type="button" value="Close" name="B2" class="ActionButton" onClick="window.close()" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</BODY>
</HTML>
