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
	'Program Name				:	PopDayBalance.asp
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

sOrgId=Request("orgid")
iAccHead=Request("Acchead")
sToDate=Request("TillDate")
sFromMonth="01/"&mid(sToDate,4,8)

set objRs  = server.CreateObject("adodb.recordset")
dOpeningAmt=GetDayOpening(sOrgId,iAccHead,sFromMonth)
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
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Daywise Balance
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
												<DIV class=frmBody id=frm1 style="width: 540; height:240;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center">Date</td>
                                        <td class="ExcelHeaderCell" align="center">Opening</td>
                                        <td class="ExcelHeaderCell" align="center">Receipt</td>
                                        <td class="ExcelHeaderCell" align="center">Payment</td>
                                        <td class="ExcelHeaderCell" align="center">Closing</td>
                                            </tr>
<%
dim bFlag,dAmount,sCurrentDate,sVouDate,sVouCRDR
	iSno=0
	bFlag=true
			sQuery="select voucheramount,crdrindication,convert(char,voucherdate,103) from Acc_T_CreatedVoucherHeader where " &_
			"OUDEFINITIONID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
			"convert(datetime,voucherdate,103) >= convert(datetime,'"& sFromMonth & "',103)" &_
			" and convert(datetime,voucherdate,103)<= convert(datetime,'"& sToDate&"',103)"

			with objRs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with
			Set objRs.ActiveConnection=nothing

			set dAmount=objRs(0)
			set sVouCRDR=objRs(1)
			set sVouDate=objRs(2)

			if not objRs.EOF then
				While not objRs.EOF
					if bFlag then
						sCurrentDate=sVouDate
						bFlag=false
					end if

					if 	sCurrentDate=sVouDate	then

						if sVouCRDR="C" then
							dCrAmount=CDbl(dCrAmount)+CDbl(dAmount)
						else
							dDrAmount=CDbl(dDrAmount)+CDbl(dAmount)
						end if
					else
						iSno=CInt(iSno)+1

%>
                                            <tr>
                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
                                        <td class="ExcelDisplayCell"><%=sCurrentDate%></td>
                                        <td class="ExcelDisplayCell" align="right"><p align="right"><%=FormatNumber(dOpeningAmt,2,,,0)%>&nbsp;<%=sOpeningCrDr%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dCrAmount,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dDrAmount,2,,,0)%></td>
<%
						if sOpeningCrDr="Cr" then
							dOpeningAmt=dOpeningAmt*-1
						end if
						dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dDrAmount)-CDbl(dCrAmount)

						if  dOpeningAmt < 0 then
							sOpeningCrDr="Cr"
							dOpeningAmt=dOpeningAmt*-1
						else
							sOpeningCrDr="Cr"
						end if
						dDrAmount=0
						dCrAmount=0
						sCurrentDate=sVouDate
%>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dOpeningAmt,2,,,0)%>&nbsp;<%=sOpeningCrDr%></td>
                                            </tr>
<%
					end if
					objRs.MoveNext
				wend
				iSno=CInt(iSno)+1

%>
                                            <tr>
                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
                                        <td class="ExcelDisplayCell"><%=sCurrentDate%></td>
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
<%
				end if
				objRs.Close
%>
                                            </tr>

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