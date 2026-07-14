<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppVoucherList.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 28,2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<%
dim sOrgId,sBookId,iBookNo,sOrgName,sBookName
dim sFlag,sFromVal,sToVal
dim objRs,objRs1
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sPartyName

sFinPeriod = Session("FinPeriod")

IF CStr(sFinPeriod) = "" Then
	sFinPeriod = Year(Date)&":"&Year(Date)+1
End IF

IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
	sFromYr = "01/04/"&sFromYr
	sToYr = "31/03/"&sToYr
End IF

sOrgId=Request("selUnitId")
sBookId=Request("selVoucher")
iBookNo=Request("selBook")
sOrgName=Request("horgName")
sBookName=Request("hBookName")
sFlag=Request("optCriteria")



select case sFlag
	case "VouNo"
			sFromVal=Request("txtNoFrom")
			sToVal=Request("txtNoFrom")
	case "VouDate"
			sFromVal=Request("hFromDate")
			sToVal=Request("hToDate")
	case "Amount"
			sFromVal=Request("txtGAmount")
			sToVal=Request("txtLAmount")
	case "AccHead"
			sFromVal=Request("SelAccHead")
			sToVal=Request("hAccHead")
	Case "Exist"
		sOrgName=Session("OrgName")
		sBookName=Session("BookName")

		sFromVal=Session("FromValue")
		sToVal=Session("ToValue")
		sFlag=Session("Flag")
end select

Session("FromValue")=sFromVal
Session("ToValue")=sToVal
Session("Flag")=sFlag

Session("OrgName")=sOrgName
Session("BookName")=sBookName

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Approve Voucher For <%=sBookName%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="96">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Voucher List
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Voucher</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
												<DIV class=frmBody id=frm4 style="width: 585; height:340;">

                                        <%IF CStr(sBookId) = "01" or CStr(sBookId) = "02" or CStr(sBookId) = "08" Then %>
                                        <table border="0" cellspacing="1" class="ExcelTable" width="575">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="10"></td>
												<td class="ExcelHeaderCell" align="center" width="75" >Voucher No</td>
												<td class="ExcelHeaderCell" align="center">Voucher Date</td>
												<td class="ExcelHeaderCell" align="center">Type</td>
												<td class="ExcelHeaderCell" align="center">Amount</td>
												<td class="ExcelHeaderCell" align="center">Created By</td>

                                            </tr>
                                        <%Else%>
											<table border="0" cellspacing="1" class="ExcelTable" width="650">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="10"></td>
												<td class="ExcelHeaderCell" align="center" width="75" >Voucher No</td>
												<td class="ExcelHeaderCell" align="center">Voucher Date</td>
												<td class="ExcelHeaderCell" align="center">Party Name</td>
												<td class="ExcelHeaderCell" align="center">Amount</td>
												<td class="ExcelHeaderCell" align="center">Created By</td>
												<td class="ExcelHeaderCell" align="center">Type</td>
                                            </tr>
                                        <%End IF %>
<%
dim sQuery,iSno,iTransNo,sVouNo,sVouDate,dAmount,sType,iCreatedBy,sCreatedByName,iApproveLevel
dim saTemp
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")


sQuery="select Distinct a.CreatedTransNo,a.CreatedVoucherNo,convert(char,a.VoucherDate,103),a.VoucherAmount,"&_
		"a.TransactionType,a.CreatedBy,b.ApprovalLevel,isNull(a.BankInstrumentType,'Nil'), isNull(A.PartyCode,0),a.VoucherDate from Acc_T_CreatedVoucherHeader a,Acc_T_VouchersForApproval b "&_
		"where a.OUDefinitionID='"&sOrgId&"' and a.BookCode='"&sBookId&"' and a.BookNumber="&iBookNo&" and "&_
		" a.CreatedTransNo=b.CreatedTransNo and ToBeApprovedBy="&getUserId&" "&_
		" and CreatedVouchStatus = '010101' "&_
		"and convert(datetime,a.VoucherDate,103) >= convert(datetime,'"&sFromYr&"',103)  "&_
		"and convert(datetime,a.VoucherDate,103) <= convert(datetime,'"&sToYr&"',103) "

select case sFlag
	case "VouNo"
		sQuery=sQuery&" and a.CreatedVoucherNo>='"&sFromVal&"' and a.CreatedVoucherNo<='"&sToVal&"' "
	case "VouDate"
		sQuery=sQuery&" and a.VoucherDate>=convert(datetime,'"&sFromVal&"',103) and a.VoucherDate<=convert(datetime,'"&sToVal&"',103) "
	case "Amount"
		sQuery=sQuery&" and a.VoucherAmount>="&sFromVal&" and a.VoucherAmount<="&sToVal&" "
	case "AccHead"
			IF sFromVal="G" then
				sQuery=sQuery&" and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where AccUnitAccountHead="&sToVal&") "
			else
				saTemp=Split(sToVal,"?")
				sQuery=sQuery&" and a.CreatedTransNo in (select distinct CreatedTransNo from Acc_T_CreatedVoucherDetails where "&_
				" AccUnitPartyType ='"&Trim(saTemp(0))&"' and AccUnitPartySubType="&Trim(saTemp(1))&" and AccUnitPartyCode="&Trim(saTemp(3))&") "
			end if
end select
	sQuery=sQuery&" order by 10"





with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

set iTransNo = objRs(0)
set sVouNo = objRs(1)
set sVouDate = objRs(2)
set dAmount = objRs(3)
set sType = objRs(4)
set iCreatedBy= objRs(5)
set iApproveLevel= objRs(6)
iSno=1
If not objRs.EOF then
	Do While Not objRs.EOF
		sQuery="select EmployeeName from  Ms_EmployeeMaster where EmployeeNumber="&iCreatedBy
		with objRs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		sCreatedByName=objRs1(0)
		objRs1.Close
		if sBookId="08" then
			sQuery="select sum(Amount) from  Acc_T_CreatedVoucherDetails where TransCrDrIndication='C' and CreatedTransNo="&iTransNo
			with objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			dAmount =objRs1(0)
			objRs1.Close
		end if

		IF CStr(sBookId) = "04" or CStr(sBookId) = "05" or CStr(sBookId) = "06" or CStr(sBookId) = "07" Then
			sQuery = "Select PartyName From App_M_PartyMaster Where PartyCode = "&objRs(8)&" "
			with objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			sPartyName =objRs1(0)
			objRs1.Close
		End IF



%>
                                            <tr>
                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
<%
	select case sBookId
	case "01"
%>
     <td class="ExcelDisplayCell" align="center"><A href="AppCashVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%	case "02"%>
     <td class="ExcelDisplayCell" align="center"><A href="AppBankVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%	case "04"%>
	 <td class="ExcelDisplayCell" align="center"><A href="AppPurVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>

<%	case "05"%>
	 <td class="ExcelDisplayCell" align="center"><A href="AppSalVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>

<%	case "06"
		IF CStr(objRs(7) = "OT")  Then%>
			<td class="ExcelDisplayCell" align="center"><A href="AppDNOTVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		ELSEIF CStr(objRs(7) = "SI") Then %>
			<td class="ExcelDisplayCell" align="center"><A href="AppDNSalInvDisplay.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		ELSEIF CStr(objRs(7) = "SC") Then %>
			<td class="ExcelDisplayCell" align="center"><A href="AppCNVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		ELSEIF CStr(objRs(7) = "OP") Then %>
			<td class="ExcelDisplayCell" align="center"><A href="AppDNPurInvVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		Else%>
			<td class="ExcelDisplayCell" align="center"><A href="AppDNVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>


<%	End IF
	case "07"
		IF CStr(objRs(7) = "SC") Then 'Sales Commission
%>
			<td class="ExcelDisplayCell" align="center"><A href="AppCNVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
 <%		elseif CStr(objRs(7) = "SR") Then 'Sales Returns %>
			<td class="ExcelDisplayCell" align="center"><A href="AppCNSalReturnDisplay.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		elseif CStr(objRs(7) = "OS") Then 'Other Sales %>
			<td class="ExcelDisplayCell" align="center"><A href="AppCNSalInvDisplay.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		elseif CStr(objRs(7) = "PA") Then 'Purchase Invoices  %>
			<td class="ExcelDisplayCell" align="center"><A href="AppCNPurInvDisplay.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		Else 'Other %>
			<td class="ExcelDisplayCell" align="center"><A href="AppCNOTVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%		End IF%>

<%	case "08"%>
     <td class="ExcelDisplayCell" align="center"><A href="AppGJVoucherView.asp?TransNo=<%=iTransNo%>&AppLevel=<%=iApproveLevel%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>
<%	end select %>

    <td class="ExcelDisplayCell" align="left"><%=sVouNo%></td>

    <td class="ExcelDisplayCell" align="center"><p align="center"><%=sVouDate%></td>
<% IF CStr(sBookId) = "04" or CStr(sBookId) = "05" or CStr(sBookId) = "06" or CStr(sBookId) = "07" Then  %>
	<td class="ExcelDisplayCell" align="left"><%=sPartyName%></td>
<%End IF%>


<%
		IF CStr(sBookId) = "04" or CStr(sBookId) = "05" or CStr(sBookId) = "06" or CStr(sBookId) = "07" Then
			if CDbl(dAmount) > 0 then
%>
				<td class="ExcelDisplayCell" align="right">Rs.&nbsp;<%=FormatNumber(dAmount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="left"><%=sCreatedByName%></td>
<%			Else%>

				<td class="ExcelDisplayCell" align="right">Rs.&nbsp;</td>
				<td class="ExcelDisplayCell" align="left"><%=sCreatedByName%></td>
<%			End IF
		End IF



		if Right(sType,1)="R" and sBookId <> "08"   then
%>
			<td class="ExcelDisplayCell" align="left">Receipt</td>
<%		elseif Right(sType,1)="P" and sBookId<>"08" then %>
			<td class="ExcelDisplayCell" align="left">Payment</td>
<%		else %>
			<td class="ExcelDisplayCell" align="left"></td>
<%		End if	%>

<%
	IF CStr(sBookId) = "01" or CStr(sBookId) = "02" or CStr(sBookId) = "08" Then
		if CDbl(dAmount) > 0 then%>
			<td class="ExcelDisplayCell" align="right">Rs.&nbsp;<%=FormatNumber(dAmount,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="left"><%=sCreatedByName%></td>
<%		else %>
            <td class="ExcelDisplayCell" align="left"></td>
            <td class="ExcelDisplayCell" align="left"><%=sCreatedByName%></td>
<%		End if
	End IF%>
                                            </tr>
<%
		objRs.MoveNext
		iSno=CInt(iSno)+1
	LOOP
end if
objRs.Close

%>

                                                </table>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
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
</form>
</BODY>
</HTML>