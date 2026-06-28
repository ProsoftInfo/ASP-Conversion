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
	'Program Name				:	DBGJView.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	N.Rajkumar
	'Created On					:	15th May 2003
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
dim objRs,objRs1,objRs2,objRs3,sQuery,iPageNo,isNo
dim sOrgId,sBookCode,iBookNo,sFromDate,sToDate,sBookName
dim iTransNo,sVouNo,sVouDate,sAccDescription,sOrgName,sAccHeadParam
dim iVocEntryNo,sAccUnitHead,sAccUnitPartyCode,sAccHeadDesc
dim sVovNarration,dAmount,sTransCrDrId,sCurrentDate,dOpeningAmt,iAccHead
dim dOpenBal,sOpenCrDr,sCloseCrDr,dCloseBal,dPayTotal,dRecTotal,sDisplayHead
dim bFlag,saTemp,sBook,sOptSel,iVouNoFrom,iVouNoTo,dGAmount,dLAmount,sVouType
dim sGetVal,sFinMonYear,sMonthDay,iCurrVouNo,sUnitDet,sUnitDesc
Dim sVouFromDate,sVouToDate,iAccHeadParam

sBookCode="08"
iPageNo=1
isNo=0

set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
set objRs3 =server.CreateObject("adodb.Recordset")

'----------- To Get The Values From the Selection Page ----------------
sFromDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
sToDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)

sGetVal=Request.QueryString("Value")
saTemp=split(sGetVal,"|")
sOrgId= saTemp(0)
sOrgName=saTemp(1)
sBookName=saTemp(2)
iBookNo= cstr(saTemp(3))
sOptSel=saTemp(4)
sVouType=saTemp(7)
sVouFromDate=saTemp(8)
sVouToDate=saTemp(9)
sFromDate=saTemp(8)
sToDate=saTemp(9)
iAccHeadParam =saTemp(12)
sAccHeadDesc=saTemp(14)

'--------------- Coding For The Option Voucher Number -----------------
Dim sOptSelArr,iCtr
sOptSelArr = Split(sOptSel,",")
For iCtr = LBound(sOptSelArr) to UBound(sOptSelArr)
	'0 sOrgId /1 sOrgName/2 sBook/3 iBookNo /4 sFlag /5 iVocNoFrom /6 iVocNoTo/7 sVouType
	'8 sFromDate/9 sToDate/10 dGAmount/11 dLAmount/12 sHead/13 sHeadCode/14 sHeadDesc
	
	If sOptSelArr(iCtr) ="VouNo" Then
		iVouNoFrom=saTemp(5)
		iVouNoTo=saTemp(6)
		sFromDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
		sToDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
		sDisplayHead="Voucher Number " & iVouNoFrom & " To " & iVouNoTo
	End IF

	if sOptSelArr(iCtr)="Amount" Then
		dGAmount=saTemp(10)
		dLAmount=saTemp(11)
		sFromDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
		sToDate="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	End IF
Next

If sOrgName = "" Then sOrgName = session("orgshortname")

'--------- Query for to Select the Account Head -----------------------

sQuery="SELECT BOOKACCOUNTHEAD FROM ACC_R_APPLICABLEACCOUNTHEADS WHERE BOOKCODE='" & sBookCode & "' AND BOOKNUMBER='"& iBookNo & "' AND OUDEFINITIONID='" & sOrgId & "'"
With objRs3
	.CursorLocation =3
	.CursorType=3
	.ActiveConnection=con
	.Source=sQuery
	.Open 
End with
set objRs3.ActiveConnection =nothing
If not objRs3.EOF then
	iAccHead =objRs3(0)
End if


objRs3.Close 

'----------Credit/Debit Indication For Opening and Closing ------------

if dOpeningAmt <0 then
	sOpenCrDr="CR"
	sCloseCrDr="CR"
	dOpeningAmt=dOpeningAmt*-1
else
	sOpenCrDr="DR"
	sCloseCrDr="DR"
end if	

'------------- Function Call For Getting Opening Amount ---------------

dOpeningAmt=FormatNumber(dOpeningAmt,2,,,0)
dCloseBal=dOpeningAmt
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>General Journal</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
</HEAD>
<script language="javascript">
function CloseWindow() {
	window.close();
}
</script>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">

	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<Tr>
		<TD class="FieldCell" class="PageTitle"><b>
		<p align=center>GJ Day Book For<%="  "&sBookName%> [ <%=sVouFromDate%> -  <%=sVouToDate%>]</td>
	</Tr>
                            
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            
                            <tr>
								<td align="center"></td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td>
												<TABLE BORDER="0" CELLSPACING=0 CELLPADDING=0>
													
												</TABLE>
											</td>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td>
												<div align="right">
													<TABLE BORDER="0" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD class="FieldCell" valign="bottom" width="0">Date</TD>
															<TD class="FieldCellSub" valign="bottom" width="0"><span class="DataOnly"><%=FormatDate(Date)%></span></TD>
														</TR>
													</TABLE>
												</div>
											</td>
										</tr>
									</table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td valign="top" width="100%">
									<div class="frmBody" id="frm2" style="width: 755; height:415;">
										<TABLE BORDER="0" CELLSPACING=1 CELLPADDING=0 width=100% class="ExcelTable">
											<tr>
												<TD class="ExcelHeaderCell" align="center">Account-Head</TD>
												<TD class="ExcelHeaderCell" align="center">A/C-Code</TD>
												<TD class="ExcelHeaderCell" align="center" rowspan="2">Debit</TD>
												<TD class="ExcelHeaderCell" align="center" rowspan="2">Credit</TD>
											</tr>
											<tr>
												<TD class="ExcelHeaderCell" align="center">Narration</TD>
												<TD class="ExcelHeaderCell" align="center">Vou-No</TD>
											</tr>
<%
	dim sTempDate, iCRAmount, iDRAmount, iCRSno, iDRSno
	
	dPayTotal=0
	dRecTotal=0
	iCRAmount=0
	iDRAmount=0
	iCRSno=0
	iDRSno=0
	bFlag=true

	'--- Selection of Transaction Entries Based On The Option Selected ----
	'--------------- Recordset For Transaction Entries --------------------

	sQuery = "Select TransactionNumber,VoucherNumber,convert(char,VoucherDate,103),TransactionType from  Acc_T_VoucherHeader"&_
		" where BookCode='"&sBookCode & "'and BookNumber="& iBookNo

	If sVouFromDate <> "" Then'If  sOptSel="VouDate" Then
		sQuery=sQuery + " and VoucherDate >= convert(datetime,'"& sFromDate & "',103) and "&_
				"VoucherDate <= convert(datetime,'"& sToDate &"',103) and OUDefinitionID='"&sOrgId&"'"	
	End IF

	If iVouNoFrom <> "" Then 'Elseif sOptSel="VouNo" Then
		sQuery =sQuery + " and VoucherNumber between '" & iVouNoFrom & "' and '" & iVouNoTo &"' and OUDefinitionID='"&sOrgId&"'"

	'Elseif sOptSel="AccHead" Then			
	'		sQuery =sQuery &" and AccountHead=" & sAccHeadParam

	'Elseif sOptSel ="Amount" Then
	'	sQuery=sQuery+ " and VoucherAmount>= "& dGAmount & " and VoucherAmount<=" & dLAmount&" and OUDefinitionID='"&sOrgId&"'"
	End if

	'------ Voucher Type Indication To Check Vocher Type in Query ---------

	If sVouType ="R" Then
		sVouType="GJ" & "R"
		sQuery=sQuery + " and TransactionType='" & sVouType & "'"
	End if

	sQuery =sQuery+ " order by VoucherDate"
		with objRs 
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection =con
			.Source =sQuery
			.Open 
		End With
		objRs.ActiveConnection =nothing
		
		Set iTransNo=objRs(0)
		Set sVouNo=objRs(1)
		Set sVouDate=objRs(2)

	'---- Setting Flag For Checking Current And Previous voucher Dates ----
	'------------- Displaying All the Transaction Details -----------------
	IF objRs.RecordCount >0 Then

	while not objRs.EOF
		if bFlag then
			iCurrVouNo=sVouNo
			bFlag=false
		end if	
	if iCurrVouNo<>sVouNo then

		if sCloseCrDr="CR" then
			dCloseBal=CDbl(dCloseBal)*-1
		end if
		
	'-------------- Updating Closing Balance For Each Entries -------------

			dCloseBal=CDbl(dCloseBal)+ CDbl(dRecTotal)-CDbl(dPayTotal)
	%>
			<tr>
				<TD COLSPAN=2 class="ExcelDisplayCell" align="Right"><P><b>Voucher Total&nbsp;</TD>
				<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dRecTotal,2,,,0)%></TD>
				<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dPayTotal,2,,,0)%></TD>
			</TR>			
			
		<%	iCurrVouNo=sVouNo
			iDRAmount = cdbl(iDRAmount) + cdbl(dRecTotal)
			iCRAmount = cdbl(iCRAmount) + cdbl(dPayTotal)
			dRecTotal=0
			dPayTotal=0
		end if	
			
		
		sQuery = "select VoucherEntryNumber, AccUnitAccountHead,AccUnitPartyType,"&_
			"AccUnitPartySubType,AccUnitPartyCode,VoucherNarration,Amount,"&_
			"TransCrDrIndication,AccountingUnit from Acc_T_VoucherDetails "&_
			"where TransactionNumber ="&iTransNo 
		
		If dGAmount <> "" Then 'If sOptSel ="Amount" Then
			sQuery=sQuery+ " and Amount>= "& dGAmount & " and Amount<=" & dLAmount
		End IF
		If iAccHeadParam <> "" Then 'Elseif sOptSel="AccHead" Then			
			sQuery =sQuery &" and AccUnitAccountHead  IN (" & sAccHeadParam &") "
		End if
				
		with objRs1
			.CursorLocation =3
			.CursorType=3
			.ActiveConnection =con
			.Source=sQuery
			.open 
		End with
		objRs1.ActiveConnection =nothing		
		
		set iVocEntryNo				=objRs1(0)
		set sAccUnitHead			=objRs1(1)
		'set sAccUnitPartyType		=objRs1(2)
		'set sAccUnitPartySubType	=objRs1(3)
		set sAccUnitPartyCode		=objRs1(4)
		set sVovNarration			=objRs1(5)
		set dAmount					=objRs1(6)
		set sTransCrDrId			=objRs1(7)
		set sUnitDet				=objRs1(8)

		while not objRs1.EOF
		
			if Trim(sUnitDet)<>Trim(sOrgId) Then
			
			sQuery="Select OrgUnitDescription,OrgUnitShortDescription from" &_
					" DCS_OrganizationUnitDefinitions where OUDefinitionID='" & sUnitDet & "'"
			with objRs2
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery 
				.Open 
			End with
			Set objRs2.ActiveConnection =nothing
			
			If not objRs2.EOF then
				sUnitDesc=objRs2(1)
			End if		
			objRs2.Close 
			
			End if
			if 	IsNull(sAccUnitHead) or IsEmpty(sAccUnitHead) then
				sQuery = "select PartyName from APP_M_PartyMaster where " &_
						" PartyCode ="&sAccUnitPartyCode&" "
				with objRs2
					.CursorLocation =3
					.CursorType=3
					.ActiveConnection =con
					.Source=sQuery
					.open 
				End with
				objRs2.ActiveConnection =nothing	
				if not objRs2.EOF then
					sAccDescription = objRs2(0)
				end if	
				objRs2.Close
			else
			
			'-------------- Query For Account Head Wise Display -------------------

				sQuery = "select AccountDescription from Acc_M_GLAccountHead where " &_
					   " AccountHead ="&sAccUnitHead&" "
				with objRs2
					.CursorLocation =3
					.CursorType=3
					.ActiveConnection =con
					.Source=sQuery
					.open 
				End with
				objRs2.ActiveConnection =nothing	
				if not objRs2.EOF then
					sAccDescription = objRs2(0)
				end if	
				objRs2.Close				
			end if	
	isNo =isNo +1			
	%>
	
	<%	if Trim(sTempDate) <> trim(sVouDate) then
			sTempDate = sVouDate
	%>
		<tr>
			<TD colspan="4" class="ExcelHeaderCell"><P><%=sTempDate%></TD>
		</TR>
	<%	end if	%>
	<tr>
		<TD class="ExcelDisplayCell"><P><%=sAccDescription%></TD>
	<% if IsNull(sAccUnitHead) or IsEmpty(sAccUnitHead) then	%>
		<TD class="ExcelDisplayCell"><P><%=sAccUnitPartyCode%></TD>
	<%	else	%>
		<TD class="ExcelDisplayCell"><P><%=sAccUnitHead%></TD>
	<%	end if	%>
	
	<%
	if sTransCrDrId="D" then
		dPayTotal=CDbl(dPayTotal)+CDbl(dAmount)
		iCRSno = iCRSno + 1 
	%>
		<TD class="ExcelDisplayCell" align="right" rowspan="2"><P><%=FormatNumber(dAmount,2,,,0)%></TD>
		<TD class="ExcelDisplayCell" align="right" rowspan="2">&nbsp;</TD>
	<%
	else
		dRecTotal=CDbl(dRecTotal)+CDbl(dAmount)
		iDRSno = iDRSno + 1
	%>
		<TD class="ExcelDisplayCell" align="right" rowspan="2"><P>&nbsp;</TD>
		<TD class="ExcelDisplayCell" align="right" rowspan="2"><%=FormatNumber(dAmount,2,,,0)%></TD>
	<%end if%>
	</tr>
	<TR>
		<TD class="ExcelDisplayCell"><P><%=sVovNarration%></TD>
		<TD class="ExcelDisplayCell"><P><%=sVouNo%></TD>
	</TR>	
	<%		objRs1.MoveNext		
			wend
			objRs1.Close
		objRs.MoveNext
	%>

	<%
	wend
	iDRAmount = cdbl(iDRAmount) + cdbl(dRecTotal)
	iCRAmount = cdbl(iCRAmount) + cdbl(dPayTotal)
	%>
	<tr>
		<TD COLSPAN=2 class="ExcelDisplayCell" align="Right"><P><b>Voucher Total&nbsp;</b></TD>
		<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dRecTotal,2,,,0)%></b></TD>
		<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dPayTotal,2,,,0)%></b></TD>
	</tr>
	<tr>
		<TD COLSPAN=4 class="ExcelHeaderCell" align="Right">&nbsp;</TD>
	</tr>
	<tr>
		<TD COLSPAN=4 class="ExcelDisplayCell" ><b>Total No. of Transactions : <%=isNo%></b></TD>
	</tr>
	<tr>
		<TD COLSPAN=1 class="ExcelDisplayCell" ><b>No. of Debit Transactions : <%=iDRSno%></b></TD>
		<TD COLSPAN=3 class="ExcelDisplayCell"><b>Total Debit Amount : <%=FormatNumber(iDRAmount,2,,,0)%></b></TD>
	</tr>
	<tr>
		<TD COLSPAN=1 class="ExcelDisplayCell" ><b>No. of Credit Transactions : <%=iCRSno%></b></TD>
		<TD COLSPAN=3 class="ExcelDisplayCell"><b>Total Credit Amount : <%=FormatNumber(iCRAmount,2,,,0)%></b></TD>
	</tr>
	<%End if%>
										</TABLE>
									</div>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
							                    <p align="center"> <input type="button" value="Ok" OnClick="CloseWindow()" class="ActionButton"  >
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack"></td>
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
