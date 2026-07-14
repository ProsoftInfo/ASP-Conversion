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
	'Program Name				:	PartyOutstandingBreakup.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	N.Rajkumar
	'Created On					:	19th June 2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	07th April 2011
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
<!--#include virtual="/include/Accpopulate.asp"-->
<%
dim objRs,objRs2,objRs3,sQuery,iPageNo,saTemp,dTotalAmt,dTotalBal,dTotAmtPaid
dim sOrgId,sOrgName,sTillDate,sPartySubType,sSubTypeName,iPartyCode,sVouDate
dim sPartyName,sGetVal,sPartyType,iInvoiceNo,sInvoiceDate,dAmount,dAmtPaid,dBalAmt,iSNo,iNoOfDays
Dim sTransType,sTransName,iDueDays,iCrTransNo,sTempVal
iPageNo=1
iSNo=0
set objRs  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
set objRs3  = server.CreateObject("adodb.recordset")

'----------- To Get The Values From the Selection Page ----------------

sGetVal=Request.QueryString("Value")
'Response.Write sGetVal
'Response.End
saTemp=split(sGetVal,"|")
sOrgId= saTemp(0)
sOrgName=saTemp(1)
sTillDate=saTemp(2)
sPartySubType =saTemp(3)
iPartyCode=saTemp(4)
sPartyType=saTemp(6)

IF CStr(iPartyCode) <> "0" Then
	sPartyName = GetPartyName(iPartyCode)
Else
	sPartyName = "All Parties"
End IF
sTempVal = iPartyCode & ":" & sPartyType & ":" & sPartySubType & ":" & sPartyName 
'To Display Organizations Full Description

sQuery="Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID='" & sOrgId & "'"

with objRs
	.CursorLocation =3
	.CursorType=3
	.ActiveConnection =con
	.Source =sQuery
	.Open 
End with 
set objrs.ActiveConnection =nothing
If not objRs.EOF then
	sOrgName =objRs(0)
End if
objRs.Close 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receivables View</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<meta http-equiv="x-ua-compatible" content="IE=edge">
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" ID="ReceivableData"><Reminder/></script>
<script type="application/xml" data-itms-xml-island="1" ID="OutData"><Reminder/></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ReportReminderCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="GetXML('<%=sGetVal%>','<%=sPartyType%>')">

<form method="POST" name="formname" action="">	
	<Input type="Hidden" name="hCnt" value="">
	<Input type="Hidden" name="hGetVal" value="<%=sGetVal%>">
	<Input type="Hidden" name="hTempVal" value="<%=sTempVal%>">
	<Input type="Hidden" name="hPartycode" value="<%=iPartyCode%>">
	
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">

	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Outstanding  Receivables  as On <%=sTillDate%> From <%=sPartyName%>
		
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
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
								<td align="center">
								</td>
								<td valign="top" width="100%">

								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                <!--<div class="frmBody" id="frm2" style="width: 755; height:395;">-->
                                <div class="frmBody" id="frm2" style="width: 695; height:310;">
                                	<TABLE BORDER="0" CELLSPACING=1 CELLPADDING=0 WIDTH=100% class="ExcelTable" ID="RecTab">
										<tr>
											<TD class="ExcelHeaderCell" rowspan="2" align="center" width="15"><P>S.No.</TD>
											<td class="ExcelHeaderCell" align="center" Rowspan="2">
												<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15" onclick="">
												</a>
											</td>
											<TD class="ExcelHeaderCell" rowspan="2" align="center" width="15"><P>Doc Type.</TD>
											<TD COLSPAN=4 class="ExcelHeaderCell" align="center" ><P>Party Invoice </TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2"><P>Amount paid <br>till date</TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2"><P>Balance Receivable</TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2" width="80"><P>No of <br>days<br>Outstanding</TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2" width="80"><P>No of days<br>Over Due</TD>
										</tr>

										<tr>
											<TD class="ExcelHeaderCell" align="center" width="140"><P>No</TD>
											<TD class="ExcelHeaderCell" align="center" width="80"><P>Date</TD>
											<TD class="ExcelHeaderCell" align="center" width="80"><P>Accounted On</TD>
											<TD class="ExcelHeaderCell" align="center" width="100"><P>Amount</TD>
										</tr>

									</TABLE>
                                </div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
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
													<input type="button" value="Add To Reminder" OnClick="AddToReminder()" class="ActionButtonX" tabindex="3"  id=button1 name=button1>
													<input type="button" value="Close" OnClick="CloseWindow()" class="ActionButton" tabindex="3"  id=button2 name=button2>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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

<%
	Function GetDueDays(iCrTrNo,iOutStdDays,iAccTrNo,sBillDate)
		Dim sQuery,sObjDueRs,iOthAppNo,iPayTerms,iPayNoDays,iPayCount,iTotalDueDay
		Dim iDurPer,iDurDays,sPayTillDate,sObjPayRs,iPayRecNo,iParPayAmt,iPayInvAmt
		Dim iPayAmtToCome 
		Set sObjDueRs = Server.CreateObject("ADODB.RecordSet")
		Set sObjPayRs = Server.CreateObject("ADODB.RecordSet")
		
		sQuery = "Select isNull(OtherApplnTransNo,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iCrTrNo
		'Response.Write sQuery
		sObjDueRs.Open sQuery,Con
		IF Not sObjDueRs.EOF Then
			iOthAppNo = sObjDueRs(0)
		End IF
		sObjDueRs.Close
		
		IF CStr(iOthAppNo) = "0" Then
			GetDueDays = 0
		Else
			sQuery = "Select InvPymtTerms From Sal_T_InvoiceHeader Where SaleTransactionNo = " & iOthAppNo
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayTerms = sObjDueRs(0)
			End IF
			sObjDueRs.Close
			
			'iPayTerms = 2
			
			sQuery = "Select Count(1) From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
			
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayCount = sObjDueRs(0)
			End IF
			sObjDueRs.Close
			
			
			
			IF iPayCount = 1 Then
				sQuery = "Select DueDay From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayNoDays = sObjDueRs(0)
				End IF
				sObjDueRs.Close
				
				'iPayNoDays = 300
				
				'Response.Write iAccTrNo &"   " & iPayNoDays &"<br>"
				iTotalDueDay = CDbl(iOutStdDays) - CDbl(iPayNoDays)
			Else
				iPayAmtToCome = 0
				sQuery = "Select ReceivableNumber,AmountReceivable From Acc_T_Receivables Where TransactionNumber = "&iAccTrNo
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayRecNo = sObjDueRs(0)
					iPayInvAmt = sObjDueRs(1)
				End IF
				sObjDueRs.Close
				
				sQuery = "Select DueDay,DuePercent From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				Do While Not sObjDueRs.EOF
					iDueDays = sObjDueRs(0)
					iDurPer = sObjDueRs(1)
					
					'Response.Write iPayInvAmt &"  " & iDurPer &"<br>"
					
					iPayAmtToCome = Cdbl((CDbl(iPayInvAmt) * CDbl(iDurPer))/100) + CDbl(iPayAmtToCome)
					
					
					sQuery = "Select Top 1 Convert(Varchar,DateAdd(day,"&iDueDays&",Convert(Datetime,'"&sBillDate&"',103)),103) From Acc_T_RcvblAdjustmentDetails  "
					'Response.Write sQuery
					sObjPayRs.Open sQuery,Con
					
					IF Not sObjPayRs.EOF Then
						sPayTillDate = sObjPayRs(0)
					End IF
					sObjPayRs.Close
					
					sQuery = "Select Sum(AmountReceived) From Acc_T_RcvblAdjustmentDetails Where ReceivableNumber = "&iPayRecNo&" and "&_
							 "Convert(Datetime,ReceivedOn,103) <=  Convert(Datetime,'"&sPayTillDate&"',103)  "
							
					'Response.Write sQuery &"<br>"
					sObjPayRs.Open sQuery,Con 
					IF sObjPayRs.EOF Then
						iParPayAmt = sObjPayRs(0)
					End IF
					sObjPayRs.Close
					
					'iParPayAmt = 119000
					'Response.Write iParPayAmt &"  " & iPayAmtToCome &"<br>"
					
					IF CDbl(iParPayAmt) >= CDbl(iPayAmtToCome) Then
						iTotalDueDay = 0
					Else
						iTotalDueDay = CDbl(iOutStdDays) - CDbl(iDueDays)
					End IF
					
					'Response.Write iTotalDueDay &"<br>"
					
					
							 
					sObjDueRs.MoveNext
				Loop
				sObjDueRs.Close
			End IF
		End IF
		
		IF iTotalDueDay <0 Then
			iTotalDueDay = 0
		End IF
		
		IF CStr(iTotalDueDay) = "" Then
			iTotalDueDay = 0
		End IF
		
		GetDueDays = iTotalDueDay
	End Function
%>
