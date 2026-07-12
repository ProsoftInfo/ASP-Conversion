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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
dim objRs,objRs2,objRs3,sQuery,iPageNo,saTemp,dTotalAmt,dTotalBal,dTotAmtPaid
dim sOrgId,sOrgName,sTillDate,sPartySubType,sSubTypeName,iPartyCode,sVouDate
dim sPartyName,sGetVal,sPartyType,iInvoiceNo,sInvoiceDate,dAmount,dAmtPaid,dBalAmt,iSNo,iNoOfDays
Dim sTransType,sTransName,iDueDays,iCrTransNo
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
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" ID="ReceivableData"><Reminder/></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/ReportReminderCompat.js"></SCRIPT>
<script>
function CheckSubmit() {
	return true;
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="GetXML('<%=sGetVal%>')">

<form method="POST" name="formname" action="">
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
                                <div class="frmBody" id="frm2" style="width: 645; height:310;">
<TABLE BORDER="0" CELLSPACING=1 CELLPADDING=0 WIDTH=100% class="ExcelTable" >
<tr>
<TD class="ExcelHeaderCell" rowspan="2" align="center" width="15">
<P>S.No.</TD>
<td class="ExcelHeaderCell" align="center" Rowspan="2">
	<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15" onclick="">
	</a>
</td>
<TD class="ExcelHeaderCell" rowspan="2" align="center" width="15">
<P>Doc Type.</TD>
<TD COLSPAN=4 class="ExcelHeaderCell" align="center" >
<P>Party Invoice </TD>
<TD class="ExcelHeaderCell" align="center" rowspan="2">
<P>Amount paid till date</TD>
<TD class="ExcelHeaderCell" align="center" rowspan="2">
<P>Balance Receivable</TD>
<TD class="ExcelHeaderCell" align="center" rowspan="2" width="80">
<P>No of days<br>Outstanding</TD>
<TD class="ExcelHeaderCell" align="center" rowspan="2" width="80">
<P>No of days<br>Over Due</TD>
</tr>
<tr>
<TD class="ExcelHeaderCell" align="center" width="140">
<P>No</TD>
<TD class="ExcelHeaderCell" align="center" width="80">
<P>Date</TD>
<TD class="ExcelHeaderCell" align="center" width="80">
<P>Accounted On</TD>
<TD class="ExcelHeaderCell" align="center" width="100">
<P>Amount</TD>
</tr>

<%
Dim iPartyCheck,dTotRec,dParOpenAmt,dParCloseAmt,sCheck,sOpenCD,sCloseCd
IF CStr(iPartyCode) = "0" Then
	'sQuery = "Select PartyCode From APP_R_OrgParty Where PartySubType = "&sPartySubType&" "&_
	'		 "and PartyType = '"&sPartyType&"' and OUDefinitionID = '"&sOrgId &"' "
	
	sQuery = "Select O.PartyCode,P.PartyName From APP_R_OrgParty O,App_M_PartyMaster P Where PartySubType = "&sPartySubType&" "&_
			 "and PartyType = '"&sPartyType&"' and OUDefinitionID = '"&sOrgId &"' and O.PartyCode = P.PartyCode Order By P.PartyName "
			 
	With Objrs2
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	iSNo = 1
	Set objRs2.ActiveConnection = Nothing
	Do While Not objRs2.EOF 
		iPartyCode = objRs2(0)
		sPartyName = objRs2(1)
		iPartyCheck = "1"
		
		sCheck = "F"
		
		'Taking the Party Opening AMount	
		sQuery = "Select ClosingAmount,OpeningAmount,OpeningCdIndication,ClosingCDIndication From Acc_T_PartyOpeningAmt Where PartyType = '"&sPartyType&"' and  "&_
			     "PartySubType = "&sPartySubType&" and PartyCode = "&iPartyCode&" and OUDefinitionID = '"&sOrgId&"' "
			     
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			dParOpenAmt = objRs(1)
			dParCloseAmt = objRs(0)
			sOpenCD = objRs(2)
			sCloseCd = objRs(3)	
		Else
			dParOpenAmt = 0
			dParCloseAmt = 0
		End IF
		objRs.Close
		
		'Response.Write sOpenCD &" " & sCloseCd
		
		'IF CStr(sOpenCD) = "C" Then
		'	dParOpenAmt = 0
		'End IF
		
		'IF CStr(sCloseCd) = "C" Then
		'	dParCloseAmt = 0
		'End IF
		
		
		
		
		
		'========= Talking for any values if the party had any advances. ========================
		'sQuery = "Select P.AdvanceReceived,H.TransactionType, Convert(Char,H.VoucherDate,103),H.CrDrIndication  "&_
		'		 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
		'		 "P.PartySubType = "&sPartySubType&" and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
		'		 "H.TransactionNumber = P.TransactionNumber and P.AdvanceReceived is Not Null "
				 
		'=======================================================================================
		'Changed on		:	26/11/2004
		'Reason			:	The display all receivables from the party the 
		'Changed By		:	Manohar Parbhu.R
		'=======================================================================================
		
		sQuery = "Select P.AdvancePaid - isNull(P.AdvanceAdjusted,0),H.TransactionType, Convert(Char,H.VoucherDate,103),H.CrDrIndication, isNull(P.AdvanceAdjusted,0), "&_
				 "P.AdvancePaid, VoucherNumber,Convert(Char,VoucherDate,103) "&_
				 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
				 "P.PartySubType = "&sPartySubType&" and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
				 "H.TransactionNumber = P.TransactionNumber and P.AdvancePaid is Not Null and P.AdvancePaid - isNull(P.AdvanceAdjusted,0) > 0"
				 
		
				 
		 'Response.Write sQuery
				 
		with objRs
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sQuery
			.Open
		End with
		set objRS.ActiveConnection =nothing
		Do While Not objRs.EOF
			dTotRec=cdbl(dTotRec)+cdbl(objRs(0))
			dTotalBal=cdbl(dTotalBal)+cdbl(objRs(0))
			dTotalAmt = CDbl(dTotalAmt) + CDbl(objRs(5))
			sCheck = "T"
			IF CStr(iPartyCheck) = "1" Then
			
	%>
				<!--<tr>
					<td colspan=10 class="ExcelDisplayCell" align="left" ><b><%Response.Write(sPartyName)%></b></td>
				</tr>-->
				
				<!--tr>
					<td colspan=7 class="ExcelDisplayCell" align="Right" ><b>Outstanding Opening Amount</b></td>
					<td class="ExcelDisplayCell" align="Right" ><b><%=FormatNumber(dParOpenAmt,2,,,0)%></b></td>
					<td class="ExcelDisplayCell" align="Right" ></td>
					
				</tr-->
	<%
				iPartyCheck = "2"
			End IF	
	%>
				
			
				<tr>
				<TD class="ExcelDisplayCell" align="center">
				<P><%=iSNo%></TD>
				<td class="ExcelDisplaycell" align="center">
					<Input type="Checkbox" Name="chkBox"  value="">
				</td>
				<TD class="ExcelDisplayCell" align="center">
				<P><%=objRs(1)%></TD>
				<TD class="ExcelDisplayCell" align="center">
				<P><%=objRs(6)%></TD>
				<TD class="ExcelDisplayCell" align="center">
				<P><%=objRs(7)%></TD>
				<TD class="ExcelDisplayCell" align="center">
				<P><%=objRs(2)%></TD>
				<TD class="ExcelDisplayCell" align="right">
				<P><%=FormatNumber(objRs(5),2,,,0)%></TD>
				<TD class="ExcelDisplayCell" align="right">
				<P><%=FormatNumber(objRs(4),2,,,0)%></TD>
				
				<TD class="ExcelDisplayCell" align="right">
				<P><%=FormatNumber(objRs(0),2,,,0)%></TD>
				<TD class="ExcelDisplayCell" align="center">
				<P></TD>
				<TD class="ExcelDisplayCell" align="center">
				<P></TD>
				</tr>
	<%
				
				iSNo = iSNo + 1
				objRs.MoveNext
				loop
	
			objRs.Close	 
	
		sQuery="Select PartyInvoiceNumber,convert(char,PartyInvoiceDate,103),AmountReceivable,"&_
				" AmountReceived,datediff(day,convert(datetime,partyInvoiceDate,103),convert(datetime,'"& sTillDate &"',103)),"&_
				" convert(char,VoucherDate,103),isNull(TransactionNumber,0) from Acc_T_Receivables"&_
				" where OuDefinitionId='"&sOrgId &"' and PartyCode="&iPartyCode &_
				" and PartySubType="&sPartySubType&"  and partyType='"&sPartyType &_
				"' and AmountReceivable>AmountReceived"&_
				" and convert(datetime,VoucherDate,103)<= convert(datetime,'"& sTillDate &"',103)"

		'Response.Write "2="& sQuery
		with objRs
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sQuery
			.Open
		End with
		set objRS.ActiveConnection =nothing
		
		

		while not objRs.EOF
			iSNo =iSNo +1 
			sCheck = "T"
			iInvoiceNo=objRs(0)
			sInvoiceDate=objRs(1)
			dAmount=objRs(2)
			dAmtPaid=objRs(3)
			iNoOfDays=objRs(4)
			sVouDate=objRs(5)
			dBalAmt=cdbl(dAmount)-cdbl(dAmtPaid)
			dTotalAmt =cdbl(dTotalAmt)+cdbl(dAmount)
			dTotAmtPaid=cdbl(dTotAmtPaid)+cdbl(dAmtPaid)
			dTotalBal=cdbl(dTotalBal)+cdbl(dBalAmt)
			dTotRec = CDbl(dTotRec) + CDbl(dBalAmt)
			
			sQuery = "Select TransactionType,VoucherNumber,Convert(Char,VoucherDate,103),CreatedTransNo From Acc_T_VoucherHeader Where TransactionNumber = "&objRs(6)&" " 
			
			
			with objRs3
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection=con
				.Source =sQuery
				.Open
			End with
			set objRs3.ActiveConnection =nothing
			IF Not objRs3.EOF Then
				sTransType = objRs3(0)
				iInvoiceNo = objRs3(1)
				sInvoiceDate = objRs3(2)
				iCrTransNo = objRs3(3)
			End IF
			objRs3.Close
			
			IF CStr(sTransType) = "SJR" Then
				iDueDays = GetDueDays(iCrTransNo,iNoOfDays,objRs(6),sInvoiceDate)			
			Else
				iDueDays = 0
			End IF
			
			IF CStr(iPartyCheck) = "1" Then 
				iPartyCheck = 2
				IF CStr(iSNo) <> "1" Then
	%>				<tr>
						<td colspan=10 class="ExcelDisplayCell" align="left" ><b>&nbsp;</b></td>
					</tr>
					
					<!--tr>
						<td colspan=9 class="ExcelDisplayCell" align="Right" ><b>CLOSING AMOUNT</b></td>
					</tr-->
				<%End IF %>
				<!--<tr>
					<td colspan=10 class="ExcelDisplayCell" align="left" ><b><%Response.Write(sPartyName)%></b></td>
				</tr>-->
				
				<!--tr>
					<td colspan=7 class="ExcelDisplayCell" align="Right" ><b>Outstanding Opening Amount</b></td>
					<td class="ExcelDisplayCell" align="Right" ><b><%=FormatNumber(dParOpenAmt,2,,,0)%></b></td>
					<td class="ExcelDisplayCell" align="Right" ></td>
					
				</tr-->
				
				
			<%End IF%>
<tr>
<TD class="ExcelDisplayCell" align="center">
<P><%=iSNo%></TD>
<td class="ExcelDisplaycell" align="center">
	<Input type="Checkbox" Name="chkBox" value="">
</td>
<TD class="ExcelDisplayCell" align="center">
<P><%=sTransType%></TD>
<TD class="ExcelDisplayCell" align="center">
<P><%=iInvoiceNo%></TD>
<TD class="ExcelDisplayCell" align="center">
<P><%=sInvoiceDate%></TD>
<TD class="ExcelDisplayCell" align="center">
<P><%=sVouDate%></TD>
<TD class="ExcelDisplayCell" align="right">
<P><%=FormatNumber(dAmount,2,,,0)%></TD>
<TD class="ExcelDisplayCell" align="right">
<P><%=FormatNumber(dAmtPaid,2,,,0)%></TD>
<TD class="ExcelDisplayCell" align="right">
<P><%=FormatNumber(dBalAmt,2,,,0)%></TD>
<TD class="ExcelDisplayCell" align="center">
<P><%=iNoOfDays%></TD>
<TD class="ExcelDisplayCell" align="center">
<P><a href="#" onclick="ShowDetails('<%=iCrTransNo%>','<%=iDueDays%>'); return false;" class="ExcelDisplayLink" alt="Over Due Details">
<%=iDueDays%></a></TD>

</tr>



<%			objRs.MoveNext 
		Wend
		objRs.Close
		IF Cstr(sCheck) = "T" Then
%>
			<!--tr>
					<td colspan=7 class="ExcelDisplayCell" align="Right" ><b>Outstanding Closing Amount</b></td>
					<td class="ExcelDisplayCell" align="Right" ><b><%=FormatNumber(dParCloseAmt,2,,,0)%></b></td>
					<td class="ExcelDisplayCell" align="Right" ></td>
					
			</tr-->
<%
		End IF
		objRs2.MoveNext
%>
		
<%
	loop
	objRs2.Close
	
	dTotalBal = abs(CDbl(dTotalBal) - CDbl(dTotRec))
	
%>
	
<%
	
Else
	Dim sCrDrInvNo,sCrDrInvDate
	iSNo = 1
	Dim sQuery1
	iPartyCheck = "1"
	
	sQuery = "Select ClosingAmount,OpeningAmount From Acc_T_PartyOpeningAmt Where PartyType = '"&sPartyType&"' and  "&_
			     "PartySubType = "&sPartySubType&" and PartyCode = "&iPartyCode&" and OUDefinitionID = '"&sOrgId&"' "
			     
		'Response.Write "3="& sQuery
		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dParCloseAmt = objRs2(0)
			dParOpenAmt = Objrs2(1)
		Else
			dParCloseAmt = 0
			dParOpenAmt = 0
		End IF
		Objrs2.Close
		
	%>
				<!--<tr>
					<td colspan=10 class="ExcelDisplayCell" align="left" ><b><%Response.Write(sPartyName)%></b></td>
				</tr>-->
				<!--tr>
					<td colspan=7 class="ExcelDisplayCell" align="Right" ><b>Outstanding Opening Amount</b></td>
					<td class="ExcelDisplayCell" align="Right" ><b><%=FormatNumber(dParOpenAmt,2,,,0)%></b></td>
					<td class="ExcelDisplayCell" align="Right" ></td>
					
				</tr-->
	<%
		
		
	iPartyCheck = "2"
	IF CStr(Trim(sPartySubType)) = "" Then
	
		'Taking the Party Opening AMount	
		
		
		
	
		sQuery1 = "Select O.PartyCode,P.PartyName,O.PARTYSUBTYPE From APP_R_OrgParty O,App_M_PartyMaster P Where "&_
			      " PartyType = '"&sPartyType&"' and OUDefinitionID = '"&sOrgId &"' and O.PartyCode = "& iPartyCode &" and O.PartyCode = P.PartyCode "
		
		With Objrs2
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = Con
			.Source = sQuery1
			.Open
		End With

		Set objRs2.ActiveConnection = Nothing
		Do While Not objRs2.EOF
			sPartySubType = sPartySubType &","&Trim(Objrs2(2))
		objRs2.MoveNext
		loop
		objRs2.Close
		sPartySubType = Mid(Trim(sPartySubType),2)
	End IF
			
	'========= Talking for any values if the party had any advances. ========================
		'sQuery = "Select P.AdvanceReceived,H.TransactionType, Convert(Char,H.VoucherDate,103),H.CrDrIndication  "&_
		'		 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
		'		 "P.PartySubType = "&sPartySubType&" and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
		'		 "H.TransactionNumber = P.TransactionNumber and P.AdvanceReceived is Not Null "
		
		'=======================================================================================
		'Changed on		:	26/11/2004
		'Reason			:	The display all receivables from the party the 
		'Changed By		:	Manohar Parbhu.R
		'=======================================================================================
		
		sQuery = "Select P.AdvancePaid - isNull(P.AdvanceAdjusted,0),H.TransactionType, Convert(Char,H.VoucherDate,103),H.CrDrIndication, isNull(P.AdvanceAdjusted,0), "&_
				 "P.AdvancePaid, VoucherNumber,Convert(Char,VoucherDate,103) "&_
				 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
				 "P.PartySubType in("&sPartySubType&") and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
				 "H.TransactionNumber = P.TransactionNumber and P.AdvancePaid is Not Null and P.AdvancePaid - isNull(P.AdvanceAdjusted,0) > 0"
			
		'Response.Write "5="& sQuery
		
				 
		with objRs
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sQuery
			.Open
		End with
		set objRS.ActiveConnection =nothing
		Do While Not objRs.EOF			
			dTotalBal=cdbl(dTotalBal)+cdbl(objRs(0))
			dTotRec=cdbl(dTotRec)+cdbl(objRs(0))
			dTotalAmt = CDbl(dTotalAmt) + CDbl(objRs(5))
			IF CStr(iPartyCheck) = "1" Then			
	%>
				<tr>
					<td colspan=10 class="ExcelDisplayCell" align="left" ><b><%Response.Write(sPartyName)%></b></td>
				</tr>
				<!--tr>
					<td colspan=8 class="ExcelDisplayCell" align="Right" ><b>Outstanding Opening Amount</b></td>
					<td class="ExcelDisplayCell" align="Right" ><b><%=FormatNumber(dParOpenAmt,2,,,0)%></b></td>
					<td class="ExcelDisplayCell" align="Right" ></td>
					
				</tr-->
	<%
				iPartyCheck = "2"
			End IF
	%>				
				<tr>
					<TD class="ExcelDisplayCell" align="center"><P><%=iSNo%></TD>				
					<td class="ExcelDisplaycell" align="center">
						<Input type="Checkbox" Name="chkBox" value="">
					</td>
					<TD class="ExcelDisplayCell" align="center"><P><%=objRs(1)%></TD>				
					<TD class="ExcelDisplayCell" align="center"><P><%=objRs(6)%></TD>				
					<TD class="ExcelDisplayCell" align="center"><P><%=objRs(7)%></TD>				
					<TD class="ExcelDisplayCell" align="center"><P><%=objRs(2)%></TD>				
					<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(objRs(5),2,,,0)%></TD>				
					<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(objRs(4),2,,,0)%></TD>					
					<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(objRs(0),2,,,0)%></TD>				
					<TD class="ExcelDisplayCell" align="center"><P></TD>				
					<TD class="ExcelDisplayCell" align="center"><P></TD>				
				</tr>
	<%
				iSNo = iSNo + 1
				objRs.MoveNext
				loop
			objRs.Close	 
			iSNo = iSNo - 1
			
			'objRs2.MoveNext
		'Loop		
		'End IF
		'objRs2.close
	dim iRcvbleNo
	sQuery = "Select PartyInvoiceNumber,convert(char,PartyInvoiceDate,103),AmountReceivable,"&_
			 " AmountReceived,datediff(day,convert(datetime,partyInvoiceDate,103),convert(datetime,'"& sTillDate &"',103)),"&_
			 " convert(char,VoucherDate,103),isNull(TransactionNumber,0),ReceivableNumber from Acc_T_Receivables"&_
			 " where OuDefinitionId='"&sOrgId &"' and PartyCode="&iPartyCode &_
			 " and PartySubType IN ("&sPartySubType&")  and partyType='"&sPartyType &_
			 "' and AmountReceivable>AmountReceived"&_
			 " and convert(datetime,VoucherDate,103)<= convert(datetime,'"& sTillDate &"',103)"

	'Response.Write "4="& sQuery
	
	with objRs
		.CursorLocation =3
		.CursorType =3
		.ActiveConnection=con
		.Source =sQuery
		.Open
	End with
	set objRS.ActiveConnection =nothing

	while not objRs.EOF
		iSNo =iSNo +1 
		iInvoiceNo=objRs(0)
		sInvoiceDate=objRs(1)
		dAmount=objRs(2)
		dAmtPaid=objRs(3)
		iNoOfDays=objRs(4)
		sVouDate=objRs(5)
		iRcvbleNo=objRs(7)
		dBalAmt=cdbl(dAmount)-cdbl(dAmtPaid)
		
		dTotalAmt =cdbl(dTotalAmt)+cdbl(dAmount)
		dTotAmtPaid=cdbl(dTotAmtPaid)+cdbl(dAmtPaid)
		dTotalBal=cdbl(dTotalBal)+cdbl(dBalAmt)
		dTotRec = CDbl(dTotRec) + CDbl(dBalAmt)
		
		sQuery = "Select TransactionType,VoucherNumber,Convert(Char,VoucherDate,103),CreatedTransNo From Acc_T_VoucherHeader Where TransactionNumber = "&objRs(6)&" " 
			
			'Response.Write "<BR><BR>A="&sQuery
			with objRs3
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection=con
				.Source =sQuery
				.Open
			End with
			set objRs3.ActiveConnection =nothing
			IF Not objRs3.EOF Then
				sTransType = objRs3(0)
				iInvoiceNo = objRs3(1)
				sInvoiceDate = objRs3(2)
				iCrTransNo = objRs3(3)
			End IF
			objRs3.Close
			
			IF CStr(sTransType) = "SJR" Then
				iDueDays = GetDueDays(iCrTransNo,iNoOfDays,objRs(6),sInvoiceDate)			
			Else
				iDueDays = 0
			End IF
			
%>
<tr>
	<TD class="ExcelDisplayCell" align="center"><P><%=iSNo%></TD>
	<td class="ExcelDisplaycell" align="center">
		<Input type="Checkbox" Name="chkBox" value="">
	</td>
	<TD class="ExcelDisplayCell" align="center"><P><%=sTransType%></TD>
	<TD class="ExcelDisplayCell" align="center"><P><%=iInvoiceNo%></TD>
	<TD class="ExcelDisplayCell" align="center"><P><%=sInvoiceDate%></TD>
	<TD class="ExcelDisplayCell" align="center"><P><%=sVouDate%></TD>
	<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(dAmount,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(dAmtPaid,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><P><%=FormatNumber(dBalAmt,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="center"><P><%=iNoOfDays%></TD>
	<TD class="ExcelDisplayCell" align="center"><P><a href="#" onclick="ShowDetails('<%=iCrTransNo%>','<%=iDueDays%>'); return false;" class="ExcelDisplayLink" alt="Over Due Details">
	<%=iDueDays%></a>&nbsp;
	<%'if trim(dBalAmt) > "0" then%>
		<!--img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Invoice Closing" width="10" height="11" onclick="InvoicePopUp('<%=iRcvbleNo%>','<%=sInvoiceDate%>','<%=dAmtPaid%>','<%=dBalAmt%>','<%=sTransType%>' )"-->
	<%'end if%>
</TD>
</tr>
<%		objRs.MoveNext 
	Wend
	objRs.Close
	
	dTotalBal = abs(CDbl(dTotalBal) - CDbl(dTotRec))
%>
	<!--tr>
		<td colspan=7 class="ExcelDisplayCell" align="Right" ><b>Outstanding Closing Amount</b></td>
		<td class="ExcelDisplayCell" align="Right" ><%Response.Write(FormatNumber(dParCloseAmt,2,,,0))%></td>
		<td class="ExcelDisplayCell" align="Right" ></td>
		
	</tr-->
<%
	
End IF
%>

<TD colspan=6 class="ExcelDisplayCell" align="right"><b>Total Outstanding Transaction Amount
<TD class="ExcelDisplayCell" align="right"><P><b><%=FormatNumber(dTotalAmt,2,,,0)%></TD>
<TD class="ExcelDisplayCell" align="right"><P><b><%=FormatNumber(dTotAmtPaid,2,,,0)%></TD>
<TD class="ExcelDisplayCell" align="right"><P><b><%=FormatNumber(dTotRec,2,,,0)%></TD>
<TD class="ExcelDisplayCell" align="right">
<TD class="ExcelDisplayCell" align="right">
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
													<input type="button" value="Add To Reminder" OnClick="CheckSubmit()" class="ActionButtonX" tabindex="3"  id=button1 name=button1>
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
