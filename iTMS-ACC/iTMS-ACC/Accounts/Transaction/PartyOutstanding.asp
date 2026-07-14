<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PartyOutstanding.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 07, 2010
	'Modified By                :   
	'Modified On                :   
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
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<%
	Dim sUnitID,iCnt,sSql,sTillDate
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo
	Dim iTotalPages,iTotalRecords,iPrevPage,iNextPage
	Dim sSentBy,sSentToVoucher,iSno
	Dim sQuery,sPartyCode,sPartySubType
	Dim sPartyName,sGetVal,sPartyType,dAmtLtThirty,dAmtGtThirty,dAmtGtSixty
	Dim iPartyCode,sOrgnPartyCode,dAmtGtNinety,dTotalBalance,dTotPgBal
	Dim dAmtGtNinetyTot,dAmtGtSixtyTot,dAmtGtThirtyTot,dAmtLtThirtyTot

	Dim Objrs,Objrs1
	
	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")
	
	sUnitID = Session("organizationcode")
	Response.Write "<font color=red>"
	
	Const iPageSize=16
	iPageNo = trim(Request("hPage"))
	if trim(iPageNo) = "" then iPageNo = 1	
	
'	response.write "<font color=red>"& Request.QueryString
	
	iCurrentPage=CInt(Request.Form("hPageSelection"))	
	
	sPartyType = Request("PartyType")
	sPartyCode = Request("hPartyCode")
	sPartySubType = Request("hPartySubType")
	'Response.Write "sPartyType = "& sPartyType 
	
	sTillDate = Trim(Request("hTillDate"))
	If sTillDate = "" Then
		sTillDate = FormatDate(Date())
	End IF
	If sPartyCode = "" Then sPartyCode = "0"
'	Response.Write "<p><font color=red>Data="&sPartyCode & "====="& sPartySubType
		
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="x-ua-compatible" content="IE=edge">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" ID="OutData"><PartyType/></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Party/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutStandingData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="GenReminder"><Root/></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/VoucherEntryCore.js"></script>
<script src="../../scripts/BankVoucher.js"></script>
<script src="../../scripts/ReportReminderCompat.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
	<form method="POST" name="formname" action="" >
	<input type="hidden" name="PartyType" value="<%=sPartyType%>">
	<input type=hidden name="hUnitNo" value="<%=sUnitID%>">
	<input type=hidden name="hUnitName" value="<%=session("orgShortName")%>">
	<input type=hidden name="hTillDate" value="<%=sTillDate%>">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<Input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
	<Input type="hidden" name="hPartySubType" value="<%=sPartySubType%>">
	<Input type="Hidden" name="hCount" value="">
	<input type="Hidden" name="hSelNode" value="">
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle">
				<p align="center">Party Outstanding
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href="#" onclick="return Div_OnClick(idUnprocessed,'',event)"  itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%">
<div id="idUnprocessed" style="display: none">
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<tr>
	<td class="FieldCellsub">Party</td>
	<td class="FieldcellSub"> 
		<span id="PartyName" class="Dataonly"></span>
		<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" onclick="SelParty(); return false;"></a>
	</td>
</tr>
<tr>
	<td class="FieldCellsub">Till Date</td>
	<td class="FieldCellsub">
		<%Response.Write InsertDatePicker("ctlTillDate")%>
	</td>
</tr>

<tr>
<td class="FieldCell"></td>
<td class="FieldCell" >
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
</td>
<td class="FieldCell" >
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<!--<div class="frmBody" id="frm4" style="width: 585; height:140;">-->
<!--<div class="frmBody" id="frm4" style="height:270;">-->
	
	<!--<div class="frmBody" id="frm4" style="height:130;">-->
		<table cellspacing="1" class="ExcelTable" width="100%" >
		<tr>
			<TD align="center" class="ExcelHeaderCell" rowspan="2"><P>S.No.</TD>
			<TD class="ExcelHeaderCell" align="center" rowspan="2"><P>Party code</TD>
			<TD COLSPAN=5 class="ExcelHeaderCell" align="center"><P>Outstanding amount for no of days</TD>
		</tr>

		<tr>
			<TD class="ExcelHeaderCell" align="center"><P>&lt; 30</TD>
			<TD class="ExcelHeaderCell" align="center"><P>30 - 60</TD>
			<TD class="ExcelHeaderCell" align="center"><P>61 - 90</TD>
			<TD class="ExcelHeaderCell" align="center"><P>&gt; 90</TD>
			<TD class="ExcelHeaderCell" align="center"><P>Total</TD>
		</tr>

		<%	

		sQuery="select distinct(partycode),orgnpartycode,partyName,PartySubType from vwOrgparty where "&_
				" OUDefinitionID='" & sUnitID & "' and partytype='" & sPartyType & "'"
				
		If cint(trim(sPartyCode)) <> "0" then 
			sQuery= sQuery &" and PartyCode = "&sPartyCode&" " 		
		End IF
		If trim(sPartySubType) <> ""  then 
			sQuery= sQuery &" and PartySubType = "&sPartySubType&" " 		
		End IF
		' Response.write sQuery
		 'Response.End
			 
		with objRs
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sQuery 
			.Open 
		End with

		If not objRs.EOF Then
			set iPartyCode =objRs(0)
			set sOrgnPartyCode=objRs(1)	
			Set sPartyName =objRS(2)
			set sPartySubType =objRs(3)
		End if

		dAmtLtThirtyTot=0
		dAmtGtThirtyTot=0
		dAmtGtSixtyTot=0
		dTotPgBal=0

		While not objRs.EOF 
		dTotalBalance=0
		if Trim(sPartyType)="DR" then

			sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=0"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<30 and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
		'	Response.Write "<font color=red>< 30 = "&sQuery &"<BR>" 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtLtThirty=objRs1(0)
			dAmtLtThirtyTot=cdbl(dAmtLtThirtyTot)+cdbl(dAmtLtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtLtThirty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=30"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<60 and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "> 30 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtThirty=objRs1(0)
			dAmtGtThirtyTot=cdbl(dAmtGtThirtyTot)+cdbl(dAmtGtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtThirty)
			objRs1.Close 

		sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=60"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<90 and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<font color=red>30  60 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtSixty=objRs1(0)
			dAmtGtSixtyTot=cdbl(dAmtGtSixtyTot)+cdbl(dAmtGtSixty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtSixty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=90"&_
				" and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<BR><BR>iPartyCode  > 90 = "&sQuery 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtNinety=objRs1(0)
			dAmtGtNinetyTot=cdbl(dAmtGtNinetyTot)+cdbl(dAmtGtNinety)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtNinety)
			objRs1.Close 
			dTotPgBal=cdbl(dtotPgBal)+cdbl(dTotalBalance)
	end if 'if Trim(sPartyType)="DR" then

			'sQuery = "Select isNull(Sum(P.AdvanceReceived - isNull(P.AdvanceAdjusted,0)),0) "&_
			'			 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
			'			 "P.PartySubType = "&sPartySubType&" and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
			'			 "H.TransactionNumber = P.TransactionNumber and P.AdvanceReceived is Not Null "
			
			'with objRs1
			'	.CursorLocation =3
			'	.CursorType =3
			'	.ActiveConnection =con
			'	.Source =sQuery
			'	.Open 
			'End with
			'set objRs1.ActiveConnection =nothing
			
			'IF Not objRs1.EOF Then
			'	dTotalBalance=cdbl(dTotalBalance)+cdbl(objRs1(0))
			'End IF
			'objRs1.Close
			
    if Trim(sPartyType)="CR" then
		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=0"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<30 and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
		'	Response.Write "<font color=red>< 30 = "&sQuery &"<BR>" 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtLtThirty=objRs1(0)
			dAmtLtThirtyTot=cdbl(dAmtLtThirtyTot)+cdbl(dAmtLtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtLtThirty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=30"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<60 and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "> 30 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtThirty=objRs1(0)
			dAmtGtThirtyTot=cdbl(dAmtGtThirtyTot)+cdbl(dAmtGtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtThirty)
			objRs1.Close 

		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=60"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<90 and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<font color=red>30  60 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtSixty=objRs1(0)
			dAmtGtSixtyTot=cdbl(dAmtGtSixtyTot)+cdbl(dAmtGtSixty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtSixty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=90"&_
				" and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<BR><BR>iPartyCode  > 90 = "&sQuery 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtNinety=objRs1(0)
			dAmtGtNinetyTot=cdbl(dAmtGtNinetyTot)+cdbl(dAmtGtNinety)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtNinety)
			objRs1.Close 
			dTotPgBal=cdbl(dtotPgBal)+cdbl(dTotalBalance)

	end if 'if Trim(sPartyType)="CR" then		
			
			
						 
		If dTotalBalance<>0 Then
		iSNo=iSNo+1
		%>
		<tr>
			<TD align="center" class="ExcelSerial"><P><%=iSNo%></TD>
			<TD class="ExcelDisplayCell" align="left">
				<A Name="<%=sPartySubType%>|<%=iPartyCode%>|<%=iPartyCode%>|<%=sPartyType%>"
				HRef="#" class="ExcelDisplayLink" onclick="ShowCCDetails(this); return false;" ALT="View Paybales Details"><%=sPartyName%></a>
			</TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtLtThirty,2,,,0)%> </TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtGtThirty,2,,,0)%></TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtGtSixty,2,,,0)%></TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtGtNinety,2,,,0)%></TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dTotalBalance,2,,,0)%></TD>
		</tr>
		<%
		End if
		objRs.MoveNext 
		wend 
		%>
		<TD Colspan=2 align="right" class="ExcelDisplayCell">
		<P><b>Total</TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtLtThirtyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtGtThirtyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtGtSixtyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtGtNinetyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dTotPgBal,2,,,0)%></TD>

		</table>
	</Div>
	
	<table>
		<tr>
			<td align="center" class="MiddlePack" colspan="3"></td>
		</tr>
	</Table>

	<!--<div class="frmBody" id="frm4" style=" height:130;">-->
		<table cellspacing="1" class="ExcelTable" width="100%" ID="RecTab">
			<tr>
				<TD class="ExcelHeaderCell" rowspan="2" align="center"><P>S.No.</TD>
				<td class="ExcelHeaderCell" align="center" Rowspan="2">
					<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15" onclick="">
					</a>
				</td>
				<TD class="ExcelHeaderCell" rowspan="2" align="center"><P>Party</TD>
				<TD class="ExcelHeaderCell" rowspan="2" align="center"><P>Invoice <br>Amount</TD>
				<TD class="ExcelHeaderCell" colspan="4" align="center"><P>Deductions</TD>
				<TD class="ExcelHeaderCell" colspan="2" align="center"><P>Amount</TD>
			</tr>
			<tr>
				<TD class="ExcelHeaderCell" align="center"><P>TDS</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Advance</TD>
				<TD class="ExcelHeaderCell" align="center"><P>C.Note</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Freight</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Paid</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Outstanding</TD>
			</tr>
		</Table>
		
	<!--</div>-->
	
	<!--<div class="frmBody" id="frm4" style=" width:570 ;height:90;">-->
		<Table>
			<tr>
				<td class="FieldCellSub">&nbsp;</td>
				</td>
			</tr>
		</Table>
	
		<Table cellspacing="0" class="BodyTable" width="100%">	
			<tr>
				<td class="FieldCellSub">To Send By</td>
				<td class="FieldCell">
					<Input type="Radio" name="radSendBy" Value="C">Courier
					<Input type="Radio" name="radSendBy" Value="E">E-Mail
				</td>
			</tr>
			<tr>
				<td class="FieldCell">Courier Company Name</td>
				<td class="FieldCell">
					<Input type="Text" name="txtCouComName" value="" class="FormElem">
				</td>
				<td class="FieldCell">Courier TransactionID</td>
				<td class="FieldCell">
					<Input type="Text" name="txtCouTransID" value="" class="FormElem">
				</td>
			</tr>
			<tr>
				<td class="FieldCell">Address</td>
				<td class="FieldCell" colspan="2">
					<Textarea type="Text" name="txtCouComAddress" value="" class="FormElem" cols="40"></Textarea>
				</td>
			</tr>
		</Table>
		
	<!--</div>

</div>-->
</td>
<td align="center" class="ClearPixel" width="5">
</td>
</tr>
<input type=hidden name="hCnt" value=<%=iSno-1%>>
<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="PaginateAcc('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage - 1%>')" id=button4 name=button4>
<%		end if	%>
<SELECT class="FormElem" onChange="PaginateAcc(this(this.selectedIndex).value)" id=select1 name=select1>
<%
For lnPage = 1 To iTotalPage
If lnPage = iCurrentPage Then
%>
<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
<%		else	%>
<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
<%		end if
next
%>
</SELECT>
<%
if iCurrentPage = iTotalPage then
%>
<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

<%		else	%>
<input type="button" value=" >> " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage + 1%>')" id=button7 name=button7>
<input type="button" value=" >| " class="ActionButtonX" onclick="PaginateAcc('<%=iTotalPage%>')" id=button8 name=button8>
<%		end if
End If
%>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td valign="middle" class="ActionCell">
<p align="center">
	<tr>
		
		<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="middle" class="ActionCell">
                        <p align="center">
                         <!--<input type="button" value="Preview Reminder" class="ActionButtonX"  id="button1" name=button1 OnClick="CheckSumbit('P')" >-->
                         <input type="button" value="Generate Reminder" class="ActionButtonX"  id="button2" name=button2 OnClick="CheckSumbit('G')" >
					</td>
				</tr>
			</table>
		</td>
		
    </tr>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel">
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
</form>
</body>
</html>
