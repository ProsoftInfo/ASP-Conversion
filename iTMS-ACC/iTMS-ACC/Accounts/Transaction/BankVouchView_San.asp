
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BankVouchView_San.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:
	'Modified By				:	KUMAR K A
	'Created On					:	August 4, 2006
	'Modified On				:	December 10, 2007
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
<!--#include file="../../include/populate.asp"-->
<%

'XML DOM Variables
Dim oDOM,nodHeader,Root,objRs,sQuery,objRsTemp
dim sNarration,sAccount,sAddtional,iSno,sNarr
dim dTotal,sOrgId,sAccHead,sType
dim EntryNode,HeaderNode,dAmount,sAccNo,sPartyName
dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim iTransNo,iBkHeadCode,bOtherUnit,iTdsAmount,sExp,TempNode,sBankInsDet
Dim sInstrType,sBankInsNo, sBankInsName,sBankInsDrawnOn,sBankInsDt
Dim sAddress1, sAddress2,sCity,sState,sPostcode,sTranIndication,sTranEntryIndication
Dim iCreatedBy,sCreatedOn,sVouStatus,sEmpName,iPartyCtrlAcc,sAdjType
Dim iEntryNo,iHeadOfAcc,sHeadOfAccName,iHeadOfAccAmt,iPartyCode,iEntryAmt,iCtr
Dim iNetAmtPaid,iTotRecovered,sNarrFlag,sAddnFlag,iTotAddnlAmt,sAdjFlag,sAdjOn,iTotAdj
Dim iAccUnitPartyCode,sDetFlag,sBankInsAmt,iAccNo
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request("TransNo")
'Response.Write iTransNo

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")

'set Root=oDOM.documentElement
'sOrgId=Root.Attributes.Item(0).nodeValue
'sOrgName=Root.Attributes.Item(1).nodeValue
'iBookCode =Root.Attributes.Item(2).nodeValue
'sBookName =Root.Attributes.Item(3).nodeValue
'sVouType=Root.Attributes.Item(4).nodeValue
'sVoucDate=Root.Attributes.Item(5).nodeValue
'iBkHeadCode=Root.Attributes.Item(6).nodeValue

'iVouNo=Root.Attributes.Item(9).nodeValue
'sApprove=Root.Attributes.Item(7).nodeValue

set objRs = Server.CreateObject("ADODB.Recordset")
set objRsTemp = Server.CreateObject("ADODB.Recordset")

sQuery = "Select OUDefinitionID,BookNumber,TransactionType,Convert(Char,VoucherDate,103),AccountHead, "&_
		 "CreatedVoucherNo From Acc_T_CreatedvoucherHeader Where CreatedTransNo = "&iTransNo
'Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

if not 	objRs.EOF then
	sOrgId = objRs(0)
	iBookCode = objRs(1)
	sVouType = objRs(2)
	sVoucDate = objRs(3)
	iBkHeadCode = objRs(4)
	iVouNo = objRs(5)
End IF
objRs.Close

sQuery="select OtherUnitTransaction from vwOrgBookNames where OUDefinitionID = '" & sOrgId & "'"&_
"and BookNumber="&iBookCode&" and BookCode='02'"
'Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

if not 	objRs.EOF then
	bOtherUnit=objRs(0)
else
	bOtherUnit=0
end if
objRs.Close
'Newly added on 10th Feb 2009 by S.Maheswari to fetch account no. from Acc_M_BankDetails table
sQuery = "Select isnull(AccountNo,'') from Acc_M_BankDetails where  BookNumber="&iBookCode&" and BookCode='02'"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

if not 	objRs.EOF then
	iAccNo=objRs(0)
end if
objRs.Close

'==============================Voucher View header queries starts here ============================

sQuery = "select Distinct CreatedVoucherNo,TransactionType,OUDefinitionID,PayToRecdFrom," &_
		"PartyCode,CreatedBy,Convert(varchar,CreatedOn,103),CreatedVouchStatus," &_
		"BankInstrumentType,BankInstrumentNo,Convert(varchar,BankInstrumentDate,103),PayableAt,DrawnOnBank,AccUnitPartyCode" &_
		" from VW_Created_BankVoucherView where CreatedTransNo="& iTransNo
' Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iVouNo = objRs(0)
	sVouType = objRs(1)
	sOrgId = objRs(2)
	sPayTo = objRs(3)
	iPartyCode = objRs(4)
	iCreatedBy = objRs(5)
	sCreatedOn = objRs(6)
	sVouStatus = objRs(7)
	sInstrType = objRs(8)
	'sBankInsNo = objRs(9)
	'sBankInsDt = objRs(10)
	'sBankInsDrawnOn = objRs(12)
	iAccUnitPartyCode = objrs("AccUnitPartyCode")
	'Response.Write "===<BR>"&iAccUnitPartyCode
	'Response.Write "===<BR>"&sPayTo
	'If sInstrType = "C" Then
	'	sBankInsName = "Cheque "
	'Elseif sInstrType = "D" Then
	'	sBankInsName = "Demand Draft "
	'ElseIf sInstrType = "B" Then
	'	sBankInsName = "Bankers Cheque "
	'Else
	'	sBankInsName = "Telegraphic Transfer "
	'End If


end if
objRs.Close

'Newly added on July 2nd 2008 by S.Maheswari to fetch bank details from Acc_T_CreatedVoucherInstrumentDet table instead of taking from CreatedVoucherHeader table
sQuery = "Select BankInstrumentNo,convert(Varchar,BankInstrumentDate,103),BankInstrumentType,InstrumentAmount,PayableAt,DrawnOnBank "&_
		" from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo = "&iTransNo&" Order by InstrumentEntryNo "
 'Response.Write sQuery
objRs.Open sQuery,con
If Not objRs.EOF then
	Do while not objRs.EOF
		sBankInsNo		= sBankInsNo&","&objRs(0)
		sBankInsDt		= sBankInsDt&","&objRs(1)
		sBankInsName	= sBankInsName&","&objRs(2)
		sBankInsAmt	    = sBankInsAmt&","&objRs(3)
		sBankInsDrawnOn = sBankInsDrawnOn&","&objRs(5)
		objRs.MoveNext
	loop
End If
objRs.Close
sBankInsNo		= mid(sBankInsNo,2)
sBankInsDt		= mid(sBankInsDt,2)
sBankInsName	= mid(sBankInsName,2)
sBankInsAmt	    = mid(sBankInsAmt,2)
sBankInsDrawnOn = mid(sBankInsDrawnOn,2)
'==================================================================================================

If trim(sVouStatus) = "010104" Then
	sQuery = "select VoucherNumber from ACC_T_VoucherHeader where CreatedTransNo="&iTransNo
	with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
	end with
	if not 	objRs.EOF then
		iVouNo = objRs(0)
	End If
	objRs.Close
End If
set objRs.ActiveConnection = nothing
If trim(sVouType) = "BAP" Then
	sTranIndication = "D"
	sTranEntryIndication = "C"
Else
	sTranIndication = "C"
	sTranEntryIndication = "D"
End If
sQuery = "Select VoucherEntryNumber,AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranIndication&"'"
'Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
'Response.write sQuery
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iEntryNo = objRs(0)
	iHeadOfAcc = objRs(1)
	iHeadOfAccAmt = FormatNumber(objRs(2),2,,,-2)
	sNarr = objRs(3)
	iPartyCtrlAcc = objRs(4)

Else
	iEntryNo = 0
	iHeadOfAccAmt = 0
	iPartyCtrlAcc = 0
end if
objRs.Close
'Response.Write iPartyCtrlAcc
IF iHeadOfAcc <> "" Then
	sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
Else
	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
End If
'Response.write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sHeadOfAccName = objRs(0)
end if
objRs.Close

sQuery = "Select OrgUnitDescription,Address1,Address2,City,State,PostCode from DCS_OrganizationUnitDefinitions where OUDefinitionID='"&sOrgId&"'"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sOrgName = objRs(0)
	sAddress1 = objRs(1)
	sAddress2 = objRs(2)
	sCity = objRs(3)
	sState = objRs(4)
	sPostcode = objRs(5)
else
	sOrgName = ""
	sAddress1 = ""
	sAddress2 = ""
	sCity = ""
	sState = ""
	sPostcode = ""
end if
objRs.Close

If iPartyCode <> "0" Then
	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCode

	With Objrs
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.Open
	End With
	set objRs.ActiveConnection = nothing
	IF not objRs.EOF  Then
		sPartyName = Trim(Objrs(0))
	End IF
	objRs.Close
End If

sEmpName = Trim(session("userid"))

'==============================Voucher View header queries ends here ============================
'sExp = "//BankInstrumentDet"
'Set tempNode = Root.selectNodes(sExp)



'	For each EntryNode in Root.childNodes
'		if EntryNode.nodeName="Entry" then
'			sPayTo = EntryNode.Attributes.Item(2).nodeValue
'		end if
'		For each HeaderNode in EntryNode.childNodes
'			if HeaderNode.nodeName="AccHead" then
'				sAccNo = Split(HeaderNode.attributes.item(0).nodeValue,"?")
'				sType = HeaderNode.Attributes.Item(4).nodeValue
'			end if
'			if HeaderNode.nodeName="Narration" then
'				sNarr= HeaderNode.text
'			end if
'		next
'	next


sQuery = "Select isNull(BankInstrumentType,''),ISNull(BankInstrumentNo,''),isNull(BankInstrumentDate,''), "&_
		 "isNull(PayableAt,'') From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

if not 	objRs.EOF then
'	sPayTo = objRs(3)
	sAccNo = objRs(1)
	sType = objRs(0)

End IF
objRs.Close
'Response.Write "sType="&sType&"<br>"
	IF sType = "P" Then
			sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&sAccNo(3)

			With Objrs
				.ActiveConnection = con
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.Open
			End With
			Set Objrs.ActiveConnection = nothing

			IF not objRs.EOF  Then
				sPartyName = Trim(Objrs(0))
			End IF
			objRs.Close
	End IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS -
		  <%
			if sVouType="BAP" then
				Response.Write "Cheque Payment Voucher"
			else
				Response.Write "Cheque Receipt Voucher"
			end if
		%>
</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="vbscript">
Function DispChequeAmt()
	If sDetFlag = True then chqAmt.innerHTML = "Rs. "& netAmt.innerHTML
End Function
FUNCTION FinalCheck(Flag)
dim iUserid,iTransNo
Dim sStatus,sValue,sOrgName,sTransNo,sVouTy
	IF document.formname.hApprover.value="Y" THEN
		IF document.formname.selUserid.selectedIndex> 0 THEN
			iUserid=document.formname.selUserid.value
			iTransNo=document.formname.hTransNo.value
			SET objhttp = CreateObject("MSXML2.XMLHTTP")
			objhttp.Open "POST","XMLVouAppUpdate.asp?BkCode=CA&TransNo="& iTransNo &"&User="& iUserid &"&Mode=E", false
			objhttp.send
			IF trim(objhttp.responseText)<>"" THEN
				MsgBox objhttp.responseText
				exit function
			END IF
		ELSE
			MsgBox ("Select Approver")
			document.formname.selUserid.focus
			exit function
		END IF
	END IF

	IF Flag="B" THEN
		document.formname.action="VouCABookSelection.asp"
		document.formname.submit
	ELSEIF Flag="PV" THEN
		document.formname.action="VouCAEntry.asp"
		document.formname.selVouType.value="C"
		document.formname.submit
	ELSEIF Flag="RV" THEN
		document.formname.action="VouCAEntry.asp"
		document.formname.selVouType.value="D"
		document.formname.submit

	END IF
END FUNCTION

Function CheckPrint()
	Dim sStatus,sValue,sTransNo,sVouTy,objhttp
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	sTransNo = document.formname.hTransNo.value
	sOrgName = document.formname.hOrgName.Value
	sVouTy = document.formname.selVouType.Value

	sValue = sTransNo
	'Added newly on  02 Sep 08 to insert Print details in table
	sBkCode = "02"
	sUserId = document.formname.hUserId.value
	objhttp.open "GET","PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId, False
	objhttp.send
	sRetVal = objhttp.responseText


	'showModalDialog "PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId&"&TransNo="&sTransNo,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	'sStatus= showModalDialog("PRNBankRecpVouViewForCheck.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
	sStatus= showModalDialog("PRNBankRecpVouView1.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")

End Function
Function PaymentAdvPrint()
Dim sStatus,sValue,sTransNo,sVouTy,objhttp
	Set objhttp = CreateObject("Microsoft.XMLHTTP")

	sTransNo = document.formname.hTransNo.value
	sOrgName = document.formname.hOrgName.Value
	sVouTy = document.formname.selVouType.Value

	sValue = sTransNo
	'Added newly on  02 Sep 08 to insert Print details in table
	sBkCode = "02"
	sUserId = document.formname.hUserId.value
	objhttp.open "GET","PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId, False
	objhttp.send
	sRetVal = objhttp.responseText

	'showModalDialog "PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId&"&TransNo="&sTransNo,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sStatus= showModalDialog("PRNPaymentAdvice.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")

End Function
Function VoucherPrint()
	Dim sStatus,sValue,sTransNo,sVouTy,objhttp
	Set objhttp = CreateObject("Microsoft.XMLHTTP")

	sTransNo = document.formname.hTransNo.value
	sOrgName = document.formname.hOrgName.Value
	sVouTy = document.formname.selVouType.Value

	sValue = sTransNo
	'Added newly on  02 Sep 08 to insert Print details in table
	sBkCode = "02"
	sUserId = document.formname.hUserId.value
	objhttp.open "GET","PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId, False
	objhttp.send
	sRetVal = objhttp.responseText

	'IF trim(sRetVal) = "" then
	'	MsgBox("Print Details Updated Successfully")
	'End IF
	'showModalDialog "PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId&"&TransNo="&sTransNo,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sStatus= showModalDialog("PRNBankReceiptVoucher.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")

End Function
Function VoucherEntry()
	Dim sStatus,sValue,sTransNo,sVouTy,objhttp
	Set objhttp = CreateObject("Microsoft.XMLHTTP")

	sTransNo = document.formname.hTransNo.value
	sOrgName = document.formname.hOrgName.Value
	sVouTy = document.formname.selVouType.Value

	sValue = sTransNo
	'Added newly on  02 Sep 08 to insert Print details in table
	sBkCode = "02"
	sUserId = document.formname.hUserId.value
	objhttp.open "GET","PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId, False
	objhttp.send
	sRetVal = objhttp.responseText

	'IF trim(sRetVal) = "" then
	'	MsgBox("Print Details Updated Successfully")
	'End IF
	'showModalDialog "PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId&"&TransNo="&sTransNo,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sStatus= showModalDialog("PRNBankVoucherEntry.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
End Function
Function ReceiptPrint()
	Dim sStatus,sValue,sTransNo,sVouTy,objhttp
	Set objhttp = CreateObject("Microsoft.XMLHTTP")

	sTransNo = document.formname.hTransNo.value
	sOrgName = document.formname.hOrgName.Value
	sVouTy = document.formname.selVouType.Value

	sValue = sTransNo
	'Added newly on  02 Sep 08 to insert Print details in table
	sBkCode = "02"
	sUserId = document.formname.hUserId.value
	objhttp.open "GET","PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId, False
	objhttp.send
	sRetVal = objhttp.responseText

	IF trim(sRetVal) = "" then
		MsgBox("Print Details Updated Successfully")
	End IF
	'showModalDialog "PrintInsert.asp?BkCode="&sBkCode&"&UserId="&sUserId&"&TransNo="&sTransNo,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	sStatus= showModalDialog("PRNBankReceiptPrint.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")

End Function
</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" OnLoad="DispChequeAmt()">
<form method="POST" name="formname" action="">
<input type="hidden" name="selUnitId" value="<%=sOrgId%>">
<input type="hidden" name="horgName" value="<%=sOrgName%>">
<input type="hidden" name="hNarr" value="<%=mid(sNarr,2)%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="selBook" value="<%=iBookCode%>">
<input type="hidden" name="selVouType" value="<%=sVouType%>">
<input type="hidden" name="hBookOtherUnit" value="<%=bOtherUnit%>">
<input type="hidden" name="hUserId" value="<%=iCreatedBy%>">
<input type="hidden" name="hBookAccHead" value="<%=iBkHeadCode%>">

<input type="hidden" name="hApprover" value="<%=sApprove%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">
				  <%
					if sVouType="BAP" then
						Response.Write "Cheque Payment Voucher"
					else
						Response.Write "Cheque Receipt Voucher"
					end if
				%>
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
											<tr>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sOrgName%></span>
												</td>
												<td class="FieldCellSub" width="72">Voucher No.
												</td>
												<td class="FieldCellSub" width="120">
													<span class="DataOnly"><%=iVouNo %>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sAddress1%><br><%=sAddress2%><br><%=sCity%>,<%=sState%>-<%=sPostcode%></span>
												</td>
												<td class="FieldCellSub" valign="top">Date
												</td>
												<td class="FieldCellSub" valign="top">
													<span class="DataOnly"><%=sVoucDate %>&nbsp;</span>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
											<tr>
												<td class="FieldCellSub" width="100">Head of Account
												</td>
												<td class="FieldCellSub" colspan="3">
													<span class="DataOnly"><%=sHeadOfAccName%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCellSub">Amount
												</td>
												<td class="FieldCellSub" colspan="3">
													<span class="DataOnly"><%=AmountWords(replace(iHeadOfAccAmt,",",""))%></span>
												</td>
											</tr>

											<tr>
												<td class="FieldCellSub">
												 <%
														if sVouType="BAP" then
															Response.Write "Paid To "
														else
															Response.Write "Received From "
														end if
												%>
												</td>
												<td class="FieldCellSub">

														<%' Response.Write "sType="&sType&"<BR>"
														if sType="P" then%>
															<span class="DataOnly"><%=sPartyName %>&nbsp;</span>
														<%else%>
															<span class="DataOnly"><%=sPayTo %>&nbsp;</span>
														<%end if%>
												</td>

												<td class="FieldCellSub" colspan=2>
													<b>Rs.&nbsp;&nbsp;</b><span class="DataOnly"><%=iHeadOfAccAmt%> </span>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" width="100%">
											<tr>
												<td class="FieldCell" rowspan="2" valign="top" >
													<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
														<tr>
															<td class="FieldCellSub" >Name of Bank
															</td>
														<tr>
														</tr>
															<td class="FieldCellSub">
																<span class="DataOnly"><%=sBankInsDrawnOn%> &nbsp;</span>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub">Account Number
															</td>
														<tr>
														</tr>
															<td class="FieldCellSub">
																<span class="DataOnly"><%=iAccNo%>&nbsp;</span>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub"><%=sBankInsName%> Number
															</td>
														<tr>
														</tr>
															<td class="FieldCellSub">
																<span class="DataOnly"><%
																		IF TRIM(sBankInsNo)="0" THEN
																		Response.Write ""
																		ELSE
																		Response.Write sBankInsNo
																		END IF
																%>&nbsp;</span>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub"><%=sBankInsName%> Date
															</td>
														<tr>
														</tr>
															<td class="FieldCellSub">
																<span class="DataOnly"><%=sBankInsDt%>&nbsp;</span>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub"><%=sBankInsName%> Amount
															</td>
														<tr>
														</tr>
															<td class="FieldCellSub">
															<!--span class="DataOnly" ID="chqAmt"><%=iHeadOfAccAmt%>&nbsp;</span-->
																<span class="DataOnly" ID="chqAmt"><%=sBankInsAmt%>&nbsp;</span>
															</td>
														</tr>

													</table>
												</td>
												<td class="FieldCell">Details
												</td>
											</tr>

											<tr>
												<td class="FieldCell" valign="top"  >
													<table border="0" cellspacing="0" cellpadding="0" width="100%" class="TableOutlineOnly">
											<tr>
												<td class="FieldCell" valign="top" >
													<table border="0" cellspacing="0" cellpadding="0" width="100%" >
														<%
															'Fetch the additional Payment / Receipt entries
															sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> "&iEntryNo&" and TransCrDrIndication='"&sTranIndication&"' "
															'Response.Write sQuery
															iCtr = 1
															iTotAddnlAmt = 0
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	iEntryAmt = cdbl(objRs(1))
																	iPartyCtrlAcc = objRs(3)
																	iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
																	sDetFlag = True
																	If iHeadOfAcc <> "" Then
																		sQuery = "Select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																	Else
																		sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																	End If
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close

																	If Not sNarrFlag Then
														%>

														<tr>
															<td class="FieldCell" valign="top" colspan=5><span class="DataOnly"><b><%=UCase(sNarr)%></b></span></td>
														</tr>
														<%
																	sNarrFlag = True
																	End If
															If Not sAddnFlag Then
																If sVouType = "BAP" Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Additonal Payments</u></b></span></td>
														</tr>
														<%
																Else
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Additonal Receipts</u></b></span></td>
														</tr>
														<%
																End If
															sAddnFlag = True
															End If
														%>
														<tr>
															<td class="FieldCellSub" valign="top"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" valign="top">
																<span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
															<td class="FieldCellSub">
															</td>
														<%
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close


														%>
														</tr>
														<%
															'Fetch the Recoveries
															sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_BankVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> "&iEntryNo&" and TransCrDrIndication='"&sTranEntryIndication&"' order by VoucherEntryNumber"
															'Response.Write sQuery
															iCtr = 1
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	iEntryAmt = cdbl(objRs(1))
																	iTotRecovered = iTotRecovered + iEntryAmt
																	sDetFlag = True
																	If iHeadOfAcc <> "" Then
																		sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																	Else
																		sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																	End If
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close

																	If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(sNarr)%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																	End If
																	If iCtr = 1 Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Recoveries</u></b></span></td>
														</tr>
														<%
																	End If
														%>
														<tr>
															<td class="FieldCellSub" colspan="2" valign="top"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" colspan="2" valign="top"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														<%
																'If iCtr < objrs.RecordCount Then
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close
															iNetAmtPaid = iHeadOfAccAmt + iTotAddnlAmt - iTotRecovered

														%>
														</tr>
														<%
															If sNarrFlag Then
														%>
														<tr>
															<!--td class="FieldCellSub" colspan=2 align=right valign="top">
																<b>Total Recovered Rs.&nbsp;</b>
															</td>
															<td class="FieldCellSub" colspan=2 valign="top">
																<span class="DataOnly"><b><%=FormatNumber(iTotRecovered,2,,,-2)%></b></span>
															</td-->
														</tr>
															<%End If%>
														<%
															'Fetch the Receivable Adjustments
															sQuery = "select ReceivableNumber,Convert(varchar,ReceivedOn,103),AmountReceived,AdjustType from Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
															'Response.Write sQuery
															iCtr = 1
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	sAdjOn = objRs(1)
																	iEntryAmt = cdbl(objRs(2))
																	sAdjType = objRs(3)
																	iTotAdj = iTotAdj + iEntryAmt
																	sDetFlag = True
																	sQuery = "select Narration from ACC_T_CreatedReceivables where Receivablenumber = "&iHeadOfAcc
																	'Response.Write sQuery
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close
																	'sHeadOfAccName = sHeadOfAccName &" "& sAdjOn
																	If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(sNarr)%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																	End If
																	If Not sAdjFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Adjustments</u></b></span></td>
														</tr>
														<%
																		sAdjFlag = True
																	End If
															If trim(sAdjType) = "A" Then sHeadOfAccName = "Less Advance Receipts "
														%>
														<tr>
															<td class="FieldCellSub" colspan="2" valign="top"><%=sHeadOfAccName%>:</td>
															<td class="FieldCellSub" colspan="2" valign="top"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														</tr>
														<%
																'If iCtr < objrs.RecordCount Then
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close


														%>
														<%
															'Fetch the Payable Adjustments
															sQuery = "select PayablesNumber,Convert(varchar,PaidOn,103),AmountPaid,AdjustType from Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
															iCtr = 1
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	sAdjOn = objRs(1)
																	iEntryAmt = cdbl(objRs(2))
																	sAdjType = objRs(3)
																	iTotAdj = iTotAdj + iEntryAmt
																	sDetFlag = True
																	sQuery = "select Narration from ACC_T_CreatedPayables where Payablesnumber ="&iHeadOfAcc
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close
																	'sHeadOfAccName = sHeadOfAccName &" "& sAdjOn
																	If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(sNarr)%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																	End If
																	If Not sAdjFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Adjustments</u></b></span></td>
														</tr>
														<%
																		sAdjFlag = True
																	End If
															If trim(sAdjType) = "A" Then sHeadOfAccName = "Less Advance Payments "
														%>
														<tr>
															<td class="FieldCellSub" colspan="2" valign="top"><%=sHeadOfAccName%>:</td>
															<td class="FieldCellSub" colspan="2" valign="top"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														</tr>
														<%
																'If iCtr < objrs.RecordCount Then
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close

															If sDetFlag <> True then %>
																<tr>
																	<td class="FieldCell" valign="top" colspan=5><span class="DataOnly"><b><%=UCase(sNarr)%></b></span></td>
																</tr>
															<% End IF %>
														'To print balance advance amount after the adjustment
														<%If (iNetAmtPaid-iTotAdj) > 0 Then%>
															<!--tr>
																<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u></u></b></span></td>'Advance Receipt / Payment
															</tr-->
															<tr>
																<%if sVouType="BAP" then%>
																	<td class="FieldCellSub" colspan="2" valign="top">Advance Paid:</td>
																<%Else%>
																	<td class="FieldCellSub" colspan="2" valign="top">Advance Received:</td>
																<%End If%>
																<td class="FieldCellSub" colspan="2" valign="top"><span class="DataOnly"><%=FormatNumber((iNetAmtPaid-iTotAdj),2,,,-2)%></span>
																</td>
															</tr>
														<%End IF%>
													</table>
											</td>
											<td valign="bottom" class="FieldCell">
													<table border="0" cellspacing="0" cellpadding="0" width="100%" >
													<%If sDetFlag = True then %>
														<tr>
															<td class="FieldCellSub" valign="top" width="80">Total Amount
															</td>
															<td class="FieldCellSub" valign="top">
																<span class="DataOnly" valign="top"><%=FormatNumber((iHeadOfAccAmt + iTotAddnlAmt),2,,,-2)%></span>
															</td>
														</tr>
														<tr>
															<td class="FieldCellSub" valign="top" >(-) Recovered
															</td>
															<td class="FieldCellSub" valign="top" >
																<span class="DataOnly" ><%=FormatNumber(iTotRecovered,2,,,-2)%></span>
															</td>
														</tr>
														<tr>
															<td class="FieldCellSub" valign="top">Net Amount
															</td>
															<td class="FieldCellSub" valign="top">
																<span class="DataOnly" id="netAmt"><%=FormatNumber(iNetAmtPaid,2,,,-2)%></span>

															</td>
														</tr>
														<%End If %>
													</table>
												</td>
											</tr>
													</table>
												</td>
											</tr>

											<tr>
												<td class="FieldCell" colspan="2">
													<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
														<tr>
															<td class="FieldCellSub">
																<span class="DataOnly"><%=sEmpName%>-<%=sCreatedOn%></span>
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly">-</span>
															</td>
															<td class="FieldCellSub">
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly">-</span>
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly">-</span>
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly">-</span>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub">Prepared By
															</td>
															<td class="FieldCellSub">Checked By
															</td>
															<td class="FieldCellSub" align="right">Passed By:
															</td>
															<td class="FieldCellSub">CA
															</td>
															<td class="FieldCellSub">CE
															</td>
															<td class="FieldCellSub">MD
															</td>
														</tr>

													</table>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">

													<input type="button" value="Print Payment Voucher" name="B8" onClick="CheckPrint()" class="ActionButtonX">

														<input type="button" value="Print Receipt" name="B8" onClick="ReceiptPrint()" class="ActionButton2">

													<input type="button" value="Print St.ment" name="B9" onClick="VoucherPrint()" class="ActionButton2">
													<input type="button" value="Payment Advice" name="B10" onClick="PaymentAdvPrint()" class="ActionButton2">
													<input type="button" value="Print Voucher Entries" name="B11" onClick="VoucherEntry()" class="ActionButton2">
													<input type="button" value="Close" name="B3" class="ActionButton" onclick="window.close()">
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
</body>
