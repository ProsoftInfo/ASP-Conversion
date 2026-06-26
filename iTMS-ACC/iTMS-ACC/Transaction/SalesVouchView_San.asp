<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	SalesVouchView_San.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Srehari M
	'Created On					:	March 04, 2006
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
<!--#include file="../../include/populate.asp"-->
<%
Dim objRs,objRs1,objRs2,iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo
dim sDiscPer,dBasicTotal,dDisTotal,dInvAmount,iFromApp
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue
dim iTransNo,sOrgName,sBookName,sParType,sParSubType,sParCode,sBookNo
Dim sAccHeadName,iRoundVal,sOrgPartyCode,sPurType,sParSubTypeName
Dim Objfs,oDOM,sExp,TempNode,sAgentName,Root,dParVal,sSelAccName,sRetVal,iSalTransNo
Dim dtInvDate
iTransNo=Request("TransNo")
Set Objfs = Server.CreateObject("Scripting.FileSystemObject")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")

sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)

'Response.Write iTransNo
if objfs.FileExists(Server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")) then
	oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
	Set Root = oDOM.documentElement

	IF Root.haschildnodes then

		sExp = "//Agent"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			sAgentName = TempNode.Item(0).Attributes.getNamedItem("Agentname").Value
		Else
			sAgentName = "NA"
		End IF

	End IF
Else
	sAgentName = "NO"
End IF

IF CStr(sAgentName) = "NO" Then
	sQuery = "Select P.PartyName From App_M_PartyMaster P,Sal_T_AdditionalAgents A,Acc_T_VoucherHeader H  "&_
			 "Where P.PartyCode = A.AgentCode and H.CreatedTransNo = "&iTransNo&" and  "&_
			 "H.TransactionNumber = A.AccTransactionNo "

	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sAgentName = objRs(0)
	Else
		sAgentName = "NA"
	End IF

	objRs.Close
End IF

sQuery = "Select VoucherAmount,isnull(OtherApplnTransNo,0),isNull(FromApplication,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo
objRs.Open sQuery,con
IF Not objRs.EOF Then
	dParVal = objRs(0)
	iSalTransNo = objRs(1)
	iFromApp = objrs(2)
End IF
objRs.Close
'Response.Write iSalTransNo




'Response.Write iTransNo

	sQuery="SELECT H.BookCode, H.BookNumber, H.PayToRecdFrom, V.OrgUnitShortDescription, V.BookName,P.PartyName,P.SubTypeName,H.CreatedVoucherNo,Convert(Varchar,H.VoucherDate,103) as VoucherDate FROM Acc_T_CreatedVoucherHeader H " _
	& "INNER JOIN VwOrgBookNames V ON H.BookCode = V.BookCode AND H.BookNumber = V.BookNumber INNER JOIN VwOrgParty P ON  "&_
	" H.PartyType = P.PartyType AND H.PartySubType = P.PartySubType AND H.PartyCode = P.PartyCode  "&_
	" WHERE H.CreatedTransNo ="& iTransNo&" and H.OUDEFINITIONID = V.OUDEFINITIONID "

	' Response.Write sQuery

	with objRs
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source=sQuery
		.Open
	end with

	if not objRs.EOF then
		sOrgName=objRs("OrgUnitShortDescription")
		sBookName=objRs("BookName")
		sRefernceNo=objRs("CreatedVoucherNo")
		sParSubTypeName=objRs("SubTypeName")
		sPartyName=objRs("PartyName")
		dtInvDate = objRs("VoucherDate")
	end if
	objRs.Close
	IF trim(dtInvDate) = "" then
		sQuery = "Select Convert(Varchar,VoucherDate,103) as VoucherDate FROM Acc_T_CreatedVoucherHeader where CreatedTransNo ="& iTransNo&"  "
		objRs.Open sQuery,con
		if not objRs.EOF then
			dtInvDate = objRs(0)
		end if
		objRs.Close
	End IF
	'Response.Write "<p>sOrgName="&sOrgName&"<BR><BR>"
	'Newly Added onb July 14th 2008 ---To Display Unit,Inv.No,Date... if Book no is Null
	If trim(sOrgName) = "" then
		sQuery="SELECT H.OUDefinitionID, H.BookNumber, H.PayToRecdFrom,P.PartyName,P.SubTypeName,H.CreatedVoucherNo FROM "&_
				" Acc_T_CreatedVoucherHeader  H INNER JOIN VwOrgParty P ON H.PartyType = P.PartyType AND H.PartySubType = P.PartySubType AND "&_
				" H.PartyCode = P.PartyCode WHERE H.CreatedTransNo = "&iTransNo
	'		Response.Write sQuery

		with objRs
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source=sQuery
			.Open
		end with
		if not objRs.EOF then
			'sOrgName=populateUnitSelected(objRs(0))
			sOrgName=objRs(0)
			sRefernceNo=objRs("CreatedVoucherNo")
			sParSubTypeName=objRs("SubTypeName")
			sPartyName=objRs("PartyName")

		end if
		objRs.Close

	End If
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Response.Write "dtInvDate="&dtInvDate
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
'*************************************************************************
Function PrintInvoice(sInvNo)
	PrintWindow "../Reports/PRNInvoiceDetails.asp?hVouNo=" & trim(sInvNo)
End Function
'*****************************************************************************
Function ViewInvoice()
	iSalTransNo = document.formname.hSalTransNo.value
	sModule = "A"
	'alert iSalTransNo
	IF iSalTransNo <> 0 then showModalDialog "../../Sales/Transaction/SaltrInvoiceDisplay.asp?InvNo="&iSalTransNo&"&Module="&sModule,"","dialogHeight:550px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No"
End Function
'*****************************************************************************
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="VouPURBookSelection.asp">
<input type=hidden name=hSalTransNo value="<%=iSalTransNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher View &nbsp;
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
                                &nbsp;
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly" width="100%">
                                <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                <tr>
                            <td class="FieldCellsub">Unit </td>
                            <td width="160" class="FieldCellSub"><span class="DataOnly"><%=populateUnitSelected(sOrgName)%>&nbsp;</span></td>
                            <td class="FieldCellSub">Voucher No. - Date</td>
                            <td class="FieldCellSub" width="160">	<span class="DataOnly"><%=sRefernceNo%> - <%=dtInvDate%>&nbsp;</span></td>
                                </tr>

                                <tr>
									<td class="FieldCellSub">Book Name</td>
									<td class="FieldCellSub" width="160"><span class="DataOnly"><%=sBookName%>&nbsp;</span></td>
									<td class="FieldCellsub">Party Name </td>
									<td class="FieldCellSub" colspan="2" ><span class="DataOnly"><%=sPartyName%>&nbsp;</span>
                                </tr>

                                <tr>
									<td class="FieldCellsub">Party Sub Type </td>
		                            <td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=sParSubTypeName%>&nbsp;</span></td>
								</tr>

								<%IF CStr(sAgentName) <> "NA" Then %>
								<tr>
									<td class="FieldCellsub">Agent Name </td>
		                            <td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=sAgentName%>&nbsp;</span></td>
								</tr>
								<%End IF %>

                                <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            <tr>
                <td></td>
                <td valign="top" width="100%">
                <div class="frmBody" id="frm2" style="width: 100%; height:200;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2"> Item Description</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Basic<br>
    Value</td>
    <td class="ExcelHeaderCell" align="center" colspan="2">Discount</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Nett<br>
    Basic</td>
        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="25">%</td>
    <td class="ExcelHeaderCell" align="center" width="60">Value</td>
        </tr>
<%
	dDisTotal=0
	sQuery= "Select D.AccUnitAccountHead,D.BasicAmount,D.Amount,D.DiscountPercent,D.DiscountAmount,G.AccountDescription,D.ItemDescription from Acc_T_CreatedVoucherDetails D inner join " _
	& "Acc_M_GLAccountHead G on D.AccUnitAccountHead=G.AccountHead where D.CreatedTransNo="& iTransNo
	with objRs
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source=sQuery
		.Open
	end with
	set objRs.ActiveConnection=nothing
	do while not objRs.EOF
		iSno=iSno+1
		sAccHeadName=objRs("ItemDescription")
		sValue=objRs("Amount")
		sDiscPer=objRs("DiscountPercent")
		sDiscount=objRs("DiscountAmount")
		sAmount=objRs("Amount")
		dDisTotal= CDbl(dDisTotal)+CDbl(sAmount)

		sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where AccountHead = "&Objrs(0)&" "
		Objrs1.Open sQuery,Con
		IF Not Objrs1.Eof Then
			sSelAccName = Trim(objrs1(0))
		End IF
		Objrs1.Close
%>
    <tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
		<td class="ExcelDisplayCell"><%=sSelAccName%> / <%=sAccHeadName%></td>
		<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sValue,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right" width="25"><%=FormatNumber(sDiscPer ,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber( sDiscount,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sAmount,2,,,0)%></td>
    </tr>
<%
	objRs.MoveNext
	loop
	objRs.Close


	with objRs1
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source="Select T.AccountHead,IsNull(T.TaxPercentage,0) TaxPercentage,T.TaxAmount,T.TransCrDrIndication,T.TaxCode,T.TaxCategoryCode,G.AccountDescription from Acc_T_CreatedVoucherTaxDet T inner join Acc_M_GLAccountHead G " _
				& "on T.AccountHead=G.AccountHead where T.CreatedTransNo="& iTransNo &" and T.TaxAmount<>0"
		.Open
	end with
	set objRs1.ActiveConnection=Nothing
	Do while not objRs1.EOF
		iSno=iSno+1
		sAccHeadName=objRs1("AccountDescription")
		dTaxValue =objRs1("TaxPercentage")
		dTax =objRs1("TaxAmount")
		iRoundVal=objRs1("TransCrDrIndication")
		IF CStr(iRoundVal) = "D" Then
			dTax = Cdbl(dTax) * - 1
		End IF

%>
	<tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
		<td class="ExcelDisplayCell"> <%=sAccHeadName%></td>
		<td class="ExcelDisplayCell"> </td>
		<td class="ExcelDisplayCell"> </td>
		<%if CStr(dTaxValue)  <> "0" then%>
			<td class="ExcelDisplayCell" align="Right" width="60"><%=dTaxValue%></td>
		<%else%>
			<td class="ExcelDisplayCell" align="Right" width="60"></td>
		<%end if

		  if iRoundVal="C" and objRs1("TaxCode")=0 and objRs1("TaxCategoryCode") then
		  dTax="-"&dTax

		  %>
			<td class="ExcelDisplayCell" align="Right"><%=FormatNumber(dTax ,2,,,0)%></td>
		<%else%>
			<td class="ExcelDisplayCell" align="Right"><%=FormatNumber(dTax ,2,,,0)%></td>
		<%end if%>
    </tr>
<%
	'Response.Write dDisTotal & dTax
	dDisTotal=CDbl(dDisTotal)+CDbl(dTax)
	objRs1.MoveNext
	loop
	objRs1.Close

	sQuery="Select Sum(Amount) BasicTotal,Sum(Amount) Total from Acc_T_CreatedVoucherDetails where CreatedTransNo="& iTransNo
	with objRs
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source=sQuery
		.Open
	end with
	set objRs.ActiveConnection=nothing
		dBasicTotal=objRs("BasicTotal")
		dTotal=objRs("Total")
	objRs.Close
 %>

        <tr>
    <td align="center" ></td>
    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dBasicTotal,2,,,0)%></b></td>
    <td class="ExcelDisplayCell" align="center" width="25">    </td>
    <td class="ExcelDisplayCell" align="right" width="60"><b></b></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
        </tr>


<%
	dInvAmount=dDisTotal
%>


        <tr>
        <td align="center" ></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <%=FormatNumber(dParVal,2,,,0)%> </td>
        </tr>
            </table>
                </div>
                </td>

								<tr>
								<td align="center" width="5" class="ClearPixel">
                                &nbsp;
								</td>
								<td valign="top" class="FieldCell" height="10">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="130" valign="top">Amount </td>
                            <td>

                            <span class="DataOnly"><%=AmountWords(dParVal)%></span>
                            </td>
                                    </table>
							</tr>
							 <tr>
              <td></td>
            <td valign="top" width="100%">

            <div class="frmBody" id="frm2" style="width: 100%; height:150;">

        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
        <td class="FieldCell" colspan="5" valign="top"><B> Receipt History </B></td>
                            <td>
                            </tr>
        <tr>

			<td class="ExcelHeaderCell" align="center"  rowspan="2"	width="20">S.No.</td>
			<td class="ExcelHeaderCell" align="center" colspan="2" width="160">Voucher</td>
			<td class="ExcelHeaderCell" align="center" rowspan="2" width="80">Amount Received</td>
			<td class="ExcelHeaderCell" align="center" rowspan="2" width="80">Received On</td>
			<td class="ExcelHeaderCell" align="center" colspan="3" width="280">Instrument</td>

        </tr>

        <tr>
			<td class="ExcelHeaderCell" align="center" width="80" >No</td>
			<td class="ExcelHeaderCell" align="center" width="80">Date</td>
			<td class="ExcelHeaderCell" align="center" width="120">Type</td>
			<td class="ExcelHeaderCell" align="center" width="80">No</td>
			<td class="ExcelHeaderCell" align="center" width="80">Date</td>
        </tr>
        <% Dim iCtr,iRcvbNo,iRcvdbyTranNo,iRcvdOn,iAmtRecvd,iVouNo,dtVouDate,sTransType,iCrTransNo,sBkInsType,BkInsNo,dtBkInsDate,dTotAmtRcvd,dBalAmt,iNoOfDays,sFlag1,sFlag2,iAdjAmt,sUserId
        sQuery = "Select ReceivableNumber from Acc_T_Receivables where TransactionNumber = (Select TransactionNumber from Acc_T_VoucherHeader where CreatedTransno = "& iTransNo &" ) "

        'Response.Write sQuery &"<BR>"
        objRs.Open sQuery,con
        iCtr = 1
        dTotAmtRcvd = 0

        sFlag1 = False
        sFlag2 = False
		if objRs.Eof then sFlag1 = False
        if not objRs.EOF then
			iRcvbNo = objRs(0)

			sQuery = "Select RecdByTransactionNo,Convert(VarChar,ReceivedOn,103) as ReceivedOn,AmountReceived from Acc_T_RcvblAdjustmentDetails where  ReceivableNumber = "& iRcvbNo &" "

			objRs2.Open sQuery,con
			do while not objRs2.EOF

				iRcvdbyTranNo = objRs2(0)
				iRcvdOn	= objRs2(1)
				iAmtRecvd    = objrs2(2)
				dTotAmtRcvd = dTotAmtRcvd + CDbl(iAmtRecvd)

				'Response.Write iRcvdbyTranNo &"==="& iAmtRecvd
				If trim(iRcvdbyTranNo) <> "" and trim(iAmtRecvd) > "0" then
					sFlag1 = true
					sQuery = "Select VoucherNumber,Convert(VarChar,VoucherDate,103) as VoucherDate,TransactionType,CreatedTransNo from Acc_T_Voucherheader where transactionnumber = "& iRcvdbyTranNo &" "
					objRs1.Open sQuery,con
					if not objRs1.EOF then
						iVouNo     = objRs1(0)
						dtVouDate  = objRs1(1)
						sTransType = objrs1(2)
						iCrTransNo = objrs1(3)
					end if
					objRs1.Close
					If trim(sTransType) = "BAR" then
						sQuery = "Select BankInstrumentType,BankInstrumentNo,convert(VarChar,BankInstrumentDate,103) from acc_T_CreatedVoucherInstrumentDet where CreatedTransno = " & iCrTransNo & " "
						'Response.Write sQuery
						objRs1.Open sQuery,con
						if not objRs1.EOF then
							sBkInsType  = objRs1(0)
							BkInsNo     = objRs1(1)
							dtBkInsDate = objrs1(2)
						else
							sBkInsType = ""
							BkInsNo = ""
							dtBkInsDate = ""
						end if
						objRs1.Close
					Else
						sBkInsType = "Cash"
					End IF
					%>
					 <tr>
						<td class="ExcelSerial" align="Center" width="20"><%=iCtr%></td>
						<td class="ExcelDisplayCell" align="Right"  width="80"><%=iVouNo%></td>
						<td class="ExcelDisplayCell" align="center" width="80"><%=dtVouDate%></td>
						<td class="ExcelDisplayCell" align="Right" width="80"><%=FormatNumber(iAmtRecvd,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="center" width="60"><%=iRcvdOn%></td>
						<td class="ExcelDisplayCell" align="left" width="120"><%=sBkInsType%></td>
						<td class="ExcelDisplayCell" align="Right" width="80"><%=BkInsNo%></td>
						<td class="ExcelDisplayCell" align="center" width="80"><%=dtBkInsDate%></td>
					</tr>

			<%
			'Else

			 iCtr = iCtr + 1
			End If 'If trim(iRcvdbyTranNo) <> "" and trim(iAmtRecvd) > "0" then

			objrs2.MoveNext

        loop
        objRs2.Close
        'Newly added by S.Maheswari on 09-MAR-09 to display entries from new table Acc_T_OutstandingClosingHistory
		sQuery = "Select isNull(ReasonForClosing,''),AmountAdjusted,convert(Varchar,ClosedOn,103),ClosedBy from Acc_T_OutstandingClosingHistory where ReceivableNumber = "& iRcvbNo &" "
		objRs1.Open sQuery,con
		If objRs1.EOF   then
			sFlag2 = False
		End IF
		do while not objRs1.EOF
			sFlag2 = True
			iAdjAmt = objrs1(1)

			'To get User id
			sQuery = "Select LoginId from Ms_EmployeeMaster where EmployeeNumber = " & objRs1(3)
			objRs2.Open sQuery,con
			if not objRs2.EOF then
				sUserId = objRs2(0)
			end if
			objRs2.Close
			%>
			<tr>
				<td class="ExcelSerial" align="Center" width="20"></td>
				<td class="ExcelDisplayCell" align="Center" Colspan = "2">Invoice Closed</td>
				<td class="ExcelDisplayCell" align="Right" width="80"><%=FormatNumber(objrs1(1),2,,,0)%></td>
				<td class="ExcelDisplayCell" align="center" width="80"><%=objRs1(2)%></td>
				<td class="ExcelDisplayCell" align="left"  Colspan = "3"><%=objrs1(0)%> - <%=sUserId%></td>
			</tr>
		<%	objRs1.MoveNext
		loop
		objRs1.Close

	end if
	objRs.Close
	'Response.Write sFlag1  &"--"&sFlag2
    If (trim(sFlag1) = "False") and (trim(sFlag2) = "False") then
		'Response.Write sFlag1  &"--"&sFlag2
	%>
		<tr>
			<td class="ExcelSerial" align="Center" width="20"></td>
			<td class="ExcelDisplayCell" align="Center"  Colspan = "7">No receipt entry available for this sales invoice</td>
		</tr>

	<%End If

    dBalAmt = CDbl(dParVal) - CDbl(dTotAmtRcvd)

     'Response.Write "dtInvDate="&dtInvDate
    IF trim(dBalAmt) > 0 then  iNoOfDays = DateDiff("d",(mid(dtInvDate,4,2)&"/"&mid(dtInvDate,1,2)&"/"&mid(dtInvDate,7,4)),date()) + 1
    'Response.Write "iAmtRecvd ="&iAmtRecvd & sFlag1
    'IF trim(iAmtRecvd) > "0"  and sFlag <> false then
    IF (trim(iAdjAmt) > "0" and sFlag2 <> False) or (trim(iAmtRecvd) > "0" and sFlag1 <> False) then
    %>
        <tr>
			<td class="ExcelSerial" align="Right" Colspan="3" ><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
			<td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotAmtRcvd,2,,,0)%></b></td>
        </tr>
        <tr>
			<td class="ExcelSerial" align="Right" Colspan="3" ><p align="right"><b>Balance To Receive</b>&nbsp;&nbsp;</td>
			<td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dBalAmt,2,,,0)%></b></td>
        </tr>
        <tr>
			<td class="ExcelSerial" align="Right" Colspan="3" ><p align="right"><b>Outstanding in days</b>&nbsp;&nbsp;</td>
			<td class="ExcelDisplayCell" align="right"><b><%=iNoOfDays%></b></td>
        </tr>
	<% End IF 'IF trim(iAmtRecvd) > "0" then%>
       </table>
       </div>
       </td></tr>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
                                                <%if trim(iFromApp)<>"0" then%>
													<input type="button" value="View Invoice" name="B3" class="ActionButtonX" onclick="ViewInvoice()" >
												<%else%>
													<input type="button" value="Print" name="B4" class="ActionButtonX" onclick="PrintInvoice('<%=iTransNo%>')" >
												<%end if%>
													<input type="button" value="Done" name="B2" class="ActionButton" onclick="window.close()" >

											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
</BODY>
</html>
