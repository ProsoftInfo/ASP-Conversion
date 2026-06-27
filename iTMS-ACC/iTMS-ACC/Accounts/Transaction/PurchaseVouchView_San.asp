<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PurchaseVouchView_San.asp
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
dim sDiscPer,dBasicTotal,dDisTotal,dInvAmount
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue
dim iTransNo,sOrgName,sBookName,sParType,sParSubType,sParCode,sBookNo
Dim sAccHeadName,iRoundVal,sOrgPartyCode,sPurType,sParSubTypeName
Dim iOtherAppNo,iSuppInvDate,iAdjAmt,sTemp

set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")

iTransNo=Request("TransNo")
'Response.Write iTransNo
iOtherAppNo = 0

	sQuery="SELECT H.BookCode, H.BookNumber, H.PayToRecdFrom, V.OrgUnitShortDescription, V.BookName,P.PartyName,P.SubTypeName,isNull(H.OtherApplnTransNo,0) OtherAppNo FROM Acc_T_CreatedVoucherHeader H " _
	& "INNER JOIN VwOrgBookNames V ON H.BookCode = V.BookCode AND H.BookNumber = V.BookNumber INNER JOIN VwOrgParty P ON  "&_
	" H.PartyType = P.PartyType AND H.PartySubType = P.PartySubType AND H.PartyCode = P.PartyCode WHERE V.OUDefinitionID = H.OUDefinitionID and H.CreatedTransNo ="& iTransNo
	'Response.Write sQuery
	with objRs
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source=sQuery
		.Open
	end with

	set objRs.ActiveConnection=Nothing
	if not objRs.EOF then
		sOrgName=objRs("OrgUnitShortDescription")
		sBookName=objRs("BookName")
		sRefernceNo=objRs("PayToRecdFrom")
		sParSubTypeName=objRs("SubTypeName")
		sPartyName=objRs("PartyName")
		iOtherAppNo = objRs("OtherAppNo")
	end if
	objRs.Close


	'Newly Added onb July 14th 2008 ---To Display Unit,Inv.No,Date... if Book no is Null
	If trim(sOrgName) = "" then
		sQuery="SELECT H.OUDefinitionID, H.BookNumber, H.PayToRecdFrom,P.PartyName,P.SubTypeName,isNull(H.OtherApplnTransNo,0) OtherAppNo FROM "&_
				" Acc_T_CreatedVoucherHeader  H INNER JOIN VwOrgParty P ON H.PartyType = P.PartyType AND H.PartySubType = P.PartySubType AND "&_
				" H.PartyCode = P.PartyCode WHERE H.CreatedTransNo = "&iTransNo
		'Response.Write sQuery

		objrs.Open sQuery,con

		set objRs.ActiveConnection=Nothing
		if not objRs.EOF then
			'sOrgName=populateUnitSelected(objRs(0))
			sOrgName=objRs(0)
			sParSubTypeName=objRs("SubTypeName")
			sPartyName=objRs("PartyName")
			iOtherAppNo = objRs("OtherAppNo")
		end if
		objRs.Close
	'Response.Write "sOrgName="&sOrgName
	End If
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	IF CStr(Trim(iOtherAppNo)) = "" Then
		iOtherAppNo = "0"
	End IF
	'Response.Write "<p><font color=red>sRefernceNo="&sRefernceNo
	IF Cstr(iOtherAppNo) <> "0" Then
		sQuery = "Select SuppInvoiceNo+'-'+Convert(Varchar,SuppInvoiceDate,103),Convert(Varchar,SuppInvoiceDate,103)   "&_
				 "From Rcv_T_InvoiceHeader Where InvoiceNumber = "&iOtherAppNo

		Objrs.Open sQuery,Con
		IF Not Objrs.Eof Then
			sRefernceNo=objRs(0)
			iSuppInvDate = objRs(1)
		End IF
		Objrs.Close
	Else
		sTemp = split(sRefernceNo,"-")
		If UBound(sTemp) > 0 Then
			iSuppInvDate = trim(sTemp(1))
		Else
			iSuppInvDate = Date()
		End IF
	End IF

'	Response.Write "iNoOfDays="&iSuppInvDate
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<script language="javascript">
function ViewInvoice() {
	var applnTransNo = document.formname.hApplnTransNo.value;
	if (String(applnTransNo) !== "0") {
		window.open("../../Purchase/Transaction/RepPurInvoiceDetailspopup.asp?iInvNo=" + encodeURIComponent(applnTransNo), "", "height=470,width=870,toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="VouPURBookSelection.asp">
<input type=hidden name="hApplnTransNo" value="<%=iOtherAppNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Voucher View &nbsp;
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
                            <td class="FieldCellSub">Invoice No. - Date</td>
                            <td class="FieldCellSub" width="160">	<span class="DataOnly"><%=sRefernceNo%>&nbsp;</span></td>
                                </tr>

                                <tr>
									<td class="FieldCellSub">Book Name</td>
									<td class="FieldCellSub" width="160"><span class="DataOnly"><%=sBookName%>&nbsp;</span></td>
									<td class="FieldCellsub">Party Name </td>
									<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=sPartyName%>&nbsp;</span>
                                </tr>

                                <tr>
									<td class="FieldCellsub">Party Sub Type </td>
		                            <td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=sParSubTypeName%>&nbsp;</span></td>
								</tr>

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
                <div class="frmBody" id="frm2" style="width: 100%; height:150;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2"> Account Head</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Quantity</td>
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

	sQuery= "Select D.AccUnitAccountHead,D.BasicAmount,D.Amount,D.DiscountPercent,D.DiscountAmount,G.AccountDescription,D.InvoicedQuantity from Acc_T_CreatedVoucherDetails D inner join " _
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
		sAccHeadName=objRs("AccountDescription")
		sValue=objRs("BasicAmount")
		sDiscPer=objRs("DiscountPercent")
		sDiscount=objRs("DiscountAmount")
		sAmount=objRs("Amount")
		sQty = objRs("InvoicedQuantity")
		dDisTotal= CDbl(dDisTotal)+CDbl(sAmount)
%>
    <tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
		<td class="ExcelDisplayCell"> <%=sAccHeadName%></td>
		<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sQty,3,,,0)%></td>
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
%>
	<tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
		<td class="ExcelDisplayCell"> <%=sAccHeadName%></td>
		<td class="ExcelDisplayCell"> </td>
		<td class="ExcelDisplayCell"> </td>
		<td class="ExcelDisplayCell"> </td>
		<%if CStr(dTaxValue)  <> "0" then%>
			<td class="ExcelDisplayCell" align="Right" width="60"><%=dTaxValue%></td>
		<%else%>
			<td class="ExcelDisplayCell" align="Right" width="60"></td>
		<%end if

		  if iRoundVal="D" and objRs1("TaxCode")=0 and objRs1("TaxCategoryCode") then
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

	sQuery="Select isNull(Sum(BasicAmount),0) BasicTotal,isNull(Sum(Amount),0) Total from Acc_T_CreatedVoucherDetails where CreatedTransNo="& iTransNo
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
    <td class="ExcelDisplayCell" align="center" width="25">    </td>
    <td class="ExcelDisplayCell" align="right" width="60"><b></b></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
        </tr>


<%
	dInvAmount=dDisTotal
%>


        <tr>
        <td align="center" ></td>
    <td class="ExcelSerial" align="right" colspan="5"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <%=FormatNumber(dInvAmount,2,,,0)%> </td>
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

                            <span class="DataOnly"><%=AmountWords(dInvAmount)%></span>
                            </td>
                                    </table>
						 </tr>
	   <tr>
              <td></td>
            <td valign="top" width="100%">

            <div class="frmBody" id="frm2" style="width: 100%; height:150;">

        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
        <td class="FieldCell" colspan="5" valign="top"><B> Payment History </B></td>
                            <td>
                            </tr>
        <tr>

			<td class="ExcelHeaderCell" align="center"  rowspan="2"	width="20">S.No.</td>
			<td class="ExcelHeaderCell" align="center" colspan="2" width="160">Voucher</td>
			<td class="ExcelHeaderCell" align="center" rowspan="2" width="80">Amount Paid</td>
			<td class="ExcelHeaderCell" align="center" rowspan="2" width="80">Paid On</td>
			<td class="ExcelHeaderCell" align="center" colspan="3" width="280">Instrument</td>

        </tr>

        <tr>
			<td class="ExcelHeaderCell" align="center" width="80" >No</td>
			<td class="ExcelHeaderCell" align="center" width="80">Date</td>
			<td class="ExcelHeaderCell" align="center" width="120">Type</td>
			<td class="ExcelHeaderCell" align="center" width="80">No</td>
			<td class="ExcelHeaderCell" align="center" width="80">Date</td>
        </tr>
        <% 
        Response.Write "<font color=red>"
        Dim iCtr,iPybNo,iPdByTranNo,dtPaidOn,iAmtPaid,iVouNo,dtVouDate,sTransType,iCrTransNo,sBkInsType,BkInsNo,dtBkInsDate,dTotAmtPaid,dBalAmt,iNoOfDays,sFlag1,sFlag2,sUserId
        sQuery = "Select PayablesNumber from Acc_T_Payables where TransactionNumber = (Select TransactionNumber from Acc_T_VoucherHeader where CreatedTransno = "& iTransNo &" ) "
        'Response.Write sQuery &"<BR>"
        objRs.Open sQuery,con
        iCtr = 1
        dTotAmtPaid = 0
        sFlag1 = False
        sFlag2 = False
		if objRs.Eof then sFlag1 = False
        if not objRs.EOF then
			iPybNo = objRs(0)

			sQuery = "Select PaidByTransactionNo,Convert(VarChar,PaidOn,103) as PaidOn,AmountPaid from Acc_T_PybleAdjustmentDetails where  Payablesnumber = "& iPybNo &" "
			objRs2.Open sQuery,con
			do while not objRs2.EOF

				iPdByTranNo = objRs2(0)
				dtPaidOn	= objRs2(1)
				iAmtPaid    = objrs2(2)
				dTotAmtPaid = dTotAmtPaid + CDbl(iAmtPaid)

				If trim(iPdByTranNo) <> "" and trim(iAmtPaid) > "0" then
					sFlag1 = True
					sQuery = "Select VoucherNumber,Convert(VarChar,VoucherDate,103) as VoucherDate,TransactionType,CreatedTransNo from Acc_T_Voucherheader where transactionnumber = "& iPdByTranNo &" "
					objRs1.Open sQuery,con
					if not objRs1.EOF then
						iVouNo     = objRs1(0)
						dtVouDate  = objRs1(1)
						sTransType = objrs1(2)
						iCrTransNo = objrs1(3)
					end if
					objRs1.Close
					If trim(sTransType) = "BAP" then
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
						<td class="ExcelDisplayCell" align="Right" width="80"><%=FormatNumber(iAmtPaid,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="center" width="60"><%=dtPaidOn%></td>
						<td class="ExcelDisplayCell" align="left" width="120"><%=sBkInsType%></td>
						<td class="ExcelDisplayCell" align="Right" width="80"><%=BkInsNo%></td>
						<td class="ExcelDisplayCell" align="center" width="80"><%=dtBkInsDate%></td>
					</tr>

					<%
					'Else


					iCtr = iCtr + 1

				End If 'If trim(iPdByTranNo) <> "" and trim(iAmtPaid) > "0" then

				objRs2.MoveNext
				loop
				objRs2.Close

				'Newly added by S.Maheswari on 09-MAR-09 to display entries from new table Acc_T_OutstandingClosingHistory
				sQuery = "Select isNull(ReasonForClosing,''),AmountAdjusted,convert(Varchar,ClosedOn,103),ClosedBy from Acc_T_OutstandingClosingHistory where PayablesNumber = "& iPybNo &" "
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
			%>
				<tr>
					<td class="ExcelSerial" align="Center" width="20"></td>
					<td class="ExcelDisplayCell" align="Center"  Colspan = "7">No payment entry available for this purchase invoice</td>
				</tr>

			<%End If
			dBalAmt = CDbl(dInvAmount) - CDbl(dTotAmtPaid)

            if Trim(iSuppInvDate)<>"" then
			    IF trim(dBalAmt) > 0 then  iNoOfDays = DateDiff("d",(mid(iSuppInvDate,4,2)&"/"&mid(iSuppInvDate,1,2)&"/"&year(iSuppInvDate)),date()) + 1
			end if  'if Trim(iSuppInvDate) then
			IF (trim(iAdjAmt) > "0" and sFlag2 <> False) or (trim(iAmtPaid) > "0" and sFlag1 <> False) then
			%>
			    <tr>
					<td class="ExcelSerial" align="Right" Colspan="3" ><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
					<td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotAmtPaid,2,,,0)%></b></td>
			    </tr>
			    <tr>
					<td class="ExcelSerial" align="Right" Colspan="3" ><p align="right"><b>Balance To Pay</b>&nbsp;&nbsp;</td>
					<td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dBalAmt,2,,,0)%></b></td>
			    </tr>
			    <tr>
					<td class="ExcelSerial" align="Right" Colspan="3" ><p align="right"><b>Outstanding in days</b>&nbsp;&nbsp;</td>
					<td class="ExcelDisplayCell" align="right"><b><%=iNoOfDays%></b></td>
			    </tr>
			<% End IF 'IF trim(iAmtPaid) > "0" then%>
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
													<input type="button" value="View Invoice" name="B3" class="ActionButtonX" onclick="ViewInvoice()" >
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
