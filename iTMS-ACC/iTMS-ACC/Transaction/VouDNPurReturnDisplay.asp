<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNPurReturnDisplay.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 18, 2003
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
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
Dim iTransNo,dHdVal

dim iSno,sDescription,sAmount,sQty,dTotal
dim sTaxName,dTax,sTaxMode,dTaxValue,sRetVal

dim sOrgName,sBookCode,sBookName,sPartyName,sInvoiceNo

dim sVouNo,sVouDate

Dim sQuery,sVouCode,iCrTransNo

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sInvoiceNo=Request("TransNo")
'Response.Write sInvoiceNo

sQuery = "Select BookCode,VoucherAmount From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sInvoiceNo&" "
'Response.Write sQuery
objRs.Open sQuery,con
IF Not objRs.EOF Then
	sVouCode = objRs(0)
	dHdVal = objRs(1)
End IF
objRs.Close

'Response.Write sBookCode

iCrTransNo = sInvoiceNo

'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")
sRetVal = GetVouchXML(sInvoiceNo)
'Response.Write sInvoiceNo
'Response.End
oDOM.Load server.MapPath(sRetVal)

set oNodRoot=oDOM.documentElement

for each oNodHeader in oNodRoot.childNodes
	if oNodHeader.nodeName="Header" then
		for Each oNodEntry in  oNodHeader.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgName=oNodEntry.text
			end if
			if oNodEntry.nodeName="Book" then
				sBookCode=oNodEntry.Attributes.Item(0).nodeValue
				sBookName=oNodEntry.text
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="PurInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue &"&nbsp; -&nbsp; "&oNodEntry.Attributes.Item(1).nodeValue
			end if
		next
	end if

	if oNodHeader.nodeName="Details" then
		set oNodDeatils=oNodHeader
	end if
	if oNodHeader.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodHeader
	end if
next

dTotal = oNodTaxRoot.Attributes.Item(0).nodeValue

sVouNo=oNodRoot.Attributes.Item(1).nodeValue
sVouDate=oNodDeatils.Attributes.Item(3).nodeValue
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<Script Language="VBScript">
Function ChkSubmit()
if document.formname.hVouCode.value="07" then
	document.formname.action="CreditVouchers.asp"
	document.formname.submit
else
	document.formname.action="DebitVouchers.asp"
	document.formname.submit
end if
End Function

Function PrintVou()
	PrintWindow("PRNCNNoteForRet.asp?iTransNo=" & document.formname.htransNo.value)
End Function

Function PrintVouDN()
	PrintWindow("PRNDNNoteForRet.asp?iTransNo=" & document.formname.htransNo.value)
End Function
</Script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%IF CStr(sVouCode) = "07" Then%>
	<form method="POST" name="formname" action="CreditVouchers.asp">
<%Else%>
	<form method="POST" name="formname" action="DebitVouchers.asp">
<%End IF %>
<Input type="hidden" name="hTransNo" Value="<%=iCrTransNo%>">
<Input type="hidden" name="hVouCode" Value="<%=sVouCode %>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%IF CStr(sVouCode) = "07" Then
			Response.Write "Credit Note "
		  Else
			Response.Write "Debit Note "
		  End IF
		%>
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
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="90">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<td align="center">Adjustments</td>
									</tr>
								  </table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								  	<tr>
								  		<tr><td align="center">Voucher</td>
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
                            <td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                        <tr>
							<td align="center" colspan="3" class="MiddlePack">
								<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
						</tr>
						<tr>
							<td align="center" width="5" class="ClearPixel">
							</td>
							<td valign="top" width="100%">
                                <table border="0" cellspacing="0" cellpadding="0">
	                                <tr>
										<td class="FieldCellSub" width="80">Book No</td>
										<td class="FieldCell" width="200"><span class="DataOnly"><%=sBookCode%></span>  </td>
										<td class="FieldCellSub" width="100">Book Name</td>
										<td class="FieldCellSub" width="145"><span class="DataOnly"><%=sBookName%></span> </td>
									</tr>
                                    <tr>
										<td class="FieldCellSub" width="80">Voucher No</td>
										<td class="FieldCell" width="200"><span class="DataOnly"><b><%=sVouNo%></b></span>  </td>
										<td class="FieldCellSub" width="100">Voucher Date</td>
										<td class="FieldCellSub" width="145"><span class="DataOnly"><%=sVouDate%></span> </td>
									</tr>
									<tr>
										<td class="FieldCellSub" width="70">Party                                                      Name</td>
										<td class="FieldCell" width="200">  <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
										<td class="FieldCellSub" width="100"><p align="left">Invoice No-Date</p></td>
										<td class="FieldCellSub" width="145"><span class="DataOnly"><%=sInvoiceNo%></span> </td>
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
                <div class="frmBody" id="frm2" style="width: 500; height:200;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
    <td class="ExcelHeaderCell" align="center">Item Description</td>
    <td class="ExcelHeaderCell" align="center" width="75">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="75">Amount</td>
        </tr>
<%
iSno=0
dTotal = 0
Dim iTaxCode
	For Each oNodEntry in oNodDeatils.childNodes

		sDescription = oNodEntry.Attributes.Item(1).nodeValue
		sAmount = oNodEntry.Attributes.Item(2).nodeValue
		sQty = oNodEntry.Attributes.Item(3).nodeValue


		if CDbl(sAmount)>0 then
			dTotal=CDbl(dTotal)+CDbl(sAmount)
			iSno=CInt(iSno)+1
%>
    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sQty,3,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sAmount,2,,,0)%></td>
        </tr>
<%
		end if
	next

	'For Each oNodEntry in oNodTaxRoot.childNodes
	'	sAccCode=oNodEntry.Attributes.Item(6).nodeValue
	'	dTax=oNodEntry.Attributes.Item(5).nodeValue
	'	if cint(sAccCode)=0 then
	'		dTotal=dTotal+CDbl(dTax)
	'	End IF
	'Next


%>

        <tr>
    <td align="center"></td>
    <td class="ExcelSerial" align="center" colspan="2"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
     <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTotal,2,,,0)%></td>
        </tr>



<%
dim dInvAmount,sAccCode
	dInvAmount=dTotal
	For Each oNodEntry in oNodTaxRoot.childNodes
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue
		sTaxName=oNodEntry.Text
		dTax=oNodEntry.Attributes.Item(5).nodeValue
		sAccCode=oNodEntry.Attributes.Item(6).nodeValue
		'if CDbl(dTax)>0 and cint(sAccCode)>0 then

		'if cint(sAccCode)>0 then
			dInvAmount=dInvAmount+CDbl(dTax)
			%>
			<tr>
				<td align="center"></td>
				<td class="ExcelSerial" align="center"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="P" then %>
					<td class="ExcelDisplayCell" align="right"><%=dTaxValue%>&nbsp;%</td>
				<%else%>
					<td class="ExcelDisplayCell" align="right"></td>
				<%end if%>
				<td class="ExcelDisplayCell" align="right"><%=dTax%></td>
		    </tr>
			<%
		'end if
	next
%>
        <tr>
        <td align="center"></td>
    <td class="ExcelSerial" align="right" colspan="2"><b>Debit Note Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dHdVal,2,,,0)%></td>
        </tr>
            </table>
                </div>
                </td>
                <td></td>
                            </tr>
                                                        <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                                        </tr>
                                                        <tr>
															<td align="center" width="5" class="ClearPixel">
															&nbsp;
															</td>
															<td valign="top" class="FieldCell" height="20">
																<table cellpadding="0" cellspacing="0" width="100%">
															<tr>
														<td class="FieldCell" width="130" valign="top">Amount </td>
														<td>
																						   <span class="DataOnly"><%=AmountWords(abs(dInvAmount))%></span>
														</td>
																</table>

															</td>
															<td align="center" class="ClearPixel" width="5">
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
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onclick="ChkSubmit()">
                                                                <%IF CStr(sVouCode) = "07" Then %>
																	<input type="button" value="Print" name="B8" onClick="PrintVou()" class="ActionButton">
																<%Else%>
																	<input type="button" value="Print" name="B8" onClick="PrintVouDN()" class="ActionButton">
																<%End IF%>


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