<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalReturnEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 01, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Feb 06,2010
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId
Dim dPreCrValue,Temparr,sCallFrom ,sTotReturnQty

Dim sFinPeriod,sFromYr,sToYr,sTempYr,sFinFrm,sFinTo

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

sFinPeriod = Session("FinPeriod")
sTempYr = Split(sFinPeriod,":")
sFinFrm = Trim(sTempYr(0))
sFinTo = Trim(sTempYr(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sInvoiceNo=Request.Form("selInvoiceNo")
sBookName=Request.Form("hBookName")
Temparr = Split(sInvoiceNo,":")
sInvoiceNo = Temparr(0)
dPreCrValue = Temparr(1)
'Response.write sInvoiceNo
iInvNo = sInvoiceNo
sCallFrom = Request.QueryString("hCallFrom")
sUserid = getUserID()

sQuery = "Select VoucherAmount From Acc_T_CreatedVoucherHeader Where CreatedTransNo "&_
		 "IN (Select CreatedTransNo From Acc_T_CreatedVoucherHeader Where BankInstrumentNo = '"&sInvoiceNo&"') "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	dPreCrValue = objRs(0)
End IF
objRs.Close
Dim sRetVal
sRetVal = GetVouchXML(sInvoiceNo)
oDOM.Load server.MapPath(sRetVal)
'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")

set oNodRoot=oDOM.documentElement

for each oNodHeader in oNodRoot.childNodes
	if oNodHeader.nodeName="Header" then
		for Each oNodEntry in  oNodHeader.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.text
			end if
			if oNodEntry.nodeName="Book" then
				oNodEntry.Attributes.Item(0).nodeValue=Request.Form("selBook")
				oNodEntry.text=Request.Form("hBookName")
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue&"&nbsp;-&nbsp;"&oNodEntry.Attributes.Item(1).nodeValue
			end if
		next
	end if

	if oNodHeader.nodeName="Details" then
		set oNodDeatils=oNodHeader
	end if
	if oNodHeader.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodHeader
	end if
	if oNodHeader.nodeName="AgentDetails" then
		set oNodtemp=oNodRoot.removeChild(oNodHeader)
	end if
next

sInvValue=oNodTaxRoot.Attributes.Item(2).nodeValue
dim dInvAmount
dInvAmount = sInvValue

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="TaxData" src="<%="../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"%>"></XML>
<XML id="GJVoucher"></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesReturnCreditNoteCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="javascript">
ITMSSalesReturnCreditNoteCompat.install();
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname">
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hPreCrValue" value="<%=dPreCrValue%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="<%=Request.Form("selBook")%>">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hBookName" value="<%=Request.Form("hBookName")%>">

<input type="hidden" name="txtNarration" value="">
<%IF sCallFrom= "CR" then%>
<input type="hidden" name="hVouCode" value="07">
<%else%>
<input type="hidden" name="hVouCode" value="08">
<%end if%>
<input type="hidden" name="hVouType" value="">
<input type="hidden" name="hVouName" value="GJ">
<input type="hidden" name="optApprove" value="">
<input type="hidden" name="hInsVou" value="">
<input type="hidden" name="hSelVouDate" value="">
<input type="hidden" name="hTransNo" value="<%=iInvNo%>">
<input type="hidden" name="CallType" value="CN">
<input type="hidden" name="CrVouNo" value="">
<input type="hidden" name="SelTDSGrp" value="">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Return Credit
          Note 		</td>
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="90">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Adjustments
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
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
                            <!--tr>
                            <td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: hand" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Voucher History">
                    <font size="4" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell">
                    &nbsp;
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Unit </td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sOrgName%></span></td>
                                                    <td class="FieldCellSub" width="100"><p align="left">Date</p></td>
                                                    <td class="FieldCellSub" width="145">
                                                 <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Party
                                                      Name</td>
                                                    <td class="FieldCell" colspan="3">  <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>

                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Invoice No-Date</td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sInvoiceNo%></span></td>
                                                    <td class="FieldCellSub" width="100">Invoice Value</td>
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=sInvValue%></span></td>
                                                        </tr>

                                                        <tr>
															<td class="FieldCellSub" width="100">Credit Note Value</td>
															<td class="FieldCell" width="200">  <span class="DataOnly"><%=FormatNumber(dInvAmount,2,,,0)%></span></td>
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
                <div class="frmBody" id="frm2" style="width: 575; height:320;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
    <td class="ExcelHeaderCell" align="center" colspan="3">Invoice</td>
    <td class="ExcelHeaderCell" align="center" colspan="2">Returned</td>

        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="75">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="75">Value</td>
    <td class="ExcelHeaderCell" align="center" width="75">Discount</td>
    <td class="ExcelHeaderCell" align="center">Qty</td>
    <td class="ExcelHeaderCell" align="center">Amount</td>

        </tr>
<%
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.GetNamedItem("No").value
		sDescription=oNodEntry.Attributes.GetNamedItem("PayTo").value
		sAmount=oNodEntry.Attributes.GetNamedItem("Amount").value
		sQty=oNodEntry.Attributes.GetNamedItem("Qty").value
		sValue=oNodEntry.Attributes.GetNamedItem("ActValue").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value
		oNodEntry.Attributes.GetNamedItem("Rate").value=CDbl(oNodEntry.Attributes.GetNamedItem("Amount").value)/CDbl(sQty)
		dTotal=CDbl(dTotal)+CDbl(sAmount)

%>

	<Input type="hidden" name=hQty<%=iSno%> value="<%=sQty%>">
	<Input type="hidden" name=hInvVal<%=iSno%> value="<%=sValue%>">

    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sQty,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sValue,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sDiscount,2,,,0)%></td>
    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtqty<%=iSno%>1" onBlur="setQty(this,'<%=iSno%>')" value="<%=FormatNumber(sQty,3,,,0)%>" class="Formelem" size="13"></td>

    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iSno%>" onBlur="setTotal(this,'<%=iSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13"></td>

        </tr>
<%
	sTotReturnQty = cdbl(sTotReturnQty) + cdbl(sQty)
	next
%>


        <tr>
    <td align="center" colspan="2"></td>
    <input type="hidden" name="hTotReturnQty" value="<%=sTotReturnQty%>">
    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
     <td align="right" colspan="2"></td>
     <td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotal" value="<%=FormatNumber(dTotal,2,,,0)%>" class="Formelem" size="13"></td>
        </tr>



<%
	Dim sCheckExp,CheckNode
'	dim dInvAmount
'	dInvAmount=dTotal
	Dim iCheck
	For Each oNodEntry in oNodTaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.GetNamedItem("CatCode").value
		sTaxCode=oNodEntry.Attributes.GetNamedItem("TaxCode").value
		sTaxMode=oNodEntry.Attributes.GetNamedItem("TaxMode").value
		sFormula=oNodEntry.Attributes.GetNamedItem("TaxFormula").value
		dTaxValue=oNodEntry.Attributes.GetNamedItem("TaxValue").value


		If sCatCode = "0" and sTaxCode = "0" and sTaxMode = "0" Then
			iRndOff = 0
		Else
			sCheckExp = "//TaxDetails/Tax[@CatCode="&sCatCode&" and @TaxCode="&sTaxCode&" and @RoundOff]"
			Set CheckNode = oNodRoot.selectNodes(sCheckExp)
			IF CheckNode.length <> 0 Then
				iRndOff = oNodEntry.Attributes.GetNamedItem("RoundOff").value
			Else
				iRndOff = 0
			End IF
		End If
		sTaxName=oNodEntry.Text
		If iRndOff = 1 Then
			dTax = FormatNumber(Round(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,0),2,,,0)
		Else
			dTax = FormatNumber(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,2,,,0)
		End If
%>
			<tr>
				<td align="center" colspan="2"></td>
				<td class="ExcelSerial" align="center" colspan="3"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="P" then %>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="Formelem">&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  onBlur="setTaxAmount('<%=sCatCode%>','<%=sTaxCode%>',this)"size="11" class="Formelem"></td>
				    </tr>
			<%
	next
oDOM.save server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")

%>

		<input type="hidden" name="hTotinvVal" value="<%=dInvAmount%>">
        <tr>
        <td align="center" colspan="2"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="Formelem">
    </td>
        </tr>
        <tr>
        </tr>
        <tr>
			<td align="left" class="FieldCell" colspan="2" valign="Top">
				Approval
			</td>
			<td align="center" class="FieldCellSub" colspan="2">
			<Input type="radio" name="optApprove" checked value="Y" onClick="EnbApp(this)"> Yes &nbsp;&nbsp;&nbsp;
			<Input type="radio" name="optApprove" value="N" onClick="EnbApp(this)"> No &nbsp;&nbsp;&nbsp;
			</td>
        </tr>
        <tr>
			<td align="left" class="FieldCell" colspan="2" valign="Top">
				Immediate Approver
			</td>
			<td align="center" class="FieldCellSub" colspan="2">
			<select size="1" name="selUserId" class="FormElem">
              <option value="I">Immediate Approver</option>
                <%=populateEmployeeWithVal(sUserId)%>
                    </select>
			</td>
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
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                                <input type="button" value="Next" name="B2" class="ActionButton" onClick="SaveXML()" >
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" onClick="Cancel('VOUCNBOOKSELECTION.ASP')" >
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
