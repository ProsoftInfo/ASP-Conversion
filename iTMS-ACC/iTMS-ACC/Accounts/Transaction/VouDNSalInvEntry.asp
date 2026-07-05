<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNOSalInvEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	October 20, 2004
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff,sFromPur
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId,sTemp
Dim sFromApp,sCallFrom,iSalInvoiceNo
Dim sFinPeriod,sArrFin,sFromDate,sTodate

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sInvoiceNo=Request.Form("selInvoiceNo")
sBookName=Request.Form("hBookName")
sFromApp = Request.Form("hFromApp")
sCallFrom = Request("CallFrom")
Response.Write "<p>"& sCallFrom

sFinPeriod = Session("FinPeriod")
sArrFin = Split(sFinPeriod,":")
sFromDate = "01/04/"&sArrFin(0)
sTodate = "31/03/"&sArrFin(1)

sTemp = Split(sInvoiceNo,":")
sInvoiceNo = sTemp(0)

iInvNo = sInvoiceNo

sQuery = "Select FromApplication,OtherApplnTransNo From Acc_T_CreatedVoucherHeader "&_
		 "Where CreatedTransNo = "&iInvNo&" and FromApplication is Not NULL "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sFromPur = "Y"
	iSalInvoiceNo = objRs(1)
Else
	sFromPur = "N"
	iSalInvoiceNo = 0
End IF
objRs.Close		 
		 

sUserid = getUserID()

Dim sRetVal
'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")
sRetVal = GetVouchXML(sInvoiceNo)
oDOM.Load server.MapPath(sRetVal)

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
			if oNodEntry.nodeName="PurInvoice" then
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
dInvAmount=oNodTaxRoot.Attributes.Item(0).nodeValue

dim dInvAmount
'dInvAmount = sInvValue

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="TaxData" src="<%="../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"%>"></XML>
<xml id="GLData"><Root></Root></xml>
<xml id="GJVoucher"></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesReturnCreditNoteEntryCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="javascript">
ITMSSalesReturnCreditNoteEntryCompat.install();
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname">
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hAccType" value="C">
<Input type="hidden" name="hCallType" value="OINV">
<Input type="hidden" name="hNoteType" value="D">
<Input type="hidden" name="hFromPur" value="<%=sFromPur%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hVouCode" value="06">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="<%=Request.Form("selBook")%>">
<input type="hidden" name="hCrAccHead" value="0">
<input type="hidden" name="hCallFromDebit" value= "<%=sCallFrom%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hVouName" value="GJ">
<input type="hidden" name="hInvoiceNo" value="<%=iSalInvoiceNo%>">
<input type="hidden" name="hFromDate" value="<%=sFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sToDate%>" />

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"> 
		
				Debit Note Sale Invoice
		
          		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
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
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
										<tr>
											<td align="center">Voucher Details
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
                            </tr-->
                            <!--tr>
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
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="100%">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <!--<td class="FieldCellSub" width="100">Unit </td>
                                                    <td class="FieldCell">  <span class="DataOnly"><%=sOrgName%></span></td>-->
                                                    <td class="FieldCellSub" width="100">Party Name</td>
                                                    <td class="FieldCell"> <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
                                                    <td class="FieldCellSub" width="75"><p align="left">Date</p></td>
                                                    <td class="FieldCellSub" width="145">
                                                 <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
                                                        </tr>
                                                        <!--<tr>
                                                    <td class="FieldCellSub" width="100">Party Name</td>
                                                    <td class="FieldCell" colspan="3">  <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
                                                        </tr>-->
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Invoice No-Date</td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sInvoiceNo%></span></td>
                                                    <td class="FieldCellSub" width="100">Invoice Value</td>
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=FormatNumber(sInvValue,2,,,0)%> </span></td>
                                                        </tr>
                                                        
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Dr Note Against</td>
                                                    <td class="FieldCell" width="200">  
                                                    <Select name="SelCrAgain" class="FormElem" onChange="SetRetVal(this,'1')">
                                                    <Option Value="0">Select</Option>
                                                    <Option Value="Q">Quantity</Option>
                                                    <Option Value="R">Rate</Option>
                                                    <Option Value="D">Discount</Option>
                                                    <Option Value="A">Quality</Option>
                                                    
                                                    </Select>
                                                    </td>
                                                    
                                                        </tr>
                                                        <tr>
														<td class="FieldCellSub" width="200">Select Account Head</td>
														<td class="FieldCell" width="200">  
														<Select name="SelAccountHd" class="FormElem" onChange="AccHead(this)">
															<Option Value="0">ITEM ACCOUNT HEAD</Option>
															<Option Value="G">GL ACCOUNT HEAD</Option>
														</Select>
														</td>
														 <td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spAccHead"></span>
														 </td>
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
                <div class="frmBody" id="frm2" style="width: 775; height:245;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
    <td class="ExcelHeaderCell" align="center" colspan="4">Invoice</td>
    <td class="ExcelHeaderCell" align="center" colspan="4">Returned</td>
     <!--td class="ExcelHeaderCell" align="center" colspan="2">Invoice Value</td-->

        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="55">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="55">Value</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="55">Value</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
        </tr>
<%
	Dim sRatePer
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.GetNamedItem("No").value
		sDescription=oNodEntry.Attributes.GetNamedItem("PayTo").value
		sAmount=oNodEntry.Attributes.GetNamedItem("Amount").value
		sQty=oNodEntry.Attributes.GetNamedItem("Qty").value
		sValue=oNodEntry.Attributes.GetNamedItem("ActValue").value
		'oNodEntry.Attributes.GetNamedItem("Rate").value=CDbl(oNodEntry.Attributes.GetNamedItem("Amount").value)/CDbl(sQty)
		sRate = oNodEntry.Attributes.GetNamedItem("Rate").value
		'sRatePer = oNodEntry.Attributes.GetNamedItem("RatePer").value
		sRatePer = 1
		sDiscPer =oNodEntry.Attributes.GetNamedItem("DisPer").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value
		
		dTotal=CDbl(dTotal)+CDbl(sAmount)

%>

    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldQty<%=isno%>"><%=FormatNumber(sQty,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldRate<%=isno%>"><%=FormatNumber(sRate,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldDis<%=isno%>"><%=FormatNumber(sDiscount,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13"></td>
    
    <td class="ExcelDisplayCell" align="Right" id="tQty<%=isno%>"><input type="text" style="text-align: Right" NAME="txtqty<%=isno%>" onBlur="setQty(this,'<%=iSno%>','Q')" value="<%=FormatNumber(sQty,3,,,0)%>" class="FormelemRead" size="13"></td>
    <td class="ExcelDisplayCell" align="Right" id="tRate<%=isno%>"><input type="text" style="text-align: Right" NAME="txtRate<%=isno%>" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')" value="<%=FormatNumber(sRate,2,,,0)%>" class="FormelemRead" size="13"></td>
    <input type="hidden" name="hRatePer" value="<%=sRatePer%>">
    </td>
    <td class="ExcelDisplayCell" align="Right" id="tDis<%=isno%>">
    <input type="text" style="text-align: Right" NAME="txtDis<%=isno%>" value="<%=FormatNumber(sDiscount,2,,,0)%>" class="FormelemRead" size="13" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')">
    <input type="hidden" name="hDisPer<%=iSNo%>" value="<%=sDiscPer%>">
    </td>
    

    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iSno%>" onBlur="setTotal(this,'<%=iSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13"></td>
	<!--td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13"></td-->
        </tr>
<%
	next
%>

        <tr>
        <Input type="hidden" name="hRowVal" value="<%=isno%>">
    <!--<td align="center" ></td>-->
    <td class="ExcelSerial" align="center" colspan="3"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <!--<td align="center" ></td>-->
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
    <td  class="ExcelDisplayCell" align="right" ></td>
    <td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotalInv" value="<%=FormatNumber(dTotal,2,,,0)%>" class="FormelemRead" size="13"></td>
     <td class="ExcelDisplayCell" align="right" colspan="3"></td>
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
				<!--<td align="center" colspan="1"></td>-->
				<td class="ExcelSerial" align="right" colspan="4"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="P" then %>
				<td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="FormelemRead" readonly>&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue" value="<%=dTax%>"  size="11" class="FormelemRead"></td>
				<td class="ExcelDisplayCell"align="center" colspan="3"></td>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  onBlur="ReTotalCr()" size="11" class="Formelem"></td>
				
				    </tr>
				    
			<%
	next
	
	
oDOM.save server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")

%>

		

        <tr>
        <!--<td align="center" colspan="1"></td>-->
    <td class="ExcelSerial" align="right" colspan="5"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <input type="text" style="text-align: Right" NAME="txtFInvValue"  size="13" value="<%=FormatNumber(sInvValue,2,,,0)%>" class="FormelemRead"></td>
    <td class="ExcelDisplayCell" align="right" colspan="3"></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(sInvValue,2,,,0)%>" class="Formelem">
    
    </td>
        </tr>
        
        <tr>
        <!--<td align="center" colspan="1"></td>-->
        <td class="ExcelSerial" align="right" colspan="9"><b>Debit Note Value&nbsp; </b></td>
    <!--td class="ExcelSerial" align="right" colspan="3"><b>Credit Note Value&nbsp; </b></td-->
    <td class="ExcelDisplayCell" align="right" id="tDrVal"> <input type="text" style="text-align: Right" NAME="txtCrNoteValue"  size="13" value="0.00" class="FormelemRead" readonly>
    
    </td>
        </tr>
        <tr>
        </tr>
        
        
            </table>
            
                </div>
                </td>
                <td></td>
                            </tr>
                            <tr>
                            <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Approval &nbsp;&nbsp;&nbsp;
			
			<Input type="radio" name="optApprove" checked value="Y" onClick="EnbApp(this)"> Yes &nbsp;&nbsp;&nbsp;
			<Input type="radio" name="optApprove" value="N" onClick="EnbApp(this)"> No &nbsp;&nbsp;&nbsp;
			</td>
        </tr>
        <tr>
			 <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Immediate Approver &nbsp;&nbsp;&nbsp;
			
			<select size="1" name="selUserId" class="FormElem">
              <option value="I">Immediate Approver</option>
                <%=populateEmployeeWithVal(sUserId)%>
                    </select>
			</td>
        </tr>
         <tr>
			 <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Narration &nbsp;&nbsp;&nbsp;
			
			<Textarea name="txtNarration" class="FormElem" cols="40" rows="4"></Textarea>
			
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
                                                                <input type="button" value="Next" name="B2" class="ActionButton" onClick="SaveXML()" >
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" onClick="Cancel('DEBITNOTETOCREATE.ASP')" >
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
