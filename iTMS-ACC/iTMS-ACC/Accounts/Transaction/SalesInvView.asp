
<% option explicit	%>
<%
	'Program Name				:	SalesInvView.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	September 30, 2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodTemp,oNodDeatils,oNodTaxRoot,oNodEntry,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo,sInvDet,sInvDate
dim sDiscPer,dBasicTotal,dDisTotal,sBookNo,sRemarks,sTerms
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue
dim iTransNo,sOrgName,sBookName,sParType,sParSubType,sParCode,sRefPartyCode
Dim iAccHead,iItemAccHead,sAccHeadName,iAccHeadCheck,sAccType,dRatePer
Dim dConRate,sConDate,sInvCurr,sOpCurr,sCurrName,sRetVal
Dim sSubTyName,sSalTypeName,sParBank,sDispName,sPayName,sParLoc,sInvTypeName
Dim iSalTrNo,sBasPric,sLoadName,sModeDesc,sPayTerms,sTransName,dItem,dClass,dOldInvVal


' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

iTransNo=Request.QueryString("TransNo")
'Response.Write "<p> iTransNo = " & iTransNo


'oDOM.load  server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")	
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)
'Response.Write "<p><font color=red>XML="&sRetVal
'Response.Write "<p>iTransNo="&iTransNo

set oNodRoot=oDOM.documentElement

for each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="Header" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="Book" then
				sBookNo=oNodEntry.Attributes.Item(0).nodeValue
				sBookName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="Party" then
				sParType = oNodEntry.Attributes.Item(0).nodeValue
				sParSubType = oNodEntry.Attributes.Item(1).nodeValue
				sParCode = oNodEntry.Attributes.Item(3).nodeValue
				sSubTyName = oNodEntry.Attributes.Item(2).nodeValue
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sInvDet = oNodEntry.Attributes.Item(0).nodeValue
				sInvDate = oNodEntry.Attributes.Item(1).nodeValue
				sRefernceNo = oNodEntry.Attributes.Item(2).nodeValue
			end if		
			if oNodEntry.nodeName="SalesType" then
				sSalTypeName = oNodEntry.Text
			end if		
			if oNodEntry.nodeName="Currency" then
				sInvCurr = oNodEntry.Attributes.Item(0).nodeValue
				sOpCurr = oNodEntry.Attributes.Item(1).nodeValue
			End IF
		next
	end if
	
	if oNodTemp.nodeName="Details" then 
		set oNodDeatils=oNodTemp
	end if

	if oNodTemp.nodeName="TaxDetails" then 
		set oNodTaxRoot=oNodTemp
	end if
next	
dInvAmount=oNodTaxRoot.Attributes.Item(0).nodeValue

if sInvCurr = "" Then sInvCurr = 1	'For Test

sQuery = "Select ConversionRate,Convert(char,ConversionAsOn,103),CurrencyShortName From "&_
		 "Ms_CurrencyMaster Where CurrencyCode = "&sInvCurr&" "
'Response.Write "<p>"&squery
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	dConRate = objRs(0)
	sConDate = objRs(1)
	sCurrName = objRs(2)
End IF
objRs.Close

sQuery = "Select OrgnPartyCode From App_M_PartyMaster Where PartyCode = "&sParCode&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sRefPartyCode = objRs(0)
End IF
objRs.Close


sQuery = "Select OtherApplnTransNo From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	iSalTrNo = objRs(0)
End IF
objRs.Close

if trim(iSalTrNo) <> "" then
	sQuery = "Select TermsDisplay,Remarks from Sal_T_InvoiceHeader where SaleTransactionNo = " & iSalTrNo
	'Response.Write "<p> sQuery = " & sQuery
	Objrs.Open sQuery,Con
	If Not Objrs.Eof Then
		sTerms = objRs(0)
		sRemarks  =objRs(1)
	End IF
	Objrs.close	
end if 


sQuery = "Select Distinct BasisOfPricing,DestinationPlaceName,LoadingPlaceName,DespatchModeDesc, "&_
		 "PaymentMode,PaymentTermsDesc,TransporterName,InvoiceValue From Vw_Sal_InvHead Where SaleTransactionNo = "&iSalTrNo&" "
'Response.Write "<p><font color=red>"&squery		 
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With
Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	sBasPric = objRs(0)
	sDispName = objRs(1)
	sLoadName = objRs(2)
	sModeDesc = objRs(3)
	sPayName = objRs(4)
	sPayTerms = objRs(5)
	sTransName = objRs(6)
	dOldInvVal = objRs(7)
End IF
objRs.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Sales Invoice</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="AppOtherSALUpdate.asp">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="05">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Invoice&nbsp;
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
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
                            <td class="FieldCellsub">Unit </td>
                            <td width="160" class="FieldCellSub">
                                                            <span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                            </td>
                            <td class="FieldCellSub">
                            Invoice No. - Date</td>
                            <td class="FieldCellSub" >
                            	<span class="DataOnly"><%=sInvDet%>&nbsp;</span>&nbsp;<span class="DataOnly"><%=sInvDate%>&nbsp;</span>
                            </td>
                                </tr>
                                
                                <tr>
                            <td class="FieldCellsub">Party Name </td>
                            <td width="500" class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sPartyName%>&nbsp;</span> 
                            </td>
                                </tr>
                                
                                <tr>
                            <td class="FieldCellsub">Party Sub Type</td>
                            <td width="500" class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sSubTyName%>&nbsp;</span> 
                            </td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Type of Sale</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sSalTypeName%>&nbsp;</span> 
									</td>
									<td class="FieldCellSub">Invoice Type
									</td>
									<td class="FieldCellSub"><span class="DataOnly">Invoice Type</span> 
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Payment Terms</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sPayTerms%>&nbsp;</span> 
									</td>
									<td class="FieldCellSub">Mode of Dispatch
									</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sModeDesc%></span> 
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Basis of Pricing</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sBasPric%>&nbsp;</span> 
									</td>
									<td class="FieldCellSub">Mode of Payment
									</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sPayName%></span> 
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Transporter</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sTransName%>&nbsp;</span> 
									</td>
									<td class="FieldCellSub">Destination Port
									</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sDispName%></span> 
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Loading Port</td>
									<td class="FieldCellSub"><span class="DataOnly"><%=sLoadName%>&nbsp;</span> 
									</td>
									<td class="FieldCellSub">Party Bank
									</td>
									<td class="FieldCellSub"><span class="DataOnly">Invoice Type</span> 
									</td>
                                </tr>
                                
                            
                           
                           <tr>
								<td class="FieldCellsub">Invoice Currency</td>
								<td class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sCurrName%>&nbsp;</span> 
								</td>
                           </tr>
                           
                           <% IF CStr(sOpCurr) <> CStr(sInvCurr) Then %>
                           <tr>
                            <td class="FieldCellsub">Conversion Rate</td>
                            <td class="FieldCellSub"><span class="DataOnly"><%=FormatNumber(dConRate)%>&nbsp;</span> 
                            </td>
                            <td class="FieldCellSub">Conversion As On </td>
                            <td class="FieldCellSub"><span class="DataOnly"><%=sConDate%>&nbsp;</span> 
                            </td>
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
                <div class="frmBody" id="frm2" style="width: 750; height:235;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
	<td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Quantity</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Rate</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Rate<br>
    Per</td>
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

	Dim dOldAmt,dOldVal,dOldDis,dOldNetVal,dOldBasTotal,dOldTotDisVal,dOldNetTotVal
	sQuery = "Select Sum(BasicAmount), Sum(DiscountAmount), Sum(BasicAmount)-Sum(DiscountAmount) "&_
			 "From Sal_T_InvoiceDetails Where SaleTransactionNo = "&iSalTrNo&" "
	With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not objRs.EOF Then
			dOldBasTotal = objRs(0)
			dOldTotDisVal = objRs(1)
			dOldNetTotVal = objRs(2)
		End IF
		objRs.Close	
		
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.Item(0).nodeValue
		sDescription=oNodEntry.Attributes.Item(1).nodeValue
		sAmount=oNodEntry.Attributes.Item(2).nodeValue
		sRate=oNodEntry.Attributes.Item(6).nodeValue
		sQty=oNodEntry.Attributes.Item(3).nodeValue &"&nbsp;"&oNodEntry.Attributes.Item(5).nodeValue
		sValue=oNodEntry.Attributes.Item(7).nodeValue
		sDiscPer=oNodEntry.Attributes.Item(8).nodeValue
		sDiscount=oNodEntry.Attributes.Item(9).nodeValue
		dRatePer = oNodEntry.Attributes.Item(16).nodeValue
		dItem = oNodEntry.Attributes.Item(10).nodeValue
		dClass = oNodEntry.Attributes.Item(11).nodeValue
	
		sQuery = "Select BasicAmount,DiscountAmount, "&_
				 "BasicAmount - DiscountAmount NetAmount From Sal_T_InvoiceDetails "&_
				 "Where SaleTransactionNo = "&iSalTrNo&" and ItemCode = "&dItem&" and ClassificationCode = "&dClass&" "
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not objRs.EOF Then
			dOldVal = objRs(0)
			dOldDis = objRs(1)
			dOldNetVal = objRs(2)
		End IF
		objRs.Close	
		
		dTotal=CDbl(dTotal)+CDbl(sAmount)
		dBasicTotal=CDbl(dBasicTotal)+CDbl(sValue)
		dDisTotal=CDbl(dDisTotal)+CDbl(sDiscount)
	
%>
    <tr>
		<td class="ExcelSerial" align="center" rowspan="2"><%=isno%></td>
		<td class="ExcelDisplayCell" rowspan="2"><%=sDescription%> </td>
		<td class="ExcelDisplayCell" align="Left" width="60" rowspan="2"><%=sQty%></td>
		<td class="ExcelDisplayCell" align="Right" width="60" rowspan="2"><%=FormatNumber(sRate,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right" width="60" rowspan="2"><%=FormatNumber(dRatePer,2,,,0)%></td>
		
		<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(dOldVal,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right" ><%=FormatNumber(sDiscPer ,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(dOldDis,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right"><%=FormatNumber(dOldNetVal,2,,,0)%></td>
   </tr>
   
   <tr>
		<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sValue,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right" ><%=FormatNumber(sDiscPer ,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber( sDiscount,2,,,0)%></td>
		<td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sAmount,2,,,0)%></td>
   </tr>
<%		
	next
'end if	
'next %>        
    
        <tr>
			<!--<td align="center" colspan="3" rowspan="2"></td>-->
			<td class="ExcelSerial" align="center" colspan="5" rowspan="2"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
			<td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dOldBasTotal,2,,,0)%></b></td>
			<td class="ExcelDisplayCell" align="center" width="25">    </td>
			<td class="ExcelDisplayCell" align="right" width="60" ><b><%=FormatNumber(dOldTotDisVal,2,,,0)%></b></td>
			<td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dOldNetTotVal,2,,,0)%></b></td>
       </tr>
       <tr>
			<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dBasicTotal,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="center" width="25">    </td>
			<td class="ExcelDisplayCell" align="right" width="60" ><%=FormatNumber(dDisTotal,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTotal,2,,,0)%></td>
       </tr>
       
        
        <input type="Hidden" name="hBasicValue" value="<%=dBasicTotal%>">
        <input type="Hidden" name="hDisValue" value="<%=dDisTotal%>">
        <input type="Hidden" name="hAmount" value="<%=dTotal%>">
<%
dim dInvAmount
	
	For Each oNodEntry in oNodTaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue 
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue 
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue 
		sFormula=oNodEntry.Attributes.Item(3).nodeValue 
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue 
		iAccHead = oNodEntry.Attributes.Item(6).nodeValue 
		sTaxName=oNodEntry.Text
		
		'sQuery = "Select TaxAmount From Sal_T_InvoiceTaxDetails Where SaleTransactionNo = "&iSalTrNo&" "&_
		'		 "and TaxCode = "&sTaxCode&" and TaxCategoryCode = "&sCatCode&" "
		'With objRs
		'	.CursorLocation = 3
		'	.CursorType = 3
		'	.Source = sQuery
		'	.ActiveConnection = Con
		'	.Open
		'End With
		'Set objRs.ActiveConnection = Nothing
		'IF Not objRs.EOF Then
		'	dTax = objRs(0)
		'End IF
		'objRs.Close	
		
		dTax=oNodEntry.Attributes.Item(5).nodeValue
			%>
			<tr>
				<!--<td align="center" colspan="3"></td>-->
				<td class="ExcelSerial" align="Right" colspan="7"><%=sTaxName%></td>
				<%if sTaxMode="P" then %>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTaxValue,2,,,0)%>&nbsp;%</td>
				<%else%>
					<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTax,2,,,0)%></td>
				    </tr>
			<%
	next		
%>        <%If dOldInvVal = "" Then dOldInvVal = cdbl("0")
            If dInvAmount = "" Then dInvAmount = cdbl("0")%>
        
        
        <tr>
        <!--<td align="center" colspan="3"></td>-->
    <td class="ExcelSerial" align="right" colspan="8"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <%=FormatNumber(dInvAmount,2,,,0)%><%'=FormatNumber(dOldInvVal,2,,,0)%> </td>
        </tr>
            </table>
                </div>
                </td>
                <td></td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
                                &nbsp;
								</td>
								<td valign="top" class="FieldCell" >
                                    <table cellpadding="0" cellspacing="0" >
                                    <tr>
                            <td class="FieldCell" valign="top">Terms</td>
                            <td class="FieldCellSub" >
                                <span class="DataOnly"><%=sTerms%></span>
                            </td>
                            </tr>
                            <tr>
                            <td class="FieldCell" valign="top">Remarks </td>
                            <td class="FieldCellSub" >
                                <span class="DataOnly" style="color:red"><%=sRemarks%></span>
                            </td>
                            </tr>
                                    <tr>
                            <td class="FieldCell" valign="top">Amount </td>
                            <td class="FieldCellSub" >
                                <span class="DataOnly"><%=AmountWords(dInvAmount)%>
                                <%'=AmountWords(dOldInvVal)%>&nbsp;</span>
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
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															
                                                                <input type="button" value="OK" name="btnAction2" onclick="window.close()" class="ActionButton">
                                                                
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
