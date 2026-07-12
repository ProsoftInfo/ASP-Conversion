<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppOtherSALView.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodTemp,oNodDeatils,oNodTaxRoot,oNodEntry,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo,sInvDet,sRefSaleTransNo
dim sDiscPer,dBasicTotal,dDisTotal,sBookNo
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue
dim iTransNo,sOrgName,sBookName,sParType,sParSubType,sParCode,sRefPartyCode
Dim iAccHead,iItemAccHead,sAccHeadName,iAccHeadCheck,sAccType,dRatePer
Dim dConRate,sConDate,sInvCurr,sOpCurr,sCurrName,sSelSalNo,sSelSalDt,sAccBookRel
Dim sRetVal,sPara,sButVal,sInvDate,sRemarks,sTerms
dim dInvAmount


' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sAccBookRel = "T" 'Book and Accounthead relation has made Yes only after the Invoice Type
				  'has been mapped for any Book voucher can be accounted

iTransNo=Request("TransNo")
sPara = Request("sPara")
'Response.Write "iTransNo="&iTransNo
IF trim(sPara) = "Edt" then
	sButVal = "Edit"
ElseIF trim(sPara) = "App" then
	sButVal = "Approve"
ElseIF trim(sPara) = "Acc" then
	sButVal = "Account"
Else
	sButVal = "Next"
End IF

if trim(iTransNo) = "" then
	iTransNo=Request.QueryString("TransNo")
end if

'oDOM.load  server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
'Response.write "<p><font color=red>" & sRetVal
'Response.Write "<p>iTransNo="&iTransNo

'Response.Write Server.MapPath(sRetVal)

oDOM.Load server.MapPath(sRetVal)
'Response.Write "<p>Query = Select TaxAmount,isNull(AccountHead,0) from Acc_T_CreatedVoucherTaxDet where isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0  and CreatedTransNo = 100"

sQuery = "Select CreatedVoucherNo,Convert(Char,VoucherDate,103) From Acc_T_CreatedVoucherHeader "&_
		 "Where CreatedTransNo = "&iTransNo

Objrs.Open sQuery,Con
If Not Objrs.Eof Then
	sSelSalNo = Objrs(0)
	sSelSalDt = Trim(Objrs(1))
End IF
Objrs.close


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
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				oNodEntry.Attributes.Item(0).nodeValue = sSelSalNo
				oNodEntry.Attributes.Item(1).nodeValue = sSelSalDt
				oNodEntry.Attributes.Item(2).nodeValue = sSelSalNo
				sInvDet=oNodEntry.Attributes.Item(0).nodeValue &"&nbsp; Dt:"&oNodEntry.Attributes.Item(1).nodeValue
				sRefernceNo=oNodEntry.Attributes.Item(2).nodeValue
				sInvDate = 	oNodEntry.Attributes.Item(1).nodeValue
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

IF CStr(sInvCurr) = "" Then
	sInvCurr = "1"
End IF

sQuery = "Select OtherApplnTransNo From Acc_T_CreatedVoucherHeader "&_
		 "Where FromApplication=3 and TransactionType='SJR' " &_
		 " and  OUDefinitionID='" & sOrgId & "' and CreatedTransNo = "&iTransNo
'Response.Write "<p> sQuery = " & sQuery
Objrs.Open sQuery,Con
If Not Objrs.Eof Then
	sRefSaleTransNo = Objrs(0)
End IF
Objrs.close

if trim(sRefSaleTransNo) <> "" then
	sQuery = "Select TermsDisplay,Remarks from Sal_T_InvoiceHeader where SaleTransactionNo = " & sRefSaleTransNo
	'Response.Write "<p> sQuery = " & sQuery
	Objrs.Open sQuery,Con
	If Not Objrs.Eof Then
		sTerms = objRs(0)
		sRemarks  =objRs(1)
	End IF
	Objrs.close	
end if 


sQuery = "Select ConversionRate,Convert(char,ConversionAsOn,103),CurrencyShortName From "&_
		 "Ms_CurrencyMaster Where CurrencyCode = "&sInvCurr&" "


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
'***************************** To Find the Account Head Defined For Both Sale Type and Item **********************************
sQuery = "Select A.AccountHead From App_R_OrgnTaxAccountHead A,Sal_T_InvoiceHeader I, "&_
		 "Acc_T_CreatedVoucherHeader V Where V.OtherApplnTransNo = I.SaleTransactionNo  "&_
		 "and I.TypeofSale = A.InvoiceType and V.CreatedTransNo = "&iTransNo&" and  "&_
		 "A.Taxcode is Null and A.TaxCategoryCode is Null and A.OUDefinitionID = '"&sOrgId&"' "

'Response.Write sQuery

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sAccType = "S"
	iAccHead = objRs(0)
Else
	iAccHead = 0
End IF
objRs.Close

IF CStr(iAccHead) = 0 Then

	sQuery = "Select T.AccountHead From Inv_M_ItemOrgAccountHead T,Sal_T_InvoiceDetails D, "&_
			 "Acc_T_CreatedVoucherHeader V Where V.OtherApplnTransNo = D.SaleTransactionNo "&_
			 "and D.ItemCode = T.ItemCode and D.Classificationcode = T.Classificationcode  "&_
			 "and V.CreatedTransNo = "&iTransNo&" and T.AccountHeadFor = 'F'  "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sAccType = "I"
		iAccHead = objRs(0)
	Else
		iAccHead = 0
	End IF
	objRs.Close
End IF



IF CStr(iAccHead) <> "0" Then
	iAccHeadCheck = "Y"
	sQuery = "Select AccountDescription From VwAccheadforSalesApp Where AccountHead = "&iAccHead&" "&_
			 "and OUDefinitionID = '"&sOrgId&"' "

	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sAccHeadName = objRs(0)
	End IF
	objRs.Close
Else
	iAccHeadCheck = "N"
End IF

'oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")

'***************************** To Find the Account Head Defined For Both Sale Type and Item **********************************

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData"><Book/></script>
<script type="application/xml" data-itms-xml-island="1" id="SaleData" data-src="<%="../xmldata/Voucher/"&iTransNo&".xml"%>"></script>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AppOtherSALViewCompat.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
ITMSAppOtherSALViewCompat.install();
</script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="VouOtherAppSALAdvance.asp">
<input type="hidden" name="hPara" value="<%=sPara%>">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="05">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hSelDate" value="<%=sInvDate%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<%IF Cstr(sCurrName) = "INR" Then%>
	<input type="hidden" name="hExpInv" value="N">
<%Else%>
	<input type="hidden" name="hExpInv" value="Y">
<%End IF %>
<Input type="hidden" name="hInvno" value="" >
<Input type="hidden" name="hRefFrom" value="Acc" >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher&nbsp;
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<!--<td class="TabCell" valign="bottom" width="125">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Application
                                              Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="96">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher List
											</td>
										</tr>
									</table>
								</td>-->
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								  	<tr>
								  		<td align="center">Voucher</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="96">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Advance
											</td>
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
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0">
                                    <!--<tr>
										<td class="FieldCell">Accounting Unit</td>
										<td class="FieldCellSub" colspan="2">
											<select size="1" name="selUnitId" class="FormElem"  onChange="DispBook(this)">
												<OPTION value="0">Select a Unit</option>
												<%populateUnitSelected(sOrgId)%>
											</select>
										</td>
										<td class="FieldCellSub">
											<%IF CStr(sInvCurr) <> CStr(sOpCurr) Then %>
												<Input type="button" value="View Invoice" class="ActionButtonX" onClick="ViewInv('<%=iTransNo%>')" id=button1 name=button1>
											<%Else
												Response.Write("&nbsp;")
											  End IF
											%>
										</td>
								 </tr>

                                <tr>
                            <td class="FieldCell" >Select Book</td>
                            <td  class="FieldCellSub">
										<select size="1" name="selBook" class="FormElem">
											<option value="S">Select Book</option>
												
										</select>
                            </td>
                            <td class="FieldCell"></td>
                                </tr>-->

                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
                                &nbsp;
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly" width="100%">
                                <tr>
                            <!--<td class="FieldCellsub">Unit </td>
                            <td width="160" class="FieldCellSub">
                                                            <span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                            </td>-->
                            <td class="FieldCellSub">Select Book</td>
                            <td width="160" class="FieldCellSub">
								<select size="1" name="selBook" class="FormElem">
									<option value="S">Select Book</option>
										<%
											IF CStr(sAccBookRel) = "T" Then
												sQuery = "Select BookNumber,BookName From Acc_R_ApplicableAccountHeads  "&_
														 "Where BookCode = '05' and OUDefinitionID = '"&sOrgId&"' "&_
														 "and BookAccountHead = "&iAccHead&" Order By BookName "
											Else
												sQuery =	"Select BookNumber,BookName from "&_
															"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='05' "
											End IF

													'"and isNull(BookAccountHead,0) = "&iAccHead
											'Response.Write "<textarea>"&sQuery&"</textarea>"
												with objRs
													.CursorLocation = 3
													.CursorType = 3
													.Source = sQuery
													.ActiveConnection = con
													.Open
												end with
												set objRs.ActiveConnection = nothing
												while not objRs.EOF
													IF objRs.RecordCount = 1 Then
														Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(1)&"</option>"
													Else
														Response.Write "<option value="""&objRs(0)&""">"&objRs(1)&"</option>"
													End IF
													objRs.MoveNext
												wend
												objRs.Close
											%>
									</select>
                            </td>
                            <td class="FieldCellSub">Voucher No. - Date</td>
                            <td class="FieldCellSub" width="160">
                            	<span class="DataOnly"><%=sInvDet%>&nbsp;</span>
                            	<%IF CStr(sInvCurr) <> CStr(sOpCurr) Then %>
									<!--<Input type="button" value="View Invoice" class="ActionButtonX" onClick="ViewInv('<%'=iTransNo%>')" id=button1 name=button1>-->
									<img src="../../assets/images/iTMS%20icons/DetailsIcon.gif" onClick="ViewInv('<%=iTransNo%>')" alt="View Invoice" style="cursor: pointer" >
								<%Else
									Response.Write("&nbsp;")
								  End IF
								%>
							</td>
                                </tr>
                                <tr>
                            <td class="FieldCellsub">Party Code </td>
                            <td width="160" class="FieldCellSub"><span class="DataOnly"><%=sRefPartyCode%>&nbsp;</span>
                            </td>
                            <td class="FieldCellSub">Reference Number</td>
                            <td class="FieldCellSub" width="160" valign="bottom"><span class="DataOnly"><%=sRefernceNo%>&nbsp;</span>&nbsp;
                            <!--
                            <a href="#" onClick="ShowInvoice('<%'=sRefSaleTransNo%>'); return false;">
						 		<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" alt="Modify / Verify Sales Invoice" width="11" height="11">
						 	</a>
						 	-->
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCellsub">Party Name </td>
                            <td width="320" class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sPartyName%>&nbsp;</span>
                            </td>
                                </tr>
                            <% IF CStr(sAccType) = "S" Then %>
                                <tr>
                            <td class="FieldCellsub">Account Head</td>
                            <td width="320" class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sAccHeadName%>&nbsp;</span>
                            </td>
                                </tr>
                           <%End IF %>

                           <tr>
				<td class="FieldCellsub">Invoice Currency</td>
				<td class="FieldCellSub"><span class="DataOnly"><%=sCurrName%>&nbsp;</span>
				<td class="FieldCellSub">Operating Currency</td>
				<td class="FieldCellSub"><span class="DataOnly">INR </span>
				</td>

                           </tr>


                           <tr>
                            <td class="FieldCellsub">Conversion Rate</td>
                            <td class="FieldCellSub">
                            <%IF Cstr(sCurrName) = "INR" Then%>
                            	<Input type="test" class="FormElemRead" name="txtConRate" Value="1.00" size="9" maxlength="7" readonly style="text-align:right">
                            <%Else%>
                            	<Input type="test" class="FormElem" name="txtConRate" Value="1.00" size="9" maxlength="7" style="text-align:right" onBlur="CalNewTax()">
                            <%End IF %>
                            </td>
                            <!--td class="FieldCellSub">Conversion As On </td>
                            <td class="FieldCellSub"><span class="DataOnly"><%=sConDate%>&nbsp;</span>
                            </td-->
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
                <div class="frmBody" id="frm2" style="height:242;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <% IF CStr(sAccType) = "I" Then %>
		<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description - Account Head</td>
	<%Else%>
		<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
	<%End IF%>
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

	sAmount = oNodDeatils.Attributes.getNamedItem("ActualValue").Value
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

		dTotal=CDbl(dTotal)+CDbl(sAmount)-cdbl(sDiscount)	'Here Discount is Less By UmaMaheswari S,On July 30,2011
		dBasicTotal=CDbl(dBasicTotal)+CDbl(sValue)
		dDisTotal=CDbl(dDisTotal)+CDbl(sDiscount)
		
		''blocked by ragav on Dec 08,2011 for updating the second time the discout value so blocked
		'sAmount = cdbl(sAmount) - cdbl(dDisAmount)	'Added By UmaMaheswari S,On July 30,2011
		''end 
%>
    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <td class="ExcelDisplayCell"><%=sDescription%>
    <% IF CStr(sAccType) = "I" Then
			Response.Write (" - ")
			Response.Write sAccHeadName
		End IF
	%>
    </td>
    <td class="ExcelDisplayCell" align="Left" width="60"><%=sQty%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sRate,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(dRatePer,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right">
    <input type="text" value="<%=FormatNumber(sValue,2,,,0)%>" name="txtBasVal" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right">
    </td>
    <td class="ExcelDisplayCell" align="Right">
    <input type="text" value="<%=FormatNumber(sDiscPer ,2,,,0)%>" name="txtDisper" class="FormElemRead" Readonly size="5" maxlength="4" style="text-align:right">
    </td>
    <td class="ExcelDisplayCell" align="Right">
    <input type="text" value="<%=FormatNumber( sDiscount,2,,,0)%>" name="txtDisVal" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right">
    </td>
    <td class="ExcelDisplayCell" align="Right">
    <input type="text" value="<%=FormatNumber(sAmount,2,,,0)%>" name="txtNetVal" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right"></td>
        </tr>
<%
	next
'end if
'next %>

        <tr>
    <!--<td align="center" colspan="3"></td>-->
    <td class="ExcelSerial" align="center" colspan="5"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right">
    <input type="text" value="<%=FormatNumber(dBasicTotal,2,,,0)%>" name="txtTotBasVal" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right">
    </td>
    <td class="ExcelDisplayCell" align="right"></td>
    <td class="ExcelDisplayCell" align="right" width="60">
    <input type="text" value="<%=FormatNumber(dDisTotal,2,,,0)%>" name="txtTotDisVal" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right">
    </td>
    <td class="ExcelDisplayCell" align="right"><b>
    <input type="text" value="<%=FormatNumber(dTotal,2,,,0)%>" name="txtTotNetVal" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right">

    </td>
        </tr>
        <input type="Hidden" name="hBasicValue" value="<%=dBasicTotal%>">
        <input type="Hidden" name="hDisValue" value="<%=dDisTotal%>">
        <input type="Hidden" name="hAmount" value="<%=dTotal%>">
<%


	For Each oNodEntry in oNodTaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		sFormula=oNodEntry.Attributes.Item(3).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue
		iAccHead = oNodEntry.Attributes.Item(6).nodeValue
		sTaxName=oNodEntry.Text

		IF CStr(iAccHead) <> "" Then
			sQuery = "Select AccountDescription From VwAccheadforSalesApp Where AccountHead = "&iAccHead&" "&_
					 "and OUDefinitionID = '"&sOrgId&"' "
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				sAccHeadName = objRs(0)
			Else
				sAccHeadName = ""
			End IF
			objRs.Close

		End IF

		dTax=oNodEntry.Attributes.Item(5).nodeValue
			%>
			<tr>
				<!--<td align="center" colspan="3"></td>-->
				<td class="ExcelSerial" align="Right" colspan="7"><%=sTaxName%>&nbsp;(<%=sAccHeadName%>)</td>
				<%if sTaxMode="P" then %>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTaxValue,2,,,0)%>&nbsp;%</td>
				<%else%>
					<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if
				  IF Cstr(dTax) = "" Then
					dTax = 0
				  End IF

				%>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" value="<%=FormatNumber(dTax,2,,,0)%>" name="txtTaxVal<%=sCatCode%>Z<%=sTaxCode%>" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right">
				</td>
				    </tr>
			<%
	next
%>
        <tr>
        <!--<td align="center" colspan="3"></td>-->
    <td class="ExcelSerial" align="right" colspan="7"><b>Invoice Value&nbsp; </b></td>
	<td class="ExcelDisplayCell" align="right"></td>
    <td class="ExcelDisplayCell" align="right">
		<input type="text" value="<%=FormatNumber(trim(dInvAmount),2,,,0)%> " name="txtInvVal" class="FormElemRead" Readonly size="15" maxlength="13" style="text-align:right">
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
                                &nbsp;
								</td>
								<td valign="top" class="FieldCell" height="20">
                                    <table cellpadding="0" cellspacing="0" width="100%">

                            <% IF CStr(iAccHeadCheck) = "N" Then %>
                                    <tr>

                            <!--td class="FieldCell" width="130" valign="top">&nbsp; </td-->
                            <td colspan="2">
                                  <span class="DataOnly">Account Head is Not Not Defined for Sale Type or Item so Accounting is Disabled</span>
                            </td>
                           <%End IF %>
                            <tr>
							<td class="FieldCell" valign="top">Terms </td>
                            <td>
								<span class="DataOnly"><%=sTerms%></span>
                            </td>
                            </tr>
                            <tr>
                            <td class="FieldCell" valign="top">Remarks </td>
                            <td>
								<span class="DataOnly" style="color:red"><%=sRemarks%></span>
                            </td>
                            </tr>
                            <tr>
                            <td class="FieldCell" valign="top">Amount </td>
                            <td>
								<span class="DataOnly"><%=AmountWords(dInvAmount)%></span>
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
															<% IF CStr(iAccHeadCheck) = "N" Then %>
                                                                <input type="button" value="<%=sButVal%>" name="btnAction" onclick="checkSubmit()" class="ActionButton" disabled>
                                                            <%Else%>
																 <input type="button" value="<%=sButVal%>" name="btnAction" onclick="checkSubmit()" class="ActionButton">
															<%End IF %>
                                                                <input type="button" value="Keep Pending" name="btnAction2" onclick="finalCancel()" class="ActionButtonX">

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
