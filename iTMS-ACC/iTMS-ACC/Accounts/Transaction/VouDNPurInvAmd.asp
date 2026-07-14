<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNPurInvAmd.asp
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff,sFromPur
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId,sTemp
Dim sFromApp,dOrgQty,dOrgRate,dOrgDis,iTransNo,sCrAgain,dTotQty,dTotRate,dTotDis
Dim ObjInv,sNarration,iCrAltAccHd,sCrAltAccName,sAccChg,iBookName
Dim sVouDate,sAmdTy,iSelBook,sRetVal,sForAmd,dInvInvAmount

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set ObjInv = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

iTransNo=Request.Form("hTransNo")
sBookName=Request.Form("hBookName")
sFromApp = Request.Form("hFromApp")
sAmdTy = Request("AmdType")
sForAmd = Request("sForAmd")


sTemp = Split(iTransNo,"-")
iTransNo = sTemp(0)

sUserid = getUserID()

'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)

'Response.Write iTransNo

set oNodRoot=oDOM.documentElement

for each oNodHeader in oNodRoot.childNodes
	if oNodHeader.nodeName="Header" then
		for Each oNodEntry in  oNodHeader.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.text
			end if
			if oNodEntry.nodeName="Book" then
				iBookName = oNodEntry.Attributes.Item(0).nodeValue
				iSelBook = oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="PurInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue&"&nbsp;-&nbsp;"&oNodEntry.Attributes.Item(1).nodeValue
				iInvNo = oNodEntry.Attributes.getNamedItem("CrTransNo").Value
			end if
		next
	end if

	if oNodHeader.nodeName="Details" then
		sVouDate = oNodHeader.Attributes.Item(3).nodeValue
		set oNodDeatils=oNodHeader
	end if
	if oNodHeader.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodHeader
	end if
	if oNodHeader.nodeName="AgentDetails" then
		set oNodtemp=oNodRoot.removeChild(oNodHeader)
	end if
next
dim dInvAmount,sExp,TempNode,iCtr

dTotQty = 0
dTotRate = 0
dTotDis = 0

sQuery = "Select Sum(InvoicedQuantity),Sum(InvoicedRate),Sum(DiscountPercent) From  "&_
		 "Acc_T_CreatedVoucherDetails Where CreatedTransNo = "&iInvNo&" "
		 
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open 
End With
Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	dOrgQty = objRs(0)
	dOrgRate = objRs(1)
	dOrgDis = objRs(2)
End IF
objRs.Close

sExp = "//Entry"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.Length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		dTotQty = CDbl(dTotQty) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Qty").Value)
		dTotRate = CDbl(dTotRate) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Rate").Value)
		dTotDis = CDbl(dTotDis) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("DisPer").Value)
	Next
End IF

IF CDbl(dOrgQty) <> CDbl(dTotQty) Then
	sCrAgain = "Q"
Elseif CDbl(dOrgRate) <> CDbl(dTotRate) Then
	sCrAgain = "R"
Elseif CDbl(dOrgDis) <> CDbl(dTotDis) Then
	sCrAgain = "D"
Else
	sCrAgain = "A"
End IF


sInvValue=oNodTaxRoot.Attributes.Item(2).nodeValue
dInvAmount=oNodTaxRoot.Attributes.Item(0).nodeValue




sExp = "//Voucher/Narration"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.Length <> 0 Then
	sNarration = Trim(TempNode.Item(0).Text)
End IF

sExp = "//AccHead"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iCrAltAccHd = TempNode.Item(0).Attributes.getNamedItem("No").Value
	sCrAltAccName = TempNode.Item(0).Attributes.getNamedItem("Name").Value
End IF

sQuery = "Select AccUnitAccountHead From Acc_T_CreatedVoucherDetails Where  "&_
		 "AccUnitAccountHead = "&iCrAltAccHd&" and CreatedTransNo = "&iInvNo&"  "&_
		 "and VoucherEntryNumber = 1 "
		 
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sAccChg = "N"
	iCrAltAccHd = 0
Else
	sAccChg = "Y"
End IF
objRs.Close

sRetVal = GetVouchXML(iInvNo)
ObjInv.Load server.MapPath(sRetVal)
Dim oInvNodRoot,oInvNodHeader,oInvNodEntry,oInvNodTaxRoot,oInvNodDeatils
Set oInvNodRoot = ObjInv.documentElement

for each oInvNodHeader in oInvNodRoot.childNodes
	if oInvNodHeader.nodeName="Details" then
		set oInvNodDeatils=oInvNodHeader
	end if
	if oInvNodHeader.nodeName="TaxDetails" then
		set oInvNodTaxRoot=oInvNodHeader
	end if
next
dInvInvAmount=oInvNodTaxRoot.Attributes.Item(0).nodeValue

ObjInv.save server.MapPath("../temp/transaction/InvDet_ForDN_"&Session.SessionID&".xml")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="TaxData" data-src="<%="../temp/transaction/Voucher EntryAmd_DN_"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="InvData" data-src="<%="../temp/transaction/InvDet_ForDN_"&Session.SessionID&".xml"%>"></script>

<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT SRC="../../scripts/SalesReturnCreditNoteEntryCompat.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script>
ITMSSalesReturnCreditNoteEntryCompat.install();
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<% IF CStr(sForAmd) <> "Y" Then %>
	<form method="POST" name="formname" action="VouDNPurInvAmdUpdate.asp">
<%Else%>
	<form method="POST" name="formname" action="AmdAccDbNtGenerate.asp">
<%End IF %>
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hAccType" value="C">
<Input type="hidden" name="hCallType" value="OINV">
<Input type="hidden" name="hNoteType" value="C">
<Input type="hidden" name="hFromPur" value="<%=sFromPur%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hVouCode" value="06">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="<%=iBookName%>">
<input type="hidden" name="hCrTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hCrAccHead" value="<%=iCrAltAccHd%>">
<input type="hidden" name="hVouDate" value="<%=sVouDate%>">
<input type="hidden" name="hAmdTy" value="<%=sAmdTy%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"> 
		Debit Note For Purchase Invoice Amendment
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
                                                    <td class="FieldCellSub" width="100">Unit </td>
                                                    <td class="FieldCell">  <span class="DataOnly"><%=sOrgName%></span></td>
                                                    <td class="FieldCellSub" width="75"><p align="left">Date</p></td>
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
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=FormatNumber(sInvValue,2,,,0)%> </span></td>
                                                        </tr>
                                                        
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Dr Note Against</td>
                                                    <td class="FieldCell" width="200">  
                                                    <Select name="SelCrAgain" class="FormElem" onChange="SetRetVal(this,'2')">
                                                    <Option Value="0">Select</Option>
                                                    <%
													  IF CStr(sCrAgain) = "Q" Then	
													%>
														<Option Value="Q" Selected>Quantity</Option>
                                                    <%Else%>
														<Option Value="Q">Quantity</Option>
													<%End IF 
                                                    
													  IF CStr(sCrAgain) = "R" Then	
													%>
														<Option Value="R" Selected>Rate</Option>
                                                    <%Else%>
														<Option Value="R">Rate</Option>
													<%End IF 
													
													  IF CStr(sCrAgain) = "D" Then	%>
														<Option Value="D" Selected>Discount</Option>
													<%Else%>
														<Option Value="D">Discount</Option>
													<%End IF 
													  IF CStr(sCrAgain) = "A" Then
													%>
														<Option Value="A" Selected>Quality</Option>
                                                    <%Else%>
														<Option Value="A">Quality</Option>
													<%End IF %>
                                                    </Select>
                                                    </td>
                                                    </tr>
                                                    <tr>
                                                    <%IF CStr(sAmdTy) = "A" Then %>
															
																<td class="FieldCellSub" width="100">Book</td>
																<td class="FieldCell" colspan="3">
																	<select size="1" name="selBook" class="FormElem">
												<%
																	sQuery = "Select BookNumber,BookName From VwOrgBookNames Where  "&_
																			 "OUDefinitionID = '"&sOrgId&"' And BookCode = '06' Order By BookName "
																	With objRs
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = Con
																		.Open
																	End With
																	Set objRs.ActiveConnection = Nothing
																	Do While Not objRs.EOF
																		IF CStr(iSelBook) = CStr(objRs(0)) Then 
												%>
																				<Option Value="<%=objRs(0)%>" Selected><%=objRs(1)%></Option>
												<%						Else%>
																				<Option Value="<%=objRs(0)%>"><%=objRs(1)%></Option>
												<%		
																		End IF
																		objRs.MoveNext
																	Loop
																	objRs.Close
																			
												%>
															</select>
														</td>
													</tr>  
												<%	End IF%>	        
                                                    
                                                        </tr>
                                                         <tr>
														<td class="FieldCellSub" width="200">Select Account Head</td>
														<td class="FieldCell" colspan="3">  
														<Select name="SelAccountHd" class="FormElem" onChange="AccHead(this)">
														<%IF CStr(sAccChg) = "N" Then %>
																<Option Value="0" Selected>ITEM ACCOUNT HEAD</Option>
																<Option Value="G">GL ACCOUNT HEAD</Option>
														<%Else%>
																<Option Value="G"  Selected>GL ACCOUNT HEAD</Option>
														<%End IF %>
														</Select>
														&nbsp; <a href="#" onclick="AccHead(document.formname.SelAccountHd); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Account Head"></a>
														</td>
														<td class="FieldCellSub" colspan="2">
														<%IF CStr(sAccChg) = "N" Then %> 
															<span class="DataOnly" id="spAccHead"></span>
														<%Else%>
															<span class="DataOnly" id="spAccHead"><%=sCrAltAccName%></span>
														<%End IF%>
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
    <td class="ExcelHeaderCell" align="center" width="55">Rate</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="55">Rate</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
        </tr>
<%
	Dim iInvSno,sInvDescription,sInvAmount,sInvQty
    Dim sInvInvValue,sInvRate,sInvRatePer,sInvDiscount,sInvDiscPer,dInvTotal,iInvItemCode,iInvClassCode
	Dim sRatePer,iItemCode,iClassCode
	
	For Each oInvNodEntry in oInvNodDeatils.childNodes
	    iSno=oInvNodEntry.getAttribute("No")
		sInvDescription=oInvNodEntry.getAttribute("PayTo")
		sInvAmount=oInvNodEntry.getAttribute("Amount")
		sInvQty=oInvNodEntry.getAttribute("Qty")
		sInvInvValue=oInvNodEntry.getAttribute("ActValue")
		sInvRate = oInvNodEntry.getAttribute("Rate")
		sInvRatePer = 1
		sInvDiscount=oInvNodEntry.getAttribute("DisAmount")
		sInvDiscPer = oInvNodEntry.getAttribute("DisPer")
		iInvItemCode = oInvNodEntry.getAttribute("ItemCode")
		iInvClassCode = oInvNodEntry.getAttribute("ClassCode")
		dInvTotal=CDbl(dInvTotal)+CDbl(sInvAmount)
		
	    sCheckExp = "//Entry[@ItemCode="& iInvItemCode &" and @ClassCode="& iInvClassCode &" and @Amount>0]"
		set oNodtemp = oNodDeatils.selectNodes(sCheckExp)
		if oNodtemp.length> 0 then
            sAmount=oNodtemp.Item(0).Attributes.GetNamedItem("Amount").value
            sQty=oNodtemp.Item(0).Attributes.GetNamedItem("Qty").value
            sValue=oNodtemp.Item(0).Attributes.GetNamedItem("ActValue").value
            sRate = oNodtemp.Item(0).Attributes.GetNamedItem("Rate").value
            sRatePer = 1
            sDiscount=oNodtemp.Item(0).Attributes.GetNamedItem("DisAmount").value
            sDiscPer = oNodtemp.Item(0).Attributes.GetNamedItem("DisPer").Value
            iItemCode = oNodtemp.Item(0).getAttribute("ItemCode")
            iClassCode = oNodtemp.Item(0).getAttribute("ClassCode")
            dTotal=CDbl(dTotal)+CDbl(sAmount)
        else
            sAmount="0"
            sQty="0"
            sValue="0"
            sRate = "0"
            sRatePer = 1
            sDiscount="0"
            sDiscPer = "0"
            iItemCode = "0"
            iClassCode = "0"
            if sCrAgain="Q" then
                sQty="0"
                sRate = sInvRate 
                sDiscount = sInvDiscount 
                sDiscPer = sInvDiscPer 
            elseif sCrAgain = "R" then
                sQty=sInvQty 
                sRate = "0"
                sDiscount = sInvDiscount 
                sDiscPer = sInvDiscPer 
            elseif sCrAgain = "D" then
                sQty=sInvQty 
                sRate = sInvRate 
                sDiscount = "0"
            end if
		end if

    %>

        <tr>
        <td class="ExcelSerial" align="center"><%=isno%></td>
        <td class="ExcelDisplayCell"><%=sInvDescription%></td>
        <td class="ExcelDisplayCell" align="Right" id="tOldQty<%=isno%>"><%=FormatNumber(sInvQty,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="Right" id="tOldRate<%=isno%>"><%=FormatNumber(sInvRate,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="Right" id="tOldDis<%=isno%>"><%=FormatNumber(sInvDiscount,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtOldAmount<%=isno%>"  value="<%=FormatNumber(sinvAmount,2,,,0)%>" class="FormelemRead" size="13"></td>
        <td class="ExcelDisplayCell" align="Right" id="tQty<%=isno%>"><input type="text" style="text-align: Right" NAME="txtqty<%=isno%>" onBlur="setQty(this,'<%=iSno%>','Q')" value="<%=FormatNumber(sQty,3,,,0)%>" class="FormelemRead" size="13"></td>
        <%IF CStr(sCrAgain) = "R" Then %>
		    <td class="ExcelInputCell" align="Right" id="tRate<%=isno%>"><input type="text" style="text-align: Right" NAME="txtRate<%=isno%>" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')" value="<%=FormatNumber(sRate,2,,,0)%>" class="Formelem" size="13"></td>
	    <%Else%>
		    <td class="ExcelDisplayCell" align="Right" id="tRate<%=isno%>"><input type="text" style="text-align: Right" NAME="txtRate<%=isno%>" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')" value="<%=FormatNumber(sRate,2,,,0)%>" class="FormelemRead" size="13" readonly></td>
	    <%End IF %>
        <input type="hidden" name="hRatePer" value="<%=sRatePer%>">
        </td>
        <%IF CStr(sCrAgain) = "D" Then %>
		    <td class="ExcelInputCell" align="Right" id="tDis<%=isno%>">
		    <input type="hidden" name="hDisPer<%=iSNo%>" value="<%=sDiscPer%>">
		    <input type="text" style="text-align: Right" NAME="txtDis<%=isno%>" value="<%=FormatNumber(sDiscount,2,,,0)%>" class="FormelemRead" size="13" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')">
	    <%Else%>
		    <td class="ExcelDisplayCell" align="Right" id="tDis<%=isno%>">
		    <input type="hidden" name="hDisPer<%=iSNo%>" value="<%=sDiscPer%>">
		    <input type="text" style="text-align: Right" NAME="txtDis<%=isno%>" value="<%=FormatNumber(sDiscount,2,,,0)%>" class="FormelemRead" size="13" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')">
	    <%End IF %>
        </td>
        

        <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iSno%>" onBlur="setTotal(this,'<%=iSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13"></td>
	    <!--td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13"></td-->
            </tr>
    <%
            
    Next	  
    %>

        <tr>
        <Input type="hidden" name="hRowVal" value="<%=isno%>">
    <td align="center" ></td>
   
    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td align="center" ></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dInvTotal,2,,,0)%></b></td>
    <td align="right" ></td>
    <td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotalInv" value="<%=FormatNumber(dInvTotal,2,,,0)%>" class="FormelemRead" size="13"></td>
     <td align="right" colspan="3"></td>
     <td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotal" value="<%=FormatNumber(dTotal,2,,,0)%>" class="Formelem" size="13"></td>
      
        </tr>



<%
	Dim sCheckExp,CheckNode
'	dim dInvAmount
'	dInvAmount=dTotal
	Dim iCheck
	Dim sInvCatCode,sInvTaxCode,sInvTaxMode,sInvFormula,dInvTaxValue,iInvRoundOff,dInvTax,sInvTaxName
	
	For Each oInvNodEntry in oInvNodTaxRoot.childNodes
	    sInvCatCode = oInvNodEntry.getAttribute("CatCode")
	    sInvTaxCode = oInvNodEntry.getAttribute("TaxCode")
	    sInvTaxMode = oInvNodEntry.getAttribute("TaxMode")
	    sInvFormula = oInvNodEntry.getAttribute("TaxFormula")
	    dInvTaxValue = oInvNodEntry.getAttribute("TaxValue")

	    For Each oNodEntry in oNodTaxRoot.childNodes
		    sCatCode=oNodEntry.Attributes.GetNamedItem("CatCode").value
		    sTaxCode=oNodEntry.Attributes.GetNamedItem("TaxCode").value
		    sTaxMode=oNodEntry.Attributes.GetNamedItem("TaxMode").value
		    sFormula=oNodEntry.Attributes.GetNamedItem("TaxFormula").value
		    dTaxValue=oNodEntry.Attributes.GetNamedItem("TaxValue").value
		    
		    If sInvCatCode = "0" and sInvTaxCode = "0" and sInvTaxMode = "0" Then
			    iInvRoundOff = 0
		    Else
			    sCheckExp = "//TaxDetails/Tax[@CatCode="&sInvCatCode&" and @TaxCode="&sInvTaxCode&" and @RoundOff]"
			    Set CheckNode = oNodRoot.selectNodes(sCheckExp)
			    IF CheckNode.length <> 0 Then
				    iInvRoundOff = oInvNodEntry.Attributes.GetNamedItem("RoundOff").value
			    Else
				    iInvRoundOff  = 0
			    End IF
		    End If
		    sInvTaxName=oNodEntry.Text
		    If iInvRoundOff = 1 Then
			    dInvTax  = FormatNumber(Round(oInvNodEntry.Attributes.GetNamedItem("TaxAmount").value,0),2,,,0)
		    Else
			    dInvTax = FormatNumber(oInvNodEntry.Attributes.GetNamedItem("TaxAmount").value,2,,,0)
		    End If


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
			    'Response.Write dTax
		    End If
		    if Trim(sInvTaxCode)=Trim(sTaxCode) and Trim(sInvCatCode)=Trim(sCatCode) then
    %>
			    <tr>
				    <td align="center" colspan="1"></td>
				    <td class="ExcelSerial" align="right" colspan="3"><%=sInvTaxName%>&nbsp;</td>
				    <%if sInvTaxMode="P" then %>
				    <td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="FormelemRead" readonly>&nbsp;%</td>
				    <%else%>
				    <td class="ExcelDisplayCell" align="right">
				    <%
					    if sTaxMode="K" then Response.Write "Per Pack"
					    if sTaxMode="Q" then Response.Write "Per Qty"
				    %>
				    </td>
				    <%end if%>
				    <td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtOldTaxValue<%=sInvCatCode%><%=sInvTaxCode%>" value="<%=dInvTax%>"  size="11" class="FormelemRead"></td>
				    <td align="center" colspan="3"></td>
				    <td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  size="11" class="Formelem" onBlur="ReTotalCr()"></td>
    				
				        </tr>
    				    
			    <%
			 end if 'if Trim(sInvTaxCode)=Trim(sTaxCode) and Trim(sInvCatCode)=Trim(sCatCode) then
	    next
    Next	    
	
	
oDOM.save server.MapPath("../temp/transaction/Voucher EntryAmd_DN_"&Session.SessionID&".xml")

%>

		

        <tr>
        <td align="center" colspan="1"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <input type="text" style="text-align: Right" NAME="txtFInvValue"  size="13" value="<%=FormatNumber(dInvInvAmount,2,,,0)%>" class="FormelemRead"></td>
    <td align="right" colspan="3"></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="Formelem">
    
    </td>
        </tr>
        
        <tr>
        <td align="center" colspan="1"></td>
        <td class="ExcelSerial" align="right" colspan="8"><b>Debit Note Value&nbsp; </b></td>
    <!--td class="ExcelSerial" align="right" colspan="3"><b>Credit Note Value&nbsp; </b></td-->
    <td class="ExcelDisplayCell" align="right" id="tDrVal"> <input type="text" style="text-align: Right" NAME="txtCrNoteValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="FormelemRead" readonly>
    
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
			
			<Textarea name="txtNarration" class="FormElem" cols="40" rows="4"><%=sNarration%></Textarea>
			
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
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" onClick="Cancel('DebitVouchers.ASP')" >
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
