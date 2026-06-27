<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNThrMiscPay.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran R
	'Created On					:	Oct 28,2011
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
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName,iAct
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff,sFromPur
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId,sTemp,sRetVal
Dim sFromApp,iItemCode,iClassCode,sSql,rsTemp,sActFor,sActQty,iAmt,sQnty,sQlty,sRat,sCallFrom,iAppRefNo

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
Set rsTemp = server.CreateObject("ADODB.Recordset")
sInvoiceNo=Request.Form("selInvoiceNo")
sBookName=Request.Form("hBookName")
sFromApp = Request.Form("hFromApp")
sCallFrom = Request("CallFrom")

IF Cstr(sInvoiceNo) = "" Then
	sInvoiceNo = "S"
End IF

IF trim(sInvoiceNo) = "S"  then
	sInvoiceNo=Request.Form("selRefNo")
	sTemp = Split(sInvoiceNo,":")
	sInvoiceNo = sTemp(0)
	iAct = sTemp(1)
Else
	sTemp = Split(sInvoiceNo,":")
	sInvoiceNo = sTemp(0)
	iAct = 0

End IF	
'Response.Write "sInvoiceNos="& sInvoiceNo &"--"& iAct

iInvNo = sInvoiceNo

sUserid = getUserID()

sRetVal = GetVouchXML(sInvoiceNo)
oDOM.Load server.MapPath(sRetVal)

set oNodRoot=oDOM.documentElement

for each oNodHeader in oNodRoot.childNodes
    sOrgId = oNodRoot.getAttribute("UnitNo")
    sOrgName = oNodRoot.getAttribute("UnitName")
next

Response.Write "<font color=red>"
sQuery = "Select M.PartyCode,M.CreatedMiscPymtNo,M.MiscTransNo,P.PartyName,M.VoucherAmount,M.AppRefNo from Acc_T_MiscPymtRequestHeader M,App_M_PartyMaster P where M.PartyCode = P.PartyCode and M.ApplicationCode = 3 and ReceiptNo = "& iInvNo
'Response.Write "<p> "& sQuery
objRs.Open sQuery,con
if not objRs.EOF then
    sPartyName = objRs(3)
    sInvoiceNo = objRs(1)        
    sInvValue = objRs(4)
    iAppRefNo = objRs(5)
end if
objRs.Close 
if Trim(iAppRefNo)<>"" then
    sQuery = "Select InvoiceValue from Sal_T_InvoiceHeader where SaleTransactionNo = "& iAppRefNo
'    Response.Write "<p> "& sQuery
    objRs.Open sQuery,con
    if not objRs.EOF then
        sInvValue = objRs(0)
    end if ' if not objRs.EOF then
end if


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="vbscript" >
Function SaveXML()
    document.formname.hVouDate.value = document.formname.ctlDate.getDate
    document.formname.action = "VouCNForMiscRecGenerate.asp"
	document.formname.submit()
End Function

Function EnbApp(sObj)
	IF sObj.value = "Y" Then
		document.formname.selUserId.disabled = False
	Else
		document.formname.selUserId.selectedIndex = 0
		document.formname.selUserId.disabled = True
	End IF
End Function

Function SetRetVal(sObj,sCallType)
	Dim sVal,sPreInvVal,sNewInvVal,dDiffVal,dRowLen
	sVal = sObj.value
	dRowLen = document.formname.hRowVal.value
	
	IF sCallType = "1" Then
		document.formname.reset()
		For iCtr = 0 To sObj.length - 1
			IF sObj.Options(iCtr).value = sVal Then
				sObj.selectedIndex = iCtr
				Exit For
			End IF
		Next
	End IF
	
	
	dPreInvVal = document.all.txtFInvValue.value
	dPreInvVal = CDbl(dPreInvVal)
	sNewInvVal = CDbl(document.formname.txtInvValue.Value)
	
	
	
	'IF CStr(document.formname.hAccType.Value) = "C" Then
	'	dDiffVal = CDbl(dPreInvVal) - CDbl(sNewInvVal)
	'Else
	'	dDiffVal = CDbl(sNewInvVal) - CDbl(dPreInvVal) 
	'End IF
	
	'dDiffVal = FormatNumber(dDiffVal,2,,,0)
	sNewInvVal = FormatNumber(sNewInvVal,2,,,0)
	
	For iCtr = 1 to dRowLen 
		IF CStr(sVal) = "Q" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElem"
			Eval("document.all.tQty"&iCtr).className = "ExcelInputCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"
			
			document.formname.txtCrNoteValue.value = sNewInvVal
			
			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = False
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"
		End IF
	
		IF CStr(sVal) = "R" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElem"
			Eval("document.all.tRate"&iCtr).className = "ExcelInputCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"
			
			document.formname.txtCrNoteValue.value = sNewInvVal
			
			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = False
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"
		End IF
	
		IF CStr(sVal) = "D" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElem"
			Eval("document.all.tDis"&iCtr).className = "ExcelInputCell"
			
			document.formname.txtCrNoteValue.value = sNewInvVal
			
			Eval("document.formname.txtdis"&iCtr).readonly = False
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"
			
		End IF
	
		IF CStr(sVal) = "0" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).class = "ExcelDisplayCell"
			
			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"
				
		End IF
		
		IF CStr(sVal) = "A" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			'Eval("document.all.tDis"&iCtr).class = "ExcelDisplayCell"
			
			
			
			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = False
			document.formname.txtCrNoteValue.className = "FormElem"
			document.all.tDrVal.classname = "ExcelInputCell"
			
			ResetTax()
		End IF
		
	Next
End Function

Function ResetTax()
	dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax
	set RootNode=TaxData.documentElement

	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	
	
	dInvAmount = 0
	dBasicTotal = 0
	dTotal = 0
	dInvAmount = CDbl(dInvAmount)
	dBasicTotal = CDbl(dBasicTotal)
	dTotal = CDbl(dTotal)
		
	For Each oNodEntry in TaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue 
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue 
		oNodEntry.Attributes.Item(5).nodeValue="0.00"
		eval("document.formname.txtTaxValue"&sCatCode&sTaxCode).value="0.00"
	next	
	
	
	dInvAmount=FormatNumber(dInvAmount,2,,,0)
	document.formname.txtInvValue.value=dInvAmount
	Taxroot.Attributes.Item(0).Nodevalue = "0.00"
	TaxRoot.Attributes.Item(1).nodeValue="0.00"
	TaxRoot.Attributes.Item(2).nodeValue="0.00"
	
End Function

Function CheckNoSer()
	Dim ObjHttp,sPassVal,sMon,sYear,sDate,sPeriod,sRetVal
	Set ObjHttp = CreateObject("MSXML2.XMLHTTP")
	IF Cstr(document.formname.hVouCode.Value) = "04" Then
		sPassVal = document.formname.selUnitId.Value 'Unit	
		sPassVal = sPassVal&":"&document.formname.hVouCode.Value 'BookCode
		sPassVal = sPassVal&":"&document.formname.hCallFrm.Value 'Call Fro Created or Accounted
		sPassVal = sPassVal&":D"
		sPassVal = sPassVal&":"&document.formname.selBook.Value 'Book Number 
	Else
		sPassVal = document.formname.hOrgid.Value 'Unit
		sPassVal = sPassVal&":"&document.formname.hVouCode.Value 'BookCode
		sPassVal = sPassVal&":"&document.formname.hCallFrm.Value 'Call Fro Created or Accounted
		IF Cstr(document.formname.hVouCRDR.Value) = "" Then
			sPassVal = sPassVal&":"&"D"
		Else
			sPassVal = sPassVal&":"&document.formname.hVouCRDR.Value ' Voucher Type C/D
		End IF
		sPassVal = sPassVal&":"&document.formname.hBookCode.Value 'Book Number 
	End IF
	
	sPeriod = document.formname.ctlDate.GetDate()
	sPassVal = sPassVal&":"&sPeriod 'Voucher Date
	ObjHttp.open "GET","NoSeriesCheck.asp?sValue="&sPassVal, False
	ObjHttp.send
	sRetVal = ObjHttp.responseText
	
	IF CStr(Trim(sRetVal)) = "T" Then
		CheckNoSer = True
	Elseif CStr(Trim(sRetVal)) = "F" Then 
		MsgBox "No Series is Not Defined "
		CheckNoSer = False
		Exit Function
	Else
		MsgBox "Error "
		CheckNoSer = False
		Exit Function
	End IF 
End Function

Function AccHead(objAcc)
dim sOrgId,sTemp,sDesc
	sOrgId=document.formname.hOrgId.value
	if objAcc.selectedIndex >0 then
		if objAcc.value="G" then
			showGLHead sOrgId
		End if 'End of select Account Head Type check GL or PARTY
	End if 'End of If any Account Head Selected Check
End function

Function showGLHead(sOrgId)
	Dim iAccCode,sRetVal,arrTemp,sDesc,sTemp
	
	iBookNo=document.formname.hBookcode.value

	OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")

	while UBound(arrTemp) = 0
		OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend

	sRetVal = OutValue

	if UBound(arrTemp) <= 1 then exit function
	sTemp = Split(sRetVal,":")
	iAccCode = sTemp(0)
	sDesc = sTemp(5)
	
	document.formname.hCrAccHead.value = iAccCode
	document.all.spAccHead.innerHTML = sDesc
	
	
End function

Function ReTotalCr()
	Dim sExp,TempNode,RootNode,iTax,iCat,dTaxVal,iCtr,sObj,dSalVal,dCrVal
	dSalVal = document.formname.txtTotal.Value
	set RootNode=TaxData.documentElement
	sExp = "//Tax"
	Set TempNode = RootNode.selectNodes(sExp)
	dTaxVal = 0
	IF TempNode.Length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iTax = TempNode.Item(iCtr).Attributes.getNamedItem("TaxCode").Value
			iCat = TempNode.Item(iCtr).Attributes.getNamedItem("CatCode").Value
			Set sObj = Eval("document.formname.txtTaxValue"&iCat&iTax)
			TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value = sObj.Value
			dTaxVal = Cdbl(dTaxVal) + CDbl(sObj.Value)
		Next
	End IF
	dCrVal = CDbl(dSalVal) + CDbl(dTaxVal)
	dCrVal = FormatNumber(dCrVal,2,,,0)
	document.formname.txtInvValue.Value = dCrVal
	document.formname.txtCrNoteValue.value = dCrVal
	
	sExp = "//TaxDetails"
	Set TempNode = RootNode.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.item(0).Value = dCrVal
	End IF
	
End Function

</script>
</HEAD>
<BODY leftMargin=0 topMargin=0>
<form method="POST" name="formname">
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hAccType" value="C">
<Input type="hidden" name="hCallType" value="OINV">
<Input type="hidden" name="hNoteType" value="C">
<Input type="hidden" name="hFromPur" value="<%=sFromPur%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hVouCode" value="06">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="<%=Request.Form("selBook")%>">
<input type="hidden" name="hCrAccHead" value="0">
<input type="hidden" name="hCallFromDebit" value= "<%=sCallFrom%>">
<input type="hidden" name="hVouDate" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"> 
				Credit Note (Misc. Receipts)
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
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<!--<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<tr><td align="center">Voucher</td>
								  	</tr>
								  </table>
								</td>-->
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
                <div class="frmBody" id="frm2" style="width: 775; height:150;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
            <tr>
                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                <td class="ExcelHeaderCell" align="center" >Description</td>
                <td class="ExcelHeaderCell" align="center" >Value</td>
            </tr>
            <%
                iSno = 0
	            For each oNodHeader in oNodRoot.childNodes
            	    if oNodHeader.nodeName="Entry" then
            	    iSno =iSno +1
            	        sInvValue = oNodHeader.getAttribute("Amount")
            	        dTotal = cdbl(dTotal) + CDbl(sInvValue)
            	        for each oNodtemp in oNodHeader.childNodes
            	            if oNodtemp.nodeName="Narration" then
            	                sDescription = oNodtemp.text
            	            end if
            	        next
            	    %>
            	        <tr>
                            <td class="ExcelHeaderCell" align="center"><%=iSno%></td>
                            <td class="ExcelDisplayCell" align="Left"><%=Trim(sDescription)%></td>
                            <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sInvValue,2,0,0,0)%></td>
                        </tr>
            	    <%
            	    end if
	            Next
            %>
            <tr>
                <td class="ExcelSerial" align="right" colspan="2"><b>Debit Note Value&nbsp; </b></td>
                <td class="ExcelDisplayCell" align="right" id="tDrVal"> <input type="text" style="text-align: Right" NAME="txtCrNoteValue"  size="13" value="<%=FormatNumber(dTotal,2,,,0)%>" class="FormelemRead" readonly>
                </td>
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
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onClick="SaveXML()" >
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
