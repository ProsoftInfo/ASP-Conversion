<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppOtherPURViewWithMulTax.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	September 30, 2006
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include File="../../Common/GetAccHeadForPartyType.asp"-->
<%
Dim oDOM,oNodRoot,oNodTemp,oNodDeatils,oNodTaxRoot,oNodEntry,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo
dim sDiscPer,dBasicTotal,dDisTotal
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue
dim iTransNo,sOrgName,sBookName,sParType,sParSubType,sParCode,sBookNo
Dim iAccHead,sAccHeadName,sAccCheck,sVouDate,sTmDate,sPurDate,dTstTotal
Dim iPurTaxTy,sChkMulAcc,iBookAccHead,sPurBillTy,sRetVal
dim dInvAmount,sExp,TempNode,sPara,sButVal,sFlagPartyAccHead,sSubTypeName


' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sTmDate = Date()+1
sTmDate = FormatDate(sTmDate)
 
iTransNo=Request("TransNo")
sPara = Request("sPara")
'Response.Write "iTransNo="&iTransNo & "--" &sPara
If trim(sPara) = "App" then  
	sButVal = "Approve"
ElseIf trim(sPara) = "Edt" then  
	sButVal = "Edit"
ElseIf trim(sPara) = "Acc" then  
	sButVal = "Account"
Else
	sButVal = "Next"
End If

    sQuery = "Select Distinct AccUnitAccountHead from Acc_T_CreatedVoucherDetails "&_
		     "Where CreatedTransNo = "&iTransNo
    With Objrs
	    .ActiveConnection = Con
	    .CursorType = 3
	    .CursorLocation = 3
	    .Source = sQuery
	    .Open 
    End With
    Set Objrs.ActiveConnection = Nothing
    sChkMulAcc = Objrs.RecordCount
    IF Not objRs.EOF Then
	    iBookAccHead = objRs(0)
    End IF
Objrs.Close

sQuery = "Select Convert(Char,VoucherDate,103),isNull(PurchaseBillType,''),VoucherAmount From Acc_T_CreatedVoucherHeader  "&_
		 "Where CreatedTransNo = "&iTransNo&" "
		 'Response.Write sQuery 
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouDate = objRs(0)
	sPurBillTy = objRs(1)
	dInvAmount = objRs(2)
End IF
objRs.Close


'oDOM.load  server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)
'Response.Write "<p><font color=red>sRetVal="&sRetVal & "  " & iTransNo

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
			if oNodEntry.nodeName="PurInvoice" then
				sRefernceNo=oNodEntry.Attributes.Item(0).nodeValue &"&nbsp; Dt:"&oNodEntry.Attributes.Item(1).nodeValue
				sPurDate = Trim(oNodEntry.Attributes.Item(1).nodeValue)
			end if


		next
	end if
	'if oNodTemp.nodeName="TaxDetails" then
	'	dTstTotal = cdbl(oNodTemp.Attributes.Item(0).nodeValue) + cdbl(oNodTemp.Attributes.Item(3).nodeValue)
	'end if

	if oNodTemp.nodeName="Details" then
		set oNodDeatils=oNodTemp
	end if

	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
	end if
next

'Response.Write dTstTotal &"<br>"
'Response.Write FormatNumber(dTstTotal,2,,,0)

iPurTaxTy = 0

sQuery = "Select Distinct InvoiceType From Acc_T_CreatedVoucherDetails  "&_
		 "Where CreatedTransNo = "&iTransNo&"  "
objRs.Open sQuery,Con
Do While Not objRs.EOF
	iPurTaxTy = iPurTaxTy&":"&objRs(0) 
	objRs.MoveNext
Loop
objRs.Close


if Trim(sParType)<>"" then
    sFlagPartyAccHead = GetAccHeadForPartyType(sParType,sParSubType,sParCode,sOrgId)
    
    sQuery= "Select SubTypeName from APP_M_PartyTypes where PartyType='"& sParType &"' and PartySubType="& sParSubType 
    
    objRs.Open sQuery,con
    if not objRs.EOF then
        sSubTypeName = Trim(objRs(0))
    end if
    objRs.Close 
    
end if 'if Trim(sParType)<>"" then



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
function appOtherPurField(name) {
	var frm = document.formname;
	return frm && (frm.elements[name] || frm[name]) || null;
}

function appOtherPurValue(name) {
	var item = appOtherPurField(name);
	return item && item.value || "";
}

function appOtherPurDateValue(name) {
	var item = appOtherPurField(name);
	if (item && typeof item.GetDate === "function") {
		return item.GetDate();
	}
	if (item && typeof item.getDate === "function") {
		return item.getDate();
	}
	return item && item.value || "";
}

function appOtherPurParseDate(value) {
	var text = String(value || "").replace(/^\s+|\s+$/g, "");
	var parts;
	if (!text) {
		return null;
	}
	parts = text.split(/[\/.-]/);
	if (parts.length < 3) {
		return null;
	}
	if (parts[0].length === 4) {
		return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
	}
	return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
}

function appOtherPurDateDiff(fromDate, toDate) {
	var fromValue = appOtherPurParseDate(fromDate);
	var toValue = appOtherPurParseDate(toDate);
	if (!fromValue || !toValue) {
		return 0;
	}
	return Math.floor((toValue.getTime() - fromValue.getTime()) / 86400000);
}

function checkSubmit() {
	var frm = document.formname;
	var selectedDate;
	var selectedBook;
	var partyAccountHead;
	if (!frm) {
		return false;
	}
	selectedBook = appOtherPurField("selBook");
	if (selectedBook && selectedBook.selectedIndex < 1) {
		alert("Select a Book");
		selectedBook.focus();
		return false;
	}
	partyAccountHead = appOtherPurValue("hFlagForParAccHead");
	if (partyAccountHead === "0" || partyAccountHead === "") {
		alert("Account Head is not Mapped for the party type " + appOtherPurValue("hSubTypeName"));
		return false;
	}
	selectedDate = appOtherPurDateValue("ctlAccDate");
	appOtherPurField("hSelDate").value = selectedDate;
	if (appOtherPurDateDiff(appOtherPurValue("hPurDate"), selectedDate) >= 0) {
		if (appOtherPurDateDiff(selectedDate, appOtherPurValue("hCurrDate")) < 0) {
			alert("Account Date Should be Between " + appOtherPurValue("hPurDate") + " To " + appOtherPurValue("hCurrDate"));
			return false;
		}
	} else {
		alert("Account Date Should be Greater Than or Equal To " + appOtherPurValue("hPurDate"));
		return false;
	}
	appOtherPurField("hSelInvDate").value = selectedDate;
	appOtherPurField("hSelDate").value = selectedDate;
	appOtherPurField("hBookName").value = selectedBook.options[selectedBook.selectedIndex].text;
	appOtherPurField("btnAction").disabled = true;
	frm.submit();
	return true;
}

function finalCancel() {
	document.formname.action = "PURCHASEVOUCHERS.ASP";
	document.formname.submit();
}

function SetDate() {
	var control = appOtherPurField("ctlAccDate");
	var value = appOtherPurValue("hPassDate");
	if (control && typeof control.SetDate === "function") {
		control.SetDate(value);
	} else if (control && typeof control.setDate === "function") {
		control.setDate(value);
	} else if (control) {
		control.value = value;
	}
}

function AmdInvPurInvoice() {
	var frm = document.formname;
	frm.action = "../../Purchase/TRANSACTION/AmdInvPurInvoiceEntry.asp?ForUnit=" + appOtherPurValue("hOrgId") +
		"&InvNo=" + appOtherPurValue("hInvNo") +
		"&hRcptNo=" + appOtherPurValue("hRcptNo") +
		"&ItemType=" + appOtherPurValue("hItemType");
	frm.submit();
}
</script>
<%
Dim iInvNo,iRcptNo,sItemType
sQuery = "Select isNull(OtherApplnTransno,0) from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iTransNo 
'Response.Write sQuery 
objRs.Open sQuery,con
If not objRs.EOF then 
	iInvNo = objRs(0)
End If
objRs.Close

sQuery = "select Distinct Receiptnumber from Pur_t_refferencenumberDet where invoicenumber = "& iInvNo
objRs.Open sQuery,con
do while not objRs.EOF 
	iRcptNo = iRcptNo &","& objRs(0)
objRs.MoveNext
loop
objRs.Close
iRcptNo = mid(iRcptNo,2)
if trim(iRcptNo)<>"" then
   ' sQuery = "Select ItemType from Rcv_T_ActualReceiptheader where receiptnumber in ("&iRcptNo&") "
    ''Response.Write sQuery
    'objRs.Open sQuery,con
    'If not objRs.EOF then 
	'    sItemType = objRs(0)
    'End If
    'objRs.Close
end if
%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">
<form method="POST" name="formname" action="VouOtherAppPURAdvance.asp">
<input type="hidden" name="hPara" value="<%=sPara%>">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="04">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hTmDate" value="<%=sTmDate%>">
<input type="hidden" name="hPurDate" value="<%=sPurDate%>">
<input type="hidden" name="hCurrDate" value="<%=FormatDate(Date)%>">
<input type="hidden" name="hSelDate" value="">
<input type="hidden" name="selApplication" value="2">
<input type="hidden" name="selVoucher" value="04">
<input type="hidden" name="hSelInvDate" value="">
<input type="hidden" name="hInvNo" value="<%=iInvNo%>">
<input type="hidden" name="hRcptNo" value="<%=iRcptNo%>">
<input type="hidden" name="hItemType" value ="<%=sItemType%>">
<input type="hidden" name="hPassDate" value ="<%=sVouDate%>">
<input type="hidden" name="hFlagForParAccHead" value="<%=sFlagPartyAccHead%>">
<input type="hidden" name="hSubTypeName" value="<%=sSubTypeName%>">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Voucher&nbsp;
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
								<td align="center" width="5" class="ClearPixel">&nbsp;</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly" width="100%">
                                <tr>
                                    <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                
                                <tr>
									<td class="FieldCellSub">Select Book</td>
									<td class="FieldcellSub">
										<select size="1" name="selBook" class="FormElem">
											<option value="S">Select Book</option>
												<%
												'sQuery="select BookNumber,OtherUnitTransaction,BookName from "&_
												'	"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='04'"
														
												sQuery="select BookNumber,OtherUnitTransaction,BookName from "&_
													"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='04' and BookAccountHead = "&iBookAccHead
												Response.Write sQuery  
												with objRs
													.CursorLocation = 3
													.CursorType = 3
													.Source = sQuery
													.ActiveConnection = con
													.Open
												end with
												set objRs.ActiveConnection = nothing
												IF Not Objrs.EOF Then
													while not objRs.EOF
														Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(2)&"</option>"
														objRs.MoveNext
													wend
												Else
													objRs.Close
													sQuery="select BookNumber,OtherUnitTransaction,BookName from "&_
															"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='04' "
													with objRs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sQuery
														.ActiveConnection = con
														.Open
													end with
													set objRs.ActiveConnection = nothing
													while not objRs.EOF
														Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(2)&"</option>"
														objRs.MoveNext
													wend
												End IF
												objRs.Close
											%>
										</select>
									</td>
									<td class="FieldCell" >Account Date</td>
									<td class="FieldCellSub"><%Response.Write InsertDatePicker("ctlAccDate")%></td>
                                </tr>
                                
                                <tr>
                            <!--<td class="FieldCell" >Unit </td>
                            <td width="160" class="FieldCellSub"><span class="DataOnly"> <%=sOrgName%> </span></td>-->
                            <td class="FieldCellSub">Bill Type</td>
                            <td class="FieldCellSub" width="160">
                            <%IF CStr(sPurBillTy) = "P" Then %>
								<span class="DataOnly">Credit Purchase </span></td>
							<%Elseif CStr(sPurBillTy) = "C" Then %>	
								<span class="DataOnly">Cash Purchase </span></td>
							<%Else%>
								<span class="DataOnly"></span></td>
							<%ENd IF %>
							
                            <td class="FieldCell">Passing Date</td>
                            <td class="FieldCellSub" width="160">	<span class="DataOnly"><%=sVouDate%> </span></td>
                                </tr>
                                <tr>
                            <td class="FieldCellSub">Party Name </td>
                            <td width="320" class="FieldCellSub" colspan="3"><span class="DataOnly"> <%=sPartyName%> </span>
                            </td>
                                </tr>

                                <tr>
                            <td class="FieldCellSub">Invoice No. - Date </td>
                            <td width="160" class="FieldCellSub" colspan="1"><span class="DataOnly" > <%=sRefernceNo%> </span>
                            &nbsp;
                            <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" width="11" alt="Modify/Verify Invoice" height="11" alt="" style="cursor:hand" onClick="AmdInvPurInvoice()">
                            
                            </td>
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
                <div class="frmBody" id="frm2" style="width: 590; height:242;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description / Account Head</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Quantity</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Rate</td>
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

	iSno = 1
	Dim oNodAcc,sTemp,iCtr,iCount,iTaxCtr,oNodTax
	sTemp = Split(iPurTaxTy,":")
	For iCtr = 0 To UBound(sTemp)
		IF Cstr(sTemp(iCtr)) <> "0" Then
			sExp = "//Details/Entry[@PurType="""&sTemp(iCtr)&"""]"
			'Response.Write sExp
			Set oNodEntry = oNodRoot.selectNodes(sExp)
			
			
			For iCount = 0 To oNodEntry.length - 1
			'Response.Write oNodEntry.Length
			'For Each oNodEntry in oNodDeatils.childNodes
				'iSno=oNodEntry.Item(iCount).Attributes.Item(0).nodeValue
				sDescription=oNodEntry.Item(iCount).Attributes.Item(1).nodeValue
				sAmount=oNodEntry.Item(iCount).Attributes.Item(2).nodeValue
				sRate=oNodEntry.Item(iCount).Attributes.Item(6).nodeValue
				sQty=oNodEntry.Item(iCount).Attributes.Item(3).nodeValue &"&nbsp;"&oNodEntry.Item(iCount).Attributes.Item(4).nodeValue
				sValue=oNodEntry.Item(iCount).Attributes.Item(7).nodeValue
				sDiscPer=oNodEntry.Item(iCount).Attributes.Item(8).nodeValue
				sDiscount=oNodEntry.Item(iCount).Attributes.Item(9).nodeValue

				dTotal=CDbl(dTotal)+CDbl(sAmount)
				dBasicTotal=CDbl(dBasicTotal)+CDbl(sValue)
				dDisTotal=CDbl(dDisTotal)+CDbl(sDiscount)

				For Each oNodAcc in oNodEntry.Item(iCount).childNodes
					IF oNodAcc.nodeName = "AccHead" Then
						iAccHead = oNodAcc.Attributes.Item(0).nodeValue
					End IF
				Next

				IF CStr(iAccHead) <> "" Then
					sQuery = "Select AccountDescription From VwAccheadforPurchaseApp Where AccountHead = "&iAccHead&" "&_
							 "and OUDefinitionID = '"&sOrgId&"' "
					objRs.Open sQuery,Con
					IF Not objRs.EOF Then
						IF CStr(sAccCheck) <> "N" Then
							sAccCheck = "Y"
						Else
							sAccCheck = "N"
						End IF
						sAccHeadName = objRs(0)
					Else
						sAccCheck = "N"
						sAccHeadName = ""
					End IF
					objRs.Close
				End IF

%>
		 <tr>
			<td class="ExcelSerial" align="center"><%=isno%></td>
			<td class="ExcelDisplayCell"><%=sDescription%> / <%=sAccHeadName%></td>
			<td class="ExcelDisplayCell" align="Left" width="60"><%=sQty%></td>
			<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sRate,5,,,0)%></td>
			<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sValue,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="Right" width="25"><%=FormatNumber(sDiscPer ,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber( sDiscount,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sAmount,2,,,0)%></td>
         </tr>
<%
				iSno = iSno + 1
			next 'For Group of Items
			
			'Response.Write "<br>"
			sExp = "//TaxDetails[@PurchaseType="""&sTemp(iCtr)&"""]"
			'Response.Write sExp & " " 
			
			Set oNodTaxRoot = oNodRoot.selectNodes(sExp)
			'Response.Write oNodTaxRoot.Length
			'Response.Write "<p><font color=red>Length="&oNodTaxRoot.Length
			For iTaxCtr = 0 To oNodTaxRoot.length - 1
				For Each oNodTax in oNodTaxRoot.Item(iTaxCtr).childNodes
					
					IF cint(oNodTax.Attributes.Item(6).nodeValue ) >0 and CStr(oNodTax.Attributes.Item(0).nodeValue ) <> "0" and CStr(oNodTax.Attributes.Item(2).nodeValue ) <> "0" then
						sCatCode=oNodTax.Attributes.Item(0).nodeValue
						sTaxCode=oNodTax.Attributes.Item(1).nodeValue
						sTaxMode=oNodTax.Attributes.Item(2).nodeValue
						sFormula=oNodTax.Attributes.Item(3).nodeValue
						dTaxValue=oNodTax.Attributes.Item(4).nodeValue
						dTax=oNodTax.Attributes.Item(5).nodeValue
						iAccHead = oNodTax.Attributes.Item(6).nodeValue
						sTaxName=oNodTax.Text
						
						'Response.Write sTemp(iCtr)&" " & iAccHead

						IF CStr(iAccHead) <> "" Then
							sQuery = "Select AccountHeadCode From VwAccheadforPurchaseApp Where AccountHead = "&iAccHead&" "&_
									 "and OUDefinitionID = '"&sOrgId&"' "
							objRs.Open sQuery,Con
							IF Not objRs.EOF Then
								sAccHeadName = objRs(0)
							Else
								sAccHeadName = ""
							End IF
							objRs.Close
						End IF



			%>
						<tr>
							<td align="center" colspan="3"></td>
							<td class="ExcelSerial" align="center" colspan="3"><%=sTaxName%>&nbsp;(<%=sAccHeadName%>)</td>
							<%if sTaxMode="F" then %>
								<td class="ExcelDisplayCell" align="right"></td>
							<%else%>
								<%if dTaxValue > 0 then %>
									<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTaxValue,2,,,0)%>&nbsp;%</td>
								<%else%>
									<td class="ExcelDisplayCell" align="right"></td>
								<%end if%>
							<%end if%>
							<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTax,2,,,0)%></td>
						</tr>
			<%
					end	if
				Next
			Next
			
	end if 'To Check Purchase Type is Not Zero
next 'List Of Distinct Purchase Type
sExp = "//TaxDetails/Tax[@CatCode=0 and @TaxCode=0]"
Set TempNode = oNodRoot.selectNodes(sExp)
'Response.Write "<p><font color=red>Length="&TempNode.length
IF CStr(iAccHead) = "" Then
	iAccHead = 0
End IF
IF TempNode.length > 0 Then
	IF CStr(TempNode.Item(0).Attributes.getNamedItem("AccHead").Value) <> "" Then
		sQuery = "Select AccountHeadCode From VwAccheadforPurchaseApp Where AccountHead = "&iAccHead&" "&_
				 "and OUDefinitionID = '"&sOrgId&"' "
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			sAccHeadName = objRs(0)
		Else
			sAccHeadName = ""
		End IF
		objRs.Close
	End IF
%>
	<tr>
		<!--<td align="center" colspan="3"></td>-->
		<td class="ExcelSerial" align="Right" colspan="6"><%=TempNode.Item(0).Text%>&nbsp;(<%=sAccHeadName%>)</td>
		<%if TempNode.Item(0).Attributes.getNamedItem("TaxValue").Value > 0 then %>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dTaxValue,2,,,0)%>&nbsp;</td>
			<%else%>
				<td class="ExcelDisplayCell" align="right"></td>
			<%end if%>
		
		<td class="ExcelDisplayCell" align="right"><%=FormatNumber(TempNode.Item(0).Attributes.getNamedItem("TaxValue").Value,2,,,0)%></td>
	</tr>
<%
			
End IF
%>

        <tr>
    <!--<td align="center" colspan="3"></td>-->
    <td class="ExcelSerial" align="center" colspan="4"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dBasicTotal,2,,,0)%></b></td>
    <td class="ExcelDisplayCell" align="center" width="25">    </td>
    <td class="ExcelDisplayCell" align="right" width="60"><b><%=FormatNumber(dDisTotal,2,,,0)%></b></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
        </tr>
        <input type="Hidden" name="hBasicValue" value="<%=dBasicTotal%>">
        <input type="Hidden" name="hDisValue" value="<%=dDisTotal%>">
        <input type="Hidden" name="hAmount" value="<%=dTotal%>">
<%
'
'
''	dInvAmount=oNodTaxRoot.Attributes.Item(0).nodeValue
'--------------------- Changed on 09/06/2004 - S Suresh	-----------------
'	sExp = "//TaxDetails[@RoundOff]"
'	Set TempNode = oNodRoot.selectNodes(sExp)
'	IF TempNode.length <> 0 Then
'		'dInvAmount=cdbl(oNodTaxRoot.Item(0).Attributes.Item(0).nodeValue) + cdbl(oNodTaxRoot.Item(0).Attributes.Item(3).nodeValue)
'		dInvAmount=cdbl(oNodTaxRoot.Item(0).Attributes.Item(0).nodeValue)
'	Else
'		dInvAmount=cdbl(oNodTaxRoot.Item(0).Attributes.Item(0).nodeValue)
'	End IF
'--------------------- End
%>
	

        <tr>
        <!--<td align="center" colspan="3"></td>-->
    <td class="ExcelSerial" align="right" colspan="7"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <%=FormatNumber(dInvAmount,2,,,0)%> </td>

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
                                    <span class="DataOnly"><%=AmountWords(dInvAmount)%></span>
                            </td>
                                    </table>

								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
							</tr>
							<% IF CStr(sChkMulAcc) <> "1" Then %>
							<!--Blocked by ragav on Oct 06,2011 for Misc Payment Adjustment -->
                            <!--        <tr>

                            <td colspan="2" align="center" class="FieldCell">
                               <font color=red>Multiply Accounthead is Selected Accounting Not Allowed</font>
                            </td>
                            </tr>-->
                            <%End IF %>
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
															<%' IF CStr(sChkMulAcc) <> "1" Then %>
                                                               <!-- <input type="button" value="<%=sButVal%>" name="btnAction" onclick="checkSubmit()" class="ActionButton" disabled>-->
                                                            <%'Else%>
																 <input type="button" value="<%=sButVal%>" name="btnAction" onclick="checkSubmit()" class="ActionButton">
															<%'End IF %>
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
