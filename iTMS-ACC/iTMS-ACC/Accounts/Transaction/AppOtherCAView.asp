
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppOtherCAView.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 28,2003
	'Modified By				:	S.Maheswari
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
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,iAppLevel,bOtherUnit
dim objRs,sQuery,sNarr
dim iVouNo,sVouDate,dAmount,sNarration
dim sAccount,sAddtional,iSno,dTotal,iTransNo
dim iBookCode,sPayTo,iBkAccHead,dBkClosingBal,sButVal
dim sExp,tempNode,sRetVal,sRefernceNo,sPurBillTy,sPara,sAccFlag

'XML DOM Variables
Dim oDOM,nodHeader,Root,Entrynode,HeaderNode

'Response.Write "<p><font color=red>asjhkdjvnfkn"
'Response.Write "<p><font color=red>sACTN="&Session("ACTN")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")

sTransNo=Request.QueryString ("TransNo")
iAppLevel=Request.QueryString ("AppLevel")
sPara = Request.QueryString ("sPara")

'Response.Write "selUser="&Request("hFormVal")
Session("ReDirVal") = Request("hFormVal")

'Response.Write "selUser="&Request("hUserID")
'Response.End
IF trim(sPara) = "Edt" then sButVal = "Edit"
IF trim(sPara) = "Acc" then sButVal = "Account"
IF trim(sPara) = "App" then sButVal = "Approve"
'Response.Write "sTransNo="&sTransNo  &"***"&sPara
'Response.End
 'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
sRetVal = GetVouchXML(sTransNo)
'Response.Write sRetVal
oDOM.Load server.MapPath(sRetVal)

sQuery = " Select InvoiceCode,convert(varchar,InvoiceDate,103)  from RCV_T_INVOICEHEADER where "&_
		 " InvoiceNumber = (Select OtherApplnTransNo from acc_t_CreatedVoucherheader where createdtransno = "& sTransNo &" )"
objrs.Open sQuery,con
if not objRs.EOF then
	sNarr = objRs(0) &" Dt:"&objRs(1)
end if
objRs.Close
set Root=oDOM.documentElement
sOrgId=Root.Attributes.GetNamedItem("UnitNo").value
sOrgName=Root.Attributes.GetNamedItem("UnitName").value
iBookCode =Root.Attributes.GetNamedItem("BookNo").value
sBookName =Root.Attributes.GetNamedItem("BookName").value
sVouType=Root.Attributes.GetNamedItem("CRDR").value
sVouDate=Root.Attributes.GetNamedItem("VouDate").value
iBkAccHead =Root.Attributes.GetNamedItem("BookAcchead").value
iVouNo =Root.Attributes.GetNamedItem("VoucherNo").value
iTransNo=Root.Attributes.GetNamedItem("TransNo").value


sExp ="//voucher/Entry[@No = 1]"
Set tempNode = Root.Selectnodes(sExp)
sPayTo = tempNode.Item(0).Attributes.getNamedItem("Payto").Value

'Newly aded to display Invoice no and date by Maheswari on June 2nd 2008
Dim sQry,rs1,rs2,rs3,iBkCode,sFrmApp,sCrAgain,sArr,sPurInvNo,sPurInvDate,sApprover,sAppTrNo
Set rs1= server.CreateObject("ADODB.Recordset")
Set rs2= server.CreateObject("ADODB.Recordset")
Set rs3= server.CreateObject("ADODB.Recordset")
sQry = "Select  BookCode From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "& sTransNo
rs1.Open sQry,con
if not rs1.EOF then
	iBkCode = rs1(0)
end if
rs1.Close

sQry = " Select PayToRecdFrom,CreatedBy,isNull(FromApplication,0),isNull(OtherApplnTransno,0),BookCode,isNull(PayableAt,'0'),isNull(AccountHead,0) AccountHead from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iTransNo &" "&_
		" and BookCode = "& iBkCode &" "
'	 Response.Write "qry="& sQry
rs2.Open sQry,con

If not rs2.EOF then
	sFrmApp = rs2(2)
	iBkCode = rs2(4)
	sCrAgain = rs2(5)
	iBkAccHead = rs2("AccountHead")
	 'Response.Write "frmapp="& sFrmApp &"--"&iBkCode & "<BR><BR>"
	If trim(sFrmApp) = "0" and iBkCode = "04" then
		sTempInvNo = rs2(0)

		sArr = split(sTempInvNo,"-")
		sPurInvNo = sArr(0)
		sPurInvDate = sArr(1)
		sApprover = rs2(1)
		'Response.Write "sPurInvNo="& sPurInvNo & "<BR><BR>"
	else
		sAppTrNo = rs2(3)
		'Response.Write "sAppTrNo="& sAppTrNo
		with rs3
			.CursorLocation = 3
			.CursorType = 3
			.Source = "Select SuppInvoiceNo,SuppInvoiceDate From RCV_T_INVOICEHEADER Where InvoiceNumber = "&sAppTrNo&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs3.EOF then
			sPurInvNo = rs3(0)
			sPurInvDate = rs3(1)
			'Response.Write "sPurInvNo="& sPurInvNo & "<BR><BR>"
		End If
		rs3.Close
	End If	'If trim(sFrmApp) = "Null" then
End If 'If not rs2.EOF then
rs2.Close
sRefernceNo = sPurInvNo &" Dt:"&sPurInvDate


IF CStr(iBkAccHead) = "" Then
	iBkAccHead = 0
End IF

sQuery = "Select AllowTransactions From VwOrgGLHeads Where AccountHead = "&iBkAccHead
'Response.Write sQuery
rs2.Open sQuery,con
If not rs2.EOF then
	sAccFlag = rs2("AllowTransactions")
Else
	sAccFlag = "A"
End IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Cash Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
function finaldone() {
	var selected;
	if (document.formname.selBook.selectedIndex < 1) {
		alert("Select a Book");
		document.formname.selBook.focus();
		return false;
	}
	document.formname.hAccDate.value = document.formname.ctlAccDate.getDate();
	selected = document.formname.selBook.options[document.formname.selBook.selectedIndex];
	document.formname.hBookName.value = selected ? selected.text : "";
	document.formname.submit();
	return true;
}

function finalCancel() {
	document.formname.action = "AppBookSelection.asp";
	document.formname.submit();
}

function SetDate() {
	document.formname.ctlAccDate.setDate(document.formname.hDate.value);
}

function AmdInvPurInvoice() {
	var orgId = document.formname.hOrgId.value;
	var invNo = document.formname.hInvNo.value;
	var rcptNo = document.formname.hRcptNo.value;
	var itemType = document.formname.hItemType.value;
	document.formname.action = "../../Purchase/TRANSACTION/AmdInvPurInvoiceEntry.asp?ForUnit=" + encodeURIComponent(orgId) + "&InvNo=" + encodeURIComponent(invNo) + "&hRcptNo=" + encodeURIComponent(rcptNo) + "&ItemType=" + encodeURIComponent(itemType);
	document.formname.submit();
}
</script>
<%
'Response.Write "<p><font color=red>"
Dim iInvNo,iRcptNo,sItemType
sQuery = "Select isNull(OtherApplnTransno,0) from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iTransNo
'Response.Write sQuery
objRs.Open sQuery,con
If not objRs.EOF then
	iInvNo = objRs(0)
End If
objRs.Close

sQuery = "select Distinct isNull(Receiptnumber,0) from Pur_t_refferencenumberDet where invoicenumber = "& iInvNo
'Response.Write "<p>Query="&sQuery
objRs.Open sQuery,con
do while not objRs.EOF
	iRcptNo = iRcptNo &","& objRs(0)
objRs.MoveNext
loop
objRs.Close
iRcptNo = mid(iRcptNo,2)
'Response.Write "<p font color=red>iRcptNo="&iRcptNo
'sQuery = "Select ItemType from Rcv_T_ActualReceiptheader where receiptnumber in ("&iRcptNo&") "
'objRs.Open sQuery,con
'If not objRs.EOF then
'	sItemType = objRs(0)
'End If
'objRs.Close
%>

<script src="/Scripts/itms-modern-compat.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="SetDate()">

<form method="POST" name="formname" action="AppOtherCashGenerate.asp">
<input type="hidden" name="hPara" value="<%=sPara%>">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="01">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hTransNo" value="<%=sTransNo%>">
<input type="hidden" name="hDate" value="<%=sVouDate%>">
<input type="hidden" name="hBookAccHead" value="<%=iBkAccHead%>">
<input type="hidden" name="hActionFlag" value="A">
<input type="hidden" name="hNarration" value="<%=sNarr%>">
<input type="hidden" name="hInvNo" value="<%=iInvNo%>">
<input type="hidden" name="hRcptNo" value="<%=iRcptNo%>">
<input type="hidden" name="hItemType" value ="<%=sItemType%>">
<input type="hidden" name="hAccDate" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%
		if sVouType="C" then
			Response.Write "Cash Payment Voucher"
		else
			Response.Write "Cash Receipt Voucher"
		end if
		%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="125">
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
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								  	<tr>
								  		<td align="center">Voucher</td>
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
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
                            <td class="FieldCell" width="145">Select Book</td>
                            <td>
										<select size="1" name="selBook" class="FormElem">
											<option value="S">Select Book</option>
<%

	sQuery="select BookNumber,OtherUnitTransaction,BookName from "&_
		"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='01'"
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	while not objRs.EOF
		IF trim(iBookCode) = trim(objRs(0)) then
			Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(2)&"</option>"
		Else
			Response.Write "<option value="""&objRs(0)&""">"&objRs(2)& "</option>"
		End IF
		objRs.MoveNext
	wend
	objRs.Close
%>
										</select>
                            </td>
                            <td class="FieldCell" width="88"></td>
                                </tr>
                                 <tr>
                            <td class="FieldCell" width="145">Account Date</td>
                            <td>
                             <%Response.Write InsertDatePicker("ctlAccDate")%>
                            </td>
                            <td class="FieldCell" width="88"></td>
                                </tr>
                            <tr>
									<td class="FieldCell"><p align="left">Accounting Unit</td>
									<td class="FieldCell" colspan="3">
										<select size="1" name="selUnitId"  class="FormElem">
										<OPTION value="0">Select a Unit</option>
									    <%populateUnitSelected(sorgID)%>
									</select>
									</td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
                            <td class="FieldCell">Unit </td>
                            <td width="160" class="FieldCell"><span class="DataOnly"> <%=sOrgName%> </span></td>
                            <td class="FieldCell">Passing Date</td>
                            <td class="FieldCell" width="160">	<span class="DataOnly"><%=sVouDate%> </span></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="145">
                             <%
								if sVouType="C" then
									Response.Write "Pay To"
								else
									Response.Write "Received From"
								end if
                            %>

                            </td>
                            <td colspan="3">
								<span class="DataOnly"><%=sPayTo%> </span>
                            </td>

                                </tr>
                                   <tr>
                            <td class="FieldCell">Invoice No. - Date </td>
                            <td width="160" class="FieldCell" colspan="1"><span class="DataOnly"> <%=sRefernceNo%> </span> &nbsp;
                            <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" width="11" alt="Modify/Verify Invoice" height="11" alt="" style="cursor: pointer" onClick="AmdInvPurInvoice()">
                            <td class="FieldCell">Bill Type</td>
                            <td class="FieldCell" width="160">
                            <%IF CStr(sVouType) = "D" Then %>
								<span class="DataOnly">Credit Purchase </span></td>
							<%Elseif CStr(sVouType) = "C" Then %>
								<span class="DataOnly">Cash Purchase </span></td>
							<%Else%>
								<span class="DataOnly"></span></td>
							<%ENd IF %>
                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell" height="20">
												<DIV class=frmBody id=frm1 style="width:600; height:200;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="583">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="75">AU</td>
                                        <td class="ExcelHeaderCell" align="center">Item Desc [Account Code - Name]/ Narration&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center" width="110">Additional Details</td>
                                        <td class="ExcelHeaderCell" align="center" width="25">Dr/Cr</td>
                                        <td class="ExcelHeaderCell" align="center" width="120" >Amount</td>
                                            </tr>
<%
dim Node,Node1,Node2
dim sTemp,sAccType,nodADD,dCRAmt,dDRAmt,sQuery1,sItemDesc
dim objrs1,sSql,iAccHead,sAccHeadName,iTaxCode,iCatCode,iTaxAmt,sCrDr,sTaxName,sTotAmt
Set objrs1 = Server.CreateObject("ADODB.RecordSet")
sQuery1 = "Select AccUnitAccountHead,Amount,TransCrDrIndication,ItemDescription from acc_t_CreatedVoucherdetails where createdtransno = "&sTransNo&" and Amount > 0 "
'Response.Write sQuery1
objRs.Open sQuery1,con
iSno = 1
dCRAmt = 0
sTotAmt = 0
Do while not objRs.EOF
	iAccHead = objRs(0)
	iTaxAmt  = objRs(1)
	sCrDr	 = objRs(2)
	sItemDesc = objRs(3)
	If trim(sCrDr) = "C" then
		dCRAmt=dCRAmt+cint(iTaxAmt)
	End If
	sSql = "Select AccountDescription from Acc_M_GLAccountHead where Accounthead = "& iAccHead
	objrs1.Open sSql,con
	if not objrs1.EOF then
		sAccHeadName = objrs1(0)
	end if
	objrs1.Close
	sAccount = sItemDesc &" ["& sAccHeadName  &"]"
	dAmount=FormatNumber(iTaxAmt,2,,,0)

'Response.Write dAmount  &"***<BR>"
If trim(sCrDr) <> "C" then
	'sTotAmt =cint(sTotAmt) + cint(dAmount)
	sTotAmt =cdbl(sTotAmt) + cdbl(dAmount)
End IF


%>
                <tr>
            <td class="ExcelSerial" align="center" rowspan="2"><%=iSno%></td>
            <td class="ExcelDisplayCell" rowspan="2"><%=sOrgName%></td>
            <td class="ExcelDisplayCell" align="right"><p align="left"><%=sAccount%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><p align="left"><%=sAddtional%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><p align="left"><%=sCrDr%>R</p>
            </td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=dAmount%></td>
                </tr>
            <tr>
				 <td class="ExcelDisplayCell" align="right"><p align="left"><%=sNarration%></td>
            </tr>

<%
iSno = iSno + 1
objRs.MoveNext
loop
objRs.Close

sQuery = "Select AccountHead,TaxCode,TaxCategoryCode,TaxAmount,TransCrDrIndication from acc_t_CreatedVouchertaxdet where createdtransno = "&sTransNo&" and TaxAmount > 0 "
objRs.Open sQuery,con

Do while not objRs.EOF
	iAccHead = objRs(0)
	iTaxCode = objRs(1)
	iCatCode = objRs(2)
	iTaxAmt  = objRs(3)
	sCrDr	 = objRs(4)
	If trim(sCrDr) = "C" then
		dCRAmt=dCRAmt+cint(iTaxAmt)
	End If
	sSql = "Select AccountDescription from Acc_M_GLAccountHead where Accounthead = "& iAccHead
	objrs1.Open sSql,con
	if not objrs1.EOF then
		sAccHeadName = objrs1(0)
	end if
	objrs1.Close
	sSql = "Select Distinct TaxName from vwPurchasetaxdetails where TaxCode = "&iTaxCode&" and TaxCategoryCode = "&iCatCode
	objrs1.Open sSql,con
	if not objrs1.EOF then
		sTaxName = objrs1(0)
	end if
	objrs1.Close
	sAccount = sAccHeadName  &" ["& sTaxName &"]"

	dAmount=FormatNumber(iTaxAmt,2,,,0)
	'Response.Write dAmount  &"***<BR>"
	If trim(sCrDr) <> "C" then
		'sTotAmt =cint(sTotAmt) + cint(dAmount)
		sTotAmt =cdbl(sTotAmt) + cdbl(dAmount)
	End IF
	set Root=oDOM.documentElement
	Set Node = oDOM.CreateElement("Entry")
	Node.setAttribute "No",iSno
	Node.setAttribute "CRDR",sCrDr
	Node.setAttribute "Payto",sPayTo
	Node.setAttribute "Amount",dAmount
	Node.setAttribute "AccUnit",sOrgId
	Node.setAttribute "AccName",sOrgName
	Node.setAttribute "TdsAmount","0"
	Node.setAttribute "TDSElgi","0"
	Node.setAttribute "TdsPercentage","0"
	Node.setAttribute "PayRecAmount","0"
	Node.setAttribute "GroupId","0"
	Root.appendchild Node
	Set Node1 = oDOM.CreateElement("AccHead")
	Node1.setAttribute "No",iAccHead
	Node1.setAttribute "CostCenter","0"
	Node1.setAttribute "Analytical","0"
	Node1.setAttribute "Name",sTaxName
	Node1.setAttribute "Type","G"
	Node1.setAttribute "TransFlag","A"
	Node.Appendchild Node1

	Set Node2 = oDOM.CreateElement("Narration")
	Node2.text = sNarr
	Node.Appendchild Node2


%>
                <tr>
            <td class="ExcelSerial" align="center" rowspan="2"><%=iSno%></td>
            <td class="ExcelDisplayCell" rowspan="2"><%=sOrgName%></td>
            <td class="ExcelDisplayCell" align="right"><p align="left"><%=sAccount%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><p align="left"><%=sAddtional%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><p align="left"><%=sCrDr%>R</p>
            </td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=dAmount%></td>
                </tr>
            <tr>
				 <td class="ExcelDisplayCell" align="right"><p align="left"><%=sNarration%></td>
            </tr>

<%

'	end if
'next'End of Voucher Node Loop

'	if StrComp(sVouType,"D")=0 then
'		dTotal=dCRAmt-dDRAmt
'	else
'		dTotal=dDRAmt-dCRAmt
'	end if
'dTotal=FormatNumber(dTotal,2,,,0)

iSno = iSno + 1
objRs.MoveNext
loop
objRs.Close

oDom.save server.MapPath("../temp/transaction/VoucherCBG_xml_"&Session.SessionID&".xml")
'Response.Write "sTotAmt = "& sTotAmt &"--"& dCRAmt

if cdbl(dCRAmt) <> 0 then
	dTotal=cdbl(sTotAmt) - cdbl(dCRAmt)
Else
	dTotal= sTotAmt
end if
dTotal=FormatNumber(dTotal,2,,,0)
%>
								<% If cint(iSno) > 1  then  %>

                                            <tr>
                                        <td class="ExcelDisplayCell" align="center" colspan="5"><p align="right"><b>Total&nbsp;</b></td>
                                        <td class="ExcelDisplayCell" align="right">Rs.&nbsp;<%=dTotal%></td>
                                            </tr>
                                  <%End If %>
                                                </table>
												</div>
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
                                &nbsp;
								</td>
								<td valign="top" class="FieldCell" height="20">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                    <% If cint(iSno) > 1  then  %>
									     <tr>
											<td class="FieldCell" width="130" valign="top">Amount </td>
											<td><span class="DataOnly"><%=AmountWords(dTotal)%></span></td>
									    </tr>
                                      <%End If %>
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
                                                                <input type="button" value="<%=sButVal%>" name="B7" onClick="finalDone('<%=sAccFlag%>')" class="ActionButton">
                                                                <!--input type="button" value="Keep Pending" name="btnAction2" onclick="finalCancel()" class="ActionButtonX"-->
                                                               </p>
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
</HTML>
