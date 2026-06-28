<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouBADisplay.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 26, 2002
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
<%

'XML DOM Variables
Dim oDOM,nodHeader,Root,objRs,sQuery
dim sNarration,sAccount,sAddtional,iSno,sAmount
dim dTotal,sOrgId,sUserId
dim EntryNode,HeaderNode,TempNode
dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim iTransNo,iBkHeadCode,bOtherUnit,sApproverName,sRedType,sRetVal
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request("TransNo")
sApprove = Request("Approver")

IF CStr(sApprove) <> "" Then
	sApproverName = GetEmployeeName(sApprove)
Else
	sApproverName = ""
End IF

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")

'Dim sRetVal
'This Call Construct The XML File From Database
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)




set Root=oDOM.documentElement
sOrgId=Root.Attributes.Item(0).nodeValue
sOrgName=Root.Attributes.Item(1).nodeValue
iBookCode =Root.Attributes.Item(2).nodeValue
sBookName =Root.Attributes.Item(3).nodeValue
sVouType=Root.Attributes.Item(4).nodeValue
sVoucDate=Root.Attributes.Item(5).nodeValue
iBkHeadCode=Root.Attributes.Item(6).nodeValue

iVouNo=Root.Attributes.Item(9).nodeValue
sApprove=Root.Attributes.Item(7).nodeValue

sUserId = Session("UserID")


set objRs = Server.CreateObject("ADODB.Recordset")

sQuery="select OtherUnitTransaction from vwOrgBookNames where OUDefinitionID = '" & sOrgId & "'"&_
"and BookNumber="&iBookCode&" and BookCode='02'"

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

if not 	objRs.EOF then
	bOtherUnit=objRs(0)
else
	bOtherUnit=0
end if
objRs.Close


For Each TempNode in Root.childNodes
	IF TempNode.nodeName = "Entry" Then
		set EntryNode = TempNode
		sPayTo= TempNode.Attributes.Item(2).nodeValue
	End IF
Next



set HeaderNode=EntryNode.childNodes(0)


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<xml id="InstData">
<BankInstrumentDet InsType="" InsNo="" InsDate="" Payableat="" Drawnon=""/>
</xml>
<script language="javascript">
function FinalCheck(flag) {
	var form = document.formname;
	var transNo = form.hTransNo.value;
	if (flag === "B") {
		form.action = "VouBABookSelection.asp";
		form.submit();
	} else if (flag === "PV") {
		form.action = "BankVoucher.asp?VouTy=P";
		form.submit();
	} else if (flag === "RV") {
		form.action = "BankVoucher.asp?VouTy=R";
		form.submit();
	} else if (flag === "P") {
		PrintWindow("PRNBankRecpVouView1.asp?Value=" + encodeURIComponent(transNo) + "&CallF=D");
	}
}

function setInstruDate(sDate) {
	var control = document.formname.ctlDate || document.getElementById("ctlDate");
	if (control && typeof control.setDate === "function") {
		control.setDate(sDate);
	} else if (control && typeof control.SetDate === "function") {
		control.SetDate(sDate);
	}
}

function SetFocus() {
	if (document.formname.B10) {
		document.formname.B10.focus();
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetFocus()">
<form method="POST" name="formname" action="">
<input type="hidden" name="selUnitId" value="<%=sOrgId%>">
<input type="hidden" name="horgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="selBook" value="<%=iBookCode%>">
<input type="hidden" name="selVouType" value="<%=sVouType%>">
<input type="hidden" name="hBookOtherUnit" value="<%=bOtherUnit%>">
<input type="hidden" name="hBookAccHead" value="<%=iBkHeadCode%>">

<input type="hidden" name="hApprover" value="<%=sApprove%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		  <%
			if sVouType="C" then
				Response.Write "Bank Payment Voucher"
				sRedType = "PV"
			else
				Response.Write "Bank Receipt Voucher"
				sRedType = "RV"
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
                                <tr>
                            <td class="FieldCell" width="145"> Book Code </td>
                            <td width="210">
                                                            <span class="DataOnly"><%=iBookCode%></span>
                            </td>
                            <td class="FieldCell" width="88">
                            Voucher No.</td>
                            <td>
                                                            <span class="DataOnly"><%=iVouNo%></span>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="145">Book Name</td>
                            <td width="210">
                                                            <span class="DataOnly"><%=sBookName%></span>
                            </td>
                            <td class="FieldCell" width="88">
                            Voucher Date</td>
                            <td>
								<span class="DataOnly"><%=sVoucDate%></span>
                            </td>
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
                            <td width="210">
								<span class="DataOnly"><%=sPayTo%></span>
                            </td>
                            <td class="FieldCell" width="88" align="left">
                                                            Unit
                            </td>
                            <td>
                                <span class="DataOnly"><%=sOrgName%></span>
                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <!--tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                        <tr>
                                                    <td class="FieldCellSub" width="118">Instrument Type</td>
                                                    <td colspan="4" class="FieldCell">
                                                        <input type="radio" name="optInsType" value="Cheque" checked class="FormElem">
                                                        Cheque&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                        <input type="radio" name="optInsType" value="Demand Draft" class="FormElem">
                                                        Demand Draft&nbsp;<br>
                                                        <input type="radio" name="optInsType" value="Bankers Cheque" class="FormElem">
                                                        Bankers
                                                        Cheque&nbsp;&nbsp; <input type="radio" name="optInsType" value="Telegraphic Transfer" class="FormElem">
                                                        Telegraphic Transfer
 </td>
                                                    <td class="FieldCellSub" width="5"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="118">Instrument Number</td>
                                                    <td>
 <input type="text" name="txtInstNo" size="20" class="Formelem">
 </td>
                                                    <td class="FieldCellSub" width="10"></td>
                                                    <td class="FieldCellSub" width="67">Payable at</td>
                                                    <td>
 <input type="text" name="txtPayableAt" size="40" class="Formelem">
 </td>
                                                    <td class="FieldCellSub" width="5"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="MiddlePack" height="2"></td>
                                                    <td height="2" class="MiddlePack">
 </td>
                                                    <td class="MiddlePack" height="2"></td>
                                                    <td class="MiddlePack" height="2"></td>
                                                    <td height="2" class="MiddlePack">
 </td>
                                                    <td class="MiddlePack" height="2"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="118">Instrument Date</td>
                                                    <td class="FieldCellSub" width="125">
                                                   <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>  </td>
                                                    <td class="FieldCellSub" width="10"></td>
                                                    <td class="FieldCellSub" width="67">Drawn on</td>
                                                    <td>
 <input type="text" name="txtDrawnOn" size="40" class="Formelem">
 </td>
                                                    <td class="FieldCellSub" width="5"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr-->
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
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
												<DIV class=frmBody id=frm1 style="width:600; height:140;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="680">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="75">AU</td>
                                        <td class="ExcelHeaderCell" align="center">Account Code - Name / Narration&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center" width="110">Additional Details</td>
                                        <td class="ExcelHeaderCell" align="center" width="25">Dr/Cr</td>
                                        <td class="ExcelHeaderCell" align="center" width="80" >Amount</td>
                                        <td class="ExcelHeaderCell" align="center" width="80" >Deduction Amount</td>
                                        <td class="ExcelHeaderCell" align="center" width="80" >Deduction Percentage</td>
                                            </tr>
<%
dim sTemp,sAccType,nodADD,dCRAmt,dDRAmt,sTdsAmt,sTdsPer

for each EntryNode in Root.childNodes
	IF EntryNode.nodeName = "Entry" Then
		iSno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(3).nodeValue
		sOrgName=EntryNode.Attributes.Item(5).nodeValue
		sTdsAmt = EntryNode.Attributes.Item(6).nodeValue
		sTdsPer = EntryNode.Attributes.Item(8).nodeValue



		if EntryNode.Attributes.Item(1).nodeValue="C" then
			dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		else
			dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		end if

		sAddtional=""
		for each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					sAccType=HeaderNode.Attributes.Item(4).nodeValue
					if sAccType="G" then
						sAccount=HeaderNode.Attributes.Item(0).nodeValue
						sAccount=sAccount& "-" & HeaderNode.Attributes.Item(3).nodeValue
					else
						sAccount=HeaderNode.Attributes.Item(3).nodeValue
					end if
			end if 'End of Check for Account head Node
			if 	HeaderNode.nodeName="Narration" then
					sNarration=HeaderNode.text
			end if 'End of Check for Narration Node
			if 	HeaderNode.nodeName="CostCenter" then
					for each  nodADD in HeaderNode.childNodes
						sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue&"-"
						sAddtional=sAddtional&nodADD.Attributes.Item(3).nodeValue &"%&nbsp;"
						sAddtional=sAddtional&nodADD.Attributes.Item(4).nodeValue&"<br>"
					next
			end if 'End of Check for Cost Center Node
			if 	HeaderNode.nodeName="Analytical" and HeaderNode.hasChildnodes then
					sAddtional=sAddtional&"---------------------------  <br>"
					for each  nodADD in HeaderNode.childNodes
						sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue&"-"
						sAddtional=sAddtional&nodADD.Attributes.Item(3).nodeValue &"%&nbsp;"
						sAddtional=sAddtional&nodADD.Attributes.Item(4).nodeValue&"<br>"
					next
			end if 'End of Check for Analytical Node
			if 	HeaderNode.nodeName="PayRec" then
					for each  nodADD in HeaderNode.childNodes
						sAddtional="Doc No:"
						sAddtional=sAddtional&nodADD.Attributes.Item(1).nodeValue&":"
						sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue &"-&nbsp;"
						sAddtional=sAddtional&nodADD.Attributes.Item(5).nodeValue&"<br>"
					next
			end if 'End of Check for PayRec Node

	next 'End of Entry Node Loop

		sAmount=FormatNumber(sAmount,2,,,0)
		sTdsAmt = FormatNumber(sTdsAmt,2,,,0)
		sTdsPer = FormatNumber(sTdsPer,2,,,0)
%>
                <tr>
            <td class="ExcelSerial" align="center" rowspan="2"><%=iSno%></td>
            <td class="ExcelDisplayCell" rowspan="2"><%=sOrgName%></td>
            <td class="ExcelDisplayCell" align="right"><p align="left"><%=sAccount%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><p align="left"><%=sAddtional%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><p align="left"><%=EntryNode.Attributes.Item(1).nodeValue%>R</p>
            </td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=sAmount%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=sTdsAmt%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=sTdsPer%></td>

                </tr>
            <tr>
				 <td class="ExcelDisplayCell" align="right"><p align="left"><%=sNarration%></td>
            </tr>

<%
End IF
next'End of Voucher Node Loop
	if StrComp(sVouType,"D")=0 then
		dTotal=dCRAmt-dDRAmt
	else
		dTotal=dDRAmt-dCRAmt
	end if
		dTotal=FormatNumber(dTotal,2,,,0)
%>
                                            <tr>
                                        <td class="ExcelDisplayCell" align="center" colspan="5"><p align="right"><b>Total&nbsp;</b></td>
                                        <td class="ExcelDisplayCell" align="right">Rs.&nbsp;<%=dTotal%></td>
                                        <td class="ExcelDisplayCell" align="right">&nbsp;</td>
                                        <td class="ExcelDisplayCell" align="right">&nbsp;</td>
                                            </tr>
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
                                <tr>
                            <td class="FieldCell" width="130" valign="top">Amount </td>
                            <td>
                                                               <span class="DataOnly"><%=AmountWords(dTotal)%></span>
                            </td>
                                </tr>
                                <% IF CStr(sApproverName) <> "" Then %>
                                <tr>
									<td class="FieldCell" width="130" valign="top">Approved By  </td>
									<td><span class="DataOnly"><%=sApproverName%></span></td>
								</tr>
                                <%End IF %>
<%
'if sApprove="Y" then
%>
                                <!--tr>
                            <td class="FieldCell" width="130">Immediate Approver</td>
                            <td class="FieldCell">
                            <select size="1" name="selUserId" class="FormElem">
                        <option value="I">Immediate Approver</option>
                        <%'=populateEmployeeWithVal(sUserId)%>
                            </select></td>
                                </tr-->
<%
'end if
%>
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
								<td valign="top" class="FieldCell" height="20">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
																<!--input type="button" value="Next Payment Voucher" name="B9" onClick="FinalCheck('PV')" class="ActionButtonX"  >
																<input type="button" value="Next Receipt Voucher" name="B10" onClick="FinalCheck('RV')" class="ActionButtonX"  -->
                                                                <!--input type="button" value="Next Book" name="B7" onClick="FinalCheck('B')" class="ActionButtonX" -->

                                                                <input type="button" value="Next Voucher" name="B10" onClick="FinalCheck('<%=sRedType%>')" class="ActionButtonX"  -->
																<input type="button" value="Print" name="B8" onClick="FinalCheck('P')" class="ActionButton" >
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
