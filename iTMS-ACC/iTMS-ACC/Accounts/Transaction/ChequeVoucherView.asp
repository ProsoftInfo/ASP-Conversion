<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ChequeVoucherView.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	September 15,2003
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
dim sOrgId,sBookCode,sBookName,sQuery,sTransNo
dim objRs,iVoucherNo

dim iBookCode,sPayTo,dTotal

dim sExp,tempNode
dim sInsType,sInsNo,sInsDate,sPayableat,sDrawnOn,bInstruFlag,sBankAddress,sACNo

'XML DOM Variables
Dim oDOM,nodHeader,Root

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")

sTransNo=Request("TransNo")
'Response.Write sTransNo

sOrgId = session("organizationcode")
iBookCode = Request("BookNo")

If 1 = 2 Then
	oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")

	set Root=oDOM.documentElement

	sOrgId=Root.Attributes.GetNamedItem("UnitNo").value
	iBookCode =Root.Attributes.GetNamedItem("BookNo").value
	sBookName =Root.Attributes.GetNamedItem("BookName").value


	sExp ="//voucher/Entry[@No = 1]"
	Set tempNode = Root.Selectnodes(sExp)
	sPayTo = tempNode.Item(0).Attributes.getNamedItem("Payto").Value
	dTotal= tempNode.Item(0).Attributes.getNamedItem("Amount").Value
	bInstruFlag=false
	sExp ="//voucher/BankInstrumentDet"
	Set tempNode = Root.Selectnodes(sExp)
	if tempNode.length > 0 then
		sInsType=tempNode.Item(0).Attributes.GetNamedItem("InsType").value
		sInsNo=tempNode.Item(0).Attributes.GetNamedItem("InsNo").value
		sInsDate=tempNode.Item(0).Attributes.GetNamedItem("InsDate").value
		sPayableat=tempNode.Item(0).Attributes.GetNamedItem("Payableat").value
		sDrawnOn=tempNode.Item(0).Attributes.GetNamedItem("Drawnon").value
		bInstruFlag=true
	end if

End IF	'If 1 = 2 Then

sQuery=" SELECT A.ACCOUNTHEAD,V.ACCOUNTDESCRIPTION,I.BANKINSTRUMENTNO, CONVERT(DATETIME,I.BANKINSTRUMENTDATE,103),"&_
			 " I.INSTRUMENTAMOUNT,A.PAYTORECDFROM,I.BANKINSTRUMENTTYPE,I.PAYABLEAT,I.DRAWNONBANK "&_
			 " FROM ACC_T_CREATEDVOUCHERHEADER A ,ACC_M_GLACCOUNTHEAD V,Acc_T_CreatedVoucherInstrumentDet I" & _
			 " WHERE A.ACCOUNTHEAD=V.ACCOUNTHEAD AND I.CREATEDTRANSNO = A.CREATEDTRANSNO AND A.BOOKCODE='02'"&_
			 " AND A.CREATEDTRANSNO IN ("& sTransNo &") AND A.BOOKNUMBER IN ("& Cstr(iBookCode) &" )"

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

set objRs.ActiveConnection = nothing

Do While Not objRs.EOF

	sBookName = objrs(1)
	sPayTo	= objrs(5)
	dTotal = objrs(4)
	sInsType= objRs(6)
	sInsNo=objrs(2)
	sInsDate=objrs(3)
	sPayableat=objRs(7)
	sDrawnOn=objRs(8)

	objRs.MoveNext
Loop
objRs.Close


sQuery=" select AccountNo,BankAddress1,isnull(BankAddress2,''),City,PinCode,State,BankName "&_
	" from Acc_M_BankDetails where OUDefinitionID = '" & sOrgId & "' and BookCode='02' and BookNumber IN ("&iBookCode&")"

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
while not objRs.EOF
	sACNo=objRs(0)
	sBankAddress=objRs(6)
	sBankAddress=sBankAddress&"<BR>"&objRs(1)

	if trim(objRs(2))<>"" then
		sBankAddress=sBankAddress&","&objRs(2)
	end if
	sBankAddress=sBankAddress&"<BR>"&objRs(3)&" - "&objRs(4)&" , "&objRs(5)
	objRs.MoveNext
wend
objRs.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Bank Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<script language="javascript">
function finaldone() {
	var chequeNo = document.formname.txtChequeNo.value;
	var transNo = document.formname.hTransNo.value;
	var total = document.formname.hTotal.value;
	var xhr;
	if (String(chequeNo) === "") {
		alert("Enter Cheque No ");
		document.formname.txtChequeNo.focus();
		return false;
	}
	xhr = new XMLHttpRequest();
	xhr.open("POST", "XMLChequeUpdate.asp?ChequeNo=" + encodeURIComponent(chequeNo) + "&TransNo=" + encodeURIComponent(transNo), false);
	xhr.send(null);
	if (String(xhr.responseText || "").replace(/^\s+|\s+$/g, "") !== "") {
		alert(xhr.responseText);
		return false;
	}
	PrintWindow("PRN_Cheque.asp?Value=" + encodeURIComponent(transNo) + "&Total=" + encodeURIComponent(total));
	return true;
}

function Submit() {
	document.formname.action = "CHEQUEPRINTING.ASP";
	document.formname.submit();
}
</script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">

<input type="hidden" name="hBookNo" value="<%=iBookCode%>">
<input type="hidden" name="hTransNo" value="<%=sTransNo%>">
<input type="hidden" name="hTotal" value="<%=dTotal%>">
<input type="hidden" name="selUnitId" value="<%=sOrgId%>">
<input type="hidden" name="selBook" value="<%=iBookCode%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">



<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="PageTitle" height="20"><p align="center">Cheque Print</td>
    </tr>
	<!--<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>-->
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<!--<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" align="center" width="96">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Instruments
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								  	<tr>
								  		<td align="center">Cheque Print</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>-->

				<tr>
					<td valign="top">
						<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
							<tr>
								<td align="center" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td class="TabBodyWithTopLine">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td colspan="3" align="center" width="5" class="ClearPixel">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<%
											sQuery=" SELECT A.ACCOUNTHEAD,V.ACCOUNTDESCRIPTION,I.BANKINSTRUMENTNO, CONVERT(DATETIME,I.BANKINSTRUMENTDATE,103),"&_
														 " I.INSTRUMENTAMOUNT,A.PAYTORECDFROM,I.BANKINSTRUMENTTYPE,I.PAYABLEAT,I.DRAWNONBANK "&_
														 " FROM ACC_T_CREATEDVOUCHERHEADER A ,ACC_M_GLACCOUNTHEAD V,Acc_T_CreatedVoucherInstrumentDet I" & _
														 " WHERE A.ACCOUNTHEAD=V.ACCOUNTHEAD AND I.CREATEDTRANSNO = A.CREATEDTRANSNO AND A.BOOKCODE='02'"&_
														 " AND A.CREATEDTRANSNO IN ("& sTransNo &") AND A.BOOKNUMBER IN ("& Cstr(iBookCode) &") "
											with objRs
												.CursorLocation = 3
												.CursorType = 3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with

											set objRs.ActiveConnection = nothing

											Do While Not objRs.EOF

												sBookName = objrs(1)
												sPayTo	= objrs(5)
												dTotal = objrs(4)
												sInsType= objRs(6)
												sInsNo=objrs(2)
												sInsDate=objrs(3)
												sPayableat=objRs(7)
												sDrawnOn=objRs(8)
										%>
										<tr>
											<td align="center" width="5" class="ClearPixel">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td class="BodyTable">
												<table border="0" cellpadding="0" cellspacing="0" width="100%" height="194">
													<tr>
														<td class="FieldCellSub" valign="top"></td>
														<td class="FieldCellSub" valign="top" height="18"></td>

														<td class="FieldCellSub" align="center" height="18" width="150"><span class="DataOnly"><%=sInsDate%></span></td>
												  	</tr>
												  	<tr>
														<td class="FieldCellSub" width="75" height="21">Pay</td>
														<td class="FieldCellSub" height="21"><span class="DataOnly"><%=sPayTo%></span>
														</td>
														<td width="150"></td>
												  </tr>
												  <tr>
													<td class="FieldCellSub" width="75" height="29">Rupees</td>
													<td class="FieldCellSub" height="29"><span class="DataOnly"><%=mid(AmountWords(dTotal),7)%></span>
													</td>
													<td width="150">
													  <table border="0" cellpadding="0" cellspacing="0" width="100%">
														<tr>
														  <td class="FieldCellSub" width="20" style="padding: 5">Rs.
														  </td>
														  <td class="FieldCellSub" style="padding: 5"><span width="100%" class="DataOnly"><%=FormatNumber(dTotal,2,,,0)%>&nbsp;</span>
														  </td>
														</tr>
													  </table>
													</td>
												  </tr>
												  <tr>
													<td style="padding-top: 10; padding-bottom: 10" height="49" colspan="3">
													  <table border="0" cellpadding="0" cellspacing="0">
														<tr>
														  <td class="FieldCellSub" width="50" style="padding: 5">A/c No.&nbsp; </td>
														  <td class="FieldCellSub" style="padding: 5"><span class="DataOnly"><B><%=sACNo%></b></span>
														  </td>
														</tr>
													  </table>
													</td>
												  </tr>
												  <tr>
													<td style="padding-top: 10; padding-bottom: 10" class="FieldCellSub" height="35" colspan="3"><%=sBankAddress%></td>
												  </tr>
												  <tr>
													<td colspan="3" style="padding-top: 10; padding-bottom: 10" class="FieldCellSub" height="42">
													  <div align="center">
														<center>
														<table border="0" cellpadding="0" cellspacing="0">
														  <tr>
															<td class="FieldCellsub">Cheque
															  Number</td>
															<td class="FieldCellsub"><input type="text" name="txtChequeNo" value="<%=sInsNo%>" size="23" class="formelem" maxlength="20"></td>
														  </tr>
														</table>
														</center>
													  </div>
													</td>
												  </tr>
												</table>
											</td>
											<td align="center" width="5" class="ClearPixel">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td colspan="3" align="center" width="5" class="ClearPixel">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<%
												objRs.MoveNext
											loop
											objRs.Close
										%>
										<tr>
											<td colspan="3" align="center" width="5" class="ClearPixel">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
															<input type="button" value="Print" name="btnAction" onclick="finaldone('A')" class="ActionButton">
															<input type="button" value="Back" name="btnAction1" class="ActionButton" onclick="Submit()">
														</td>
														<td align="center" class="ClearPixel" width="5">
															<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
														</td>
													</tr>
												</table>
											</td>
											<td align="center" width="5" class="ClearPixel">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td class="MiddlePack" colspan="3">
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</Table>
					</td>
				</tr>
			</Table>

		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
