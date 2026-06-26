<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNCommDelView.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 01,2003
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
<!--#include file="../../include/Accpopulate.asp"-->
<%

'XML DOM Variables
Dim oDOM,nodHeader,Root
dim sNarration,sAccount,sAddtional,iSno,sAmount
dim dTotal,sOrgId
dim EntryNode,HeaderNode
dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim sParType,sParSubType,sParCode,iTransNo
Dim sStr,TempNode
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request("hTransNo")
'Response.Write iTransNo

oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")


set Root=oDOM.documentElement
sStr = "//voucher"
Set TempNode = Root.selectNodes(sStr)
IF TempNode.length <> 0 Then
	Set Root = TempNode.Item(0)
End IF

sOrgId=Root.Attributes.Item(0).nodeValue
sOrgName=Root.Attributes.Item(1).nodeValue
iBookCode =Root.Attributes.Item(2).nodeValue
sBookName =Root.Attributes.Item(3).nodeValue
sVoucDate=Root.Attributes.Item(5).nodeValue
iVouNo=Root.Attributes.Item(10).nodeValue
sVouType="C"

sStr = "//Entry"
Set TempNode = Root.selectNodes(sStr)
IF TempNode.length <> 0 Then
	'dTotal=TempNode.Item(0).Attributes.Item(2).nodeValue
	sPayTo=TempNode.Item(0).Attributes.getNamedItem("Payto").Value
End IF
'Response.Write sPayTo

sStr = "//Party"
Set TempNode = Root.selectNodes(sStr)
IF TempNode.length <> 0 Then
	sParType=TempNode.Item(0).Attributes.Item(0).nodeValue
	sParSubType=TempNode.Item(0).Attributes.Item(1).nodeValue
	sParCode=TempNode.Item(0).Attributes.Item(2).nodeValue

End IF

sStr = "//Entry"
Set TempNode = Root.selectNodes(sStr)
set EntryNode=TempNode.Item(0)



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>

<script>
function FinalCheck(Flag,iTraNo)
{
	if (Flag=='B')
	{
		document.formname.action="VouCNBookSelection.asp";
		document.formname.submit();
		return true;
	}
}
</script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="VouCnCommDeletion.asp">
<input type="hidden" name="selUnitId" value="<%=sOrgId%>">
<input type="hidden" name="horgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="selBook" value="<%=iBookCode%>">
<input type="hidden" name="selVouType" value="<%=sVouType%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		  <%
				Response.Write "Sales Commission Credit Note"
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
									<td class="FieldCell" width="145">Unit </td>
									<td colspan="3"><span class="DataOnly"><%=sOrgName%></span></td>

                                </tr>

                                <tr>
									<td class="FieldCell" width="145"> Book Code </td>
									<td width="210"><span class="DataOnly"><%=iBookCode%></span></td>
									<td class="FieldCell" width="88"> Voucher No.</td>
									<td><span class="DataOnly"><%=iVouNo%></span></td>
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
									Response.Write "Pay To"
                            %>

                            </td>
                            <td colspan="3">
								<span class="DataOnly"><%=sPayTo%></span>
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
												<DIV class=frmBody id=frm1 style="width:600; height:140;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="700">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center">Account Code - Name / Narration&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center" width="150">Additional Details</td>
                                        <td class="ExcelHeaderCell" align="center" width="120" >Amount</td>
                                        <td class="ExcelHeaderCell" align="center" width="80" >Deduction Amount</td>
                                        <td class="ExcelHeaderCell" align="center" width="80" >Deduction Percentage</td>

                                            </tr>
<%
dim sTemp,sAccType,nodADD,dCRAmt,dDRAmt,iCtr,dTdsAmt,dTdsPer,CheckNode,sStr2





	sStr = "//Entry"
	Set TempNode = Root.selectNodes(sStr)
	For iCtr = 0 To TempNode.length - 1
		iSno = TempNode.Item(iCtr).Attributes.Item(0).nodeValue
		sAmount = TempNode.Item(iCtr).Attributes.getNamedItem("Amount").value

		sStr2 = "//Entry[@No="&iSno&" and @TdsAmount]"
		Set CheckNode = Root.selectNodes(sStr2)

		IF CheckNode.length <> 0 Then
			dTdsAmt = TempNode.Item(iCtr).Attributes.getNamedItem("TdsAmount").value
			dTdsPer = TempNode.Item(iCtr).Attributes.getNamedItem("TdsPercentage").value
		Else
			dTdsAmt = 0
			dTdsPer = 0
		End IF

		sAddtional=""

		for each HeaderNode in TempNode.Item(iCtr).childNodes
			if HeaderNode.nodeName="AccHead" then
					sAccType=HeaderNode.Attributes.Item(4).nodeValue

						sAccount=HeaderNode.Attributes.Item(0).nodeValue
						sAccount=sAccount& "-" & HeaderNode.Attributes.Item(3).nodeValue
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
			if 	HeaderNode.nodeName="Analytical" then
					for each  nodADD in HeaderNode.childNodes
						sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue&"-"
						sAddtional=sAddtional&nodADD.Attributes.Item(3).nodeValue &"%&nbsp;"
						sAddtional=sAddtional&nodADD.Attributes.Item(4).nodeValue&"<br>"
					next
			end if 'End of Check for Analytical Node
		next 'End of Entry Node Loop

		dTotal = CDbl(sAmount) + dTotal

%>
                <tr>
            <td class="ExcelSerial" align="center" rowspan="2"><%=iSno%></td>
            <td class="ExcelDisplayCell" align="right"><p align="left"><%=sAccount%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><p align="left"><%=sAddtional%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=FormatNumber(sAmount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=FormatNumber(dTdsAmt,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right" rowspan="2"><%=FormatNumber(dTdsPer,2,,,0)%></td>

                </tr>
            <tr>
				 <td class="ExcelDisplayCell" align="right"><p align="left"><%=sNarration%></td>
            </tr>
  <%Next%>

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
																<input type="submit" value="Delete" name="B7" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="B8" class="ActionButton" onClick="Cancel('../AccountsHome.asp')">

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