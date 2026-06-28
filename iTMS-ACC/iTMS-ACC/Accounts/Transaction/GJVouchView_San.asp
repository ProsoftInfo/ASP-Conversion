<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GJVouchView_San.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:	Sre Hari.M
	'Modified By				:
	'Created On					:	Mar 02, 2006
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

'XML DOM Variables
Dim oDOM,nodHeader,Root,objRs,sQuery
dim sNarration,sAccount,sAddtional,iSno,sNarr
dim dTotal,sOrgId,sAccHead,dTotDr,dTotCr
dim EntryNode,HeaderNode,dAmount
dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim iTransNo,iBkHeadCode,bOtherUnit,iTdsAmount,sExp,TempNode,sBankInsDet
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request("TransNo")
'Response.Write iTransNo

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
oDOM.Load server.MapPath(GetVouchXML(iTransNo))
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

set EntryNode= Root.childNodes(0)
sPayTo = EntryNode.Attributes.Item(2).nodeValue
set HeaderNode=EntryNode.childNodes(0)
'sPayTo= HeaderNode.Attributes.Item(3).nodeValue

	For each EntryNode in Root.childNodes
		For each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="Narration" then
				sNarr=sNarr &"," & HeaderNode.text
			end if
		next
	next


sExp = "//BankInstrumentDet"
Set tempNode = Root.selectNodes(sExp)
IF tempNode.length <> 0 Then
	sBankInsDet = tempNode.Item(0).Attributes.getNamedItem("InsType").value
	sBankInsDet = sBankInsDet&" No: "
	sBankInsDet = sBankInsDet& tempNode.Item(0).Attributes.getNamedItem("InsNo").value
	sBankInsDet = sBankInsDet& " "
	sBankInsDet = sBankInsDet& tempNode.Item(0).Attributes.getNamedItem("InsDate").value
	sBankInsDet = sBankInsDet& " Drawn On "
'	sInsPayat = tempNode.Item(0).Attributes.getNamedItem("Payableat").value

	sBankInsDet = sBankInsDet & tempNode.Item(0).Attributes.getNamedItem("Drawnon").value
End IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Cash Vouchers</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../Scripts/itms-modern-compat.js"></SCRIPT>
<script language="javascript">
function field(form, name) {
	return form.elements[name] || form.elements[String(name).toLowerCase()] || null;
}

function openModernDialog(url, features) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, "", features || "dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No", function () {});
	} else {
		window.open(url, "_blank", "height=200,width=300,resizable=no,status=no");
	}
}

function approveVoucherIfRequired(form) {
	var xhr;
	var response;
	if (form.hApprover.value === "Y") {
		if (form.selUserid.selectedIndex > 0) {
			xhr = new XMLHttpRequest();
			xhr.open("POST", "XMLVouAppUpdate.asp?BkCode=CA&TransNo=" + encodeURIComponent(form.hTransNo.value) + "&User=" + encodeURIComponent(form.selUserid.value) + "&Mode=E", false);
			xhr.send(null);
			response = String(xhr.responseText || "").trim();
			if (response !== "") {
				alert(response);
				return false;
			}
		} else {
			alert("Select Approver");
			form.selUserid.focus();
			return false;
		}
	}
	return true;
}

function FinalCheck(flag) {
	var form = document.formname;
	var transNo;
	var orgNameField;
	var value;
	var url;
	if (!approveVoucherIfRequired(form)) {
		return;
	}
	if (flag === "B") {
		form.action = "VouCABookSelection.asp";
		form.submit();
	} else if (flag === "PV") {
		form.action = "VouCAEntry.asp";
		form.selVouType.value = "C";
		form.submit();
	} else if (flag === "RV") {
		form.action = "VouCAEntry.asp";
		form.selVouType.value = "D";
		form.submit();
	} else if (flag === "P") {
		transNo = form.hTransNo.value;
		orgNameField = field(form, "hOrgName");
		value = transNo + ":" + (orgNameField ? orgNameField.value : "");
		url = (form.selVouType.value === "D" ? "PRNCashRecpVouView.asp?Value=" : "PRNCashPayVouView.asp?Value=") + encodeURIComponent(value);
		openModernDialog(url);
	}
}

function CheckPrint(sPara) {
	var transNo = document.formname.hTransNo.value;
	var url = sPara === "GJ" ? "PrnGJView.asp?iTransNo=" + encodeURIComponent(transNo) : "PRNGJCNoteNew.asp?iTransNo=" + encodeURIComponent(transNo);
	openModernDialog(url);
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="selUnitId" value="<%=sOrgId%>">
<input type="hidden" name="horgName" value="<%=sOrgName%>">
<input type="hidden" name="hNarr" value="<%=mid(sNarr,2)%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="selBook" value="<%=iBookCode%>">
<input type="hidden" name="selVouType" value="<%=sVouType%>">
<input type="hidden" name="hBookOtherUnit" value="<%=bOtherUnit%>">

<input type="hidden" name="hBookAccHead" value="<%=iBkHeadCode%>">

<input type="hidden" name="hApprover" value="<%=sApprove%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		  Journal Voucher
		</td>
    </tr>

	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
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
                            <td class="FieldCell" width="145"> </td>
                            <td colspan=4 width="210">
                                 <span class="DataOnly"></span>
                            </td>
                            <td class="FieldCell" >Date</td>
								<td colspan="4">
                                       <span class="DataOnly"><%=FormatDate(date) %> </span>
								 </td>
                            </tr>

                            <tr>
								<td class="FieldCell" width="145">Voucher Date</td>
								<td colspan="4">
									<span class="DataOnly"><%=sVoucDate  %></span>
								</td>

								<td class="FieldCell" >Voucher No</td>
								<td colspan="4">
									<span class="DataOnly"><%=iVouNo %></span>
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
												<DIV class=frmBody id=frm1 style="width:600; height:170;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="3%">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="77%">Particulars</td>
                                        <td class="ExcelHeaderCell" align="center" width="20%" >Code</td>
                                        <td class="ExcelHeaderCell" align="center" width="20%" >Debit</td>
                                        <td class="ExcelHeaderCell" align="center" width="20%">Credit</td>
                                            </tr>
<%
dim sTemp,sAccType,nodADD,dCRAmt,dDRAmt,dTdsTotalAmt,bPayRec,sAmount,CheckNode
Dim dTdsAmt,dTdsPer,sCrDr
bPayRec=false
iSno = 0
for each EntryNode in Root.childNodes
	'iSno=EntryNode.Attributes.Item(0).nodeValue

If EntryNode.Attributes.Item(1).nodeValue = "D" then
	iSno=iSno + 1
	sCrDr    = EntryNode.Attributes.Item(1).nodeValue
	sAmount  = EntryNode.Attributes.Item(3).nodeValue
	sOrgName = EntryNode.Attributes.Item(5).nodeValue
'Response.Write sCrDr &"<---->"
	sExp = "//Entry[@No="&iSno&" and @TdsAmount]"
	Set CheckNode = Root.selectNodes(sExp)

	IF CheckNode.length <> 0 Then
		dTdsAmt = EntryNode.Attributes.Item(6).nodeValue
		dTdsPer = EntryNode.Attributes.Item(8).nodeValue
	Else
		dTdsAmt = 0
		dTdsPer = 0
	End IF

	sAmount=FormatNumber(sAmount,2,,,0)
	dTdsAmt=FormatNumber(dTdsAmt,2,,,0)
	dTdsPer=FormatNumber(dTdsPer,2,,,0)


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
					sAccount=HeaderNode.Attributes.Item(3).nodeValue
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
				sAddtional=sAddtional&"--------------------  <br>"
				for each  nodADD in HeaderNode.childNodes
					sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue&"-"
					sAddtional=sAddtional&nodADD.Attributes.Item(3).nodeValue &"%&nbsp;"
					sAddtional=sAddtional&nodADD.Attributes.Item(4).nodeValue&"<br>"
				next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
				for each  nodADD in HeaderNode.childNodes
					'sAddtional="Doc No:"
					'sAddtional=sAddtional&nodADD.Attributes.Item(1).nodeValue
					'sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue &"-&nbsp;"
					'sAddtional=sAddtional&nodADD.Attributes.Item(5).nodeValue&"<br>"
				next
		end if 'End of Check for PayRec Node
	next 'End of Entry Node Loop

	IF CStr(sAddtional) = "" Then
		sAddtional = sAccount
	End IF

%>
                <tr>
					<td class="ExcelSerial" align="Center" valign="top"><%=iSno%></td>
					<td class="ExcelDisplayCell" align="right" valign="top"><p align="left" ><%=sAddtional%></td>
					<td class="ExcelDisplayCell" align="Center" valign="top">&nbsp;</td>
					 <%'if EntryNode.Attributes.Item(1).nodeValue="D" then%>
						<td class="ExcelDisplayCell" align="right" valign="top"><%=sAmount%></td>
						<td class="ExcelDisplayCell" align="right"  valign="top"></td>
					<%'else%>
						<!--td class="ExcelDisplayCell" align="right" valign="top" ></td>
						<td class="ExcelDisplayCell" align="right" valign="top"><%=sAmount%></td-->
					<%'end if%>
                </tr>


<%
End IF
next'End of Voucher Node Loop

for each EntryNode in Root.childNodes
	'iSno=EntryNode.Attributes.Item(0).nodeValue

If EntryNode.Attributes.Item(1).nodeValue = "C" then
	iSno=iSno + 1
	sCrDr    = EntryNode.Attributes.Item(1).nodeValue
	sAmount  = EntryNode.Attributes.Item(3).nodeValue
	sOrgName = EntryNode.Attributes.Item(5).nodeValue

	sExp = "//Entry[@No="&iSno&" and @TdsAmount]"
	Set CheckNode = Root.selectNodes(sExp)

	IF CheckNode.length <> 0 Then
		dTdsAmt = EntryNode.Attributes.Item(6).nodeValue
		dTdsPer = EntryNode.Attributes.Item(8).nodeValue
	Else
		dTdsAmt = 0
		dTdsPer = 0
	End IF

	sAmount=FormatNumber(sAmount,2,,,0)
	dTdsAmt=FormatNumber(dTdsAmt,2,,,0)
	dTdsPer=FormatNumber(dTdsPer,2,,,0)


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
					sAccount=HeaderNode.Attributes.Item(3).nodeValue
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
				sAddtional=sAddtional&"--------------------  <br>"
				for each  nodADD in HeaderNode.childNodes
					sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue&"-"
					sAddtional=sAddtional&nodADD.Attributes.Item(3).nodeValue &"%&nbsp;"
					sAddtional=sAddtional&nodADD.Attributes.Item(4).nodeValue&"<br>"
				next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
				for each  nodADD in HeaderNode.childNodes
					'sAddtional="Doc No:"
					'sAddtional=sAddtional&nodADD.Attributes.Item(1).nodeValue
					'sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue &"-&nbsp;"
					'sAddtional=sAddtional&nodADD.Attributes.Item(5).nodeValue&"<br>"
				next
		end if 'End of Check for PayRec Node
	next 'End of Entry Node Loop

	IF CStr(sAddtional) = "" Then
		sAddtional = sAccount
	End IF

%>
                <tr>
					<td class="ExcelSerial" align="Center" valign="top"><%=iSno%></td>
					<td class="ExcelDisplayCell" align="right" valign="top"><p align="left" ><%=sAddtional%></td>
					<td class="ExcelDisplayCell" align="Center" valign="top">&nbsp;</td>
					<td class="ExcelDisplayCell" align="right" valign="top" ></td>
					<td class="ExcelDisplayCell" align="right" valign="top"><%=sAmount%></td>

                </tr>


<%
End IF
next'End of Voucher Node Loop
%>
                                        <tr>
											<td class="ExcelDisplayCell" align="center" colspan="2"><p align="right"><b>Total Debit&nbsp;</b></td>
											<td class="ExcelDisplayCell" align="right">&nbsp;</td>
											<td class="ExcelDisplayCell" align="right"><B><%= FormatNumber(dDRAmt,2,,,0)%></B></td>
											<td class="ExcelDisplayCell" align="right"></td>
                                        </tr>

                                        <tr>
											<td class="ExcelDisplayCell" align="center" colspan="2"><p align="right"><b>Total Credit&nbsp;</b></td>
											<td class="ExcelDisplayCell" align="right" colspan="4" ><B><%=FormatNumber(dCRAmt,2,,,0)%></B></td>
                                        </tr>

                                                </table>
                                               Being <%=sNarration  %>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
                                                               <span class="DataOnly"><%=AmountWords(sAmount)%></span>
                            </td>
                                </tr>

                                <!--tr>
                            <td class="FieldCell" width="130">Immediate Approver</td>
                            <td class="FieldCell">
                            <select size="1" name="selUserId" class="FormElem">
                        <option value="I">Immediate Approver</option>

                            </select></td>
                                </tr-->

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
																<input type="button" value="Print GJ" name="B8" onClick="CheckPrint('GJ')" class="ActionButtonX">
																<input type="button" value="Print Credit Note" name="B10" onClick="CheckPrint('CN')" class="ActionButtonX">
																<input type="button" value="Close" name="B9" onClick="window.close()" class="ActionButtonX"  >
																<!--input type="button" value="Next Receipt Voucher" name="B10" onClick="FinalCheck('RV')" class="ActionButtonX"  >
                                                                <input type="button" value="Next Book" name="B7" onClick="FinalCheck('B')" class="ActionButtonX" >-->


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
