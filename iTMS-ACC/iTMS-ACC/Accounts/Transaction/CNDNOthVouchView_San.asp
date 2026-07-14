<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CNDNOthVouchView_San.asp
	'Module Name				:	ACCOUNTS (Transaction)
	'Author Name				:	Sre Hari M
	'Created On					:	Mar 21, 2006
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

'XML DOM Variables
Dim oDOM,nodHeader,Root,objRs,sQuery,EntryNode,HeaderNode,TempNode
dim sNarration,sAccount,sAddtional,iSno,sNarr
dim dTotal,sOrgId,sAccHead,sType,sAccNo
dim dAmount,cAmount,sPartyName,sPartyCode,sExp,sPAdd1,sPAdd2,sPCity,sPState,sPCountry
dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim iTransNo,iBkHeadCode,arrTemp,iTdsAmount,sSelAccHead,sRetVal

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request("TransNo")
'Response.Write iTransNo

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)

set Root=oDOM.documentElement
set objRs = Server.CreateObject("ADODB.Recordset")


sQuery = "Select CreatedVoucherNo,TransactionType,Convert(Char,VoucherDate,103) VoucherDate,VoucherAmount From  "&_
		 "Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo&" "
		' Response.Write sQuery  & "<BR><BR>"
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With
Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	iVouNo=objRs("CreatedVoucherNo")
	sVouType = objRs("TransactionType")
	sVoucDate = objRs("VoucherDate")
End IF
objRs.Close
sPartyCode=Root.Attributes.getNamedItem("PartyCode").value
if InStr(1,sPartyCode,"?")>0 then
	arrTemp=Split(sPartyCode,"?")
	sPartyCode=arrTemp(3)
else
	sPartyCode=sPartyCode
end if
sQuery = "SELECT AddressLine1,isNull(AddressLine2,''),isNull(City,''),isNull(State,''),isNull(Country,''),PartyName " & _
		 "From App_M_PartyMaster Where PartyCode ="& sPartyCode

With Objrs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
Set Objrs.ActiveConnection = nothing
IF not objRs.EOF  Then
	sPAdd1 = objRs(0)
	sPAdd2 = objRs(1)
	sPCity = objRs(2)
	sPState = objRs(3)
	sPCountry = objRs(4)
	sPayTo=objRs(5)
End IF
objRs.Close

sQuery = "Select M.AccountDescription From Acc_M_GLAccountHead M,Acc_T_CreatedVoucherDetails D  "&_
		 "Where D.AccUnitAccountHead = M.AccountHead and D.CreatedTransNo = "&iTransNo&" "
Objrs.Open sQuery,Con
IF Not Objrs.Eof Then
	sSelAccHead	= Objrs(0)
End IF
Objrs.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Credit Vouchers</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
function CheckPrint() {
	var transNo = document.formname.hTransNo.value;
	var isCredit = document.formname.selVouType.value === "CNR";
	var bookCode = isCredit ? "07" : "06";
	var xhr = new XMLHttpRequest();
	xhr.open("GET", "PrintInsert.asp?BkCode=" + encodeURIComponent(bookCode) + "&UserId=1234", false);
	xhr.send(null);
	window.open((isCredit ? "PRNCNNoteView.asp" : "PrnDBNoteView.asp") + "?iTransNo=" + encodeURIComponent(transNo), "", "height=200,width=300,toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
}
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="selUnitId" value="<%=sOrgId%>">
<input type="hidden" name="horgName" value="<%=sOrgName%>">
<input type="hidden" name="hNarr" value="<%=sNarr%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="selBook" value="<%=iBookCode%>">
<input type="hidden" name="selVouType" value="<%=sVouType%>">

<input type="hidden" name="hBookAccHead" value="<%=iBkHeadCode%>">

<input type="hidden" name="hApprover" value="<%=sApprove%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">


<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		  <%
			if sVouType="CNR" then
				Response.Write "Credit Note"
			else
				Response.Write "Debit Note"
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
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" height="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">

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
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" > Voucher No. </td>
                            <td align="left">
                                    <span class="DataOnly"><%=iVouNo %>  </span>
                            </td>
                            <td class="FieldCell" >
                            Date</td>
                            <td align="left">
                                     <span class="DataOnly"><%=sVoucDate %>  </span>
                            </td>
                            </tr>

                            <tr>

								<td class="FieldCell" >A/C Head </td>
								<td align="left">
									<span class="DataOnly"> <%=sSelAccHead  %>  </span>
								   </td>
                            </tr>

                                <tr>
									<td colspan="4">&nbsp;</td>
                                </tr>
                                <tr>

                             <td class="FieldCell"><%=" To "%></td>
                            <td colspan="4"><span class="DataOnly"><%=sPayTo %> </span></td>
                            </tr>
                            <tr><td class="FieldCell" ></td>
                            <td ><span class="DataOnly"><%=sPAdd1 %> </span></td></tr>
                            <tr>
                            <tr><td class="FieldCell"></td>
                            <td colspan="4"><span class="DataOnly"><%=sPAdd2 %> </span></td></tr>
                            <tr>
                            <tr><td class="FieldCell" ></td>
                            <td colspan="4"><span class="DataOnly"><%=sPCity %> </span></td></tr>
                            <tr>
                            <tr><td class="FieldCell" ></td>
                            <td colspan="4"><span class="DataOnly"><%=sPState%> </span></td></tr>
                            <tr>
                            <tr><td class="FieldCell" ></td>
                            <td colspan="4"><span class="DataOnly"><%=sPCountry %> </span></td></tr>
                             <tr><td class="FieldCell" ></td></tr>


                            <tr><td class="FieldCell" ></td><td class="FieldCellSub" ><%
								IF CStr(sVouType) = "DNR" Then
									Response.Write "Sir / Sirs We have debited your account with us as detailed below :"
								Else
									Response.Write  "Sir / Sirs We have credited your account with us as detailed below :"
								End IF
							%>
							</td></tr>
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
												<DIV class=frmBody id=frm1 style="width:100%; height:120;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
										<td class="ExcelHeaderCell" align="center" width="75" colspan="2">Particulars</td>
                                        <td class="ExcelHeaderCell" align="center" width="80" >Amount</td>

                                            </tr>
<%
dim sCheck,sAccType,nodADD,dCRAmt,dDRAmt,dTdsTotalAmt,bPayRec,sParType
sCheck="F"
iSno=0

sExp = "//TaxDetails"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.Length <> 0 Then
		dAmount = TempNode.Item(0).Attributes.Item(0).nodeValue

		sExp = "//SaleInvoice"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.Length <> 0 Then
			sCheck = "T"
			sNarration = sNarration & TempNode.Item(0).Attributes.getNamedItem("InvNo").value
			sNarration = sNarration &"-"& TempNode.Item(0).Attributes.getNamedItem("InvDate").value
		End IF

		IF CStr(sCheck) = "F" Then
			sExp = "//PurInvoice"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.Length <> 0 Then
				sCheck = "T"
				sNarration = sNarration & TempNode.Item(0).Attributes.getNamedItem("PurInvNo").value
				sNarration = sNarration &"-"& TempNode.Item(0).Attributes.getNamedItem("PurInvDate").value
			End IF
		End IF

		sExp="//Narration"
		set TempNode=Root.selectNodes(sExp)
		if TempNode.length<>0 then
			sNarr="Being  "& TempNode.Item(0).text
		end if

		sNarration = sNarr & "  For Inv No " & sNarration
		iSno=iSno+1
else
	dAmount = 0
	set EntryNode=Root.ChildNodes
		for each EntryNode in Root.ChildNodes
		if EntryNode.nodeName="Entry" then

			dAmount = Cdbl(dAmount) + Cdbl(EntryNode.Attributes.Item(3).nodeValue)

			sExp="//Narration"
			set TempNode=Root.selectNodes(sExp)
			if TempNode.Length<>0 then
				sNarration = sNarration & EntryNode.text
			end if
		end if
		next

		sNarration="Being "& sNarration
		dAmount=FormatNumber(dAmount,2,,,0)
		iSno=iSno+1
end if

%>
             <tr>
            <td class="ExcelSerial" align="center" valign="top"><%=iSno %></td>
			<td class=ExcelDisplayCell colspan="2"><%=sNarration %><br></td>
            <td class="ExcelDisplayCell" align="right" valign="top"><%=FormatNumber(dAmount,2,,,0)%></td>
              </tr>
               <tr>
                <td class="ExcelDisplayCell" align="center" colspan="3"><p align="right"><b>Total&nbsp;</b></td>
                <td class="ExcelDisplayCell" align="right">Rs.&nbsp;<%=FormatNumber(dAmount,2,,,0)%></td>

                    </tr>
                        </table>
						</div>
		</td>
		<td align="center" class="ClearPixel" width="5">
		</td>
	</tr>
    <!--tr>
		<td align="center" class="MiddlePack" colspan="3">
		</td>
    </tr-->
	<tr>
		<td align="center" width="5" class="ClearPixel">
        &nbsp;
		</td>
		<td valign="top" class="FieldCell" height="20%">
            <table cellpadding="0" cellspacing="0" width="100%">
        <tr>
    <td class="FieldCell" width="130" valign="top">Amount </td>
                            <td>
                                                               <span class="DataOnly"><%=AmountWords(dAmount)%></span>
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
																<input type="button" value="Print" name="B8" onClick="CheckPrint()" class="ActionButton">
																<input type="button" value="Close" name="B9" onClick="window.close()" class="ActionButton"  >
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
