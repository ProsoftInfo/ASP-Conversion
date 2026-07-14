<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNCommAmend.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	Feburary 14 2003
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo,sFlag
dim sBookName,sInvoiceNo,sTemp,arrPartyCode,sPartyCode,sPartyName
Dim sInvTemp,iCtr,sVouTemp,sVouchTy,sNarr,sAmount,sTransno,ODom
Dim sStr,TempNode,VouRoot,iSelBook,sChkGj
Dim sVouDate,sVouUnit,sVouNumber,sVouAmt,sAmdTy

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Dim sFinPeriod,sFromYr,sToYr,sTempYr

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("horgName")
iBookNo=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sInvoiceNo=Request.Form("selInvoiceNo")
sVouchTy = Request.Form("selVoucherType")
sInvTemp = Split(sInvoiceNo,",")
sTransno = Request.Form("hTransNo")
sFlag = Request("sFlag")
sAmdTy = Request("AmdType")
sChkGj = Request("hChkVal")
'Response.Write "sChkGj="&sChkGj

IF CStr(sAmdTy) = "A" Then
	sFlag = "True"
End IF

'Response.Write "Flag = "& sFlag

sVouTemp = Split(Request.Form("hVouDetails"),":")
IF CStr(sVouchTy) <> "SC" Then
	sTemp=Split(Request.Form("hVouDetails"),"-")
Else
	sNarr = "For Invoice "
	sAmount = 0
	sAmount = CDbl(sAmount)
	For iCtr = 0 To UBound(sVouTemp)
		sTemp = Split(sVouTemp(iCtr),"-")
		sNarr = sNarr&" "&sTemp(0)&" "&sTemp(1)&", "
		sAmount = CDbl(sAmount + sTemp(2))
	Next
End IF

'Response.Write sTransno
'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
oDOM.load server.MapPath(GetVouchXML(sTransNo))
oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CN_"&Session.SessionID&".xml")

Set VouRoot = ODom.documentElement
sStr = "//voucher"
Set TempNode = VouRoot.selectNodes(sStr)
IF TempNode.length <> 0 Then
	sOrgId = TempNode.Item(0).Attributes.getNamedItem("UnitNo").value
	sVouDate = TempNode.Item(0).Attributes.getNamedItem("VouDate").value
	sVouUnit = TempNode.Item(0).Attributes.getNamedItem("UnitName").value
	sVouNumber = TempNode.Item(0).Attributes.getNamedItem("VoucherNo").value
	iSelBook = TempNode.Item(0).Attributes.getNamedItem("BookNo").value
End IF

sStr = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr)
IF TempNode.length <> 0 Then
	sVouAmt = TempNode.Item(0).Attributes.getNamedItem("Amount").value
End IF

sPartyName=Request.Form("txtPartyName")
arrPartyCode=split(Request.Form("hPartyCode"),"?")

Set objRs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=edge">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="DetData">
<Root>

</Root>
</script>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script src="../../scripts/ExcelFunctions.js"></script>

<!--XML ISLAND FOR VOUCHER DATA -->
<script type="application/xml" data-itms-xml-island="1" id="VoucherData" data-src="<%="../temp/transaction/Voucher AMD_CN_"&Session.SessionID&".xml"%>"></script>
<!--XML ISLAND FOR ENTRY DATA -->
<script type="application/xml" data-itms-xml-island="1" id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="" TDSElgi="0" TdsPercentage="0" />
</script>

<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<script>
window.CNCommisionEntryConfig = {
	dataIsland: "VoucherData",
	saveIsland: "VoucherData",
	saveName: "Voucher AMD",
	checkApp: false,
	checkFinancialDate: false,
	forceMultiEntryFlow: true,
	includeAccUnit: true,
	setDateFromVoucher: true,
	setFinancialDateLimits: true,
	alwaysDisplayVoucher: true,
	showTotalRow: false,
	updateActionWhenFlagTrue: true,
	flagField: "hFlag",
	flagAction: "AmdAccCrNtGenerate.asp"
};
</script>
<script src="../../scripts/VouCNCommisionEntryCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitVouCNCommisionEntry()">

<form method="POST" name="formname" action="VouCNAmdGenerate.asp">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sVouUnit%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hInvDate" value="<%=sVouDate%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hVouchTy" value="<%=sVouchTy%>">
<input type="hidden" name="hEntryNo" value="">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hTransno" value="<%=sTransno%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hVouType" value="OT">
<input type="hidden" name="hCrTransNo" value="<%=sTransno%>">
<input type="hidden" name="hChkGj" value="<%=sChkGj%>">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Other
          Amendment
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
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
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">

                            </table>
                            </td>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="590">
                                    <tr>
										<td class="FieldCell" width="93">Unit</td>
										<td colspan="3"><span class="DataOnly"><%=sVouUnit%>&nbsp;</span></td>

	                                </tr>
	                                <%IF CStr(sPartyName) <> "" Then %>
										<tr>
											<td class="FieldCell" width="93">Party Name</td>
											<td colspan="3" ><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
										</tr>
									<%End IF %>
	                                <tr>
										<td class="FieldCell" width="100">Invoice No-Date</td>
										<td colspan="3"><span class="DataOnly"><%=sVouNumber%>-&nbsp;<%=sVouDate%> </span></td>
	                                </tr>
									<tr>
										<td class="FieldCell" width="113">Voucher Amount</td>
										<td width="230"><span class="DataOnly"><%=FormatNumber(sVouAmt,2,,,0)%>&nbsp;-&nbsp;CR&nbsp;</span></td>
										<td class="FieldCell" width="100"></td>
										<td></td>
	                                </tr>

                                    </table>

								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="5" width="139"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Head</td>
                                                    <td class="FieldCell">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selGLHead(this)">
															<option value="A">Select Account Head</option>
															<option value="G">General Ledger</option>

                                                    </select>
                                                    <input type="hidden" name="hHeadCount" value="1">
													 </td>
                                                    <td class="FieldCell" colspan="2"><p align="center">Date
                                                    </td>
                                                    <td class="FieldCell"> <p align="center">
                                                    <% ' Function Call to Insert Date Picker
															Response.Write InsertDatePicker("ctlDate")
													%>

														</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td>
 </td>
                                                    <td colspan="2"><p align="center"><!--Number--></p>
                                                    </td>
                                                    <td class="FieldCellSub">  </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td class="FieldCell" colspan="4">
                                                    <input type="text" name="txtPayTo" size="40" class="Formelem">
                                                    &nbsp; <a href="#" onclick="SelMisParty(); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139" valign="top">Narration</td>
                                                    <td class="FieldCell" colspan="2" valign="top">

														<textarea rows="3" name="txtNarration" cols="50" class="FormElem"></textarea> </td>

                                                    <td class="FieldCell" colspan="2" valign="middle">
 </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell" colspan="4">

														<input type="text" name="txtAmount" value="0.00" size="15" style="text-align:right" class="Formelem"> </td>

                                                        </tr>
                                                        <tr>
														<td class="FieldCellSub" width="133">Deduction @</td>
														<td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
														% On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
														</td>
														    </tr>
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="1">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td >
<DIV class=frmBody id="Disaddtional" style="height:1; visibility: hidden;">
<div id="DisCCANL" class=frmBody style="height:1; visibility: hidden;">
	<table cellpadding="0" cellspacing="0" >
		<tr>
			<td class=MiddlePack colspan="3"> </td>
		</tr>
		<tr>
			<td class=FieldCell>
				<DIV class=frmBody id="DisCost" style="width:280;height:100;">
					<table border="0" id="tblCost" cellspacing="1" class="ExcelTable">
						<tr>
							<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Cost Center Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
						 </tr>
					</table>
				</div><!--End of CostCenter Display Division -->
			</td>
			<td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">                   </td>
			<td class=FieldCell>
				<DIV class=frmBody id="DisAnal" style="width:280; height:100;">

					<table border="0" id="tblAnal" cellspacing="1" class="ExcelTable">
						<tr>
								<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Analytical Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
					    </tr>
					</table>
				</div>	<!--End of Analytical Display Division -->
			</td>
		</tr>
		<tr>
			<td class=MiddlePack  colspan="3"></td>
		</tr>
	</table>
</div> <!--End of CCANAL Display Division -->
</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">

																 <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
																 <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >
                                                                 <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >
                                                                 <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                 <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('CreditVouchers.asp')" class="ActionButton" >


														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="35">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

							<tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
<DIV class=frmBody id="DisVoucher" style="width:100%; visibility:hidden; height:1;">
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="700">
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="35">&nbsp;</td>
		<td class="ExcelHeaderCell" align="center" width="75">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="125">Amount</td>
		<td class="ExcelHeaderCell" align="center" >Additional Details</td>
		<td class="ExcelHeaderCell" align="center" >Deduction Amount</td>
		<td class="ExcelHeaderCell" align="center" >Deduction Percentage</td>

	</tr>
	</table>
</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
