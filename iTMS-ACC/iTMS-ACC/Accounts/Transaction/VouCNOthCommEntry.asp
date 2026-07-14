<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNOthCommEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	Feburary 14 2003
	'Modified By				:	Manohar Prabhu.R
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
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo
dim sBookName,sInvoiceNo,sTemp,arrPartyCode,sPartyCode,sPartyName
Dim sInvTemp,iCtr,sVouTemp,sVouchTy,sNarr,sAmount,sTempAmt,sUserid
Dim oDom,Root,MainNode,PartyNode,sVouchDetails

Set oDom = Server.CreateObject("Microsoft.XMLDOM")

sUserid = getUserID()

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("horgName")
iBookNo=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sInvoiceNo=Request.Form("selInvoiceNo")
sVouchTy = Request.Form("selVoucherType")
sVouchDetails = Request.Form("hVouDetails")
sInvTemp = Split(sInvoiceNo,",")

sVouTemp = Split(Request.Form("hVouDetails"),":")

sPartyName=Request.Form("txtPartyName")
arrPartyCode=split(Request.Form("hPartyCode"),"?")

Set objRs = Server.CreateObject("ADODB.RecordSet")

Set Root = oDom.createElement("Root")
oDom.appendChild Root
For iCtr = 0 To UBound(sInvTemp)
	Set MainNode = oDom.createElement("voucher")
	MainNode.setAttribute "UnitNo", sOrgId
	MainNode.setAttribute "UnitName", sOrgName
	MainNode.setAttribute "BookNo", iBookNo
	MainNode.setAttribute "BookName", sBookName
	MainNode.setAttribute "VouDate", ""
	MainNode.setAttribute "Approver", ""
	MainNode.setAttribute "SalTransNo", sInvTemp(iCtr)
	MainNode.setAttribute "SalVouNo", ""
	MainNode.setAttribute "SalVouDate", ""

	Set PartyNode = oDom.createElement("Party")
	PartyNode.setAttribute "ParType", trim(arrPartyCode(0))
	PartyNode.setAttribute "ParSubType", trim(arrPartyCode(1))
	PartyNode.setAttribute "ParCode", trim(arrPartyCode(3))
	PartyNode.text = sPartyName

	MainNode.appendChild PartyNode
	Root.appendChild MainNode
Next

oDOM.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-CNCommEntry.xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=edge">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="DetData" data-src="<%="../Temp/Transaction/"&Session.SessionID&"-CNCommEntry.xml"%>">
</script>

<script type="application/xml" data-itms-xml-island="1" id="EntryData">
<Entry No="0" Payto="" Amount="" CRDR="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script src="../../scripts/VouTransactions.js"></script>

<script>
window.CNCommisionEntryConfig = {
	dataIsland: "DetData",
	saveIsland: "DetData",
	saveMod: "CN",
	saveName: "Voucher Entry",
	checkFinancialDate: false,
	forceMultiEntryFlow: true
};
</script>
<script src="../../scripts/VouCNCommisionEntryCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitVouCNCommisionEntry()">

<form method="POST" name="formname" action="VouCNOthCommGen.asp">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hInvDate" value="0">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hVouchTy" value="<%=sVouchTy%>">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hEditEntry" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Commission
          Entry
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
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: pointer" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">�</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font size="3" face="Webdings">�</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
                    <font size="4" face="Webdings">�</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell">
                    &nbsp;
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
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
										<td colspan="3"><span class="DataOnly"><%=sOrgName%>&nbsp;</span></td>

	                                </tr>
									<tr>
										<td class="FieldCell" width="93">Agent Name</td>
										<td width="230"><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
									<%IF CStr(sVouchTy) <> "SC" Then %>
										<td class="FieldCell" width="100">Invoice No-Date</td>
										<td><span class="DataOnly">1&nbsp;-&nbsp;1&nbsp;</span></td>
									<%End IF %>
	                                </tr>
									<!--tr>
										<td class="FieldCell" width="113">Commision Amount</td>
										<td width="230"><span class="DataOnly"><%=FormatNumber(sAmount,2,,,0)%></span></td>
										<td class="FieldCell" width="100"></td>
										<td></td>
	                                </tr-->
	                                <tr>
										<td class="FieldCell" width="113">Entry Type</td>
										<td width="230" class="FieldCellSub">
										<Input type="radio" name="OptCRDR" value="C" class="FormElem">Credit &nbsp;&nbsp;
										<Input type="radio" name="OptCRDR" value="D" class="FormElem" checked>Debit &nbsp;&nbsp;</td>
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
                                                    &nbsp; <a href="#" onclick="SelMisParty(); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a>
                                                    </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139" valign="top">Narration</td>
                                                    <td class="FieldCell" colspan="2" valign="top">

														<textarea rows="3" name="txtNarration" cols="50" class="FormElem"><%=Trim(sVouchDetails)%></textarea> </td>

                                                    <td class="FieldCell" colspan="2" valign="middle">
 </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell" colspan="4">

														<input type="text" name="txtAmount" value="" size="15" style="text-align:right" class="Formelem"> </td>

                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Deduction @</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    </td>
                                                        </tr>

                                                        <tr>
															<td class="FieldCellSub" width="139">Approval</td>
															<td class="FieldCell" colspan="4">
																<Input type="radio" name="optApprove" class="FormElem" value="Y" checked onClick="SetApp('Y')">Yes &nbsp;
																&nbsp;&nbsp;&nbsp;
																<Input type="radio" name="optApprove" class="FormElem" value="N" onClick="SetApp('N')">No
																&nbsp;&nbsp;&nbsp; Immediate Approver &nbsp;&nbsp; <select size="1" name="selUserId" class="FormElem">
																		<option value="I">Immediate Approver</option>
																		<%=populateEmployeeWithVal(sUserId)%>
																		    </select>
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

																 <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntrySC('A')" class="ActionButton" >
																 <input type="Button" value="Update" name="btnUpdate" onClick="AddEntrySC('U')" class="ActionButton" disabled>
																 <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" class="ActionButton" disabled>



                                                                <input type="button" value="Next" onClick="AddEntrySC('S')" name="btnNext" class="ActionButton" >

                                                                <input type="reset" value="Cancel" name="B8" class="ActionButton" >

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
<DIV class=frmBody id="DisVoucher" style="width:585; visibility:hidden; height:1;">
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="700">
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
		<td class="ExcelHeaderCell" align="center" width="75">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="125">Amount</td>
		<td class="ExcelHeaderCell" align="center" >Additional Details</td>
		<td class="ExcelHeaderCell" align="center" width="80">Deduction Amount</td>
		<td class="ExcelHeaderCell" align="center" width="80">Deduction Percentage</td>

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
