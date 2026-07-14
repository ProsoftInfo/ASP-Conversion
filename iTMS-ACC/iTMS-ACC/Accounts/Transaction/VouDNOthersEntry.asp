<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNOtherEntry.asp
	'Module Name				:	ACCOUNTS (Transcation Debit Note Amendment For Other Voucher Type)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 31,2011
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
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo
dim sBookName,sInvoiceNo,sTemp,arrPartyCode,sPartyCode,sPartyName
Dim sInvTemp,iCtr,sVouTemp,sVouchTy,sNarr,sAmount,sTransno,ODom
Dim sStr,TempNode,VouRoot,sUserID
Dim sVouDate,sVouUnit,sVouNumber,sVouAmt,sSelVouTy,sSelInvNo
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sCallFrom,sVouCode,sVouName

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
'sOrgId=Request.Form("selUnitId")
sOrgId = Session("organizationcode")
iBookNo=Request.Form("selBook")
'Response.Write "iBookNo = "& iBookNo
sUserID = getUserID()

sPartyName=Request.Form("txtPartyName")
'sOrgName=Request.Form("horgName")
sOrgName = Session("OrgShortName")
sBookName=Request.Form("hBookName")
sPartyCode=Request.Form("hPartyCode")
sVouTemp = Request.Form("hVouDetails")
sSelVouTy = Request.Form("selVoucherType")
sSelInvNo = Request.Form("selInvoiceNo")

Set objRs = Server.CreateObject("ADODB.RecordSet")
sCallFrom = Request("CallFrom")
if Trim(sCallFrom)="GJ" then
    sVouCode = "08"
    sVouName = "GJ"
else
    sVouCode = "06"
    sVouName = "DN"
end if
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
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root /></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script src="../../scripts/ExcelFunctions.js"></script>

<!--XML ISLAND FOR VOUCHER DATA -->
<script type="application/xml" data-itms-xml-island="1" id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=iBookNo%>" BookName="<%=sBookName%>" CRDR="D" VouDate="" PartyCode="<%=sPartyCode%>" PartyName="<%=Replace(sPartyName,"&","and")%>" Approver=""/></script>
<!--XML ISLAND FOR ENTRY DATA -->
<script type="application/xml" data-itms-xml-island="1" id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName=""/>
</script>

<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="GJVoucher"></script>
<script src="../../scripts/VouDNOthersEntryCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitVouDNOthersEntry()">

<form method="POST" name="formname">
<input type="hidden" name="hVouCode" value="<%=sVouCode%>">
<input type="hidden" name="hVouName" value="<%=sVouName%>">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hTransNo" value="0">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hSelVouTy" value="<%=sSelVouTy%>">
<input type="hidden" name="hInvNos" value="<%=sSelInvNo%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hAction" value="New" >
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Debit Note Other

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
							<!--tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr-->
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
                            <!--tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr-->
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border=0 cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
										<td class="FieldCell" width="75">Agent Name</td>
										<td ><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
										<td class="FieldCell" width="75">Voucher Date</td>
										<td class="FieldCell" width="75"> <p align="center">
                                        <% ' Function Call to Insert Date Picker
												Response.Write InsertDatePicker("ctlDate")
										%>

										</td>

	                                </tr>
									<tr>


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
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td class="MiddlePack" colspan="5"></td>
                                                        </tr>
                                                        <tr>
                                                            <td class="FieldCellSub" width="95">Accounting Head</td>
                                                            <td class="FieldCell">
                                                                    <select size="1" name="selAccHead" class="FormElem" onChange="selGLHead(this)">
															        <option value="A">Select Account Head</option>
															        <option value="G">General Ledger</option>
                                                            </select>
                                                            <input type="hidden" name="hHeadCount" value="1">
													         </td>
													         <td colspan=3 class=FieldCellSub>Entry No&nbsp;&nbsp;<span id="spEntryNo" class="DataOnly">1</span></td>
                                                        </tr>
                                                        <tr>
                                                            <td class="FieldCellSub" width="95"></td>
                                                            <td></td>
                                                            <td colspan="2"><p align="center"><!--Number--></p>
                                                            </td>
                                                            <td class="FieldCellSub">  </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="FieldCellSub" width="95"></td>
                                                            <td class="FieldCell" colspan="4">
                                                            <input type="text" name="txtPayTo" size="40" class="Formelem"> </td>
                                                        </tr>
                                                        <tr>
                                                                <td class="FieldCellSub" width="95" valign="top">Narration&nbsp;&nbsp;&nbsp;
														            <a href="#" onclick="showNarration(); return false;"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
                                                                </td>
                                                                <td class="FieldCell" colspan="2" valign="top">

														            <textarea rows="3" name="txtNarration" cols="50" class="FormElem" onKeyPress="return ChkEnter(event)"></textarea> </td>

                                                                <td class="FieldCell" colspan="2" valign="middle">
                                                                </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="95">Amount</td>
                                                    <td class="FieldCell" colspan="4">

														<input type="text" name="txtAmount" value="0.00" size="15" style="text-align:right" class="Formelem" onblur="popAddAmount()"> </td>

                                                        </tr>

                                                        <tr>
														    <td class="FieldCellSub" width="95">Approval</td>
														    <td class="FieldCell" colspan="3">
														    <input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="SetApp('Y')">Yes&nbsp;
														    <input type="radio" value="N" name="optApprove" class="FormElem" onClick="SetApp('N')"> No
														    &nbsp; Immediate Approver &nbsp;
														    <select size="1" name="selUserId" class="FormElem">
															    <option value="I">Immediate Approver</option>
															    <%=populateEmployeeWithVal(sUserId)%>
														    </select>
														    &nbsp;<input type="Button" value="Add Entry" name="btnAdd" onClick="AddNew()" class="AddButton" >
														    </td>
													    </tr>
    					                                <tr>
								                            <td colspan=4 align=center>
                                                                    <DIV class=frmBody id="DisVoucher" style="width:95%; visibility:hidden; height:1;">
	                                                                    <table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="95%">
	                                                                    <tr>
		                                                                    <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="35">&nbsp;</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="35">&nbsp;</td>
		                                                                    <!--<td class="ExcelHeaderCell" align="center" width="75">AU</td>-->
		                                                                    <td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="100">Amount</td>
		                                                                    <td class="ExcelHeaderCell" align="center" >Additional Details</td>
	                                                                    </tr>
	                                                                    </table>
                                                                    </div>
								                            </td>
                                                        </tr>
                                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="1">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                            <!--tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr-->
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">

																 <!--<input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >-->
                                                                <!--<input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >-->
                                                                <input type="button" value="Save" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouCNBookSelection.asp')" class="ActionButton" >


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
