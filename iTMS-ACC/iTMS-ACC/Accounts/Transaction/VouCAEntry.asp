<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCAEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January 06,2003
	'Modified By				:	Manohar Prabhu.R
	'Modified On				:	Sep 28, 2004
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,sQuery
dim iVouNo,objRs,objRs1,sVouDate,bActionFlag
dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sAccHead,sParType,sPartSubType
dim iEnNo,Entrynode,HeaderNode,dOpeningBal
dim sParCode,sNarration,sAccHeadname,sAccUnitName,bOtherUnits,iBookAccHead,dTransLimit

dim sAccount,sAddtional,iSno
dim dTotal
dim sVoucDate,iBookCode,sPayTo,sUserId
sUserId = getUserID()
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("horgName")
sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("selVouType")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")

oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS Cash Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<meta http-equiv="x-ua-compatible" content="IE=10">
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script src="../../scripts/ExcelFunctions.js"></script>
<script src="../../scripts/VouSelection.js"></script>
<script src="../../scripts/VoucherEntryCore.js"></script>
<script src="../../scripts/CashVoucher.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/VouCAEntryCompat.js"></script>

<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=sBookCode%>" BookName="<%=sBookName%>" CRDR="<%=sVouType%>" VouDate="" BookAcchead="<%=iBookAccHead%>" Approver=""/></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="0" TDSElgi="0" TdsPercentage="0" PayRecAmount="0" /></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="TDSData"  ><Root/></xml>
<xml id="GLHeadData"><Root /></xml>
<xml id="PartyHeadData"><Root /></xml>
<XML id="AccHeadData">
<account/>
</XML>
<XML ID="UnitBookData">
<Book/>
</XML>
<XML ID="TDSFlagData">
<Root/>
</XML>
<XML id="VoucherAmdData"></XML>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();popAccHead();DisplayBalamt();">
<form method="POST" name="formname" action="VouGenerate.asp">
<input type="hidden" name="hVouCode" value="01">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="CA">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=sBookCode%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hBookAccHead" value="<%=iBookAccHead%>">
<input type="hidden" name="hOtherUnitFlag" value="<%=bOtherUnits%>">
<input type="hidden" name="hActionFlag" value="<%=bActionFlag%>">
<input type="hidden" name="hTransLimit" value="<%=dTransLimit%>">
<input type="hidden" name="hTransNo" value="0">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hPayTo" value="">
<input type="hidden" name="hTDSElgi" value="0">
<input type="hidden" name="hTotalAmt" value="0">
<input type="hidden" name="hPayRecCount" value="0">
<input type="hidden" name="hSelPayRecCount" value="0">
<input type="hidden" name="hTotType" value="N">

<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<% IF CStr(sVouType) = "C" Then
				Response.Write("Cash Payment Voucher")
		   Else
				Response.Write("Cash Receipt Voucher ")
		   End IF
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
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="left">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: pointer" Title="Month wise Balance" >
                    <p align="center"><font face="Webdings" size="5">�</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font face="Webdings" size="5">�</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
                    <font face="Webdings" size="5">�</font>
                    </span>
                    </p>
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" height="8">
                                                      <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                                       <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Voucher
                                                            Date</td>
                                                          <td class="FieldCellSub" width="125">
                                                          <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>

                                                          </td>
                                                          <td class="FieldCellSub" width="80">Entry
                                                      Number</td>
                                                          <td class="FieldCellSub"><span class="DataOnly" id="spEntryNo">1&nbsp;</span></td>
                                                          <td class="FieldCellSub" width="100">
                                                    Book Balance</td>
                                                          <td class="FieldCellSub">
   <span class="DataOnly">
                                                            <%
                                                             dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
                                                             dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                             if dOpeningBal<0 then
                                                             %>
																<span class="DataOnly" id="spBookBal"><%Response.Write dOpeningBal*-1 &"&nbsp;Cr"%></span>
                                                             <%

															 else
															%>
																<span class="DataOnly" id="spBookBal"><%Response.Write dOpeningBal &"&nbsp;Dr"%></span>
															<%

															 end if
                                                            %></span>
   &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Entry
                                                            Type</td>
                                                          <td class="FieldCellSub" colspan="3">
											<%if sVouType="C" then%>
                                                            <input type=radio name="selCRDR" value="C" disabled>Receipts
                                                            <input type=radio name="selCRDR" value="D" checked>Payments
                                                            <%else%>
                                                            <input type=radio name="selCRDR" value="C" checked>Receipts
                                                            <input type=radio name="selCRDR" value="D" disabled >Payments&nbsp;
											<%end if%>
                                                            </td>
                                                          <td class="FieldCellSub" width="100">Current
                                                      Balance&nbsp;</td>
                                                          <td class="FieldCellSub"><span class="DataOnly" id="iCurrentBal">
                                                            <%
                                                             dOpeningBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
                                                             dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                             if dOpeningBal<0 then
                                                            %>
																<span class="DataOnly" id="spCurrBal"><%Response.Write dOpeningBal*-1 &"&nbsp;Cr"%></span>
                                                            <%

															 else
															%>
																<span class="DataOnly" id="spCurrBal"><%Response.Write dOpeningBal &"&nbsp;Dr"%></span>
															<%

															 end if
                                                            %></span>
                                                            &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                      </table>
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
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="100%">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="2" width="139"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Unit</td>
                                                    <td class="FieldCell">
													<%	if bOtherUnits=1 then%>
                                                     <select size="1" name="selAccUnitId" onChange="popAccHead()" class="FormElem">
													<option value="A">Account Unit</option>
													<option value="<%=sOrgId%>" selected><%=sOrgName%></option>
															<%=popIUTUnits(sOrgId)%>
													</select>
													<%
														else
															Response.Write sOrgName
														end if
													%>


 </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Head</td>
                                                    <td class="FieldCell">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
															<option value="A">Select Account Head</option>
															<%
																dim iHeadCount
															 	'iHeadCount=popFrequentHead(sOrgId,"01",sBookCode)

															%>
																<option value="G">General Ledger</option>
															<%populatePartyType(sOrgId)%>
                                                    </select>
                                                    </td>
                                                    <input type="hidden" name="hHeadCount" value="0">

														</tr>
                                                    	<tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td><span class="DataOnly" id="spAccHead"></span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Pay to / Received from</td>
                                                    <td class="FieldCell"> <input type="text" name="txtPayTo" size="40" class="Formelem">
                                                    &nbsp; <a href="javascript:SelMisParty()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a></td>
                                                        </tr>
                                                        <tr>
                                                    <td width="139" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td width="50%" class="FieldCellSub">Narration</td>
                                                          <td width="50%" class="FieldCellSub">
<%

sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
	" OUDefinitionID='"&sOrgId&"'and BookCode='01' and BookNumber="&sBookCode

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

if objRs(0)>0 then
%>
                                                            <p align="left">
                                                    <a href="javascript:showNarration('01')"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
<%
end if
objRs.Close
%>
                                                           </td>
                                                        </tr>
                                                      </table>
                                                      &nbsp;</td>
                                                    <td class="FieldCell" valign="top"> <textarea rows="3" name="txtNarration" cols="50" class="FormElem"></textarea> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell"> <input type="text" name="txtAmount" size="15" value="0.00" style="text-align:right" maxlength="13" class="Formelem" onblur="popAddAmount()"> </td>
                                                        </tr>
                                                         <tr>
                                                    <td class="FieldCellSub" width="133">Deduction @</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    &nbsp; % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    </td>
                                                        </tr>
                                                         <!--tr>
															<td class="FieldCellSub" width="133">Approval</td>
															<td class="FieldCell" width="591">
															<input type="radio" value="Y" checked name="optApprove" class="FormElem">
															Yes&nbsp;&nbsp;
															<input type="radio" value="N" name="optApprove" class="FormElem"> No </td>
														</tr-->
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
								<td class="FieldCellSub" width="639">Approval

								<input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="SetApp('Y')">
								Yes&nbsp;&nbsp;
								<input type="radio" value="N" name="optApprove" class="FormElem" onClick="SetApp('N')"> No
								&nbsp;&nbsp; Approver &nbsp; <select size="1" name="selUserId" class="FormElem">
											<option value="I">Immediate Approver</option>
											<%=populateEmployeeWithVal(sUserId)%>
											    </select></td>
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
				<DIV class=frmBody id="DisCost" style="width:260;height:100;">
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
				<DIV class=frmBody id="DisAnal" style="width:260; height:100;">

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
	<DIV class=frmBody id="DisPayable" style="width: 555; visibility: hidden; height:1;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="555">
			<tr>
				<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
				<td class="ExcelHeaderCell" align="center" colspan="2">Document</td>
				<td class="ExcelHeaderCell" align="center" width="275" colspan="5">Amount</td>
		    </tr>
		   <tr>
				<td class="ExcelHeaderCell" align="center">Detail</td>
				<td class="ExcelHeaderCell" align="center">Date</td>
				<td class="ExcelHeaderCell" align="center">Amount</td>
				<td class="ExcelHeaderCell" align="center">Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To Account</td>
				<td class="ExcelHeaderCell" align="center">To be Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To adjust</td>

		   </tr>
		</table>
	</div>
</div><!--End of Addtional Details Display  -->
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
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
                                                                <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >
                                                                <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >
                                                                <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouCABookSelection.asp')" class="ActionButton" >
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
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
<DIV class=frmBody id="DisVoucher" style="width:585; visibility:hidden; height:1;">
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" style="width:660;" >
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="25"></td>
		<td class="ExcelHeaderCell" align="center">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center">Additional Details</td>
		<td class="ExcelHeaderCell" align="center">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="70">Amount</td>
		<td class="ExcelHeaderCell" align="center" width="70">Deduction Amount</td>
		<td class="ExcelHeaderCell" align="center" width="70">Deduction Percentage</td>
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
