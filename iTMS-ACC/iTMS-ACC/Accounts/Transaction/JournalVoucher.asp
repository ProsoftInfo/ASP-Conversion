<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	JournalVoucher.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 28,2011
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
<!--#include File="../../include/CheckACCPrevFinYear.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,sQuery
dim iVouNo,objRs,objRs1,sVouDate,bActionFlag
dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sAccHead,sParType,sPartSubType
dim iEnNo,Entrynode,HeaderNode,dOpeningBal
dim sParCode,sNarration,sAccHeadname,sAccUnitName,bOtherUnits,iBookAccHead,dTransLimit
Dim sFinPeriod,sFinTemp,sMinDate,sMaxDate,sCurrFinYr,sVal,sTempVal,sCallFrm
Dim sFinFrm,sFinTo,sValTemp,sToDate,sFromDate

dim sAccount,sAddtional,iSno
dim dTotal
dim sVoucDate,iBookCode,sPayTo
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,sUserId

sUserId = getUserID()

sCallFrm = "N"
sVal = Request.QueryString("Val")
IF Cstr(sVal) <> "" Then
	sTempVal = Split(sVal,"~")
	sCallFrm = sTempVal(2)
	sTransNo = sTempVal(0)
End IF

'Response.Write sTransNo

sFinPeriod = Session("FinPeriod")
'Response.Write sFinPeriod
IF CStr(sFinPeriod) <> "" Then
	sFinTemp = Split(sFinPeriod,":")
	sTodate = "31/03/"&sFinTemp(1)
	sFromDate = "01/04/"&sFinTemp(0)
End IF

sCurrFinYr = Cdbl(Year(Now)-1)&":"&Cdbl(Year(Now))

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("horgName")
sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("selVouType")
'sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")

sOrgId = Session("organizationcode")
sOrgName = Session("OrgShortName")

sQuery = "Select Top 1 OUDefinitionID,OrgUnitDescription From DCS_OrganizationUnitDefinitions "&_
		 "Where Len(OUDefinitionID) > 4 Order By OUDefinitionID "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
'	sOrgId = objRs(0)
'	sOrgName = objRs(1)
End IF
objRs.Close

sQuery = "Select Top 1 BookNumber,BookName,isNull(BookAccountHead,0),OtherUnitTransaction From vwOrgBookNames Where  "&_
		 "OUDefinitionID = '"&sOrgId&"' and BookCode = '08' Order By BookNumber "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sBookCode = objRs(0)
	sBookName = objRs(1)
	iBookAccHead = objRs(2)
	bOtherUnits = objRs(3)
Else
	sBookCode = "08"
	sBookName = ""
	iBookAccHead = 0
	bOtherUnits = 1
End IF
objRs.Close

sFinPeriod = Session("FinPeriod")
sValTemp = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp(0))
sFinTo = Trim(sValTemp(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"
bOtherUnits = 1
oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
oDOM.Save server.MapPath("../temp/transaction/Voucher Entry_GJ_"&Session.SessionID&".xml")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS GJ Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<script language="javascript" src="../../scripts/VouSelection.js"></script>
<script language="javascript" src="../../scripts/VoucherEntryCore.js"></script>
<script language="javascript" src="../../scripts/JournalVoucher.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=sBookCode%>" BookName="" CRDR="" VouDate="" BookAcchead="0" Approver=""/></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<xml id="GLHeadData"><Root /></xml>
<xml id="PartyData"><Root /></xml>
<XML ID="UnitBookData">
<Book/>
</XML>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="setdate();SetUnBook();DisplayBook();">
<form method="POST" name="formname" action="VouGenerate.asp">
<input type="hidden" name="hVouCode" value="08">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="GJ">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=sBookCode%>">
<input type="hidden" name="hOtherUnitFlag" value="<%=bOtherUnits%>">
<input type="hidden" name="hActionFlag" value="<%=bActionFlag%>">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hTransNo" value="<%=sTransNo%>">
<input type="hidden" name="hEntryNo" value="1">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hAmendTy" value="N">
<input type="hidden" name="hToDate" value="<%=sToDate%>">
<input type="hidden" name="hFromDate" value="<%=sFromDate%>">
<input type="hidden" name="hAmendChk" value="N">
<input type="hidden" name="hCallFrm" value="<%=sCallFrm%>">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hPreBookNo" value="1">
<input type="hidden" name="hAction" value="New">

<input type="hidden" name="hEditEntNo" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class=PageTitle>
          General Journal With Adjustments
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20px" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<!--td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td-->
								<td class="TabCurrentCell" valign="bottom" align="center" width="110px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70px">
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
					<TD class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
                            <td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
							<tr>
							    <td align="center" width="5px"></td>
								<td align="center" height="8px">
                                                      <table border="0" width="100%" cellspacing="1px" class="TableOutlineOnly">
                                                        <tr>
                                                       <!--   <td class="FieldCellSub" width="110">Organization</td>
                                                          <td class="FieldCellSub">
                                                          <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
																<OPTION value="0">Select</option>
																<%populateOrganizationListDBWithVal(sOrgID)%>
															</select>
                                                            </td>-->
                                                          <td class="FieldCellSub" width="40px" style="height: 22px">Book</td>
                                                          <td class="FieldCellSub" colspan="2" style="height: 22px">
																<select size="1" name="selBook" class="FormElem">
																	<option value="S">Select</option>
																</select>
															</td>
															<td class="FieldCellSub" width="100px" style="height: 22px" >Voucher
                                                            Date&nbsp;</td>
                                                          <td class="FieldCellSub" width="100px" style="height: 22px"><% ' Function Call to Insert Date Picker
																Response.Write InsertDatePicker("ctlDate")
															%>
															</td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" colspan=3 style="height: 20px"></td>
                                                          <td class="FieldCellSub" width="100px" style="height: 20px">Voucher Number</td>
                                                          <td class="FieldCellSub" width="100px" style="height: 20px">
														      <input type="text" name="txtVouNo" size="15" class="FormElem" readonly>
														  </td>
                                                        </tr>

                                                        <tr>
                                                          <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                      </table>
								</td>
								<td align="center" width="5px"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8px">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>
							<tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="100%">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="2" width="139px"></td>
                                                        </tr>
                                                    <tr>
                                                          <td class="FieldCellSub" width="90px">Entry Type</td>
                                                          <td class="FieldCell">
                                                            <input type=hidden name="hAccUNitId"  value = "<%=sOrgId %>">
                                                          <table border=0 width=100%>
                                                              <tr>
                                                              <td class="FieldCell">
                                                                  <input type=radio name="selCRDR" value="D" >Debit
                                                                  <input type=radio name="selCRDR" value="C" checked>Credit</td>
                                                              <td class="FieldCellSub" align=right>Entry Number&nbsp;
                                                                <span class="DataOnly" id="spEntryNo">1&nbsp;</span>
                                                              </td>
                                                              </tr>
                                                           </table>
                                                          </td>
                                                           </tr>
                                                        <tr>
                                                            <td class="FieldCellSub" width="139">Accounting Head</td>
                                                            <td class="FieldCell">
                                                                    <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead()">
															        <option value="A">Select Account Head</option>
															        <%
																        dim iHeadCount
															 	        'iHeadCount=popFrequentHead(sOrgId,"08",sBookCode)
															 	        iHeadCount = 0
															        %>
																        <option value="G">General Ledger</option>
															        <%populatePartyType(sOrgId)%>
                                                            </select> &nbsp; <a href="javascript:selAccountHead()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Account Head"></a>
                                                            </td>
                                                            <input type="hidden" name="hHeadCount" value="<%=iHeadCount%>">
														</tr>
                                                    	<tr>
                                                    <td class="FieldCellSub" width="139px"></td>
                                                    <td><span class="DataOnly" id="spAccHead"></span> </td>
                                                        </tr>
                                                        <!--tr>
                                                    <td class="FieldCellSub" width="139">Pay to / Received from</td>
                                                    <td class="FieldCell"> <input type="text" name="txtPayTo" size="40" class="Formelem">
                                                     &nbsp; <!--a href="javascript:SelMisParty()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a--><!--/td>
                                                        </tr-->
                                                        <tr>
                                                    <td width="139" valign="top">
                                                      <table border="0" width="100%" cellspacing="1px">
                                                        <tr>
                                                          <td width="50%" class="FieldCellSub">Narration</td>
                                                          <td width="50%" class="FieldCellSub">
																<a href="javascript:showNarration('08')"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
                                                           </td>
                                                        </tr>
                                                      </table>
                                                      &nbsp;</td>
                                                    <td class="FieldCell" valign="top"> <textarea rows="3" name="txtNarration" cols="50" class="FormElem" onKeyPress="ChkEnter()"></textarea> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell"> <input type="text" name="txtAmount" size="15" style="text-align:right" maxlength="13" class="FormElem" onblur="popAddAmount()"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Deduction @</td>
                                                    <td class="FieldCell" width="591px"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="FormElem" readOnly>
                                                    % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="FormElem" readOnly>&nbsp;
                                                    <input type="Button" value="Add Entry" name="btnAdd" onClick="AddNew()" class="AddButton" >
                                                    </td>
                                                </tr>
                                                <tr>
								                    <td width="100%" colspan=2 align="center">
                                                        <DIV class="frmBody" id="DisVoucher" style="width:98%; visibility:hidden; height:1px;">
	                                                        <table border="0" cellspacing="1px" id="tblVoucher" class="ExcelTable" width="98%">
	                                                        <tr>
		                                                        <td class="ExcelHeaderCell" width="10px">S.No.</td>
		                                                        <td class="ExcelHeaderCell" width="25px"></td>
		                                                        <td class="ExcelHeaderCell" width="25px"></td>
		                                                        <!--<td class="ExcelHeaderCell" align="center" width="75">AU</td>-->
		                                                        <td class="ExcelHeaderCell">Account Code - Name</td>
		                                                        <td class="ExcelHeaderCell" width="125px">Narration</td>
		                                                        <td class="ExcelHeaderCell" width="100px">Cr Amount</td>
		                                                        <td class="ExcelHeaderCell" width="100px">Dr Amount</td>
		                                                        <td class="ExcelHeaderCell">CC/AH Details</td>
		                                                        <td class="ExcelHeaderCell">Deduction Amount</td>
		                                                        <td class="ExcelHeaderCell">Deduction Percentage</td>
	                                                        </tr>
	                                                        </table>
                                                        </div>
								                    </td>
                                                </tr>
                                         </table>
								</td>
								<td align="center" class="ClearPixel" width="5px">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>
							<tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td class="FieldCellSub" width="639px">Approval

								<input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="SetApp('Y')">
								Yes&nbsp;&nbsp;
								<input type="radio" value="N" name="optApprove" class="FormElem" onClick="SetApp('N')"> No
								&nbsp;&nbsp; Approver &nbsp; <select size="1" name="selUserId" class="FormElem">
											<option value="I">Immediate Approver</option>
											<%=populateEmployeeWithVal(sUserId)%>
											    </select></td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8px">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5px" class="ClearPixel">
								</td>
								<td >
<DIV class=frmBody id="Disaddtional" style="height:1; visibility: hidden;">
<div id="DisCCANL" class=frmBody style="height:1px; visibility: hidden;">
	<table cellpadding="0" cellspacing="0" >
		<tr>
			<td class=MiddlePack colspan="3"> </td>
		</tr>
		<tr>
			<td class="FieldCell">
				<DIV class="frmBody" id="DisCost" style="width:275px;height:100px;">
					<table border="0" id="tblCost" cellspacing="1px" class="ExcelTable">
						<tr>
							<td class="ExcelHeaderCell" align="center" width="10px">S.No.</td>
								<td class="ExcelHeaderCell" width="150px">Cost Center Head</td>
								<td class="ExcelHeaderCell">Ratio</td>
								<td class="ExcelHeaderCell">Amount</td>
						 </tr>
					</table>
				</div><!--End of CostCenter Display Division -->
			</td>
			<td class=ClearPixel width="5px">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">                   </td>
			<td class=FieldCell>
				<DIV class=frmBody id="DisAnal" style="width:275px; height:100px;">

					<table border="0" id="tblAnal" cellspacing="1px" class="ExcelTable">
						<tr>
								<td class="ExcelHeaderCell" width="10px">S.No.</td>
								<td class="ExcelHeaderCell" width="150px">Analytical Head</td>
								<td class="ExcelHeaderCell">Ratio</td>
								<td class="ExcelHeaderCell">Amount</td>
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
	<DIV class="frmBody" id="DisPayable" style="width: 555px; visibility: hidden; height:1px;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="555px">
			<tr>
				<td class="ExcelHeaderCell" rowspan="2" width="10px">S.No.</td>
				<td class="ExcelHeaderCell" colspan="2">Document</td>
				<td class="ExcelHeaderCell" width="275px" colspan="5">Amount</td>
		    </tr>
		   <tr>
				<td class="ExcelHeaderCell">Detail</td>
				<td class="ExcelHeaderCell">Date</td>
				<td class="ExcelHeaderCell">Amount</td>
				<td class="ExcelHeaderCell">Adjusted</td>
				<!--td class="ExcelHeaderCell" align="center">Cr/Dr Note Value</td-->
				<td class="ExcelHeaderCell">To Account</td>
				<td class="ExcelHeaderCell">To be adjusted</td>
				<td class="ExcelHeaderCell">To adjust</td>
		   </tr>
		</table>
	</div>
</div><!--End of Addtional Details Display  -->
								</td>
								<td align="center" class="ClearPixel" width="5px">
								</td>
                            </tr>

							<tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="ActionCell">
                                                               <!-- <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >-->
                                                                <!--<input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >-->
                                                                <input type="button" value="Save" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouGJBookSelection.asp')" class="ActionButton" >
                                                                <!--<input type="Button" value="Delete Voucher" name="btnVouDel" onClick="DelVou()" disabled=true class="ActionButtonX" >-->
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5px">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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
