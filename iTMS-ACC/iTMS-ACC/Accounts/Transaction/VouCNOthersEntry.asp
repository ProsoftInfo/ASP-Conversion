
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNOthersEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
dim sOrgId,sOrgName,sBookName,objRs,sQuery,iBookNo
dim sPartyCode,sPartyName,sUserID,sSelVouTy,sVouNarr,sSelInvNo
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sCallFrom,sVouCode,sVouName

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

'sOrgId=Request.Form("selUnitId")
sOrgId = Session("organizationcode")
sOrgName = Session("OrgShortName")
iBookNo=Request.Form("selBook")
sPartyName=Request.Form("txtPartyName")
'sOrgName=Request.Form("horgName")
sBookName=Request.Form("hBookName")
sPartyCode=Request.Form("hPartyCode")
sSelVouTy = Request.Form("selVoucherType")
sSelInvNo = Request.Form("selInvoiceNo")

'Response.Write "<p> Vou Type = "& sSelVouTy 

IF CStr(sSelVouTy) <> "OT" Then
	sVouNarr = Request.Form("hVouDetails")
Else
	sVouNarr = ""
End IF

sUserID = getUserID()


Set objRs = Server.CreateObject("ADODB.RecordSet")
sCallFrom = Request("CallFrom")
if Trim(sCallFrom)="GJ" then
    sVouCode = "08"
    sVouName = "GJ"
else
    sVouCode = "07"
    sVouName = "CN"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=edge">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script src="../../scripts/ExcelFunctions.js"></script>
<!--XML ISLAND FOR VOUCHER DATA -->
<script type="application/xml" data-itms-xml-island="1" id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=iBookNo%>" BookName="<%=sBookName%>" CRDR="C" VouDate="" PartyCode="<%=Replace(sPartyCode,"&"," and ")%>" Approver="" PartyName="<%=Replace(Trim(sPartyName),"&"," and ")%>" /></script>
<!--XML ISLAND FOR ENTRY DATA -->
<script type="application/xml" data-itms-xml-island="1" id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="GJVoucher"></script>
<script src="../../scripts/VouCNOthersEntryCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitVouCNOthersEntry()">

<form method="POST" name="formname">
<input type="hidden" name="hVouCode" value="<%=sVouCode%>">
<input type="hidden" name="hVouName" value="<%=sVouName%>">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hTransNo" value="0">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hSelVouTy" value="<%=sSelVouTy%>">
<input type="hidden" name="hInvNos" value="<%=sSelInvNo%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hAction" value="New">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20">Other Credit Note 		</td>
    </tr>
	<tr>
		<td align="center" class=MiddlePack height="20"><p align="center"> 		</td>
    </tr>
    <tr>
								<td align="center">
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
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
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                             <table cellpadding="0" cellspacing="0" width=100% border=0>
                                <tr>
                                    <td colspan=4 width=100%>
                                        <table border=0 cellpadding=0 cellspacing=0>
                                            <tr>
                                                <td class="FieldCellSub" width="15%">Party Name </td>
                                                <td class="FieldCell" width="55%" ><span class="DataOnly"><%=sPartyName%></span>
                                                </td>
                                                <td  class="FieldCell">
                                                    Voucher Date
                                                </td>
                                                <td align=right class="FieldCell">
                                                 <% ' Function Call to Insert Date Picker
								                    Response.Write InsertDatePicker("ctlDate")
							                    %>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan=4>
                                    <table border=0 cellspacing=0 cellpadding=0 class="TableOutlineOnly" width=100%>
                                            <tr>
                                        <td class="FieldCellSub" width="160">GL Account Head</td>
                                        <td class="FieldCell">
                                        <select size="1" name="selAccountHead" class="FormElem" onChange="PopAccHead(this) ">
							            <option value="S">Select Account Head</option>
							            <%
										            dim iHeadCount
										            iHeadCount=popFrequentHead(sOrgId,"07",iBookNo)

							            %>
							            <option value="G">GL Account Head</option>
                                        </select>
                                        </td>
                                        <td class="FieldCellSub" width="150" colspan=2 align=right>Entry No&nbsp;&nbsp;
                                        <span class="DataOnly" id="spEntryNo">1</span></td>
                                            </tr>
                                            <tr>
                                        <td class="FieldCellSub" width="125"></td>
                                        <td class="FieldCell" colspan="3"> <span class="DataOnly" id="spAccHead"></span></td>
                                            </tr>
                                            <tr>
                                                                <td width="139" valign="top">
                                                                 <table border="0" width="100%" cellspacing="1">
                                                                    <tr>
                                                                      <td width="50%" class="FieldCellSub">Narration</td>
                                                                      <td width="50%" class="FieldCellSub">
            <%

            sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
	            " OUDefinitionID='"&sOrgId&"'and BookCode='07' and BookNumber="&iBookNo

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
                                                                <a href="#" onclick="showNarration('07'); return false;"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
            <%
            end if
            objRs.Close
            %>
                                                                       </td>
                                                                    </tr>
                                                                  </table>

                                                                </td>
                                                                <td class="FieldCell" colspan="3" valign="top"> <textarea rows="3" name="txtNarration" cols="50" class="FormElem"><%=Trim(sVouNarr)%></textarea> </td>

                                            </tr>
                                            <tr>
                                        <td class="FieldCellSub" width="115">Amount</td>
                                        <td class="FieldCell" colspan="3">
                                        <input type="text" name="txtAmount" value="0.00" size="15" maxlength="15" style="text-align:right" class="FormElem" onblur="popAddAmount()"> </td>
                                            </tr>
                                            <tr>
                                            <td class="FieldCellSub" width="133">Deduction @</td>
                                            <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                            % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                            </td>
                                                </tr>

                                            <tr>
                                                                <td class="FieldCellSub" width="139">Approval</td>
                                                                <td class="FieldCell" colspan="3">
                                                                <input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="SetApp('Y')">Yes&nbsp;&nbsp;

													            <input type="radio" value="N" name="optApprove" class="FormElem" onClick="SetApp('N')"> No
													            &nbsp;&nbsp; Immediate Approver &nbsp;
													            <select size="1" name="selUserId" class="FormElem">
														            <option value="I">Immediate Approver</option>
														            <%=populateEmployeeWithVal(sUserId)%>
													            </select>
													            &nbsp;<input type="button" value="Add Entry" name="btnAdd" class="AddButton" onclick="AddNew()" >
													            </td>
                                                            </tr>
                                                            <tr>
							                                    <td align=center width=100% colspan=4>
                                                                        <DIV class=frmBody id="DisVoucher" style="width:98%; visibility:hidden; height:1;">
                                                                            <table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="98%">
                                                                            <tr>
	                                                                            <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
	                                                                            <!--<td class="ExcelHeaderCell" align="center" width="75">AU</td>-->
	                                                                            <td class="ExcelHeaderCell" align="center">Account Code - Name</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="125">Narration</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="80">Amount</td>
	                                                                            <td class="ExcelHeaderCell" align="center">CC/AH Details</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="80">Amount</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="80">Amount</td>
                                                                            </tr>
                                                                            </table>
                                                                        </div>
							                                    </td>
                                                            </tr>
                                                    </table>
								                </td>
								            <td align="center" class="ClearPixel" width="5" height="1">
                                        &nbsp;
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
								<td valign="top" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <!--<input type="button" value="Update" name="btnUpdate" class="ActionButton" onclick="AddEntry('U')" disabled>
                                                                <input type="button" value="Delete" name="btnDel" class="ActionButton" onclick="DelEntry()" disabled>-->
                                                                <input type="button" value="Save" onClick="AddEntry('S')" name="btnNext" class="ActionButton" >
                                                                <input type="button" value="Cancel" onClick="CancelAction('CREDITNOTETOCREATE.asp')" name="B8" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="35">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
