
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	CashVoucher.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	June 10, 2005
	'Modified By                :   Ragavendran R
	'Modified On                :   Jan 18,2011
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
<!--#include File="../../include/IncludeDatePicker.asp" -->
<!--#include File="../../include/CheckACCPrevFinYear.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,sQuery
dim iVouNo,objRs,objRs1,sVouDate,bActionFlag,sVal,sValTemp
dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sAccHead,sParType,sPartSubType
dim iEnNo,Entrynode,HeaderNode,dOpeningBal
dim sParCode,sNarration,sAccHeadname,sAccUnitName,bOtherUnits,iBookAccHead,dTransLimit
Dim sVouCkTy,sLastVouDt,sSelVouTy
Dim sFinPeriod,sFinFrm,sFinTo,sValTemp2,sFormVal,sSelArg
dim sAccount,sAddtional,iSno,sAction
dim dTotal
dim sVoucDate,iBookCode,sPayTo,sUserId,iPreBookVal
Dim sRetVal,sFinFromDate,sFinToDate
sUserId = session("userid")
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,sLogUID

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("hVouType")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")
sSelVouTy = Request("VOUTY")

sSelArg = Request("voutype")
sFormVal = Request("hFormVal")

iPreBookVal = sBookCode

sOrgId = Session("organizationcode")
sOrgName = Session("OrgShortName")
'Response.Write "sOrgID = "& sOrgId
'Response.Write "sOrgName = "& sOrgName
'Response.Write sVouType
sAction = Session("ACTN")
sVal=Request("Val")
'Response.Write sVal
sValTemp=Split(sVal,"~")
IF CStr(sVouType) = "" Then
	IF CStr(sSelVouTy) = "P" Then
		sVouType = "C"
	Else
		sVouType = "D"
	End IF
End IF
'Response.Write sVouType
IF Cstr(sVal) <> "" Then
	sQuery = "Select H.OUDEFINITIONID,D.OrgUnitDescription,isNull(AccountHead,0),BookNumber From DCS_OrganizationUnitDefinitions D, "&_
		 "Acc_T_CreatedVoucherHeader H Where H.OUDEFINITIONID = D.OUDEFINITIONID "&_
		 "and H.CreatedTransNo = "&sValTemp(0)&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
'		sOrgId = objRs(0)
'		sOrgName = objRs(1)
        iBookAccHead = objrs(2)
        sbookCode = objrs(3)
	End IF
	objRs.Close

Else

	sQuery = "Select Top 1 OUDefinitionID,OrgUnitDescription From DCS_OrganizationUnitDefinitions "&_
			 "Where Len(OUDefinitionID) > 4 Order By OUDefinitionID "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
'		sOrgId = objRs(0)
'		sOrgName = objRs(1)
	End IF
	objRs.Close
End IF

IF CStr(iBookAccHead) = "" Then
	sQuery = "Select Top 1 BookNumber,BookName,isNull(BookAccountHead,0),OtherUnitTransaction From vwOrgBookNames Where  "&_
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '01' Order By BookName "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sBookCode = objRs(0)
		sBookName = objRs(1)
		iBookAccHead = objRs(2)
		bOtherUnits = objRs(3)
	Else
		iBookAccHead = 0
	End IF
	objRs.Close
else
sQuery = "Select Top 1 OtherUnitTransaction From vwOrgBookNames Where  "&_
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '01' Order By BookName "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		bOtherUnits = objRs(0)
	Else
		bOtherUnits = 1
	End IF
	objRs.Close
End IF

oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

''Blocked and Added by Ragav on Jan 13,2012
''Begin
'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
sRetVal = GetVouchXML(sTransNo)
IF Request.Form("hCallFrm") = "A" then
    oDOM.Load server.MapPath(sRetVal)
End IF
''End

'Response.Write Request.Form("hCallFrm")
'IF Request.Form("hCallFrm") <> "A" then
	'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")
'End IF


IF CStr(sVouType) = "C" Then
	sVouCkTy = "CAP"
Else
	sVouCkTy = "CAR"
End IF

sQuery = "Select Convert(Char,VoucherDate,103),CreatedBy From Acc_T_CreatedVoucherHeader Where CreatedTransNo =  "&_
		 "(Select Max(CreatedTransNo) From Acc_T_CreatedVoucherHeader Where BookCode = '01'  "&_
		 "and OUDefinitionID = '"&sOrgId&"' and TransactionType ='"&sVouCkTy&"' ) "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sLastVouDt = Trim(objRs(0))
	'sUserId = Trim(objRs(1))
End IF
objRs.Close

'IF CStr(sUserId) = "" Then
sUserId = session("userid")
'End IF

sLogUID = session("userid")

sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp2(0))
sFinTo = Trim(sValTemp2(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"
bOtherUnits = 1
sFinFromDate = "01/04/"& sValTemp2(0)
sFinToDate = "31/03/"&sValTemp2(1)



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS Cash Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<meta http-equiv="x-ua-compatible" content="IE=edge">
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
<!--XML ISLAND FOR VOUCHER DATA -->
<script type="application/xml" data-itms-xml-island="1" id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=sBookCode%>" BookName="<%=sBookName%>" CRDR="<%=sVouType%>" VouDate="" BookAcchead="<%=iBookAccHead%>" Approver=""/></script>
<!--XML ISLAND FOR ENTRY DATA -->
<script type="application/xml" data-itms-xml-island="1" id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="0" TDSElgi="0" TdsPercentage="0" PayRecAmount="0" /></script>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="TDSData"  ><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyHeadData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData">
<Book/>
</script>
<script type="application/xml" data-itms-xml-island="1" ID="TDSFlagData">
<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="VoucherAmdData"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();SelUnBook();DisplayBook('<%=sOrgID%>')">
<form method="POST" name="formname" action="VouGenerate.asp" >
<input type="hidden" name="hVouCode" value="01">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="CA">
<input type="hidden" name="hTdsAmt" value="">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=sBookCode%>">
<input type="hidden" name="hOtherUnitFlag" value="<%=bOtherUnits%>">
<input type="hidden" name="hActionFlag" value="<%=bActionFlag%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hPayTo" value="">
<input type="hidden" name="hTDSElgi" value="0">
<input type="hidden" name="hTotalAmt" value="0">
<input type="hidden" name="hPayRecCount" value="0">
<input type="hidden" name="hSelPayRecCount" value="0">
<input type="hidden" name="hTotType" value="N">
<input type="hidden" name="hUpdate" value="N">
<input type="hidden" name="hTdsNew" value="N">
<input type="hidden" name="hAction" value="New">

<input type="hidden" name="hVouType" value="<%=sVouType %>">

<%if Trim(sVal)<>"" then%>
	<input type="hidden" name="hTransNo" value="<%=sValTemp(0)%>">
	<input type="hidden" name="hAmendDet" value="<%=sValTemp(1)%>">
	<input type="hidden" name="hCallFrm" value="<%=sValTemp(2)%>">
<%else%>
	<input type="hidden" name="hCallFrm" value="C">
	<input type="hidden" name="hTransNo" value="0">
<%End if%>

<input type="hidden" name="hLastVouDt" value="<%=sLastVouDt%>">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hAmendTy" value="N">
<input type="hidden" name="hBookAccHead" value="<%=iBookAcchead%>">
<input type="hidden" name="hBookOtherUnit" value="1">
<input type="hidden" name="hPreBookSel" value="<%=iPreBookVal%>">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hFormVal" value="<%=sFormVal%>">
<input type="hidden" name="voutype" value="<%=sSelArg%>">
<input type="hidden" name="hFromDate" value="<%=sFinFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sFinToDate%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
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
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20" valign="bottom">
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
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Entry Details
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
                            <!--tr>
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
                    <p align="center"><font face="Webdings" size="5">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font face="Webdings" size="5">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
                    <font face="Webdings" size="5">?</font>
                    </span>
                    </p>
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
								</td>
							</tr>
							<tr>
							    <td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td align="center" height="8">
                                  <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                    <tr>
                                      <td class="FieldCellSub" width="110">Book</td>
                                      <td class="FieldCell" width="110">
										<select size="1" name="selBook" class="FormElem" onChange="SetBookAccHead()">
												<option value="S">Select</option>
											</select>
									   </td>
									   	<td class="FieldCellSub" width="90">Date</td>
										<td class="FieldCell" >
										    <% ' Function Call to Insert Date Picker
									Response.Write InsertDatePicker("ctlDate")%>
										</td>

										</tr>

										<tr>
											<td class="FieldCellSub" width="90">Current Balance
											<td class="FieldCell">
											<span class="DataOnly" id="spCurrBal">
                                                        <%
                                                         dOpeningBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
                                                         'dOpeningBal = 0
                                                         dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                         if dOpeningBal<0 then
															Response.Write dOpeningBal*-1 &"&nbsp;Cr"
														 else
															Response.Write dOpeningBal &"&nbsp;Dr"
														 end if
                                                        %> </span>
                                                        &nbsp;&nbsp;&nbsp;Book Balance
                                                        &nbsp;
															<span class="DataOnly" id="spBookBal">
															<%
															dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
															'dOpeningBal = 0
															dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
															if dOpeningBal<0 then
															Response.Write dOpeningBal*-1 &"&nbsp;Cr"
															else
															Response.Write dOpeningBal &"&nbsp;Dr"
															end if
															%> </span>&nbsp;
									        </td>
                                      <td class="FieldCellSub" width="80">Voucher No</td>

                                      <td class="FieldCell" width="110">
                                      <%if Trim(sVal)<>"" then%>
													<input type="text" name="txtVouNo" size="20" class="FormElem" value=<%=sValTemp(1)%> readonly>
												<%else%>
													<input type="text" name="txtVouNo" size="20" class="FormElem" readonly>
												<%end if%>

								</td>
                                </tr>
                                </tr>

                            </table>
                            </td>
                            <td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
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
                                                    <td class="MiddlePack" colspan="2" width=100%>
                                                            <tr>
                                                                <td class="FieldCellSub" width="100">Entry Type</td>
                                                                <td class="FieldCell">
                                                                <table border=0 width="100%">
                                                                        <tr>
                                                                        <td class="FieldCell">
													                     <input type=hidden name="hAccUnitId" value="<%=sOrgId%>">
													                    <%IF CStr(sVouType) = "D" Then %>
                                                                            <input type=radio name="selCRDR" value="C" checked>Receipts
                                                                            <input type=radio name="selCRDR" value="D" disabled>Payments&nbsp;
                                                                        <%Else%>
														                    <input type=radio name="selCRDR" value="C" disabled>Receipts
														                    <input type=radio name="selCRDR" value="D"  checked>Payments&nbsp;
                                                                        <%End IF %>
                                                                        </td>
                                                                        <td class="FieldCell"  align=right>
                                                                        &nbsp;&nbsp;&nbsp;&nbsp;Entry No&nbsp;&nbsp;&nbsp;
													                    <span class="DataOnly" id="spEntryNo"><b>1&nbsp;</b></span></td>
                                                                    </tr>
                                                            </table>
                                                   </tr>
                                                   <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Head</td>
                                                    <td class="FieldCell">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
															<option value="A">Select Account Head</option>
															<%
																dim iHeadCount
															 	'iHeadCount=popFrequentHead(sOrgId,"01",sBookCode)
																iHeadCount=0
															%>
																<option value="G">General Ledger</option>
															<%populatePartyType(sOrgId)%>
                                                    </select> &nbsp; <a href="#" onclick="selAccountHead((document.forms.formname || document.forms[0]).selAccHead); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Account Head"></a>
                                                    </td>
                                                    <input type="hidden" name="hHeadCount" value="<%=iHeadCount%>">

														</tr>
                                                    	<tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td><span class="DataOnly" id="spAccHead"></span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Pay to / Received from</td>
                                                    <td class="FieldCell"> <input type="text" name="txtPayTo" size="40" class="FormElem" maxlength="50">
                                                    &nbsp; <a href="#" onclick="SelMisParty(); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a></td>
                                                        </tr>
                                                        <tr>
                                                    <td width="139" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td width="50%" class="FieldCellSub">Narration</td>
                                                          <td width="50%" class="FieldCellSub">
<%

IF Cstr(sBookCode) = "" Then
	sBookCode = 0
End IF
sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
	" OUDefinitionID='"&sOrgId&"'and BookCode='01' and BookNumber="&sBookCode
'Response.write "<textarea>"& sQuery &"</textarea>"


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
                                                    <a href="#" onclick="showNarration('01'); return false;"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
<%
end if
objRs.Close
%>
                                                           </td>
                                                        </tr>
                                                      </table>
                                                      &nbsp;</td>
                                                    <td class="FieldCell" valign="top"> <textarea rows="3" name="txtNarration" cols="50" class="FormElem" onKeyPress="return ChkEnter(event)"></textarea> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell"> <input type="text" name="txtAmount" size="15" value="0.00" style="text-align:right" maxlength="13" class="FormElem" onblur="popAddAmount();TDSAmount()"> </td><!--popAddAmount()-->
                                                        </tr>
                                                        <tr>
                                                        <td colspan=2>
                                                        <div id="DisCCANL" class=frmBody style="height:1px; visibility: hidden;">
	                                                                <table cellpadding="0" cellspacing="0" >
		                                                                <tr>
			                                                                <td class=MiddlePack colspan="4"> </td>
		                                                                </tr>
		                                                                <tr>
		                                                                    <td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"></td>
			                                                                <td class=FieldCell>
				                                                                <DIV class=frmBody id="DisCost" style="width:260;height:100;">
					                                                                <table border="0" id="tblCost" cellspacing="1" class="ExcelTable">
						                                                                <tr>
							                                                                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								                                                                <td class="ExcelHeaderCell" align="center" width="150">
								                                                                Cost Center Head
								                                                                <img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Click Here to add Cost Center or Analytical Head" onclick="PopCCAH()">
								                                                                </td>
								                                                                <td class="ExcelHeaderCell" align="center">Ratio</td>
								                                                                <td class="ExcelHeaderCell" align="center">Amount</td>
						                                                                 </tr>
					                                                                </table>
				                                                                </div><!--End of CostCenter Display Division -->
			                                                                </td>
			                                                                <td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"></td>
			                                                                <td class=FieldCell>
				                                                                <DIV class=frmBody id="DisAnal" style="width:260; height:100;">

					                                                                <table border="0" id="tblAnal" cellspacing="1" class="ExcelTable">
						                                                                <tr>
								                                                                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								                                                                <td class="ExcelHeaderCell" align="center" width="150">Analytical Head
								                                                                <img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Click Here to add Cost Center or Analytical Head" onclick="PopCCAH()">
								                                                                </td>
								                                                                <td class="ExcelHeaderCell" align="center">Ratio</td>
								                                                                <td class="ExcelHeaderCell" align="center">Amount</td>
					                                                                    </tr>
					                                                                </table>
				                                                                </div>	<!--End of Analytical Display Division -->
			                                                                </td>
		                                                                </tr>
		                                                                <tr>
			                                                                <td class=MiddlePack  colspan="4"></td>
		                                                                </tr>
	                                                                </table>
                                                                </div> <!--End of CCANAL Display Division -->
                                                            </td>
                                                          </tr>
                                                         <tr>
                                                    <td class="FieldCellSub" >Select TDS Group</td>
                                                    <td class="FieldCell" width="591">
                                                    <select size="1" name="SelTDSGrp" class="FormElem" onchange="TDSAmount()">
                                                    <Option Value="0" selected> Select </option>
		                                                  <% Dim sUseable,sGrpID,sTemp

																sQuery = "Select GroupID,GroupName from ACC_M_TDSGroup where OUDefinitionID = '"& sOrgId &"' and isNull(Useable,'Y') <> 'N' "
																	'Response.Write sQuery
																	With objRs1
																		.CursorLocation = 3
																		.CursorType = 3
																		.ActiveConnection = con
																		.Source = sQuery
																		.Open
																	End With
																	Do while Not objRs1.EOF
																		sGrpId = objRs1(0)
																	Response.Write objRs1(1)& "<BR>"%>
																	<option value="<%=objRs1(0)%>" <%'If trim(sGroupName) = trim(objRs1(0)) then Response.Write "selected" %>> <%=objRs1(1)%> </option>
																	<%objRs1.MoveNext
																Loop
																objrs1.Close
															%>
                                                    </select>
                                                    &nbsp; % On Amount &nbsp;
                                                    <input type="text" name="txtTdsAmount" Value="" size="15" style="text-align:right" maxlength="13" class="FormElemRead" readonly >
                                                    <a href="#" onclick="TDSCalc(); return false;"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="TDS Group Selection" width="10" height="11"></a>
                                                     <!--input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled-->
                                                     &nbsp;&nbsp;<input type="Button" value="Add Entry" name="btnAdd" onClick="AddNew()" class="AddButton">
                                                    </td>
                                                        </tr>
<tr>
								<td colspan=2 width="100%" align=center>
                                    <DIV class=frmBody id="DisVoucher" style="width:98%; visibility:hidden; height:1px;">
	                                    <table border="0" cellspacing="1px" id="tblVoucher" class="ExcelTable" style="width:98%;" >
	                                    <tr>
		                                    <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		                                    <td class="ExcelHeaderCell" align="center" width="25"></td>
		                                    <td class="ExcelHeaderCell" align="center" width="25"></td>
		                                    <!--<td class="ExcelHeaderCell" align="center">AU</td>-->
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
							</tr>                                                         <!--tr>
															<td class="FieldCellSub" width="133">Approval</td>
															<td class="FieldCell" width="591">
															<input type="radio" value="Y" checked name="optApprove" class="FormElem">
															Yes&nbsp;&nbsp;
															<input type="radio" value="N" name="optApprove" class="FormElem"> No </td>
														</tr-->
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5px">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
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
								<td align="center" width="5px" class="ClearPixel">
								</td>
								<td >
<DIV class=frmBody id="Disaddtional" style="height:1px; visibility: hidden;">
	<DIV class=frmBody id="DisPayable" style="width: 585px; visibility: hidden; height:1px;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="565px">
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
								<td align="center" class="ClearPixel" width="5px">
								</td>
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
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell" align="center">
													<!--<input type="Button" value="Update Entry" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButtonX" >-->
													<!--<input type="Button" value="Delete Entry" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButtonX" >-->
													<input type="button" value="Save" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
													<!--input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouCABookSelection.asp')" class="ActionButton" -->
													<!--input type="button" value="Delete Voucher" name="btnDelVou" onClick="DelVouch()" class="ActionButtonX" disabled-->
													<!--input type="button" value="Print" name="btnPrnVou" onClick="PrnVouch()" class="ActionButtonX" -->
													<input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('CASHVOUCHERS.ASP')" class="ActionButtonX">
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
