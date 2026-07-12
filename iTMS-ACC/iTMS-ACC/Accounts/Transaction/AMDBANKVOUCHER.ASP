<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AMDBankVoucher.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	June 10, 2005
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
Dim sVouCkTy,sLastVouDt,sCallVouTy,sPopVouTy,sFinPeriod,sFinTemp,sMaxDate,sMinDate
Dim iPreBookVal,objDOM
Dim sFinFrm,sFinTo,sValTemp2,sFormVal,sSelArg
Dim sConStatus,sRetVal

'Response.Write Request.ServerVariables("SCRIPTNAME")

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sFinTemp = Split(sFinPeriod,":")
	sMaxDate = "31/03/"&sFinTemp(1)
	sMinDate = "01/04/"&sFinTemp(0)
End IF

dim sAccount,sAddtional,iSno
dim dTotal
dim sVoucDate,iBookCode,sPayTo,sUserId,sVal,sValTemp
sUserId = getUserID()
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2
sCallVouTy = Request("VouTy")
sSelArg = Request("voutype")
sFormVal = Request("hFormVal")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("horgName")
sOrgId   = session("organizationcode")
sOrgName = session("orgshortname")

sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("selVouTypepe")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")

iPreBookVal = sBookCode
sConStatus = GetContraStatus(sTransNo)
'Response.Write "sTransNo="&sTransNo
sVal=Request("Val")
sValTemp=Split(sVal,"~")

IF Cstr(sVal) <> "" Then
	sQuery = "Select H.OUDEFINITIONID,D.OrgUnitDescription From DCS_OrganizationUnitDefinitions D, "&_
		 "Acc_T_CreatedVoucherHeader H Where H.OUDEFINITIONID = D.OUDEFINITIONID "&_
		 "and H.CreatedTransNo = "&sValTemp(0)&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sOrgId = objRs(0)
		sOrgName = objRs(1)
	End IF
	objRs.Close

Else

	sQuery = "Select Top 1 OUDefinitionID,OrgUnitDescription From DCS_OrganizationUnitDefinitions "&_
			 "Where Len(OUDefinitionID) > 4 Order By OUDefinitionID "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sOrgId = objRs(0)
		sOrgName = objRs(1)
	End IF
	objRs.Close
End IF

IF CStr(iBookAccHead) = "" Then
	sQuery = "Select Top 1 BookNumber,BookName,isNull(BookAccountHead,0),OtherUnitTransaction From vwOrgBookNames Where  "&_
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '02' Order By BookName "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sBookCode = objRs(0)
		sBookName = objRs(1)
		iBookAccHead = objRs(2)
		bOtherUnits = objRs(3)
	Else
		sBookCode = "02"
		sBookName = ""
		iBookAccHead = 0
		bOtherUnits = 1
	End IF
	objRs.Close
End IF

'Response.Write sTransNo


oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

'Response.Write sTransNo

'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_BA_"&Session.SessionID&".xml")

IF CStr(sTransNo) <> "" Then
 	sRetVal = GetVouchXML(sTransNo)
	oDOM.Load server.MapPath(sRetVal)
	oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_BA_"&Session.SessionID&".xml")
End IF


IF CStr(sTransNo) = "" Then
	sTransNo = 0
End IF

IF CStr(sCallVouTy) = "R" Then
	sVouType = "D"
	sVouCkTy = "BAP"
	sPopVouTy = "D"
Else
	sVouType = "C"
	sVouCkTy = "BAR"
	sPopVouTy = "C"
End IF

'IF CStr(sVouType) = "C" Then
'	sVouCkTy = "BAP"
'Else
'	sVouCkTy = "BAR"
'End IF

sQuery = "Select Convert(Char,VoucherDate,103) From Acc_T_CreatedVoucherHeader Where CreatedTransNo =  "&_
		 "(Select Max(CreatedTransNo) From Acc_T_CreatedVoucherHeader Where BookCode = '02'  "&_
		 "and OUDefinitionID = '"&sOrgId&"' and TransactionType ='"&sVouCkTy&"' ) "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sLastVouDt = Trim(objRs(0))
End IF
objRs.Close

'sQuery = "Select CreatedBy From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sTransNo&" "
'objRs.Open sQuery,Con
'IF Not objRs.EOF Then
'	sUserId = Trim(objRs(0))
'End IF
'objRs.Close


'IF CStr(sUserId) = "" Then
sUserId = session("userid")
'End IF

sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp2(0))
sFinTo = Trim(sValTemp2(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"

Dim TempNode
Set Root = objDOM.CreateElement("Root")
objDOM.appendChild Root
Set Entrynode = objDOM.CreateElement("Details")
Root.appendChild Entrynode
'****************** This Voucher that is been Adjusted in Any Purchase/Sales Vouchers Will be
'****************** taken here for consideration *********************************************

sQuery = "Select H.CreatedVoucherNo,Convert(Varchar,H.VoucherDate,103) VoucherDate,H.VoucherAmount,C.AmountAdjusted, "&_
		 "A.CreatedAdvanceNo,C.CreatedTransNo,H.CrDrIndication From Acc_T_CreatedVoucherHeader H, "&_
		 "ACC_T_CreatedAdvanceAdj C,Acc_T_CreatedAdvances A Where A.CreatedTransNo = "&sTransNo&" "&_
		 "And A.CreatedAdvanceNo = C.CreatedAdvanceNo and C.CreatedTransNo = H.CreatedTransNo  "

'Response.Write sQuery
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountAdjusted")
	TempNode.setAttribute "AdjCrTrNo", objRs("CreatedTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CreatedAdvanceNo")
	TempNode.setAttribute "ValFrm", "CRADJADV"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close

sQuery = "Select CreatedVoucherNo,Convert(Varchar,VoucherDate,103) VoucherDate,VoucherAmount, "&_
		 "AmountPaid,AdjCrTransNo,CradvNo From CrPayAdvAdjDet Where CreatedTransNo = "&sTransNo
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountPaid")
	TempNode.setAttribute "AdjCrTrNo", objRs("AdjCrTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CradvNo")
	TempNode.setAttribute "ValFrm", "Pay"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close

sQuery = "Select CreatedVoucherNo,Convert(Varchar,VoucherDate,103) VoucherDate,VoucherAmount, "&_
		 "AmountReceived,AdjCrTransNo,CradvNo From CrRecAdvAdjDet Where CreatedTransNo = "&sTransNo
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountReceived")
	TempNode.setAttribute "AdjCrTrNo", objRs("AdjCrTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CradvNo")
	TempNode.setAttribute "ValFrm", "Rec"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close


objDOM.Save server.MapPath("../temp/transaction/Voucher ADJAMD_BA_"&Session.SessionID&".xml")



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS Bank Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script src="../../scripts/ExcelFunctions.js"></script>
<script src="../../scripts/VouSelection.js"></script>
<script src="../../scripts/VoucherEntryCore.js"></script>
<script src="../../scripts/BankVoucher.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>

<!--XML ISLAND FOR VOUCHER DATA -->
<script type="application/xml" data-itms-xml-island="1" id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=sBookCode%>" BookName="<%=sBookName%>" CRDR="<%=sVouType%>" VouDate="" BookAcchead="0" Approver=""/></script>
<!--XML ISLAND FOR ENTRY DATA -->
<script type="application/xml" data-itms-xml-island="1" id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="0" TDSElgi="0" TdsPercentage="0" PayRecAmount="0" /></script>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="TDSData"  ><Root/></script>
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
<script type="application/xml" data-itms-xml-island="1" id="IUTUnits"></script>
<script type="application/xml" data-itms-xml-island="1" id="InsDet"></script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyHeadData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="AdjDet" data-src="<%="../Temp/Transaction/Voucher ADJAMD_BA_"&Session.SessionID&".xml"%>"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();DisplayBook();SelUnBook()">
<form method="POST" name="formname" action="AmdAccGenerate.asp" >
<input type="hidden" name="hVouCode" value="02">
<input type="hidden" name="hVouType" value="<%=sCallVouTy%>">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="BA">
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
<input type="hidden" name="hPopVouTy" value="<%=sPopVouTy%>">
<input type="hidden" name="hMinDate" value="<%=sMaxDate%>">
<input type="hidden" name="hMaxDate" value="<%=sMinDate%>">
<input type="hidden" name="hBookOtherUnit" value="1">
<input type="hidden" name="hPreBookSel" value="<%=iPreBookVal%>">
<input type="hidden" name="hInsDet" value="">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">

<input type="hidden" name="hFormVal" value="<%=sFormVal%>">
<input type="hidden" name="voutype" value="<%=sSelArg%>">
<input type="hidden" name="ContraStatus" value="<%=sConStatus%>">
<input type="hidden" name="hCurrConStatus" value="0">
<input type="hidden" name="hNorToContra" value="N">

<input type="hidden" name="hUpdate" value="N">
<input type="hidden" name="hTdsNew" value="N">

<input type="hidden" name="hInsrAmt" value="0">
<input type="hidden" name="hAction" value="New">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<% IF CStr(sCallVouTy) = "R" Then
				Response.Write("Bank Receipt Voucher Amendment")
		   Else
				Response.Write("Bank Payment Voucher Amendment")
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
								<td align="center" colspan="3" height="8">
                                  <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                   <tr>
                                      <td class="MiddlePack" colspan="6"></td>
                                    </tr>
                                    <tr>
                                       <td class="FieldCellSub" width="110">Book</td>
										<td class="FieldCell" colspan="3">
											<select size="1" name="selBook" class="FormElem" onChange="SetBookAccHead()">
												<!--option value="S">Select</option-->
											</select>
										</td>
										<td class="FieldCellSub" width="90">Voucher Type</td>
										<td class="FieldCellSub" colspan="2">
										<input type=hidden name=hAccUnit value="<%=sOrgId%>">
											<select size="1" name="selVouType" class="FormElem" onChange="CheckEntryType(this)" disabled>
												<option value="S">Select</option>
												<%IF CStr(sVouType) = "D" Then %>
													<option value="D" Selected>Receipt</option>
													<option value="C">Payment</option>
												<%Else%>
													<option value="D">Receipt</option>
													<option value="C" Selected>Payment</option>
												<%End IF %>
											</select>
										</td>

										</tr>
										
										<tr>

											<td class="FieldCellSub" width="110">Voucher No
											<td class="FieldCell" colspan="3">
												<%if Trim(sVal)<>"" then%>
													<input type="text" name="txtVouNo" size="20" class="FormElemRead" value="<%=sValTemp(1)%>" readonly>
												<%else%>
													<input type="text" name="txtVouNo" size="20" class="FormElemRead" readonly>
												<%end if%>
											</td>
											<!--td class="FieldCellSub">
												<a href="#" onclick="popVoucherNo('C','02','BA',document.formname.selUnitId.value,document.formname.hPopVouTy.Value,document.formname.selBook.value); return false;">
												<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
												 &nbsp;&nbsp;
												 <a href="#" onclick="popVoucherNo('A','02','BA',document.formname.selUnitId.value,document.formname.hPopVouTy.Value,document.formname.selBook.value); return false;">
												<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers"></a>
											</td-->
                                      <td class="FieldCellSub" width="90">Date</td>
                                      <td class="FieldCell" width="110">
									      <%'Function Call to Insert Date Picker
											Response.Write InsertDatePicker("ctlDate")%>
 									  </td>
								</tr>
								
								<tr>
									<td class="FieldCellSub" width="110">Entry No</td>
									<td class="FieldCell" colspan="3"><span class="DataOnly" id="spEntryNo">1&nbsp;</span>
									</td>
								</tr>
								<tR>
									<td class="FieldCellSub" width="110">&nbsp;</td>
									<td class="FiledCell">
										<Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()">
									</td>
									<td class="FieldCell" colspan=5>
								    <span id="spInsDet" class="DataOnly">&nbsp;</span>
								    &nbsp; <a href="#" onclick="PopInsDet(); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Enter Instrument Details"></a>
								</td>
								</tr>
								
								<!--<tr>-->
									<!--td class="FieldCellSub" >
										<Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()">
									</td-->
									<!--<td class="FieldCellSub" colspan="2">Entry No &nbsp; <span class="DataOnly" id="spEntryNo">1&nbsp;</span>-->
									<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()" -->
									<!--</td>
								</tr>-->
								<tr>

									<td class="FieldCellSub">Current Balance</td>
									<td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spCurrBal">
                                <%

										dOpeningBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
										'dOpeningBal = 0
										dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
										if dOpeningBal<0 then
											Response.Write FormatNumber(dOpeningBal*-1,2,,,0) &"&nbsp;Cr"
										else
											Response.Write FormatNumber(dOpeningBal,2,,,0) &"&nbsp;Dr"
										end if
                                %> </span>
                                &nbsp;</td>
                                <td class="FieldCellSub">Book Balance</td>
								<td class="FieldCellSub">
									<span class="DataOnly" id="spBookBal">
									<%
									dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
									'dOpeningBal = 0
									dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
									if dOpeningBal<0 then
									Response.Write FormatNumber(dOpeningBal*-1,2,,,0) &"&nbsp;Cr"
									else
									Response.Write FormatNumber(dOpeningBal,2,,,0) &"&nbsp;Dr"
									end if
									%> </span>&nbsp;
								</td>

								</tr>
								<tr id="BankDet" style="visibility: hidden;">
									<td class="FieldCellSub">Inst No</td>
									<td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spInsNo">
									<td class="FieldCellSub">Inst Date</td>
									<td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spInsDT">

								</tr>

								<tr>
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
                                                    <td class="MiddlePack" colspan="2" ></td>
                                                        </tr>
                                                        <!--<tr>
                                                    <td class="FieldCellSub">Accounting Unit</td>
                                                    <td class="FieldCellSUB">

                                                     <select size="1" name="selAccUnitId" onChange="popAccHead()" class="FormElem">
														<option value="A" >Account Unit</option>
														<%'populateOrganizationListDBWithVal(sOrgID)%>
														<OPTION value="<%=sOrgId%>" Selected><%=sOrgName%></option>
													</select>
													<%
														'else
														'	Response.Write sOrgName
														'end if
													%>
													</td>-->
													<!--td class="FieldCellSub">Entry No
													&nbsp; <span class="DataOnly" id="spEntryNo">1&nbsp;</span>
													</td-->
                                                        <!--</tr>-->
                                                        <tr>
                                                        <td class="FieldCellSub" width="110">Entry Type
									<td class="FieldCell" colspan="2">
									<% IF CStr(sCallVouTy) = "R" Then %>
										<input type=radio name="selCRDR" value="C" checked>Receipts
										<input type=radio name="selCRDR" value="D" disabled>Payments&nbsp;
                                    <%Else%>
										<input type=radio name="selCRDR" value="C"disabled>Receipts
										<input type=radio name="selCRDR" value="D" checked>Payments&nbsp;
									<%End IF %>
									</td>
									</tr>
									
									<tr>
                                        <td class="FieldCellSub" >Accounting Head</td>
                                        <td class="FieldCellSub" colspan="2">
                                                <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
												<option value="A">Select Account Head</option>
												<%
													dim iHeadCount
												 	'iHeadCount=popFrequentHead(sOrgId,"02",sBookCode)
													iHeadCount=0
												%>
													<option value="G">General Ledger</option>
												<%populatePartyType(sOrgId)%>
                                        </select> &nbsp; <a href="#" onclick="selAccountHead(document.formname.selAccHead); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Account Head"></a>
                                        </td>
                                        <input type="hidden" name="hHeadCount" value="<%=iHeadCount%>">
									</tr>
									
                                    <tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td class="FieldCellSub"><span class="DataOnly" id="spAccHead"></span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" >Pay to / Received from</td>
                                                    <td class="FieldCellSub" colspan="2"> <input type="text" name="txtPayTo" size="55" class="Formelem" maxlength="50">
                                                    &nbsp; <a href="#" onclick="SelMisParty(); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a></td>
                                                        </tr>
                                                        <tr>
                                                    <td width="139" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td class="FieldCellSub">Narration</td>
                                                          <td class="FieldCellSub">
<%

'sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
'	" OUDefinitionID='"&sOrgId&"'and BookCode='02' and BookNumber="&sBookCode

'with objRs
'	.CursorLocation = 3
'	.CursorType = 3
'	.Source = sQuery
'	.ActiveConnection = con
'	.Open
'end with
'set objRs.ActiveConnection = nothing

'if objRs(0)>0 then
%>
                                                            <p align="left">
                                                    <a href="#" onclick="showNarration('02'); return false;"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
<%
'end if
'objRs.Close
%>
                                                           </td>
                                                        </tr>
                                                      </table>
                                                      &nbsp;</td>
                                                    <td class="FieldCellSub" valign="top"> <textarea rows="2" name="txtNarration" cols="50" class="FormElem" onKeyPress="return ChkEnter(event)"></textarea> </td>
                                                        </tr>
                                                        <tr>
                                                  <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtAmount" size="15" value="0.00" style="text-align:right" maxlength="13" class="Formelem" onblur="popAddAmount();TDSAmount()"> </td>
                                                        </tr>
                                                        <tr>
                                                        <td colspan="2">
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
						                                                                <td class="ExcelHeaderCell" align="center" width="150">Cost Center Head
						                                                                <img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Click Here to add Cost Center or Analytical Head" onclick="PopCCAH()">
						                                                                </td>
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
	                                                                <td class=MiddlePack  colspan="3"></td>
                                                                </tr>
                                                            </table>
                                                        </div> <!--End of CCANAL Display Division -->
                                                        </td>
                                                        </tr>
                                                         <tr>
                                                    <td class="FieldCellSub" >Select TDS Group</td>
                                                    <td class="FieldCellSub" width="591">
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
                                                    &nbsp; % On Amount &nbsp; <input type="text" name="txtTdsAmount" Value="" size="15" style="text-align:right" maxlength="13" class="Formelemread" readonly >
                                                    <a href="#" onclick="TDSCalc(); return false;"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="TDS Group Selection" width="10" height="11"></a>
                                                    <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
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
                                                                <input type="Button" value="Update Entry" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButtonX" >
                                                                <input type="Button" value="Delete Entry" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButtonX" >
                                                                <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Delete Voucher" name="btnDelVou" onClick="DelVouch()" class="ActionButtonX" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('BANKVOUCHERS.ASP')" class="ActionButton">

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