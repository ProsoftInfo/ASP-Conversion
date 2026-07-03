<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CreditNoteNewPage.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	S.Maheswari
	'Created On					:	sep 16, 2008
	'Modified By				:	Ragavenran R
	'Modified On				:	Oct 25,2011
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<!-- XML Data Island -->
<XML id="UnitBook"></XML>
<XML ID="UnitBookData"><Book/></XML>
<XML ID="OutData"><PartyType/></XML>
<XML id="AccHeadData">
<account/>
</XML>
<xml id="PartyData"><Root /></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/VouTransactions.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<script language="javascript">
window.__itmsCndnNewPageConfig = { mode: "creditCreate", unitField: "hUnitID", validateUnitFromSelect: false, unitBookXml: "UnitBook", nonGjBookCode: "07" };
</script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/CreditDebitNoteCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="FnInit();popPartyType()">
<%

	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

	Dim sFinPeriod,Objrs,Objrs1,Objrs2,iCnt,sSql,iCrTransNo,sOptType
	Dim dcrs,sUnitLID,sUnitLName,sUnitSName,sBankType,AccVoucherNo
	Dim sFormVal,sTemparr,sUnitID,sBookNo,sFrmDate,sToDate,sFrmAmt,sToAmt
	Dim sChkSR,sChkSI,sChkPI,sChkSC,sChkMI
	Dim sSelVouTy,sOrgId,sPartyType,sPartypeName,sOrgName,sPayToRecdFrom
	Dim dtFromDate,dtToDate

	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")

	'sOrgId = Request("selUnitId")
	sOrgId = Session("organizationcode")
	sOrgName = Session("OrgShortName")
	'sPartyType = Request("selPartyType")
	sPartyType = Request.form("hPartyCode")
	IF trim(Request.form("hPartyCode")) <> "" then
		sTemp  = split(Request.form("hPartyCode"),"?")
		sPartypeName = sTemp(2)
		sParName =  Request.form("hParName")
		'Response.Write "PartyType="&sPartyType &"--"& sParName
	End if
	If trim(sOrgId) = "" or  trim(sOrgId) = "0" then sOrgId = "010101"
	sSelVouTy = Request("voutype")
	'Response.Write "sSelVouTy="&sSelVouTy
	IF trim(sSelVouTy) = "" then
		sSelVouTy = "SI"
		sChkSI = "Checked"
		sChkSR = ""
		sChkPI = ""
		sChkSC = ""
	End IF
	IF trim(sSelVouTy) = "SI" then
		sChkSI = "Checked"
		sChkSR = ""
		sChkPI = ""
		sChkSC = ""
		sChkMI = ""
	ElseIF trim(sSelVouTy) = "SC" then
		sChkSC = "Checked"
		sChkSR = ""
		sChkSI = ""
		sChkPI = ""
		sChkMI = ""
	ElseIF trim(sSelVouTy) = "SR" then
		sChkSR = "Checked"
		sChkSI = ""
		sChkPI = ""
		sChkSC = ""
		sChkMI = ""
	ElseIF trim(sSelVouTy) = "PI" then
		sChkPI = "Checked"
		sChkSI = ""
		sChkSR = ""
		sChkSC = ""
		sChkMI = ""
	ElseIF trim(sSelVouTy) = "MI" then
	    sChkMI = "Checked"
		sChkPI = ""
		sChkSI = ""
		sChkSR = ""
		sChkSC = ""
	End IF
	'Response.Write "VouType="&sSelVouTy
	iCurrentPage=CInt(Request.Form("hPageSelection"))
	dtFromDate = Request.Form("hFromDate")
	dtToDate = Request.Form("hToDate")

	if trim(dtFromDate)="" then
		dtFromDate = "01/04/"& split(session("FinPeriod"),":")(0)
	end if
	if trim(dtToDate)="" then
		dtToDate = FormatDate(date)
	end if
	'iCnt=Request.Form("hCnt")

%>
	<form method="POST" name="formname" action="CreditNoteToCreate.asp">

	<input type=hidden name="TransNo" value="">
	<input type=hidden name="hTransNo" value="">
	<input type=hidden name="hInvVal" value="<%=sSelVouTy%>">
	<input type=hidden name="hUnitID" value="<%=sOrgID%>">
	<input type=hidden name="hUnitName" value="<%=sOrgName%>">
	<input type=hidden name="hPartyCode" value="">
	<input type=hidden name="hParType" value="">
	<input type=hidden name="hSubParType" value="">
	<input type=hidden name="hParCode" value="">
	<input type=hidden name="hParName" value="<%=sParName%>">
	<input type=hidden name="hPartypeName" value="<%=Request.form("hPartyCode")%>">
	<input type=hidden name="hBookNo" value="">
	<input type=hidden name="selInvoiceNo" value="">
	<input type=hidden name="selBook" value="">
	<input type=hidden name="hVouType" value="">
	<input type=hidden name="hChkVal" value="">
	<input type=hidden name="hVouDetails" value="">
	<input type=hidden name="hBookName" value="">
	<input type=hidden name="hFromDate" value="<%=dtFromDate%>">
	<input type=hidden name="hToDate" value="<%=dtToDate%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Credit Vouchers
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="center" class="SubTitle">&nbsp;&nbsp;
<Input type=radio name=voutype value="SI" <%=sChkSI%> onclick=setInvoiceNo(this)>Sal Inv.&nbsp;
<Input type=radio name=voutype value="SC"  <%=sChkSC%>  onclick=setInvoiceNo(this) >Sal Comm.&nbsp;
<Input type=radio name=voutype value="SR" <%=sChkSR%> onclick=setInvoiceNo(this) >Sal Return&nbsp;
<Input type=radio name=voutype value="PI" <%=sChkPI%> onclick=setInvoiceNo(this) >Pur Invoice&nbsp;
<Input type=radio name=voutype value="MI" <%=sChkMI%> onclick=setInvoiceNo(this) >Misc Receipts&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
<td width="100%">
<div id="idUnprocessed" style="width: 575px; display: none">
<table cellpadding="0" cellspacing="0">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<!--<tr>
	<td class="FieldCellSub">Unit Name</td>
	<td class="FieldCellSub" colspan="4">
	 <select size="1" name="selUnitId" class="FormElem"> <%'onChange="DisplayBook()"%>

	  	<%populateOrganizationListDB%>
    </select> &nbsp;
	</td>
</tr>-->
<tr>
<td class="FieldCellSub" width="168px">Select Party Type</td>
<td class="FieldCellSub" colspan="3">
<select size="1" name="selPartyType" class="FormElem" onChange="selParty(this)">
	<option value="S">Select Party Type</option>
	</select>
	</td>
    </tr>
    <tr>
<td class="FieldCellSub" width="108px">Party Name</td>
<td class="FieldCellSub" colspan="3"> <input type="text" name="txtPartyName" size="40" class="FormElem"></td>
</tr>
  <tr>
<td class="FieldCellSub" width="108px">From Date</td>
<td class="FieldCellSub" ><%Response.Write insertdatepicker("ctlFromDate")%></td>
<td class="FieldCellSub" align="right" width="108">To Date</td>
<td class="FieldCellSub" ><%Response.Write insertdatepicker("ctlToDate")%></td>
</tr>
<tr>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell" colspan="2">
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()" >
</td>
</table>
</div>
</td>
</tr>
<tr>
<td align="center" class="MiddlePack">
</td>
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
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
<table border="0" cellspacing="1px" class="ExcelTable" width="100%">
<tr>
<td class="ExcelHeaderCell" width="10px" >S.No.
</td>
<td class="ExcelHeaderCell" width="10px" >
</td>
<td class="ExcelHeaderCell">Number
</td>
<td class="ExcelHeaderCell">Date
</td>
<%
    if Trim(sSelVouTy)="MI" then
        Response.Write "<td class='ExcelHeaderCell' align='Center'>Recd From</td>"
    end if
%>

<td class="ExcelHeaderCell">Party Name
</td>
<%if trim(sSelVouTy)="SI" then%>
	<td class="ExcelHeaderCell"><%Response.Write "Invoice Value"%></td>
<%elseif trim(sSelVouTy)="SC" then%>
	<td class="ExcelHeaderCell"><%Response.Write "Commission Value"%></td>
<%end if%>
<!--td class="ExcelHeaderCell" align="center" >Amount
</td>
<td class="ExcelHeaderCell" align="center" >Status
</td-->
</tr>

<SCRIPT LANGUAGE=vbscript RUNAT=Server>

</SCRIPT>
<%
Dim sQuery,sFromApp,sVouStatus,sTemp,sPartyName
Dim sParType,iParCode,iSubParType,sParName
iCnt = 0
Const iPageSize=15
'Response.Write "Test="&sPartyType
IF trim(sPartyType) = "" or trim(sPartyType) = "S" then
	IF trim(sSelVouTy) = "SI" then 'Sales Invoice
		sQuery = "select  distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
				 "H.PayToRecdFrom from Acc_T_VoucherHeader H, Acc_T_CreatedReceivables P where H.OUDefinitionID='"&sOrgId&"' "&_
				 "and H.BookCode='05' and P.AmountReceivable > P.AmountReceived and Convert(datetime,H.VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103) "&_
				 "and Convert(datetime,H.VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103)  Order By 5 "
	ElseIF trim(sSelVouTy) = "SC" then 'Sales Commission
		sQuery="select VA.AccTransactionNo,VA.VoucherNumber,convert(char,VA.VoucherDate,103),VA.AgentCommission,VA.AgentNAme,VA.CommissionType,VA.CommissionToPay,VS.CreatedTransNo,VA.AgentCode,VA.CurrencyCode from "&_
		" VwAgentCommisionDetails VA,VwSalCommAccDet VS where OUDefinitionID='"&sOrgId&"' and Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103) "&_
		"and Convert(datetime,VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103) and VS.TransactionNumber=VA.AccTransactionNo  and VA.CommissionToPay = 1" 'CommisionToPay 0 case should be check
	ElseIF trim(sSelVouTy) = "SR" then 'Sales Return
	sFromApp = 3
	sVouStatus = "010104"
		sQuery = "Select distinct H.CreatedTransNo,H.CreatedVoucherNo,convert(char,H.VoucherDate,103),H.VoucherAmount,H.PayToRecdFrom,P.ReceivableNumber "&_
				 "from Acc_T_CreatedVoucherHeader H,Sal_T_SalesReturnHeader S,Acc_T_CreatedReceivables P,Sal_T_SalesReturnAction A  Where H.BookCode = '05' "&_
				 "and H.CreatedVouchStatus = '"& sVouStatus &"' and H.FromApplication = "&sFromApp&" and S.SaleTransactionNo = H.OtherApplnTransNo "&_
				 "and H.CreatedTransNo = P.CreatedTransNo and P.AmountReceivable > P.AmountReceived and S.SalesReturnNo = A.SalesReturnNo and A.ActionNo = 11 "&_
				 "and Convert(datetime,H.VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103) "&_
				 "and Convert(datetime,H.VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103)  Order By 1 "
	ElseIF trim(sSelVouTy) = "PI" then 'Purchase Invoice
		sQuery = "select Distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
	  			 "H.PayToRecdFrom,P.PayablesNumber from Acc_T_VoucherHeader H, Acc_T_Payables P "&_
				 "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='04' and H.TransactionNumber = P.TransactionNumber "&_
				 "and P.AmountPayable > P.AmountPaid and Convert(datetime,H.VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103) "&_
				 "and Convert(datetime,H.VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103)  Order By 1 "
    Elseif Trim(sSelVouTy) = "MI" then  ' Misc Payment
        sQuery = "Select V.CreatedTransNo,V.VoucherNumber,Convert(varchar,V.VoucherDate,103), "&_
                " V.VoucherAmount,V.PayToRecdFrom,M.MiscTransNo from Acc_T_MiscPymtRequestHeader M,Acc_T_VoucherHeader V"&_
                " where M.ReceiptNo = V.CreatedTransNo and V.OUDefinitionID = '"& sOrgId &"' and M.ApplicationCode = 3 and isNull(M.AdjustmentStatus,'N')='N'"
	End IF
ElseIF trim(sPartyType) <> "" then
 'Response.Write sPartyType
	sTemp = split(sPartyType,"?")
	sParType = sTemp(0)
	iSubParType = sTemp(1)
	sParName = sTemp(2)
	iParCode = sTemp(3)

	IF trim(sSelVouTy) = "SI" then 'Sales Invoice
		sQuery = "select  distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
				 "H.PayToRecdFrom from Acc_T_VoucherHeader H, Acc_T_CreatedReceivables P where H.OUDefinitionID='"&sOrgId&"' "&_
				 "and H.BookCode='05' and P.AmountReceivable > P.AmountReceived and H.PartyType = '"& sParType &"' and "&_
				 "H.PartySubType= '"&iSubParType &"' and H.PartyCode = "& iParCode & " "&_
				 "and Convert(datetime,H.VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103)"&_
				 "and Convert(datetime,H.VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103)  Order By 5 "
	ElseIF trim(sSelVouTy) = "SC" then 'Sales Commission
	sQuery="select VA.AccTransactionNo,VA.VoucherNumber,convert(char,VA.VoucherDate,103),VA.AgentCommission,VA.AgentNAme,VA.CommissionType,VA.CommissionToPay,VS.CreatedTransNo,VA.AgentCode,VA.CurrencyCode from "&_
		" VwAgentCommisionDetails VA,VwSalCommAccDet VS where OUDefinitionID='"&sOrgId&"'  and AgentCode="& iParCode &" and Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103) "&_
		"and Convert(datetime,VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103) and VS.TransactionNumber=VA.AccTransactionNo and VA.CommissionToPay = 1" 'CommisionToPay 0 case should be check
	ElseIF trim(sSelVouTy) = "SR" then 'Sales Return
	sFromApp = 3
	sVouStatus = "010104"
		sQuery = "Select distinct H.CreatedTransNo,H.CreatedVoucherNo,convert(char,H.VoucherDate,103),H.VoucherAmount,H.PayToRecdFrom,P.ReceivableNumber "&_
				 "from Acc_T_CreatedVoucherHeader H,Sal_T_SalesReturnHeader S,Acc_T_CreatedReceivables P,Sal_T_SalesReturnAction A  Where H.BookCode = '05' "&_
				 "and H.CreatedVouchStatus = '"& sVouStatus &"' and H.FromApplication = "&sFromApp&" and S.SaleTransactionNo = H.OtherApplnTransNo "&_
				 "and H.CreatedTransNo = P.CreatedTransNo and P.AmountReceivable > P.AmountReceived  and H.PartyType = '"& sParType &"' and "&_
				 "H.PartySubType= '"&iSubParType &"' and H.PartyCode = "& iParCode & "	 and S.SalesReturnNo = A.SalesReturnNo and A.ActionNo = 11 "&_
				 " and Convert(datetime,H.VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103) "&_
				 " and Convert(datetime,H.VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103)  Order By 1 "
	ElseIF trim(sSelVouTy) = "PI" then 'Purchase Invoice
		sQuery = "select Distinct H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
	  			 "H.PayToRecdFrom,P.PayablesNumber from Acc_T_VoucherHeader H, Acc_T_Payables P "&_
				 "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='04' and H.TransactionNumber = P.TransactionNumber "&_
				 "and P.AmountPayable > P.AmountPaid  and H.PartyType = '"& sParType &"' and "&_
				 "H.PartySubType= '"&iSubParType &"' and H.PartyCode = "& iParCode & ""&_
				 " and Convert(datetime,H.VoucherDate,103) >= Convert(datetime,'"&dtFromDate  &"',103) "&_
				 "and Convert(datetime,H.VoucherDate,103) <= Convert(datetime,'"&dtToDate  &"',103)  Order By 1 "
    Elseif Trim(sSelVouTy) = "MI" then  ' Misc Payment
        sQuery = "Select V.CreatedTransNo,V.VoucherNumber,Convert(varchar,V.VoucherDate,103), "&_
                " V.VoucherAmount,V.PayToRecdFrom,M.MiscTransNo from Acc_T_MiscPymtRequestHeader M,Acc_T_VoucherHeader V"&_
                " where M.ReceiptNo = V.CreatedTransNo and V.OUDefinitionID = '"& sOrgId &"' and M.ApplicationCode = 3 and isNull(M.AdjustmentStatus,'N')='N'"
	End IF
End IF
   'Response.Write sQuery
	with Objrs
	.ActiveConnection=con
	.CursorLocation=3
	.CursorType=3
	.Source=sQuery
	.Open
	end with
	set Objrs.ActiveConnection=Nothing
	IF  not Objrs.EOF then

	'******* Start of Paging
	Objrs.PageSize=iPageSize
	if iCurrentPage=0 then iCurrentPage=1
	Objrs.AbsolutePage=iCurrentPage
	iTotalPage=objrs.PageCount
	For iPageCtr=1 to objrs.PageSize
	'	Do while not Objrs.EOF
	IF trim(sSelVouTy) <> "SC" then
		sSql = "select PartyName from App_M_PartyMaster where PartyCode in (Select PartyCode from Acc_t_CreatedVoucherHeader where CreatedTransno = "& trim(Objrs(0)) &") "
		Objrs1.Open sSql,con
		If not Objrs1.EOF then
			sPartyName = Objrs1(0)
		End If
		Objrs1.Close
	Else
		sPartyName = Objrs(4)
	End IF

		if Trim(sSelVouTy)="MI" then
		    sQuery = "Select PayToRecdFrom from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& trim(Objrs(0))
		    Objrs1.Open sQuery,con
		    if not Objrs1.EOF then
		        sPayToRecdFrom = trim(Objrs1(0))
		    end if
		    Objrs1.Close

		    sQuery = "Select PartyName from App_M_PartyMaster where PartyCode in("&_
		             "Select PartyCode from Acc_T_MiscPymtRequestHeader where ReceiptNo ="&Trim(Objrs(0))&")"
		    objrs1.Open sQuery,con
		    if not Objrs1.EOF then
		        sPartyName = trim(Objrs1(0))
		    end if
		    objrs1.Close
		end if 'if Trim(sSelVouTy)="MI" then


	   iCnt = iCnt + 1
	%>
<tr>
<td class="ExcelSerial" align="center"><%=iCnt%></td>
<td class="ExcelDisplayCell" align="center" width="10">
<%if trim(sSelVouTy)<>"SC" then%>
<input type="Checkbox"  name="OptCriteria<%=iCnt%>"  value="<%=trim(Objrs(0))%>:<%=trim(Objrs(1))%>:<%=trim(Objrs(2))%>:<%=trim(Objrs(4))%>" onclick="SetVal(this)">
<%else%>
<input type="Checkbox"  name="OptCriteria<%=iCnt%>"  value="<%=trim(Objrs(7))%>" onclick="SetVal(this)">
<input type="hidden"  name="hOptCriteria<%=iCnt%>"  value="<%=trim(Objrs(0))%>:<%=trim(Objrs(1))%>:<%=trim(Objrs(2))%>:<%=trim(Objrs(3))%>:<%=trim(objrs(4))%>:<%=trim(objrs(5))%>:<%=trim(objrs(6))%>:<%=trim(objrs(7))%>:<%=trim(objrs(8))%>:<%=trim(objrs(9))%>">
<%end if%>
<td class="ExcelDisplayCell" align="left" ><a href="#" LANGUAGE="VBSCRIPT" onclick="ShowVouch('<%=trim(Objrs(0))%>')" class="ExcelDisplayLink"><%=Objrs(1)%></a></td>
<td class="ExcelDisplayCell" align="left" ><%=Objrs(2)%></td>
<%
    if Trim(sSelVouTy)="MI" then
        Response.Write "<td class='ExcelDisplayCell' align='left'>"& sPayToRecdFrom &"</td>"
    end if
%>
<td class="ExcelDisplayCell" align="left" ><%=sPartyName%></td>

<%if trim(sSelVouTy)="SI" then%>
	<td class="ExcelDisplayCell" align="left" ><%=FormatNumber(Objrs(3),2,,,-1)%></td>
<%elseif trim(sSelVouTy)="SC" then%>
	<td class="ExcelDisplayCell" align="left" ><%=FormatNumber(Objrs(3),2,,,-1)%></td>
<%end if%>
<!--td class="ExcelDisplayCell" align="right" ></td-->
</tr>
<%
		Objrs.MoveNext
	'loop
	if Objrs.EOF then exit for
	next
	End IF

	Objrs.Close

%>
</table>
<!--/div-->
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
</td>
<td valign="top" align="right">
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hCnt" value=<%=iCnt%>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
		if iCurrentPage = 1 then
%>
			<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
			<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
	<%	else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
	<%	end if%>

<SELECT class="FormElem" onChange="Paginate(this(this.selectedIndex).value)" id=select1 name=select1>
	<%
	For lnPage = 1 To iTotalPage
		If lnPage = iCurrentPage Then
	%>
			<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
	<%	else%>
			<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
	<%	end if
	next
	%>
</SELECT>
	<%if iCurrentPage = iTotalPage then%>
		<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
		<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>
	<%else%>
	<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
	<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotalPage%>')" id=button8 name=button8>
	<%end if
End If
%>
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
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td class="ActionCell">
<input type="button" value="Create GJ Voucher" name="B9" class="ActionButtonX" tabindex="3" onclick="Voucher('GJ')">
<%if trim(sSelVouTy)<>"SC" then%>
	<input type="button" value="Create CR Voucher" name="B10" class="ActionButtonX" tabindex="3" onclick="Voucher('CR')">
<%end if%>

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
</body>
</html>

