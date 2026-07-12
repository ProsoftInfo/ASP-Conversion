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
	'Program Name				:	AdjustPurchaseVoucherForPayRec.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU
	'Created On					:	January 03,2002
	'Modified By				:	Manohar Prabhu.R
	'Modified On				:
	'Modified By				:	UmaMaheswari S
	'Modified On				:	23 June 2011
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
dim sOrgId,sPartyValue,objRs,sQuery,sVouType,sParCode,sParSubType,sParType,sVouDate
Dim sQuery2,sQuery3,sQuery4,objRs2,sTemp,iCrTransNo,iTrNo,oDOM,dAmount,sUnit,sParName
Dim iPiTrNo,iDNTrNo,iCNTrNo,iITrNo,iAdvRecTrNo,iAdvPayTrNo,sAdjTy,iCtr,sSelParTy
Dim sExp,TempNode,Root,sDispClass,sTransTy,sAmdTy,iAdvNo,sParChk,dAdjAmt,sPrintAdjusted

Dim blnSalesInvoice,blnDebitNode,blnPurBills,blnCreditNode,blnAdvancePay,blnAdvanceRec
Dim nVoucherNo,sVoucherDate,nVoucherNoQ,sCallFrom,sRetVal,sDebitNoteQry
Dim nAmtOutstanding,nInvNumber,nBillNo,sCheck,newElement,sCheckData,sDebitAdjNarr,sCrTransNo

sParChk = Request("sType")
sCheck = Request("hType")
Response.Write "<font color=red>"
sCallFrom = Request.Form("hCallType")	'From Payables(PAY) r Receivables(REC)
If sCallFrom = "" Then
	sCallFrom = Request("hCallFrom")
End IF
dAdjAmt = 0.00

IF CStr(sParChk) = "" Then
	sTemp = Split(Request("Val"),"~")
	iCrTransNo = sTemp(0)
	sAmdTy = Request("AmdTy")
	IF CStr(sAmdTy) = "A" Then
		iAdvNo = sTemp(1)
	Else
		iAdvNo = 0
	End IF
Else
	'iAdvNo = 0
	iCrTransNo = 0
	sOrgId = Request.Form("SelOrg")
	sParCode = Request.Form("hdPartyCode")
	sUnit = Request.Form("hUnitName")
	sParName = Request.Form("hParName")
	sParType = Request.Form("SelPartyType")
End IF

IF CStr(iAdvNo) = "" Then
	iAdvNo = 0
	dAdjAmt = "0.00"
End IF


if trim(sVouDate)="" then
	sVouDate=Day(date)&"/"&Month(date)&"/"&Year(date)
end if

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

IF CStr(iCrTransNo) <> 0 Then
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	'oDOM.load server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	
	sRetVal = GetVouchXML(iCrTransNo)
	oDOM.Load server.MapPath(sRetVal)
	oDOM.Save server.MapPath("../temp/transaction/PurVouAdjForPayRec"&Session.SessionID&".xml")

	sQuery = " select V.TransactionNumber,H.PartyType,H.PartySubType,H.PartyCode, "&_
			 " H.OUDefinitionID,Convert(Varchar,H.VoucherDate,103) VoucherDate,H.VoucherAmount, "&_
			 " H.TransactionType,isNull(D.AccUnitPartyType,'NA'),AccUnitPartyType,H.CreatedVoucherNo  "&_
			 " From Acc_T_CreatedVoucherHeader H,Acc_T_CreatedVoucherDetails D,Acc_T_VoucherHeader V Where H.CreatedTransNo = "&_
			 " D.CreatedTransNo And V.CreatedTransNo = H.CreatedTransNo And H.CreatedTransNo = " & iCrTransNo
			 
	'Response.Write "<p>squery"&squery
	
	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		iTrNo = objRs("TransactionNumber")
		sParType = objRs("PartyType")
		sParSubType = objRs("PartySubType")
		sParCode = objRs("PartyCode")
		sOrgId = objRs("OUDefinitionID")
		sVouDate = objRs("VoucherDate")
		dAmount = objRs("VoucherAmount")
		sTransTy = objRs("TransactionType")
		sSelParTy = objRs("PartyType")
		nVoucherNoQ = objRs("CreatedVoucherNo")
	End IF
	objRs.Close
End IF

Set Root = oDOM.documentElement

IF CStr(iCrTransNo) = "0" and CStr(iAdvNo) <> "0" Then
	sQuery = "Select Top 1 B.OUDefinitionID,B.PartyCode,P.PartyName,D.OrgUnitShortDescription,B.VoucherAmt  "&_
			 "From Acc_T_BillsClosingDetails As B Inner Join APP_M_PartyMaster As P  "&_
			 "On B.PartyCode = P.PartyCode Inner Join DCS_OrganizationUnitDefinitions As D "&_
			 "On B.OUDefinitionID = D.OUDefinitionID Where B.AdvCloseNo = "&iAdvNo

	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		sOrgId = objRs("OUDefinitionID")
		sParCode = objRs("PartyCode")
		sUnit = objRs("OrgUnitShortDescription")
		sParName = objRs("PartyName")
		dAdjAmt = FormatNumber(objRs("VoucherAmt"),2,,,0)
	End IF
	objRs.Close
End IF
'Response.Write Root.xml
If sCheck = "A" Then
	oDOM.load server.MapPath("../temp/transaction/Adjust_Det_"&Session.SessionID&".xml")
	Set Root = oDOM.documentElement
End IF
'Response.Write "<p><font color=red>iTrNo="&iTrNo 
'Response.Write "<p><font color=red>iCrTransNo="&iCrTransNo & "-" & sCallFrom & "-"&sTransTy
'Response.Write "<p>sCheck = " & sCheck 

If sCheck = "" Then

	sExp = "//AdvanceDetails"
	set TempNode = Root.selectNodes(sExp)
	If TempNode.length > 0 Then
	

	sDebitNoteQry = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyInvoiceDate,103),"&_
				" P.AmountReceivable, S.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, "&_
				" Convert(Char,H.VoucherDate,103),H.CreatedTransNo ,P.TransactionNumber,H.TransactionType "&_
				" From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C,Acc_T_CreatedRcvbleAdjDet S Where P.OUDefinitionID = '"& sOrgId &"' "&_
				" and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"& sVouDate &"',103) "&_
				" and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.CreatedReceivable "&_
				" and C.ReceivableNumber = S.ReceivableNumber and C.AmountReceivable >= C.AmountReceived and "&_
				"  H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='DNR' and S.CreatedTransNo="& iCrTransNo
	objRs.Open sDebitNoteQry,con
'	Response.Write "<p><font color=red>"&sDebitNoteQry
	Do while Not objrs.EOF 
		set newElement = oDOM.CreateElement("Advance")
		newElement.setAttribute "TransNo",objrs(9)
		newElement.setAttribute "VoucherNo",objrs(1)
		newElement.setAttribute "VoucherDate",objrs(2)
		newElement.setAttribute "AmountRec",cdbl(objrs(3))
		newElement.setAttribute "AmountAdj","0"
		newElement.setAttribute "AmountToAdj",cdbl(objrs(4))'cdbl(objrs(4))-cdbl(objrs(3))
		newElement.setAttribute "CreatedTransNo",objrs(8)
		newElement.setAttribute "AdvNo",objrs(6)
		newElement.setAttribute "AdjType","D"
		newElement.setAttribute "ToAccount","0"
		TempNode.item(0).appendchild newElement
		objRs.MoveNext  
	Loop
	objrs.Close 
	
	sDebitNoteQry = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyInvoiceDate,103),"&_
				" P.AmountReceivable, P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, "&_
				" Convert(Char,H.VoucherDate,103),H.CreatedTransNo ,P.TransactionNumber,H.TransactionType "&_
				" From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where P.OUDefinitionID = '"& sOrgId &"' "&_
				" and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"& sVouDate &"',103) "&_
				" and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.CreatedReceivable "&_
				" and C.AmountReceivable >= C.AmountReceived and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='DNR' and P.AmountReceivable-P.AmountReceived>0 "
	objRs.Open sDebitNoteQry,con
'	Response.Write "<p><font color=red>"&sDebitNoteQry
	Do while Not objrs.EOF 
		set newElement = oDOM.CreateElement("Advance")
		newElement.setAttribute "TransNo",objrs(9)
		newElement.setAttribute "VoucherNo",objrs(1)
		newElement.setAttribute "VoucherDate",objrs(2)
		newElement.setAttribute "AmountRec",cdbl(objrs(3))-cdbl(objrs(4))
		newElement.setAttribute "AmountAdj","0"
		newElement.setAttribute "AmountToAdj","0" 'cdbl(objrs(4))-cdbl(objrs(3))
		newElement.setAttribute "CreatedTransNo",objrs(8)
		newElement.setAttribute "AdvNo",objrs(6)
		newElement.setAttribute "AdjType","D"
		newElement.setAttribute "ToAccount","0"
		TempNode.item(0).appendchild newElement
		objRs.MoveNext  
	Loop
	objrs.Close 
	
	
	
	sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),"&_
	      " isnull(a.AdvanceAdjusted,0),isnull(a.AdvancePaid,0),A.CreatedTransNo,A.CreatedAdvanceNo from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b,"&_
	      "  Acc_T_CreatedAdvances C where  a.OUDefinitionID='"&sOrgId &"' and a.PartyCode="& sParCode &""&_
	      "  and  b.TransactionNumber = a.TransactionNumber and b.VoucherDate<= convert(datetime,'"& sVouDate &"',103) and "&_
	      " A.CreatedAdvanceNo = C.CreatedAdvanceNo  and A.CreatedAdvanceNo = C.CreatedAdvanceNo and isNull(a.AdvancePaid,0)-isNull(a.AdvanceAdjusted,0)>0"
		objRs.Open sQuery,con
'	Response.Write "<p><font color=red>"&sQuery 
	Do while Not objrs.EOF 
	   ' Response.Write "<p>TransNo = "&objRs(0)
		set newElement = oDOM.CreateElement("Advance")
		newElement.setAttribute "TransNo",objrs(0)
		newElement.setAttribute "VoucherNo",objrs(1)
		newElement.setAttribute "VoucherDate",objrs(2)
		newElement.setAttribute "AmountRec",objrs(4)
		newElement.setAttribute "AmountAdj",cdbl(objrs(3))
		newElement.setAttribute "AmountToAdj","0"   'cdbl(objrs(4))-cdbl(objrs(3))
		newElement.setAttribute "CreatedTransNo",objrs(5)
		newElement.setAttribute "AdvNo",objrs(6)
		newElement.setAttribute "AdjType","P"
		newElement.setAttribute "ToAccount","0"
		TempNode.item(0).appendchild newElement
		objRs.MoveNext  
	Loop
	objrs.Close 
	
	
	sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),"&_
	      " isnull(a.AdvanceAdjusted,0),isnull(a.AdvanceReceived,0),A.CreatedTransNo,A.CreatedAdvanceNo from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b,"&_
	      "  Acc_T_CreatedAdvances C where  a.OUDefinitionID='"&sOrgId &"' and a.PartyCode="& sParCode &""&_
	      "  and  b.TransactionNumber = a.TransactionNumber and b.VoucherDate<= convert(datetime,'"& sVouDate &"',103) and "&_
	      " A.CreatedAdvanceNo = C.CreatedAdvanceNo  and A.CreatedAdvanceNo = C.CreatedAdvanceNo and isNull(a.AdvanceReceived,0)-isNull(a.AdvanceAdjusted,0)>0"
		objRs.Open sQuery,con
'	Response.Write "<p><font color=red>"&sQuery 
	Do while Not objrs.EOF 
	   ' Response.Write "<p>TransNo = "&objRs(0)
		set newElement = oDOM.CreateElement("Advance")
		newElement.setAttribute "TransNo",objrs(0)
		newElement.setAttribute "VoucherNo",objrs(1)
		newElement.setAttribute "VoucherDate",objrs(2)
		newElement.setAttribute "AmountRec",objrs(4)
		newElement.setAttribute "AmountAdj",cdbl(objrs(3))
		newElement.setAttribute "AmountToAdj","0"   'cdbl(objrs(4))-cdbl(objrs(3))
		newElement.setAttribute "CreatedTransNo",objrs(5)
		newElement.setAttribute "AdvNo",objrs(6)
		newElement.setAttribute "AdjType","R"
		newElement.setAttribute "ToAccount","0"
		TempNode.item(0).appendchild newElement
		objRs.MoveNext  
	Loop
	objrs.Close 
	
	
		sQuery = "Select P.PayablesNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyBillDate,103), "&_
	         " P.AmountPayable, S.AmountPaid,C.AmountPaid - P.AmountPaid,P.PayablesNumber,"&_
	         " Convert(Char,H.VoucherDate,103),H.CreatedTransNo ,P.TransactionNumber,H.TransactionType "&_
	         " From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C,Acc_T_CreatedPybleAdjDet S Where "&_
	         " P.OUDefinitionID = '"& sOrgId &"' and P.PartyCode = "& sParCode &" and P.VoucherDate <= convert(datetime,'"& sVouDate &"',103)"&_
	         " and P.TransactionNumber = H.TransactionNumber and S.PayablesNumber= C.PayablesNumber  "&_
	         " and (C.PayablesNumber = P.CreatedPayablesNumber or C.PayablesNumber = P.CrCreatedPayable) and C.AmountPayable >= C.AmountPaid "&_
	         "  and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='CNR' and S.CreatedTransNo =  "& iCrTransNo
	objRs.Open sQuery,con
'	Response.Write "<p><font color=red>"&sQuery 
	Do while Not objrs.EOF 
		set newElement = oDOM.CreateElement("Advance")
		newElement.setAttribute "TransNo",objrs(9)
		newElement.setAttribute "VoucherNo",objrs(1)
		newElement.setAttribute "VoucherDate",objrs(2)
		newElement.setAttribute "AmountRec",objrs(3)
		newElement.setAttribute "AmountAdj","0"
		newElement.setAttribute "AmountToAdj",cdbl(objrs(4))'cdbl(objrs(4))-cdbl(objrs(3))
		newElement.setAttribute "CreatedTransNo",objrs(8)
		newElement.setAttribute "AdvNo",objrs(6)
		newElement.setAttribute "AdjType","C"
		newElement.setAttribute "ToAccount","0"
		TempNode.item(0).appendchild newElement
		objRs.MoveNext  
	Loop
	objrs.Close 

	
	
	sQuery = "Select P.PayablesNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyBillDate,103), "&_
	         " P.AmountPayable, P.AmountPaid,C.AmountPaid - P.AmountPaid,P.PayablesNumber,"&_
	         " Convert(Char,H.VoucherDate,103),H.CreatedTransNo ,P.TransactionNumber,H.TransactionType "&_
	         " From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
	         " P.OUDefinitionID = '"& sOrgId &"' and P.PartyCode = "& sParCode &" and P.VoucherDate <= convert(datetime,'"& sVouDate &"',103)"&_
	         " and P.TransactionNumber = H.TransactionNumber and "&_
	         " (C.PayablesNumber = P.CreatedPayablesNumber or C.PayablesNumber = P.CRCreatedPayable) and C.AmountPayable >= C.AmountPaid "&_
	         "  and H.CreatedTransNo = C.CreatedTransNo and P.AmountPayable-P.AmountPaid>0 and H.TransactionType='CNR' "
	objRs.Open sQuery,con
'	Response.Write "<p><font color=red>"&sQuery 
	Do while Not objrs.EOF 
		set newElement = oDOM.CreateElement("Advance")
		newElement.setAttribute "TransNo",objrs(9)
		newElement.setAttribute "VoucherNo",objrs(1)
		newElement.setAttribute "VoucherDate",objrs(2)
		newElement.setAttribute "AmountRec",cdbl(objrs(3))-cdbl(objrs(4))
		newElement.setAttribute "AmountAdj","0"
		newElement.setAttribute "AmountToAdj","0"'cdbl(objrs(4))-cdbl(objrs(3))
		newElement.setAttribute "CreatedTransNo",objrs(8)
		newElement.setAttribute "AdvNo",objrs(6)
		newElement.setAttribute "AdjType","C"
		newElement.setAttribute "ToAccount","0"
		TempNode.item(0).appendchild newElement
		objRs.MoveNext  
	Loop
	objrs.Close 
	
	End IF
End If 'If sCheck = "" Then
'Response.Write "<p>CreatedTransNo = "& iCrTransNo 

oDOM.Save server.MapPath("../temp/transaction/PurVouAdjForPayRecNew"&Session.SessionID&".xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<root No="0">
	<PayRec/>
	<RecCount Val="0" />
</root>
</script>
<%If sCheck = "" Then%>
	<script type="application/xml" data-itms-xml-island="1" id="AdjDet" data-src="<%="../temp/transaction/PurVouAdjForPayRecNew"&Session.SessionID&".xml"%>"></script>
<%Else%>
	<script type="application/xml" data-itms-xml-island="1" id="AdjDet" data-src="<%="../temp/transaction/Adjust_Det_"&Session.SessionID&".xml"%>"></script>
<%End IF%>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AdjustPurchaseVoucherForPayRecCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetTotalVal('<%=iAdvNo%>')">

<form method="POST" name="frm1" action="">
<input type=hidden name="hVouType" value="<%=sVouType%>">
<input type=hidden name="hCrTrNo" value="<%=iCrTransNo%>">
<input type=hidden name="hTrNo" value="<%=iTrNo%>">
<input type=hidden name="hVouAmt" value="<%=dAmount%>">
<input type=hidden name="hTransTy" value="<%=sTransTy%>">
<input type=hidden name="hAdvNo" value="<%=iAdvNo%>">
<input type=hidden name="hParCode" value="<%=sParCode%>">
<input type=hidden name="hUnitID" value="<%=sOrgId%>">
<input type=hidden name="hVouDate" value="<%=sVouDate%>">
<input type=hidden name="hDrCrAmt" value="0">
<input type=hidden name="hType" value="<%=scheck%>">
<input type=hidden name="hVouNumber" value="<%=nVoucherNoQ%>">
<input type=hidden name="hCallFrom" value="<%=sCallFrom%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Adjustments
		</td>
    </tr>
     <%IF CStr(sSelParTy) <> "NA" Then %>
	<tr>
		<td align="center" class="TopPack">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>

							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
								
								<DIV class=frmBody id=frm1 width="100%" >
												<table border="0" cellpadding="0" cellspacing="1" width="100%" class="TableOutlineOnly">
								<%	Dim sBookName,sPartyName,sPurchaseType,sParSubTypeName,sPrintName
									'IF CStr(iCrTransNo) <> 0 Then
										If sCallFrom = "PAY" Then
											sPrintName = "Purchase"
										Else
											sPrintName = "Sales"
										End IF
										
										sExp = "//Header/Book"
										Set TempNode = Root.selectNodes(sExp)
										sBookName = TempNode.Item(0).Text
										'Response.Write "<p>nVoucherNoQ="&nVoucherNoQ
										'If nVoucherNo = "" Then nVoucherNo = cdbl(nVoucherNoQ)
										If nVoucherNo = "" Then nVoucherNo = nVoucherNoQ
										%>
											<tr>
											  <td width="25%" class="FieldCellSub"><%=sPrintName%> Book</td>
											  <td width="45%" class="FieldCellSub"><span class="DataOnly"><%=sBookName%> </span></td>
											  <td width="25%" class="FieldCellSub">Invoice Number</td>
											  <td width="10%" class="FieldCellSub">
												<span class="DataOnly"><%=nVoucherNo%></span> &nbsp;&nbsp;
												<img src="../../assets/images/iTMS%20icons/DetailsIcon.gif" onClick="ShowVouch(<%=iCrTransNo%>)" alt="View Voucher Details" style="cursor:pointer" >
											  </td>
											</tr>
											<%
												sExp = "//Header/Party"
												Set TempNode = Root.selectNodes(sExp)
												sPartyName = TempNode.Item(0).Text
												sParSubTypeName = TempNode.item(0).Attributes.getNamedItem("ParSubTypeName").value
												
												If sCallFrom = "PAY" Then
													sExp = "//Header/PurchaseType"
												Else
													sExp = "//Header/SalesType"
												End IF
												Set TempNode = Root.selectNodes(sExp)
												sPurchaseType = TempNode.Item(0).Text
												
												If sCallFrom = "PAY" Then
													sExp = "//Header/PurInvoice"
													Set TempNode = Root.selectNodes(sExp)
													sVoucherDate = TempNode.Item(0).Attributes.getNamedItem("PurInvDate").Value
												Else
													sExp = "//Header/SaleInvoice"
													Set TempNode = Root.selectNodes(sExp)
													sVoucherDate = TempNode.Item(0).Attributes.getNamedItem("InvDate").Value
												End IF
												
												
											%>
											<tr>
										    <td class="FieldCellSub"><%=sPrintName%> Type</td>
										    <td width="45%" class="FieldCellSub"><span class="DataOnly"><%=sPurchaseType%> </span></td>
										    <td width="15%" class="FieldCellSub">Invoice Date</td>
											<td width="15%" class="FieldCellSub"><span class="DataOnly"><%=sVoucherDate%> </span></td>
										  </tr>

										  <tr>
										    <td class="FieldCellSub">Party Type</td>
										    <td width="45%" class="FieldCellSub"><span class="DataOnly"><%=sParType%>-<%=sParSubTypeName%> </span></td>
										    <td width="15%" class="FieldCellSub">Amount</td>
											<td width="15%" class="FieldCellSub"><span class="DataOnly" id="spVouAmt"><%=FormatNumber(dAmount,2,,,0)%> </span></td>
										  </tr>
										  <tr>
										    <td class="FieldCellSub">Party Name</td>
										    <td width="45%" class="FieldCellSub"><span class="DataOnly"><%=sPartyName%> </span></td>
										  </tr>

										  <tr>
										    <td width="100%" class="FieldCellSub" colspan="2">Voucher Adjustment Details</td>
										  </tr>

										  <tr>
										    <td width="100%" class="FieldCellSub" colspan="4">
										      <table border="0" cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
										        <tr>
										          <td width="2%" class="ExcelHeaderCell" align="center">S.No</td>
										          <td class="ExcelHeaderCell" align="center">
													<img border="0" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Remove Bill Details" onclick="DeleteData()"></a>	
												  </td>
										          <td width="55%" class="ExcelHeaderCell" align="center">Voucher Details</td>
										          <td width="19%" class="ExcelHeaderCell" align="center">Voucher Amount</td>
										          <td width="16%" class="ExcelHeaderCell" align="center">Adjusted Amount</td>
										        </tr>
										        <%	Dim nAmtToAdjust,sCheckDelAtt,nTotAmtAdjusted,sPrintMsg,nTransactionNo,sTransType,sNarration
										            Dim iSLNO
													nTotAmtAdjusted = cdbl("0")
													sPrintMsg = ""
													
													sExp = "//AdvanceDetails/Advance" 
													set TempNode = Root.selectNodes(sExp)
													For iCtr = 0 To TempNode.length - 1
													
														sCheckDelAtt = ""
														If sCheck = "A" Then
															sCheckDelAtt = TempNode.Item(iCtr).Attributes.getNamedItem("Check").Value
														End If
														'If nVoucherNo = "" Then nVoucherNo = cdbl(nVoucherNoQ)
														
														sAdjTy = TempNode.Item(iCtr).Attributes.getNamedItem("AdjType").Value
														nAmtToAdjust = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("AmountToAdj").Value,2,,,0)
														nVoucherNo  = TempNode.Item(iCtr).Attributes.getNamedItem("VoucherNo").Value
														sVoucherDate = TempNode.Item(iCtr).Attributes.getNamedItem("VoucherDate").Value
														nTransactionNo = TempNode.Item(iCtr).Attributes.getNamedItem("TransNo").Value
														sCrTransNo = TempNode.Item(iCtr).Attributes.getNamedItem("CreatedTransNo").Value
														
														If sCheckDelAtt <> "DEL" and cdbl(nAmtToAdjust) > "0" Then
														
															sPrintMsg = "Y"
															
														'	sQuery= "select H.TransactionType From Acc_T_CreatedVoucherHeader  H,Acc_T_AdvancePayments P where H.CreatedTransNo = P.CreatedTransNo and P.TransactionNumber = "& nTransactionNo &" "
														
														    sQuery = "Select TransactionType from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& sCrTransNo
														   ' Response.Write "<tr><td>"& sQuery &"</td></tr>"
															objRs2.Open sQuery,con
															If Not objRs2.EOF Then
																sTransType = Trim(objRs2(0))
															End IF
															objRs2.Close 
														'	Response.Write "<tr><td>"& sTransType &"</td></tr>"
															If sTransType = "BAP" or sTransType = "BAR" Then
																sNarration = "BANK"
															Elseif sTransType = "CAP" or sTransType = "CAR" Then
																sNarration = "CASH"
														    Elseif sTransType = "DNR" then
														        sNarration = "Debit Note "
														    Elseif sTransType = "CNR" then
														        sNarration = "Credit Note "
															End IF
															iSLNO = iSLNO + 1
													%>
														<tr>
															<td class="ExcelHeaderCell" align="center"><%Response.Write(iSLNO)%></td>
															<td class="ExcelHeaderCell" align="center">
																<Input type="Checkbox" Name="ChkBillDet"  value="<%=nVoucherNo&":"&nAmtToAdjust&":"&nTransactionNo%>" >
															</td>
															<td class="ExcelDisplayCell" align="left"><%Response.Write sNarration & " VOU No:"%><%=nVoucherNo%> DT:<%=sVoucherDate%></td>
															<td class="ExcelDisplayCell" align="right"><%Response.Write(FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("AmountRec").Value,2,,,0))%></td>
															<td class="ExcelDisplayCell" align="right"><%Response.Write(FormatNumber(nAmtToAdjust,2,,,0))%></td>
														</tr>
													<%
															nTotAmtAdjusted = cdbl(nTotAmtAdjusted) + cdbl(TempNode.Item(iCtr).Attributes.getNamedItem("AmountToAdj").Value)
														End IF 'If nAmtToAdjust <> "0" Then
														'End IF 'If sCheckDelAtt <> "DEL" Then 
													Next
													%>

										<%'Response.Write "<p><font color=red>nTotAmtAdjusted="&sCheckDelAtt
										If trim(sPrintMsg) = "" Then iCtr = 0
										'Response.Write "<p><font color=red>iCtr="&iCtr
										IF CStr(iCtr) = "0" Then
										%>
											<tr>
												<td class="ExcelDisplayCell" align="center" colspan="5"><b>No Bills is Selected for Adjustments!! Adjustment with New Bills is Not Possible.</b></td>
											</tr>

										<%End IF %>
											<tr>
												<td class="ExcelDisplayCell" colspan="4" align="Right"> <b>Total Adjusted Amount</b></td>
												<td class="ExcelDisplayCell" align="right">
													<Input type="Text" Name="TxtTotAdjustedAmount" style="text-align:right" Value="<%=FormatNumber(nTotAmtAdjusted,2,,,0)%>" Readonly Class="FormElemRead">
												</td>
											</tr>
										
										
										      </table>
										    </td>
										  </tr>
										  
										</table>
									<table border="0" cellspacing="0" width="100%">
									 <tr>
										<td align="center" class="MiddlePack" colspan="3">
										</td>
									</tr>
								<%'Else%>
										
										<tr>
										  <td width="25%" class="FieldCellSub">Debit Amount</td>
										  <td Width="25%" class="FieldCellSub"><span class="DataOnly" id="spDRAmt"><%=dAdjAmt%> </span></td>
										  <td width="25%" class="FieldCellSub">Credit Amount</td>
										  <td Width="25%" class="FieldCellSub"><span class="DataOnly" id="spCRAmt"><%=dAdjAmt%> </span></td>
										</tr>
										<tr>
										  <td width="25%" class="FieldCellSub">Differences</td>
										  <td Width="25%" class="FieldCellSub"><span class="DataOnly" id="spDiffAmt">0.00 </span></td>
										  <td width="25%" class="FieldCellSub">&nbsp;</td>
										  <td Width="25%" class="FieldCellSub">&nbsp;</td>
										</tr>
								<%'End IF %>
									</table>
									</Div>
									<DIV class=frmBody id=frm1 width="100%" style="height:200;">
								<%' IF CStr(iCtr) <> "0" Then %>
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                               <tr>
													<td class="ExcelHeaderCell" align="Left" colspan="81">Bills Type</td>
											   </tr>
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"  rowspan="2"></td>
                                        <td class="ExcelHeaderCell" align="center" colspan="3">Document</td>
                                        <td class="ExcelHeaderCell" align="center" colspan="4">Amount </td>

                                            </tr>
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center">Number</td>
                                        <td class="ExcelHeaderCell" align="center">Date</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                        <td class="ExcelHeaderCell" align="center"> Adjusted</td>
                                        <td class="ExcelHeaderCell" align="center">To Account</td>
                                        <td class="ExcelHeaderCell" align="center">Outstanding</td>
                                        <td class="ExcelHeaderCell" align="center">To Adjust</td>

                                            </tr>
 

<%

'		P  -- Advance Payments
'		R  -- Advance Receipts
'*********************** Adjustment Type Details ******************************************

Dim iDocNo,iSno,iPayNo,iRecCount,sCrDrNoDesc
Dim sAmtReceived,sAmtReceiavble,sInvDate,sInvNo,sVouNo,dToAccount,dToAdjusted
Dim sCrDrNo,sCrDrType,iPlace
Response.Write "<font color=red>"
'Response.Write "<p>PartyType = "&sParType 
iSno = 1
IF CStr(iCrTransNo) <> "0" Then
'************************ Advance Payments/Receipts ****************************************
	IF CStr(sParType) = "CR" Then
		sQuery = " select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(c.AdvanceAdjusted,0), "&_
			 " isnull(a.AdvancePaid,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 " b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 " A.CreatedAdvanceNo,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 " from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 " a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 " b.TransactionNumber = a.TransactionNumber "&_
			 " and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 " A.CreatedAdvanceNo = C.CreatedAdvanceNo "&_
			 " and A.CreatedAdvanceNo = C.CreatedAdvanceNo "
	Else
		 sQuery = " select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(C.AdvanceAdjusted,0), "&_
			 " isnull(a.AdvanceReceived,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 " b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 " A.CreatedAdvanceNo,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 " from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 " a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 " b.TransactionNumber = a.TransactionNumber "&_
			 " and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 " A.CreatedAdvanceNo = C.CreatedAdvanceNo "&_
			 " and A.CreatedAdvanceNo = C.CreatedAdvanceNo " 
	End IF

'	Response.Write "<textarea>Query="&sQuery &"<BR></Textarea>"

	IF CStr(sParType) = "CR" Then
		blnAdvancePay = true
		sCrDrType = "P"
		sCrDrNoDesc = "Advance Payments  "
		%>
		<tr>
			<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Payments</b></td>
		</tr>
	<%Else
		blnAdvanceRec = true
		sCrDrType = "R"
		sCrDrNoDesc = "Advance Receipts  "
	%>
		<tr>
			<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Receipts</b></td>
		</tr>
	<%
	End IF

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	'iSno=1
	Do While Not objRs.EOF
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		sAmtReceived = objRs(3)
		sAmtReceiavble = objRs(4)
		'iCrDocNo = objRs(5)
		iPayNo = objRs(12)
		dToAccount = Objrs(13)
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		
		'Response.Write "<p>ToAdjusted = "& dToAdjusted 
		
		IF CStr(iCrTransNo) <> 0 Then
			sDispClass = "ExcelDisplayCell"
			sExp = "//AdvanceDetails/Advance[@TransNo="&iDocNo&"]"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.length <> 0 Then
				dToAdjusted = CDbl(dToAdjusted) + CDbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value)
			End IF
		End IF
		
		IF CDbl(dToAdjusted) <> 0 Then
			blnAdvancePay = false
			blnAdvanceRec = false
			sPrintAdjusted = ""
			IF CStr(iCrTransNo) <> 0 Then
				IF TempNode.length <> 0 Then
					'Response.Write "<p><font color=red>val="&CDbl(TempNode.item(0).Attributes.getNamedItem("VoucherNo").Value)
					If CDbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value) <> "0" Then
						sPrintAdjusted = "Y"	'Already Adjusted
					Else
						sPrintAdjusted = "N"	'Not Adjusted
					End IF
				End If
			End If
		'	Response.Write "<p><font color=red>sPrintAdjusted="&sPrintAdjusted
			If sPrintAdjusted <> "" Then
			%>	
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>?<%=trim(objRs(11))%>">
				
				<%If sPrintAdjusted = "N" Then%>
					<td class="ExcelSerial" align="center"><%=iSno%></td>
					<td class="ExcelInputCell" align="right" width="10">
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
					<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
					<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right">
						<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
					</td>
					
				<%Elseif sPrintAdjusted = "Y" Then
					nAmtOutstanding = 0
					sAmtReceived = objRs(3)
					nAmtOutstanding = CDbl(sAmtReceiavble) - (CDbl(sAmtReceived)+ CDbl(dToAccount))
					%>
					<td class="ExcelSerial" align="center"><%=iSno%><FONT color="red"><b>*</b></font></td>
					<td class="ExcelInputCell" align="right" width="10">
						<%If sCheck = "A" Then
							If TempNode.item(0).Attributes.getNamedItem("Check").Value = "DEL" Then
								If cdbl(sAmtReceived) > 0  Then
									sAmtReceived = cdbl(sAmtReceived) - cdbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value)
								End IF
								nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))%>
								<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
							<%Else%>
								<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
							<%End IF%>
						<%Else%>
							<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
						<%End IF%>
					<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
					<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right">
						<Input type="Text" Name="TxtAdjust<%=sInvNo%>" Value="<%=FormatNumber(sAmtReceived,2,,,0)%>" style="text-align:right" size="10" class="FormElem" Readonly>
					</td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right">
						<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(nAmtOutstanding,2,,,0)%><%'=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
					</td>
					<%
				End If
				%>
				
				<td class="ExcelDisplayCell" align="right">
					<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly onchange="DisplayAmt('<%=iDocNo&"Z"&iPayNo&"Z"&iSno%>')">
					<!--<input type=text value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>">-->
					<Input type="hidden" name="hCheckType<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrType%>">
				</td>
			</tr>
<%			iSno=cint(iSno)+1
			End IF
		End IF
	objRs.MoveNext
	Loop
	objRs.Close

'	Response.Write "sCrDrType = "&sCrDrType
'	Response.Write  blnAdvancePay
'	Response.Write blnAdvanceRec
	if trim(sCrDrType)="P" and blnAdvancePay then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Payments Found !!</b></td>
		</tr>
	<%
	elseif trim(sCrDrType)="R" and blnAdvanceRec then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Receipts Found !!</b></td>
		</tr>
	<%
	Else%>
	<tr>
		<td colspan="8" class="ExcelDisplayCell" align="Right"><b>Total Amount To Be Adjusted</b></td>
		<td class="ExcelDisplayCell" align="right">
			<Input type="Text" Name="TxtTotToAdjustedAmount" style="text-align:right" Value="0.00" Readonly Class="FormElemRead">
		</td>
	</tr>
	<%end if

	IF CStr(iCrTransNo) <> "0" and  1= 2 Then
	%>
		<tr>
			<td class="ExcelDisplayCell" colspan="9"><b><FONT color="red">*</Font> Denotes Bill(s) Adjusted to this Voucher </b></td>
		</tr>
	<%
	End IF
'End IF
Else
	'Response.Write "==========="
	sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
			 "isnull(a.AdvancePaid,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 "a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 "isnull(a.AdvancePaid,0)>isnull(a.AdvanceAdjusted,0) and "&_
			 "b.TransactionNumber = a.TransactionNumber "&_
			 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 "A.CreatedAdvanceNo = C.CreatedAdvanceNo  and A.AdvanceNumber NOT IN("&iAdvPayTrNo&")  Order By b.VoucherDate"

	sQuery2 = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
			 "isnull(a.AdvanceReceived,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 "a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 "isnull(a.AdvanceReceived,0)>isnull(a.AdvanceAdjusted,0) and "&_
			 "b.TransactionNumber = a.TransactionNumber "&_
			 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 "A.CreatedAdvanceNo = C.CreatedAdvanceNo and A.AdvanceNumber NOT IN("&iAdvRecTrNo&")  Order By b.VoucherDate "
%>
	<tr>
		<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Payments</b></td>
	</tr>
<%
	sCrDrType = "P"
	sCrDrNoDesc = "Advance Payments  "
	blnAdvancePay = true
'	Response.Write "<p>"& sQuery 
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	'iSno=1
	Do While Not objRs.EOF
		blnAdvancePay = false
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		sAmtReceived = objRs(3)
		sAmtReceiavble = objRs(4)
		'iCrDocNo = objRs(5)
		iPayNo = objRs(12)
		dToAccount = Objrs(13)
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		IF CDbl(dToAdjusted) <> 0 Then

%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>?<%=trim(objRs(11))%>">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
				<td class="ExcelDisplayCell" align="right"></td>
			</tr>

<%			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
		Loop
		objRs.Close

		IF CStr(iAdvNo) <> "0" Then
			sQuery = "Select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103) "&_
					 "VoucherDate,D.BillAdjAmt, D.BillAmt,B.TransactionType,A.CreatedTransNo, "&_
					 "A.AdvanceNumber From Acc_T_AdvancePayments A, Acc_T_VoucherHeader b, "&_
					 "Acc_T_CreatedAdvances C,Acc_T_BillsClosingDetails D where a.OUDefinitionID='"&sOrgId&"' and "&_
					 "A.PartyCode="&sParCode&" and B.TransactionNumber = A.TransactionNumber and  "&_
					 "A.CreatedAdvanceNo = C.CreatedAdvanceNo And D.BillTrNo = C.CreatedAdvanceNo And "&_
					 "D.AdvCloseNo = "&iAdvNo&" And D.BillType = 'P' Order By b.VoucherDate "

		'	Response.Write sQuery
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection = nothing
			'iSno=1
			Do While Not objRs.EOF
				blnAdvance = false
			iDocNo = objRs("TransactionNumber")
			sInvNo = objRs("VoucherNumber")
			sInvDate = objRs("VoucherDate")
			sAmtReceived = objRs("BillAdjAmt")
			sAmtReceiavble = objRs("BillAmt")
			'iCrDocNo = objRs(5)
			iPayNo = objRs("AdvanceNumber")
			dToAccount = 0
			sAmtReceiavble = CDbl(sAmtReceiavble)
			sAmtReceived = CDbl(sAmtReceived)
			dToAccount= CDbl(dToAccount)
			dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
			dToAdjusted = sAmtReceived
			IF CDbl(dToAdjusted) <> 0 Then

%>
				<tr>
					<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>?<%=objrs(7)%>">
					<td class="ExcelSerial" align="center"><%=iSno%></td>
					<td class="ExcelInputCell" align="right" width="10">
					<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
					<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
					<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right">
						<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
					</td>
					<td class="ExcelDisplayCell" align="right"></td>
				</tr>

<%				iSno=cint(iSno)+1
			End IF
			objRs.MoveNext
		Loop
		objRs.Close

	End IF

	if blnAdvancePay then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Payments Found !!</b></td>
		</tr>
	<%
	end if

%>
	<tr>
		<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Receipts</b></td>
	</tr>
<%
	sCrDrType = "R"
	sCrDrNoDesc = "Advance Receipts  "
	blnAdvanceRec = true
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery2
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	'iSno=1
	Do While Not objRs.EOF
		blnAdvanceRec =false
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		sAmtReceived = objRs(3)
		sAmtReceiavble = objRs(4)
		'iCrDocNo = objRs(5)
		iPayNo = objRs(12)
		dToAccount = Objrs(13)
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		IF CDbl(dToAdjusted) <> 0 Then

%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>?<%=objrs(11)%>">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" ></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
				<td class="ExcelDisplayCell" align="right"></td>
			</tr>

<%			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
		Loop
		objRs.Close

		IF CStr(iAdvNo) <> "0" Then
			sQuery = "Select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103) "&_
					 "VoucherDate,D.BillAdjAmt, D.BillAmt,B.TransactionType,A.CreatedTransNo, "&_
					 "A.AdvanceNumber From Acc_T_AdvancePayments A, Acc_T_VoucherHeader b, "&_
					 "Acc_T_CreatedAdvances C,Acc_T_BillsClosingDetails D where a.OUDefinitionID='"&sOrgId&"' and "&_
					 "A.PartyCode="&sParCode&" and B.TransactionNumber = A.TransactionNumber and  "&_
					 "A.CreatedAdvanceNo = C.CreatedAdvanceNo And D.BillTrNo = C.CreatedAdvanceNo And "&_
					 "D.AdvCloseNo = "&iAdvNo&" And D.BillType = 'R' Order By b.VoucherDate "


			'Response.Write sQuery
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			Set objRs.ActiveConnection = nothing
			Do While Not objRs.EOF
				blnAdvanceRec = false
				iDocNo = objRs("TransactionNumber")
				sInvNo = objRs("VoucherNumber")
				sInvDate = objRs("VoucherDate")
				sAmtReceived = objRs("BillAdjAmt")
				sAmtReceiavble = objRs("BillAmt")
				'iCrDocNo = objRs(5)
				iPayNo = objRs("AdvanceNumber")
				dToAccount = 0
				sAmtReceiavble = CDbl(sAmtReceiavble)
				sAmtReceived = CDbl(sAmtReceived)
				dToAccount= CDbl(dToAccount)
				'dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
				dToAdjusted = sAmtReceived
				IF CDbl(dToAdjusted) <> 0 Then

%>
					<tr>
						<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>?<%=objrs(7)%>">
						<td class="ExcelSerial" align="center"><%=iSno%></td>
						<td class="ExcelInputCell" align="right" width="10">
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
						<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
						<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right">
							<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
						</td>
						<td class="ExcelDisplayCell" align="right"></td>
					</tr>

<%				iSno=cint(iSno)+1
			End IF
			objRs.MoveNext
		Loop
		objRs.Close
		End IF

	if blnAdvancePay then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Receipts Found !!</b></td>
		</tr>
	<%
	end if

	End IF
End IF

If Trim(sCallFrom) = "PAY" Then
%>
 <tr>
	<td colspan="9" class="ExcelSerial"><b>Debit Notes</b></td>
<%
Dim sCrDrVouDate,sCheckDebQry
sCrDrType = "D"
blnDebitNode = true
sCrDrNoDesc = "DEBIT NOTE NO: "
iCNTrNo = 0	'Test
iSno = 1

sQuery2 = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyInvoiceDate,103),P.AmountReceivable, "&_
		  "P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, Convert(Char,H.VoucherDate,103),H.CreatedTransNo "&_
		  "From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where "&_
		  " P.OUDefinitionID = '"&sOrgId&"'  "&_
		  "and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "&_
		  "and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.CreatedReceivable  "&_
		  "and C.AmountReceivable > C.AmountReceived and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='DNR' and C.AmountReceived > 0"
		  'P.AmountReceivable > P.AmountReceived and

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery2
	.ActiveConnection = con
	.Open
end with

iRecCount = objRs.RecordCount + CDbl(iRecCount)
set objRs.ActiveConnection = nothing

sQuery2 = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyInvoiceDate,103),"&_
				" P.AmountReceivable, P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, "&_
				" Convert(Char,H.VoucherDate,103),H.CreatedTransNo ,P.TransactionNumber,H.TransactionType "&_
				" From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where P.OUDefinitionID = '"& sOrgId &"' "&_
				" and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"& sVouDate &"',103) "&_
				" and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.CreatedReceivable "&_
				" and C.AmountReceivable >= C.AmountReceived and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='DNR' "
objRs2.Open sQuery2,con
If Not objRs2.EOF Then
	iRecCount = "0"
End IF
objRs2.Close 

If iRecCount = "0" Then
	objrs.Close 
	sQuery2 = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyInvoiceDate,103),"&_
				" P.AmountReceivable, P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, "&_
				" Convert(Char,H.VoucherDate,103),H.CreatedTransNo ,P.TransactionNumber,H.TransactionType "&_
				" From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where P.OUDefinitionID = '"& sOrgId &"' "&_
				" and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"& sVouDate &"',103) "&_
				" and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.CreatedReceivable "&_
				" and C.AmountReceivable >= C.AmountReceived and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='DNR' "
	objRs.Open sQuery2,con
	sCheckDebQry = "Y"
Else
	sCheckDebQry = "N"
End IF
'=====================================================================================
'Response.Write "<p><font color=red>"&sQuery2
'Response.Write "<p><font color=red>sCheckDebQry="&sCheckDebQry
	Do While Not objRs.EOF
		
		blnDebitNode = false
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		iPayNo = objRs(0)
		sCrDrNo = objRs(6)
		sCrDrVouDate = objRs(7)

		sAmtReceiavble = Trim(objRs(3))
		sAmtReceived = Trim(objRs(4))
		dToAccount= Trim(objRs(5))

		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)

		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		
		IF CStr(iCrTransNo) <> 0 and sCheckDebQry = "Y" Then
			sDispClass = "ExcelDisplayCell"
			sExp = "//AdvanceDetails/Advance[@TransNo="&objRs(9)&"]"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.length <> 0 Then
				dToAdjusted = CDbl(dToAdjusted) + CDbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value)
			End IF
		End IF
		
		'Response.Write "<p><font color=red>sCheckDebQry="&sCheckDebQry

		IF CDbl(dToAdjusted) > 0 Then
			If sCheckDebQry = "N" Then
				sPrintAdjusted = "N"	'Not Adjusted
			Else
				sPrintAdjusted = "Y"	'Adjusted
			End IF
			'Response.Write "<p>sPrintAdjusted="&sPrintAdjusted
			IF CStr(iCrTransNo) <> 0 and sCheckDebQry = "Y" Then
				IF TempNode.length <> 0 Then
					If CDbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value) <> "0" Then
						sPrintAdjusted = "Y"	'Already Adjusted
					Else
						sPrintAdjusted = "N"	'Not Adjusted
					End IF
				End If
			End If
			'Response.Write "<p>sPrintAdjusted="&sPrintAdjusted

%>
		<tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?D?<%=objrs(8)%>">
			<%If sPrintAdjusted = "N" Then%>
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
			<%Elseif sPrintAdjusted = "Y" Then
				nAmtOutstanding = 0
				sAmtReceived = objRs(4)
				nAmtOutstanding = CDbl(sAmtReceiavble) - (CDbl(sAmtReceived)+ CDbl(dToAccount))
				%>
				<td class="ExcelSerial" align="center"><%=iSno%><FONT color="red"><b>*</b></font></td>
				<td class="ExcelInputCell" align="right" width="10">
					<%If sCheck = "A" Then
						If TempNode.item(0).Attributes.getNamedItem("Check").Value = "DEL" Then
							If cdbl(sAmtReceived) > 0  Then
								sAmtReceived = cdbl(sAmtReceived) - cdbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value)
							End IF
							nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))%>
							<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
						<%Else%>
							<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
						<%End IF%>
					<%Else%>
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
					<%End IF%>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<Input type="Text" Name="TxtAdjust<%=sInvNo%>" Value="<%=FormatNumber(sAmtReceived,2,,,0)%>" style="text-align:right" size="10" class="FormElem" Readonly>
				</td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(nAmtOutstanding,2,,,0)%><%'=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
			<%End IF%>
			<td class="ExcelDisplayCell" align="right">
				<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly onchange="DisplayAmt('<%=iDocNo&"Z"&iPayNo&"Z"&iSno%>')">
			</td>
			<Input type="hidden" name="hCheckType<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="D">
       </tr>

<%
		iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop

objRs.Close

IF CStr(iAdvNo) <> "0" Then
	sQuery = "Select Distinct B.PayRecNo,B.BillAmt,B.BillAdjAmt,R.CreatedReceivable,R.Narration, "&_
			 "Convert(Varchar(10),R.PartyInvoiceDate,103) PartyInvoiceDate,H.CreatedTransNo From "&_
			 "Acc_T_BillsClosingDetails B,Acc_T_Receivables R, " &_
			 "Acc_T_VoucherHeader H Where R.ReceivableNumber = B.PayRecNo And  "&_
			 "B.AdvCloseNo = "&iAdvNo&" And H.TransactionNumber = B.BillTrNo  "&_
			 "And H.BookCode = '06' "

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	iRecCount = objRs.RecordCount
	Set objRs.ActiveConnection = nothing
	Do While Not objRs.EOF
	blnDebitNode = false
		sInvNo = objRs("Narration")
		sInvDate = objRs("PartyInvoiceDate")
		iPayNo= objRs("PayRecNo")
		iDocNo = objRs("CreatedReceivable")
		sAmtReceiavble = Trim(objRs("BillAmt"))
		sAmtReceived = CDbl(objRs("BillAmt")) - CDbl(objRs("BillAdjAmt"))
		dToAccount= 0
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
%>
		<tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?D?<%=objrs(6)%>">
            <td class="ExcelSerial" align="center"><%=iSno%></td>
            <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
			<td class="ExcelDisplayCell" align="right">
				<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly onchange="DisplayAmt('<%=iDocNo&"Z"&iPayNo&"Z"&iSno%>')">
			</td>
			<Input type="hidden" name="hCheckType<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="D">
       </tr>

<%
		iSno=cint(iSno)+1
		objRs.MoveNext
	Loop
	objRs.Close
End IF

	if blnDebitNode then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Debit Notes Found !!</b></td>
		</tr>
	<%
	end if

Else   'if Trim(sCallFrom) = "PAY" Then
	sCrDrType = "C"
	sCrDrNoDesc = "Credit Notes"
	iCNTrNo = "0"
	%>
		<tr><td colspan="9" class="ExcelSerial"><b>Credit Notes </b></td></tr>
	<%
	blnCreditNode = True
	
	sQuery4 = "Select P.PayablesNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyBillDate,103), "&_
		  "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ,P.PayablesNumber, Convert(Char,H.VoucherDate,103),H.CreatedTransNo "&_
		  "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
		  "P.Narration is Null and P.AmountPayable > P.AmountPaid "&_
		  "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and "&_
		  "P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) and P.TransactionNumber = "&_
		  "H.TransactionNumber and (C.PayablesNumber = P.CrCreatedPayable or C.PayablesNumber = P.CreatedPayablesNumber) "&_
		  " and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='CNR' and C.AmountPaid >0"
		  
		'Response.Write sQuery4

		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery4
			.ActiveConnection = con
			.Open
		end with

		iRecCount = objRs.RecordCount + CDbl(iRecCount)
		set objRs.ActiveConnection = nothing
		
		iRecCount = objRs.RecordCount + CDbl(iRecCount)
set objRs.ActiveConnection = nothing
sQuery2 = "Select P.PayablesNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyBillDate,103), "&_
		  "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ,P.PayablesNumber, Convert(Char,H.VoucherDate,103),H.CreatedTransNo,"&_
		  "P.TransactionNumber,H.TransactionType From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
          "P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and  P.AmountPayable > P.AmountPaid  and "&_
		  "P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) and P.TransactionNumber = "&_
		  "H.TransactionNumber and (C.PayablesNumber = P.CrCreatedPayable or C.PayablesNumber = P.CreatedPayablesNumber) "&_
		  "And C.AmountPayable > C.AmountPaid and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='CNR'"
		  
'	Response.Write "<p>"& sQuery2 
objRs2.Open sQuery2,con
If Not objRs2.EOF Then
	iRecCount = "0"
End IF
objRs2.Close 

If iRecCount = "0" Then
	objrs.Close 
    sQuery2 = "Select P.PayablesNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyBillDate,103), "&_
		      "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ,P.PayablesNumber, Convert(Char,H.VoucherDate,103),H.CreatedTransNo,"&_
		      "P.TransactionNumber,H.TransactionType From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
              "P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and "&_
		      "P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) and P.TransactionNumber = "&_
		      "H.TransactionNumber and (C.PayablesNumber = P.CrCreatedPayable or C.PayablesNumber = P.CreatedPayablesNumber) "&_
		      " and H.CreatedTransNo = C.CreatedTransNo and H.TransactionType='CNR'"
'Response.Write sQuery2 
	objRs.Open sQuery2,con
	sCheckDebQry = "Y"
Else
	sCheckDebQry = "N"
End IF
		'=====================================================================================
		Do While Not objRs.EOF
			blnCreditNode = false
			iDocNo = objRs(0)
			sInvNo = objRs(1)
			sInvDate = objRs(2)
			iPayNo = objRs(0)
			sCrDrNo = objRs(6)
			sCrDrVouDate = objRs(7)

			sAmtReceiavble = Trim(objRs(3))
			sAmtReceived = Trim(objRs(4))
			dToAccount= Trim(objRs(5))

			sAmtReceiavble = CDbl(sAmtReceiavble)
			sAmtReceived = CDbl(sAmtReceived)
			dToAccount= CDbl(dToAccount)
			dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
			
			
			IF CStr(iCrTransNo) <> 0 and sCheckDebQry = "Y" Then
			    sDispClass = "ExcelDisplayCell"
			    sExp = "//AdvanceDetails/Advance[@TransNo="&objRs(9)&"]"
			    Set TempNode = Root.selectNodes(sExp)
			    IF TempNode.length <> 0 Then
				    dToAdjusted = CDbl(dToAdjusted) + CDbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value)
			    End IF
		    End IF
    		
			IF CDbl(dToAdjusted) > 0 Then
	   			If sCheckDebQry = "N" Then
				    sPrintAdjusted = "N"	'Not Adjusted
			    Else
				    sPrintAdjusted = "Y"	'Adjusted
			    End IF
			    'Response.Write "<p>sPrintAdjusted="&sPrintAdjusted
			    IF CStr(iCrTransNo) <> 0 and sCheckDebQry = "Y" Then
				    IF TempNode.length <> 0 Then
					    If CDbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value) <> "0" Then
						    sPrintAdjusted = "Y"	'Already Adjusted
					    Else
						    sPrintAdjusted = "N"	'Not Adjusted
					    End IF
				    End If
			    End If
'Response.Write "<p>sPrintAdjusted  = "& sPrintAdjusted 
		%>
		 <tr>
		 	<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?C?<%=objrs(8)%>">
		 	
		 	<%If sPrintAdjusted = "N" Then%>
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
			<%Elseif sPrintAdjusted = "Y" Then
				nAmtOutstanding = 0
				sAmtReceived = objRs(4)
				nAmtOutstanding = CDbl(sAmtReceiavble) - (CDbl(sAmtReceived)+ CDbl(dToAccount))
				%>
				<td class="ExcelSerial" align="center"><%=iSno%><FONT color="red"><b>*</b></font></td>
				<td class="ExcelInputCell" align="right" width="10">
					<%If sCheck = "A" Then
						If TempNode.item(0).Attributes.getNamedItem("Check").Value = "DEL" Then
							If cdbl(sAmtReceived) > 0  Then
								sAmtReceived = cdbl(sAmtReceived) - cdbl(TempNode.item(0).Attributes.getNamedItem("AmountToAdj").Value)
							End IF
							nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))%>
							<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
						<%Else%>
							<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
						<%End IF%>
					<%Else%>
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
					<%End IF%>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<Input type="Text" Name="TxtAdjust<%=sInvNo%>" Value="<%=FormatNumber(sAmtReceived,2,,,0)%>" style="text-align:right" size="10" class="FormElem" Readonly>
				</td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(nAmtOutstanding,2,,,0)%><%'=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
			<%End IF%>
			<td class="ExcelDisplayCell" align="right">
				<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly onchange="DisplayAmt('<%=iDocNo&"Z"&iPayNo&"Z"&iSno%>')">
			</td>
			<Input type="hidden" name="hCheckType<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="C">
		</tr>

	<%			iSno=cint(iSno)+1
			End IF
			objRs.MoveNext
		Loop

	objRs.Close

IF CStr(iAdvNo) <> "0" Then
	sQuery = "Select Distinct B.PayRecNo,B.BillAmt,B.BillAdjAmt,R.CreatedPayablesNumber,H.VoucherNumber, "&_
			 "Convert(Varchar(10),H.VoucherDate,103) VoucherDate,H.CreatedTransNo From "&_
			 "Acc_T_BillsClosingDetails B,Acc_T_Payables R, " &_
			 "Acc_T_VoucherHeader H Where R.PayablesNumber = B.PayRecNo And  "&_
			 "B.AdvCloseNo = "&iAdvNo&" And H.TransactionNumber = B.BillTrNo  "&_
			 "And H.BookCode = '07' "

	'Response.Write sQuery

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	iRecCount = objRs.RecordCount
	Set objRs.ActiveConnection = nothing
	Do While Not objRs.EOF
		blnCreditNode = false
		sInvNo = objRs("VoucherNumber")
		sInvDate = objRs("VoucherDate")
		iPayNo= objRs("PayRecNo")
		iDocNo = objRs("CreatedPayablesNumber")
		sAmtReceiavble = Trim(objRs("BillAmt"))
		sAmtReceived = CDbl(objRs("BillAmt")) - CDbl(objRs("BillAdjAmt"))
		dToAccount= 0
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)


%>

		<tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?PI?<%=objrs(6)%>">
			<td class="ExcelSerial" align="center"><%=iSno%></td>
		    <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
			<td class="ExcelDisplayCell" align="right">
				<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly onchange="DisplayAmt('<%=iDocNo&"Z"&iPayNo&"Z"&iSno%>')">
			</td>
			<Input type="hidden" name="hCheckType<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sPrintAdjusted%>" value="C">
			
       </tr>

<%
		iSno=cint(iSno)+1
		objRs.MoveNext
	Loop
objRs.Close
End IF
'Response.Write "<p>blnCreditNode="& blnCreditNode 
if blnCreditNode then
%>
	<tr>
		<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Credit Notes Found !!</b></td>
	</tr>
<%
end if
End IF 'if Trim(sCallFrom) = "PAY" Then
if Trim(sCrDrType)="D" then
  sDebitAdjNarr = "Total Debit Note Adjusted"
else
  sDebitAdjNarr = "Total Credit Note Adjusted"
end if
%>
	<tr>
		<td colspan="8" class="ExcelDisplayCell" align="Right"><b><%=sDebitAdjNarr%></b></td>
		<td class="ExcelDisplayCell" align="right">
			<Input type="Text" Name="TxtTotDebitNoteAdj" style="text-align:right" Value="0.00" Readonly Class="FormElemRead">
		</td>
	</tr>
 
<%IF CStr(iCrTransNo) <> "0" Then%>
	<tr>
		<td class="ExcelDisplayCell" colspan="9"><b><FONT color="red">*</Font> Denotes Bill(s) Adjusted to this Voucher </b></td>
	</tr>
<%End IF%>
	</table>                                              
		<Input type="hidden" name="hRecCount" value="<%=iRecCount%>">
  		</div>
  		</td>
  		<td align="center"></td>
  	</tr>

  	<tr>
  		<td align="center" class="MiddlePack" colspan="3">
  		</td>
  	</tr>

							<tr>
								<td align="center">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                            <%'IF CStr(iCtr) <> "0" Then %>
																	<input type="button" value="Done" name="B8" class="ActionButton" onClick="finaldone()" >
																	<!--input type="button" value="Cancel" name="B8" class="ActionButton" onClick="finalcancel()" -->
																	<!--input type="reset" value="Reset" name="B9" class="ActionButton"-->
															<%'Else%>
																<input type="button" value="Cancel" name="B8" class="ActionButton" onClick="finalcancel()">
															<%'End IF%>
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
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


