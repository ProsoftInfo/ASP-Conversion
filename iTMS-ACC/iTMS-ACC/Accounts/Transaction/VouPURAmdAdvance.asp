<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURAmdAdvance.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu R
	'Created On					:	OCT 28, 2004
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
<!--#include file="../../include/Accpopulate.asp"-->

<%
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTax,oNodItem
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo,sorgName
dim sParCode,sParSubType,sParType,dItemValue,dInvAmount,dBasicValue
dim bCommFlag,ObjDOM
dim dBasicTotal,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue
dim sInvoiceNo,sInvoiceDt,Objrs2,iCrTransNo,sAmdType,sFlag,iVouCrNo,sFormVal
Dim ChkRoot,sSelAdvNo,sRetVal,sAdVouDate

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set ObjDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

sAmdType = Request.Form("hAmdType")
sFormVal = Request.Form("hFormVal")
sFlag = Request.Form("hFlag")
iVouCrNo = Request("hTransNo")
iCrTransNo = Request("hTransNo")

oDOM.load server.MapPath("../temp/transaction/Voucher Amd_PUR_"&Session.SessionID&".xml")
Set oNodRoot=oDOM.documentElement

'ObjDOM.load server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
sRetVal = GetVouchXML(iCrTransNo)
ObjDOM.Load server.MapPath(sRetVal)

Set ChkRoot=ObjDOM.documentElement
sExp = "//AdvanceDetails/Advance"
Set TempNode = ChkRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		sSelAdvNo = sSelAdvNo & ","& TempNode.Item(iCtr).Attributes.getNamedItem("AdvNo").Value
	Next
End IF
sSelAdvNo = Mid(sSelAdvNo,2)


bCommFlag=false
for each oNodDeatils in oNodRoot.childNodes
	if oNodDeatils.nodeName="Header" then
		for Each oNodEntry in  oNodDeatils.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sorgName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SalesType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="PurInvoice" then
				sInvoiceNo=oNodEntry.Attributes.getNamedItem("PurInvNo").Value
				sInvoiceDt=oNodEntry.Attributes.getNamedItem("PurInvDate").Value
			end if
		next
	end if
	if oNodDeatils.nodeName="Details" then
		sVouDate = oNodDeatils.Attributes.getNamedItem("VouDate").Value
		IF CStr(oNodDeatils.Attributes.Item(2).nodeValue) <> "" Then
			dBasicValue=oNodDeatils.Attributes.Item(2).nodeValue
			set oNodItem=oNodDeatils
		End IF
	end if


	if oNodDeatils.nodeName="TaxDetails" then
		dInvAmount=oNodDeatils.Attributes.Item(0).nodeValue
		set oNodTax	=oNodDeatils
	end if
	if oNodDeatils.nodeName="AgentDetails" then
		set oNodAgent=oNodDeatils
		bCommFlag=true
	end if


next
dim iItemCount,dTemp,dItemTotal
dim saTemp
dim dBasItemValue

'------TO UPDATE THE ITEM VALUE FOR THE TAX WITHOUT ACCOUNT HEAD------
iItemCount=CInt(oNodItem.childNodes.length)

dTemp=0
dItemTotal = 0
dBasItemValue = 0

FOR EACH oNodDeatils in oNodItem.childNodes
	dItemValue=oNodDeatils.Attributes.Item(2).nodeValue
	dBasItemValue = oNodDeatils.Attributes.Item(0).nodeValue

	For Each oNodEntry in oNodTax.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		sFormula=oNodEntry.Attributes.Item(3).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue

		'Response.Write sTaxMode &" "

		if sTaxMode="F" then
			dTax=CDbl(dTaxValue)/iItemCount
		else
			'dTax=CalculateTax(oNodTax,sFormula,dBasItemValue,dItemValue,dTaxValue)
		end if
		dTax = Round(dTax,2)

		'saTemp= Split(CStr(dTax),".")
		'if UBound(saTemp)>0 then
			'oNodEntry.Attributes.Item(7).nodeValue=saTemp(0)&"."&left(saTemp(1),2)
		'else
		oNodEntry.Attributes.Item(7).nodeValue=dTax
		'end if
	next


	For Each oNodEntry in oNodTax.childNodes
		'Response.Write " ================= "
		if cint(oNodEntry.Attributes.Item(6).nodeValue)=0 then
			'dItemValue=FormatNumber(CDbl(dItemValue),2,,,0)+CDbl(oNodEntry.Attributes.Item(7).nodeValue)
			dItemValue=Round(CDbl(dItemValue),2)+CDbl(oNodEntry.Attributes.Item(7).nodeValue)
			'Response.Write dItemValue &"   " & CDbl(oNodEntry.Attributes.Item(7).nodeValue) &"<br>"

		else
			'dTemp = Round(CDbl(dTemp)+CDbl(oNodEntry.Attributes.Item(7).nodeValue),2)
		end if
	next

	dItemTotal=CDbl(dItemValue)+CDbl(dItemTotal)
	oNodDeatils.Attributes.Item(2).nodeValue=dItemValue
next

'Response.Write dTemp &" " & iItemCount &" " & dItemTotal
dItemTotal=CDbl(dTemp/iItemCount)+CDbl(dItemTotal)

dTemp=oNodItem.childNodes(CInt(oNodItem.childNodes.length)-1).Attributes.Item(2).nodeValue
'dTemp=CDbl(dTemp)+ Round(CDbl(dInvAmount)-CDbl(dItemTotal),2)

oNodItem.childNodes(CInt(oNodItem.childNodes.length)-1).Attributes.Item(2).nodeValue=dTemp
'Response.write dTemp &" " & dItemValue &" " & dItemTotal

'============== Calculation for Amount Match =====================================
'============== Included on 04/05/2004 =====================================
Dim iActualVal,sExp,TempNode,iCtr,iTotalActVal,iTotalNewVal,iTotTaxVal,iDiffAmt
Dim iOldVal,iNewVal
iTotalActVal = 0
iTotalNewVal = 0

sExp = "//Entry"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
'		iTotalActVal = iTotalActVal + TempNode.Item(iCtr).Attributes.getNamedItem("ActValue").value
		iTotalActVal = iTotalActVal + TempNode.Item(iCtr).Attributes.getNamedItem("ActValue").value - TempNode.Item(iCtr).Attributes.getNamedItem("DisAmount").value
		iTotalNewVal = iTotalNewVal + TempNode.Item(iCtr).Attributes.getNamedItem("Amount").value
	Next
End IF
sExp = "//Tax"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		IF CStr(TempNode.Item(iCtr).Attributes.getNamedItem("AccHead").Value) = "0" Then
			iTotTaxVal = Cdbl(iTotTaxVal) + Cdbl(TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").value)
		End IF
	Next
End IF

iActualVal = CDbl(iTotalActVal + iTotTaxVal)
iDiffAmt = CDbl(iActualVal - iTotalNewVal)

Set oNodRoot = oDOM.documentElement
sExp = "//Entry"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then

	iOldVal = TempNode.Item(0).Attributes.getNamedItem("Amount").value
	iOldVal = CDbl(iOldVal)
	iNewVal = iOldVal + iDiffAmt
	iNewVal = abs(iNewVal)
	TempNode.Item(0).Attributes.Item(2).NodeValue = iNewVal
End IF

oDOM.save server.MapPath("../temp/transaction/Voucher Amd_PUR_"&Session.SessionID&".xml")

'Response.Write iNewVal


'============== Calculation for Amount Match =====================================

%>

<%
function CalculateTax(oNodTaxRoot,sFormula,dBValue,dDValue,dPercentage)
	dim saTemp,dTaxAmount,iCounter,iTemp
	dim oNodTemp
	dim saTemp1

	saTemp=Split(sFormula,",")
	if trim(saTemp(0))="BV" then
		dTaxAmount=dBValue
		iTemp=1
	elseif trim(saTemp(0))="BD" then
		dTaxAmount=dDValue
		iTemp=1
	else
		dTaxAmount=0
		iTemp=0
	end if

	for iCounter=iTemp to UBound(saTemp)
		saTemp1=Split(trim(saTemp(iCounter)),"#")
		For Each oNodTemp in oNodTaxRoot.childNodes
			if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
				dTaxAmount=Round(CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(7).nodeValue),2)
			end if
		next
	next

	if trim(dPercentage)<>"" then
		CalculateTax=dTaxAmount*(cdbl(dPercentage)/100)
	else
		CalculateTax=dTaxAmount
	end if
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<XML id="AdvanceData" src="<%="../temp/transaction/Voucher Amd_PUR_"&Session.SessionID&".xml"%>"></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/AdvanceAdjustmentCompat.js"></SCRIPT>
<script language="javascript">
ITMSAdvanceAdjustmentCompat.install({
	invoiceAmount: "<%=dInvAmount%>",
	saveUrl: "XMLSave.asp?Mod=PUR&Name=Voucher AMD",
	subtractToAccount: true,
	zeroBlankToAccount: true,
	emptyFieldAction: {
		fieldName: "hAmdType",
		emptyAction: "VouPURAmdGenerate.asp",
		valueAction: "AmdAccPURGenerate.asp"
	}
});
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="VouPURAmdGenerate.asp">
<Input type="hidden" name="hAmdType" value="<%=sAmdType%>">
<Input type="hidden" name="hTransNo" value="<%=iCrTransNo%>">
<input type="hidden" name="hFormVal" value="<%=sFormVal%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Voucher  	</td>
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
								<!--<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>-->
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<td align="center">Invoice Details</td>
									</tr>
								  </table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								  	<tr>
									  	<td align="center">Advance</td>
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
                            <!--tr>
                            <td align="center" colspan="3" class="MiddlePack" height="7">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <tr>
                            <td align="center" height="39">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center" height="39">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: hand" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">ª</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Daywise Balance"><font size="3" face="Webdings">¦</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Voucher History">
                    <font size="4" face="Webdings">¨</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell">
                    &nbsp;
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center" height="39">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="75">Party Name</td>
                                                    <td colspan="3" class="FieldCellSub"><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="75">Invoice No</td>
                                                    <td class="FieldCellSub" width="200">  <span class="DataOnly"><%=sInvoiceNo%>&nbsp;</span></td>

                                                    <td class="FieldCellSub" width="110"><p align="left">Invoice Date</p></td>
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=sInvoiceDt%>&nbsp;</span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub">Invoice Amount</td>
                                                    <td width="145" class="FieldCellSub"><span class="DataOnly"><%=dInvAmount%>&nbsp;</span></td>

                                                        </tr>
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
<%if bCommFlag then%>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="125"><p align="center">Commission
                                                              Details
                                                            </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack> </td>
														</tr>

														<tr>
															<td class=FieldCellSub width="10">
												<DIV class=frmBody id=frm3 style="width: 336; height:100;">
                                                                    <table border="0" id = "tblBin0" cellspacing ="1" class="ExcelTable" width="100%">
                                                                      <tr>
                                                                        <td class="ExcelHeaderCell" align="center" width="25">S.No.</td>
                                                                        <td class="ExcelHeaderCell" align="center">Agent Name</td>
                                                                        <td class="ExcelHeaderCell" align="center" colspan=2>Commission</td>
                                                                      </tr>
							<%
							dim sAgentName,sAgentCode,sCommType,dCommision,sCurrency
							iSno=0

							For Each oNodEntry in oNodAgent.childNodes
									iSno=CDbl(iSno)+1
									sAgentCode=oNodEntry.Attributes.Item(0).nodeValue
									sAgentName=oNodEntry.Attributes.Item(1).nodeValue
									sCommType=oNodEntry.Attributes.Item(2).nodeValue
									dCommision=oNodEntry.Attributes.Item(3).nodeValue
									sCurrency=oNodEntry.Attributes.Item(5).nodeValue
									if sCommType="P" then
										dCommision=(CDbl(dCommision)*dBasicValue)/100
									end if
									oNodEntry.Attributes.Item(6).nodeValue=dCommision
							%>
																	<tr>
                                                                        <td class="ExcelHeaderCell" align="center" width="25"><%=iSno%></td>
                                                                        <td class="ExcelHeaderCell" align="left"><%=sAgentName%></td>
                                                                        <td class="ExcelHeaderCell" width="40" align="right"><%=sCurrency%></td>
                                                                        <td class="ExcelHeaderCell" width="75" align="right"><%=dCommision%></td>
                                                                      </tr>

<%
next


%>
                                                                    </table>
												</div>
                                                   </td>
												</center>
														</tr>

														<tr>
															<td class=MiddlePack width="267">
                                                   </td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
                                                        </div>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
<%end if%>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
												<DIV class=frmBody id=frm2 style="width: 580; height:240;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"  rowspan="2">&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center" colspan="4">Document</td>
                                        <td class="ExcelHeaderCell" align="center" colspan="3">Amount</td>

                                            </tr>
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center">Number</td>
                                        <td class="ExcelHeaderCell" align="center">Date</td>
                                        <td class="ExcelHeaderCell" align="center">Details</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                        <td class="ExcelHeaderCell" align="center">Adjusted</td>
                                        <td class="ExcelHeaderCell" align="center">To Account</td>
                                        <td class="ExcelHeaderCell" align="center">To be adjusted</td>
                                            </tr>
<%
dim iDocNo,iSno,oNodAdvRoot,newElem
dim dAmtPaid,dAmtAdjusted,sVouDate,sVouNo,iCrDocNo,sRowCheck,dToAccount,dToPayable,iAdvNo
Dim sVouType,sRecdFrom,sInstType,sInstNo,sInstDate,sInstBank,iCrAdvNo,dTransAdjAmt,sInstDet
Set oNodAdvRoot = oDOM.createElement("AdvanceDetails")

'sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(char,b.VoucherDate,103),"&_
'		 "isnull(a.AdvanceAdjusted,0),isnull(a.AdvancePaid,0),a.CreatedTransNo,A.AdvanceNumber from Acc_T_AdvancePayments a ,Acc_T_VoucherHeader b" &_
'		 " where a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"'"&_
'		 " and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and "&_
'		 "isnull(a.AdvancePaid,0)>isnull(a.AdvanceAdjusted,0)  and b.TransactionNumber=a.TransactionNumber"

'sQuery = "Select Distinct TransactionNumber,VoucherNumber,VoucherDate,AdvanceAdjusted, "&_
'		 "AdvancePaid,TransactionType,PayToRecdFrom,BankInstrumentType,BankInstrumentNo, "&_
'		 "BankInstrumentDate,DrawnOnBank,CreatedTransNo,CreatedAdvanceNo,isNull(AmountAdjusted,0),isNull(ToAccount,0) ToAccount From  "&_
'		 "VwPurAdvanceDet Where OUDefinitionID = '"&sOrgId&"' and PartyType = '"&sParType&"' and PartySubType = "&sParSubType&" "&_
'		 "and PartyCode = "&sParCode&" and CONVERT(DATETIME,VOUCHERDATE,103)<= CONVERT(DATETIME,'"&sVouDate&"',103) "

'sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(char,b.VoucherDate,103),"&_
'		 "isnull(a.AdvanceAdjusted,0),isnull(a.AdvancePaid,0),a.CreatedTransNo,A.AdvanceNumber, "&_
'		 "isNull(C.AdvanceAdjusted,0) CrAdvanceAdj,isnull(a.AdvancePaid,0) - isnull(a.AdvanceAdjusted,0) ToAdjust "&_
'		 "from Acc_T_AdvancePayments a ,Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C" &_
'		 " where a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"'"&_
'		 " and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and "&_
'		 "isnull(a.AdvancePaid,0)>isnull(a.AdvanceAdjusted,0)  and b.TransactionNumber=a.TransactionNumber "&_
'		 "and B.CreatedTransNo = C.CreatedTransNo "


'********************** Previous Non Adjusted Bills Will COme ***********************************************************************************************
IF Cstr(sSelAdvNo) = "" Then
	sSelAdvNo = 0
End IF
sQuery = "SELECT A.TRANSACTIONNUMBER,B.VOUCHERNUMBER,CONVERT(CHAR,B.VOUCHERDATE,103), "&_
		 "ISNULL(A.ADVANCEADJUSTED,0),ISNULL(A.ADVANCEPAID,0),B.TRANSACTIONTYPE,B.PAYTORECDFROM, "&_
		 "ISNULL(B.BANKINSTRUMENTTYPE,''),ISNULL(B.BANKINSTRUMENTNO,0),ISNULL(CONVERT(CHAR,B.BANKINSTRUMENTDATE,103),'01/01/1900'),ISNULL(DRAWNONBANK,''),  "&_
		 "A.CREATEDTRANSNO, A.ADVANCENUMBER,ISNULL(C.ADVANCEADJUSTED,0) - ISNULL(A.ADVANCEADJUSTED,0)  "&_
		 "TOACCOUNT,C.CREATEDADVANCENO  FROM ACC_T_ADVANCEPAYMENTS A, ACC_T_VOUCHERHEADER B,  "&_
		 "ACC_T_CREATEDADVANCES C WHERE A.OUDEFINITIONID='"&sOrgId&"' AND A.PARTYTYPE='"&sParType&"' AND  "&_
		 "A.PARTYSUBTYPE="&sParSubType&" AND A.PARTYCODE="&sParCode&" AND ISNULL(A.ADVANCEPAID,0)>ISNULL(A.ADVANCEADJUSTED,0)  "&_
		 "AND B.TRANSACTIONNUMBER = A.TRANSACTIONNUMBER AND B.VOUCHERDATE<=  "&_
		 "CONVERT(DATETIME,'"&sVouDate&"',103) AND A.CREATEDADVANCENO = C.CREATEDADVANCENO   "&_
		 "And C.CREATEDADVANCENO Not In("&sSelAdvNo&") "


'Response.write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
If not objRs.EOF then
	iSno=0
	Do While Not objRs.EOF

		iDocNo = objRs(0)
		sVouNo = objRs(1)
		sAdVouDate = objRs(2)
		dAmtAdjusted = objRs(3)
		dAmtPaid = objRs(4)
		sVouType = objRs(5)
		sRecdFrom = objRs(6)
		sInstType = objRs(7)
		sInstNo = objRs(8)
		sInstDate = objRs(9)
		sInstBank = objRs(10)
		dTransAdjAmt = objRs(13)
		dToAccount = objRs(13)
		iCrAdvNo = objRs(14)


		sQuery = "Select AmountAdjusted From ACC_T_CreatedAdvanceAdj Where  "&_
				 "CreatedAdvanceNo = "&iCrAdvNo&" and CreatedTransNo = "&iCrTransNo&" "

	'	Response.Write sQuery &"<br><br>"
		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dTransAdjAmt = Objrs2(0)
		Else
			dTransAdjAmt = 0
		End IF
		Objrs2.Close

		'Response.Write dToAccount &" -- " & dTransAdjAmt &"<br>"
		IF CDbl(dToAccount) > 0 and CStr(sAmdType) <> "A"   Then
			dToAccount = CDbl(dToAccount) - CDbl(dTransAdjAmt)
		Elseif CStr(sAmdType) <> "A" Then
			dToAccount = 0
		Elseif CStr(sAmdType) = "A" Then
			dAmtAdjusted = CDbl(dAmtAdjusted) - CDbl(dTransAdjAmt)
		End IF

		dToAccount = FormatNumber(dToAccount,2,,,0)

		'Response.Write dToAccount &" -- " & dAmtPaid &"<br>"
		IF CDbl(dAmtPaid) > CDbl(dToAccount) Then

			Set newElem = oDOM.createElement("Advance")
			newElem.setAttribute "TransNo",iDocNo
			newElem.setAttribute "VoucherNo",trim(sVouNo)
			newElem.setAttribute "VoucherDate",trim(sAdVouDate)
			newElem.setAttribute "AmountRec",dAmtPaid
			newElem.setAttribute "AmountAdj",dAmtAdjusted
			newElem.setAttribute "AmountToAdj","0"
			newElem.setAttribute "CreatedTransNo", iCrDocNo
			newElem.setAttribute "AdvNo", iCrAdvNo
			newElem.setAttribute "AdjType", "I"
			newElem.setAttribute "ToAccount", dToAccount
			oNodAdvRoot.appendChild newElem

			IF CStr(Trim(sInstDate)) = "01/01/1900" Then
				sInstDate = ""
			End IF

			if sVouType="CAR" then
				sInstDet="Received From : &nbsp; " &sRecdFrom
			else
				sInstDet=sInstType
				sInstDet=sInstDet &"No:&nbsp;"& sInstNo
				sInstDet=sInstDet &"<br> Date:&nbsp;"& sInstDate
				sInstDet=sInstDet &"<br>Bank:&nbsp;"& sInstBank
			end if
			iSno = Cdbl(iSno) + 1


%>
              <tr>
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<%IF CDbl(dTransAdjAmt) <> 0 Then %>
					<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>" value="<%=iDocNo%>" class="FormElem" Checked></td>
				<%Else%>
					<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>" value="<%=iDocNo%>" class="FormElem"></td>
				<%End IF %>
				<td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sAdVouDate%></td>
				<td class="ExcelDisplayCell"><p align="left"><%=sInstDet%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtPaid,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>

				<td class="ExcelInputCell" align="right">
				<input type="text" style="text-align:right" name="txtAmount<%=iDocNo%>Z<%=iCrAdvNo%>" value="<%=FormatNumber(dTransAdjAmt,2,,,0)%>" maxlength="13" size="15" class="Formelem"> </td>
			 </tr>

<%
			End IF
		objRs.MoveNext
	Loop
End IF
objRs.Close

'*****************Only Previous Adjusted Bills Will Come *****************************************************************
sQuery = "SELECT A.TRANSACTIONNUMBER,B.VOUCHERNUMBER,CONVERT(CHAR,B.VOUCHERDATE,103), "&_
		 "ISNULL(A.ADVANCEADJUSTED,0),ISNULL(A.ADVANCEPAID,0),B.TRANSACTIONTYPE,B.PAYTORECDFROM, "&_
		 "ISNULL(B.BANKINSTRUMENTTYPE,''),ISNULL(B.BANKINSTRUMENTNO,0),ISNULL(CONVERT(CHAR,B.BANKINSTRUMENTDATE,103),'01/01/1900'),ISNULL(DRAWNONBANK,''),  "&_
		 "A.CREATEDTRANSNO, A.ADVANCENUMBER,ISNULL(C.ADVANCEADJUSTED,0) - ISNULL(A.ADVANCEADJUSTED,0)  "&_
		 "TOACCOUNT,C.CREATEDADVANCENO  FROM ACC_T_ADVANCEPAYMENTS A, ACC_T_VOUCHERHEADER B,  "&_
		 "ACC_T_CREATEDADVANCES C WHERE A.OUDEFINITIONID='"&sOrgId&"' AND A.PARTYTYPE='"&sParType&"' AND  "&_
		 "A.PARTYSUBTYPE="&sParSubType&" AND A.PARTYCODE="&sParCode&"   "&_
		 "AND B.TRANSACTIONNUMBER = A.TRANSACTIONNUMBER AND B.VOUCHERDATE<=  "&_
		 "CONVERT(DATETIME,'"&sVouDate&"',103) AND A.CREATEDADVANCENO = C.CREATEDADVANCENO   "&_
		 "And C.CREATEDADVANCENO In("&sSelAdvNo&") "

		 '"ISNULL(A.ADVANCEPAID,0) - ISNULL(C.ADVANCEADJUSTED,0) - ISNULL(A.ADVANCEADJUSTED,0) > 0  "


'Response.Write sQuery &"<br><br>"

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
If not objRs.EOF then
	iSno=0
	Do While Not objRs.EOF

		iDocNo = objRs(0)
		sVouNo = objRs(1)
		sAdVouDate = objRs(2)
		dAmtAdjusted = objRs(3)
		dAmtPaid = objRs(4)
		sVouType = objRs(5)
		sRecdFrom = objRs(6)
		sInstType = objRs(7)
		sInstNo = objRs(8)
		sInstDate = objRs(9)
		sInstBank = objRs(10)
		dTransAdjAmt = objRs(13)
		dToAccount = objRs(13)
		iCrAdvNo = objRs(14)


		sQuery = "Select AmountAdjusted From ACC_T_CreatedAdvanceAdj Where  "&_
				 "CreatedAdvanceNo = "&iCrAdvNo&" and CreatedTransNo = "&iCrTransNo&" "

'		Response.Write sQuery &"<br><br>"
		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dTransAdjAmt = Objrs2(0)
		Else
			dTransAdjAmt = 0
		End IF
		Objrs2.Close

		'Response.Write dToAccount &" -- " & dTransAdjAmt &"<br>"
		IF CDbl(dToAccount) > 0 and CStr(sAmdType) <> "A"   Then
			dToAccount = CDbl(dToAccount) - CDbl(dTransAdjAmt)
		Elseif CStr(sAmdType) <> "A" Then
			dToAccount = 0
		Elseif CStr(sAmdType) = "A" Then
			dAmtAdjusted = CDbl(dAmtAdjusted) - CDbl(dTransAdjAmt)
		End IF

		dToAccount = FormatNumber(dToAccount,2,,,0)

		'Response.Write dToAccount &" -- " & dAmtPaid &"<br>"
		IF CDbl(dAmtPaid) > CDbl(dToAccount) Then

			Set newElem = oDOM.createElement("Advance")
			newElem.setAttribute "TransNo",iDocNo
			newElem.setAttribute "VoucherNo",trim(sVouNo)
			newElem.setAttribute "VoucherDate",trim(sAdVouDate)
			newElem.setAttribute "AmountRec",dAmtPaid
			newElem.setAttribute "AmountAdj",dAmtAdjusted
			newElem.setAttribute "AmountToAdj","0"
			newElem.setAttribute "CreatedTransNo", iCrDocNo
			newElem.setAttribute "AdvNo", iCrAdvNo
			newElem.setAttribute "AdjType", "I"
			newElem.setAttribute "ToAccount", dToAccount
			oNodAdvRoot.appendChild newElem

			IF CStr(Trim(sInstDate)) = "01/01/1900" Then
				sInstDate = ""
			End IF

			if sVouType="CAR" then
				sInstDet="Received From : &nbsp; " &sRecdFrom
			else
				sInstDet=sInstType
				sInstDet=sInstDet &"No:&nbsp;"& sInstNo
				sInstDet=sInstDet &"<br> Date:&nbsp;"& sInstDate
				sInstDet=sInstDet &"<br>Bank:&nbsp;"& sInstBank
			end if
			iSno = Cdbl(iSno) + 1


%>
              <tr>
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<%IF CDbl(dTransAdjAmt) <> 0 Then %>
					<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>" value="<%=iDocNo%>" class="FormElem" Checked></td>
				<%Else%>
					<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>" value="<%=iDocNo%>" class="FormElem"></td>
				<%End IF %>
				<td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sAdVouDate%></td>
				<td class="ExcelDisplayCell"><p align="left"><%=sInstDet%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtPaid,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>

				<td class="ExcelInputCell" align="right">
				<input type="text" style="text-align:right" name="txtAmount<%=iDocNo%>Z<%=iCrAdvNo%>" value="<%=FormatNumber(dTransAdjAmt,2,,,0)%>" maxlength="13" size="15" class="Formelem"> </td>
			 </tr>

<%
			End IF
		objRs.MoveNext
	Loop
End IF
objRs.Close


'=================================================================================
	'Added On		:	29/11/2004
	'Reason			:	To display credit notes for this party.
	'=================================================================================
	'sQuery = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Char,isNull(P.PartyInvoiceDate,H.VoucherDate),103),P.AmountReceivable, "&_
	'		 "P.AmountReceived,C.AmountReceived - P.AmountReceived ToAccount,P.ReceivableNumber,P.DRCreatedReceivable "&_
	'		 "From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where "&_
	'		 "P.CreatedReceivable = 0 and P.Narration is Null and  "&_
	'		 "P.AmountReceivable > P.AmountReceived and P.OUDefinitionID = '"&sOrgId&"' and P.PartyType = 'CR' "&_
	'		 "and P.PartyCode = "&sParCode&" and P.TransactionNumber = H.TransactionNumber "&_
	'		 "and C.ReceivableNumber = P.DRCreatedReceivable "

	'sQuery = "Select P.PayablesNumber,H.VoucherNumber,Convert(Char,isNull(P.PartyBillDate,H.VoucherDate),103), "&_
	'		 "P.AmountPayable, P.AmountPaid,C.AmountPaid - P.AmountPaid,P.PayablesNumber  "&_
	'		 "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C  "&_
	'		 "Where P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable >  "&_
	'		 "P.AmountPaid and P.OUDefinitionID = '"&sOrgId&"' and P.PartyType = 'CR' and  "&_
	'		 "P.PartyCode = "&sParCode&" and P.TransactionNumber = H.TransactionNumber and  "&_
	'		 "C.PayablesNumber = P.CRCreatedPayable "

	'Response.Write sQuery &"<br><br>"

	'with objRs
	'	.CursorLocation = 3
	'	.CursorType = 3
	'	.Source = sQuery
	'	.ActiveConnection = con
	'	.Open
	'end with
	'set objRs.ActiveConnection = nothing



	'Do While Not objRs.EOF
	'
	'	iDocNo = objRs(0)
	'	sVouNo = objRs(1)
	'	sVouDate = objRs(2)
	'	iCrDocNo = objRs(6)
	'	iAdvNo = objRs(7)
	'
	'	sQuery = "Select AmountReceived From Acc_T_CreatedRcvbleAdjDet Where  "&_
	'			 "CreatedTransNo = "&iCrTransNo&" and ReceivableNumber = "&iAdvNo&" "
	'
	'	'Response.Write sQuery &"<br>"
	'
	'	Objrs2.Open sQuery,Con
	'	IF Not Objrs2.EOF Then
	'		dTransAdjAmt = Objrs2(0)
	'	Else
	'		dTransAdjAmt = 0
	'	End IF
	'	Objrs2.Close
	'
	'
	'	dToPayable = Trim(objRs(3))
	'	dAmtPaid = Trim(objRs(4))
	'	dToAccount = Trim(objRs(5))
'
	'	dToPayable = CDbl(dToPayable)
	'	dAmtPaid = CDbl(dAmtPaid)
	'	dToAccount = CDbl(dToAccount)
'
'
	'	'dAmtAdjusted = CDbl(dToPayable - dAmtPaid - dToAccount)
	'	dAmtAdjusted = dAmtPaid
	'	'Response.Write dToAccount &" -- " & dTransAdjAmt &"<br><br>"
	'
	'	IF CDbl(dToAccount) > 0 Then
	'		dToAccount = CDbl(dToAccount) - CDbl(dTransAdjAmt)
	'	Else
	'		dToAccount = 0
	'	End IF
'
	'	'Response.Write dToPayable &" -- " & dToAccount &"<br><br>"
	'	IF CDbl(dToPayable) > CDbl(dToAccount) or CDbl(dTransAdjAmt) > 0 Then
	'		sRowCheck = "Y"
	'		Set newElem = oDOM.createElement("Advance")
	'		newElem.setAttribute "TransNo",iDocNo
	'		newElem.setAttribute "VoucherNo",trim(sVouNo)
	'		newElem.setAttribute "VoucherDate",trim(sVouDate)
	'		newElem.setAttribute "AmountRec",dToPayable
	'		newElem.setAttribute "AmountAdj",dAmtPaid
	'		newElem.setAttribute "AmountToAdj","0"
	'		newElem.setAttribute "CreatedTransNo", iCrDocNo
	'		newElem.setAttribute "AdvNo", "0"
	'		newElem.setAttribute "AdjType", "D"
	'		newElem.setAttribute "ToAccount", dToAccount
'
	'		oNodAdvRoot.appendChild newElem
'
	%>

	<%

			'End IF
			'objRs.MoveNext
			'iSno=cint(iSno)+1
		'Loop


		if CStr(sRowCheck) = "N" then
			Set oNodAdvRoot = oDOM.createElement("AdvanceDetails")
			oNodRoot.appendChild oNodAdvRoot
			oDOM.save server.MapPath("../temp/transaction/Voucher Amd_PUR_"&Session.SessionID&".xml")
			IF CStr(sAmdType) = "" Then 'Called from Created or Approved stage of amedment
				'Response.Redirect("VouPURGenerate.asp")
			Else
				'Response.Redirect("AmdAccPURGenerate.asp")
			End IF
		end if

	'objRs.Close

	Dim CheckNode,iChkCtr
	sExp = "//AdvanceDetails"
	Set CheckNode = oNodRoot.selectNodes(sExp)
	IF CheckNode.length <> 0 Then
		'CheckNode.removeAll
		For iChkCtr = 0 To CheckNode.length - 1
			oNodRoot.RemoveChild CheckNode.Item(iChkCtr)
		Next
	End IF

oNodRoot.appendChild oNodAdvRoot
oDOM.save server.MapPath("../temp/transaction/Voucher Amd_PUR_"&Session.SessionID&".xml")
%>
                                                </table>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell">
                                                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr>
                                        <td valign="middle" class="ActionCell"><p align="center">
                                        <input type="button" value="Next" name="B7" class="ActionButton" onClick="actionDone()">
                                        <input type="button" value="Cancel" name="B8" class="ActionButton"></td>
                                            </tr>
                                                </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="BottomPack">
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