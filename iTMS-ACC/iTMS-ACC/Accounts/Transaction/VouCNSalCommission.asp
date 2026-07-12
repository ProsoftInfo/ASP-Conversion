<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalCommission.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran
	'Created On					:	Feb 16,2010
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1,sTxtNarration
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff,sFromSal
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId,sTemp,iitemCode,iClassCode,iSalRetQty
Dim sRatePer,nArrCount,sTempVal,sCommtypename,sAgentCode,sSalTransNo

Dim sFinPeriod,sFinFrm,sFinTo,sValTemp,objDOM1,sdomRoot,sdomPage
Dim sCallFrom ,iCnt,sBookNumber,sInvFrom,sTempComm,sTempCommVal
Dim oDOMGJ,oDGjRoot,oDGjEntry,oDGjAcc,oDGjNarr,oDSubNode,nRatePer
sFinPeriod = Session("FinPeriod")
sValTemp = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp(0))
sFinTo = Trim(sValTemp(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"
iCnt =0
sBookNumber = Request.Form("selBook")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set oDOMGJ = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
Set objDOM1 = Server.CreateObject("Microsoft.XMLDOM")
sInvoiceNo=Request.Form("selInvoiceNo")
Response.Write sInvoiceNo
'Response.end
sTemp = Split(sInvoiceNo,":")
sBookName=Request.Form("hBookName")

sCallFrom= Request.QueryString("hCallFrom")
sInvoiceNo = sTemp(0)
iInvNo = sInvoiceNo
sTempComm = Request.Form("hVouDetails")
sTempCommVal=split(sTempComm,",")
sUserid = getUserID()

sQuery = "Select FromApplication From Acc_T_CreatedVoucherHeader "&_
		 "Where CreatedTransNo = "&iInvNo&" and FromApplication is Not NULL "
Response.Write sQuery
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sFromSal = "Y"
Else
	sFromSal = "N"
End IF
objRs.Close
'For GJVoucher
Dim sDocNO,sInvNo,sInvDate,sTransAmt,sAmtAdjusted,sAmtToAdjust,sDocType
Dim sAmtToAcc,sPayableNo, sAdjType,sAccPay,sAccRec,sAccType,sAccAdv
Dim sRetVal
'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")
sRetVal = GetVouchXML(sInvoiceNo)

oDOM.Load server.MapPath(sRetVal)
set oNodRoot = oDOM.documentElement
''CN Case XML Creation
for each oNodHeader in oNodRoot.childNodes
	if oNodHeader.nodeName="Header" then
		for Each oNodEntry in  oNodHeader.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.text
			end if
			if oNodEntry.nodeName="Book" then
				oNodEntry.Attributes.Item(0).nodeValue=Request.Form("selBook")
				oNodEntry.text=Request.Form("hBookName")
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue&"&nbsp;-&nbsp;"&oNodEntry.Attributes.Item(1).nodeValue
			end if
		next
	end if

	if oNodHeader.nodeName="Details" then
		set oNodDeatils=oNodHeader
	end if
	if oNodHeader.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodHeader
	end if
	if oNodHeader.nodeName="AgentDetails" then
		set oNodtemp=oNodRoot.removeChild(oNodHeader)
	end if
next

dim dInvAmount
sInvFrom = Request.Form("hInvVal")

'dInvAmount = sInvValue
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<meta http-equiv="x-ua-compatible" content="IE=edge">
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="TaxData" data-src="<%="../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="GJVoucher"></script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root/></script>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script src="../../scripts/VouCNSalCommissionCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitVouCNSalCommission()">
<form method="POST" name="formname">
<Input type="hidden" name="hInvCallFrom" value="<%=sInvFrom%>">
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hAccType" value="C">
<Input type="hidden" name="hCallType" value="OINV">
<Input type="hidden" name="hNoteType" value="C">
<Input type="hidden" name="hFromSal" value="<%=sFromSal%>">
<Input type="hidden" name="hOrgid" value="<%=sOrgId%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<Input type="hidden" name="hBookCode" value="<%=Request.Form("selBook")%>">
<input type="hidden" name="hCrAccHead" value="0">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hCallFromVoucher" value="<%=sCallFrom%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<%IF sCallFrom= "CR" then%>
<input type="hidden" name="hVouCode" value="07">
<%else%>
<input type="hidden" name="hVouCode" value="08">
<%end if%>
<input type="hidden" name="hVouType" value="">
<input type="hidden" name="hVouName" value="GJ">
<input type="hidden" name="optApprove" value="">
<input type="hidden" name="hInsVou" value="">
<input type="hidden" name="hSelVouDate" value="">
<input type="hidden" name="hTransNo" value="<%=iInvNo%>">
<input type="hidden" name="CallType" value="CN">
<input type="hidden" name="CrVouNo" value="">
<input type="hidden" name="SelTDSGrp" value="">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Credit Note For Sales Commission

		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
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
                            <!--tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: pointer" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
                    <font size="4" face="Webdings">?</font>
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
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="575">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Unit </td>
                                                    <td class="FieldCell">  <span class="DataOnly"><%=sOrgName%></span></td>
                                                    <td class="FieldCellSub" width="75"><p align="left">Date</p></td>
                                                    <td class="FieldCellSub" width="145">
                                                 <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Party
                                                      Name</td>
                                                    <td class="FieldCell" colspan="3">  <span id="sPartyName" class="DataOnly">&nbsp;</span></td>

                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Cr Note Against</td>
                                                    <td class="FieldCell" width="200"><span class="DataOnly">Sales Commission</span></td>
                                                    </tr>
                                                    <tr>
														<td class="FieldCellSub" width="200">Select Account Head</td>
														<td class="FieldCell" width="200">
														<Select name="SelAccountHd" class="FormElem" onChange="AccHead(this)">
															<Option Value="0">Select</Option>
															<Option Value="G">GL ACCOUNT HEAD</Option>
														</Select>
														</td>
														 <td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spAccHead"></span>
														 </td>
                                                     </tr>


                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            <tr>
                <td></td>
                <td valign="top" width="100%">
                <div class="frmBody" id="frm2" style="width: 575; height:150;">
            <table border="0" cellspacing="1" class="ExcelTable" width="575">
			    <tr>
					<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
					<td class="ExcelHeaderCell" align="center" colspan="2">Invoice</td>
					<td class="ExcelHeaderCell" align="center"  rowspan="2" width="150">Commission Type</td>
					<td class="ExcelHeaderCell" align="center"  rowspan="2" width="150">Commission Value</td>
			    </tr>
			    <tr>
						<td class="ExcelHeaderCell" align="center"  width="150" >Number</td>
						<td class="ExcelHeaderCell" align="center"  width="150" >Date</td>
			    </tr>
						<%

							set sdomRoot=objDOM1.createElement("Root")
							objdom1.appendChild sdomRoot

							iSno = 0
							sTxtNarration = "CR Note for "
							for nArrCount = 0 to UBound(sTempCommVal) - 1
								sTempVal = split(sTempCommVal(nArrCount),":")
								iSno = iSno + 1

								if Trim(sTempVal(5)) = "A" then
									sCommtypename = "Per Qty"
								elseif Trim(sTempVal(5)) = "B" then
									sCommtypename = "% of Basic"
								elseif Trim(sTempVal(5)) = "C" then
									sCommtypename = "% of Total Inv"
								elseif Trim(sTempVal(5)) = "P" then
									sCommtypename = "Per Packing"
								End if

								sTxtNarration = sTxtNarration & " : "& sTempVal(1) &" - "& sTempVal(2)
						%>
						    <tr>
						    <td class="ExcelSerial" align="center"><%=isno%></td>
						    <td class="ExcelDisplayCell"><%=sTempVal(1)%></td>
						    <td class="ExcelDisplayCell" ><%=sTempVal(2)%></td>
						    <td class="ExcelDisplayCell" align="Center"><%=sCommtypename%></td>
						    <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sTempVal(3),2,,,-1)%></td>
						    </tr>
						<%
								sAgentCode = sTempVal(8)
								dTotal = cdbl(dTotal) + cdbl(sTempVal(3))

								set sdomPage=objDOM1.createElement("CommDet")
								sdomPage.setAttribute "AgentCode",sTempVal(8)
								sdomPage.setAttribute "CommissionType",sTempVal(5)
								sdomPage.setAttribute "CurrCode",sTempVal(9)
								sdomPage.setAttribute "CommValue",sTempVal(3)
								objRs.Open "Select Saletransactionno from VwSalCommAccDet where TransactionNumber = "& sTempVal(0),con
								if not objRs.EOF then
									sSalTransNo = objRs(0)
								end if
								objRs.Close
								sdomPage.setAttribute "SalTransNo",sSalTransNo
								sdomPage.setAttribute "AccTransNo",sTempVal(0)
								sdomRoot.appendChild sdomPage

							next

							objDOM1.save server.MapPath("../temp/transaction/VoucherSalCommDet_"&Session.SessionID&".xml")

							with objRs
								.CursorLocation = 3
								.CursorType = 3
								.ActiveConnection = con
								.Source = "Select PartyType,PartySubType,R.PartyCode,PartyName from App_M_PartyMaster P,APP_R_orgparty R where  R.PartyCode = P.PartyCode and R.PartyCode = "& sAgentCode
								'Response.Write objRs.Source
								.Open
							end with
							if not objRs.EOF then
							%>
								<input type=hidden name="hAgentType" value="<%=objrs(0)%>">
								<input type=hidden name="hAgentSubType" value="<%=objrs(1)%>">
								<input type=hidden name="hAgentCode" value="<%=objrs(2)%>">
								<input type=hidden name="hAgentName" value="<%=objrs(3)%>">
							<%end if%>
        <tr>
			<Input type="hidden" name="hRowVal" value="<%=isno%>">
			<td class="ExcelSerial" align="center" colspan=4 ><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
			<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTotalInv" value="<%=FormatNumber(dTotal,2,,,0)%>" class="Formelem" size="13"></td>
			 <td align="right" colspan="3"></td>
        </tr><input type="hidden" name="hTxtnarr" value="<%=sTxtNarration%>">
			<%
				oDOM.save server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")
			%>
	        </table>

                </div>
                </td>
                <td></td>
                            </tr>
                            <tr>
                            <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Approval &nbsp;&nbsp;&nbsp;

			<Input type="radio" name="optApprove" checked value="Y" onClick="EnbApp(this)"> Yes &nbsp;&nbsp;&nbsp;
			<Input type="radio" name="optApprove" value="N" onClick="EnbApp(this)"> No &nbsp;&nbsp;&nbsp;
			</td>
        </tr>
        <tr>
			 <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Immediate Approver &nbsp;&nbsp;&nbsp;

			<select size="1" name="selUserId" class="FormElem">
              <option value="I">Immediate Approver</option>
                <%=populateEmployeeWithVal(sUserId)%>
                    </select>
			</td>
        </tr>
        <tr>
			 <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Narration &nbsp;&nbsp;&nbsp;

			<Textarea name="txtNarration" class="FormElem" cols="40" rows="4" maxlength="200" readonly ></Textarea>

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
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                                <input type="button" value="Next" name="B2" class="ActionButton" onClick="SaveXML()" >
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" onClick="Cancel('VOUCNBOOKSELECTION.ASP')" >
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
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</html>
