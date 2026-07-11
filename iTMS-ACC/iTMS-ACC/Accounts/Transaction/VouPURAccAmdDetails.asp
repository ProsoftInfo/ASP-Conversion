<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURAccAmdDetails.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu R
	'Created On					:	Oct 28, 2004
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
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo,sVal,sValTemp
dim sReferenceNo,sInvoiceNo,iBkAccHead,sPartyName,sSetInvDate
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,objfs,iTransNo
Dim sExp,TempNode,sVouDate
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sBkAccDesc
dim sCode,sValue,bCostcenter,bAnalytical,sAMdTy,sFlag

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF
sAMdTy = Request("AmdType")

'Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

sOrgId=Request.Form("selUnitId")
iBookNo=Request.Form("selBook")
sReferenceNo=Request.Form("txtReferenceNo")
sInvoiceNo=Request.Form("txtInvoiceNo")
iBkAccHead=Request.Form("hBkAccHead")
sPartyName=Request.Form("txtPartyName")
sSetInvDate = Request.Form("hInvDate")
sVouDate = Request.Form("hVouDate")

iTransNo = Request("hTransNo")
sFlag = Request("sFlag")

'Response.Write sFlag
'IF objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")) then
'	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"))
'End IF
'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
'Response.Write iTransNo
'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")

oDOM.load server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")

Set Root = oDOM.documentElement
sExp = "//Header/Book"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iBkAccHead = TempNode.item(0).Attributes.getNamedItem("BKOtherUnits").Value
End IF

sExp = "//Party"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sPartyName = TempNode.Item(0).Text
End IF

sExp = "//Organization"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sOrgId = TempNode.Item(0).Attributes.getNamedItem("OrgId").Value
	sOrgName = TempNode.Item(0).Text
End IF

sExp = "//Book"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iBookNo = TempNode.Item(0).Attributes.getNamedItem("BookId").Value
End IF

sExp = "//PurInvoice"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sInvoiceNo = TempNode.Item(0).Attributes.getNamedItem("PurInvNo").Value
	sSetInvDate = TempNode.Item(0).Attributes.getNamedItem("PurInvDate").Value
End IF

'Response.Write sInvoiceNo

sExp = "//Details"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sVouDate = TempNode.Item(0).Attributes.getNamedItem("VouDate").Value
End IF

Set objRs = Server.CreateObject("ADODB.RecordSet")
sQuery = "Select M.AccountHead,M.AccountDescription From Acc_M_GLAccountHead M, "&_
		 "VwOrgBookNames V Where V.BookCode = '04' and V.BookNumber = "&iBookNo&" "&_
		 "and V.OUDefinitionID = '"&sOrgId&"' and V.BookAccountHead = M.AccountHead "
objRs.Open sQuery,con
IF Not objRs.EOF Then
	iBkAccHead = objRs(0)
	sBkAccDesc = objRs(1)
Else
	iBkAccHead = 0
	sBkAccDesc = ""
End IF
objRs.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<meta http-equiv="x-ua-compatible" content="IE=10">
<XML id="DetData">
<Details BasicValue="" Discount="" ActualValue="" VouDate=""/></XML>
<XML id="EntryData"><Entry No="0" PayTo="" Amount="" Qty="" UOM="" UOMValue="" Rate="" ActValue="" DisPer="" DisAmount="" ItemCode="" ClassCode=""/></XML>
<XML id="TaxData"></XML>

<XML id="AccHeadData">
<account/>
</XML>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData" src="<%="../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"%>"></XML>
<XML id="GLHeadData"><Root></Root></XML>


<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/checkdate.js"></script>
<SCRIPT SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script src="../../scripts/VouTransactions.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/PurchaseVoucherEntryCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitPurchaseVoucherAmdDetails()">

<form method="POST" name="formname" action="VouPURAmdTaxEntry.asp" >
<input type="hidden" name="hVouCode" value="04">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hSetInvDate" value="<%=sSetInvDate%>">
<input type="hidden" name="hVouDate" value="<%=sVouDate%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hSalAccCode" value="<%=iBkAccHead%>">
<input type="hidden" name="hSalAccName" value="<%=sBkAccDesc%>">
<input type="hidden" name="hAmdType" value="<%=sAMdTy%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hItemCode" value="0">
<input type="hidden" name="hClassCode" value="0">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Voucher Amendment	</td>
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
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Invoice Details</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center"> Advance</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr><td align="center">Voucher</td></a>
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
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                        <tr>
                            <td class="FieldCell" width="115">Party Name </td>
                            <td class="FieldCell" colspan="3">
                            <span class="DataOnly"><%=sPartyName%></span>
                            </td>


                                </tr>
                                <tr>
                                <%
									'sQuery = "Select OUDefinitionID,BookNumber,PayToRecdFrom,VoucherDate from Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
									'Response.Write sQuery
									 ' 	With objRs
									  '		.CursorLocation = 3
									  '		.CursorType = 3
									  '		.Source = sQuery
									  '		.ActiveConnection = con
									  '		.Open
									  '	End with
									  '	Set objRs.Activeconnection = nothing
										'	sOrgId=objRs("OUDefinitionID")
											'iBookNo=objRs("BookNumber")
											'IF InStr(1,objRs("PayToRecdFrom"),"-") <> 0 Then
											'	sInvoiceNo =Left(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-")-1)
											'	sSetInvDate=Mid(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-"))
											'Else
											'	sInvoiceNo = ""
											'	sSetInvDate = ""
											'End IF

											'if sSetInvDate="" then
											'	sSetInvDate=objRs("VoucherDate")
											'end if
										'objRs.Close
                                %>
                            <td class="FieldCell" width="115">Invoice Number </td>
                            <td class="FieldCell" width="125">
                            	<span class="DataOnly" id="tInvNo"><%=sInvoiceNo%></span>
                            </td>
                            <td width="90" class="FieldCell">
                                Voucher Date
                            </td>
                            <td class="FieldCell">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlDate")
%>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="135">Purchase Account Head</td>
                            <td class="FieldCell" colspan="3">
                            <% IF CStr(iBkAccHead) = "0" Then %>
								<select size="1" name="selAccountHead" class="FormElem" onChange="popSalesHead(this) ">
								<option value="S" Selected>Purchase Account Head</option>
								<option value="G">GL Account Head</option>
							<%Else%>
								<select size="1" name="selAccountHead" class="FormElem" onChange="popSalesHead(this)" disabled>
								<option value="S">Purchase Account Head</option>
								<option value="G" Selected>GL Account Head</option>
							<%ENd IF %>
                            </select>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="125"></td>
                            <td class="FieldCell" colspan="3">
                            <% IF CStr(iBkAccHead) = "0" Then %>
								<span class="DataOnly" id="spAccHead"></span>
							<%Else%>
								<span class="DataOnly" id="spAccHead"><%Response.Write(sBkAccDesc)%> </span>
							<%End IF %>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Item Description</td>
                            <td class="FieldCell" colspan="3">
                            <input type="text" name="txtDescription" size="50" class="FormElem"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Quantity</td>
                            <td class="FieldCell" colspan="3" align="left">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td><input type="text" name="txtQty" size="15"  maxlength="14" style="text-align: Right" class="FormElem" value="0.00"></td>
                                <td width="10">
                                </td>
                                <td>
                            <select size="1" name="selUOM" class="FormElem">
                         <%

								sQuery = "Select UoMCode,UoMShortDescription from Ms_UnitOfMeasurement"

							  	With objRs
							  		.CursorLocation = 3
							  		.CursorType = 3
							  		.Source = sQuery
							  		.ActiveConnection = con
							  		.Open
							  	End with
							  	Set objRs.Activeconnection = nothing
							  	Set sCode = objRs(0)
							  	Set sValue = objRs(1)

							  	Do while not objRs.EOF
									Response.Write "<option value="""&sCode&""">"&sValue&"</option>"
									objRs.MoveNext
								Loop
								objRs.Close
							%>
                            </select></td>
                              </tr>
                            </table>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Rate</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtRate" onBlur="calculateField(1)" size="15"  maxlength="13" style="text-align:right" class="FormElem" value="0.00"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Actual Value</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtValue" size="15"  maxlength="13" style="text-align:right" class="FormElem" value="0.00"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Discount</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="60" class="FieldCell"><input type="text" name="txtDisPercentage" onBlur="calculateField(2)" size="6"  maxlength="5" style="text-align:right" value="0" class="FormElem">%</td>
                                <td>
                            <input type="text" name="txtDisAmount" size="15" onBlur="calculateField(3)"  maxlength="13" style="text-align:right" value="0.00" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>

                                <tr>
                            <td class="FieldCell" width="115">Purchase Value</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtAmount" size="15" readonly maxlength="13" style="text-align:right" class="FormElem" onBlur="popAddAmount1()" value="0.00"></td>
                              </tr>


                            </table>
                                  </td>
                                </tr>

                                <!--tr>
                            <td class="FieldCell" width="108">Purchase Type&nbsp;</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selPurType" class="FormElem" >
									<option value="0">Select Purchase Type</option>
									<%
										'dim sCode,sValue
										sQuery = "Select PurchaseType,PurchaseTypeName from APP_M_PurchaseTypes Where Active = 'Y' Order By PurchaseTypeName "
									  	With objRs
									  		.CursorLocation = 3
									  		.CursorType = 3
									  		.Source = sQuery
									  		.ActiveConnection = con
									  		.Open
									  	End with
									  	Set objRs.Activeconnection = nothing
									  	Set sCode = objRs(0)
									  	Set sValue = objRs(1)
									  	Do while not objRs.EOF
									%>
											<option value="<%Response.Write sCode%>"><%Response.Write sValue%></option>
									<%

											objRs.MoveNext
										Loop
										objRs.Close
								    %>
								</select>
                            </td>
                                </tr-->
                           <!--tr>
									<td class="FieldCell" width="115">VAT Elgibility</td>
									<td class="FieldCell" colspan="3" align="left">
										<input type="checkbox" value="Y" name="chkVatElg" class="FormElem" Checked> YES
									</td>
                            </tr-->

                                <tr>
                            <td class="FieldCell" width="115">Approval</td>
                            <td class="FieldCell" colspan="3">
                            <!--table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td class="FieldCell" -->
                            <input type="radio" value="Y" checked name="optApprove" class="FormElem">
                             Yes&nbsp;&nbsp;
                            <input type="radio" value="N" name="optApprove" class="FormElem"> No
                            </td>
                              <!--/tr>


                            </table>
                                  </td-->
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
								<td valign="top" class="FieldCell" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Add Entry" name="btnAdd" class="ActionButton" onclick="AddEntry('A')" >
                                                                <input type="button" value="Update" onClick="AddEntry('U')" name="btnUpdate" class="ActionButton" disabled>
                                                                <input type="button" value="Delete" onClick="DelEntry()" name="btnDel" class="ActionButton" disabled>
                                                                <input type="button" value="Next" onClick="AddEntry('S')" name="btnNext" class="ActionButton" >

                                                               <input type="button" value="Cancel" name="btnCancel" onClick="Cancel('VouPURBookSelection.asp')" class="ActionButton" >
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
								<td align="center" width="5" class="ClearPixel" >&nbsp;
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" >
												<DIV class=frmBody id=DisVoucher style="width: 600; height:140;">
                                                <table border="0" id="tblVoucher" cellspacing="1" class="ExcelTable" width="584">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center">Account Head</td>
                                        <td class="ExcelHeaderCell" align="center">Rate</td>
                                        <td class="ExcelHeaderCell" align="center">Quantity</td>
                                        <td class="ExcelHeaderCell" align="center">Value</td>
                                        <td class="ExcelHeaderCell" align="center">Discount</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                            </tr>
                                                </table>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5" >
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
