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
	'Program Name				:	VouSALAccAmdEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	Feburary 14 2003
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
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo,sSalType,saTemp
dim sReferenceNo,sInvoiceNo,iBkAccHead,sPartyName,sInvDate,iTransNo
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,objfs,sExp,TempNode
Dim sVouNo,sVouDate,sSetInvDate,iSalTy,iSalAccHead,sSalAccHdName
Dim sCode,sValue,sAgentName,sSelUOM,sSelPack,sFlag,sFrmOthApp

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("hOrgName")

sSalType=Request.Form("hSalType")
'saTemp=Split(sSalType,"-- --")
'sSalType=saTemp(1)
iBookNo=Request.Form("selBook")
sReferenceNo=Request.Form("txtRefNo")
sInvoiceNo=Request.Form("txtInvoiceNo")
iBkAccHead=Request.Form("hBkAccHead")
sPartyName=Request.Form("txtPartyName")
sInvDate=Request.Form("hInvDate")
iSalTy = Request.Form("selSaleType")
iTransNo = Request("hTransNo")
sFlag = Request("sFlag")


oDOM.load server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")


Set Root = oDOM.documentElement
Set objRs = Server.CreateObject("ADODB.RecordSet")

sExp = "//Party"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sPartyName = TempNode.Item(0).Text
End IF

sExp = "//SalesType"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sSalType = TempNode.Item(0).Text
	iSalTy = TempNode.Item(0).Attributes.Item(0).nodeValue
End IF

sExp = "//Voucher"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sVouNo = TempNode.Item(0).Attributes.getNamedItem("CreatedVouNo").Value
End IF

sExp = "//Organization"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sOrgId = TempNode.Item(0).Attributes.getNamedItem("OrgId").Value
	sOrgName = TempNode.Item(0).Text
End IF

sExp = "//SaleInvoice"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sReferenceNo = TempNode.Item(0).Attributes.getNamedItem("RefNo").Value
	sInvoiceNo = TempNode.Item(0).Attributes.getNamedItem("InvNo").Value
	sInvDate = TempNode.Item(0).Attributes.getNamedItem("InvDate").Value
End IF

sExp = "//Book"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iBookNo = TempNode.Item(0).Attributes.getNamedItem("BookId").Value
End IF

sExp = "//Details"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sVouDate = TempNode.Item(0).Attributes.getNamedItem("VouDate").Value
End IF

sExp = "//Agent"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sAgentName = TempNode.Item(0).Attributes.getNamedItem("Agentname").Value
End IF



sQuery = "Select AccountHead From App_R_OrgnTaxAccountHead Where TaxCode is Null and TaxCategoryCode  "&_
		 "is Null and InvoiceType = "&iSalTy&" and OUDefinitionID = '"&sOrgId&"'  "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	iSalAccHead = objRs(0)
Else
	iSalAccHead = 0
End IF
objRs.Close

sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where AccountHead = "&iSalAccHead&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sSalAccHdName = objRs(0)
Else
	sSalAccHdName = ""
End IF
objRs.Close

sQuery = "Select isNull(FromApplication,'0') From Acc_T_CreatedVoucherHeader Where CreatedTransNo = " & iTransNo
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sFrmOthApp = objRs(0)
Else
	sSalAccHdName = ""
End IF
objRs.Close




%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=10">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData">
<Details BasicValue="" Discount="" ActualValue="" VouDate="<%=sInvDate%>"/></XML>
<XML id="EntryData"><Entry No="0" PayTo="" Amount="" Qty="" UOM="" UOMValue="" Rate="" ActValue="" DisPer="" DisAmount="" RndOff="" NoofPack="" PackType="" RatePer="" ItemCode="" ClassCode="" /></XML>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData" src="<%="../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"%>"></XML>
<XML id="GLHeadData"><Root/></XML>
<XML id="ItemData"><Root/></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/checkdate.js"></script>

<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<script language="javascript" src="../../scripts/SalesVoucherEntryCompat.js"></script>
<script language="javascript" src="../../scripts/VouSALAmdEntryCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitVouSALAmdEntry()">

<form method="POST" name="formname" action="VouSALAmdTaxEntry.asp">
<input type="hidden" name="hVouCode" value="05">
<input type="hidden" name="hVouName" value="SJR">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hInvDate" value="<%=sInvDate%>">
<input type="hidden" name="hRefNo" value="<%=sReferenceNo%>">
<input type="hidden" name="hInvNo" value="<%=sInvoiceNo%>">
<input type="hidden" name="hSalAccCode" value="<%=iSalAccHead%>">
<input type="hidden" name="hSalAccName" value="<%=sSalAccHdName%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hItemCode" value="0">
<input type="hidden" name="hClassCode" value="0">
<input type="hidden" name="hAmdTy" value="A">
<input type="hidden" name="hFrmOthApp" value="<%=Trim(sFrmOthApp)%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Accounted Sales Voucher Amendment
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
								  		<td align="center">Advance</td>
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
                            <!--tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
							<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
							<span style="cursor: hand" Title="Month wise Balance" >
							<p align="center"><font size="4" face="Webdings">?</font>
							</span>
							</td>
							<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
							<p align="center">
							<span style="cursor: hand" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
							</span>
							</p>
							</td>
							<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
							<p align="center">
							<span style="cursor: hand" Title="Voucher History">
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
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr-->
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" colspan="2">
                              <table border="0" cellspacing="0" width="100%" class="TableOutlineOnly" cellpadding="0">
                                <tr>
                                  <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                <tr>
                                <%
									sQuery = "Select H.OUDefinitionID,H.BookNumber,H.PayToRecdFrom,convert(char,H.VoucherDate,103) VoucherDate,H.CreatedVoucherNo,D.OrgUnitShortDescription from Acc_T_CreatedVoucherHeader H inner join " _
									& "DCS_OrganizationUnitDefinitions D on H.OUDefinitionID=D.OUDefinitionID where H.CreatedTransNo="&iTransNo
									  	With objRs
									  		.CursorLocation = 3
									  		.CursorType = 3
									  		.Source = sQuery
									  		.ActiveConnection = con
									  		.Open
									  	End with
									  	Set objRs.Activeconnection = nothing
									  	if not objRs.EOF then
											sOrgName =objRs("OrgUnitShortDescription")
											iBookNo=objRs("BookNumber")
											sReferenceNo=objRs("CreatedVoucherNo")
											sInvoiceNo =objRs("PayToRecdFrom")
											'sInvoiceNo =Left(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-")-1)
											'sSetInvDate=Mid(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-"))
											'sInvDate=objRs("VoucherDate")
										objRs.Close
										end if
                                %>
                                  <!--<td class="FieldCellSub" width="165">Unit Name</td>
                                  <td class="FieldCell"><span class="DataOnly"><%=sOrgName%>&nbsp;<span></td>-->

								  <td class="FieldCellSub" width="100">Voucher No-Date</td>
								  <td class="FieldCell" width="165">
								  <span class="DataOnly"><%=sVouNo%>-<%=sVouDate%></span></td>

                                </tr>
                                <tr>
                                  <td class="FieldCellSub" width="165">Party Name</td>
                                  <td class="FieldCell" colspan="3"> <span class="DataOnly"><%=sPartyName%></span></td>
                                 </tr>

                                 <tr>

                                  <td class="FieldCellSub" width="75">Sale Type</td>
                                  <td class="FieldCell" colspan="2"><span class="DataOnly"><%=sSalType%>&nbsp;<span></td>
                                </tr>
                                <tr>
                                  <td class="FieldCellSub" width="165">Reference / Invoice Number</td>
                                  <td class="FieldCell">                            	<span class="DataOnly"><%=sReferenceNo%>&nbsp;/&nbsp;<%=sInvoiceNo%>&nbsp;</span>
</td>
                                  <td class="FieldCellSub" width="75">Invoice Date</td>
                                  <td class="FieldCell" width="125"><span class="DataOnly"><%=sInvDate%></span>     </td>
                                </tr>
                                <%IF CStr(sAgentName) <> "" Then %>
									<tr>

									  <td class="FieldCellSub" width="75">Agent Name</td>
									  <td class="FieldCell" colspan="2"><span class="DataOnly"><%=sAgentName%>&nbsp;<span></td>
									</tr>
								<%End IF %>

                                <tr>
                                  <td class="MiddlePack" colspan="4"></td>
                                </tr>
                              </table>
                            </td>
                                </tr>
                                <tr>
                            <td class="MiddlePack" colspan="2"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Sales Account Head</td>
                            <td class="FieldCell">
                            <select size="1" name="selAccountHead" class="FormElem" onChange="popSalesHead(this) ">
							<%IF CStr(iSalAccHead) = "0" Then %>
								<option value="S" Selected>Sales Account Head</option>
								<option value="G">GL Account Head</option>
							<%Else%>
								<option value="S">Sales Account Head</option>
								<option value="G"  Selected>GL Account Head</option>
							<%End IF %>
                            </select>
                            </td>
                            <Input type="hidden" name="hHeadCount" value="1">
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115"></td>
                            <td class="FieldCell">

                            <%IF CStr(iSalAccHead) <> "0" Then %>
								<span class="DataOnly" id="spAccHead"><%=sSalAccHdName%> </span>
							<%Else%>
								<span class="DataOnly" id="spAccHead"></span>
							<%End IF %>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Item Description</td>
                            <td class="FieldCell">
                            <input type="text" name="txtDescription" size="40" class="FormElem">
                            <a href="#" onClick="GetItem()">
									<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item Description" width="15" height="15">
								</a>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Quantity</td>
                            <td class="FieldCell">
							    <table border="0" cellpadding="0" cellspacing="0">
							     <tr>
							       <td width="65"></td>
							       <td><input type="text" name="txtQty" value="0.00" size="13"  maxlength="11" style="text-align: Right" class="FormElem"></td>
							       <td width="10">
							       </td>
							       <td>
							   <select size="1" name="selUOM" class="FormElem" onChange="ChDisp(this)">
							<%

									sQuery = "Select UoMCode,UoMShortDescription from Ms_UnitOfMeasurement"

								  	With objRs
								  		.CursorLocation = 3
								  		.CursorType = 3
								  		.Source = sQuery
								  		.ActiveConnection = con
								  		.Open
								  	End with

								  	IF Not objRs.EOF Then
								  		sSelUOM = objRs(0)
								  	End IF

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
							   <td class="FieldCell">&nbsp;&nbsp;In</td>
							   <td class="FieldCell"><input type="text" name="txtBagno" class="FormElem" size="6" style="text-align: Right"></td>
							   <td>
							   <select size="1" name="selPack" class="FormElem" onChange="ChDisp(this)">
							   <option value="0">Select</option>
							<%

									sQuery = "Select PackingCode,PackingShortName From APP_M_PackingType Order By PackingShortName "

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

								  	IF Not objRs.EOF Then
								  		sSelPack = objRs(1)
								  	End IF


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
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtRate" value="0.00" onBlur="calculateField(1)" size="15"  maxlength="13" style="text-align:right" class="FormElem"></td>
                            <td class="FieldCell">&nbsp;&nbsp;Per</td>
								<td class="FieldCell"><input type="text" name="txtRatePer" class="FormElem" size="6" style="text-align: Right" value="1" onBlur="calculateField(1)">
								&nbsp;&nbsp;<span id="spUOM" class="ExcelDisplayCell"><%=sSelUOM%> </span>&nbsp;&nbsp;
								In a <span id="spPack" class="ExcelDisplayCell"><%=sSelPack%> </span>
								</td>

                              </tr>

                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Actual Value</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtValue" readonly size="15" value="0.00" maxlength="13" style="text-align:right" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Discount</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="60" class="FieldCell"><input type="text" name="txtDisPercentage" onBlur="calculateField(2)" size="6"  maxlength="5" style="text-align:right" value="0" class="FormElem">%</td>
                                <td>
                            <input type="text" name="txtDisAmount" size="15" value="0.00" onBlur="calculateField(3)"  maxlength="13" style="text-align:right" value="0" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Sales Value</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtAmount" size="15" readonly value="0.00" maxlength="13" style="text-align:right" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                  <tr>
                            <td class="FieldCell" width="115">Approval</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td class="FieldCell">
                           <Input type="radio" name="optApproval" value="Y" class="FormElem" checked>Yes &nbsp;&nbsp;
                           <Input type="radio" name="optApproval" value="N" class="FormElem">No &nbsp;&nbsp;
                           </td>
                              </tr>
                            </table>
                                  </td>
                                  <tr>
                            <td class="FieldCell">Rounded Off</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td class="FieldCell">
                           <Input type="radio" name="optRound" value="Y" class="FormElem" >Yes &nbsp;&nbsp;
                           <Input type="radio" name="optRound" value="N" class="FormElem" checked>No &nbsp;&nbsp;
                           </td>
                              </tr>

                            </table>
                                  </td>
                                </tr>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="1">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                                                                 <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
                                                                <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >
                                                                <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >
                                                                <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouSalBookSelection.asp')" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="35">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                                        <td class="ExcelHeaderCell" align="center">Account Head - Item Description</td>
                                        <td class="ExcelHeaderCell" align="center" width="60">Quantity</td>
                                        <td class="ExcelHeaderCell" align="center">Rate</td>
                                        <td class="ExcelHeaderCell" align="center">Value</td>
                                        <td class="ExcelHeaderCell" align="center">Discount</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                            </tr>
                                                </table>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5" >
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
