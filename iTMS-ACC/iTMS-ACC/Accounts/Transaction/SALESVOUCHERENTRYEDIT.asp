
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SALESVOUCHERENTRYEDIT.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	May 25, 2011
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
<!--#include file="../../include/Salpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include File="../../include/CheckACCPrevFinYear.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo,sSalType,saTemp
dim sReferenceNo,sInvoiceNo,iBkAccHead,sPartyName,sInvDate,sSetInvDate
Dim sAccUnit,sAccUnitName,iSalTy,iSalAccHead,sSalAccHdName,sAgName,sPartyCode
Dim sAccBookRel,sName,nCreatedTransNo,nPartyType,nSaleType,sPayToRecFrom,nInvNo

sOrgId=Session("organizationcode")
sOrgName=Session("OrgShortName")
sAccUnit = sOrgId
sAccUnitName= sOrgName

Set objRs = Server.CreateObject("ADODB.RecordSet")

nCreatedTransNo = Request.QueryString("nTransNo")
Response.Write "<p><font color=red>nCreatedTransNo="&nCreatedTransNo

sQuery = " SELECT A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.PARTYTYPE,"&_
		 " A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,A.BANKINSTRUMENTTYPE,V.PARTYNAME,A.BOOKCODE,A.PayToRecdFrom,A.PartySubType,A.PARTYCODE "&_
		 " FROM ACC_T_CREATEDVOUCHERHEADER AS A INNER JOIN APP_M_PARTYMASTER AS V "&_
		 " ON A.PARTYCODE=V.PARTYCODE WHERE  A.OUDEFINITIONID='"& sOrgId &"' AND "&_
		 " isNull(A.OtherApplnTransNo,0) = 0 AND A.CREATEDTRANSNO = '"& nCreatedTransNo &"' "
objRs.Open sQuery,con
Response.Write sQuery
If Not objRs.EOF Then
	nPartyType = objrs(2)
	sPartyName = objRs(6)
	iBookNo = objrs(7)
	nSaleType = objRs(5)
	sPayToRecFrom = objrs(8)
	sPartyCode = nPartyType & "?"& objrs(9)& "?" & "" & "?"& objRs(10)
End IF
objRs.Close 
If sPayToRecFrom <> "" Then 
	If InStr(1,sPayToRecFrom,"-") Then
		nInvNo = split(sPayToRecFrom,"-")(0)
		sSetInvDate = split(sPayToRecFrom,"-")(1)
	Else
		nInvNo = sPayToRecFrom
		sSetInvDate = date()
	End IF
End IF 

Dim objfs,oDOM,sRetVal
Set objfs = CreateObject("Scripting.FileSystemObject")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sRetVal = GetVouchXML(nCreatedTransNo)
oDOM.Load server.MapPath(sRetVal)

if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")) then
	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"))
End IF

oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<meta http-equiv="x-ua-compatible" content="IE=10">
<XML id="DetData">
<Details BasicValue="" Discount="" ActualValue="" VouDate=""/></XML>
<XML id="EntryData"><Entry No="0" PayTo="" Amount="" Qty="" UOM="" UOMValue="" Rate="" ActValue="" DisPer="" DisAmount="" ItemCode="" ClassCode="" TransBasicamt="" TransRate="" TransDisAmt="" TransInvAmt="" RatePer="" NoofPack="" PackType="" RndOff="" /></XML>
<XML id="AccHeadData">
<account/>
</XML>
<xml id="PartyData"><Root /></xml>
<xml id="OutData"><Root /></xml>
<xml id="UnitBookData"><Book /></xml>
<xml id="SaleTypeData"><Book /></xml>
<!--<xml id="VoucherData"><Voucher/></xml>-->
<XML ID="VoucherData" src="<%="../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"%>"></XML>
<xml id="TEMPXML"><Root></Root></xml>
<xml id="GLHeadData"><Root></Root></xml>
<xml id="ItemData"><Root></Root></xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<script language="javascript" src="../../scripts/checkdate.js"></script>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script language="javascript" src="../../scripts/VouTransactions.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script language="javascript" src="../../scripts/SalesVoucherEntryCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="InitSalesVoucherEntryEdit()">

<form method="POST" name="formname" action="VouSALAmdTaxEntry.asp">
<input type="hidden" name="hVouCode" value="04">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hAccUnit" value="<%=sAccUnit%>">
<input type="hidden" name="hAccUnitName" value="<%=sAccUnitName%>">
<input type="hidden" name="hSalAccCode" value="<%=iSalAccHead%>">
<input type="hidden" name="hSalAccName" value="<%=sSalAccHdName%>">
<input type="hidden" name="hItemCode" value="0">
<input type="hidden" name="hClassCode" value="0">
<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
<input type="hidden" name="hParCode" value="">
<input type="hidden" name="hCommName" value="">
<input type="hidden" name="hCurrDate" value="<%=date()%>">
<input type="hidden" name="hBkAccHead" value="">

<Input type="hidden" name="hPartyType" value="<%=nPartyType%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hTransNo" value="<%=nCreatedTransNo%>">
<input type="hidden" name="hInvDate" value="<%=sSetInvDate%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher Entry
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
								<!--<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>-->
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
								<td class="TabCell" valign="bottom" align="center" width="100">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Commission</td>
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
								  	<tr><td align="center">Voucher</td></tr>
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
			        <td height="20" valign="bottom">
			            <table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>

								<td valign="top" width="100%" align=left>
								    <table cellpadding="0" cellspacing="0" width="100%" class="TableOutLineOnly">
                                    <tr>
                                            <td height="8" colspan=2></td>
                                    </tr>
                                    <tr>
                                            <td class="FieldCellSub" width="108">Sales Book</td>
                                            <td class="FieldCell">
												<select size="1" name="selBook" class="FormElem" onChange="PopulateSalTy()">
													<option value="S">Select Book</option>
												</select>
                                            </td>
                                            <td class="FieldCellSub" width="108">Reference Number</td>
                                            <td class="FieldCell"><input type="text" name="txtRefNo" size="20" maxlength="30" class="FormElem">
                                            </td>
                                    </tr>
                                        <tr>
                                            <td class="FieldCellSub" width="108">Sale Type&nbsp;</td>
                                            <td class="FieldCell">
                                            	<select size="1" name="selSaleType" class="FormElem" onchange="GetAccHead(this)" >
									                <option value="0">Select Sale Type</option>
									                <%

										                IF CStr(sAccBookRel) <> "T" Then 'Book and Account Head is Not Done
											                sQuery = "Select InvoiceType,InvTypeShortName,InvoiceTypeName from Sal_M_InvoiceTypes where TobeAccounted=1 and Useable = 1 Order By InvoiceTypeName "
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
									  		                Set sName = objRs(2)
									  		                Do while not objRs.EOF
									  							If Trim(sCode) = trim(nSaleType) Then
									  							%>
																	<option value="<%Response.Write sCode%>" selected><%=trim(sName)%></option>
																<%
																Else
																%>
																	<option value="<%Response.Write sCode%>"><%=trim(sName)%></option>
																<%
																End IF
												                objRs.MoveNext
											                Loop
											                objRs.Close
										                End IF
								                    %>
								                </select>
                                            </td>
                                            <td class="FieldCellSub" width="108">Invoice Number</td>
											<td class="FieldCell">
												<input type="text" name="txtInvoiceNo" size="20" class="FormElem" onblur="VouCreate()">
                                            </td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub" width="108">Party Name</td>
                                            <td class="FieldCell" valign=middle>
                                                <input type="text" name="txtPartyName" size="40" class="FormElemRead" value="<%=sPartyName%>">
                                                &nbsp;&nbsp;<a href="#" onclick="SelPartyName()">
                                                <img src="../../assets/images/iTMS%20icons/Entry.gif" alt='Select Party Name' height="10" width="10"></a>
                                            </td>
                                            <td class="FieldCellSub" width="120"> Invoice Date</td>
											<td class="FieldCell">
												  <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
												  %>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub" width="108">Party Type</td>
                                            <td class="FieldCell" colspan="3">
                            	                <select size="1" name="selParType" class="FormElem" onchange="PartyType(this)">
								                <option value="A">Select Party Type</option>
								                </select>
                                              </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub" width="108">Agent </td>
                                            <td class="FieldCell" colspan="3">
                                                <table border="0" cellpadding="0" cellspacing="0">
                                                  <tr>
                                                    <td class="FieldCell">
                                                        <input type="radio" value="C" name="optAgentExist" onClick="showAgent('1')" class="Formelem">
                                                    </td>
                                                    <td class="FieldCell">Commission Agent</td>
                                                    <td class="FieldCell">
                                                        <input type="radio" value="D" name="optAgentExist" onClick="showAgent('2')" class="Formelem">
                                                    </td>
                                                    <td class="FieldCell">Depo Agent</td>
                                                    <td class="FieldCell"> <input type="radio" value="No" onClick="showAgent('N')" checked name="optAgentExist" class="Formelem">
                                                    </td>
                                                    <td class="FieldCell">
                                                  		No Agent</td>
                                                  	<td>
                                                  		<span ID="spAgentName" class="DataOnly"></span>
                                                  	</td>
                                                  </tr>
                                                </table>
                                            </td>
                                        </tr>
										<!--tr>
											<td class="FieldCell" width="108">&nbsp;</td>
											<td class="FieldCell"><span ID="spAgentName" class="DataOnly"></span></td>
											<td class="FieldCell" width="120"></td>
											<td class="FieldCell"></td>
										</tr-->
                                    </table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
        			        </td>
			                </tr>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td valign="top" width="100%" >
                                <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                <td class="MiddlePack" colspan="2"></td>
                                </tr>
                                <tr>
                            <td class="FieldCellSub">Sales Account Head</td>
                            <td class="FieldCell">
                            <select size="1" name="selAccountHead" class="FormElem" onfocus="VouCreate()" onChange="popSalesHead(this)" >
                            <%IF CStr(iSalAccHead) = "0" Then %>
								<option value="S" Selected>Sales Account Head</option>
								<option value="G">GL Account Head</option>
							<%Else%>
								<option value="S">Sales Account Head</option>
								<option value="G"  Selected>GL Account Head</option>
							<%End IF %>
                            </select>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell"></td>
                            <td class="FieldCell">
                            <%IF CStr(iSalAccHead) <> "0" Then %>
								<span class="DataOnly" id="spAccHead"><%=sSalAccHdName%> </span>
							<%Else%>
								<span class="DataOnly" id="spAccHead"></span>
							<%End IF %>
							</td>
                                </tr>
                                <tr>
                            <td class="FieldCellSub">Item Description</td>
                            <td class="FieldCell">
								<input type="text" name="txtDescription" size="40" class="FormElem" onfocus="VouCreate()">
                                <a href="#" onClick="GetItem()">
									<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item Description" width="15" height="15">
								</a>
							</td>
							</tr>
                            <tr>
                            <td class="FieldCellSub">Quantity</td>
                            <td class="FieldCell">
							    <table border="0" cellpadding="0" cellspacing="0">
							     <tr>
							       <td width="65"></td>
							       <td><input type="text" name="txtQty" value="0.00" size="13"  maxlength="11" style="text-align: Right" class="FormElem" onBlur="calculateField(1)"></td>
							       <td width="10">
							       </td>
							       <td>
							   <select size="1" name="selUOM" class="FormElem" onChange="ChDisp(this)">
							<%
									Dim sCode,sValue,sSelUOM,sSelPack
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

								  	IF Not objRs.EOF Then
								  		sSelUOM = objRs(1)
								  	End IF

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
                            <td class="FieldCellSub">Rate</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
									<input type="text" name="txtRate" value="0.00" onBlur="calculateField(1)" size="15"  maxlength="13" style="text-align:right" class="FormElem">
								</td>
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
                            <td class="FieldCellSub">Actual Value</td>
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
                            <td class="FieldCellSub">Discount</td>
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
                            <td class="FieldCellSub">Sales Value</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtAmount" size="15" readonly value="0.00" maxlength="13" style="text-align:right" class="FormElem" onBlur="popAddAmount1()"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                  <tr>
                            <td class="FieldCellSub">Approval</td>
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
                                </tr>
                                <tr>
                            <td class="FieldCellSub">Rounded Off</td>
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
								<td valign="top" class="FieldCell" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Add Entry" name="btnAdd" class="ActionButton" onclick="AddEntry('A')" >
                                                                <input type="button" value="Update" onClick="AddEntry('U')" name="btnUpdate" class="ActionButton" disabled>
                                                                <input type="button" value="Delete" onClick="DelEntry()" name="btnDel" class="ActionButton" disabled>
                                                                <input type="button" value="Next" onClick="AddEntry('S')" name="btnNext" class="ActionButton" >
                                                                <input type="reset" value="Cancel" name="btnCancel" class="ActionButton" onClick="Cancel('VouSALBookSelection.asp')" >
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
									<td valign="top" class="FieldCell" >
												<DIV class=frmBody id=DisVoucher style="width: 100%; height:140;">
                                                <table border="0" id="tblVoucher" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center">Account Head</td>
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
