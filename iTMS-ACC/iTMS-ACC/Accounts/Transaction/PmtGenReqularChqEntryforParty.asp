<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtGenReqularChqEntryforParty.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April  24, 2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	April 06, 2011
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
dim objRs,sQuery,objRs1
dim sOrgName,iPymtNo,sPymtFor,sOrgId,iAmount
dim sDate,sRequestBy,dAmount,sPayTo,sReason,nTransactionNo
dim sTemp,sApprovedBy,iSno,sReqType,sReqFrom,sParTemp,iPartyVal
dim iParCode,iParSubType,sParType,sParName,sParSubTypeName
set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")

'sTemp=split(trim(Request("selRequestNo")),"?")

'iPymtNo=sTemp(0)
'sRequestBy=sTemp(1)
'sApprovedBy=sTemp(2)
iPymtNo = Trim(Request("RequestNo"))

sOrgName=Request("hUnitName")
sOrgId=Request("hUnitNo")
sReqType=Request("hReqTypeS")
'sReqFrom= Request("selReqFrom")

'Response.Write iPymtNo & sReqType
sQuery = "Select Distinct RequestedBy,isNull(ApprovedBy,0) from Acc_T_PaymentRequestHdr where PaymentRequestNo = "& iPymtNo &" "
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

IF Not objRs.EOF then

	with objRs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "select EmployeeName from Ms_EmployeeMaster where EmployeeNumber="&objRs(0)
		.ActiveConnection = con
		.Open
	end with
	set objRs1.ActiveConnection = nothing	
	IF Not objRs1.EOF Then		
		sRequestBy=objRs1(0)
	End IF
	objRs1.Close
				
	with objRs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "select EmployeeName from Ms_EmployeeMaster where EmployeeNumber="&objRs(1)
		.ActiveConnection = con
		.Open
	end with
	set objRs1.ActiveConnection = nothing	
	IF Not objRs1.EOF Then		
		sApprovedBy=objRs1(0)
	End IF
	objRs1.Close
	
End IF
objrs.Close 
'-------------Newly Added on Dec 28 th 2007 by Maheswari to fetch Party Name -------------------
sQuery = "Select AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,isNull(TransactionNumber,0) from Acc_T_PaymentRequestDet where PaymentRequestNo = "& iPymtNo &" "
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
'Response.Write "<p>sQuery="&sQuery
IF Not objRs.EOF then
	sParType	= objRs(0)
	iParSubType = objRs(1)
	iParCode	= objRs(2)
	nTransactionNo = objrs(3)
	
End IF 
objrs.Close 
If iParSubType <> "" Then
	sQuery = "Select isNull(PartyName,''),isNull(SubTypeName,'') from VwOrgParty where PartyCode = "&iParCode &" and PartyType = '"&sParType&"' and PartySubType = "& iParSubType &" "  	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		'Response.Write "<p>sQuery="&sQuery
		.Open
	end with

	set objRs.ActiveConnection = nothing
	IF Not objRs.EOF then
		sParName		= objRs(0)														
		sParSubTypeName = objRs(1)			
		sParTemp = sParType &"-"& sParSubTypeName											
	End IF 
	objrs.Close 
End IF
iPartyVal = sParType&"?"&iParSubType&"?"&sParTemp&"?"&iParCode
'Response.Write iPartyVal
'------------------------------------------------------------------------------------------------------------------                                                    
sQuery="select ReasonForPayment,AmountToPay,ToBePaidTo,PayablesNumber from Acc_T_PaymentRequestDet where PaymentRequestNo="&iPymtNo
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
		sPayTo=objRs(2)		
		iAmount=FormatNumber(objRs(1),2,,,0)
		
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=edge">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="VoucherData">
	<voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="0" BookName="0" CRDR="C" VouDate="" BookAcchead="" Approver=""/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="EntryData">
	<Entry No="0" CRDR="D" Payto="" Amount="" AccUnit="" AccName=""/></script>
<script type="application/xml" data-itms-xml-island="1" id="RequestData">
	<RequestDetails/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
	<account/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempXMLData"><Root></Root></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script src="../../scripts/VouTransactions.js"></script>
<SCRIPT SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script src="../../scripts/checkdate.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/PmtGenRegularChqEntryForPartyCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="InitPmtGenRegularChqEntryForParty()">
<form method="POST" name="formname" action="PmtGenerate.asp">
<input type="hidden" name="hUnitId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hPaymentNo" value="<%=iPymtNo%>">
<input type="hidden" name="hParSubName" value="<%=sParSubTypeName%>">
<input type="hidden" name="hParName" value="<%=sParName%>">
<input type="hidden" name="hVouType" value="<%=Left(sParType,1)%>">
<input type="hidden" name="hAmount" value="<%=iAmount%>">
<input type="hidden" name="hInsDet" value="">
<input type="hidden" name="hParCode" value="<%=iParCode%>">
<input type="hidden" name="hParValue" Value="<%=iPartyVal%>">
<input type="hidden" name="hTransNo" Value="<%=nTransactionNo%>">
<%if sReqType="A" then%>
<input type="hidden" name="hVouCode" value="02">
<input type="hidden" name="hVouName" value="BA">
<%else%>
<input type="hidden" name="hVouCode" value="01">
<input type="hidden" name="hVouName" value="CA">
<%end if%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">&nbsp;Regular Payment -
		<%if sReqType="A" then
			Response.Write "CHEQUE"
		else
			Response.Write "CASH"
		end if
		%>
		Generation for <%=sOrgname%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0" width="90%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="110"><p align="center">Request Details
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
															<td class=MiddlePack colspan="3"> </td>
														</tr>
                                                        <tr>
								<td align="center">&nbsp;	</td>
								<td valign="top">
                                                   <table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCell width="78"> Raised
                                                              By</td>
															<td class="FieldCell">
                                                            <span class="DataOnly"><%=sRequestBy%>&nbsp;</span></td>
															<td width="20" class="FieldCell">
                                                            </td>
															<td width="80" class="FieldCell">
                                                            Approved By</td>
															<td class="FieldCell">
                                                            <span class="DataOnly"><%=sApprovedBy%>&nbsp;</span></td>
														</tr>
														<tr>
															<td class=FieldCell width="78"> Pay To</td>
															<td class="FieldCell"><span class="DataOnly"><%=Trim(sPayTo)%>&nbsp;</span></td>
														</tr>

														<tr>
															<td class=MiddlePack colspan="5"> </td>
														</tr>

														<tr>
															<td class=FieldCell width="178" colspan="5">
															<DIV class=frmBody id=frm1 style="width: 385; height:50;">

                                    <table border="0" id="tblPayable0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" >Reason</td>
                                        <td class="ExcelHeaderCell" align="center" >Amount To Pay</td>
                                            </tr>
                                            <%
												iSno=1
												objRs.MoveFirst
                                        	do while not objRs.EOF
                                        %>
                                        	    <tr>
													<td class="ExcelSerial" align="center" width="30"><%=iSno%></td>
													<td class="ExcelDisplayCell" align="Left"><%=objRs(0)%></td>
													<td class="ExcelDisplayCell" align="right"><%=FormatNumber(objRs(1),2,,,0)%></td>
                                            </tr>
                                        <%
												iSno=cint(iSno)+1
												objRs.MoveNext
											loop
											objRs.Close
										%>

                                                </table>
                                                												</div>

 </td>
														</tr>

													</table>
								</td>
								<td align="center">&nbsp;</td>
                                                        </tr>
												</center>
														<tr>
															<td class=MiddlePack colspan="3">
                                                   </td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
                                                        </div>
								</td>
								<td align="center">
								</td>
                              </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                     <table border="0" cellspacing="0" cellpadding="0">
                                       <tr>
											<td class="FieldCellSub"> <Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()">&nbsp;</td >
                                       </tr>
                                       <tr>
											<td class="FieldCellSub" width="139">Inst No</td>
											<td>                 
												 <span id="spInsNo" class="DataOnly">-</span>
											</td >&nbsp;
                                            <td class="FieldCellSub" width="139">Inst Date</td>
											<td width="296">
												<span id="spInsDate" class="DataOnly">-</span>
                                            </td>
                                        </tr>
                                        <tr>
                                           <td class="FieldCellSub" width="139">Select Book</td>
                                           <td width="296">
                                               <select size="1" name="selBookId" class="FormElem">
                                                    <option value="Select">Select</option>
														<%
															if sReqType="A" then
																sQuery="select BookNumber,BookName,BookAccountHead from vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode=02"
															else
																sQuery="select BookNumber,BookName,BookAccountHead from vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode=01"
															end if
															with objRs1
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs1.ActiveConnection = nothing

														do while not objRs1.EOF
															Response.Write    "<option value="""&trim(objRs1(0))&"?"&trim(objRs1(2))&""">"&trim(objRs1(1))&"</option>"
															objRs1.MoveNext
														loop
														objRs1.Close
														%>
												</select>
											</td>
                                            <!--<td class="FieldCell" colspan="2" valign="bottom"><p align="center">Payment
                                             </td>-->
                                         </tr>
                                                         
                                          <tr>
                                                <td class="FieldCellSub" width="139">Party Name</td>
                                                 <!--select size="1" name="selAccType" class="FormElem" onChange="selAccountHead(this)"-->
                                                 <td class="FieldCell" width="296">
													<span id="ParNamID" class="DataOnly"><%If sParSubTypeName <> "" Then Response.Write sParSubTypeName Else Response.Write "-" End IF%></span>                                                    
										         </td>
                                                 <td class="FieldCell">Date</td>
                                                    <td class="FieldCell"> <p align="center">
													<% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>
											</tr>
											<tr>
                                                 <td class="FieldCellSub" width="139">Pay to </td>
                                                 <td class="FieldCell" colspan="3"> <span id="ParNamID" class="DataOnly"><%=sParName%></span>  </td>
                                            </tr>
                                            <tr>
                                                 <td class="FieldCellSub" width="139" valign="top">Narration</td>
                                                 <td class="FieldCell" colspan="3"> <textarea rows="3" name="txtNarration" cols="70" class="FormElem"></textarea> </td>
                                            </tr>
                                            <tr>
                                                 <td class="FieldCellSub" width="139">Amount</td>
                                                 <td class="FieldCell" colspan="3"> <input type="text" name="txtAmount" value="<%=iAmount%>" size="15" style="text-align:right" class="Formelem"> </td>
                                            </tr>
                                     </table>
								</td>
								<td align="center">
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
	<DIV class=frmBody id="DisPayable" style="width: 555; visibility: hidden; height:1;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="555">
			<tr>
				<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
				<td class="ExcelHeaderCell" align="center" width="300" colspan="3">Document</td>
				<td class="ExcelHeaderCell" align="center" width="250" colspan="3">Amount</td>
		    </tr>
		   <tr>
				<td class="ExcelHeaderCell" align="center">Number</td>
				<td class="ExcelHeaderCell" align="center">Date</td>
				<td class="ExcelHeaderCell" align="center">Type</td>
				<td class="ExcelHeaderCell" align="center">Amount</td>
				<td class="ExcelHeaderCell" align="center">Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To adjust</td>
		   </tr>
		   <tr>
				<td class="ExcelSerial" align="center">1</td>
				<td class="ExcelDisplayCell">xNumber</td>
				<td class="ExcelDisplayCell" align="right"><p align="left">xDate</td>
				<td class="ExcelDisplayCell" align="right"><p align="left">xType</td>
				<td class="ExcelDisplayCell" align="right">xAmount
				</td>
				<td class="ExcelDisplayCell" align="right"><p align="right">xAdjusted
				</td>
				<td class="ExcelDisplayCell" align="right">xToAdjust</td>
			</tr>

		</table>
	</div>
</div><!--End of Addtional Details Display  -->
						</td>
								<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                                                </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
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
												<input type="button" value="Create" name="B4" class="ActionButton" onClick="AddEntry()">
												 <input type="reset" value="Reset" name="B1" class="ActionButton" >
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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
