<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TDSGroupSelectionCash.ASP
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MAHESHWARI S
	'Created On					:	Feb 26,2007
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	BankVoucher.asp
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

<%
Dim Root,docobj,sCallFrom,sVouName,sNewVal,sUpdate,iNewGrpId,iChkId
set docobj=server.CreateObject("Microsoft.XMLDOM")
set Root=docobj.DocumentElement

set Root=docobj.CreateElement("Root")
docobj.appendchild Root
sCallFrom = Request("CallFrom")
sVouName = Request("VouName")
sNewVal = Request("NewVal")
iNewGrpId = Request("GrpId")
sUpdate = Request("Update")

'iChkId = Request.Form("hGrpId")
'Response.Write "sVouName="&sVouName
' Response.Write "sNewVal="&sNewVal &"<BR>"
'Response.Write "sCallFrom="&sCallFrom&"<BR>"
'Response.Write "sUpdate="&sUpdate&"<BR>"
'Response.Write "GrpId="&iNewGrpId&"<BR>"
'Response.Write Session.SessionID
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>TDS</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/DataValidation.js"></SCRIPT>

<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" id="TempData" data-src="<%="../temp/transaction/TDS_Cash_"&Session.SessionID&".xml"%>"  ><Root/>
</script>
<% 'Response.Write sCallFrom &"    "& sUpdate
	If sCallFrom = "A" then
		If trim(sUpdate) = "Y" then %>
			<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/NewVoucher_"&sVouName&"_"&Session.SessionID&".xml"%>"  ><Root/></script>
	<%	Else %>
			<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml"%>"  ><Root/></script>
	<% End If  %>
<% ElseIf sCallFrom ="C" then %>
	<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml"%>"  ><Root/></script>
<% End If  %>
<script type="application/xml" data-itms-xml-island="1" id="NewData" data-src="<%="../temp/transaction/NewVoucher_"&sVouName&"_"&Session.SessionID&".xml"%>"  ><Root/></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/TDSGroupSelectionCashCompat.js"></script>
</HEAD>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hCompForm" value="<%=sCompForm%>">
<input type="hidden" name="hPer" value="">
<input type="hidden" name="hCheck" value="">
<input type="hidden" name="hCtr" value="">
<input type="hidden" name="hReCalc" value="">
<input type="hidden" name="hNewVal" value="<%=sNewVal%>">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hNewGrpId" value="<%=iNewGrpId%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		TDS Calculation Details </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell" width="100%">

								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
                        <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                        </tr>
                        <tr>
							<td align="center" class="ClearPixel">
							</td>
							<td valign="top" class="FieldCell" width="100%" align="center">
                            <DIV class=frmBody id=frm3 style="width: 360; height:265;">

                                        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                        <tr>
                                        <td class="ExcelHeaderCell" align="center" width="30" colspan="2">S.No.</td>
                                        <td class="ExcelHeaderCell" align="middle" width="300" colspan="8">Head Name/Account Head
										</td>
										<td class="ExcelHeaderCell" align="middle" width="100" colspan="2"> %
										</td>
										<td class="ExcelHeaderCell" align="middle" width="100" colspan="4">Value
										</td>
										<td class="ExcelHeaderCell" align="middle" width="40" colspan="2">Round <br>
										Off
										</td>
										</tr>

											<%	'Query
												Dim objFS,sQry,rs,sTDSAmt,sGrpId,sGrpName,sCompMode,sAcHeadCode,sCompForm,sEntNo,iGrpId,sNewAmt
												Dim sTemp,Arr,Arr1,sVal,Value,iCtr,n,i,j,sTemp1,sArrVal,sTempVal,sVal1,sVal2,sTotValue
												Dim iVal1,iVal2,iVal3,iVal4,newElem,RootElem,subnode,sNewPer,sNewCtr,sArrGrp,ArrGrpId,iGrpHeadId
												Dim TDSnode,TFormula,Tamt,iCnt,sArrTemp,sArrTemp1,sArrTempVal,k,sPer,sNewArr1,sNewArr2
													Set objFs = server.CreateObject("Scripting.FileSystemObject")
													Set rs = server.CreateObject("ADODB.Recordset")
													sEntNo = Request("EntNo")
													sTDSAmt = Request("Amount")
													iGrpId = Request("GrpId")
													sNewAmt = Request("NewAmt")
													'Response.Write "iGrpId="& iGrpId &"<BR><BR>"
													If trim(sTDSAmt) <> "" then
													sQry = " Select GroupHeadID,GroupHeadName,ComputeMode,AcHeadCode,isNull(ComputeFormula,'') "&_
														   " from ACC_M_TDSHeadComputation where GroupId = "&iGrpId&" "
														   'Response.Write sQry
													with rs
														.CursorLocation = 3
														.CursorType = 3
														.ActiveConnection = con
														.Source = sQry
														.Open
													end with
													iCtr = 1

													Do while not rs.EOF
														sGrpId = rs(0)
														sGrpName = rs(1)
														sCompMode = rs(2)
														sAcHeadCode = rs(3)
														sCompForm = rs(4)
														iGrpHeadId = iGrpId &"-"& sGrpId
														sTemp1 = ""
														 'Response.Write "sCompForm=" &sCompForm
														'Response.Write "sCompMode="& sCompMode
														If trim(sCompForm) = "" then
															sTemp1 = "0"
															Value = "0"
														Else
															IF trim(sCompMode) <> "F" then
																Arr = split(sCompForm,",")
																sTemp =Arr(0)
																IF sTemp <> "0" then
																	Arr1 = split(sTemp,"-")
																	sTemp1 = Arr1(1)
																End If
															Else
																sTemp1 = sCompForm
															End If	'IF trim(sCompMode) <> "F" then
														End If
														Set newElem = docobj.CreateElement("TDS")
														newElem.SetAttribute "Ctr",iCtr
														newElem.SetAttribute "AccHeadCode",sAcHeadCode
														newElem.SetAttribute "TDSAmount",Round(sTDSAmt,2)
														newElem.SetAttribute "TdsPercentage",sTemp1
														newElem.SetAttribute "PayRecAmount","0"
														newElem.SetAttribute "Formula",sCompForm
														newElem.SetAttribute "CompMode",sCompMode
														newElem.SetAttribute "GroupHeadId",iGrpHeadId
														newElem.setAttribute "TdsRndOff","N"
														Root.appendchild newElem


														'If trim(sCompMode)= "F" then
														'	sCompMode = ""
														'	sCompForm = ""
														'	Value = "0"
														'	sTemp1 = "0"
														'End If
														If trim(sCompForm) = "" then
															sTemp1 = "0"
															Value = "0"
														End If
														If trim(sCompMode) = "P" then
															sCompMode = "P"

															Value=CalculatePer(Root,sCompForm,sTDSAmt)

																For each TDSnode in Root.childnodes
																	If trim(iCtr) = TDSnode.getAttribute("Ctr") then
																		TDSnode.setAttribute "PayRecAmount",Value
																		TDSnode.setAttribute "TdsRndOff","N"
																	End If
																next

														'Response.Write "Value="&Value

														End If 'If trim(sCompMode) = "P" then
														'Response.Write Value

														sTotValue = sTotValue + Value
														'Response.Write sTotValue



														 %>
														<TR>
															<td class="ExcelSerial" align="center" colspan="2" width="30"><%=iCtr%></td>
															<td class="ExcelDisplayCell" colspan="8"  width="300"><%= sGrpName%>
															</td>
															<%IF sCompMode = "P" Then %>
																<td class="ExcelInputCell" colspan="2"><p align="right">
					 												<input type="text" name="txtPer<%=iCtr%>" value="<%=FormatNumber(sTemp1,2,,,0)%>" size="5" OnChange="TDScalc('<%=iCtr%>')" maxlength="5" class="Formelem" style="text-align:Right" >

					 												</p>
																</td>
															<%Else%>
																<td class="ExcelDisplayCell" colspan="2"><p align="right">
					 												<input type="text" name="txtPer<%=iCtr%>" value="" size="5" maxlength="5" class="FormelemRead" style="text-align:Right" readonly>

					 												</p>
																</td>
															<%End IF %>
															<%IF sCompMode = "P" Then %>
																<td class="ExcelDisplayCell" align="Right" colspan="4"  width="100">
																<Input Type="text" name="value<%=iCtr%>" Value="<%=FormatNumber(Value,2,,,0)%>" class="FormelemRead"  onchange="DisplayTot('<%=iCtr%>')" style="text-align:Right">
																</td>
															<%Else%>
																<td class="ExcelInputCell" align="Right" colspan="4"  width="100">
																	<Input Type="text" name="value<%=iCtr%>" Value="<%=FormatNumber(Value,2,,,0)%>" class="Formelem"  onchange="DisplayTot('<%=iCtr%>')" style="text-align:Right">
																</td>
															<%End IF %>
															<td class="ExcelDisplayCell" colspan="2" width="40"><p align="center">
					 												<input type="checkbox" name="ChkRnd<%=iCtr%>" Value="" size="5"  maxlength="5" class="Formelem" style="text-align:center" onclick="SetRoundOff('<%=iCtr%>')" >
					 												</p>
																</td>
														</TR>
														<input type="hidden" name="hRndOffVal<%=iCtr%>" value = "N">
												<%

													docobj.save server.MapPath("../temp/transaction/TDS_Cash_"&Session.SessionID&".xml")
												 iCtr = iCtr + 1
												rs.MoveNext
													loop
												rs.Close

												%>
												<tr>
													<td class="ExcelSerial" align="right" colspan="12"><b>Total</b></td>
													<td class="ExcelDisplayCell" align="center">
														<Input Type="text" name="txtTdsTotal" Value="<%=FormatNumber(sTotValue,2,,,0)%>" class="Formelemread" style="text-align:Right"  >
													</td>

												</tr>
										<input type="hidden" name="hTotVal" value="<%=sTotValue%>">
										<input type="hidden" name="hTDSAmt" value="<%=sTDSAmt%>">
										<input type="hidden" name="hNewAmt" value="<%=sNewAmt%>">
										<input type="hidden" name="htxtPer" value ="<%=iCtr-1%>">

									<%  End If 'If trim(sTDSAmt) <> "" then

									%>
									<input type="hidden" name="hGrpId" value="<%=iGrpId%>">
							</div>
							</td>
						</tr>


					</table>
					</td>
				</tr>
				<tr>
								<td align="center" class="ClearPixel">
								</td>
                        </tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell">
								<table border="0" cellpadding="0" cellspacing="0" width="100%" >
									<tr>
										<td valign="middle" class="ActionCell" align="center">

											<input type="button" value="Save" onclick="Window_onunload()" name="B1" class="ActionButton" tabindex="3"> &nbsp;&nbsp;&nbsp;
											<input type="button" value="Close" onClick="Window_onunload()" name="B2" class="ActionButton" tabindex="3">

										</td>
									</tr>
								</table>
								</td>
								<td align="center" class="ClearPixel">
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
</form>
</BODY>
</HTML>
