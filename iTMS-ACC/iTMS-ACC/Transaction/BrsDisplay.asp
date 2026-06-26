<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BrsDisplay.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 27,2003
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->


<%

'XML DOM Variables
Dim oDOM,nodHeader,Root,sQuery,objrs
dim sBookName,dPassBalance,sPassCrDr,dBookBalance,sBookCrDr,sFromDt,sToDt
Dim sRecVal,sAccName,iBookNo,sQry,iAccHead,sOrgId
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml")

set Root=oDOM.documentElement
Set objrs = Server.CreateObject("ADODB.RecordSet")
sOrgId=root.Attributes.Item(0).nodeValue
iBookNo = root.Attributes.Item(1).nodeValue
sBookName= root.Attributes.Item(2).nodeValue
sFromDt= root.Attributes.Item(3).nodeValue
sToDt= root.Attributes.Item(4).nodeValue
dPassBalance= root.Attributes.Item(5).nodeValue
sPassCrDr= root.Attributes.Item(6).nodeValue
dBookBalance= root.Attributes.Item(7).nodeValue
sBookCrDr= root.Attributes.Item(8).nodeValue

sRecVal = Request("hChkVal")
'Response.Write "sChkVal="&sChkVal
sQuery="select AccountHead,AccountDescription from Acc_M_GLAccountHead where AccountHead=(select BankChargesHead from Acc_M_BankDetails where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='02' and BookNumber="&iBookNo&")"
'Response.Write sQuery
with objRs
.CursorLocation = 3
.CursorType = 3
.Source = sQuery
.ActiveConnection = con
.Open
end with
set objRs.ActiveConnection = nothing

if not objRs.EOF then
	iAccHead=objRs(0)
	sAccName=objRs(1)
end if
objrs.close
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>BRS Display</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<xml id="OutData" src="<%="../temp/transaction/Bank Recon_BA_"&Session.SessionID&".xml"%>">
</xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="vbscript" >
Function BankCharges()
	IF document.formname.C1.checked = true then
		document.formname.hChkVal.value = True

		document.formname.ctlDate.Enable = True
		document.formname.selCRDR(0).disabled  = False
		document.formname.selCRDR(1).disabled  = False
		document.formname.txtNarration.disabled  = False
		document.formname.txtAmount.disabled  = False


	Else
		document.formname.hChkVal.value = False

		document.formname.ctlDate.Enable = False
		document.formname.selCRDR(0).disabled  = True
		document.formname.selCRDR(1).disabled  = True
		document.formname.txtNarration.disabled  = True
		document.formname.txtAmount.disabled  = True

		document.formname.hVouDate.value = ""
		document.formname.hAmt.value = ""
		'document.formname.hCrDr.value = document.formname.selCRDR(0).value
		document.formname.hCrDr.value = ""

	End IF
End Function
Function BrsPopup()
Set Root = OutData.documentelement
'alert(Root.xml)
sInsDet =  document.formname.hInsDet.value
'alert sInsDet
IF document.formname.C1.checked <> True then
	document.formname.hChkVal.value = False

	Exit function
Else
	document.formname.hChkVal.value = True
	set OutValue= showModalDialog("BrsCommEntry.asp?InsDet="&sInsDet,"","dialogHeight:350px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
	  'alert(OutValue.xml)
	 IF OutValue.xml = "" then exit function
	 IF OutValue.xml <> "" then
		For each node in OutValue.childnodes
			dVouDate = OutValue.getAttribute("VouDate")
			document.formname.hVouDate.value = dVouDate
			IF trim(node.NodeName) = "Entry" then
				sCrDr = node.getAttribute("CRDR")
				iAmt  = node.getAttribute("Amount")
				For each AccNode in node.childnodes
					IF trim(AccNode.NodeName) = "AccHead" then
						iAccHead = AccNode.getAttribute("No")
						document.formname.hAccNo.value = iAccHead
					End IF
					IF trim(AccNode.NodeName) = "Narration" then
						document.formname.hNarr.value = AccNode.text
					End IF
				Next

				document.formname.hAmt.value = iAmt
				document.formname.hCrDr.value = sCrDr
			End If
		Next
	 End IF
'	alert document.formname.hVouDate.value &"--"&	document.formname.hAmt.value  &"--"&document.formname.hCrDr.value

End If
End Function
Function finalDone(bFlag)
	IF trim(bFlag) ="P" then
		IF document.formname.C1.checked <> True then
			RetVal = MsgBox("Bank Charges Not Applicable.Do U want to Continue?",4)
			'alert(RetVal)
			IF RetVal <> 6 then Exit Function

		End IF
		IF document.formname.C1.checked = True then
			document.formname.hVouDate.value = document.formname.ctlDate.GetDate
			IF document.formname.txtAmount.value = "" then
				Alert("Enter Amount")
				Exit Function
			End IF
			document.formname.hAmt.value = document.formname.txtAmount.value

			IF document.formname.selCRDR(0).checked = true then
				document.formname.hCrDr.value = document.formname.selCRDR(0).value
			Else
				document.formname.hCrDr.value = document.formname.selCRDR(1).value
			End If
		End IF
		'alert(document.formname.hCrDr.value)

		document.formname.action="BrsGenerate.asp"
		document.formname.submit()
	End If
End Function
Function CheckRecon()
	document.formname.ctlDate.Enable  = False
	'alert document.formname.hRecVal.value
	IF trim(document.formname.hRecVal.value) = "BR" then
		'document.formname.C1.checked = True
		'BrsPopup()
	End IF
End Function
Function SelChk()
	document.formname.txtNarration.value = ""
	For i = 1 to document.formname.hCtr.value -1
		Set sObj = eval("document.formname.Chk"&i)
		IF trim(sObj.checked) = "True" then
			document.formname.txtNarration.value = document.formname.txtNarration.value &","&sObj.Value

		End IF
	Next
	document.formname.txtNarration.value = mid(document.formname.txtNarration.value,2)
End Function
</script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="CheckRecon()">

<form method="POST" name="formname">
<input type="hidden" name="hRecVal" Value="<%=sRecVal%>">
<input type="hidden" name="hChkVal" value="False">
<input type="hidden" name="hVouDate" value="">
<input type="hidden" name="hAccNo" value="<%=iAccHead%>">
<input type="hidden" name="hAmt" value="">
<input type="hidden" name="hCrDr" value="">
<input type="hidden" name="hNarr" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Bank Reconciliation</td>
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
											<td align="center">Instruments
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
                                    <tr>
                                    	<td align="center">Reconciled List</td>
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
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
												<DIV class=frmBody id=frm1 style="width: 585; height:140;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"></td>
                                        <td class="ExcelHeaderCell" align="center">Instrument No. - Date</td>
                                        <td class="ExcelHeaderCell" align="center" width="150">Paid/Recd </td>
                                        <td class="ExcelHeaderCell" align="center" width="80">Receipt</td>
                                        <td class="ExcelHeaderCell" align="center" width="80">Payment</td>
                                        <td class="ExcelHeaderCell" align="center" width="50">Date</td>
                                            </tr>
<%
dim sVouNo,sInstDet,dAmont,sTransType,sClearedOn,sPayRec,iVouNo,iInsDet
dim dPayTotal,dRecTotal,iSNo,dTotal,sExp,tempNode,iCounter,iTransNo,iCtr
iCtr=1

sExp="//Voucher"
set tempNode=Root.selectNodes(sExp)
for iCounter=0 to tempNode.length-1
	iSNo = tempNode.item(iCounter).Attributes.Item(0).nodeValue
	iTransNo = tempNode.item(iCounter).Attributes.Item(1).nodeValue
	sVouNo=tempNode.item(iCounter).Attributes.Item(3).nodeValue
	sInstDet=tempNode.item(iCounter).Attributes.Item(4).nodeValue
	sPayRec=tempNode.item(iCounter).Attributes.Item(5).nodeValue
	dAmont=tempNode.item(iCounter).Attributes.Item(6).nodeValue
	sTransType=tempNode.item(iCounter).Attributes.Item(7).nodeValue
	sClearedOn=tempNode.item(iCounter).Attributes.Item(8).nodeValue

	if tempNode.item(iCounter).Attributes.Item(2).nodeValue="Y" then

	'Response.Write sVouNo & "<br>"
	iVouNo = iVouNo &","& sVouNo
	iInsDet = iInsDet &","& sInstDet
%>
                                            <tr>
                                        <td class="ExcelSerial" align="center"><%=iCtr%></td>
                                        <td class="ExcelDisplayCell" align="right"><p align="left">
                                        <input type="Checkbox" name="Chk<%=iCtr%>" value="<%=sInstDet%>" onclick="SelChk()" checked >
                                        </td>
                                        <td class="ExcelDisplayCell" align="right"><p align="left"><%=sInstDet%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=sPayRec%></td>
                                        <td class="ExcelDisplayCell" align="right">
                                        <%if sTransType="R" then Response.Write FormatNumber(dAmont,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right">
                                        <%if sTransType="P" then Response.Write FormatNumber(dAmont,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=sClearedOn%></td>
                                            </tr>
<%	iCtr=CInt(iCtr)+1

	else

		if sTransType="P" then
			dPayTotal=CDbl(dPayTotal)+ CDbl(dAmont)
		else
			dRecTotal=CDbl(dRecTotal)+ CDbl(dAmont)
		end if

	end if
	'iCtr=CInt(iCtr)+1
next
iVouNo = mid(iVouNo,2)
iInsDet = mid(iInsDet,2)
%>
   <input type="hidden" name="hVouNo" value="<%=iVouNo%>">
   <input type="hidden" name="hInsDet" value="<%=iInsDet%>">
   <input type="hidden" name="hTransNo" value="<%=iTransNo%>">
   <input type="hidden" name="hCtr" value="<%=iCtr%>">
                                                </table>
												</div>
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
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly">
                                <tr>
									<td class="MiddlePack" colspan="3"> </td>
                                </tr>
                                <tr>
                                <!--td class="FieldCellSub" >Bank Charges Applicable</td>
								<td class="FieldCell"><input type="checkbox" name="ChkBanChrg" class="formelem" value="" onclick="BrsPopup()"></td-->
                                </tr>
									<tr>
												<td class="FieldCell">Book Balance
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=formatnumber(dBookBalance,2,,,0)%>&nbsp;<%=sBookCrDr%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Passbook Balance as on <%=sToDt%>
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=formatnumber(dPassBalance,2,,,0)%>&nbsp;<%=sPassCrDr%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Cheque / DD deposited but yet to be cleared
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=FormatNumber(dRecTotal,2,,,0)%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Cheque / DD issued but yet to be cleared
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=FormatNumber(dPayTotal,2,,,0)%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell" valign="top">Reason
												</td>
												<td class="FieldCellSub"><textarea rows="2" name="S1" cols="40" class="FormElem"></textarea>
												</td>
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

								<tr>
									<td align="center" width="5" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%"><center>
										<div align="left">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td>
														<table cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td class="GroupTitleLeft" width="10">&nbsp;
																</td>
																<td class="GroupTitle" width="120">
																	<p align="center">
																	<input type="checkbox" name="C1" value="ON" class="FormElem" onclick="BankCharges()">
 																	Bank Charges
																</td>
															</center><td class="GroupTitleRight">
																<p align="left">&nbsp;
															</td>
														</tr>

													</table>
												</td>
											</tr>

											<tr>
												<td class="GroupTable"><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="MiddlePack" colspan="3">
															</td>
														</tr>

														<tr>
															<td class="ClearPixel" width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
															<td class="FieldCell" width="100%"></center>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td class="FieldCell">Book Name
																	</td>
																	<td class="FieldCellSub">
																		<span class="DataOnly"><%=sBookName%>&nbsp;</span>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell">Charges Account Head
																	</td>
																	<td class="FieldCellSub">
																		<span class="DataOnly"><%=sAccName%>&nbsp;</span>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell">Charge Type
																	</td>
																	<td class="FieldCellSub">
																		<input type="radio" value="C" checked name="selCRDR" class="FormElem" disabled>
 																		Receipts
																		<input type="radio" name="selCRDR" value="D" class="FormElem" disabled>
 																		Payments
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell">Charges Date
																	</td>
																	<td class="FieldCellSub">
																	 <% ' Function Call to Insert Date Picker
																		'	Response.Write InsertDatePicker("ctlDate")
																		%>
																		<object id="ctlDate" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext >
																			<param name="_ExtentX" value="2355">
																			<param name="_ExtentY" value="529">
																		</object>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell" valign="top">Narration
																	</td>
																	<td class="FieldCellSub"><textarea rows="2" name="txtNarration" cols="60" class="FormElem" Disabled ><%=iInsDet%></textarea>
																	</td>
																</tr>

																<tr>
																	<td class="FieldCell" valign="top">Amount
																	</td>
																	<td class="FieldCellSub">
																		<input type="text" name="txtAmount" size="10" class="FormElem" disabled>
																	</td>
																</tr>

															</table>
														</td>
														<td class="ClearPixel" width="5">
															<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
														</td>
													</tr>

													<tr>
														<td class="MiddlePack" width="267" colspan="3">
														</td>
													</tr>

													</table>
												</td>
											</tr>

										</table>
										</div>
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

												<td valign="middle" class="ActionCell">
													<p align="center">
                                                        <!--input type="button" value="Save & Print" onClick="finalDone('P')" name="B2" class="ActionButtonX">
                                                        <input type="button" value="Done" name="B6" onClick="finalDone('S')" class="ActionButton"-->
                                                        <% IF trim(sRecVal) = "BR" then %>
															<input type="button" value="Save" onClick="finalDone('P')" name="B2" class="ActionButtonX">
														<% Else %>
															<input type="button" value="Save" onClick="finalDone('P')" name="B2" class="ActionButtonX">
                                                        <% End IF %>
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
</form>
</BODY>
</HTML>