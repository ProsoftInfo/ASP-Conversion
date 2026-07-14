<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BankInsDetails.asp
	'Module Name				:	Fixed Deposit(Transaction)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Jun 15,2005
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<%
	Dim oDOM,Root,node,sTemp,sTemparr,sDHname,sJDname,sNDname,sGDname,iAmt,sVouTy,sVouDate,sBookNo,sOrgId,Elem
	Dim sQry,rs,sPrintCheq,dtIssue,sDrawOn,sPayAt,iEntNo,sFlag,sVouType,iTransNo,iUsInsNo,sVouName
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")


		sTemp = Trim(Request("sTemp"))
		'Response.Write "sTemp="&sTemp
		
		If 1 = 2 Then
		sVouTy = Right(sTemp,1)
		sTemparr = Split(sTemp,":")
		sVouDate = sTemparr(2)
		sBookNo  = sTemparr(3)
		sOrgId	 = sTemparr(4)
		sVouType = sTemparr(5)
		iTransNo = sTemparr(6)
		'sTemparr = Split(sTemp,":")
		'sDHname = sTemparr(0)
		'iAmt = sTemparr(1)

		sVouName = sTemparr(7)
		End IF
		
		
		sTemparr = Split(sTemp,":")
		sVouTy   = Right(sTemparr(0),1)
		sVouDate = sTemparr(1)
		sBookNo  = sTemparr(2)
		sOrgId	 = sTemparr(3)
		sVouType = sTemparr(0)
		iTransNo = sTemparr(4)
		sVouName = sTemparr(5)
		
		iUsInsNo = ""

		sFlag = False

	If trim(iTransNo) <> "0" then
		'oDOM.Load server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml")
		oDOM.Load server.MapPath("../temp/transaction/"&iTransNo&".xml")

		set Root=oDOM.documentElement
		IF Root.haschildnodes then
			For each node in Root.childnodes
				IF trim(node.NodeName) = "BankInstrumentDet" then
					Root.Removechild node
				End If
			Next
		End IF
	End IF ' If trim(iTransNo) <> "0" then
'	Response.Write "iTransNo="&sVouType

	IF trim(sVouType) = "P" then
		sQry = "Select PrintCheques from Acc_M_BankDetails where OUDefinitionID = '"&sOrgId&"' and BookNumber ="&sBookNo
		'Response.Write sQry
		rs.Open sQry,con
		IF not rs.EOF then
			sPrintCheq = rs(0)
		End If
		rs.Close
		' Response.Write "sPrintCheq="&sPrintCheq
		IF trim(sPrintCheq) = "1" then
			sQry = "Select convert(varchar,dateOfIssue,103),DrawnOn,PayableAt,EntryNo from Acc_R_BankInstrumentDetails where OUDefinitionID = '"&sOrgId&"' and BookNumber = "&sBookNo
			'Response.Write sQry
			rs.Open sQry,con
			IF not rs.EOF then
				sFlag = True
				dtIssue = rs(0)
				sDrawOn = rs(1)
				sPayAt  = rs(2)
				iEntNo  = rs(3)
			End If
			rs.Close
			'sQry = "Select EntryNo,InstrumentEntryNo,InstrumentNo from Acc_R_BankInstrumentUsage where CreatedTransNo = "&iTransNo&" and Status = 'U'"

			sQry = "Select I.InstrumentEntryNo,I.BankInstrumentEntryNo,U.EntryNo,U.InstrumentEntryNo,U.InstrumentNo,I.BankInstrumentType,convert(VarChar,I.BankInstrumentDate,103),"&_
				   "I.PayableAt,I.DrawnOnBank,I.InstrumentAmount from Acc_R_BankInstrumentUsage as U,Acc_T_CreatedVoucherInstrumentDet as I where U.CreatedTransNo = "&iTransNo&" and "&_
				   "I.CreatedTransNo = U.CreatedTransNo and I.BankInstrumentEntryNo = U.EntryNo and Isnull(I.InstrumentEntryNo1,0) = U.InstrumentEntryNo and U.Status = 'U'"

			 'Response.Write sQry

			 'Response.End
			rs.Open sQry,con
			IF not rs.EOF then

				Do while not rs.EOF
					Set Elem = oDOM.CreateElement("BankInstrumentDet")
					Elem.SetAttribute "SlNo",rs(0)

					IF trim(rs(1)) = "0" then
						Elem.SetAttribute "InsNo",rs(4)
					Else
						iUsInsNo = rs(2)&"-"&rs(3)&"-"&rs(4)
						Elem.SetAttribute "InsNo",iUsInsNo
					End IF
					Elem.SetAttribute "InsType",rs(5)
					Elem.SetAttribute "InsDate",rs(6)
					Elem.SetAttribute "PayAt",rs(7)
					Elem.SetAttribute "DrawnOn",rs(8)
					Elem.SetAttribute "InsAmt",rs(9)
					Elem.setAttribute "Option","Y"
					Elem.setAttribute "Action","0"
					Root.appendchild Elem
					rs.MoveNext
				loop
			End If
			rs.Close
			' Response.Write iUsInsNo
		End If

	End IF ' IF trim(sVouType) = "P" then
oDOM.Save Server.MapPath ("../temp/transaction/Voucher Amd_"&sVouName&"_"&Session.SessionID&".xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Bank Voucher - Instrument Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" ID="NewData">
<Root>
</Root>
</script>
<!--XML ID="OutData" src="<%="../temp/transaction/Voucher Amd_"&sVouName&"_"&Session.SessionID&".xml"%>"-->
<script type="application/xml" data-itms-xml-island="1" ID="OutData">

<!--BankInstrumentDet>
</BankInstrumentDet-->

</script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script src="../../scripts/BankInsDetailsCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">

<form method="POST" name="formname" action="">
<Input Type="hidden" name="hExists" Value="">
<Input Type="hidden" name="hEditNo" Value="">

<Input Type="hidden" name="hTransNo" value="<%=iTransNo%>">
<Input Type="hidden" name="hUsInsNo" value="<%=iUsInsNo%>">
<Input Type="hidden" name="hVouDate" Value="<%=sVouDate%>">
<input type="hidden" name="hDate" value="<%=dtIssue%>">
<input type="hidden" name="hFlag" value = "<%=sFlag%>">
<input type="hidden" name="hCtr" value="1">
<input type="hidden" name="hInsType" value="C">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Instrument Details
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0">
                                                        <!--tr>
															<td class=FieldCell> Name of First Deposit Holder</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">

                                                            </span>
                                                            </td>
                                                        </tr>

                                                        <tr>
															<td class=FieldCell> Name of Second Deposit Holder</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">

                                                            </span>
                                                            </td>
                                                        </tr>

                                                        <tr>
															<td class=FieldCell> Name of Guardian</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">

                                                            </span>
                                                            </td>
                                                        </tr>


                                                        <tr>
															<td class=FieldCell> Deposit amount in Rupees</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">Rs &nbsp;
                                                            <%=iAmt%>
                                                            </span>
                                                            </td>
                                                        </tr>
                                                        -->
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
                                                <table  cellpadding="0" cellspacing="0">
                                            <tr>
                                        <td>
                                        <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                <td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p>
                                </td>
                                <td class="GroupTitle" width="126" align="center"><p align="center">Instrument Details</td>
                                <td class="GroupTitleRight"><p align="left">&nbsp;</td>
                                    </tr>
                                        </table>
                                        </td>
                                            </tr>
                                            <tr>
                                        <td class="GroupTable">
                                        <table cellpadding="0" cellspacing="0">
                                    <tr>
                                <td class="MiddlePack" colspan="5"><p align="left"></td>
                                    </tr>
                                    <tr>
                                <td class="FieldCellSub"><p align="left">Instrument Type</p>
                                </td>
                                <td class="FieldCellSub" colspan="6"><p align="left">
                                <Input type="radio" name="optInsType" value="C" checked class="FormElem" onclick="OptFun(this)"> Cheque
                                  &nbsp;<input type="radio" name="optInsType" value="D" class="FormElem" onclick="OptFun(this)"> Demand Draft
                                  &nbsp;<input type="radio" name="optInsType" value="B" class="FormElem" onclick="OptFun(this)">
                                Bankers
                                Cheque&nbsp;<input type="radio" name="optInsType" value="T" class="FormElem" onclick="OptFun(this)">
                                RTGS&nbsp;
                                <input type="radio" name="optInsType" value="W" class="FormElem" onclick="OptFun(this)">
                                <%IF CStr(sVouTy) = "C" Then
										Response.Write "Cash Withdrawn"
								  Else
										Response.Write "Cash Deposited"
								  End IF
                                %>

                                </td>

                                    </tr>
                                    <tr>
										<td class="FieldCellSub"><p align="left">Instrument Number
										</td>
										<td class="FieldCellSub"><p align="left">
										<% IF trim(sPrintCheq) = "1" then %>
											<input type="text" name="txtInsNo" size="10" class="Formelem" maxlength="6" disabled>&nbsp;
											<%If trim(iTransNo) <> "0" then %>
												<Select size=1 name="SelInsNo" class=FormElem disabled>
											<%Else%>
												<Select size=1 name="SelInsNo" class=FormElem >
											<%End IF%>

											<option value=0>Select</option>
											<%
											 	sQry = "Select EntryNo,InstrumentEntryNo,InstrumentNo from Acc_R_BankInstrumentUsage where EntryNo = "&iEntNo&" and Status = 'N'"
											 	Response.Write sQry
											 	rs.Open sQry,con
											 	Do while not rs.Eof
											 	%>
											 	<option value="<%=rs(0)%>-<%=rs(1)%>-<%=rs(2)%>"><%=rs(2)%></option>
											 	<%
											 	rs.MoveNext
											 	loop
											 	rs.Close
											%>
											</Select>
										<% Else  %>
											<input type="text" name="txtInsNo" size="10" class="Formelem" maxlength="6">
										<% End IF 'IF trim(sPrintCheq) = "1" then%>

										<%If trim(iTransNo) <> "0" then %>

											&nbsp;<Select size="1" name="SelAct" class="Formelem" onchange="SelAction()">
											<option value="0">Select</option>
											<option value="C">Cancel</option>
											<option value="R">Reuse</option>
											</Select>

										<%End If %>
										</td></p>

										<td class="FieldCellSub"></td>
										<td class="FieldCellSub"><p align="left">Payable at
										</td>
										<td class="FieldCellSub"><p align="left">
										<!--<input type="text" name="txtIntDate" size="11" class="Formelem" value="<%=Formatdate(date)%>">-->
										<input type="text" name="txtPayableAt" size="20" value="<%=sPayAt%>"  class="Formelem">
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub"><p align="left">Instrument Date
										</td>
										<td class="FieldCellSub"><p align="left">
										<% ' Function Call to Insert Date Picker
											'	Response.Write InsertDatePicker("ctlDate")
										%><input type="date" id="ctlDate" name="ctlDate" class="formelem itms-date-picker" style="width:89px"></p>
										</td>
										<td class="FieldCellSub"></td>
										<td class="FieldCellSub"><p align="left">Drawn On
										</td>
										<td class="FieldCellSub"><p align="left">
										<!--<input type="text" name="txtIntDate" size="11" class="Formelem" value="<%=Formatdate(date)%>">-->
											<input type="text" name="txtDrawnOn" value="<%=sDrawOn%>" size="20" class="Formelem">
										</td>

									</tr>
									<tr>
									<td class="FieldCellSub"><p align="left">Instrument Amount</td>
										<td class="FieldCellSub"><p align="left">
											<input type="text" name="txtAmount" value="" size="15" class="Formelem">
										</td>
										<td class="FieldCellSub"><p align="left">
										<td class="FieldCell"><p align="left">
											<input type="Button" name="ButAddList" Value="Add To List" class="ActionButtonX" Onclick="AddFun()">
										</td>

									</tr>
									<tr>
												<td class="FieldCell" colspan="8"><p align="left">
												<div class="frmBody" id="frm1" style="width: 655; height:80;">
													<table ID="InsTab" border="0" cellspacing="1" class="ExcelTable" width ="100%" ></table>
												</div>
												</td>
											</tr>

                                <td class="MiddlePack" colspan="5"><p align="left"></td>
                                    </tr>

                                        </table>

                                        </td>

                                            </tr>
                                            <tr>

									</tr>



                                                </table>


								</td>
								<td align="center">
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
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
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
