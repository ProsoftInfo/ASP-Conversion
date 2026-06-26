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
	'Program Name				:	InstrumentDetails.asp
	'Module Name				:	Fixed Deposit(Transaction)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Jun 15,2005
	'Modified By				:	S.Maheswari
	'Modified On				:	Jun 13,2008
	'Modified By				:   UmaMaeswari S
	'Tables Used				:	04 April 2010
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Bank Voucher - Instrument Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<xml id="VoucherData" >
<Root></Root>
</xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=Vbscript>
Function CheckSubmit()
Dim Root,objhttp,objhttp1
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set objhttp1 = CreateObject("MSXML2.XMLHTTP")
	IF trim(document.formname.hStatusFlag.value) = "True" then
		alert("This Cheque is Already Used.Updation is Not Possible")
		Exit function
	End IF
	IF document.formname.txtStartNo.value = "" then
		alert("Enter Cheque Start No")
		document.formname.txtStartNo.focus()
		Exit Function
	End IF
	IF document.formname.txtEndNo.value = "" then
		alert("Enter Cheque End No")
		document.formname.txtEndNo.focus()
		Exit Function
	End IF
	IF cint(document.formname.txtStartNo.value) > cint(document.formname.txtEndNo.value) then
		alert("Cheque End No. should be greater than Start No.")
		document.formname.txtEndNo.focus()
		exit function
	End IF

	set Root = VoucherData.DocumentElement
	If Root.haschildnodes then
		For each node in Root.childnodes
			IF trim(node.NodeName)="InstrumentDetails" then
				Set InsDet = node
			End IF
			Root.RemoveChild InsDet
		Next
	End If
	Set Elem = VoucherData.createElement("InstrumentDetails")
	Elem.setAttribute "EntryNo",document.formname.hEntNo.value
	Elem.setAttribute "UnitId",document.formname.hUnitId.value
	Elem.setAttribute "UnitName",document.formname.hUnitName.value
	Elem.setAttribute "BookId",document.formname.hBookId.value
	Elem.setAttribute "BookName",document.formname.hBookName.value
	Elem.setAttribute "AccType",document.formname.hAccType.value
	Elem.setAttribute "AccNo",document.formname.hAccNo.value
	Elem.setAttribute "IssuedOn",document.formname.ctlIssueDate.GetDate
	Elem.setAttribute "StartNo",document.formname.txtStartNo.value
	Elem.setAttribute "EndNo",document.formname.txtEndNo.Value
	Elem.setAttribute "DrawnOn",document.formname.txtDrawnOn.Value
	Elem.setAttribute "PayAt",document.formname.txtPayAt.Value

	IF document.formname.ChkStatus(0).checked = True then
		Elem.setAttribute "Status",document.formname.ChkStatus(0).value
	ElseIF document.formname.ChkStatus(1).checked = True then
		Elem.setAttribute "Status",document.formname.ChkStatus(1).value
	ElseIF document.formname.ChkStatus(2).checked = True then
		Elem.setAttribute "Status",document.formname.ChkStatus(2).value
	End IF
	Root.AppendChild Elem
	'alert(Root.xml)
	objhttp.Open "POST","XMLSave.asp?Name=BankInsDet&Mod=BA", false
	objhttp.send VoucherData.XMLDocument

	'objhttp1.open "GET","BankInsDetailsUpdate.asp?Value="&VoucherData, False
	objhttp1.open "GET","BankInsDetailsUpdate.asp", False
	objhttp1.send
	sRetVal = objhttp1.responseText
	'alert(sRetVal)
	IF trim(sRetVal) = "" then
		MsgBox("Bank  Instrument Details Updated Successfully")
		window.close()
	End IF
	'alert(sRetVal)
End Function

Function FnInit()
	If document.formname.hIssueDate.value <> "" then
		document.formname.ctlIssueDate.SetDate = document.formname.hIssueDate.value
	Else
	 	document.formname.ctlIssueDate.SetDate = Date()
	End IF
End Function
</SCRIPT>
<%
Dim sUnitId,sBookId,sUnitName,sBookName,sAccNo,sAccType,sChkN,sChkH,sChkC
Dim sQry,objrs,objrs1,oDOM,iEntNo,sDrawnOn,sPayAt,iStartNo,iEndNo,dtIssueDate,sStatus
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objrs = Server.CreateObject("ADODB.Recordset")
Set objrs1 = Server.CreateObject("ADODB.Recordset")
sUnitId = Request("UnitId")
sBookId = Request("BookId")
sUnitName = Request("UnitName")
sBookName = Request("BookName")
sAccType  = Request("AccType")
sAccNo    = Request("AccNo")
sDrawnOn  = Request("DrwOn")
sPayAt    = Request("PayAt")
IF trim(sAccType) = "S" then sAccType = ""

'Response.Write "sBookId="&sBookId
sQry = "Select EntryNo,DrawnOn,PayableAt,StartNo,EndNo,convert(VarChar,DateOfIssue,103),Status from "&_
	   "Acc_R_BankInstrumentDetails where OUDefinitionID ='"&sUnitId&"' and BookNumber = '"&sBookId&"' "
'Response.Write sQry
objrs.Open sQry,con
IF Not objrs.EOF then
	iEntNo		= objrs(0)
	'sDrawnOn	= objrs(1)
	'sPayAt		= objrs(2)
	iStartNo	= objrs(3)
	iEndNo		= objrs(4)
	dtIssueDate	= objrs(5)
	sStatus		=  objrs(6)
End IF
objrs.Close
IF sStatus	= "" or sStatus	= "N" then sChkN = "Checked"
IF sStatus	= "H" then sChkH = "Checked"
IF sStatus	= "C" then sChkC = "Checked"
dim sStatusFlag
sStatusFlag = False
'Validation for updating ins details
'Response.Write iEntNo &"-----"
IF trim(iEntNo) <> "" and trim(iEntNo) <> "0" then
	sQry = "Select InstrumentEntryNo from Acc_R_BankInstrumentUsage where EntryNo = "&iEntNo
	objrs.Open sQry,con
	Do while not  objrs.EOF
		sQry = "Select Status from Acc_R_BankInstrumentUsage where EntryNo = "&iEntNo&" and InstrumentEntryNo = "&objrs(0)
		'Response.Write sQry
		objrs1.Open sQry,con
		IF Not objrs1.EOF then
			If objrs1(0) = "U" then sStatusFlag = True
		End If
		objrs1.Close
	objrs.MoveNext
	loop
	objrs.Close
End IF 'IF trim(iEntNo) <> "" and trim(iEntNo) <> "0" then
'Response.Write sStatusFlag
'Set Root = VoucherData.documentElement
%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit()" >

<form method="POST" name="formname">
<input type="hidden" name="hUnitId" value="<%=sUnitId%>">
<input type="hidden" name="hBookId" value="<%=sBookId%>">
<input type="hidden" name="hUnitName" value="<%=sUnitName%>">
<input type="hidden" name="hBookNAme" value="<%=sBookName%>">
<input type="hidden" name="hAccType" value="<%=sAccType%>">
<input type="hidden" name="hAccNo" value="<%=sAccNo%>">
<input type="hidden" name="hEntNo" value="<%=iEntNo%>">
<input type="hidden" name="hIssueDate" value="<%=dtIssueDate%>">
<input type="hidden" name="hStatusFlag" value="<%=sStatusFlag%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Instrument
          Details
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
                                                <table cellpadding="0" cellspacing="0">
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
                                        <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                <td class="MiddlePack" colspan="4"><p align="left"></td>
                                    </tr>
                                    <!--<tr>
                                <td class="FieldCellSub"><p align="left">Unit</p>
                                </td>
                                <td class="FieldCellSub" colspan="3"><p align="left"><span id="spUnit" class="DataOnly">
                                <%=sUnitName%> </span>
                                </td>

                                    </tr>-->
                                    <tr>
										<td class="FieldCellSub"><p align="left">Book
                                          Name
										</td>
										<td class="FieldCellSub" colspan="3"><p align="left"><span id="spBook" class="DataOnly">
                                        <%=sBookName%></span> </p>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Account Type
										</td>
										<td class="FieldCellSub" colspan="3"><span id="spAccType" class="DataOnly"><%=sAccType%></span>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Account Number
										</td>
										<td class="FieldCellSub" colspan="3"><span id="spAccNo" class="DataOnly"><%=sAccNo%></span>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Cheque Book
                                          Issued On
										</td>
										<td class="FieldCellSub" colspan="3">
										<object id="ctlIssueDate"   classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"         codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
											<param name="_ExtentX" value="2355">
											<param name="_ExtentY" value="529">
										</object>
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub">Cheque Start No
										</td>
										<td class="FieldCellSub"><input type="text" name="txtStartNo" value="<%=iStartNo%>" size="15" maxlength="6" class="FormElem">
										</td>
										<td class="FieldCellSub">Cheque End No
										</td>
										<td class="FieldCellSub"><input type="text" name="txtEndNo" Value="<%=iEndNo%>" size="15" maxlength="6" class="FormElem">
										</td>
									</tr>

                                    <tr>
										<td class="FieldCellSub">Drawn On
										</td>
										<td class="FieldCellSub"><input type="text" name="txtDrawnOn" Value="<%=sDrawnOn%>" size="20" class="FormElem" disabled>
										</td>
										<td class="FieldCellSub">Payable At
										</td>
										<td class="FieldCellSub"><input type="text" name="txtPayAt" value="<%=sPayAt%>" size="20" class="FormElem" disabled>
										</td>
                                    </tr>
                                    <tr>
										<td class="FieldCellSub">Status
										</td>
										<td class="FieldCellSub" colspan="3"><input type="radio" value="N" name="ChkStatus" class="FormElem" <%=sChkN%>>
                                          New&nbsp; <input type="radio" name="ChkStatus" value="H" class="FormElem"  <%=sChkH%>>
                                          Hold&nbsp; <input type="radio" name="ChkStatus" value="C" class="FormElem"  <%=sChkC%>>&nbsp;
                                          Closed
										</td>
                                    </tr>
                                    <tr>
										<td class="FieldCellSub">Last Used
                                          Cheque No
										</td>
										<td class="FieldCellSub" colspan="3"><span id="spChkLastNo" >&nbsp;</span>&nbsp;<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" width="11" alt="" height="11" valign="top" style="cursor:hand" onClick="">
										</td>
                                    </tr>

                                    <tr>
                                <td class="MiddlePack" colspan="4"><p align="left"></td>
                                    </tr>
                                        </table>
                                        </td>
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