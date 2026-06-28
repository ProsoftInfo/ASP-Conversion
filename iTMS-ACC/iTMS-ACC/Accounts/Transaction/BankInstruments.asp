<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BankInstruments.asp
	'Module Name				:	Accounts
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 02, 2010
	'Modified By				:
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<%
Dim rsObj,rsObj1
Dim nSlNo,iStartRec,iEndRec,sorgID
Dim iPrevPage,iTotalPages,nPageCtr,iNextPage,iPageSize,iPageNo
Dim iTotalRecords,sBookNo,sBankName,sAccType,nAccNo,nInsNo
Dim sQuery,sBookName,sBookTypeName,sIssuedOn

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsObj1 = Server.CreateObject("ADODB.Recordset")

sorgID = Session("organizationcode")

sBookNo   = Request("hBookNo")
sBankName = Request("hBankName")
sAccType  = Request("hAccType")
nAccNo	  = Request("hAccNo")
nInsNo	  = Request("hInsNo")
sIssuedOn = Request("hIssuedOn")

iPageSize = 20

iPageNo = trim(Request("hPage"))
if trim(iPageNo) = "" then iPageNo = 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Accounts</title>
<XML id="OutData">
<Root/>
</XML>
<xml id="GLHeadData"><Root></Root></xml>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<Script Language="javascript">
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function field(name) {
	var lower = String(name).toLowerCase();
	var form = document.formname;
	if (form.elements[name]) {
		return form.elements[name];
	}
	for (var i = 0; i < form.elements.length; i += 1) {
		if (String(form.elements[i].name).toLowerCase() === lower) {
			return form.elements[i];
		}
	}
	return null;
}

function fields(name) {
	var item = field(name);
	if (!item) {
		return [];
	}
	if (item.length != null && !item.tagName) {
		return Array.prototype.slice.call(item);
	}
	return [item];
}

function selectedRowIndex() {
	var radios = fields("radButton");
	for (var i = 0; i < radios.length; i += 1) {
		if (radios[i].checked) {
			return i + 1;
		}
	}
	return 0;
}

function openModernDialog(url, args, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	} else {
		window.open(url, "_blank", "height=400,width=600,resizable=no,status=no");
	}
}

function accountTypeName(value) {
	var text = trim(value);
	if (text === "CC") {
		return "Cash Credit Account";
	}
	if (text === "CA" || text === "CU") {
		return "Current Account";
	}
	return text;
}

function InsDetails() {
	var rowIndex = selectedRowIndex();
	var url;
	if (rowIndex === 0) {
		alert("Select any One Book");
		return;
	}
	url = "InstrumentDetails.asp?UnitId=" + encodeURIComponent(document.formname.hOrgCode.value) +
		"&UnitName=" + encodeURIComponent(document.formname.hOrgName.value) +
		"&BookId=" + encodeURIComponent(field("BookNo" + rowIndex).value) +
		"&BookName=" + encodeURIComponent(field("BookName" + rowIndex).value) +
		"&AccType=" + encodeURIComponent(accountTypeName(field("AccType" + rowIndex).value)) +
		"&AccNo=" + encodeURIComponent(field("AccNo" + rowIndex).value) +
		"&DrwOn=" + encodeURIComponent(field("BankName" + rowIndex).value) +
		"&PayAt=";
	openModernDialog(url, window.OutData || document.getElementById("OutData"), "dialogHeight:330px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No");
}

function AssignPage(nPage) {
	document.formname.hPage.value = nPage;
	document.formname.submit();
}

function selectedRadioValue(name) {
	var list = fields(name);
	for (var i = 0; i < list.length; i += 1) {
		if (list[i].checked) {
			return list[i].value;
		}
	}
	return "";
}

function CheckSubmit() {
	var issuedOnField = field("txtFromDate");
	var issuedOn = trim(issuedOnField.value);
	var month;
	document.formname.hBookNo.value = document.formname.selBook.options[document.formname.selBook.selectedIndex].value;
	document.formname.hBankName.value = document.formname.txtBankName.value;
	document.formname.hAccType.value = selectedRadioValue("optAccType");
	document.formname.hAccNo.value = document.formname.txtAccNo.value;
	document.formname.hInsNo.value = document.formname.txtInsNo.value;
	if (issuedOn !== "") {
		if (issuedOn.length < 10) {
			alert("Enter Valid Date");
			issuedOnField.focus();
			return;
		}
		month = Number(issuedOn.split("/")[1]);
		if (month >= 13) {
			alert("Enter Valid Month");
			issuedOnField.focus();
			return;
		}
		document.formname.hIssuedOn.value = issuedOnField.value;
	}
	document.formname.submit();
}

function ShowBankDetails(sOrgCode, sBookCode, sBookNumber) {
	var url = "BankDetails.asp?OrgCode=" + encodeURIComponent(sOrgCode) + "&BookCode=" + encodeURIComponent(sBookCode) + "&BookNumber=" + encodeURIComponent(sBookNumber);
	openModernDialog(url, "", "dialogHeight:400px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No", function (returnValue) {
		if (returnValue === "Done") {
			document.formname.submit();
		}
	});
}

function ShowBankBookDet(sOrgCode, sBookCode, sBookNumber, FromAccHead) {
	var url = "BankBookDetailsPopup.asp?OrgCode=" + encodeURIComponent(sOrgCode) + "&BookCode=" + encodeURIComponent(sBookCode) + "&BookNumber=" + encodeURIComponent(sBookNumber) + "&FromAcc=" + encodeURIComponent(FromAccHead || "");
	openModernDialog(url, "", "Status:No;");
}
</Script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hOrgCode" value="<%=sorgID%>">
	<input type="hidden" name="hOrgName" value="<%=Session("orgshortName")%>">
	<input type="hidden" name="hBookNo" value="<%=sBookNo%>">
	<input type="hidden" name="hBankName" value="<%=sBankName%>">
	<input type="hidden" name="hAccType" value="<%=sAccType%>">
	<input type="hidden" name="hAccNo" value="<%=nAccNo%>">
	<input type="hidden" name="hInsNo" value="<%=nInsNo%>">
	<Input type="hidden" name="hIssuedOn" value="<%=sIssuedOn%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Bank Instrument Control
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>

								<tr>
									<td align="center" width="5" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
											<tr>
												<td>
													<div>
														<table class="CollapseBand" cellspacing="0" cellpadding="0">
															<tr>
																<td valign="center"><a style="width: 1em; height: 1em;" title href onclick="Div_OnClick(idUnprocessed,'')" >
																	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
																	</a>
																</td>
																<td valign="center" class="SubTitle">&nbsp;&nbsp;
																	<%

																		'if cstr(sBookType)="" or cstr(sBookType)="0" then
																		'	Response.Write "All Books"
																		'else
																		'	Response.Write sBookTypeName
																		'end if
																	%>
																</td>
															</tr>

														</table>
														<table border="0" cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="display: none">
																		<table cellpadding="0" cellspacing="0" class="BodyTable" Width="100%">
																			<tr>
																				<td class="MiddlePack">
																				</td>
																				<td class="MiddlePack" colspan="4">
																				</td>
																			</tr>

																			<tr>

																				<td class="FieldCellSub">Bank Book</td>
																				<td class="FieldCellSub">
																					<Select name="selBook" class="FormElem">
																						<OPTION value="0">Select</option>
																						<%
																							'sQuery = "Select BookCode,BookName from ACC_M_DayBooks"
																							sQuery ="select BookNumber,Upper(BookName) From vwOrgBookNames where BookCode=02 AND BookAccountHead is not null "
																							rsObj.Open sQuery,con
																							if not rsObj.EOF then
																								do while not rsObj.EOF
																									if CInt(rsObj(0))=CInt(sBookNo) then
																										Response.Write "<OPTION value="& rsObj(0) &" Selected>"& rsObj(1) &"</option>"
																									else
																										Response.Write "<OPTION value="&  rsObj(0) &" >"& rsObj(1) &"</option>"
																									end if
																									rsObj.MoveNext
																								loop
																							end if
																							rsObj.Close
																						%>
																					</Select>
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub">Bank Name</td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtBankName" value="" class="FormElem">
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub">Account Type</td>
																				<td class="FieldCellSub">
																					<Input type="Radio" Name="optAccType" Value="CC" <%If sAccType="CC" Then Response.Write "checked"%>>Cash Credit Account
																					<Input type="Radio" Name="optAccType" Value="CA" <%If sAccType="CA" Then Response.Write "checked"%>>Current Account
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub">Account Number</td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtAccNo"  class="FormElem" maxLength="30">
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub">Instrument No</td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtInsNo" value="" class="FormElem">
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub">Cheque Book Issued On</td>
																				<td class="FieldCellSub">
																					<Input Type="Text" Name="txtFromDate" class="FormElem" MaxLength="10" size="10">
																				</td>
																				<td>
																					<input type="button" name="btnGo" value="GO" class="ActionButtonX" onClick="CheckSubmit()">
																				</td>
																			</tr>

																			<tr>
																				<td class="Middlepack" colspan="4">

																				</td>
																			</tr>

																		</table>
																	</div>
																</td>
															</tr>

														</table>
													</div>
												</td>
											</tr>

										</table>
									</td>
									<td align="center" class="ClearPixel" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>

								<tr>
									<td align="center" class="MiddlePack" colspan="3">
									</td>
								</tr>

								<tr>
									<td align="center" width="5" class="ClearPixel">
									</td>
									<td valign="top">
										<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
										<table border="0" cellspacing="1px" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" width="10px" Rowspan="2">S.No.
												</td>
												<td class="ExcelHeaderCell" Rowspan="2"><a style="width: 1em; height: 1em;" title href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
													<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Expands this section for more search criteria." width="15px" height="15px">
													</a>
												</td>
												<td class="ExcelHeaderCell" Rowspan="2">Book Name
												</td>
												<td class="ExcelHeaderCell" Rowspan="2">Bank Name
												</td>
												<td class="ExcelHeaderCell" Rowspan="2">Account Type
												</td>
												<td class="ExcelHeaderCell" Rowspan="2">Account No
												</td>
												<td class="ExcelHeaderCell" colspan="2">Instrument
												</td>
											</tr>
											<tr>
												<td class="ExcelHeaderCell">Nos</td>
												<td class="ExcelHeaderCell">Last Used No</td>
											</tr>

											<%	Dim iCtr


												sQuery=" select BookNumber,BookName,isnull(BookAccountHead,0),BookCode,Useable from Acc_R_ApplicableAccountHeads Where"&_
														" OUDefinitionID='"& sorgID &"' "

												sQuery = " select A.BookNumber,A.BookName,isnull(A.BookAccountHead,0),A.BookCode,A.Useable,"&_
														 " B.BankName,B.AccountType,B.AccountNo "&_
														 " from Acc_R_ApplicableAccountHeads A,Acc_M_BankDetails B Where "&_
														 " A.OUDefinitionID='"& sorgID &"' and A.BookCode=B.BookCode and A.BookNumber=B.BookNumber"


												if Trim(sBookNo) <> "" and sBookNo <> "0" then
													sQuery = sQuery & " and A.BookNumber = '"& sBookNo &"' "
												end if

												If Trim(sBankName) <> "" Then
													sQuery = sQuery & " and B.BankName like '"& sBankName &"%'"
												End IF

												If Trim(sAccType) <> "" Then
													sQuery = sQuery & " and B.AccountType = '"& sAccType & "' "
												End IF

												If nAccNo <> "" Then
													sQuery = sQuery & "and B.AccountNo =' "& nAccNo & "' "
												End IF

												If nInsNo <> "" Then
													sQuery = sQuery & " and A.BookNumber IN (select Distinct BookNumber From Acc_R_BankInstrumentDetails Where StartNo  = "& cint(nInsNo) &"  ) "
												End IF

												If Trim(sIssuedOn) <> "" Then
													sQuery = sQuery & " and A.BookNumber IN (select Distinct BookNumber From Acc_R_BankInstrumentDetails Where convert(VarChar,DateOfIssue,103) = Convert(Varchar,'"& sIssuedOn &"',103) )"
												End IF

												'Response.Write sQuery
												with rsObj
													.CursorLocation = 3
													.CursorType = 3
													.ActiveConnection = con
													.Source = sQuery
													.Open
												end with

												nSlNo = 1
												iCtr = 1
												If not rsObj.EOF Then
													iTotalPages = rsObj.PageCount
													iTotalRecords = rsObj.RecordCount
													rsObj.AbsolutePage = iPageNo
												Else
													iTotalPages = 0
													iTotalRecords = 0

													iStartRec = 0
													iEndRec = 0
												End If

												if trim(iPageNo) = 1 then
													iPrevPage = 0
												else
													iPrevPage = iPageNo - 1
												end if


												if iTotalPages >= iPageNo + 1 then
													iNextPage = iPageNo + 1
												else
													iNextPage = 0
												end if

												do while not rsObj.EOF and nSlNo < iPageSize

											%>

											<tr>
												<td class="ExcelSerial" align="center"><%=nSlNo%>
												</td>

												<Input type="Hidden" Name="BookNo<%=iCtr%>" value="<%=rsobj(0)%>">
												<Input type="Hidden" Name="BookName<%=iCtr%>" value="<%=rsobj(1)%>">
												<Input type="Hidden" Name="BankName<%=iCtr%>" value="<%=rsobj(5)%>">
												<Input type="Hidden" Name="AccType<%=iCtr%>" value="<%=rsobj(6)%>">
												<Input type="Hidden" Name="AccNo<%=iCtr%>" value="<%=rsobj(7)%>">

												<td class="ExcelDisplayCell" align="center" width="10">
													<input type="radio" name="radButton" value="<%=rsObj(0)%>">
												</td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(1)%></td>
												<td class="ExcelDisplayCell" align="left">
													<a href="#" class="ExcelDisplayLink" onClick="ShowBankDetails('<%=sorgID%>','<%=rsObj(3)%>','<%=rsObj(0)%>')">
													<%=rsobj(5)%></a></td>
												<td class="ExcelDisplayCell" align="Left"><%If rsObj(6)="CC" Then Response.Write "Cash Credit Account" Else Response.Write "Current Account" End IF %></td>
												<td class="ExcelDisplayCell" align="left"><%=rsobj(7)%></td>
													<%
														Dim sInstNo
														squery = "Select StartNo,EndNo from Acc_R_BankInstrumentDetails Where BookNumber="& rsObj(0)&""

														rsObj1.Open sQuery,con
														if not rsObj1.EOF then
															sInstNo = rsObj1(0) & "-" & rsObj1(1)
														end if
														rsObj1.Close
													%>
												<td class="ExcelDisplayCell" align="left"><%=sInstNo%></td>
												<td class="ExcelDisplayCell" align="left"></td>
											</tr>
											<%
													nSlNo=nSlNo + 1
													iCtr = iCtr + 1
													rsObj.MoveNext
												loop
											%>
											<input type="hidden" name="hCnt" value="<%=iCtr-1%>" >
										</table>
										<!--/div-->
									</td>
									<td align="center" class="ClearPixel" width="5">
									</td>
								</tr>

								<tr>
									<td align="center" class="MiddlePack" colspan="3">
									</td>
								</tr>

								<tr>
									<td align="center" width="5" class="ClearPixel">
									</td>
									<td valign="top" align="right">

									<input type="button" value=" |< " class="ActionButtonX" id=ButFirst name=ButFirst onClick="AssignPage('1')">

									<%if trim(iPrevPage) = "0" then  %>
										<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev >
									<%else%>
										<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev onClick="AssignPage('<%=iPrevPage%>')">
									<%end if %>


									<SELECT class="FormElem" onChange="AssignPage(this.value)"  id="mCmbPage" name="mCmbPage">

									<%for nPageCtr= 1 to iTotalPages %>
										<option value="<%=nPageCtr%>" <%if trim(iPageNo) = trim(nPageCtr) then Response.Write "Selected" %> >Page <%=nPageCtr%> of <%=iTotalPages %></option>
									<%next%>

									</SELECT>
									<%if trim(iNextPage) = "0" then  %>
										<input type="button" value=" >> " class="ActionButtonX" id=ButNext name=ButNext >
									<%else%>
										<input type="button" value=" >> " class="ActionButtonX" onclick="AssignPage('<%=iNextPage%>')" id=ButNext name=ButNext >
									<%end if%>

									<input type="button" value=" >| " class="ActionButtonX" id=ButLast name=ButLast OnClick="AssignPage('<%=iTotalPages %>')">

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
												<td class="ActionCell">
													<input type="button" value="Manage Instruments" name="BtnNewParty" class="ActionButtonX" onclick="InsDetails()">
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
</body>
</html>
