<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BookNarrations.asp
	'Module Name				:	Accounts
	'Author Name				:	UmaMaheswari S
	'Created On					:	March 26, 2011
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
<!--#include file="../../include/sessionVerify.asp"-->
<%
Dim nSlNo,sNarration,sSelDayBook,iBookCode,sQuery,sBookCode,sDayBookLabel
Dim iStartRec,iEndRec,iPrevPage,iTotalPages,nPageCtr,iNextPage,iPageSize,iPageNo
Dim iTotalRecords,sBookName

Dim rsObj,rs

set rsObj = Server.CreateObject("ADODB.Recordset")
set rs = Server.CreateObject("ADODB.Recordset")

sBookCode = trim(Request("BookCode"))
sSelDayBook = Request("hDayBook")
sNarration  = Request("hNarration")

If trim(sBookCode) = "01" Then
	sDayBookLabel = "Cash Book"
ElseIf trim(sBookCode) = "02" Then
	sDayBookLabel = "Bank Book"
Else
	sDayBookLabel = "Day Book"
End If

iPageSize = 20

iPageNo = trim(Request("hPage"))
if trim(iPageNo) = "" then iPageNo = 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Accounts</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="RetData"><Root/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<Script Language="VBScript">
Dim sType
'**********************************************
Function ViewContactDeatils(sPartyCode)
	showModalDialog "ParDisplayContactDetails.asp?PartyCode="&sPartyCode,"","dialogHeight:350px;Status:no"
End Function
'************************************************
Function CreateNew()
	sBookCode = document.formname.hBookCode.value
	sType = "N"
	set OutValue = showModalDialog("NarrationEntryPopUp.asp?BookCode="&sBookCode&"&Type="&sType,"","dialogHeight:190px;Status:no")
	If OutValue.getAttribute("Done") = "Y" Then
		document.formname.submit()
	End IF
End Function


Function DeleteData()
	Dim nCnt,iNoOfSelect,iCtr,nNarrationNo,nBookCode,nBookNo

	nCnt = document.formname.hCnt.value

	For iCtr = 1 to nCnt
		n1 = Eval("document.formname.chkbox"+Trim(iCtr)).value
		check = Eval("document.formname.chkbox"+Trim(iCtr)).checked

		If check Then
			If n1 <> "" Then
				Arr = Split(n1,":")
				nNarrationNo = nNarrationNo & "," & Arr(0)
				'nBookCode = Arr(1)
				'nBookNo = Arr(2)
			End IF
			iNoOfSelect = iNoOfSelect + 1
		End IF

	Next

	If iNoOfSelect = 0 Then
		alert("Select any one Narration For Delete")
		Exit Function
	End IF

	If nNarrationNo <> "" Then nNarrationNo = mid(nNarrationNo,2)

	document.formname.action = "NarrationDelete.asp?NarrationNo="&nNarrationNo&"&BookCode="&document.formname.hBookCode.value
	document.formname.submit
End Function

'**********************************
Function AssignPage(nPage)
	document.formname.hPage.value = nPage
	document.formname.submit()
End Function
'************************************************
Function CheckSubmit()
	document.formname.hDayBook.value= document.formname.selDayBook(document.formname.selDayBook.selectedIndex).value
	document.formname.hNarration.value = document.formname.txtNarration.value
	document.formname.submit
End Function
'************************************************
Function EditNarration(nNarrNo)
	sBookCode = document.formname.hBookCode.value
	sType = "E"
	set OutValue = showModalDialog("NarrationEntryPopUp.asp?BookCode="&sBookCode&"&Type="&sType&"&NarrNo="&nNarrNo,"","dialogHeight:190px;Status:no")
	If OutValue.getAttribute("Done") = "Y" Then
		document.formname.submit()
	End IF
End Function
'************************************************
'************************************************
</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "bookNarrations" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hDayBook" value="<%=sSelDayBook%>">
	<input type="hidden" name="hNarration" value="<%=sNarration%>">
	<input type="hidden" name="hCallTy" value="">
	<input type="hidden" name="hBookCode" value="<%=sBookCode%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Narration
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
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" width="5" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
											<tr>
												<td>
													<div>
														<table class="CollapseBand" cellspacing="0" cellpadding="0" class="BodyTable">
															<tr>
																<td valign="center"><a style="width: 1em; height: 1em;" title href onclick="Div_OnClick(idUnprocessed,'')" >
																	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
																	</a>
																</td>
																<td valign="center" class="SubTitle">&nbsp;&nbsp;

																</td>
															</tr>

														</table>
														<table border="0" cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="display: none">
																		<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
																			<tr>
																				<td class="MiddlePack" colspan="3">
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub"><%=sDayBookLabel%></td>
																				<td class="FieldCellSub">
																					<Select name="selDayBook" class="FormElem">
																						<option value="S">Select</option>
																						<%
																						with rsObj
																							.CursorLocation = 3
																							.CursorType = 3
																							.Source = "Select BookNumber,BookName From VWOrgBookNames Where BookCode="& sBookCode &" Order by BookCode"
																							.ActiveConnection = con
																							.Open
																						end with
																						If Not rsObj.EOF Then
																							Do while Not rsObj.EOF
																								%><option value="<%=rsObj(0)%>"><%=rsObj(1)%></option><%
																							rsObj.MoveNext
																							Loop
																						End IF
																						rsObj.Close
																						%>
																					</Select>
																				</td>
																				<td class="FieldCell">&nbsp;</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub">Narration</td>
																				<td class="FieldCellSub" colspan="2">
																					<input type="text" name="txtNarration" value="" class="FormElem">&nbsp;
																					<input type="button" name="btnGo" value="  GO  " class="ActionButtonX" onClick="CheckSubmit()">
																				</td>
																			</tr>
																			<tr>
																				<td class="Middlepack" colspan="5">

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
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" width="10" >S.No.</td>
												<td class="ExcelHeaderCell" width="20px">
													<a style="width: 1em; height: 1em;">
														<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Click here to delete the Narration" onclick="DeleteData()" width="15px" height="15px">
													</a>
												</td>
												<td class="ExcelHeaderCell">Short Description</td>
												<td class="ExcelHeaderCell">Description</td>
												<td class="ExcelHeaderCell">Used in Book</td>
											</tr>

											<%
												sQuery = " Select Distinct NarrationNumber,NarrationShortDesc,NarrationDesc,isNull(BookCode,''),isNull(BookNumber,'') From VW_ACC_BookNarrations Where BookCode = '"& sBookCode &"' "

												If sNarration <> "" Then
													sQuery = sQuery & " And NarrationDesc like '"& sNarration &"%' "
												End IF

												If sSelDayBook <> "" and sSelDayBook <> "S" Then
													If sNarration <> "" Then
														sQuery = sQuery & " and BookNumber ='"& sSelDayBook &"' And NarrationDesc like '"& sNarration &"%' "
													Else
														sQuery = sQuery & " and BookNumber ='"& sSelDayBook &"' "
													End IF
												End IF

												sQuery = sQuery & " order by NarrationNumber"

												'Response.Write sQuery

												with rsObj
													.CursorLocation = 3
													.CursorType = 3
													.ActiveConnection = con
													.Source = sQuery
													.Open
												end with

												nSlNo = 1
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
													iBookCode = rsObj(3)

													with rs
														.CursorLocation = 3
														.CursorType = 3
														.ActiveConnection = con
														.Source = "Select Distinct BookName From VWOrgBookNames Where BookCode='"& rsObj(3)&"' and  BookNumber='"& rsObj(4)&"'"
														.Open
													end with
													If Not rs.EOF Then
														sBookName = rs(0)
													End IF
													rs.Close
											%>

											<tr>
												<td class="ExcelSerial" align="center"><%=nSlNo%></td>
												<td class="ExcelDisplayCell" align="center" width="10px">
													<input type="CheckBox" name="chkbox<%=nSlNo%>" value="<%=rsObj(0)%>:<%=rsObj(3)%>:<%=rsObj(4)%>">
												</td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(1)%></td>
												<td class="ExcelDisplayCell" align="left"><a href=# class=ExcelDisplayLink language=VBScript onclick="EditNarration(<%=rsObj(0)%>)"><%=rsObj(2)%></a></td>
												<td class="ExcelDisplayCell" align="left"><%=sBookName%></td>
											</tr>
											<%
													nSlNo=nSlNo + 1
													rsObj.MoveNext
												loop
											%>
											<input type="hidden" name="hCnt" value="<%=nSlNo-1%>" >
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
									<td align="center" width="5px" class="ClearPixel">
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
									<td align="center" width="5px" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td class="ActionCell">
													<input type="button" value="Add New" name="BtnNew" class="ActionButtonX" onclick="CreateNew()">
												</td>
											</tr>

										</table>
									</td>
									<td align="center" class="ClearPixel" width="5px">
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
