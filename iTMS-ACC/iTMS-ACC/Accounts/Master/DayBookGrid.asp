<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	DayBookGrid.asp
	'Module Name				:	Accounts
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 09,2010
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
Dim rsObj,rsObj1,rsObj2,rsObj3
Dim nSlNo
Dim iPartyCode,iStartRec,iEndRec,iUseable
Dim iPrevPage,iTotalPages,nPageCtr,iNextPage,iPageSize,iPageNo
Dim iTotalRecords,iAccHead
Dim sorgID
Dim sQuery,sBookType,sBookName,sBookTypeName,sFromAccHead,sGLAccHead

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsObj1 = Server.CreateObject("ADODB.Recordset")
set rsObj2 = Server.CreateObject("ADODB.Recordset")
set rsObj3 = Server.CreateObject("ADODB.Recordset")
' Response.Write "Request(parName)="& Request("hParName")
sBookType = Request("hBookType")
sBookTypeName = Request("hBookTypeName")
sBookName = Request("hBookName")
'sorgID =  Request("hOrgCode")
sorgID = Session("organizationcode")
if trim(sorgID)="" then
	sorgID = "010101"
end if

iPageSize = 20

iPageNo = trim(Request("hPage"))
if trim(iPageNo) = "" then iPageNo = 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Accounts</title>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root></Root></script>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "dayBookGrid" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hBookName" value="<%=sBookName%>">
	<input type="hidden" name="hBookType" value="<%=sBookType%>">
	<input type="hidden" name="hOrgCode" value="<%=sorgID%>">
	<input type="hidden" name="hPartyCode" value="">
	<input type="hidden" name="hBookTypeName" value="<%=sBookTypeName%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Day Books
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
									<td align="center" width="5px" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
											<tr>
												<td>
													<div>
														<table class="CollapseBand" cellspacing="0" cellpadding="0">
															<tr>
																<td valign="center"><a style="width: 1em; height: 1em;" title href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" >
																	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
																	</a>
																</td>
																<td valign="center" class="SubTitle">&nbsp;&nbsp;
																	<%

																		if cstr(sBookType)="" or cstr(sBookType)="0" then
																			Response.Write "All Books"
																		else
																			Response.Write sBookTypeName
																		end if
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
																		<!--	<tr>
																				<td width=100></td>
																				<td class="FieldCell">Organisation</td>
																				<td class="FieldCellSub">
																					<Select name="selOrganisaion" class="FormElem">
																						<%
																							populateOrganizationList()
																						%>
																					</Select>
																				</td>
																			</tr>-->
																			<tr>
																				<td width=100></td>
																				<td class="FieldCell">Day Book Type</td>
																				<td class="FieldCellSub">
																					<Select name="selBookType" class="FormElem">
																						<OPTION value="0">Select a Day Book</option>
																						<%
																							sQuery = "Select BookCode,BookName from ACC_M_DayBooks"
																							rsObj.Open sQuery,con
																							if not rsObj.EOF then
																								do while not rsObj.EOF
																									if cstr(rsObj(0))=cstr(sBookType) then
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
																				<td width="100px"></td>
																				<td class="FieldCell">Book Name</td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtBookName" value="<%=sBookName%>" class="FormElem">
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
									<td align="center" class="ClearPixel" width="5px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>

								<tr>
									<td align="center" class="MiddlePack" colspan="3">
									</td>
								</tr>

								<tr>
									<td align="center" width="5px" class="ClearPixel">
									</td>
									<td valign="top">
										<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
										<table border="0" cellspacing="1px" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10px" >S.No.
												</td>
												<td class="ExcelHeaderCell" align="center" ><a style="width: 1em; height: 1em;" title href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
													<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Expands this section for more search criteria." width="15px" height="15px">
													</a>
												</td>
												<td class="ExcelHeaderCell">Book Name
												</td>
												<td class="ExcelHeaderCell">GL Account Head
												</td>
												<td class="ExcelHeaderCell">Contra
												</td>
												<td class="ExcelHeaderCell">Active
												</td>
											</tr>

											<%


												sQuery=" select BookNumber,BookName,isnull(BookAccountHead,0),BookCode,Useable from Acc_R_ApplicableAccountHeads Where"&_
														" OUDefinitionID='"& sorgID &"' "

												if trim(sBookName)<>"" then
													sQuery = sQuery & " and BookName like '"& sBookName &"%'"
												end if
												if trim(sBookType)<>"0" and trim(sBookType)<>"" then
													sQuery = sQuery & " and BookCode = '"& sBookType &"'"
												end if

												  ' Response.Write sQuery
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
													iPartyCode = rsObj(0)
													iUseable = rsObj(4)
											%>

											<tr>
												<td class="ExcelSerial"><%=nSlNo%>
												</td>
												<td class="ExcelDisplayCell" width="10px">
													<input type="radio" name="radButton" value="<%=iPartyCode%>">
												</td>
												<%
													sQuery = "select count(a.ToAccountHead) from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
													"where a.OUDefinitionID='"& sorgID &"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead = "& rsObj(2)
													'Response.Write sQuery
													rsObj1.Open sQuery,con
													if not rsObj1.EOF then
														if rsObj(3)="01" or rsObj(3)="02" then
															sFromAccHead = rsObj(2)
														else
															sFromAccHead = 0
														end if
													end if
													rsObj1.Close
												%>
												<td class="ExcelDisplayCell" align="left">
													<%If rsObj(3) = "02" Then%>
														<img src="../../assets/images/iTMS%20icons/DetailsIcon.gif" onClick="ShowBankBookDet('<%=sorgID%>','<%=rsObj(3)%>','<%=rsObj(0)%>','<%=sFromAccHead%>')" alt="Bank Details" style="cursor: pointer" >
													<%End IF%>
														<a href="#" class="ExcelDisplayLink" onClick="EditBook('<%=sorgID%>','<%=rsObj(3)%>','<%=rsObj(0)%>','<%=sFromAccHead%>'); return false;"><%=rsObj(1)%></a>
												</td>
												<td class="ExcelDisplayCell" align="left">
												<%
													sGLAccHead = ""
													if rsObj(2)<>"0" then

														sQuery="select AccountDescription,AccountHeadCode from Acc_M_GLAccountHead where AccountHead="&rsObj(2)
														rsObj1.Open sQuery,con
														if not rsObj1.EOF then
															sGLAccHead =  rsObj1(0)
														end if
														rsObj1.Close
													end if ' if rsObj(2)<>"0" then

													sQuery="Select count(1) from Acc_T_CreatedVoucherHeader where BookNumber="&trim(rsObj(0))&_
																" and OUDefinitionID='"&sorgID&"' and BookCode='"&rsObj(3)&"'"
													with rsObj1
														.CursorLocation = 3
														.CursorType = 3
														.Source = sQuery
														.ActiveConnection = con
														.Open
													end with
													if not rsObj1.EOF then
														if rsObj1(0) >  0 then
															Response.Write sGLAccHead
														else
															if Trim(sGLAccHead)="" then sGLAccHead = "Map Account Head"
															%>
																<a href="#" class="ExcelDisplayLink" onClick="GlChange('<%=sorgID%>','<%=rsObj(0)%>','<%=rsObj(3)%>'); return false;"><%=sGLAccHead%></a>
															<%
														end if

													end if
													rsObj1.Close
												%>
												</td>
												<td class="ExcelDisplayCell">
												<%
													sQuery = "select count(a.ToAccountHead) from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
														"where a.OUDefinitionID='"& sorgID &"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead = "& rsObj(2)
														'Response.Write sQuery
														rsObj1.Open sQuery,con
														if not rsObj1.EOF then
															if rsObj(3)="01" or rsObj(3)="02" then
																Response.Write "<a href='#' onClick=""ViewContraDet('"& sorgID &"','"& rsObj(2) &"'); return false;"" class='ExcelDisplayLink'>"& rsObj1(0) &"</a>"
															else
																Response.Write "NA"
															end if
														end if
														rsObj1.Close
												%>
												</td>
												<td class="ExcelDisplayCell" align="left">
													<%
														if iUseable = 0 then
															Response.Write "Yes"
														else
															Response.Write "No"
														end if
													%>
												</td>
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
									<td align="center" class="ClearPixel" width="5px">
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
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td class="ActionCell">

													<input type="button" value="Create New" name="BtnNewParty" class="ActionButtonX" onclick="CreateNewParty()">
												</td>
											</tr>

										</table>
									</td>
									<td align="center" class="ClearPixel" width="5px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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
