<%@ Language=VBScript %>
<%	option explicit
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ContactsList.asp
	'Module Name				:	Accounts
	'Author Name				:	Kalaiselvi R
	'Created On					:	Sep 27,2011
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
Dim rsObj,rsTemp
Dim nSlNo
Dim iContactNumber,iStartRec,iEndRec,iPartyCode
Dim iPrevPage,iTotalPages,nPageCtr,iNextPage,iPageSize,iPageNo
Dim iTotalRecords
Dim sQuery,sParName,sCity,sSearch,sCallFrom

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")


sParName = Request("hParName")
sCity = Request("hCity")
sSearch = Request("hSearch")
sCallFrom = Request("CallFrom")

iPageSize = 20

iPageNo = trim(Request("hPage"))
if trim(iPageNo) = "" then iPageNo = 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Accounts</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "contactsList" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="ContactsList.asp">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hParName" value="<%=sParName%>">
	<input type="hidden" name="hCity" value="<%=sCity%>">

	<input type="hidden" name="hContactNo" value="">
	<input type="hidden" name="hSearch" value="<%=sSearch%>">
	<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">


	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Contact Master
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
														<table class="CollapseBand" cellspacing="0" cellpadding="0" class="BodyTable">
															<tr>
																<td valign="center"><a style="width: 1em; height: 1em;" title href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" >
																	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
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
																				<td class="MiddlePack">
																				</td>
																				<td class="MiddlePack" colspan="4">
																				</td>
																			</tr>
																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell">Contact Name</td>
																				<td class="FieldCellSub" colspan="3">
																					<Select name="selParSearchType" class="FormElem">
																						<Option value="SB" <%if sSearch="SB" then Response.Write "Selected"%> >Start With</Option>
																						<Option value="WA" <%if sSearch="WA" then Response.Write "Selected"%> >Any Where</Option>
																					</Select>
																					<input type="text" name="txtContactName" value="<%=sParName%>" class="FormElem">
																				</td>
																			</tr>
																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell">City</td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtCity" value="<%=sCity%>" class="FormElem">
																				</td>
																				<td class="FieldCell">&nbsp;</td>
																				<td class="FieldCellSub">
																					<input type="button" name="btnGo" value="  GO  " class="ActionButtonX" onClick="CheckSubmit()">
																				</td>
																			</tr>
																			<tr>
																				<td class="Middlepack" colspan="5">&nbsp;

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
												<td class="ExcelHeaderCell" align="center" ><a style="width: 1em; height: 1em;" title href="#" onclick="DelContact(); return false;" itms_state="0">
													<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Click here to delete the Party" width="15px" height="15px">
													</a>
												</td>
												<td class="ExcelHeaderCell">Contact Name
												</td>
												<td class="ExcelHeaderCell">Designation
												</td>
												<td class="ExcelHeaderCell">Contact Person For
												</td>
												<td class="ExcelHeaderCell">City
												</td>
												<td class="ExcelHeaderCell">Phone No
												</td>
												<!--
												<td class="ExcelHeaderCell" align="center" >PartyCode
												</td>
												-->
												<td class="ExcelHeaderCell">Status
												</td>
											</tr>

											<%
												sQuery = "Select ContactNumber,ContactName,Designation,ContactPersonFor,City,isNull(PhoneNos,''),isNull(Useable,0),isNull(PartyCode,0) from APP_M_Contacts where isNull(ContactName,'')<>'' "

												if trim(sParName)<>"" then
													if sSearch = "SB" then
														sQuery = sQuery & " and ContactName like '"& sParName &"%' "
													else
														sQuery = sQuery & " and ContactName like '%"& sParName &"%' "
													end if
												end if

												if trim(sCity)<>"" then
													sQuery = sQuery & " and City like '"& sCity & "%'"
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
													iContactNumber = rsObj(0)

													iPartyCode = ""
													if trim(rsObj(7)) <> "0" then
														sQuery = "Select OrgnPartyCode from APP_M_PartyMaster where PartyCode =" & rsObj(7) & ""
														' Response.Write sQuery
														with rsTemp
															.CursorLocation = 3
															.CursorType = 3
															.ActiveConnection = con
															.Source = sQuery
															.Open
														end with

														if not rsTemp.EOF then
															iPartyCode = rsTemp(0)
														end if
														rsTemp.Close
													end if 'if trim(rsObj(7)) <> "0" then
											%>

											<tr>
												<td class="ExcelSerial"><%=nSlNo%>
												</td>
												<td class="ExcelDisplayCell" width="10px">
													<input type="radio" name="radButton" value="<%=iContactNumber%>">
												</td>
												<td class="ExcelDisplayCell" align="left">
													<input type=hidden name="hPartyName<%=nSlNo%>" value="<%=rsObj(1)%>">
													<a href="#" onClick="ViewContactDeatils('<%=iContactNumber%>'); return false;" class="ExcelDisplayLink"><%=rsObj(1)%></a>
												</td>
												<td class="ExcelDisplayCell" align="left">
													<%=rsObj(2)%>
												</td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(3)%>
												</td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(4)%></td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(5)%></td>
												<!--<td class="ExcelDisplayCell" align="left"><%=iPartyCode%></td>-->
												<td class="ExcelDisplayCell" align="left">
												<%
													if trim(rsObj(6)) = "0"  then
														Response.Write "Active"
													else
														Response.Write "InActive"
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

													<input type="button" value="Create New" name="BtnNewParty" class="ActionButtonX" onclick="CreateNewContacts()">
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
