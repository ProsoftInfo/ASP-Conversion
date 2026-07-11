<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLAccHeadGrid.asp
	'Module Name				:	Accounts
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 11,2010
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
<%

Dim rsObj,rsObj1,rsObj2,rsConn
Dim iCatagoryCode,iChildCount,iCnt,iAccountGroupCode,CurPoss,iAccGroup1,iAccGroup2
Dim iAccParentGroup,iParantCursor,nSpace,iSpCnt,iAccRowCount
Dim bSummaryPosting,bEligibleForTDS,bPartySubLed,bContra
Dim sQuery,sSpace,sOrgID
Dim iSumPosting,iETDS,iSubLed,iContra

Set rsConn = Server.CreateObject("ADODB.Connection")
Set rsObj  = Server.CreateObject("ADODB.Recordset")
Set rsObj1 = Server.CreateObject("ADODB.Recordset")
Set rsObj2 = Server.CreateObject("ADODB.Recordset")

iCatagoryCode =Request.QueryString("Category")

if trim(iCatagoryCode)="" then
	iCatagoryCode = Request("hCategory")
end if

if iCatagoryCode="" then
	iCatagoryCode = "01"
end if

'sOrgID = Request.Form("hOrgCode")
sOrgID = Session("organizationcode")
iSubLed = Request.Form("hSubLed")
iContra = Request.Form("hContra")
iETDS = Request.Form("hTDS")
iSumPosting  = Request.Form("hSumPosting")
if trim(sOrgID)="" then
'	sOrgID = "010101"
end if

nSpace = 5
iAccRowCount = 0
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Accounts</title>
<XML id="OutData">
<Root/>
</XML>
<XML id="AccHeadData">
	<Root/>
</XML>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "glAccountHeadGrid" };
</script>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="init()">
<form method="POST" name="formname" action="">
	<input type="hidden" name="hCategory" value="<%=iCatagoryCode%>">
	<input type="hidden" name="hHeadValue" value="">
	<input type="hidden" name="hHeadName" value="">
	<input type="hidden" name="GCode" value="">
	<input type="hidden" name="GName" value="">
	<input type="hidden" name="selUnitId" value="">
	<input type="hidden" name="hOrgCode" value="<%=sOrgID%>">
	<input type="hidden" name="hSubLed" value="<%=iSubLed%>">
	<input type="hidden" name="hContra" value="<%=iContra%>">
	<input type="hidden" name="hTDS" value="<%=iETDS%>">
	<input type="hidden" name="hSumPosting" value="<%=iSumPosting%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				GL Account Heads
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
																	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
																	</a>
																</td>
																<td valign="center" class="SubTitle">&nbsp;&nbsp;
																<%
																	sQuery = "Select CategoryCode,CategoryName from Acc_M_AccountCategory"
																	rsObj1.Open sQuery,Con
																	if not rsObj1.EOF then
																		iCnt = 0
																		do while not rsObj1.EOF
																			if trim(iCatagoryCode)=trim(rsObj1(0)) then
																				Response.Write "<input type=radio name=radAccHead value="&rsObj1(0)&" checked onClick=AccountsCategory("& iCnt &")>"
																			else
																				Response.Write "<input type=radio name=radAccHead value="&rsObj1(0)&" onClick=AccountsCategory("& iCnt &")>"
																			end if
																			Response.Write rsObj1(1)
																			Response.Write "&nbsp;&nbsp;"
																			iCnt = iCnt + 1
																			rsObj1.MoveNext
																		loop
																	end if
																	rsObj1.Close
																%>
																</td>
															</tr>

														</table>
														<table border="0" cellpadding="0" cellspacing="0">
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="width: 575px; display: none">
																		<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
																			<tr>
																				<td class="MiddlePack">
																				</td>
																				<td class="MiddlePack" colspan="4">
																				</td>
																			</tr>
																			<!--<tr>
																				<td width=100></td>
																				<td class="FieldCell">Organisation</td>
																				<td class="FieldCellSub">
																					<Select name="selOrganisaion" class="FormElem">
																						<option value="0">Not Applicable</option>
																						<%
																							sQuery = "Select OUDefinitionID,OrgUnitDescription from DCS_OrganizationUnitDefinitions where Len(OUDefinitionID)>4"
																							rsObj.Open sQuery,con
																							if not rsObj.EOF then
																								do while not rsObj.EOF
																										Response.Write "<option value="& rsObj(0) &">"& rsObj(1) &"</option>"
																									rsObj.MoveNext
																								loop
																							end if
																							rsObj.Close
																						%>
																					</Select>
																				</td>
																			</tr>-->
																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell">GL Group</td>
																				<td class="FieldCellSub">
																					<input type=text name="txtGLGroup" value ="" class="FormElem"> &nbsp;
																					<img src="../../assets/images/iTMS%20icons/Entryicon.gif" onClick="">
																				</td>
																			</tr>

																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell">Account Head</td>
																				<td class="FieldCellSub">
																					<input type=text name="txtAccountHead" value ="" class="FormElem">
																				</td>
																			</tr>

																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell" colspan="2">
																				<input type=checkbox name=chkSummary class="FormElem">Summary Posting &nbsp;&nbsp;&nbsp;
																				<input type=checkbox name=chkParty class="FormElem">Party Control
																				</td>
																			</tr>

																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell" colspan="2">
																				<input type=checkbox name=chkContra class="FormElem">Eligible For Contra &nbsp;&nbsp;
																				<input type=checkbox name=chkTDS class="FormElem">Eligible For TDS &nbsp;&nbsp;
																				<input type=button name=btnGO value="GO" class="ActionButton"  onClick="CheckSubmit()"></td>
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
										<div class="frmbody" id="frm4" style="width: 585; height:300;">
										<table border="0" cellspacing="1px" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="Left" colspan=8>GL Group</td>
											</tr>
											<tr>
												<td></td>
												<td colspan=7 width=100%>
													<table border=0 cellspacing="1px" width="100%">
														<td class="ExcelHeaderCell" align="center" width="25px">
														<a style="width: 1em; height: 1em;" title href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
															<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Expands this section for more search criteria." width="15px" height="15px">
															</a>
														</td>
														<td class="ExcelHeaderCell" width="13%" >Code
														</td>
														<td class="ExcelHeaderCell" width="200px" >Name
														</td>
														<td class="ExcelHeaderCell" width="30px" >Summary
														</td>
														<td class="ExcelHeaderCell" width="30px" >Party
														</td>
														<td class="ExcelHeaderCell" width="30px" >Contra
														</td>
														<td class="ExcelHeaderCell" width="30px" > TDS
														</td>
													</table>
												</td>
											</tr>

											<%
												sQuery = "Select AccountsGroupCode,AccountsGroupName,ChildCount,GroupCategory,AccountsParentGroup from ACC_M_AccountGroups where AccountsGroupCode = AccountsParentGroup and GroupCategory = '"& iCatagoryCode &"'"
												with rsObj1
													.CursorLocation = 3
													.CursorType = 3
													.ActiveConnection = con
													.source = sQuery
													.Open
												end with
												do while not rsObj1.EOF
													iAccParentGroup =  rsObj1(0)

														set rsConn = con
														rsConn.CursorLocation = 3
														sQuery = "Select AccountsGroupCode,AccountsGroupName,ChildCount,GroupCategory,AccountsParentGroup from ACC_M_AccountGroups "
														set rsObj = rsConn.execute(sQuery)
														set rsObj.ActiveConnection = Nothing

												'		rsConn.close
														rsObj.Filter = "AccountsGroupCode ='"& iAccParentGroup &"'"
														set rsConn = Nothing

														do while not rsObj.EOF
															CurPoss = rsObj.AbsolutePosition

															iAccGroup1 = rsObj(0)
															iChildCount = rsObj(2)
															Response.Write "<tr><td class=ExcelDisplayCell colspan=8>"&trim(rsObj(1))&"</td></tr>"

																	if trim(sOrgID)<>"0" and trim(sOrgID)<>"" then
																		sQuery = "Select AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
																		         "SummaryPosting,OUDefinitionID,GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vwOrgGLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup1 &"'  and OUDefinitionID = '"& sOrgID &"' "
																	else
																		sQuery = "Select Distinct AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
																				"SummaryPosting,'0',GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vw_GLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup1 & "'  "
																	end if 'if trim(sOrgID)<>"" then

																	if iSubLed = "1" then sQuery = sQuery & " and SubLedger = 1 "

																	if iContra = "1" then sQuery = sQuery & " and EligibleForContras = 1 "

																	if iETDS = "1" then sQuery = sQuery & " and EligibleForTDS = 1 "

																	if iSumPosting = "1" then sQuery = sQuery & " and  SummaryPosting = 1 "

																	sQuery= sQuery & " Order By AccountHead"
															'Response.Write sQuery
															rsObj2.Open sQuery,con
															if not rsObj2.EOF then
																Response.Write "<tr>"
																Response.Write "<td>"& sSpace &"</td>"
																Response.Write "<td colspan=7 class=ExcelTable><table width='100%' border=0 cellspacing=1>"
																do while not rsObj2.EOF
																	bSummaryPosting = rsObj2(5)
																	bEligibleForTDS = rsObj2(4)
																	bPartySubLed = rsObj2(9)
																	bContra = rsObj2(10)

																	iAccRowCount = iAccRowCount + 1

																		Response.Write "<tr>"& vbCrLf
																		Response.Write "<td width='10' class=ExcelDisplayCell><input type=radio name=radAccount value="& rsObj2(0) &"></td>" & vbCrLf
																		Response.Write "<td width='15%' class=ExcelDisplayCell><font color=red>"&trim(rsObj2(1))&"</font></td>" & vbCrLf
																		Response.Write "<td width='200' class=ExcelDisplayCell><font color=red><a href='#' onClick=""EditAccHead('"& rsObj2(6) &"','"& rsObj2(0) &"','"& rsObj2(2) &"','"& rsObj2(7) &"','"& rsObj2(8) &"'); return false;"" class=ExcelDisplayLink>"&rsObj2(2)&"</a></font></td>" & vbCrLf
																		if bSummaryPosting = "1" then
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td>" & vbCrLf
																		else
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td>"& vbCrLf
																		end if

																		if bPartySubLed = "1" then
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td>" & vbCrLf
																		else
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td>"& vbCrLf
																		end if

																		if bContra  = "1" then
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td>" & vbCrLf
																		else
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td>"& vbCrLf
																		end if

																		if bEligibleForTDS ="1" then
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td></tr>"& vbCrLf
																		else
																			Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td></tr>"& vbCrLf
																		end if
																	rsObj2.MoveNext
																loop
																Response.Write "</table></td></tr>"
															else
																if iChildCount="0" then
																		Response.Write "<tr><td align=center colspan=8 class=ExcelDisplayCell><font color=red>No Records Found</font></td></tr>"& vbCrLf
																end if
															end if
															rsObj2.Close

															For iCnt = 1 to iChildCount
																iAccountGroupCode = iAccGroup1 & Right("0"&cstr(iCnt),2)
																GLAccountHead iAccountGroupCode,2
																set iAccountGroupCode = Nothing
															Next

															rsObj.Filter = "AccountsGroupCode='"& iAccParentGroup &"'"
															rsObj.AbsolutePosition  = CurPoss
															rsObj.MoveNext
														loop

													rsObj1.MoveNext
												loop

												set rsObj = Nothing

											%>
											<input type="hidden" name="hCnt" value="<%=iAccRowCount%>" >
										</table>
										</div>
									</td>
									<td align="center" class="ClearPixel" width="5px">
									</td>
								</tr>

								<tr>
									<td align="center" class="MiddlePack" colspan="3">
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

													<input type="button" value="Create New" name="BtnNewParty" class="ActionButtonX" onclick="CreateNewHead()">
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
<%
Sub GLAccountHead(iAccGroupCode,Level)
	Dim iChildCnt,iAccGrpCode,iCurPoss,iCount
	rsObj.Filter = "AccountsGroupCode='"&iAccGroupCode &"'"
	do while not rsObj.EOF
		iCurPoss = rsObj.AbsolutePosition
		iAccGroup2 = rsObj(0)
		sSpace = ""

		For iSpCnt = 1 to (nSpace*Level)
			sSpace = sSpace & "&nbsp;"
		Next

		iChildCnt = rsObj(2)

		Response.Write "<tr>"
		Response.Write "<td colspan=8 class=ExcelDisplayCell>"& sSpace & rsObj(1) &"</td></tr>"

		'Response.Write "iChildCnt = "& iChildCnt
		if trim(sOrgID)<>"0" and trim(sOrgID)<>"" then
			sQuery = "Select AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
						"SummaryPosting,OUDefinitionID,GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vwOrgGLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup2 & "'  and OUDefinitionID = '"& sOrgID &"' "
		else
			sQuery = "Select Distinct AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
						 "SummaryPosting,'0',GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vw_GLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup2 & "' "
		end if 'if trim(sOrgID)<>"" then

		if iSubLed = "1" then sQuery = sQuery & " and SubLedger = 1 "

		if iContra = "1" then sQuery = sQuery & " and EligibleForContras = 1 "

		if iETDS = "1" then sQuery = sQuery & " and EligibleForTDS = 1 "

		if iSumPosting = "1" then sQuery = sQuery & " and  SummaryPosting = 1 "

		sQuery= sQuery & " Order By AccountHead"

		'Response.Write sQuery
		rsObj2.Open sQuery,con
		if not rsObj2.EOF then
			Response.Write "<tr>"
			Response.Write "<td>"& sSpace &"</td>"
			Response.Write "<td colspan=7 class=ExcelTable><table width='100%' border=0 cellspacing=1>"
			do while not rsObj2.EOF
				bSummaryPosting = rsObj2(5)
				bEligibleForTDS = rsObj2(4)
				bPartySubLed = rsObj2(9)
				bContra = rsObj2(10)
				iAccRowCount = iAccRowCount + 1

					Response.Write "<tr>"& vbCrLf
					Response.Write "<td class=ExcelDisplayCell width=10><input type=radio name=radAccount value="& rsObj2(0) &"></td>" & vbCrLf
					Response.Write "<td class=ExcelDisplayCell width='15%' ><font color=red>"&trim(rsObj2(1))&"</font></td>"& vbCrLf
					Response.Write "<td class=ExcelDisplayCell width='200' ><font color=red><a href='#' onClick=""EditAccHead('"& rsObj2(6) &"','"& rsObj2(0) &"','"& rsObj2(2) &"','"& rsObj2(7) &"','"& rsObj2(8) &"'); return false;"" class=ExcelDisplayLink>"&rsObj2(2)&"</a></font></td>"& vbCrLf
					if bSummaryPosting = "1" then
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td>"& vbCrLf
					else
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td>"& vbCrLf
					end if
					if bPartySubLed = "1" then
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td>" & vbCrLf
					else
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td>"& vbCrLf
					end if

					if bContra  = "1" then
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td>" & vbCrLf
					else
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td>"& vbCrLf
					end if
					if bEligibleForTDS ="1" then
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>Yes</font></td></tr>"& vbCrLf
					else
						Response.Write "<td align=center width='40' class=ExcelDisplayCell><font color=red>No</font></td></tr>"& vbCrLf
					end if
				rsObj2.MoveNext
			loop
			Response.Write "</table></td></tr>"
		else
			if iChildCnt="0" then
					Response.Write "<tr><td align=center colspan=8 class=ExcelDisplayCell><font color=red>No Records Found</font></td></tr>"& vbCrLf
			end if
		end if
		rsObj2.Close



		For iCount = 1 to iChildCnt
			iAccGrpCode = iAccGroup2 &right("0"&cstr(iCount),2)
			GLAccountHead iAccGrpCode,Level+1
		Next
		rsObj.Filter = "AccountsGroupCode='"& iAccGroupCode &"'"
		rsObj.AbsolutePosition = iCurPoss
		rsObj.MoveNext
	loop
End Sub
%>
