
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ActivityCreationMain.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	December 14 ,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/populate.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID = "OutData"><Root/></XML>
<XML ID=PRData></XML>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AdminRoleActivityCompat.js"></SCRIPT>
<SCRIPT>
ITMSAdminRoleActivityCompat.installActivityCreationMain();
</SCRIPT>
<% 	Dim sSql,sSql1,iCtr,sLoginID,sUnitID,sEmpValue,sEmpName,nKK,sActivityName,sPracticeCode
	Dim iCurrentPage,iTotalPage,lnPage,nSlNo,sDefValue,sProcessCode,sTempVal,nSelAppCode
	Dim sCondition,sCondition1,blnPrintProcessName,iTempCnt

	Const iPageSize = 20
	Dim objRs,objRs1,objRs2,rsMain

	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	Set objRs  = Server.CreateObject("ADODB.RecordSet")
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")
	Set objRs2 = Server.CreateObject("ADODB.RecordSet")

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	if iCurrentPage=0 then iCurrentPage=1

	if trim(sUnitID) = "" then	sUnitID = Session("organizationcode")

	sProcessCode  = Request("selProcess")
	sPracticeCode = Request("selActivityType")
	sActivityName = Request("hActivityName")

	nSelAppCode = cint(Request.Form("selApplication"))

	'Response.Write "<p>sProcessCode = "&sProcessCode & "  ++  " & nSelAppCode

	'IF sProcessCode <> "0"  and sProcessCode <> "" Then
	'	nSelAppCode = sProcessCode
	'End IF

	If nSelAppCode = "0" Then
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES"
			.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing

		if not objRs.EOF then
			nSelAppCode = objrs(0)
		End IF
		objrs.Close
	End IF
	'Response.Write "<p>nRoleID="& nSelAppCode

%>
</Head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="PopulatePractice('<%=nSelAppCode%>')">
<form method="POST" name="formname">

	<Input type=hidden name=hProcessName  value="<%=sProcessCode%>">
	<Input type=hidden name=hPracticeName value="<%=sPracticeCode%>">
	<Input type=hidden name=hActivityName value="<%=sActivityName%>">
	<input type=hidden name="hAppCode" value="<%=nSelAppCode%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0" >
		<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Activity Creation
		</td>
		</tr>

		<tr>
		<td valign="top">
			<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >

				<tr>
				<TD class=TabBody>
				<!--<td class="TabBodyWithTopLine"><div style="height:130px;">-->
					<table border="0" cellpadding="0" cellspacing="0" >
						<tr>
						<td align="center" colspan="3" class="MiddlePack" height="7" >
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						</td>
						</tr>

						<tr>
						<td align="center" width="5" class="ClearPixel">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						</td>
						<td valign="top" width="100%">
						<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
						<tr>
						<td>
						<div>
						<table class="CollapseBand" cellspacing="0" cellpadding="0" >
						<tr>
						<td valign="center">
						<a style="width: 1em; height: 1em;" title="" onclick="return Div_OnClick(idUnprocessed,event);" >
						<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
						</a>
						</td>
						<td valign="right" class="SubTitle">
						</td>
						</tr>
						</table>

						<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<tr>
						<td width="100%">
						<div id="idUnprocessed" style="display: none">
						<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">

						<tr>
							<td class="FieldCellSub">Unit Name</td>
							<td class="FieldCellSub" colspan="2">
							<select size="1" name="selUnitId" class="FormElem">
								<%populateUnits%>
							</select>
							</td>
						</tr>
						<tr>
							<td class="FieldCellSub">Process Name</td>
							<td class="FieldCellSub">
								<!--<input type = text name=TxtProcessName  class="Formelem">-->
								<select size="5" name="selProcess" class="FormElem" onChange="PopulatePractice(this.value)">
									<option value="0" selected>Select</option>
								<%	'Calling the Function to populate Applications / Process List
									'PopulateProcess
									with objRs
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
									.ActiveConnection = con
									.Open
								end with
								set objRs.ActiveConnection = nothing

								if not objRs.EOF then

									Do While Not objRs.EOF

										If cint(nSelAppCode) = cinT(objRs(0)) Then%>
											<option value=<%=trim(objRs(0))%> selected ><%=trim(trim(objRs(1)))%></option>
										<%Else%>
											<option value=<%=trim(objRs(0))%>><%=trim(trim(objRs(1)))%></option>
										<%End IF
									objRs.MoveNext
									Loop
								end if
								objRs.Close

								%>
								</select>
							</td>

							<td class="FieldCellSub">Practice Name</td>
							<td class="FieldCellSub">
								<!--<input type = text name=txtPracticeName  class="Formelem">-->
								<select size="5" name="selActivityType" class="FormElem" Multiple>
									<option value="S" selected>Select</option>
								</select>
							</td>
						</tr>

						<tr>
							<td class="FieldCellSub">Activity Name</td>
							<td class="FieldCellSub" colspan="2">
								<input type = text name=txtActivityName  class="Formelem">
							</td>
						</tr>

						<tr>
							<td class="FieldCellSub"></td>
							<td class="FieldCellSub" >
								<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
							</td>
							<td class="FieldCell" >
								<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
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
						<td>
						</td>
						<td valign="top" width="100%">
							<table border="0" cellspacing="1" class="ExcelTable" width="100%">

								<tr>
									<td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
									<!--<td class="ExcelHeaderCell" align="center" width="10">
										<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15"></a>
									</td>-->
									<td class="ExcelHeaderCell" align="center" >Mapped Activities</td>
								</tr>

							<%
							Dim sPrevAppName,sPrevProcessName,sPrevRoleID,iSpan,nNoOfActivity,nNoOfModules,nNoOfPractice
							Dim iSpan1
							sPrevAppName = ""
							sPrevProcessName = ""
							sPrevRoleID = cint("0")

							'sSql = " select Distinct R.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONCODE,A.APPLICATIONNAME,AP.PROCESSCODE,AP.PROCESSNAME,AP.ORDERNUMBER "&_
							'	   " From MS_ROLEACTIVITY R,Ms_ApplicationProcess AP,Ms_Applications A ,MS_ROLES M "&_
							'	   " where R.APPLICATIONCODE = AP.APPLICATIONCODE and A.APPLICATIONCODE = R.APPLICATIONCODE and R.RoleID = M.RoleID "


							'sSql = " select Distinct R.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONCODE,A.APPLICATIONNAME,AP.PROCESSCODE,AP.PROCESSNAME,AP.ORDERNUMBER "&_
							'	   " From MS_ROLEACTIVITY R,Ms_ApplicationProcess AP,Ms_Applications A ,MS_ROLES M "&_
							'	   " where R.APPLICATIONCODE = AP.APPLICATIONCODE and R.PROCESSCODE = AP.PROCESSCODE "&_
							'	   " and R.APPLICATIONCODE = A.APPLICATIONCODE and R.RoleID = M.RoleID "

							sSql = " select Distinct A.APPLICATIONCODE,A.APPLICATIONNAME,AP.PROCESSCODE,AP.PROCESSNAME,AP.ORDERNUMBER "&_
								   " From Ms_ApplicationProcess AP,Ms_Applications A where A.APPLICATIONCODE = AP.APPLICATIONCODE  "


							If nSelAppCode <> "0" Then
								sSql = sSql & " AND A.APPLICATIONCODE= "& nSelAppCode&" "
							End IF

							if Trim(sProcessCode) <> "S" and Trim(sProcessCode) <> "" then
								'sSql = sSql & " and A.APPLICATIONCODE = "& sProcessCode &" "
							End if

							if Trim(sPracticeCode) <> "S" and Trim(sPracticeCode) <> "" then
								sSql = sSql & " and AP.PROCESSCODE IN ("& sPracticeCode &") "
							End if

							sCondition = sSql

							sSql = sSql & "ORDER BY AP.PROCESSCODE "
							'Response.Write sSql

							With rsMain
								.ActiveConnection = Con
								.CursorLocation = 3
								.CursorType = 3
								.Source = sSql
								.Open
							End With

							Set rsMain.ActiveConnection = Nothing

							iCtr = 1

							If Not rsMain.EOF Then

								blnPrintProcessName = True

								Do While Not rsMain.EOF

									'sSql = " select Distinct M.ROLEID,M.ROLEDESCRIPTION,V.APPLICATIONNAME,V.PROCESSNAME,V.PROCESSCODE, "&_
									'	   " R.ACTIVITYCODE,V.ACTIVITYNAME,R.APPLICATIONCODE,A.ORDERNUMBER from MS_ROLEACTIVITY R,VWACTIVITY V,MS_ROLES M,"&_
									'	   " VWUSERACTIVITY U, Ms_ApplicationProcess A where R.APPLICATIONCODE = V.APPLICATIONCODE and R.ACTIVITYCODE = V.ACTIVITYCODE "&_
									'	   " and R.ROLEID = M.ROLEID and R.ACTIVITYCODE = U.ACTIVITYCODE and V.PROCESSCODE=U.PROCESSCODE "&_
									'	   " and A.APPLICATIONCODE =V.APPLICATIONCODE and A.PROCESSCODE = V.PROCESSCODE "&_
									'	   " and A.PROCESSCODE = "& rsMain(4) &" and A.APPLICATIONCODE = "& rsMain(2)&" and R.ROLEID = "& rsMain(0)&" "

									If 1 = 2 Then
									sSql = " select Distinct A.APPLICATIONNAME,V.PROCESSNAME,V.PROCESSCODE, "&_
										   " V.ACTIVITYCODE,V.ACTIVITYNAME,V.APPLICATIONCODE,V.ORDERNUMBER from "&_
										   " VWUSERACTIVITY V, Ms_Applications A where V.APPLICATIONCODE = A.APPLICATIONCODE  "&_
										   " and V.PROCESSCODE = "& rsMain(2) &" and V.APPLICATIONCODE = "& rsMain(0)&" "

									if Trim(sActivityName) <> "" then
										sSql = sSql & " and V.ACTIVITYNAME like '%"& sActivityName &"%'"
									End if
									sCondition1 = sSql
									sSql = sSql & " ORDER BY V.ORDERNUMBER,V.PROCESSCODE,V.ACTIVITYCODE "

									'Response.Write ssql
									End IF

									sSql = " select Distinct A.APPLICATIONNAME,AP.PROCESSNAME,AA.PROCESSCODE,AA.ACTIVITYCODE,AA.ACTIVITYNAME, "&_
										   " AA.APPLICATIONCODE,AP.ORDERNUMBER from Ms_Applications A, Ms_ApplicationActivity AA , Ms_ApplicationProcess AP "&_
										   " where A.APPLICATIONCODE = AA.APPLICATIONCODE and AA.APPLICATIONCODE = AP.APPLICATIONCODE "&_
									       " and AA.PROCESSCODE = AP.PROCESSCODE and AA.PROCESSCODE = "& rsMain(2) &" and AA.APPLICATIONCODE = "& rsMain(0)&"  "

									if Trim(sActivityName) <> "" then
										sSql = sSql & " and AA.ACTIVITYNAME like '%"& sActivityName &"%'"
									End if

									sSql = sSql & " ORDER BY AP.ORDERNUMBER,AA.PROCESSCODE,AA.ACTIVITYCODE "
									sCondition1 = sSql

									With objRs
										.ActiveConnection = Con
										.CursorLocation = 3
										.CursorType = 3
										.Source = sSql
										.Open
									End With

									Set objRs.ActiveConnection = Nothing

									nSlNo = 1


									If not objRs.EOF then
										objRs.PageSize = iPageSize
										objRs.AbsolutePage = iCurrentPage
										iTotalPage = objRs.PageCount
									Elseif 1 = 2 Then
										sTempVal = ""
										'sTempVal = objrs(0) &":"& Trim(objrs(7)) &":"& Trim(objrs(2)) &":"& Trim(objrs(4)) &":"& Trim(objrs(3))

										'sTempVal = rsMain(0) &":"& rsMain(2) &":"& rsMain(3) &":"& Trim(rsMain(4)) &":"& Trim(rsMain(5))

										sTempVal = "1" &":"& rsMain(0) &":"& rsMain(1) &":"& Trim(rsMain(2)) &":"& Trim(rsMain(3))

										'ROLEID,APPLICATIONCODE,APPLICATIONNAME,PROCESSCODE,PROCESSNAME
										%>
										<tr>
											<td class=ExcelDisplayCell align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><%=rsMain(3)%></B>
												&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('PRACTICE','<%=sTempVal%>'); return false;">EDIT</a>]
												&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowRoleActivityMap('ADD','<%=sTempVal%>'); return false;">ADD ACTIVITY</a>]
											</td>
										</tr>
										<%
									End if 'If not objRs.EOF then

									nKK = 1
									nNoOfModules  = 0
									nNoOfPractice = 0
									nNoOfActivity = 0

									Do while not objRs.EOF and nSlNo  <= objRs.PageSize

										blnPrintProcessName = False

										sTempVal = ""
										%>

										<%	'sTempVal = objrs(0) &":"& Trim(objrs(7)) &":"& Trim(objrs(2)) &":"& Trim(objrs(4)) &":"& Trim(objrs(3))

											'sTempVal = "1" &":"& Trim(objrs(3)) &":"& Trim(objrs(0)) &":"& Trim(objrs(2)) &":"& Trim(objrs(1))


											sTempVal = "" &":"& Trim(objrs(5)) &":"& Trim(objrs(0)) &":"& Trim(objrs(2)) &":"& Replace(Trim(objrs(1)),"&","and") 'Trim(objrs(1))

											iSpan = 200

											'Process Name
											IF sPrevAppName = "" or sPrevAppName <> objrs(0) then
												sPrevAppName = objrs(0)%>
												<tr>
												<td class="ExcelSerial" align="center" rowspan=<%=iSpan%> valign=top><%=nSlNo%></td>
												<td class=ExcelDisplayCell align=left><B><%=UCase(objrs(0))%></B>
													&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('PROCESS','<%=sTempVal%>'); return false;">EDIT</a>]
													&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('ADDPRACTICE','<%=sTempVal%>'); return false;">ADD PRACTICE</a>]
													&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="DelActivity(); return false;">DELETE ACTIVITY</a>]
												</td>

											<%End IF

											'Practice Name
											If sPrevProcessName = "" or sPrevProcessName <> objrs(1) Then
												sPrevProcessName = objrs(1)
											%>
												<tr>
													<td class=ExcelDisplayCell align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><%=objrs(1)%></B>
														&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('PRACTICE','<%=sTempVal%>'); return false;">EDIT</a>]
														&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowRoleActivityMap('ADD','<%=sTempVal%>'); return false;">ADD ACTIVITY</a>]
													</td>
												</tr>
											<%
											End IF


											'ActivityName
											'sTempVal = sTempVal & ":" & Trim(objrs(5)) &":"& Trim(objrs(6))
											'sTempVal = sTempVal & ":" & Trim(objrs(3)) &":"& Trim(objrs(4))
											sTempVal = sTempVal & ":" & Trim(objrs(3)) &":"& ""

											'Response.Write Trim(objrs(5)) &":"& Trim(objrs(6))&":"& objrs(7) & ":" & objRs(4) &":"& objrs(0)
											
											sSql = "Select Count(*) from Ms_ApplicationActivityTemplates where ApplicationCode = "& objrs(5) &" and ProcessCode ="& objrs(2) &" and ActivityCode = "& objrs(3)
											objrs2.open sSql,con
											if not objrs2.eof then
											    iTempCnt = objrs2(0)
											end if
											objrs2.close
											
											%>
											<tr>
												<td class=ExcelDisplayCell align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<input type="checkbox" name="Chkbox<%=iCtr%>" value="<%Response.Write Trim(objrs(3)) &":"& Trim(objrs(4))&":"& objrs(5) & ":" & objRs(2) &":"& "1"%>" <%if Cint(iTempCnt)>1 then Response.write "Disabled" %> >
												<a href="#" class="ExcelDisplayLink" onclick="ShowRoleActivityMap('EDT','<%=sTempVal%>'); return false;"><%=objrs(4)%>&nbsp;&nbsp;
												</a>
												<%
												if cint(iTempCnt)>1 then
												    Response.write "[Available Templates : "& iTempCnt &"]"
												end if
												
												%>
												</td>
											</tr>
										</tr>

										<%
											iCtr = iCtr + 1
											nSlNo = nSlNo + 1

										objRs.MoveNext
									Loop
									objRs.Close

									rsMain.MoveNext
								Loop
							Else
								Dim sApplicationName,nApplicationCode
								nSlNo = "1"
								with objRs
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME FROM MS_APPLICATIONS WHERE APPLICATIONCODE = "& nSelAppCode &" "
									.ActiveConnection = con
									.Open
								end with
								set objRs.ActiveConnection = nothing

								if not objRs.EOF then
									nApplicationCode = objRs(0)
									sApplicationName = objrs(1)
								End IF
								objrs.Close
								sTempVal = "NEW"& ":" & nApplicationCode & ":" & sApplicationName
								%>
								<tr>
									<td class="ExcelSerial" align="center" rowspan=<%=iSpan%> valign=top><%=nSlNo%></td>
									<td class=ExcelDisplayCell align=left><B><%=UCase(sApplicationName)%></B>
										&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('PROCESS','<%=sTempVal%>'); return false;">EDIT</a>]
										&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('ADDPRACTICE','<%=sTempVal%>'); return false;">ADD PRACTICE</a>]
										&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="DelActivity(); return false;">DELETE ACTIVITY</a>]
									</td>
								</tr>
								<%
							End IF 'If Not rsMain.EOF Then
							rsMain.Close


							'Process with out Mapped Activities
							With rsMain
								.ActiveConnection = Con
								.CursorLocation = 3
								.CursorType = 3
								.Source = sCondition
								.Open
							End With

							Set rsMain.ActiveConnection = Nothing

							If blnPrintProcessName Then
								sTempVal = "1" &":"& rsMain(0) &":"& rsMain(1) &":"& Trim(rsMain(2)) &":"& Trim(rsMain(3))
								%>
								<tr>
									<td class="ExcelSerial" align="center" rowspan=<%=iSpan%> valign=top><%=nSlNo%></td>
									<td class=ExcelDisplayCell align=left><B><%=UCase(rsMain(1))%></B>
										&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('PROCESS','<%=sTempVal%>'); return false;">EDIT</a>]
										&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('ADDPRACTICE','<%=sTempVal%>'); return false;">ADD PRACTICE</a>]
										&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="DelActivity(); return false;">DELETE ACTIVITY</a>]
									</td>
								</tr>
								<%
							End IF

							Do While Not rsMain.EOF

									sSql = " select Distinct A.APPLICATIONNAME,AP.PROCESSNAME,AA.PROCESSCODE,AA.ACTIVITYCODE,AA.ACTIVITYNAME, "&_
										   " AA.APPLICATIONCODE,AP.ORDERNUMBER from Ms_Applications A, Ms_ApplicationActivity AA , Ms_ApplicationProcess AP "&_
										   " where A.APPLICATIONCODE = AA.APPLICATIONCODE and AA.APPLICATIONCODE = AP.APPLICATIONCODE "&_
									       " and AA.PROCESSCODE = AP.PROCESSCODE and AA.PROCESSCODE = "& rsMain(2) &" and AA.APPLICATIONCODE = "& rsMain(0)&"  "

									if Trim(sActivityName) <> "" then
										sSql = sSql & " and AA.ACTIVITYNAME like '%"& sActivityName &"%'"
									End if

									sSql = sSql & " ORDER BY AP.ORDERNUMBER,AA.PROCESSCODE,AA.ACTIVITYCODE "

									With objRs
										.ActiveConnection = Con
										.CursorLocation = 3
										.CursorType = 3
										.Source = sSql
										.Open
									End With

									Set objRs.ActiveConnection = Nothing

									If Not objrs.EOF Then
									Else

										sTempVal = "1" &":"& rsMain(0) &":"& rsMain(1) &":"& Trim(rsMain(2)) &":"& Trim(rsMain(3))
										'ROLEID,APPLICATIONCODE,APPLICATIONNAME,PROCESSCODE,PROCESSNAME

										%>
										<tr>
											<%If blnPrintProcessName Then%>
												<td class="ExcelSerial" align="center" rowspan=<%=iSpan%> valign=top>&nbsp;</td>
											<%End IF%>
											<td class=ExcelDisplayCell align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><%=rsMain(3)%></B>
												&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowDetails('PRACTICE','<%=sTempVal%>'); return false;">EDIT</a>]
												&nbsp;&nbsp;[<a href="#" class="ExcelDisplayLink" onclick="ShowRoleActivityMap('ADD','<%=sTempVal%>'); return false;">ADD ACTIVITY</a>]
											</td>
										</tr>
										<%
									End IF
									objrs.Close

								rsMain.MoveNext
							Loop
							rsMain.Close


							'Response.Write "<p>iCtr="&iCtr

							%>
							<input type="hidden" name="hCnt" value="<%=iCtr-1%>">

							</table>
							</div>
							</td>
								<td>
								</td>
                            </tr>

							<tr>
								<td colspan="3" class="MiddlePack">
								</td>
							</tr>

							<tr>
							<td align="center" width="5" class="ClearPixel">
							</td>
							<td valign="top" align="right">
							<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
							<!--<input type=hidden name="hCnt" value=<%'=iCtr -1  %>>-->
							<input type=hidden name="hPageSelection" value="0">

							<Select class="FormElem" Name="selApplication" onchange="AssaignApplicationCode(this.value)">
								<% 'PopulateRole %>
								<%
								with objRs
									.CursorLocation = 3
									.CursorType = 3
									'.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES"
									.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
									.ActiveConnection = con
									.Open
								end with
								set objRs.ActiveConnection = nothing

								if not objRs.EOF then

									Do While Not objRs.EOF

										If cint(nSelAppCode) = cinT(objRs(0)) Then%>
											<option value=<%=trim(objRs(0))%> selected ><%=trim(trim(objRs(1)))%></option>
										<%Else%>
											<option value=<%=trim(objRs(0))%>><%=trim(trim(objRs(1)))%></option>
										<%End IF
									objRs.MoveNext
									Loop
								end if
								objRs.Close
								%>
							</Select>


							<%	If iTotalPage >= 2 Then
							if iCurrentPage = 1 then
							%>
							<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
							<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
							<%		else%>
							<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
							<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
							<%		end if	%>
							<SELECT class="FormElem" onChange="Paginate(this.options[this.selectedIndex].value)" id=select1 name=select1>
							<%
							For lnPage = 1 To iTotalPage
							If lnPage = iCurrentPage Then
							%>
							<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
							<%		else	%>
							<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
							<%		end if
							next
							%>
							</SELECT>
							<%
							if iCurrentPage = iTotalPage then
							%>
							<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
							<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

							<%		else	%>
							<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
							<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotalPage%>')" id=button8 name=button8>
							<%		end if
							End If
							%>
							</td>
							<td align="center" class="ClearPixel" width="5">
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							</tr>

							<tr>
							<td>
								<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							<td valign="top">

							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td valign="middle" class="ActionCell">
									<p align="center">
										<Input type="button" value="Done" name="ButOpt" class="ActionButton" tabindex="3">
										<!--<Input type="button" value="Edit" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction('EDT')" onclick="GotoAction('CRN')">-->
									</td>
								</tr>
							</table>

							</td>
								<td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td colspan="3" class="BottomPack">
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
<%
Function populateUnits()
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM VWITEMLIST) ORDER BY ORGANIZATIONUNITID"
		.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	set sUnitID = dcrs(0)
	set sUnitName = dcrs(1)
	If not dcrs.EOF then
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
	end if
	dcrs.Close

End Function

' Function to populate Applications / Process
Function PopulateProcess()
		' Declaration of variables
		Dim dcrs
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			Do While Not dcrs.EOF

				Response.Write "<option value="""&trim(dcrs(0))&""">"&trim(trim(dcrs(1)))&"</option>" & vbCrLf

			dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
