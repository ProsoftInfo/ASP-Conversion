<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppActivityRole.asp
	'Module Name				:	Admin (Activity Creation Amendment)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	December 17, 2010
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<%
	Dim nEmployeeNo,sEmployeeName,sProcessCode,sPracticeCode,sSql,iCtr,sActivityName
	Dim iCurrentPage,iTotalPage,lnPage,nSlNo,sDefValue,sTempVal,nSelAppCode,nKK

	nEmployeeNo = Request.QueryString("EmpNo")

	Const iPageSize = 20
	Dim objRs,objRs1,rsRole,rsMain

	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	Set objRs  = Server.CreateObject("ADODB.RecordSet")
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")
	Set rsRole = Server.CreateObject("ADODB.RecordSet")

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	if iCurrentPage=0 then iCurrentPage=1

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "Select (isNull(UserName,'')+isNull(MiddleName,'')+isNull(LastName,'')) From DCS_User where InternalUserID = "& nEmployeeNo &" "
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing

	If not objRs.EOF then
		sEmployeeName = objRs(0)
	End IF
	objRs.close

	sProcessCode  = Request("selProcess")
	sPracticeCode = Request("selActivityType")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Practice Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"></XML>
<XML ID="SelectedData"><Root/></XML>
<XML ID=PRData></XML>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AppActivityRoleCompat.js"></SCRIPT>
<SCRIPT>
ITMSAppActivityRoleCompat.installMain();
</SCRIPT>
</head>
<!--<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="populateList('<%=nEmployeeNo%>','1')">-->
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<form method="POST" name="formname">

	<Input type="hidden" name="hEmpID" value="<%=nEmployeeNo%>">
	<Input type="hidden" name="hItemRows" value="">
	<Input type=hidden name=hProcessCode  value="<%=sProcessCode%>">
	<Input type=hidden name=hPracticeCode value="<%=sPracticeCode%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0" >
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">User Activity
			</td>
		</tr>

		<tr>
		<td valign="top">
			<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				<tr>
				<td height="1" valign="bottom" class="TabCellEnd">&nbsp;
				</td>
                </tr>

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
							<td>
							</td>
							<td valign="top" width="100%">
							</td>

							<td>
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
						<table class="CollapseBand" cellspacing="0" cellpadding="0" >
						<tr>
						<td valign="center">
						<a style="width: 1em; height: 1em;" title="" onclick="return Div_OnClick(idUnprocessed,event);" >
						<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
						</a>
						</td>
						<td valign="right" class="SubTitle"><B>USER NAME : <%=sEmployeeName%></B>
						</td>
						</tr>
						</table>

						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
						<td width="100%">
						<div id="idUnprocessed" style="width: 575; display: none">
						<table cellpadding="0" cellspacing="0">

						<tr>
							<td class="FieldCellSub">Process Name</td>
							<td class="FieldCellSub">
								<!--<input type = text name=TxtProcessName  class="Formelem">-->
								<select size="5" name="selProcess" class="FormElem" onChange="PopulatePractice(this.value)">
									<option value="0" selected>Select</option>
								<%
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
											%>
												<option value=<%=trim(objRs(0))%>><%=trim(trim(objRs(1)))%></option>
											<%
										objRs.MoveNext
										Loop
									end if
									objRs.Close

								%>
								</select>
							</td>

							<td class="FieldCellSub">Practice Name</td>
							<td class="FieldCellSub">
								<select size="5" name="selActivityType" class="FormElem" Multiple>
									<option value="S" selected>Select</option>
								</select>
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

						<tr>
						<td align="center" class="MiddlePack">
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
							<tR>
								<td class="ExcelHeaderCell" align="center" width=10>S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="20">
									<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record"  onclick="DeleteItem()" width="15" height="15"></a>
								</td>
								<td class="ExcelHeaderCell" align="center" >Activity</td>
							</tr>
							<%
							Dim sPrevAppName,sPrevProcessName,sPrevRoleID,iSpan,nNoOfActivity,nNoOfModules,nNoOfPractice
							Dim iSpan1

							sPrevAppName = ""
							sPrevProcessName = ""
							sPrevRoleID = cint("0")
							
							sSql = "Select Distinct M.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONNAME,AP.PROCESSNAME,R.PROCESSCODE, "&_
											   " R.ACTIVITYCODE,AA.ACTIVITYNAME,R.APPLICATIONCODE,AAT.ActivityTemplateNo,AAT.ActivityTemplateName "&_
											   " from MS_ROLES M,Ms_Applications A,MS_ROLEACTIVITY R,Ms_ApplicationProcess AP,Ms_ApplicationActivity AA,Ms_ApplicationActivityTemplates AAT "&_
											   " where M.RoleID = R.ROleID and A.ApplicationCode = R.ApplicationCode and R.ApplicationCode = AP.ApplicationCode and "&_
											   " R.ProcessCode = AP.Processcode and A.ApplicationCode = AA.ApplicationCode and AP.ProcessCode= AA.ProcessCode "&_
											   " and R.ActivityCode = AA.ActivityCode and AA.ApplicationCode = AAT.APplicationCode and AA.ProcessCode = AAT.ProcessCode and "&_
											   " AA.ActivityCode = AAT.ActivityCode and R.ActivityTemplateNo = AAT.ActivityTemplateNo AND M.ROLEID IN ( SELECT DISTINCT ROLEID FROM Ms_Roles) "&_
										       " AND Cast(R.APPLICATIONCODE as Varchar)+ ':' + cast(R.PROCESSCODE as Varchar) + ':' + cast(R.ACTIVITYCODE as Varchar)+ ':' + cast(R.ActivityTemplateNo as Varchar) IN  ( "&_
											   " select DISTINCT (Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) + ':' + cast(ActivityTemplateNo as Varchar) )  FROM MS_USERACTIVITY WHERE INTERNALUSERID =  "& nEmployeeNo & " ) "
										
										If sProcessCode <> "S" and sProcessCode <> "" and sProcessCode <> "0" Then
											sSql = sSql & " and R.APPLICATIONCODE = "& sProcessCode &" "
										End IF

										If sPracticeCode <> "S" and sPracticeCode <> "" and sProcessCode <> "0" Then
											sSql = sSql & " and R.PROCESSCODE IN ( "& sPracticeCode &" ) "
										End IF

										sSql = sSql & " ORDER BY M.ROLEID,R.APPLICATIONCODE,R.PROCESSCODE,R.ACTIVITYCODE,AAT.ActivityTemplateNo "

					
							With rsMain
								.ActiveConnection = Con
								.CursorLocation = 3
								.CursorType = 3
								.Source = sSql
								.Open
							End With

							Set rsMain.ActiveConnection = Nothing

							iCtr  = 1
							nSlNo = 1
							nKK   = 1

							If Not rsMain.EOF Then

								rsMain.PageSize = iPageSize
								rsMain.AbsolutePage = iCurrentPage
								iTotalPage = rsMain.PageCount

								Do While Not rsMain.EOF and nSlNo <= rsMain.PageSize

									iSpan = 200

									If sPrevRoleID = cint("0") or cint(sPrevRoleID) <> cint(rsMain(0)) Then
										sPrevRoleID = cint(rsMain(0))
										sPrevAppName = ""
										sPrevProcessName = ""
										%>
										<tr>
											<td class="ExcelSerial" align="Left" Colspan="3"><B><%=Trim(rsMain(1))%></b></td>
									    </tr>
										<%
									End IF

									'Module Name
									IF sPrevAppName = "" or sPrevAppName <> rsMain(2) then
										sPrevAppName = rsMain(2)%>
										<tr>
											<td class=ExcelSerial align=left colspan=3>&nbsp;&nbsp;&nbsp;<B><%=UCase(rsMain(2))%></B></td>
										</tr>
									<%End IF

									'Practice Name
									If sPrevProcessName = "" or sPrevProcessName <> rsMain(3) Then
										sPrevProcessName = rsMain(3)%>
										<tr>
											<td class=ExcelSerial align=left colspan=3>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><%=rsMain(3)%></B></td>
										</tr>
									<%End IF

									'ActivityName
									'rsMain(0) & ":" & rsMain(7) & ":" & rsMain(4) & ":" & rsMain(5) 'ROLEID,APPCODE,PRCODE,ACTCODE
									sTempVal = rsMain(0) & ":" & rsMain(7) & ":" & rsMain(4) & ":" & rsMain(5) &":"& rsMain(8)
									%>
									<tr>
										<td class="ExcelSerial" align="center" rowspan=<%'=iSpan%> ><%=nSlNo%></td>
										<td class="ExcelDisplayCell" align="center" width="10" rowspan=<%'=iSpan%> >
											<input type="checkbox" name="Chkbox<%=nSlNo%>" value="<%=sTempVal%>">
										</td>
										<td class=ExcelDisplayCell align=left><%=rsMain(6)%>-<%=rsMain(9) %></td>
									</tr>

									<%

									iCtr = iCtr + 1
									nSlNo = nSlNo + 1

									rsMain.MoveNext
								Loop
							End IF 'If Not rsMain.EOF Then
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
										<Input type="button" value="Assign Activity" name="ButOpt" class="ActionButtonX" tabindex="3" onclick="ShowActivityDetails()">
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
