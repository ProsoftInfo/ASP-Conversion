
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ApplicationRoles.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	December 07 ,2010
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
<script type="application/xml" data-itms-xml-island="1" ID = "OutData"><Root/></script>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AdminRoleActivityCompat.js"></SCRIPT>
<SCRIPT>
ITMSAdminRoleActivityCompat.installApplicationRoles();
</SCRIPT>
<% 	Dim sSql,sSql1,iCtr,sLoginID,sUnitID,sEmpValue,sEmpName,nKK,nSelRoleID,nProcessCode
	Dim iCurrentPage,iTotalPage,lnPage,nSlNo,sDefValue

	Const iPageSize = 16
	Dim objRs,objRs1,objRs2

	Set objRs = Server.CreateObject("ADODB.RecordSet")
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")
	Set objRs2 = Server.CreateObject("ADODB.RecordSet")

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	if iCurrentPage=0 then iCurrentPage=1

	if trim(sUnitID) = "" then	sUnitID = Session("organizationcode")

	nSelRoleID = cint(Request.Form("selRole"))

	If nSelRoleID = "0" Then
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES"
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing

		if not objRs.EOF then
			nSelRoleID = objrs(0)
		End IF
		objrs.Close
	End IF
	'Response.Write "<p>nRoleID="& nSelRoleID
	nProcessCode = cint(Request.Form("selProcess"))
	'Response.Write "<p>nProcessCode="&nProcessCode
%>
</Head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname">

	<input type=hidden name="hRoleID" value="<%=nSelRoleID%>">
	<input type=hidden name="hProcessCode" value="<%=nProcessCode%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0" >
		<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Application Roles
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
						<td>
						</td>
						<td valign="top" width="100%">
							<table border="0" cellspacing="1" class="ExcelTable" width="100%">

								<tr>
									<td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
									<td class="ExcelHeaderCell" align="center" width="10">
										<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" onclick="DeleteRole()" width="15" height="15"></a>
									</td>
									<td class="ExcelHeaderCell" align="center" >Role Name</td>
									<td class="ExcelHeaderCell" align="center" >Mapped Activities</td>
								</tr>

							<%
							Dim sPrevAppName,sPrevProcessName,sPrevRoleID,iSpan,nNoOfActivity,nNoOfModules,nNoOfPractice,sCondition

							sPrevAppName = ""
							sPrevProcessName = ""
							sPrevRoleID = cint("0")

							'sSql = "SELECT DISTINCT ROLEID,ROLEDESCRIPTION FROM MS_ROLES ORDER BY ROLEID "

							'sSql = " select Distinct M.ROLEID,M.ROLEDESCRIPTION,V.APPLICATIONNAME,V.PROCESSNAME,V.PROCESSCODE, "&_
							'		" R.ACTIVITYCODE,V.ACTIVITYNAME,R.APPLICATIONCODE from MS_ROLEACTIVITY R,VWACTIVITY V,MS_ROLES M "&_
							'		" where R.APPLICATIONCODE = V.APPLICATIONCODE and R.ACTIVITYCODE = V.ACTIVITYCODE "&_
							'		" and R.ROLEID = M.ROLEID ORDER BY V.PROCESSCODE,R.ACTIVITYCODE"

							'sSql = " select Distinct M.ROLEID,M.ROLEDESCRIPTION,V.APPLICATIONNAME,V.PROCESSNAME,V.PROCESSCODE, "&_
							'	   " R.ACTIVITYCODE,V.ACTIVITYNAME,R.APPLICATIONCODE from MS_ROLEACTIVITY R,VWACTIVITY V,MS_ROLES M,"&_
							'	   " VWUSERACTIVITY U where R.APPLICATIONCODE = V.APPLICATIONCODE and R.ACTIVITYCODE = V.ACTIVITYCODE "&_
							'	   " and R.ROLEID = M.ROLEID and R.ACTIVITYCODE = U.ACTIVITYCODE and V.PROCESSCODE=U.PROCESSCODE "

							'sSql = " select Distinct M.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONNAME,R.PROCESSNAME,R.PROCESSCODE, "&_
							'	   " R.ACTIVITYCODE,R.ACTIVITYNAME,R.APPLICATIONCODE from MS_ROLES M,Ms_Applications A,"&_
							'	   " VwROLEACTIVITY R where R.APPLICATIONCODE = A.APPLICATIONCODE  and R.ROLEID = M.ROLEID"

							'sSql = " select Distinct M.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONNAME,AP.PROCESSNAME,R.PROCESSCODE, "&_
							'	   " R.ACTIVITYCODE,R.ACTIVITYNAME,R.APPLICATIONCODE from MS_ROLES M,Ms_Applications A,"&_
							'	   " VwROLEACTIVITY R,Ms_ApplicationProcess AP where R.APPLICATIONCODE = A.APPLICATIONCODE  and R.ROLEID = M.ROLEID"&_
							'	   " and R.APPLICATIONCODE = AP.APPLICATIONCODE AND R.PROCESSCODE = AP.PROCESSCODE "

							sSql = " select Distinct M.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONNAME,AP.PROCESSNAME,R.PROCESSCODE, "&_
									" R.ACTIVITYCODE,AA.ACTIVITYNAME,R.APPLICATIONCODE from MS_ROLES M,Ms_Applications A,  "&_
									" MS_ROLEACTIVITY R,Ms_ApplicationProcess AP,Ms_ApplicationActivity AA where R.APPLICATIONCODE = A.APPLICATIONCODE "&_
									" and R.ROLEID = M.ROLEID and R.APPLICATIONCODE = AP.APPLICATIONCODE "&_
									" AND R.PROCESSCODE = AP.PROCESSCODE  "&_
									" and R.APPLICATIONCODE = AA.APPLICATIONCODE and R.PROCESSCODE = AA.PROCESSCODE and "&_
									" R.ACTIVITYCODE = AA.ACTIVITYCODE "
									
							sSql =  "Select Distinct M.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONNAME,AP.PROCESSNAME,R.PROCESSCODE, "&_
							        " R.ACTIVITYCODE,AA.ACTIVITYNAME,R.APPLICATIONCODE,AAT.ActivityTemplateName from MS_ROLES M,Ms_Applications A,  "&_
							        "  MS_ROLEACTIVITY R,Ms_ApplicationProcess AP,Ms_ApplicationActivity AA,Ms_ApplicationActivityTemplates AAT where R.APPLICATIONCODE = A.APPLICATIONCODE "&_
							        "  and R.ROLEID = M.ROLEID and R.APPLICATIONCODE = AP.APPLICATIONCODE  AND R.PROCESSCODE = AP.PROCESSCODE  "&_
							        "  and R.APPLICATIONCODE = AA.APPLICATIONCODE and R.PROCESSCODE = AA.PROCESSCODE and  R.ACTIVITYCODE = AA.ACTIVITYCODE "&_
							        " and AA.ApplicationCode = AAT.ApplicationCode and AA.ProcessCode = AAT.ProcessCode and AA.ActivityCode = AAT.ActivityCode "&_
							        " and R.ActivityTemplateNo = AAT.ActivityTemplateNo "

							If nSelRoleID <> "0" Then
								sSql = sSql & " AND R.ROLEID= "& nSelRoleID&" "
							End IF
							If nProcessCode <> "0" Then
								sSql = sSql & " AND R.APPLICATIONCODE= "& nProcessCode&" "
							End IF

							sCondition = sSql
							sSql = sSql & " ORDER BY M.ROLEID,R.APPLICATIONCODE,R.PROCESSCODE,R.ACTIVITYCODE "	'and M.RoleID = 1

							'Response.Write ssql

							With objRs
								.ActiveConnection = Con
								.CursorLocation = 3
								.CursorType = 3
								.Source = sSql
								.Open
							End With

							Set objRs.ActiveConnection = Nothing
							iCtr = 1
							nSlNo = 1

							If not objRs.EOF then
								objRs.PageSize = iPageSize
								objRs.AbsolutePage = iCurrentPage
								iTotalPage = objRs.PageCount
							Else
								With objRs1
									.ActiveConnection = Con
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES WHERE ROLEID="& nSelRoleID&" "
									.Open
								End With

								Set objRs1.ActiveConnection = Nothing

								If Not objRs1.EOF Then
									iCtr = iCtr + 1
									%>
									<tr>
										<td class="ExcelSerial" align="center"><%=nSlNo%></td>
										<td class="ExcelDisplayCell" align="center" width="10" >
											<input type="checkbox" name="Chkbox<%=nSlNo%>" value="<%=objrs1(0)%>">
										</td>
										<td class="ExcelDisplayCell" align="Left">
											<a href="#" class="ExcelDisplayLink" onclick="ShowRoleActivityMap('<%=objrs1(0)%>','<%=Trim(objrs1(1))%>'); return false;">
											<%=Trim(objrs1(1))%></a>
										</td>
										<td class=ExcelDisplayCell align=left>&nbsp;</td>
									</tr>
									<%
								End IF
								objRs1.Close
							End if 'If not objRs.EOF then

							nKK = 1
							nNoOfModules  = 0
							nNoOfPractice = 0
							nNoOfActivity = 0

							Do while not objRs.EOF and nSlNo <= objRs.PageSize

								If sPrevRoleID = cint("0") or cint(sPrevRoleID) <> cint(objrs(0)) Then

									sPrevRoleID = cint(objrs(0))

									'For NoOF MOdule
									With objRs1
										.ActiveConnection = Con
										.CursorLocation = 3
										.CursorType = 3
										'.Source = "select count(*) from Ms_Applications where APPLICATIONCODE = "& objrs(7)&" "
										.Source = " SELECT count(*) FROM Ms_ApplicationActivity WHERE ApplicationCode = "& objrs(7)&" "
										.Open
									End With

									Set objRs1.ActiveConnection = Nothing

									If Not objRs1.EOF Then
										nNoOfModules = objrs1(0)
									End IF
									objRs1.Close

									'For NoOf Practices
									With objRs1
										.ActiveConnection = Con
										.CursorLocation = 3
										.CursorType = 3
										.Source = "select count(*) from Ms_ApplicationProcess where APPLICATIONCODE = "& objrs(7)&" "
										.Open
									End With

									Set objRs1.ActiveConnection = Nothing

									If Not objRs1.EOF Then
										nNoOfPractice = objrs1(0)
									End IF
									objRs1.Close

									'For Noof Activity
									With objRs1
										.ActiveConnection = Con
										.CursorLocation = 3
										.CursorType = 3
										.Source = "select count(RoleID) from ("& sCondition &") as temp  where  RoleID='"& objrs(0) &"' "
										'.Source = "select count(*) from Ms_RoleActivity where RoleId = "& objrs(0) &" and ACTIVITYCODE IN (SELECT distinct ACTIVITYCODE FROM VWUSERACTIVITY WHERE PROCESSCODE = "& objrs(4)&")"
										.Open
									End With

									Set objRs1.ActiveConnection = Nothing

									If Not objRs1.EOF Then
										'nNoOfActivity = objrs1(0)
										iSpan = objrs1(0)
									End IF
									objRs1.Close

									iSpan = cint(nNoOfModules) + cint(nNoOfPractice) + cint(nNoOfActivity)
									iCtr = iCtr + 1
									'Response.Write "<p>iSpan="&nNoOfModules &"+"& nNoOfPractice & " + "& nNoOfActivity & "="& iSpan
									'Response.Write "<p>span="&iSpan
									iSpan = "200"
								%>
								<tr>
									<td class="ExcelSerial" align="center" rowspan=<%=iSpan%> ><%=nSlNo%></td>
									<td class="ExcelDisplayCell" align="center" width="10" rowspan=<%=iSpan%> >
										<input type="checkbox" name="Chkbox<%=nSlNo%>" value="<%=objrs(0)%>">
									</td>
									<td class="ExcelDisplayCell" align="Left" rowspan="<%=iSpan%>">
										<a href="#" class="ExcelDisplayLink" onclick="ShowRoleActivityMap('<%=objrs(0)%>','<%=Trim(objrs(1))%>'); return false;">
										<%=Trim(objrs(1))%></a>
									</td>
								<%End IF%>
									<!--<td class="ExcelDisplayCell" align="Left">-->
								<%
									'Module Name
									IF sPrevAppName = "" or sPrevAppName <> objrs(2) then
										sPrevAppName = objrs(2)%>
										<td class=ExcelDisplayCell align=left><B><%=UCase(objrs(2))%></B></td></tr>
									<%End IF

									'Practice Name
									If sPrevProcessName = "" or sPrevProcessName <> objrs(3) Then
										sPrevProcessName = objrs(3)%>
										<tr>
											<td class=ExcelDisplayCell align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B><%=objrs(3)%></B></td>
										</tr>
									<%End IF


									'ActivityName
									%>
									<tr>
										<td class=ExcelDisplayCell align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=objrs(6)%>&nbsp;-&nbsp;<%=objrs(8)%></td>
									</tr>
								

								<%
									nSlNo = nSlNo + 1

								objRs.MoveNext
							Loop
							objRs.Close
							'Response.Write "<p>iCtr="&iCtr
								%>
								

							</table>
							<input type="hidden" name="hCnt" value="<%=iCtr-1%>">
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

							<Select class="FormElem" Name="selRole" onchange="AssaignRoleID(this.value)">
								<% 'PopulateRole %>
								<%
								with objRs
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES"
									.ActiveConnection = con
									.Open
								end with
								set objRs.ActiveConnection = nothing

								if not objRs.EOF then

									Do While Not objRs.EOF

										If cint(nSelRoleID) = cinT(objRs(0)) Then%>
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

							<Select class="FormElem" Name="selProcess" Onchange="AssaignProcessCode(this.value)">
								<option value="0" SELECTED>ALL</option>
								<%
								with objRs
									.CursorLocation = 3
									.CursorType = 3
									.Source = "select APPLICATIONCODE,APPLICATIONNAME from Ms_Applications "
									.ActiveConnection = con
									.Open
								end with
								set objRs.ActiveConnection = nothing

								if not objRs.EOF then

									Do While Not objRs.EOF

										If cint(nProcessCode) = cinT(objRs(0)) Then%>
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
										<Input type="button" value="Create" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction('CRN')" >
										<Input type="button" value="Edit" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction('EDT')" >
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
