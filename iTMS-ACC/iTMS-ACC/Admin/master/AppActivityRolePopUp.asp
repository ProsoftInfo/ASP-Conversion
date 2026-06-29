<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppActivityRolePopUp.asp
	'Module Name				:	Admin (Role Creation)
	'Author Name				:	UmaMaheswari S
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>

<title>Role Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"><Root/></XML>
<XML ID=PRData></XML>
<XML ID="RetData"><Root Done=""/></XML>
<XML ID="SelectedData"><Root/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript SRC="../../scripts/Cancel.vbs"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>

	Function ShowData()
		Dim sTemp,CallFrom,sData,nEmployeeNo,nRoleID,nProcessCode,nPracticeCode

		nEmployeeNo = document.formname.hEmpID.value
		'nRoleID  = document.formname.hRoleID.value
		'sTemp = nEmployeeNo & ":" & nRoleID

		nRoleID = document.formname.selRole.value
		nProcessCode = document.formname.selProcess.value
		nPracticeCode = document.formname.selActivityType.value

		sTemp = nEmployeeNo & ":" & nRoleID & ":" & nProcessCode & ":" & nPracticeCode

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLSelect.asp?sWho=UANM&sPassData="& sTemp , false
		'objhttp.Open "GET","XMLSelect.asp?sWho=UANM&sUser="& nEmployeeNo, false
		objhttp.send

		'alert(objhttp.responseXML.xml)
		'Exit Function

		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
		else
			Exit Function
		end if
		DisplayTable()
	End Function

	Function DisplayTable()
		Dim Root,Node,i,iCtr,sProcessName,sPracticeName,sTempValue,sRoleName,nRoleID

		Set Root = OutData.DocumentElement

		ClearTable()
		i = 1
		j = 0
		iSNo = 1

		If Root.hasChildNodes Then

			For Each Node In Root.childNodes
				If Node.NodeName = "ROLE" Then

					'j = j + 1

					iCtr = iCtr + 1

					'set oRow = document.all.tblData.insertRow(j)
					set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length)

					set headerCell=oRow.insertCell()
					headerCell.innerHTML=ucase(Node.Attributes.getNamedItem("ROLENAME").value)
					headerCell.className="ExcelHeaderCell"
					headerCell.colspan = 3
					headerCell.align="left"

					nRoleID = Node.Attributes.getNamedItem("ROLEID").value

					For Each SubNode in Node.childNodes
						If SubNode.nodeName = "TOMAP" Then

							'Process Name
							If SubNode.Attributes.getNamedItem("APPName").value <> sProcessName Then
								j = j + 1
								sProcessName = SubNode.Attributes.getNamedItem("APPName").value

								'set oRow = document.all.tblData.insertRow(j)
								set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length)
								set headerCell=oRow.insertCell()
								headerCell.innerHTML=ucase(sProcessName)
								headerCell.className="ExcelHeaderCell"
								headerCell.colspan = 3
								headerCell.align="left"
								j = j + 1
							End IF

							'Practice Name
							If SubNode.Attributes.getNamedItem("PAName").value <> sPracticeName Then

								sPracticeName = SubNode.Attributes.getNamedItem("PAName").value

								'set oRow = document.all.tblData.insertRow(j)
								set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length)
								set headerCell=oRow.insertCell()
								headerCell.innerHTML = SubNode.Attributes.getNamedItem("PAName").value
								headerCell.className="ExcelHeaderCell"
								headerCell.colspan = 3
								headerCell.align="left"
								j = j + 1
							End IF

							'set oRow = document.all.tblData.insertRow(j)
							set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length)

							sTempValue = SubNode.getAttribute("APPCode") &":" & SubNode.getAttribute("PRCode") & ":" & SubNode.getAttribute("ACCode") & ":" & nRoleID

							set headerCell=oRow.insertCell()
							headerCell.innerHTML= iSNo
							headerCell.className="ExcelSerial"
							headerCell.align="center"

							set headerCell=oRow.insertCell()
							headerCell.innerHTML="<input type=CheckBox NAME='mChkValue"+cstr(iSNo)+"' size='11' value=" & trim(sTempValue) & "  class='Formelem'>"
							headerCell.className="ExcelDisplayCell"
							headerCell.align="Center"

							set headerCell=oRow.insertCell()
							headerCell.innerHTML= SubNode.Attributes.getNamedItem("ACName").value
							headerCell.className="ExcelDisplayCell"
							headerCell.align="left"

							iSNo = iSNo  + 1

							document.formname.hItemRows.value = iSNo

						End IF
						j = j + 1
					Next
				End If
			Next

		End IF	'If Root.hasChildNodes Then

	End Function

	Function ClearTable()
		Dim i
		For i = 1 to document.all.tblData.rows.length - 1
			document.all.tblData.deleteRow(1)
		Next
	End Function

	Function CheckSubmit()
		Dim sRoot,Root,SelRoot

		set Root = OutData.documentElement
		set sRoot = RetData.documentElement
		set SelRoot = SelectedData.documentElement

		'nRows=document.formname.hItemRows.value - 1
		nRows=document.formname.hCnt.value

		'set OutValue = showModalDialog ("ItemTypeSelection.asp","","dialogHeight:300px;dialogWidth:290px;center:Yes;help:No;resizable:No;status:No;scroll:no")
		'document.formname.hItemTypeID.value = OutValue.getAttribute("ItemTypeID")

		'SelRoot.setAttribute "RoleID",document.formname.hRoleID.value
		SelRoot.setAttribute "UserID",document.formname.hEmpID.value
		'SelRoot.setAttribute "ItemType",document.formname.hItemTypeID.value

		for nTempCtr = 1 to  nRows
			set objChk = eval("document.formname.mchkValue"&cstr(nTempCtr))

			if objchk.checked then

				'ROLEID,APPCODE,PRCODE,ACTCODE
				Set Node = SelectedData.createElement("ACTIVITY")
				Node.setAttribute "APPCode",split(objchk.value,":")(1)
				Node.setAttribute "PRCode",split(objchk.value,":")(2)
				Node.setAttribute "ACCode",split(objchk.value,":")(3)
				Node.setAttribute "RoleID",split(objchk.value,":")(0)
				Node.setAttribute "TempNo",split(objchk.value,":")(4)
				SelRoot.appendChild Node
			End IF
		Next

		SaveData()

	End Function

	Function SaveData()

		Dim Root,objHttp,sRoot

		Set Root = SelectedData.DocumentElement
		set sRoot = RetData.documentElement

		sCallFrom = "ADDAPPUSERROLE"

		sExp = "./ACTIVITY"
		set TempNode = Root.selectNodes(sExp)
		If TempNode.length > 0 Then
		Else
			alert("Select Any One Activity For Add")
			Exit Function
		End IF

		set OutValue = showModalDialog ("ItemTypeSelection.asp","","dialogHeight:250px;dialogWidth:240px;center:Yes;help:No;resizable:No;status:No;scroll:no")
		If OutValue.getAttribute("Done") = "" Then
			alert("Select Item Type")
			Exit Function
		End IF

		document.formname.hItemTypeID.value = OutValue.getAttribute("ItemTypeID")
		Root.setAttribute "ItemType",document.formname.hItemTypeID.value

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","RoleActivityMappingInsert.asp?sPassData="&sCallFrom,False
		objHttp.send SelectedData.XMLDocument

		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Added Successfully")
			sRoot.setAttribute "Done","Y"
			document.formname.submit
		End if
	End Function

	Function FinalSubmit()
		document.formname.action = "AppActivityRole.asp?EmpNo="&document.formname.hEmpID.value
		document.formname.submit
	End Function

	'Function window_onunload()
	'	set window.returnvalue = RetData.documentElement
	'End Function

	Function PopulatePractice(obj)
		IF obj = "0" Then
			alert("Select Prcoess Name")
			document.formname.selProcess.focus
			Exit Function
		End IF

		document.formname.selActivityType.options.length = 1

		dim Root,HeaderNode

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLSelect.asp?sWho=PR&sProcess="& obj, false

		objhttp.send

		if objhttp.responseXML.xml <> "" then
			PRData.loadXML objhttp.responseXML.xml
			Set Root = PRData.documentElement
			if Root.HaschildNodes() then
				For Each HeaderNode In Root.childNodes
					document.formname.selActivityType.length = document.formname.selActivityType.length+1
					document.formname.selActivityType.options(document.formname.selActivityType.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
					document.formname.selActivityType.options(document.formname.selActivityType.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
				next
			end if
		else
			alert("No Practice defined for the Process Selected")
			document.formname.selProcess.focus
			Exit Function
		end if

	end Function

	Function Validate()
		document.formname.hRoleID.value = document.formname.selRole.value
		document.formname.hProcessCode.value = document.formname.selProcess.value

		For iCtr = 0 to document.formname.selActivityType.length - 1
			If document.formname.selActivityType.options(iCtr).selected Then
				If document.formname.selActivityType.options(iCtr).value = "S" Then
				Else
					sPracticeCode = sPracticeCode & "," & document.formname.selActivityType.options(iCtr).value
				End IF
			End IF
		Next
		'sPracticeCode = Trim(document.formname.selActivityType.value)
		If sPracticeCode <> "" Then sPracticeCode = mid(sPracticeCode,2)

		document.formname.hPracticeCode.value = sPracticeCode
		document.formname.submit
	End Function

	Function Paginate(nPage)
		document.formname.hPageSelection.value = nPage
		document.formname.selRole.value = document.formname.hRoleID.value
		document.formname.selProcess.value = document.formname.hProcessCode.value
		document.formname.submit()
	End function

</SCRIPT>
<%
	Dim sTemp,sArr,sPassType,sSql,nRoleID,sRoleName,nEmpID,sProcessCode
	DIm nSlNo,sDefValue,sTempVal,nSelAppCode,nKK,iCtr,sItemType,sPracticeCode
	Dim iCurrentPage,iTotalPage,lnPage

	Dim dcrs,ObjRs,rsMain
	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set ObjRs = Server.CreateObject("ADODB.RecordSet")

	Const iPageSize = 15
	iCurrentPage=CInt(Request.Form("hPageSelection"))
	if iCurrentPage=0 then iCurrentPage=1

	nEmpID     = Request.QueryString("PassData")
	'sArr      = Split(sTemp,":")
	'nEmpID    = Trim(sArr(0))
	'nRoleID   = Trim(sArr(1))
	'sItemType = Trim(sArr(2))

	nRoleID = Request("selRole")
	sProcessCode = Request("selProcess")
	sPracticeCode  = Request("selActivityType")

	'Response.Write "<p>nRoleID = "&nRoleID
	'Response.Write "<p>sProcessCode = "&sProcessCode
	'Response.Write "<p>sPracticeCode = "&sPracticeCode

%>
</head>
<!--<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="ShowData()">-->
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<form method="POST" name="formname" action="">
		<Input type="hidden" name="hItemRows" value="">

		<Input type="hidden" name="hRoleID" value="<%=nRoleID%>">
		<Input type="hidden" name="hEmpID" value="<%=nEmpID%>">
		<Input type="hidden" name="hProcessCode" value="<%=sProcessCode%>">
		<Input type="hidden" name="hPracticeCode" value="<%=sPracticeCode%>">
		<Input type="hidden" name="hItemTypeID" value="">

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
						<a style="width: 1em; height: 1em;" title="" onclick="Div_OnClick(idUnprocessed);" >
						<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
						</a>
						</td>
						<td valign="right" class="SubTitle">
						</td>
						</tr>
						</table>

						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
						<td width="100%">
						<div id="idUnprocessed" style="width: 575; display: none">
						<table cellpadding="0" cellspacing="0">

						<tr>
							<td class="FieldCellSub">Role</td>
							<td class="FieldCellSub">
								<select size="5" name="selRole" class="FormElem">
									<option value="ALL" selected>ALL</option>
								<%
									with objRs
										.CursorLocation = 3
										.CursorType = 3
										.Source = "SELECT DISTINCT ROLEID,ROLEDESCRIPTION FROM Ms_Roles "
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


							<td class="FieldCellSub">Process Name</td>
							<td class="FieldCellSub">
								<select size="5" name="selProcess" class="FormElem" onChange="PopulatePractice(this.value)">
									<option value="S" selected>Select</option>
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
								<input type="button" value="Show" name="Cmdgo" class="ActionButton" onclick="Validate()">
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
						<td></td>
						<td valign="top" width="100%">
						 <!--<div class="frmBody" id="frm2" style="width: 100%; height:230;">-->

							<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
								<tr>
									<td class="ExcelHeaderCell" align="center" width=10>S.No.</td>
									<td class="ExcelHeaderCell" align="center" width="20">
									</td>
									<td class="ExcelHeaderCell" align="center" >Activity</td>
								</tr>
								<%

										Dim sPrevAppName,sPrevProcessName,sPrevRoleID,nNoOfActivity,nNoOfModules,nNoOfPractice


										sPrevAppName = ""
										sPrevProcessName = ""
										sPrevRoleID = cint("0")
										
										
										sSql = "Select Distinct M.ROLEID,M.ROLEDESCRIPTION,A.APPLICATIONNAME,AP.PROCESSNAME,R.PROCESSCODE, "&_
											   " R.ACTIVITYCODE,AA.ACTIVITYNAME,R.APPLICATIONCODE,AAT.ActivityTemplateNo,AAT.ActivityTemplateName "&_
											   " from MS_ROLES M,Ms_Applications A,MS_ROLEACTIVITY R,Ms_ApplicationProcess AP,Ms_ApplicationActivity AA,Ms_ApplicationActivityTemplates AAT "&_
											   " where M.RoleID = R.ROleID and A.ApplicationCode = R.ApplicationCode and R.ApplicationCode = AP.ApplicationCode and "&_
											   " R.ProcessCode = AP.Processcode and A.ApplicationCode = AA.ApplicationCode and AP.ProcessCode= AA.ProcessCode "&_
											   " and R.ActivityCode = AA.ActivityCode and AA.ApplicationCode = AAT.APplicationCode and AA.ProcessCode = AAT.ProcessCode and "&_
											   " AA.ActivityCode = AAT.ActivityCode and R.ActivityTemplateNo = AAT.ActivityTemplateNo AND M.ROLEID IN "
											   
										If nRoleID <> "" and nRoleID <> "ALL" Then
											 ssql = ssql & " ( SELECT DISTINCT ROLEID FROM Ms_Roles WHERE ROLEID = "& nRoleID &" ) "
										Else
											 ssql = ssql & " ( SELECT DISTINCT ROLEID FROM Ms_Roles) "
										End IF
										
											  sSql = sSql & " AND Cast(R.APPLICATIONCODE as Varchar)+ ':' + cast(R.PROCESSCODE as Varchar) + ':' + cast(R.ACTIVITYCODE as Varchar)+ ':' + cast(R.ActivityTemplateNo as Varchar) NOT IN  ( "&_
											   " select DISTINCT (Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) + ':' + cast(ActivityTemplateNo as Varchar) )  FROM MS_USERACTIVITY WHERE INTERNALUSERID =  "& nEmpID & " ) "
										
										If sProcessCode <> "S" and sProcessCode <> "" and sProcessCode <> "0" Then
											sSql = sSql & " and R.APPLICATIONCODE = "& sProcessCode &" "
										End IF

										If sPracticeCode <> "S" and sPracticeCode <> "" and sProcessCode <> "0" Then
											sSql = sSql & " and R.PROCESSCODE IN ( "& sPracticeCode &" ) "
										End IF

										sSql = sSql & " ORDER BY M.ROLEID,R.APPLICATIONCODE,R.PROCESSCODE,R.ACTIVITYCODE,AAT.ActivityTemplateNo "
										'Response.Write sSql

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
												sTempVal = rsMain(0) & ":" & rsMain(7) & ":" & rsMain(4) & ":" & rsMain(5)&":"& rsMain(8)
												%>
												<tr>
													<td class="ExcelSerial" align="center"><%=nSlNo%></td>
													<td class="ExcelDisplayCell" align="center" width="10" >
														<input type="checkbox" align=center name="mchkValue<%=nSlNo%>" value="<%=sTempVal%>">
													</td>
													<td class=ExcelDisplayCell align=left><%=rsMain(6)%>-<%=rsMain(9)%></td>
												</tr>

												<%

												iCtr = iCtr + 1
												nSlNo = nSlNo + 1

												rsMain.MoveNext
											Loop
										End IF 'If Not rsMain.EOF Then
										rsMain.Close
										%>
										<input type="hidden" name="hCnt" value="<%=iCtr-1%>">
										</table>
										<!--</div>-->
									</td>
									<!--<td width=30% valign=top>
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="FieldCellSub">Item Type <Br>
													<select size="5" name="selIType" class="FormElem">
														<option value="S" selected>Select</option>
														<option value="">Not Applicable</option>
														<%	'Calling the Function which populates the Item Type list
															'populateItemType
														%>
													</select>
												</td>
											</tr>
										</Table>
									</td>-->

							</td>
								<td></td>
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
							<SELECT class="FormElem" onChange="Paginate(this(this.selectedIndex).value)" id=select1 name=select1>
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
										<Input type="button" value="ADD" name="ButOpt" class="ActionButtonX" tabindex="3" Onclick="CheckSubmit()">
										<Input type="button" value="Done" name="ButOpt" class="ActionButtonX" tabindex="3" onClick="FinalSubmit()">
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
	Function populateItemType()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT ITEMTYPEID,ITEMTYPENAME FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>
