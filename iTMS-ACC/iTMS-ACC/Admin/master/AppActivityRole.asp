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
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript SRC="../../scripts/Cancel.vbs"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
	Function populateList(nEmployeeNo,RoleID)

		Dim Root,HeaderNode,OutDataRoot,sTemp

		sTemp = nEmployeeNo & ":" & RoleID

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'objhttp.Open "GET","XMLSelect.asp?sWho=UAMAP&sUser="& nEmployeeNo, false
		objhttp.Open "GET","XMLSelect.asp?sWho=UAMAP&sPassData="& sTemp, false
		objhttp.send

		'alert(objhttp.responseXML.xml)
		'Exit Function

		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
		end if

		DisplayTable()

		'document.formname.selRole.selectedIndex = RoleID
	end Function

	Function DisplayTable()
		Dim Root,Node,i,iCtr,sProcessName,sPracticeName,sTempValue,sRoleName

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

							sTempValue = SubNode.getAttribute("APPCode") &":" & SubNode.getAttribute("PRCode") & ":" & SubNode.getAttribute("ACCode")

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


	Function DisplayTable_OLD()
		Dim Root,Node,i,iCtr,sProcessName,sPracticeName,sTempValue,sRoleName

		Set Root = OutData.DocumentElement

		ClearTable()
		i = 1
		j = 0
		iSNo = 1

		If Root.hasChildNodes Then

			sRoleName = Root.Attributes.getNamedItem("ROLENAME").value

			For Each Node In Root.childNodes
				If Node.NodeName = "TOMAP" Then

					iCtr = iCtr + 1
					j = j + 1

					'Role Name
					If sRoleName <> "" Then

						set oRow = document.all.tblData.insertRow(j)
						set headerCell=oRow.insertCell()
						headerCell.innerHTML=ucase(sRoleName)
						headerCell.className="ExcelHeaderCell"
						headerCell.colspan = 3
						headerCell.align="left"
						sRoleName = ""
						j = j + 1
					End IF

					'Process Name
					If Node.Attributes.getNamedItem("APPName").value <> sProcessName Then
						sProcessName = Node.Attributes.getNamedItem("APPName").value

						set oRow = document.all.tblData.insertRow(j)
						set headerCell=oRow.insertCell()
						headerCell.innerHTML=ucase(sProcessName)
						headerCell.className="ExcelHeaderCell"
						headerCell.colspan = 3
						headerCell.align="left"
						j = j + 1
					End IF

					'Practice Name
					If Node.Attributes.getNamedItem("PAName").value <> sPracticeName Then
						sPracticeName = Node.Attributes.getNamedItem("PAName").value

						set oRow = document.all.tblData.insertRow(j)
						set headerCell=oRow.insertCell()
						headerCell.innerHTML = Node.Attributes.getNamedItem("PAName").value
						headerCell.className="ExcelHeaderCell"
						headerCell.colspan = 3
						headerCell.align="left"
						j = j + 1
					End IF


					set oRow = document.all.tblData.insertRow(j)

					sTempValue = Node.getAttribute("APPCode") &":" & Node.getAttribute("PRCode") & ":" & Node.getAttribute("ACCode")

					set headerCell=oRow.insertCell()
					headerCell.innerHTML= iSNo
					headerCell.className="ExcelSerial"
					headerCell.align="center"

					set headerCell=oRow.insertCell()
					headerCell.innerHTML="<input type=CheckBox NAME='mChkValue"+cstr(iSNo)+"' size='11' value=" & trim(sTempValue) & "  class='Formelem'>"
					headerCell.className="ExcelDisplayCell"
					headerCell.align="Center"

					set headerCell=oRow.insertCell()
					headerCell.innerHTML= Node.Attributes.getNamedItem("ACName").value
					headerCell.className="ExcelDisplayCell"
					headerCell.align="left"

					iSNo = iSNo  + 1

					document.formname.hItemRows.value = iSNo

					'document.formname.hItemRows.value =iCtr
				End IF
				i = i + 1
			Next
		End IF

	End Function

	Function ClearTable()
		Dim i
		For i = 1 to document.all.tblData.rows.length - 1
			document.all.tblData.deleteRow(1)
		Next
	End Function


	Function CheckSubmit()
		document.formname.action ="ApplicationUserGrid.asp"
		document.formname.submit()
	end Function


	Function ShowActivityDetails_OLD()
		Dim sTemp,nEmpID,sItemType

		If document.formname.selRole.value = "S" Then
			alert("Select Role")
			document.formname.selRole.focus()
			Exit Function
		Elseif document.formname.selIType.value = "S" Then
			alert("Select Item Type")
			document.formname.selIType.focus
			Exit Function
		End IF

		nEmpID = document.formname.hEmpID.value
		sItemType = Trim(document.formname.selIType.value)

		sTemp = nEmpID & ":" & document.formname.selRole.value & ":" & sItemType

		set OutValue = showModalDialog ("AppActivityRolePopUp.asp?PassData="&sTemp,"","dialogHeight:320px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No;scroll:no")

		If OutValue.getAttribute("Done") = "Y" Then
			populateList nEmpID,document.formname.selRole.value
		End IF
	End Function

	Function DeleteItem_OLD()
		Dim SelRoot,nRows,objHttp
		set SelRoot = SelectedData.documentElement

		nRows=document.formname.hItemRows.value - 1
		SelRoot.setAttribute "RoleID",document.formname.selRole.value
		SelRoot.setAttribute "UserID",document.formname.hEmpID.value

		for nTempCtr = 1 to  nRows
			set objChk = eval("document.formname.mchkValue"&cstr(nTempCtr))

			if objchk.checked then

				Set Node = SelectedData.createElement("ACTIVITY")
				Node.setAttribute "APPCode",split(objchk.value,":")(0)
				Node.setAttribute "PRCode",split(objchk.value,":")(1)
				Node.setAttribute "ACCode",split(objchk.value,":")(2)
				SelRoot.appendChild Node
			End IF
		Next

		'alert(SelectedData.xml)
		'Exit Function

		sExp = "./ACTIVITY"
		set TempNode = SelRoot.selectNodes(sExp)
		If TempNode.length > 0 Then
		Else
			alert("Select Activity For deletion")
			Exit Function
		End IF

		sCallFrom = "DELAPPUSERROLE"

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","RoleActivityMappingInsert.asp?sPassData="&sCallFrom,False
		objHttp.send SelectedData.XMLDocument

		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Deleted Successfully")
		End if
		populateList document.formname.hEmpID.value,document.formname.selRole.value
	End Function
'---------------------------------------------------------------------------------------------------
	'Added
	Function Validate()
		Dim sProcessCode,sPracticeCode,iCtr

		sProcessCode  = Trim(document.formname.selProcess.value)

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

		document.formname.hProcessCode.value = sProcessCode
		document.formname.hPracticeCode.value = sPracticeCode

		document.formname.submit()
	End Function

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

	Function Paginate(nPage)
		document.formname.hPageSelection.value = nPage
		document.formname.selProcess.value = document.formname.hProcessCode.value
		document.formname.selActivityType.value  = document.formname.hPracticeCode.value
		document.formname.submit()
	End function

	Function DeleteItem()
		Dim SelRoot,nRows,objHttp
		set SelRoot = SelectedData.documentElement

		nRows=document.formname.hCnt.value

		SelRoot.setAttribute "UserID",document.formname.hEmpID.value

		for nTempCtr = 1 to  nRows
			set objChk = eval("document.formname.Chkbox"&cstr(nTempCtr))

			if objchk.checked then

				Set Node = SelectedData.createElement("ACTIVITY")
				Node.setAttribute "RoleID",split(objchk.value,":")(0)
				Node.setAttribute "APPCode",split(objchk.value,":")(1)
				Node.setAttribute "PRCode",split(objchk.value,":")(2)
				Node.setAttribute "ACCode",split(objchk.value,":")(3)
				Node.setAttribute "TempNo",split(objchk.value,":")(4)
				SelRoot.appendChild Node
			End IF
		Next

		sExp = "./ACTIVITY"
		set TempNode = SelRoot.selectNodes(sExp)
		If TempNode.length > 0 Then
		Else
			alert("Select Activity For deletion")
			Exit Function
		End IF

		sCallFrom = "DELAPPUSERROLE"

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","RoleActivityMappingInsert.asp?sPassData="&sCallFrom,False
		objHttp.send SelectedData.XMLDocument

		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Deleted Successfully")
			document.formname.submit()
		End if

	End Function

	Function ShowActivityDetails()
		Dim sTemp,nEmpID,sItemType

		nEmpID = document.formname.hEmpID.value
		'sTemp = nEmpID & ":" & document.formname.selRole.value & ":" & sItemType

		'set OutValue = showModalDialog ("AppActivityRolePopUp.asp?PassData="&nEmpID,"","dialogHeight:380px;dialogWidth:540px;center:Yes;help:No;resizable:No;status:No;scroll:YES")
		document.formname.action = "AppActivityRolePopUp.asp?PassData="&nEmpID
		document.formname.submit
	End Function

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
						<a style="width: 1em; height: 1em;" title="" onclick="Div_OnClick(idUnprocessed);" >
						<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
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
									<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record"  onclick="DeleteItem()" width="15" height="15"></a>
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
