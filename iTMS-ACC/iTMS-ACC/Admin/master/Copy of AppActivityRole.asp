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
	Dim nEmployeeNo,sEmployeeName
	nEmployeeNo = Request.QueryString("EmpNo")

	Dim dcrs

	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "Select (isNull(UserName,'')+isNull(MiddleName,'')+isNull(LastName,'')) From DCS_User where EmployeeNumber = "& nEmployeeNo &" "
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	If not dcrs.EOF then
		sEmployeeName = dcrs(0)
	End IF
	dcrs.close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Practice Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"></XML>
<XML ID="SelectedData"><Root/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
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


	Function ShowActivityDetails()
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

	Function DeleteItem()
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
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="populateList('<%=nEmployeeNo%>','1')">
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >

	<form method="POST" name="formname" action="">
		<Input type="hidden" name="hEmpID" value="<%=nEmployeeNo%>">
		<Input type="hidden" name="hItemRows" value="">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Role Details
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>

									<td class="TabCellEnd" valign="bottom" align="left">&nbsp;
									</td>
								</tr>

							</table>
						</td>
					</tr>

					<tr>
						<td class="TabBody">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<div align="left">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td class="FieldCell">User Name</td>
													<td class="FieldCellSub">
														<span class=dataonly><%=sEmployeeName%></span>
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Role</td>
													<td class="FieldCellSub">
														<select size="4" name="selRole" class="FormElem" onChange="populateList('<%=nEmployeeNo%>',this.value)">

														<%	'Calling the Function to populate Applications / Process List
															PopulateRole
														%>
														</select>
													</td>
													<td class="FieldCell">Item Type</td>
													<td class="FieldCellSub">
														<select size="5" name="selIType" class="FormElem">
															<option value="S" selected>Select</option>
															<option value="">Not Applicable</option>
															<%	'Calling the Function which populates the Item Type list
																populateItemType
															%>
														</select>
													</td>
												</tr>
												<tr>
													<td class="FieldCell"></td>
													<td class="FieldCellSub"><Input type=button name=btnView value="View & Add" class="ActionButton" onclick="ShowActivityDetails()"></td>
												</tr>

											</table>
										</div>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center"></td>
									<td width="100%">
										<div class="frmBody" id="frm2" style="width:700; height:200;">
											<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" align="center" width=10>S.No.</td>
													<td class="ExcelHeaderCell" align="center" width="20">
														<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record"  onclick="DeleteItem()" width="15" height="15"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" >Activity</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<!--<input type="button" value="Amend" name="B2" class="ActionButton" onClick="CheckSubmit()">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">-->
 													<input type="Button" value="Done" name="B1" class="ActionButton" onclick="CheckSubmit()">
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="BottomPack">
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
	' Function to populate Roles
	Function PopulateRole()
		'Declaration of variables
		Dim dcrs,iCtr
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		iCtr = 0

		if not dcrs.EOF then
			Do While Not dcrs.EOF
				iCtr = iCtr + 1
				If iCtr = 1 Then
					Response.Write "<option value="""&trim(dcrs(0))&""" Selected >"&trim(trim(dcrs(1)))&"</option>" & vbCrLf
				Else
					Response.Write "<option value="""&trim(dcrs(0))&""">"&trim(trim(dcrs(1)))&"</option>" & vbCrLf
				End IF
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

