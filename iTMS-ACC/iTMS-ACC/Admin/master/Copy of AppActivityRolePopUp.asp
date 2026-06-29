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
<SCRIPT LANGUAGE=vbscript>

	Function ShowData()
		Dim sTemp,CallFrom,sData,nEmployeeNo,nRoleID

		nEmployeeNo = document.formname.hEmpID.value
		nRoleID  = document.formname.hRoleID.value
		sTemp = nEmployeeNo & ":" & nRoleID

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLSelect.asp?sWho=UANM&sPassData="& sTemp , false
		'objhttp.Open "GET","XMLSelect.asp?sWho=UANM&sUser="& nEmployeeNo, false
		objhttp.send

		'alert(objhttp.responseXML.xml)

		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
		else
			Exit Function
		end if
		DisplayTable()
	End Function

	Function DisplayTable()
		Dim Root,Node,i,iCtr,sProcessName,sPracticeName,sTempValue

		Set Root = OutData.DocumentElement

		ClearTable()
		i = 1
		j = 0
		iSNo = 1

		If Root.hasChildNodes Then
			For Each Node In Root.childNodes
				If Node.NodeName = "TOMAP" Then

					iCtr = iCtr + 1
					j = j + 1

					'Node.Attributes.getNamedItem("CTR").Value = iCtr
					sTempValue = Node.getAttribute("APPCode") &":" & Node.getAttribute("PRCode") & ":" & Node.getAttribute("ACCode")

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

					set headerCell=oRow.insertCell()
					headerCell.innerHTML= iSNo
					headerCell.className="ExcelSerial"
					headerCell.align="center"

					set headerCell=oRow.insertCell()
					headerCell.innerHTML="<input type=CheckBox NAME='mChkValue"+cstr(iSNo)+"' value=" & trim(sTempValue) & "  size='11' class='Formelem'>"
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
		Dim sRoot,Root,SelRoot

		set Root = OutData.documentElement
		set sRoot = RetData.documentElement
		set SelRoot = SelectedData.documentElement

		nRows=document.formname.hItemRows.value - 1

		SelRoot.setAttribute "RoleID",document.formname.hRoleID.value
		SelRoot.setAttribute "UserID",document.formname.hEmpID.value
		SelRoot.setAttribute "ItemType",document.formname.hItemType.value

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
			window.close
		End IF

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","RoleActivityMappingInsert.asp?sPassData="&sCallFrom,False
		objHttp.send SelectedData.XMLDocument


		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Added Successfully")
			sRoot.setAttribute "Done","Y"
			set window.returnvalue = RetData.documentElement
		End if
		Window.close

	End Function

	Function window_onunload()
		set window.returnvalue = RetData.documentElement
	End Function

</SCRIPT>
<%
	Dim sTemp,sArr,sPassType,sSql,nRoleID,sRoleName,nEmpID,sItemType

	Dim dcrs
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sTemp   = Request.QueryString("PassData")
	sArr    = Split(sTemp,":")
	nEmpID  = Trim(sArr(0))
	nRoleID	= Trim(sArr(1))
	sItemType = Trim(sArr(2))
%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="ShowData()">

	<form method="POST" name="formname" action="">
		<Input type="hidden" name="hItemRows" value="">
		<Input type="hidden" name="hLastSelectedPractice" value="">
		<Input type="hidden" name="hRoleID" value="<%=nRoleID%>">
		<Input type="hidden" name="hEmpID" value="<%=nEmpID%>">
		<Input type="hidden" name="hItemType" value="<%=sItemType%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Role Activity
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
									<td class="TabCell" valign="bottom" align="center" width="95">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">

										</table>
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
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center"></td>
									<td width="100%">
										<div class="frmBody" id="frm2" style="width: 100%; height:190;">
											<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" align="center" width=10>S.No.</td>
													<td class="ExcelHeaderCell" align="center" width="15">
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
													<input type="Button" value="Done" name="B1" class="ActionButton" onclick="CheckSubmit()">
													<!--<input type="button" value="Close" name="B2" class="ActionButton" onClick="window.close()">-->
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
