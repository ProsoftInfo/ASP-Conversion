<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppActivityDetailsPopUp.asp
	'Module Name				:	Admin
	'Author Name				:	UmaMaheswari S
	'Created On					:	December 14, 2010
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
<%
	Dim sTemp,sArr,sPassType,sSql,nRoleID,nProcessCode,sProcessName,nPracticeCode,sType
	Dim sPracticeName,sPrintHeading,iCtr,sCheck

	Dim dcrs
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sTemp = Request.QueryString("PassData")
	sArr  = Split(sTemp,":")
	'Response.Write "<p>sTemp="&sTemp

	If Trim(sArr(0)) = "NEW" Then
		sCheck		  = Trim(sArr(0))
		nProcessCode  = Trim(sArr(1))
		sProcessName  = Trim(sArr(2))
		sType		  = Trim(sArr(3))	'PROCESS	'PRACTICE	'ADDPRACTICE
	Else
		nRoleID	      = Trim(sArr(0))
		nProcessCode  = Trim(sArr(1))
		sProcessName  = Trim(sArr(2))
		nPracticeCode = Trim(sArr(3))
		sPracticeName = Trim(sArr(4))
		sType		  = Trim(sArr(5))	'PROCESS	'PRACTICE	'ADDPRACTICE
	End IF

	'sArr  = Split(sTemp,":")
	'nRoleID	      = Trim(sArr(0))
	'nProcessCode  = Trim(sArr(1))
	'sProcessName  = Trim(sArr(2))
	'nPracticeCode = Trim(sArr(3))
	'sPracticeName = Trim(sArr(4))
	'sType		  = Trim(sArr(5))	'PROCESS	'PRACTICE	'ADDPRACTICE

	'Response.Write "<p>sType="&sType

	If Trim(sType) = "PROCESS" Then
		sPrintHeading = "PROCESS DETAILS"
	Elseif Trim(sType) = "ADDPRACTICE" Then
		sPrintHeading = "PRACTICE DETAILS"
	Else
		sPrintHeading = "PRACTICE DETAILS"
	End IF

	'Response.Write "<p>sTemp="&sTemp
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Role Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"><Root/></XML>
<XML ID=PRData></XML>
<XML ID="RetData"><Root Done=""/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
	Function GoToAction()
		Dim iCtr
		iCtr =  document.formname.hItemRows.value
		IF document.formname.hCheck.value = "NEW" Then
		'	document.formname.hItemRows.value = "1"
			Exit Function
		End IF

		If document.formname.hType.value="PROCESS" Then

			For i = 1 to iCtr
				If document.formname.ChkboxEdit.checked Then
					Eval("document.formname.txtPractice"&i).readonly = False
					Eval("document.formname.txtOrder"&i).readonly = False
					'Eval("document.formname.txtPractice"&i).Disabled = False
					Eval("document.formname.txtOrder"&i).Disabled = False
					'Eval("document.formname.ChkboxEdit"&cstr(i)).Disabled = False
				Else
					Eval("document.formname.txtPractice"&i).readonly = True
					Eval("document.formname.txtOrder"&i).readonly = True
					'Eval("document.formname.txtPractice"&i).Disabled = True
					Eval("document.formname.txtOrder"&i).Disabled = True
					'Eval("document.formname.ChkboxEdit"&cstr(i)).Disabled = True
				End IF
			Next

		Elseif document.formname.hType.value="PRACTICE" Then

			Set Root = OutData.DocumentElement

			For i = 1 to iCtr
				If document.formname.ChkboxEdit.checked Then
					Eval("document.formname.txtOrder"&i).Disabled	 = False
					Eval("document.formname.txtOrder"&i).Readonly	 = False
					Eval("document.formname.txtPractice"&i).readonly = False
				Else
					Eval("document.formname.txtOrder"&i).Disabled	 = True
					Eval("document.formname.txtPractice"&i).readonly = True
				End If
			Next

		Elseif document.formname.hType.value="ADDPRACTICE" Then

			For i = 1 to iCtr
				If document.formname.ChkboxEdit.checked Then
					Eval("document.formname.txtPractice"&i).readonly = False
					Eval("document.formname.txtOrder"&i).readonly = False
					'Eval("document.formname.txtPractice"&i).Disabled = False
					Eval("document.formname.txtOrder"&i).Disabled = False
					'Eval("document.formname.ChkboxEdit"&cstr(i)).Disabled = False
				Else
					Eval("document.formname.txtPractice"&i).readonly = True
					Eval("document.formname.txtOrder"&i).readonly = True
					'Eval("document.formname.txtPractice"&i).Disabled = True
					Eval("document.formname.txtOrder"&i).Disabled = True
					'Eval("document.formname.ChkboxEdit"&cstr(i)).Disabled = True
				End IF
			Next

		End IF
	End Function

	Function addRow()
		Dim nRow,iCtr,oRow,Root

		nRow = document.formname.hItemRows.value
		iCtr = document.all.tblData.rows.length

		Set Root = OutData.DocumentElement

		set oRow = document.all.tblData.insertRow()

		set headerCell = oRow.insertCell()
		headerCell.innerHTML = iCtr
		headerCell.className="ExcelSerial"
		headerCell.align = "center"

		set headerCell = oRow.insertCell()
		headerCell.innerHTML = "<input type=CheckBox NAME='ChkboxEdit"+cstr(iCtr)+"'  size='5' class='Formelem' value=""1"" DISABLED>"
		headerCell.className="ExcelDisplayCell"
		headerCell.align="Center"

		Set headerCell 	= oRow.insertCell()
		headerCell.className  = "ExcelInputCell"
		Set oText = document.createElement("<input type=""TEXT"" name='txtPractice" & iCtr&"' size=""40""   class=""FormElem"" maxlength=""50"" >")
		headerCell.appendChild(oText)

		Set headerCell 	= oRow.insertCell()
		headerCell.className  = "ExcelInputCell"
		headerCell.align="Center"
		Set oText = document.createElement("<input type=""TEXT"" name='txtOrder" & iCtr&"' size=""11""    class=""FormElem"" maxlength=""4"" >")
		headerCell.appendChild(oText)


		set NewElem = OutData.CreateElement("DETAILS")
		NewElem.setAttribute "CTR",iCtr
		NewElem.setAttribute "PROCESSCODE",""
		NewElem.setAttribute "PROCESSNAME",Eval("document.formname.txtPractice"&iCtr).value
		NewElem.setAttribute "ORDERNUMBER",Eval("document.formname.txtOrder"&iCtr).value
		Root.appendchild NewElem

	End Function

	Function ShowData(CallFrom,nProcessCode,nPracticeCode)

		Dim sTemp,nRoleID

		sTemp = CallFrom & ":" & nProcessCode & ":" & nPracticeCode

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLSelect.asp?sWho=ACTONLOAD&sPassData="& sTemp , false
		objhttp.send

		'alert(objhttp.responseXML.xml)

		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
		else
			Exit Function
		end if
		DisplayTable()

		If document.formname.hType.value="ADDPRACTICE" Then
			addRow()
		End IF
	End Function

	Function DisplayTable()
		Dim Root,Node,i,iCtr,sProcessName
		Set Root = OutData.DocumentElement

		ClearTable()
		i = 1

		If Root.hasChildNodes Then
			For Each Node In Root.childNodes
				If Node.NodeName = "DETAILS" Then

					iCtr = iCtr + 1

					set oRow = document.all.tblData.insertRow(i)

					set headerCell=oRow.insertCell()
					headerCell.width	= 10
					headerCell.innerHTML= iCtr
					headerCell.className="ExcelSerial"
					headerCell.align="center"

					set headerCell=oRow.insertCell()
					If document.formname.hType.value="PROCESS" Then
						headerCell.innerHTML="<input type=CheckBox NAME='ChkboxEdit"+cstr(iCtr)+"'  size='5' class='Formelem'>"
					Else
						headerCell.innerHTML="<input type=CheckBox NAME='ChkboxEdit"+cstr(iCtr)+"'  size='5' class='Formelem' DISABLED>"
					End IF
					headerCell.className="ExcelDisplayCell"
					headerCell.align="Center"

					Set headerCell 	= oRow.insertCell()
					headerCell.className  = "ExcelInputCell"
					Set oText = document.createElement("<input type=""TEXT"" name='txtPractice" & iCtr&"' size=""40"" value="""&Node.Attributes.getNamedItem("PROCESSNAME").value&""" READONLY class=""FormElem"" maxlength=""50"" DISABLED>")
					headerCell.appendChild(oText)

					Set headerCell 	= oRow.insertCell()
					headerCell.className  = "ExcelInputCell"
					headerCell.align="Center"
					Set oText = document.createElement("<input type=""TEXT"" name='txtOrder" & iCtr&"' size=""11"" value="""&Node.Attributes.getNamedItem("ORDERNUMBER").value&"""  READONLY class=""FormElem"" maxlength=""4"" DISABLED>")
					headerCell.appendChild(oText)

					document.formname.hItemRows.value =iCtr
				End IF
				i = i + 1
			Next
		End IF

		If document.formname.hType.value="PRACTICE" Then
			If Root.hasChildNodes Then
				For Each Node In Root.childNodes
					If Node.NodeName = "DETAILS" Then
						If Node.attributes.getNamedItem("PROCESSCODE").value=document.formname.hPracticeCode.value Then

							i = Node.attributes.getNamedItem("CTR").value
							Eval("document.formname.txtPractice"&i).Disabled = False
							Eval("document.formname.ChkboxEdit"&cstr(i)).Disabled = False
							Eval("document.formname.txtPractice"&i).readonly = False
						End IF
					End IF
				Next
			End IF

		End IF

	End Function

	Function ClearTable()
		Dim i
		For i = 1 to document.all.tblData.rows.length - 1
			document.all.tblData.deleteRow(1)
		Next
	End Function

	Function DeleteItem()
		Dim nRows,nTempCtr,nActivityCode,nApplicationCode,nProcessCode

		nRows = document.formname.hItemRows.value
		set Root = OutData.DocumentElement

		'nProcessCode  = document.formname.hPracticeCode.value
		nPracticeCode = document.formname.hProcessCode.value

		for nTempCtr = 1 to  nRows

			set objChk = eval("document.formname.ChkboxEdit"&cstr(nTempCtr))
			if objchk.checked then
				for each Node in Root.childnodes
					If Node.nodeName="DETAILS" then
						if cint(Node.getAttribute("CTR"))=cint(nTempCtr) then
							Root.removeChild(Node)
							nProcessCode = nProcessCode & "," & Node.getAttribute("PROCESSCODE")
						End IF
					End if
				next
			Else
				'alert("Select any One option for Delete")
				'Exit Function
			End IF
		Next


		If nProcessCode <> "" Then
			nProcessCode = mid(nProcessCode,2)
		Else
			alert("Select any One option for Delete")
			Exit Function
		End IF

		if confirm("This will delete the Activity which u selected. Do you want to continue?") then
			DeleteData nPracticeCode,nProcessCode
		End IF
		DisplayTable()
		GoToAction()

	End Function

	Function Max()
		Dim Root,nOrderNo,TempArr
		Set Root = OutData.DocumentElement
		For each Node in Root.childnodes
			If Node.nodeName="DETAILS" then
				nOrderNo = nOrderNo & "," & Node.getAttribute("ORDERNUMBER")
			End IF
		Next

		nOrderNo = mid(nOrderNo,2)
		TempArr = Split(nOrderNo,",")

		For i = 0 to UBound(TempArr)
			n1 =TempArr(i)
			If i = UBound(TempArr) Then
				n2 = TempArr(0)
			Else
				n2 = TempArr(i+1)
			End IF
			If n1 > n2 Then
				nMaxNo = TempArr(i)
			End IF
		Next
		Max = nMaxNo
	End FUnction

	Function SaveData(sCallFrom)
		Dim Root,objHttp,nProcessCode,nPracticeCode,sTemp,nAppName
		Set Root = OutData.DocumentElement

		nRows = document.formname.hItemRows.value

		'sCallFrom     = Trim(document.formname.hType.value)
		nProcessCode  = document.formname.hPracticeCode.value
		nPracticeCode = document.formname.hProcessCode.value
		nAppName = document.formname.txtPrcessName.value

		sTemp = sCallFrom & ":" & nProcessCode & ":" & nPracticeCode & ":" & nAppName
		sPrevOrderNo = "0"

		nMaxNo = Max()

		If sCallFrom = "ADDPRACTICE" Then
			nRows = nRows + 1
		End IF

		For nTempCtr = 1 to  nRows
			for each Node in Root.childnodes
				If Node.nodeName="DETAILS" then
					if cint(Node.getAttribute("CTR"))=cint(nTempCtr) then

						If sPrevOrderNo = Eval("document.formname.txtOrder"&nTempCtr).value Then
							alert("Same Order Number is Not allowed")
							Exit Function
						Elseif Eval("document.formname.txtPractice"&nTempCtr).value = "" Then
							alert("Enter Practice Name")
							Exit Function
						Elseif Eval("document.formname.txtOrder"&nTempCtr).value = "" Then
							alert("Enter Order No")
							Exit Function
						Elseif Not IsNumeric(Eval("document.formname.txtOrder"&nTempCtr).value) Then
							alert("Enter Numerals Only")
							Exit Function
						'Elseif cint(Eval("document.formname.txtOrder"&nTempCtr).value) > cint(nMaxNo+1) Then
						'	alert("Enter Unique value")
						'	Exit Function
						End IF

						sPrevOrderNo =  Eval("document.formname.txtOrder"&nTempCtr).value

						Node.Attributes.getNamedItem("ORDERNUMBER").Value = Eval("document.formname.txtOrder"&nTempCtr).value

						If Trim(sCallFrom) = "PRACTICE" or  Trim(sCallFrom) = "ADDPRACTICE" Then
							Node.Attributes.getNamedItem("PROCESSNAME").Value = Eval("document.formname.txtPractice"&nTempCtr).value
						End IF

					End IF
				End if
			next
		Next

		'alert(Root.xml)
		'Exit Function

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","AppActivityUpdate.asp?sPassData="&sTemp,False
		objHttp.send OutData.XMLDocument

		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Udated Successfully")

			set sRoot = RetData.documentElement
			sRoot.setAttribute "Done","Y"
			set window.returnvalue = RetData.documentElement

			ShowData sCallFrom,nPracticeCode,nProcessCode
		End if
		window.close
	End Function

	Function DeleteData(nPracticeCode,nProcessCode)
		Dim sCallFrom

		sTemp = nPracticeCode & ":" & nProcessCode

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","AppActivityDelete.asp?sPassData="&sTemp,False
		objHttp.send

		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Deleted Successfully")
		End if

	End Function

	Function CheckSubmit()
		Dim sRoot,Root,objHttp,sCallFrom

		set Root  = OutData.documentElement
		set sRoot = RetData.documentElement

		sExp = "./DETAILS"
		set TempNode = Root.selectNodes(sExp)
		Exit Function

		If TempNode.length > 0 Then
			sRoot.setAttribute "Done","Y"
			set window.returnvalue = RetData.documentElement
		End IF

		Window.close
	End Function

	Function window_onunload()
		set window.returnvalue = RetData.documentElement
	End Function

</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="ShowData('<%=sType%>','<%=nProcessCode%>','<%=nPracticeCode%>')">

	<form method="POST" name="formname" action="">
		<Input type="hidden" name="hItemRows" value="">
		<Input type="hidden" name="hLastSelectedPractice" value="">
		<Input type="hidden" name="hRoleID" value="<%=nRoleID%>">
		<Input type="hidden" name="hType" value="<%=sType%>">
		<Input type="hidden" name="hPracticeCode" value="<%=nPracticeCode%>">
		<Input type="hidden" name="hProcessCode" value="<%=nProcessCode%>">
		<Input type="hidden" name="hCheck" value="<%=sCheck%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center"><%=sPrintHeading%>
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
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<div align="left">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td class="FieldCell">Process Name</td>
													<td class="FieldCellSub">
														<input type=text name=txtPrcessName class=FormElem value="<%=sProcessName%>" <%If sType = "PRACTICE" or sType="ADDPRACTICE" Then Response.Write "ReadOnly"%>>
													</td>
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
										<div class="frmBody" id="frm2" style="width: 100%; height:155;">
											<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" align="center" width=10>S.No.</td>
													<td class="ExcelHeaderCell" align="center" width="15">
														<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record"  onclick="DeleteItem()" width="15" height="15"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" >Practice</td>
													<td class="ExcelHeaderCell" align="center" >Display Order<br>
														<input type="checkbox" name="ChkboxEdit" onclick="GoToAction()"> Edit
													</td>
												</tr>
											</table>
											<input type=hidden name="hiCtr" value="<%=iCtr%>">
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
													<input type="Button" value="Update" name="B1" class="ActionButton" onclick="SaveData('<%=sType%>')">
													<input type="button" value="Close" name="B2" class="ActionButton" onClick="window.close()">
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
