<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppRoleActivityMappingPopUp.asp
	'Module Name				:	Admin (Role Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	December 08, 2010
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
<XML id="ActivityData"><Root Eligible="N"></Root></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
	Function PopulatePractice(obj)
		'if obj.selectedIndex = 0 then exit function
		Dim sCallFrom,nPracticeCode

		sCallFrom = "FROMPROCESS"
		nPracticeCode = "0"

		ShowData sCallFrom,obj.value

		If obj.selectedIndex = 0 Then
			document.formname.selPractice.options.length = 1
			document.formname.selPractice.selectedIndex = 0
			document.formname.selActivtyDes.options.length = 1
			document.formname.selActivtyDes.selectedIndex = 0
			Exit Function
		End IF

		document.formname.selPractice.options.length = 1

		dim Root,HeaderNode

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLSelect.asp?sWho=PR&sProcess="& obj.value, false

		objhttp.send

		if objhttp.responseXML.xml <> "" then
			PRData.loadXML objhttp.responseXML.xml
			Set Root = PRData.documentElement
			if Root.HaschildNodes() then
				For Each HeaderNode In Root.childNodes
					document.formname.selPractice.length = document.formname.selPractice.length+1
					document.formname.selPractice.options(document.formname.selPractice.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
					document.formname.selPractice.options(document.formname.selPractice.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
				next
			end if
			document.formname.selPractice.selectedIndex = 0
			document.formname.selActivtyDes.options.length = 1
			document.formname.selActivtyDes.selectedIndex = 0
		else
			Exit Function
		end if

	End Function

	Function PopulateActivities(obj)
		Dim  Root,HeaderNode,nProcess,nRoleID

		nProcess = document.formname.selProcess.value
		nRoleID = document.formname.hRoleID.value

		sCallFrom = "FROMPRACTICE"
		nPracticeCode = "0"

		If obj.selectedIndex = 0 Then
			document.formname.selActivtyDes.options.length = 1
			document.formname.selActivtyDes.selectedIndex = 0
			Exit Function
		End IF

		'alert(obj.value & ":" & nProcess & ":" & nRoleID)

		document.formname.selActivtyDes.options.length = 1

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLSelect.asp?sWho=PPA&sProcess="& obj.value &":"& nProcess&":"&nRoleID, false
		objhttp.send

		'alert(objhttp.responseText)

		if objhttp.responseXML.xml <> "" then
			PRData.loadXML objhttp.responseXML.xml
			Set Root = PRData.documentElement
			if Root.HaschildNodes() then
				For Each HeaderNode In Root.childNodes
			        document.formname.selActivtyDes.length = document.formname.selActivtyDes.length+1
				    document.formname.selActivtyDes.options(document.formname.selActivtyDes.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
				    document.formname.selActivtyDes.options(document.formname.selActivtyDes.length-1).Value = HeaderNode.Attributes.Item(1).nodeValue
				next
			end if
			document.formname.selActivtyDes.selectedIndex = 0
		else
			Exit Function
		end if

		ShowData sCallFrom,obj.value

		document.formname.hLastSelectedPractice.value = obj.value
		document.formname.hLastSelectedProcess.value = nProcess
	End Function

	Function AddEntry()
		Dim nAppCode,sAppName,nProcessCode,sProcessName,nActCode,sActName,nKK
		Dim Root
		Set Root = OutData.DocumentElement

		nAppCode = document.formname.selProcess.value
		nProcessCode = document.formname.selPractice.value
		nActCode = document.formname.selActivtyDes.value

		If nAppCode = "S" Then
			alert("Select Process")
			document.formname.selProcess.focus()
			Exit Function
		elseif nProcessCode = "S" Then
			alert("Select Practice")
			document.formname.selPractice.focus()
			Exit Function
		elseif nActCode = "S" then
			alert("Select Activity")
			document.formname.selActivtyDes.focus()
			Exit Function
		End IF

		sAppName = document.formname.selProcess.options(document.formname.selProcess.selectedIndex).text
		sProcessName = document.formname.selPractice.options(document.formname.selPractice.selectedIndex).text

		nKK = document.formname.selActivtyDes.length
		nActCode = ""
		sActName = ""
		For i = 0 to nKK - 1
			If document.formname.selActivtyDes.options(i).selected Then
				nActCode = nActCode & "," & split(document.formname.selActivtyDes.options(i).value,":")(0)
				sActName = sActName & "," & document.formname.selActivtyDes.options(i).text
			End If
		Next
		nActCode = mid(nActCode,2)
		sActName = mid(sActName,2)

		'alert(Root.xml)

		Flag = True
		sExp = "./ACTIVITYMAPPING"
		set TempNode = Root.selectNodes(sExp)
		If TempNode.length > 0 Then
			set Node = TempNode.item(0)
			Flag = False
		End IF

		If Flag = True Then
			set Node = OutData.createElement("ACTIVITYMAPPING")
			Node.setAttribute "CTR","0"
			Node.setAttribute "APPCODE",trim(nAppCode)
			Node.setAttribute "APPNAME",sAppName
			Node.setAttribute "PROCESSCODE",trim(nProcessCode)
			Node.setAttribute "PROCESSNAME",sProcessName
			Root.appendchild Node
		End IF
		
		
		set objhttp = createObject("Microsoft.XMLHTTP")
		objhttp.open "GET","XMLGetActTemp.asp?AppCode="&nAppCode&"&ProcessCode="&nProcessCode&"&ActCode="&nActCode,false
		objhttp.send 
		'alert(objhttp.responseText)
		if trim(objhttp.responseXML.xml)<>"" then
		    ActivityData.loadXML(objhttp.responseXML.xml)
		end if
		
		set ndTempRoot = ActivityData.documentElement
		'alert(ndTempRoot.xml)
		sEligible = ndTempRoot.getAttribute("Eligible")
		
		if sEligible="Y" then
		   set OutActData = showModalDialog("ActRolTempSelPop.asp?AppCode="&nAppCode&"&ProcessCode="&nProcessCode&"&ActCode="&nActCode,"","dialogWidth:500px;dialogHeight:500px")
		   'alert(OutActData.xml)
		   if OutActData.hasChildNodes() then
		        for each ndRole in OutActData.childNodes    
		            if ndRole.nodeName="Role" then
		                nActivityCode = ndRole.getAttribute("ActCode")
		                sActivityName = ndRole.getAttribute("ActName")
		                nTempNo = ndRole.getAttribute("TempNo")
		                sTempName = ndRole.getAttribute("TempName")
		                
		                If Root.hasChildNodes then
				            For Each MapNode in Root.childNodes
					            If MapNode.nodeName = "ACTIVITYMAPPING" Then
						            For Each ActNode in MapNode.childNodes
							            If ActNode.nodeName = "ACTIVITY" Then
								            If trim(ActNode.Attributes.getNamedItem("NAME").Value) &"-"& trim(ActNode.Attributes.getNamedItem("TNAME").value) = Trim(sActivityName)&"-"&trim(sTempName)  Then
									            alert("Already Selected Activity is Mapped")
									            Exit Function
								            End IF
							            End IF
						            Next
					            End IF
				            Next
			            End IF
			            
			            Set ActNode = OutData.createElement("ACTIVITY")
			                ActNode.setAttribute "ACTCTR","0"
			                ActNode.setAttribute "CODE",nActivityCode
			                ActNode.setAttribute "NAME",sActivityName
			                ActNode.setAttribute "TCODE",nTempNo
			                ActNode.setAttribute "TNAME",sTempName
			            Node.appendChild ActNode
		                
		            end if
		        next
		   end if
		else
	        For i = 0 to nKK - 1
		        If document.formname.selActivtyDes.options(i).selected Then
		            nActivityCode = document.formname.selActivtyDes.options(i).value
			        nTempNo = 1
			        sActivityName = document.formname.selActivtyDes.options(i).text
			        sTempName =  sActivityName
			        if ndTempRoot.hasChildNodes() then
			            for each ndTempAct in ndTempRoot.childNodes
			                if ndTempAct.nodeName="Activity" then   
			                    for each ndTemplate in ndTempAct.childNodes
			                        if ndTemplate.nodeName="Template" then
			                            nTempNo = ndTemplate.getAttribute("TempNo")
			                            sTempName = ndTemplate.getAttribute("TempName")
			                            exit for
			                        end if
			                    next
			                end if
			            next
			        end if
			        If Root.hasChildNodes then
				        For Each MapNode in Root.childNodes
					        If MapNode.nodeName = "ACTIVITYMAPPING" Then
						        For Each ActNode in MapNode.childNodes
							        If ActNode.nodeName = "ACTIVITY" Then
								        If ActNode.Attributes.getNamedItem("NAME").Value = Trim(sActivityName)  Then
									        alert("Already Selected Activity is Mapped")
									        Exit Function
								        End IF
							        End IF
						        Next
					        End IF
				        Next
			        End IF


			        Set ActNode = OutData.createElement("ACTIVITY")
			        ActNode.setAttribute "ACTCTR","0"
			        ActNode.setAttribute "CODE",nActivityCode
			        ActNode.setAttribute "NAME",sActivityName
			        ActNode.setAttribute "TCODE",nTempNo
			        ActNode.setAttribute "TNAME",sTempName
			        Node.appendChild ActNode
		        End If
	        Next
		end if
		
		DisplayTable()
		SaveData()
	End Function

	Function DisplayTable()
		Dim Root,Node,i,iCtr,sProcessName
		Set Root = OutData.DocumentElement

		ClearTable()
		i = 1
		j = 0
		iSNo = 1

		If Root.hasChildNodes Then
			For Each Node In Root.childNodes
				If Node.NodeName = "ACTIVITYMAPPING" Then

					iCtr = iCtr + 1
					j = j + 1

					Node.Attributes.getNamedItem("CTR").Value = iCtr

					'Process Name
					If Node.Attributes.getNamedItem("APPNAME").value <> sProcessName Then
						sProcessName = Node.Attributes.getNamedItem("APPNAME").value

						set oRow = document.all.tblData.insertRow(j)
						set headerCell=oRow.insertCell()
						headerCell.innerHTML=ucase(sProcessName)
						headerCell.className="ExcelHeaderCell"
						headerCell.colspan = 3
						headerCell.align="left"
						j = j + 1
					End IF

					'Practice Name
					set oRow = document.all.tblData.insertRow(j)

					set headerCell=oRow.insertCell()
					headerCell.innerHTML = Node.Attributes.getNamedItem("PROCESSNAME").value
					headerCell.className="ExcelHeaderCell"
					headerCell.colspan = 3
					headerCell.align="left"
					'iSNo = 1

					For Each SubNode In Node.childNodes
						If SubNode.nodeName = "ACTIVITY" Then

							j = j + 1
							set oRow = document.all.tblData.insertRow(j)

							SubNode.Attributes.getNamedItem("ACTCTR").Value = iSNo

							set headerCell=oRow.insertCell()
							headerCell.innerHTML= iSNo
							headerCell.className="ExcelSerial"
							headerCell.align="center"

							set headerCell=oRow.insertCell()
							headerCell.innerHTML="<input type=CheckBox NAME='mChkValue"+cstr(iSNo)+"' 'onclick=DeleteItem("+cstr(sPassCode)+") size='11' class='Formelem'>"
							headerCell.className="ExcelDisplayCell"
							headerCell.align="Center"

							set headerCell=oRow.insertCell()
							headerCell.innerHTML= SubNode.Attributes.getNamedItem("NAME").value &"-"& SubNode.Attributes.getNamedItem("TNAME").value
							headerCell.className="ExcelDisplayCell"
							headerCell.align="left"
							

							iSNo = iSNo  + 1
						End IF
						document.formname.hItemRows.value = iSNo
					Next

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

	Function DeleteItem()
		Dim nRows,nTempCtr,nActivityCode,nApplicationCode

		nRows=document.formname.hItemRows.value - 1
		set Root = OutData.DocumentElement
		nActivityCode = ""
		nApplicationCode = ""

		for nTempCtr = 1 to  nRows
			set objChk = eval("document.formname.mchkValue"&cstr(nTempCtr))
			if objchk.checked then

				for each Node in Root.childnodes
					If Node.nodeName="ACTIVITYMAPPING" then
						For Each SubNode in Node.childNodes
							If SubNode.nodeName = "ACTIVITY" then
								if cint(SubNode.getAttribute("ACTCTR"))=cint(nTempCtr) then
									'Node.removeChild(SubNode)
									nApplicationCode = Node.getAttribute("APPCODE")
									nProcessCode = Node.getAttribute("PROCESSCODE")
									nActivityCode = nActivityCode & "," & SubNode.getAttribute("CODE")
									nTempNo = nTempNo &","& SubNode.getAttribute("TCODE")
								End IF
							End If
						Next
					End if
				next
			End IF
		Next
		DisplayTable()

		If nActivityCode <> "" Then nActivityCode = mid(nActivityCode,2)
		if nTempNo <>"" then nTempNo = mid(nTempNo,2)
		if confirm("This will delete the Activity which u selected. Do you want to continue?") then
			DeleteData nApplicationCode,nProcessCode,nActivityCode,nTempNo
		End IF
	End Function

	Function ShowData(CallFrom,sData)

		Dim sTemp,nRoleID

		nRoleID = document.formname.hRoleID.value

		sTemp = CallFrom & ":" & nRoleID & ":" & sData	'If FROMPROCESS sData carries AppCode

		If CallFrom = "FROMPRACTICE" Then
			sTemp = CallFrom & ":" & nRoleID & ":" & document.formname.selProcess.value &":"& sData
		End IF

		'alert(sTemp)									'If FROMPRACTICE sdata carries Practicecode

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'objhttp.Open "GET","XMLSelect.asp?sWho=PPAONLOAD&RoleID="& document.formname.hRoleID.value , false
		objhttp.Open "GET","XMLSelect.asp?sWho=PPAONLOAD&sPassData="& sTemp , false
		objhttp.send
		
        'alert(objhttp.responseText)
		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
		else
			Exit Function
		end if
		'alert(OutData.xml)
		DisplayTable()
	End Function

	Function SaveData()
		Dim Root,objHttp
		Set Root = OutData.DocumentElement

		sCallFrom = "ADD"
		nRoleID = document.formname.hRoleID.value

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","RoleActivityMappingInsert.asp?sPassData="&sCallFrom & ":" & nRoleID,False
		objHttp.send OutData.XMLDocument

		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Added Successfully")
		End if

		'document.formname.selProcess.selectedIndex = 1
		'document.formname.selProcess.selectedIndex = 1
		For nKK = 0 to document.formname.selPractice.length - 1
			If document.formname.selPractice.options(nKK).value = document.formname.hLastSelectedPractice.value Then
				document.formname.selPractice.SelectedIndex = nKK
			End If
		Next

		For nKK = 0 to document.formname.selProcess.length - 1
			If document.formname.selProcess.options(nKK).value = document.formname.hLastSelectedProcess.value Then
				document.formname.selProcess.SelectedIndex = nKK
			End If
		Next
		'document.formname.selPractice.selectedIndex = 1

		PopulateActivities document.formname.selPractice
		ClearTable()
		DisplayTable()
	End Function

	Function DeleteData(ApplicationCode,ProcessCode,ActivityCode,TemplateNo)
		Dim sCallFrom,nRoleID

		sCallFrom = "DEL"
		nRoleID = document.formname.hRoleID.value
		sTemp = sCallFrom & ":" & ProcessCode & ":" & ActivityCode & ":" & ApplicationCode & ":" & nRoleID &":"& TemplateNo

		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","RoleActivityMappingInsert.asp?sPassData="&sTemp,False
		objHttp.send OutData.XMLDocument

		If objHttp.responseText <> "" Then
			alert(objHttp.responseText)
		Else
			alert("Record Deleted Successfully")
		End if

		'alert(document.formname.hLastSelectedPractice.value)
		'document.formname.selProcess.selectedIndex = 1
		For nKK = 0 to document.formname.selPractice.length - 1
			If document.formname.selPractice.options(nKK).value = document.formname.hLastSelectedPractice.value Then
				document.formname.selPractice.SelectedIndex = nKK
			End If
		Next

		For nKK = 0 to document.formname.selProcess.length - 1
			If document.formname.selProcess.options(nKK).value = document.formname.hLastSelectedProcess.value Then
				document.formname.selProcess.SelectedIndex = nKK
			End If
		Next
		'alert(document.formname.hLastSelectedProcess.value)
		'document.formname.selPractice.selectedIndex = 1

		PopulateActivities document.formname.selPractice

		ClearTable()
		DisplayTable()

	End Function

	Function CheckSubmit()
		Dim sRoot,Root
		set Root = OutData.documentElement
		set sRoot = RetData.documentElement

		sExp = "./ACTIVITYMAPPING"
		set TempNode = Root.selectNodes(sExp)

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
<%
	Dim sTemp,sArr,sPassType,sSql,nRoleID,sRoleName

	Dim dcrs
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sTemp = Request.QueryString("PassData")
	sArr  = Split(sTemp,":")
	nRoleID	  = Trim(sArr(0))
	sRoleName = Trim(sArr(1))

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >

	<form method="POST" name="formname" action="">
		<Input type="hidden" name="hItemRows" value="">
		<Input type="hidden" name="hLastSelectedPractice" value="">
		<Input type="hidden" name="hLastSelectedProcess" value="">
		<Input type="hidden" name="hRoleID" value="<%=nRoleID%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Role Activity Mapping
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
													<td class="FieldCell">Role Name</td>
													<td class="FieldCellSub"><span class=DataOnly><%=sRoleName%></span>
													</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<div align="left">
											<table cellpadding="0" cellspacing="0">
												<tr>
												</tr>
												<tr>
													<td class="FieldCell">Process</td>
													<td class="FieldCellSub">Practices</td>
													<td class="FieldCellSub">Activities</td>
												</tr>
												<tr>
													<td class="FieldCell">
														<select size="5" name="selProcess" class="FormElem" onChange="PopulatePractice(this)">
															<option value="S" selected>Select</Option>
															<%PopulateProcess%>
														</Select>
													</td>
													<td class="FieldCellSub">
														<select size="5" name="selPractice" class="FormElem" onchange="PopulateActivities(this)">
															<option value="S" selected>Select</Option>
														</Select>
													</td>
													<td class="FieldCellSub">
														<select size="5" name="selActivtyDes" class="FormElem" multiple>
															<option value="S" selected>Select</Option>
														</Select>
													</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<div align="center">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td class="FieldCell">&nbsp;</td>
													<td class="FieldCellSub" align=left>
														<input type="button" value="Add" name="BtnAdd" class="AddButton" onClick="AddEntry()">
													</td>
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
									<td align="center"></td>
									<td width="100%">
										<div class="frmBody" id="frm2" style="width: 100%; height:150;">
											<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" align="center" width=10>S.No.</td>
													<td class="ExcelHeaderCell" align="center" width="15">
														<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record"  onclick="DeleteItem()" width="15" height="15"></a>
													</td>
													<!--<td class="ExcelHeaderCell" align="center" >Process</td>
													<td class="ExcelHeaderCell" align="center" >Practice</td>-->
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
