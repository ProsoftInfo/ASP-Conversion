<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ActivityMappingPopUp.asp
	'Module Name				:	Admin (Role Creation)
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Role Creation</title>
<base target="_self"/>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"><Root/></XML>
<XML ID="PRData"></XML>
<XML ID="RetData"><Root Done=""/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
'******************************************
Function DelActivity()
 set ndRoot = ActivityData.documentElement
    if ndRoot.hasChildNodes() then
        for each ndActivity in ndRoot.childNodes
            if ndActivity.nodeName="Activity" then
                for each ndTemp in ndActivity.childNodes
                    if ndTemp.nodeName="Template" then
                        TempNo = ndTemp.getAttribute("No")
                        set objChk = eval("document.formname.ChkTempZ"&TempNo)
                        if objChk.checked = true then
                            ndTemp.setAttribute "Del","Y"
                        end if
                    end if
                next
            end if
        next
    end if
    
    Set ObjHttp = CreateObject("Microsoft.XMLHTTP")
		ObjHttp.open "POST","AppActDel.asp",False
		ObjHttp.send ActivityData.XMLDocument
		'alert(ObjHttp.responseText)
		If ObjHttp.responseText <> "" Then
		    ArrValue = split(ObjHttp.responseText,":")
		    if UBound(ArrValue)=1 then
		        if trim(ArrValue(0))="ActNo"  then 
		            if trim(ArrValue(1))<>"0" then
		                sTempVal =  "EDT::"&document.formname.hAppCode.value&":"& document.formname.hAppName.value&":"&document.formname.hProcessCode.value&":"&document.formname.hProcessName.value&":"&ArrValue(1)
		                document.formname.action = "ActivityMappingPopUp.asp?PassData="&sTempVal
		                document.formname.submit
		            else
		                window.close
		            end if
		        end if
		    else
		        alert(objhttp.responseText)
		    end if
		End IF
End Function
'*******************************************
Function EditTemplate(TempNo)
    set ndRoot = ActivityData.documentElement
    if ndRoot.hasChildNodes() then
        for each ndActivity in ndRoot.childNodes
            if ndActivity.nodeName="Activity" then
                for each ndTemp in ndActivity.childNodes
                    if ndTemp.nodeName="Template" then
                        if TempNo = ndTemp.getAttribute("No") then
                            document.formname.txtTempName.value = ndTemp.getAttribute("Name")
                            document.formname.txtTempDesc.value = ndTemp.getAttribute("Description")
                            document.formname.txtPath.value = ndTemp.getAttribute("ProgramPath")
                            sStatus = ndTemp.getAttribute("Status") 
                            sProgramType = ndTemp.getAttribute("FileType")
                            sEmail = ndTemp.getAttribute("EMAIL")
                            sSMS = ndTemp.getAttribute("SMS")
                            if trim(sStatus) = "A" then
                                document.formname.radStatus(0).checked = true
                            else
                                document.formname.radStatus(1).checked = true
                            end if
                            For iCnt = 0 to document.formname.selFileType.length-1
                                if lcase(trim(document.formname.selFileType(iCnt).value))=lcase(trim(sProgramType)) then
                                    document.formname.selFileType.selectedIndex = iCnt
                                    exit for
                                end if
                            Next
                            if trim(sEmail)="Y" then
                                document.formname.chkEmail.checked = true
                            else
                                document.formname.chkEmail.checked = false
                            end if
                            
                            if trim(sSMS)="Y" then
                                document.formname.chkSMS.checked = true
                            else
                                document.formname.chkSMS.checked = false
                            end if
                                
                        end if  'if TempNo = ndTemp.getAttribute("No) then
                    end if
                next
            end if
        next
    end if
    document.formname.hTempNo.value = TempNo
End Function
'*********************************************
Function Init()
    ClearTable
    DisplayTable
End Function
'**************************************
Function DisplayTable() 
    iSer = 0
    iTempCnt = 0 
    set ndRoot = ActivityData.documentElement
    if ndRoot.hasChildNodes() then
        for each ndActivity in ndRoot.childNodes
            if ndActivity.nodeName="Activity" then
                for each ndTemp in ndActivity.childNodes
                    if ndTemp.nodeName="Template" then
                        TempNo = ndTemp.getAttribute("No")
                        TempName = ndTemp.getAttribute("Name")
                        ProgPath = ndTemp.getAttribute("ProgramPath")
                        
                        iSer=iSer + 1
                        iTempCnt =  iTempCnt + 1
                        set oRow = document.all.tblTemp.insertRow(document.all.tblTemp.rows.length)
	                        set headerCell=oRow.insertCell()
	                        headerCell.innerText = iSer 
	                        headerCell.className="ExcelSerial"
	                        headerCell.align="center"
	                    
	                        set headerCell=oRow.insertCell()
	                        set oText = document.createElement("<input type=Checkbox name=ChkTempZ"&TempNo&">")
	                        headerCell.appendChild(oText)
	                        headerCell.className="ExcelDisplayCell"
	                        headerCell.align="center"
	                        
	                        set headerCell=oRow.insertCell()
	                        headerCell.innerHtml = "<a href='#' onClick=EditTemplate('"& TempNo &"') class='ExcelDisplayLink' alt='Click to Edit Template'>"& TempName &"</a>"
	                        headerCell.className="ExcelDisplayCell"
	                        
	                        set headerCell=oRow.insertCell()
	                        headerCell.innerText = ProgPath 
	                        headerCell.className="ExcelDisplayCell"
	                    
                    end if
                next
            end if
        next
    end if
    document.formname.hTempCnt.value = iTempCnt
End Function
'**************************************
Function ClearTable()
    K = document.all.tblTemp.rows.length - 3
	for	i = 1 to  K
		document.all.tblTemp.deleteRow(1)
	next
End Function
'***************************************
	Function CheckSubmit()
		Dim Root,nProcessCode,nProcessName,nPracticeName,nPracticeCode
		Dim sActivityDescription,sProgramPath,sStatus
		
		set ndRoot = ActivityData.documentElement

    if trim(document.formname.hTempCnt.value)="0" then
		If trim(document.formname.txtDesc.value) = "" then
			alert("Enter Activity Description")
			document.formname.txtDesc.select()
			exit function
		elseif trim(document.formname.txtTempName.value)="" then
		    alert("Enter Template Name")
		    document.formname.txtTempName.select()
		    exit function
		elseif document.formname.selFileType.selectedIndex<=0 then
		    alert("Select File Type")
		    document.formname.selFileType.focus
		    exit function
		elseIf trim(document.formname.txtPath.value) = "" then
			alert("Enter Program Path")
			document.formname.txtPath.select()
			exit function
		End IF
	else
	    if trim(document.formname.txtTempName.value)<>"" then
	        if trim(document.formname.txtPath.value)="" then    
	            alert("You have entered Template name so please enter the path for the template")
	            document.formname.txtPath.select()
	            exit function
	        end if
	    end if
	    
	    if trim(document.formname.txtPath.value)<>"" then
	        if trim(document.formname.txtTempName.value)="" then    
	            alert("You have entered Path for the Template but the Template name is not entered so please enter")
	            document.formname.txtTempName.select()
	            exit function
	        end if
	    end if
	    
	    if trim(document.formname.txtTempName.value)="" and trim(document.formname.txtPath.value) = "" then
	        if confirm("Template Details not available you want to update Activity Details?") then
	            if ndRoot.hasChildNodes() then
                    for each ndActivity in ndRoot.childNodes
                        if ndActivity.nodeName="Activity" then
                            if document.formname.radStatus(0).checked = true then
                                sStatus = document.formname.radStatus(0).value
                            else
                                sStatus = document.formname.radStatus(1).value
                            end if
                                    
                            ndActivity.setAttribute "STATUS",sStatus
                        end if
                    next
                end if
	        else
	            exit function
	        end if
	    end if
	end if 
		


    sTempNo = document.formname.hTempNo.value
    
    if sTempNo<>"0" then
        if ndRoot.hasChildNodes() then
            for each ndActivity in ndRoot.childNodes
                if ndActivity.nodeName="Activity" then
                    ndActivity.setAttribute "ACTIVITYNAME",document.formname.txtDesc.value
                    
                    for each ndTemp in ndActivity.childNodes
                        if ndTemp.nodeName="Template" then
                            if sTempNo = ndTemp.getAttribute("No") then
                                ndTemp.setAttribute "Name",document.formname.txtTempName.value
                                ndTemp.setAttribute "Description",document.formname.txtTempDesc.value
                                ndTemp.setAttribute "ProgramPath",document.formname.txtPath.value
                                
                                if document.formname.radStatus(0).checked = true then
                                    sStatus = document.formname.radStatus(0).value
                                else
                                    sStatus = document.formname.radStatus(1).value
                                end if
                                
                                ndTemp.setAttribute "Status",sStatus 
                                ndTemp.setAttribute "FileType",document.formname.selFileType(document.formname.selFileType.selectedIndex).value
                                if document.formname.chkEmail.checked then
                                    ndTemp.setAttribute "EMAIL","Y"
                                else
                                    ndTemp.setAttribute "EMAIL","N"
                                end if
                                
                                if document.formname.chkSMS.checked then
                                    ndTemp.setAttribute "SMS","Y"
                                else
                                    ndTemp.setAttribute "SMS","N"
                                end if
                                
                                exit for
                            end if  'if TempNo = ndTemp.getAttribute("No) then
                        end if
                    next
                end if
            next
        end if
    else
        
        if trim(document.formname.txtTempName.value)<>"" or trim(document.formname.txtPath.value)<>"" then
            
        end if
        
        if ndRoot.hasChildNodes() then
            for each ndActivity in ndRoot.childNodes
                if ndActivity.nodeName="Activity" then
                ndActivity.setAttribute "ACTIVITYNAME",document.formname.txtDesc.value
                
                set ndTemplate = ActivityData.CreateElement("Template")
                    ndTemplate.setAttribute "No",""
                    ndTemplate.setAttribute "Name",document.formname.txtTempName.value
                    ndTemplate.setAttribute "Description",document.formname.txtTempDesc.value
                    ndTemplate.setAttribute "ProgramPath",document.formname.txtPath.value
                    
                    if document.formname.radStatus(0).checked = true then
                        sStatus = document.formname.radStatus(0).value
                    else
                        sStatus = document.formname.radStatus(1).value
                    end if
                    
                    ndTemplate.setAttribute "Status",sStatus 
                    ndTemplate.setAttribute "FileType",document.formname.selFileType(document.formname.selFileType.selectedIndex).value
                    if document.formname.chkEmail.checked then
                        ndTemplate.setAttribute "EMAIL","Y"
                    else
                        ndTemplate.setAttribute "EMAIL","N"
                    end if
                    
                    if document.formname.chkSMS.checked then
                        ndTemplate.setAttribute "SMS","Y"
                    else
                        ndTemplate.setAttribute "SMS","N"
                    end if
                    ndActivity.appendchild ndTemplate
                end if
            next
        end if
    end if
    

		Set ObjHttp = CreateObject("Microsoft.XMLHTTP")
		ObjHttp.open "POST","AppActivityCreationAndAmendInsert.asp",False
		ObjHttp.send ActivityData.XMLDocument
		'alert(ObjHttp.responseText)
		If ObjHttp.responseText <> "" Then
		    ArrValue = split(ObjHttp.responseText,":")
		    if UBound(ArrValue)=1 then
		        if trim(ArrValue(0))="ActNo" then 
		            sTempVal =  "EDT::"&document.formname.hAppCode.value&":"& document.formname.hAppName.value&":"&document.formname.hProcessCode.value&":"&document.formname.hProcessName.value&":"&ArrValue(1)
		            document.formname.action = "ActivityMappingPopUp.asp?PassData="&sTempVal
		            document.formname.submit
		        end if
		    else
		        alert(objhttp.responseText)
		    end if
		End IF

	end Function
	Function window_onunload()
	    set sRoot = RetData.documentElement
			sRoot.setAttribute "Done","Y"
		set window.returnvalue = RetData.documentElement
	End Function

</SCRIPT>

<%
    Dim objXML,ndRoot,ndActivity,ndTemplate,objFSO
	Dim sTemp,sArr,sPassType,sSql,nRoleID,nAppCode,sAppName,nProcessCode,sProcessName
	Dim nActivityCode,sActivityName,sCallFrom,sProgramPath,sStatus,sActStatus

	Dim dcrs
	set objXML = Server.CreateObject("Microsoft.XMLDOM")
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sTemp = Request.QueryString("PassData")

	sArr  = Split(sTemp,":")
	sCallFrom = sArr(0)
	nRoleID	  = Trim(sArr(1))
	nAppCode = Trim(sArr(2))
	sAppName = Trim(sArr(3))
	nProcessCode = Trim(sArr(4))
	sProcessName = Trim(sArr(5))
	
	set ndRoot = objXML.createElement("Root")
	objXML.appendChild ndRoot
	
	set ndActivity = objXML.createElement("Activity")
	    ndActivity.setAttribute "TYPE",sCallFrom
	    ndActivity.setAttribute "APPCODE",nAppCode
	    ndActivity.setAttribute "APPNAME",sAppName
	    ndActivity.setAttribute "PROCESSCODE",nProcessCode
	    ndActivity.setAttribute "ACTIVITYCODE",""
	    ndActivity.setAttribute "ACTIVITYNAME",""
	    ndActivity.setAttribute "STATUS",""
	    ndRoot.appendChild ndActivity
	
	If sCallFrom = "EDT" Then
		nActivityCode = Trim(sArr(6))
	End IF
	
	if trim(nActivityCode)<>"" then
	    sSql =  " Select Distinct ACTIVITYNAME,STATUS  From Ms_ApplicationActivity Where APPLICATIONCODE = "& nAppCode &" and "&_
			    "  ProcessCode = "& nProcessCode &" and ActivityCode = "& nActivityCode &" "
			    dcrs.open sSql,con
		        If Not dcrs.Eof Then
			        sActivityName = dcrs(0)
			        sActStatus = dcrs(1)
		        End IF
		        dcrs.close
        ndActivity.setAttribute "ACTIVITYCODE",nActivityCode
	    ndActivity.setAttribute "ACTIVITYNAME",sActivityName
	    ndActivity.setAttribute "STATUS",sActStatus
	   
	    sSql = "Select ActivityTemplateNo,IsNull(ActivityTemplateName,''),IsNull(TemplateDescription,''),IsNull(ProgramPath,''),Status,IsNull(ProgramType,''),IsNull(SendAsEmail,'N'),IsNull(SendAsSMS,'N') from Ms_ApplicationActivityTemplates where ApplicationCode = "& nAppCode &" and ProcessCode = "& nProcessCode &" and ActivityCode = "& nActivityCode
	    dcrs.open sSql,con
	    if not dcrs.eof then
	        do while not dcrs.eof
            	set ndTemplate = objXML.CreateElement("Template")
                    ndTemplate.setAttribute "No",dcrs(0)
                    ndTemplate.setAttribute "Name",dcrs(1)
                    ndTemplate.setAttribute "Description",dcrs(2)
                    ndTemplate.setAttribute "ProgramPath",dcrs(3)
                    ndTemplate.setAttribute "Status",dcrs(4)
                    ndTemplate.setAttribute "FileType",dcrs(5)
                    ndTemplate.setAttribute "EMAIL",dcrs(6)
                    ndTemplate.setAttribute "SMS",dcrs(7)
                    ndActivity.appendchild ndTemplate
	            dcrs.movenext
	        loop
	    end if
	    dcrs.close
	end if 'if trim(nActivityCode)<>"" then
	
	objXML.save(server.mappath("../temp/ActivityData_"&session.sessionID&".xml"))

%>
<%if objFSO.fileexists(Server.Mappath("../temp/ActivityData_"&session.sessionID&".xml")) then%>
    <XML id="ActivityData" src="<%="../temp/ActivityData_"&session.sessionID&".xml"%>"></XML>
<%else %>
    <XML id="ActivityData"><Root></Root></XML>
<%end if%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="Init()" >

	<form method="POST" name="formname" action="">
		<Input type="hidden" name="hItemRows" value="">
		<Input type="hidden" name="hLastSelectedPractice" value="">
		<Input type="hidden" name="hRoleID" value="<%=nRoleID%>">
		<Input type="hidden" name="hAppCode" value="<%=nAppCode%>">
		<Input type="hidden" name="hAppName" value="<%=sAppName%>">
		<Input type="hidden" name="hProcessCode" value="<%=nProcessCode%>">
		<Input type="hidden" name="hProcessName" value="<%=sProcessName%>">
		<Input type="hidden" name="hActCode" value="<%=nActivityCode%>">
		<Input type="hidden" name="hActName" value="<%=sActivityName%>">
		<input type="hidden" name="hTempNo" value="0">
		<input type="hidden" name="hTempCnt" value="0">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Role Activity Mapping
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
													<td class="FieldCell">Process</td>
													<td class="FieldCellSub"><span class=DataOnly><%=sAppName%></span>
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Practice</td>
													<td class="FieldCellSub"><span class=DataOnly><%=sProcessName%></span>
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Activity Description</td>
													<td class="FieldCellSub">
														<%If sCallFrom = "EDT" Then%>
															<input type="text" name="txtDesc" maxlength=50 size="50" class="FormElem" value="<%=sActivityName%>">
														<%Else%>
															<input type="text" name="txtDesc" maxlength=50 size="50" class="FormElem">
														<%End IF%>
													</td>
												</tr>
												<tr>
													<td class="FieldCell"> Template Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtTempName" maxlength=50 size="50" class="FormElem">
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Template Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtTempDesc" maxlength=50 size="50" class="FormElem">
													</td>
												</tr>
												
												<%
												If sCallFrom = "EDT" Then
													sSql = " Select ProgamPath,Status From VW_Admin_ApplicationActivities Where APPLICATIONCODE = "& nAppCode &" and "&_
														   " ProcessCode = "& nProcessCode &" and ActivityCode = "& nActivityCode &" "

													with dcrs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sSql
														.ActiveConnection = con
														.Open
													end with
													set dcrs.ActiveConnection = nothing
													If Not dcrs.Eof Then
														sProgramPath = Trim(dcrs(0))
														sStatus = Trim(dcrs(1))
													End IF
													dcrs.close
												End IF
												%>
												<tr>
												    <td class="FieldCell">Program Type</td>
												    <td class="FieldCellSub">
												        <select id="selFileType" class="FormElem">
												            <option value="sel">Select</option>
												            <option value="WebForm">Web Form</option>
												            <option value="WinForm">Win Form</option>
												        </select>
												    </td>
												</tr>
												<tr>
													<td class="FieldCell">Program Path</td>
													<td class="FieldCellSub">
															<input type="text" name="txtPath" maxlength=200 size="50" class="FormElem">
														 &nbsp;
														 <input type="button" name="btnAdd" class="AddButtonX" value="Add" onclick="CheckSubmit()">
													</td>
													
												</tr>
												<tr>
												    <td class="FieldCell">Send As</td>
												    <td class="FieldCellSub">
												        <input type="checkbox" name="chkEmail" value="E" class="FormElem"/>Email
												        <input type="checkbox" name="chkSMS" value="M"/ class="FormElem">SMS
												    </td>
												</tr>
												<tr>
													<td class="FieldCell">Status</td>
													<td class="FieldCellSub">
														<input type="Radio" name="radStatus" class="FormElem" Value="A" <%if trim(sActStatus)="A" then response.write "Checked" %>>Active
														<input type="Radio" name="radStatus" class="FormElem" Value="I" <%if trim(sActStatus)="I" then response.write "Checked" %>>InActive
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
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="button" value="Done" name="B1" class="ActionButton" onclick="window.close()">
													
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
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>
								<tr>
								    <td colspan="3" width="100%">
								        <table width="100%" class="ExcelTable" id="tblTemp">
								            <tr>
								                <td class="ExcelHeaderCell" align="center">S.No.</td>
								                <td class="ExcelHeaderCell" align="center">
								                    <img src="../../assets/images/iTMS%20icons/DeleteIcon.gif" alt="Click here to delete the template" onclick="DelActivity()">
								                </td>
								                <td class="ExcelHeaderCell" align="center">Template Name</td>
								                <td class="ExcelHeaderCell" align="center">Program Path</td>
								            </tr>
								        </table>
								    </td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
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
