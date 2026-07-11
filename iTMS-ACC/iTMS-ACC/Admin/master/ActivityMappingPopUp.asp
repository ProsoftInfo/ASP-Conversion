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
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/ActivityMappingPopupCompat.js"></SCRIPT>
<SCRIPT>
ITMSActivityMappingPopupCompat.install();
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
													<input type="button" value="Done" name="B1" class="ActionButton" onclick="DoneClose()">
													
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
