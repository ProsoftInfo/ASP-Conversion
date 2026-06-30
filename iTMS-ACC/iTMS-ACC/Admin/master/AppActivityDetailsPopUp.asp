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
	Dim sTemp,sArr,sSql,nRoleID,nProcessCode,sProcessName,nPracticeCode,sType
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
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/AppActivityDetailsPopupCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
ITMSAppActivityDetailsPopupCompat.install();
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
		<Input type="hidden" name="hNoOfRows" value="">

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
													<input type="Button" value="Save" name="B1" class="ActionButton" onclick="SaveData('<%=sType%>')">
													<input type="button" value="Close" name="B2" class="ActionButton" onClick="ClosePopup()">
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
