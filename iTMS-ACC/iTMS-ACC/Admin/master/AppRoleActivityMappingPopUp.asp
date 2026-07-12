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
<script type="application/xml" data-itms-xml-island="1" ID="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" ID=PRData></script>
<script type="application/xml" data-itms-xml-island="1" ID="RetData"><Root Done=""/></script>
<script type="application/xml" data-itms-xml-island="1" id="ActivityData"><Root Eligible="N"></Root></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AppRoleActivityMappingPopupCompat.js"></SCRIPT>
<SCRIPT>
ITMSAppRoleActivityMappingPopupCompat.install();
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
														<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record"  onclick="DeleteItem()" width="15" height="15"></a>
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
