<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AmendUserCreationDetailsEntry.asp
	'Module Name				:	Admin (User Creation Amendment)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 29, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	AmendUserCreationInsert.asp
	'Procedures/Functions Used	:	populateOrganizationUnit and populateCountry
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
<!--#include file="../../include/populate.asp"-->

<%
	dim dcrs,sSql,sCallFrom,sTempArr,sUserType
	dim sLoginID,sOrgID,sOrgName,sPassword,sFirstName,sMiddleName,sLastName,sTitle,sAlternateName
	dim sShortName,sEmpCode,sDesig,sWorkPhone,sWorkFax,sWorkEmail,sCellPhone
	dim sAddress,sCity,sState,sCountry,sPostalCode,sHomePhone,sHomeFax,sHomeEmail,sNotes
	Dim sUserAccessMode

	sCallFrom = Request.Form("hCallFrom")

	'Response.Write "<p>sCallFrom="&sCallFrom

	If sCallFrom = "FM" Then
		sLoginID = Request.Form("selUser")
		sOrgID = Request.Form("selUnit")
		sOrgName = Request.Form("hOrgName")
	Elseif sCallFrom = "FG" then
		sTempArr = Request.QueryString("sTemp")
		sLoginID = Split(sTempArr,":")(0)
		sOrgID = Split(sTempArr,":")(1)

		set dcrs=server.CreateObject("ADODB.Recordset")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID ='"& sOrgID &"' "
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		If Not dcrs.EOF Then
			sOrgName = Trim(dcrs(1))
		End IF
		dcrs.Close
	End IF

	set dcrs=server.CreateObject("ADODB.Recordset")

	sSql = "SELECT PASSWORD, USERNAME, ISNULL(MIDDLENAME, '') AS MIDDLENAME, LASTNAME, TITLE, ISNULL(ALTERNATENAME, '') AS ALTERNATENAME, ISNULL(SHORTNAME, '') AS SHORTNAME, InternalUserID, DESIGNATION, ISNULL(WORKPHONE, '') AS WORKPHONE, ISNULL(WORKFAX, '') AS WORKFAX, ISNULL(WORKEMAIL, '') AS WORKEMAIL, ISNULL(CELLPHONE, '') AS CELLPHONE, ISNULL(HOMESTREETADDRESS, '') AS HOMESTREETADDRESS, ISNULL(HOMECITY, '') AS HOMECITY, ISNULL(HOMESTATEPROVINCE, '') AS HOMESTATEPROVINCE, ISNULL(HOMECOUNTRY, '') AS HOMECOUNTRY, ISNULL(HOMEPOSTALCODE, '') AS HOMEPOSTALCODE, ISNULL(HOMEPHONE, '') AS HOMEPHONE, ISNULL(HOMEFAX, '') AS HOMEFAX, ISNULL(HOMEEMAIL, '') AS HOMEEMAIL, ISNULL(NOTES, ''),USERTYPE,UserAccessMode,UserCode FROM DCS_USER WHERE LOGINID = " & pack(sLoginID) & " AND OUDEFINITIONID = " & pack(sOrgID)
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = Nothing

	if not dcrs.EOF then
		sPassword = dcrs(0)
		sFirstName = dcrs(1)
		sMiddleName = dcrs(2)
		sLastName = dcrs(3)
		sTitle = dcrs(4)
		sAlternateName = dcrs(5)
		sShortName = dcrs(6)
		sEmpCode = dcrs(24)
		sDesig = dcrs(8)
		sWorkPhone = dcrs(9)
		sWorkFax = dcrs(10)
		sWorkEmail = dcrs(11)
		sCellPhone = dcrs(12)
		sAddress = dcrs(13)
		sCity = dcrs(14)
		sState = dcrs(15)
		sCountry = dcrs(16)
		sPostalCode = dcrs(17)
		sHomePhone = dcrs(18)
		sHomeFax = dcrs(19)
		sHomeEmail = dcrs(20)
		sNotes = dcrs(21)
		sUserType = dcrs(22)
		sUserAccessMode = dcrs(23)
	end if
	dcrs.Close
	

	if sCountry = "" then sCountry = 0
	if sCountry = "S" then sCountry = 0
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>User Amendment</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../scripts/AmendcreateUser.js"></SCRIPT>
<SCRIPT SRC="../../scripts/UserCreationCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetIndex((document.forms.formname || document.forms[0]).selCountry,<%=sCountry%>);SetUserType()">

<form method="POST" name="formname" action="">
<INPUT TYPE=HIDDEN NAME="hLoginID" VALUE="<%=sLoginID%>">
<INPUT TYPE=HIDDEN NAME="hOrgID" VALUE="<%=sOrgID%>">
<INPUT TYPE=HIDDEN NAME="hOrgName" VALUE="<%=sOrgName%>">
<INPUT TYPE=HIDDEN NAME="hCallFrom" VALUE="<%=sCallFrom%>">
<Input type=hidden name=hUserType value="<%=sUserType%>">
<input type="hidden" name="hUAM" value="<%=sUserAccessMode%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">User Amendment
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
						<td class="TabBodywithTopLine">
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
													<td class="FieldCell">Login ID </td>
													<td class="FieldCellSub"><span class="DataOnly"><%=sLoginID%> </span>
														<!--input type="text" name="txtLoginID" size="11" maxlength=10 class="FormElem"-->
													</td>
													<!--td class="FieldCellSub"></td>
													<td class="FieldCellSub">Password *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtPassword" maxlength=10 size="11" class="FormElem" value="<%=sPassword%>" >
													</td-->
												</tr>

												<tr>
													<td class="FieldCell">User Type</td>
													<td class="FieldCellSub" colspan=4>
														<Input type=Radio Name=radUserType Class=FormElem Value="AU">Application User
														<Input type=Radio Name=radUserType Class=FormElem Value="AD">Administrator
														<Input type=Radio Name=radUserType Class=FormElem Value="SU">Super User
													</td>
												</tr>

												<tr>
													<td class="FieldCell">First Name *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtFName" maxlength=50 size="22" class="FormElem" value="<%=sFirstName%>">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Middle Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtMName" maxlength=50 size="22" class="FormElem" value="<%=sMiddleName%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Last Name *</td>
													<td class="FieldCellSub" colspan=4>
														<input type="text" name="txtLName" maxlength=50 size="22" class="FormElem" value="<%=sLastName%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Title *</td>
													<td class="FieldCellSub" colspan=4>
														<input type="text" name="txtTitle" maxlength=100 size="22" class="FormElem" value="<%=sTitle%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Alternate Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtAName" maxlength=100 size="22" class="FormElem" value="<%=sAlternateName%>">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Short Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtSName" maxlength=50 size="22" class="FormElem" value="<%=sShortName%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Employee ID *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtEmployeeID" maxlength=10 size="22" class="FormElem" value="<%=sEmpCode%>" readonly>
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Designation *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtDesignation" maxlength=50 size="22" class="FormElem" value="<%=sDesig%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell" valign="top">Organization </td>
													<td class="FieldCellSub" colspan=4><span class="DataOnly"><%=sOrgName%> </span>
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Work Phone</td>
													<td class="FieldCellSub">
														<input type="text" name="txtWorkPhone" maxlength=50 size="22" class="FormElem" value="<%=sWorkPhone %>">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Work Fax</td>
													<td class="FieldCellSub">
														<input type="text" name="txtWorkFax" maxlength=50 size="22" class="FormElem" value="<%=sWorkFax %>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Work Email</td>
													<td class="FieldCellSub">
														<input type="text" name="txtWorkEmail" maxlength=50 size="22" class="FormElem" value="<%=sWorkEmail %>">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Cell Phone</td>
													<td class="FieldCellSub">
														<input type="text" name="txtCell" maxlength=50 size="22" class="FormElem" value="<%=sCellPhone%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Street Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtStreet" maxlength=50 size="22" class="FormElem" value="<%=sAddress%>">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">City</td>
													<td class="FieldCellSub">
														<input type="text" name="txtCity" maxlength=50 size="22" class="FormElem" value="<%=sCity%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">State</td>
													<td class="FieldCellSub" colspan=4>
														<input type="text" name="txtState" maxlength=50 size="22" class="FormElem" value="<%=sState%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Country</td>
													<td class="FieldCellSub">
														<select size="1" name="selCountry" class="FormElem" >
															<option value="select">Select</option>
															<%	'Calling the Function which populates the Country list
																populateCountry
															%>
														</select>
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Postal Code</td>
													<td class="FieldCellSub">
														<input type="text" name="txtPostal" maxlength=50 size="22" class="FormElem" value="<%=sPostalCode%>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Home Phone</td>
													<td class="FieldCellSub">
														<input type="text" name="txtHomePhone" maxlength=50 size="22" class="FormElem" value="<%=sHomePhone %>">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Home Fax</td>
													<td class="FieldCellSub">
														<input type="text" name="txtHomeFax" maxlength=50 size="22" class="FormElem" value="<%=sHomeFax %>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Home Email</td>
													<td class="FieldCellSub" colspan=4>
														<input type="text" name="txtHomeEmail" maxlength=50 size="22" class="FormElem" value="<%=sHomeEmail %>">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">Notes</td>
													<td class="FieldCellSub" colspan=4>
														<textarea name="txtNotes" maxlength=200 rows=5 col=50 class="FormElem" ><%=sNotes%></textarea>
													</td>
												</tr>

											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack"></td>
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
													<input type="button" value="Amend" name="B3" class="ActionButton" onClick="CheckSubmit()">
													<input type="button" value="Delete" name="B2" class="ActionButton" onClick="CheckDelete()">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
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
