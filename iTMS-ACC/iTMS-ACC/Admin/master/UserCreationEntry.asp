<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	UserCreationEntry.asp
	'Module Name				:	Admin (User Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 29, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	UserCreationInsert.asp
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
Dim sUserAccessMode
sUserAccessMode = Request("UAM")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>User Creation</title>

<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../scripts/createUser.js"></SCRIPT>
<XML id="EMPADDXML"><Root></Root></XML>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/UserCreationCompat.js"></SCRIPT>

</head>

<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<form method="POST" name="formname" action="UserCreationInsert.asp">
		<Input type=hidden name=hUserType value="">
		<input type="hidden" name="hUAM" value="<%=sUserAccessMode%>">
		<input type="hidden" name="hCode" value="" />
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">User Creation (
				<%
				    if trim(sUserAccessMode)="I" then
				        Response.write "Internal"
				    else
				        Response.write "External"
				    end if
				%>
				)
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
													<td class="FieldCell">Employee ID *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtEmployeeID" size="22" class="FormElemRead" readonly="true">&nbsp;
														<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" width="11" height="11" alt="Select Supplier" style="cursor: pointer" onclick="SelectEmployee()">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Login ID *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtLoginID" maxlength=20 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub">Password *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtPassword" maxlength=20 size="22" class="FormElem">
													</td>
												</tr>

												<tr>
													<td class="FieldCell">User Type</td>
													<td class="FieldCellSub" colspan=4>
														<Input type=Radio Name=radUserType Class=FormElem Value="AU" checked>Application User
														<Input type=Radio Name=radUserType Class=FormElem Value="AD">Administrator
														<Input type=Radio Name=radUserType Class=FormElem Value="SU">Super User
													</td>
													<td class="FieldCellSub">Organisation *</td>
													<td class="FieldCellSub">
														<select size="1" name="selUnit" class="FormElem">
															<option value="S">Select</option>
															<%	'Calling the Function which populates the Organization Units Definition list
																populateUnit
															%>
														</select>
													</td>
												</tr>

												<tr>
													<td class="FieldCell">First Name *</td>
													<td class="FieldCellSub">
														<input type="text" name="txtFName" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Middle Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtMName" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub">Last Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtLName" maxlength=50 size="22" class="FormElem">
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Title</td>
													<td class="FieldCellSub">
														<input type="text" name="txtTitle" maxlength=100 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Alternate Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtAName" maxlength=100 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub">Short Name</td>
													<td class="FieldCellSub">
														<input type="text" name="txtSName" maxlength=50 size="22" class="FormElem">
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Address 1</td>
													<td class="FieldCellSub">
														<input type="text" name="txtStreet" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Address 2</td>
													<td class="FieldCellSub">
														<input type="text" name="txtAddr2" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub">City</td>
													<td class="FieldCellSub">
														<input type="text" name="txtCity" maxlength=50 size="22" class="FormElem">
													</td>
												</tr>
												<tr>
													<td class="FieldCell">State</td>
													<td class="FieldCellSub">
														<input type="text" name="txtState" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Country</td>
													<td class="FieldCellSub">
														<select size="1" name="selCountry" class="FormElem">
															<option value="S">Select</option>
															<%	'Calling the Function which populates the Country list
																populateCountry
															%>
														</select>
													</td>
													<td class="FieldCellSub">Postal Code</td>
													<td class="FieldCellSub">
														<input type="text" name="txtPostal" maxlength=50 size="22" class="FormElem">
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Work Phone</td>
													<td class="FieldCellSub">
														<input type="text" name="txtWorkPhone" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Work Fax</td>
													<td class="FieldCellSub">
														<input type="text" name="txtWorkFax" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub">Work Email</td>
													<td class="FieldCellSub">
														<input type="text" name="txtWorkEmail" maxlength=50 size="22" class="FormElem">
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Cell Phone</td>
													<td class="FieldCellSub">
														<input type="text" name="txtCell" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub"></td>
													<td class="FieldCellSub">Home Phone</td>
													<td class="FieldCellSub">
														<input type="text" name="txtHomePhone" maxlength=50 size="22" class="FormElem">
													</td>
													<td class="FieldCellSub">Home Email</td>
													<td class="FieldCellSub">
														<input type="text" name="txtHomeEmail" maxlength=50 size="22" class="FormElem">
													</td>
												</tr>
												<tr>
													<td class="FieldCell">Notes</td>
													<td class="FieldCellSub" colspan=4>
														<textarea name="txtNotes" maxlength=200 rows=2 col=200 class="FormElem"></textarea>
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
													<input type="button" value="Create" name="B2" class="ActionButton" onClick="CheckSubmit()">
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
