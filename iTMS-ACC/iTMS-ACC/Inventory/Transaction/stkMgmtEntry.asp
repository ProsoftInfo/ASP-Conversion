<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtEntry.asp
	'Module Name				:	Inventory (Stock Management)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 27, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	stkMgmtDetailsEntry.asp
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
<!-- #include File="../../include/populate.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock Management</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Output/>
</script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/stkMgmtEntry.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%
	if CheckFinYear(formatDate(date())) <> "0" then
%>
<SCRIPT LANGUAGE=javascript>
	alert("Since Year End closing has been done / Transaction date entered is in current Financial Year, this transaction cannot be performed for this current Financial Year.")
	window.location.href = "../welcome_Inventory.asp"
</SCRIPT>
<%
	Response.End
	end if
%>

<form method="POST" name="formname" action="">
<input type=hidden name="hSelectedValue" value="">
<input type=hidden name="hItemNames" value="">
<input type=hidden name="hClassName" value="">
<input type=hidden name="hOrgName" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Stock Management</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Header</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td width="100%" align="center">Control</td>
										</tr>
									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    <font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
										<tr>
                                            <td class="FieldCell">Select Organization</td>
                                            <td class="FieldCellSub" width="175">
												<select size="1" name="selUnit" class="FormElem" >
													<option value="select">Select</option>
													<%	'Calling the Function which populates Organization Unit list
														populateUnit
													%>
												</select>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Select Classification</td>
                                            <td class="FieldCellSub">
												<select size="1" name="selClass" class="FormElem" onChange="popItmDisplay()">
													<option value="select">Select</option>
												</select>
											<input type="button" value="Select" name="btnSelect" class="AddButton" onClick="AddClass()">
											</td>
                                        </tr>
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <table border="0" cellspacing="1" width="100%" class="TableOutlineOnly">
                                        <tr>
                                            <td colspan="2" class="TableHeader" width="50%"><p align="center">Enter few characters to select&nbsp;
												<input TYPE="TEXT" NAME="txtSearch" VALUE SIZE="10" ONKEYUP="selectTheItem(this,'selFrombox')" class="formelem">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="TableHeader"><p align="center">Select Item</td>
                                            <td class="TableHeader" align="center">Selected Item</td>
                                        </tr>
                                        <tr>
                                            <td width="50%" class="TableInput"><p align="center">
												<select size="15" name="selFrombox" multiple class="FormElem">
												</select>
											</td>
                                            <td width="50%" class="TableInput"><p align="center">
												<select size="15" name="selTobox" multiple class="FormElem">
												</select>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="TableFooter" width="50%"><p align="center">
												<input type="button" value="Add >>" name="add" ONCLICK="addclick('selTobox','selFrombox','remove')" class="AddButton">
											</td>
                                            <td class="TableFooter" width="50%"><p align="center">
												<input type="button" value="<< Remove" name="remove" ONCLICK="removeclick('selTobox','selFrombox','remove')" class="AddButton">
                                            </td>
                                        </tr>
                                    </table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
                            <tr>
                                <td class="FieldCellSub" colspan="3">Select Control &nbsp;
									<select size="1" name="selControl" class="FormElem">
										<option value="select">Select</option>
										<option value="stkMgmtSMEntry.asp">Stock Mgt.</option>
										<option value="stkMgmtPAEntry.asp">Physical Adj.</option>
										<!--option value="RE">Returns</option-->
										<option value="stkMgmtSTEntry.asp">Stock Transfer</option>
									</select>
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Next" name="next" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="reset" value="Reset" name="B1" class="ActionButton">
                                                    <input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
</BODY>
</HTML>
