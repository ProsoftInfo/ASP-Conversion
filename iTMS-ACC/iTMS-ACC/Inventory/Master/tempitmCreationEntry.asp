
<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	tempitmCreationEntry.asp
	'Module Name				:	Inventory (Temporary Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 10, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	tempitmCreationInsert.asp
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
<HTML><HEAD><TITLE>Temporary Item Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="Data">
<root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<root/>
</script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/tempitmCreate.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%
	dim iAppCode,iModCode,sCreationStage,sItemtype,sOrgCode
	iAppCode = trim(Request("appCode"))
	iModCode = trim(Request("modCode"))
	sCreationStage = trim(Request("creStage"))
	sItemtype = trim(Request("itmType"))
	sOrgCode = trim(Request("sOrgCode"))
%>
<form method="POST" name="formname" action="">
<input type=hidden name="hAppCode" value="<%=iAppCode%>">
<input type=hidden name="hModCode" value="<%=iModCode%>">
<input type=hidden name="hCreStage" value="<%=sCreationStage%>">
<input type=hidden name="hOrgCode" value="<%=sOrgCode%>">
<input type=hidden name="hItemType" value="<%=sItemtype%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Temporary Item Creation</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td width="100%">
												<table cellpadding="0" cellspacing="0">
												<!--	<tr>
														<td class='FieldCell'>Item Type </td>
														<td class='FieldCellSub'>
															<select size="1" name="selItmType" class="FormElem" onChange="LetIType(this)">
																<option value="select">Select</option>
																<%	'Calling the Function which populates the Item Type list
																	'populateItemTypeSelected sItemtype
																%>
															</select>
													    </td>
													</tr>-->
													<tr>
														<td class=FieldCell> Item Code</td>
														<td class='FieldCellSub'>
														<%	if sItemtype = "YRN" then %>
															<input type="text" name="itmCode" size="19" maxlength=15 class="Formelem" readonly>
                                                            <input type="button" value="Yarn Code" name="btnYrnCode" class="AddButton" onClick="CreateItemCode(this)">
                                                        <%	else %>
															<input type="text" name="itmCode" size="19" maxlength=15 class="Formelem">
                                                            <input type="button" value="Yarn Code" name="btnYrnCode" class="AddButton" onClick="CreateItemCode(this)" disabled>
                                                        <%	end if %>
                                                        </td>
													</tr>
													<tr>
														<td class=FieldCell> Description</td>
														<td class='FieldCellSub'><input type="text" name="txtItmDesc" size="60" maxlength=60 class="Formelem"></td>
													</tr>
													<tr>
														<td class=FieldCell> Short Description</td>
														<td class='FieldCellSub'><input type="text" name="txtItmShDesc" size="25" maxlength=25 class="Formelem"></td>
													</tr>
													<tr>
														<td class=FieldCell> Additional Description</td>
														<td class='FieldCellSub'><input type="text" name="txtItmAddDesc" size="80" maxlength=250 class="Formelem"></td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="MiddlePack" width="100%">
											</td>
                                        </tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Save" name="B1" class="ActionButton" onClick="CheckSubmit()">
																<input type="reset" value="Reset" name="B2" class="ActionButton">
														</td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="BottomPack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
									</table>
								</td>
								<td align="center">
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

