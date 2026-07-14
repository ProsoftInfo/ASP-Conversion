<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmEditOpUoMEntry.asp
	'Module Name				:	Inventory (Item creation / Definition)
	'Author Name				:	S.MAHESHWARI
	'Created On					:	September 10, 2007
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - UoM Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	dim sUoM

	sUoM = Request.QueryString("UOM")

%>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemOpUom.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function CheckUoMEntered() {
	iType = ""
	if ((iType != "PLA") && (document.forms[0].selUoMPurchase.selectedIndex != "0") && (document.forms[0].selUoMPurchase.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToPur.value) == "")) {
		alert("Enter Stores To Purchase Conversion");
		document.forms[0].txtStToPur.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMPurchase.selectedIndex != "0") && (document.forms[0].selUoMPurchase.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToPur.value) != "") && (isNaN(document.forms[0].txtStToPur.value))) {
		alert("Only Numerals are Allowed");
		document.forms[0].txtStToPur.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMPurchase.selectedIndex != "0") && (document.forms[0].selUoMPurchase.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToPur.value) != "") && (document.forms[0].selStToPur.selectedIndex == "0")) {
		alert("Select Stores To Purchase Operator");
		document.forms[0].selStToPur.focus();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMManu.selectedIndex != "0") && (document.forms[0].selUoMManu.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToManu.value) == "")) {
		alert("Enter Stores To Manufacturing Conversion");
		document.forms[0].txtStToManu.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMManu.selectedIndex != "0") && (document.forms[0].selUoMManu.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToManu.value) != "") && (isNaN(document.forms[0].txtStToManu.value))) {
		alert("Only Numerals are Allowed");
		document.forms[0].txtStToManu.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMManu.selectedIndex != "0") && (document.forms[0].selUoMManu.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToManu.value) != "") && (document.forms[0].selStToManu.selectedIndex == "0")) {
		alert("Select Stores To Manufacturing Operator");
		document.forms[0].selStToManu.focus();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMSales.selectedIndex != "0") && (document.forms[0].selUoMSales.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToSales.value) == "")) {
		alert("Enter Stores To Sales Conversion");
		document.forms[0].txtStToSales.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMSales.selectedIndex != "0") && (document.forms[0].selUoMSales.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToSales.value) != "") && (isNaN(document.forms[0].txtStToSales.value))) {
		alert("Only Numerals are Allowed");
		document.forms[0].txtStToSales.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMSales.selectedIndex != "0") && (document.forms[0].selUoMSales.value != document.forms[0].selUoMStores.value) && (trimTrue(document.forms[0].txtStToSales.value) != "") && (document.forms[0].selStToSales.selectedIndex == "0")) {
		alert("Select Stores To Sales Operator");
		document.forms[0].selStToSales.focus();
		return false;
	}
	else
		return true;
}

//-->
</SCRIPT>

</HEAD>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">UoM Details
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
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<table border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td valign="top">
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class='GroupTitleLeft' width="10">&nbsp;</td>
																	<td class='GroupTitle' width="147"><p align="center">UoM &amp; Basic Conversion</td>
																	<td class='GroupTitleRight'><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class=MiddlePack colspan="3"> </td>
																</tr>
																<tr>
																	<td class=FieldCellSub> Stores</td>
																	<td class='FieldCellSub'>
																	<select size="1" name="selUoMStores" class="FormElem" DISABLED>
																		<option value="select">Select</option>
																		<option value="<%=sUoM%>" selected><%=sUoM%></option>
															        </select>
															        </td>
																	<td class='FieldCellSub'>
															        </td>
																</tr>
																<tr>
																	<td class=FieldCellSub> Purchase</td>
																	<td class='FieldCellSub'>
																		<select size="1" name="selUoMPurchase" class="FormElem">
																			<option value="select">Select</option>
																			<%	'Calling the Function which populates the UoM list
																				populateUoM
																			%>
																		</select>
																	</td>
																	<td class='FieldCellSub'>
																	</td>
																</tr>
																<tr>
																	<td class=FieldCellSub> Manufacturing</td>
																	<td class='FieldCellSub'>
																	<select size="1" name="selUoMManu" class="FormElem">
																		<option value="select">Select</option>
																		<%	'Calling the Function which populates the UoM list
																			populateUoM
																		%>
																	</select>
																	</td>
																	<td class='FieldCellSub'></td>
																</tr>
																<tr>
																	<td class=FieldCellSub> Sales</td>
																	<td class='FieldCellSub'>
																		<select size="1" name="selUoMSales" class="FormElem">
																			<option value="select">Select</option>
																			<%	'Calling the Function which populates the UoM list
																				populateUoM
																			%>
																		</select>
																	</td>
																	<td class='FieldCellSub'></td>
																</tr>
															</table>
                                                        </td>
													</tr>
												</table>
                                            </td>
                                            <td>
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                            </td>
                                            <td valign="bottom">
                                                <table border="0" cellspacing="1" class="TableOutlineOnly" width="100%">
													<tr>
														<td class="ExcelHeaderCell" width="140"><p align="center">Conversion</p></td>
														<td class="ExcelHeaderCell" width="50"><p align="center">Factor</td>
														<td class="ExcelHeaderCell"><p align="center">Operator</td>
													</tr>
													<tr>
														<td class="MiddlePack" colspan="3"></td>
													</tr>
													<tr>
														<td class="FieldCellSub" width="140">Stores to Purchase</td>
														<td width="50"><input type="text" name="txtStToPur" size="8" maxlength=9 class="Formelem"></td>
														<td class="FieldCell">&nbsp;
															<select size="1" name="selStToPur" class="FormElem">
																<option value="select">Select</option>
																<option value="0">*</option>
																<option value="1">/</option>
														    </select>
														</td>
													</tr>
													<tr>
														<td class="FieldCellSub" width="140">Stores to Manufacturing</td>
														<td width="50"><input type="text" name="txtStToManu" size="8" maxlength=9 class="Formelem"></td>
														<td class="FieldCell">&nbsp;
															<select size="1" name="selStToManu" class="FormElem">
																<option value="select">Select</option>
																<option value="0">*</option>
																<option value="1">/</option>
														    </select>
														</td>
													</tr>
													<tr>
														<td class="FieldCellSub" width="140">Stores to Sales</td>
														<td width="50"><input type="text" name="txtStToSales" size="8" maxlength=9 class="Formelem"></td>
														<td class="FieldCell">&nbsp;
															<select size="1" name="selStToSales" class="FormElem">
																<option value="select">Select</option>
																<option value="0">*</option>
																<option value="1">/</option>
														    </select>
														</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Alternate UoM For</td>
											<td class="FieldCellSub">
												<select size="1" name="selFor" class="FormElem">
													<option value="select">Select</option>
													<option value="P">Purchase</option>
													<option value="S">Sales</option>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Enter few characters</td>
											<td class="FieldCellSub">
												<input type="text" name="txtSearch" size="11" class="Formelem"  ONKEYUP="javascript:selectTheItem(this,'selItem')">
											</td>
										</tr>
										<tr>
										    <td class="FieldCell" valign="top"> Unit of Measurement</td>
										    <td class="FieldCellSub">
												<select size="5" name="selItem" class="FormElem">
											<%	'Calling the Function which populates UoM List
												populateUoMList
											%>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Factor</td>
											<td class="FieldCellSub">
												<input type="text" name="txtFactor" size="8" maxlength=9 class="FormElem">
											</td>
										</tr>
										<tr>
										    <td class="FieldCell">Operator</td>
										    <td class="FieldCellSub">
												<select size="1" name="selOpe" class="FormElem">
													<option value="select">Select</option>
													<option value="0">*</option>
													<option value="1">/</option>
												</select>&nbsp;
												<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
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
								<td width="100%">
									<div class="frmBody" id="frm2" style="width: 100%; height:90;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Alternate UoM</td>
												<td class="ExcelHeaderCell" align="center" width="100">Factor</td>
												<td class="ExcelHeaderCell" align="center" width="60">Operator</td>
												<td class="ExcelHeaderCell" align="center" width="60">For</td>
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
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="window.close()">
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
</BODY>
</HTML>
<%
	' Function to populate UoM List
	Function populateUoMList()
		' Declaration of variables
		Dim dcrs,sUomDesc,sUomShDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT ORDER BY UOMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUoMCode = dcrs(0)
		set sUomDesc = dcrs(1)
		set sUomShDesc = dcrs(2)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUoMCode)&""">"&trim(sUomShDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
