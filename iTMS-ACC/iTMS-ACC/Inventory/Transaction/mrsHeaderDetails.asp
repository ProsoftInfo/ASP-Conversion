<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	mrsHeaderDetails.asp
	'Module Name				:	Inventory (MRS Header Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	August 27, 2003
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/SourceReferenceDetails.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>MR Header Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
	function CheckSubmit() {
		document.formname.action = document.formname.hsAct.value
		document.formname.submit()
	}
//-->
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%
	dim dcrs,sMRNo,sAction,sSelected,sIssueCode
	dim sUnit,sUnitName,sType,sItmType,sDate,sDept,sUsage,sRemarks,sItmTypeName
	sMRNo = trim(Request("mrs"))
	sAction = trim(Request("sAct"))
	'Response.Write sAction
	sSelected = trim(Request("hAction"))
	sItmType = "-"
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT DISTINCT CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,DEPTNAME,ITEMTYPENAME,ISNULL(REMARKS,''),MRSTYPE,MRSFORUNIT FROM VWMRSHEADERDETAILS WHERE MRSNUMBER = " & sMRNo & ""
		.Source = "SELECT DISTINCT CONVERT(CHAR,MRSDATE,103),ISSUEDFORDESCRIPTION,ORGUNITSHORTDESCRIPTION,ITEMTYPENAME,ISNULL(ITEMTYPENAME,'-'),ISNULL(REMARKS,''),MRSTYPE,MRSFORUNIT,ISNULL(MRSCODE,MRSNUMBER),ItemTypeID FROM VWMRSHEADERDETAILS WHERE MRSNUMBER = " & sMRNo & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sDate = trim(dcrs(0))
		sUsage = trim(dcrs(1))
		sUnitName = trim(dcrs(2))
		sDept = trim(dcrs(3))
		sItmTypeName = trim(dcrs(4))
		sRemarks = trim(dcrs(5))
		if trim(dcrs(6)) = 0 then
			sType = "Returnable"
		else
			sType = "Non Returnable"
		end if
		sUnit = trim(dcrs(7))
		sIssueCode = trim(dcrs(8))
		sItmType = trim(dcrs(9))
	end if
	dcrs.close
%>
<form method="POST" name="formname" action="">
<input type="hidden" name="mrs" value="<%=sMRNo%>">
<input type="hidden" name="hsAct" value="<%=sAction%>">
<input type="hidden" name="hUnit" value="<%=sUnit%>">
<input type="hidden" name="hItmType" value="<%=sItmType%>">
<input type="hidden" name="hAction" value="<%=sSelected%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Material Requisition Header Details
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
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td width="100%">
                                                <table border="0" cellpadding="0" cellspacing="0">
													<tr>
													    <td class="FieldCell">Material Requisition Number&nbsp;</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sIssueCode%>&nbsp;</span></td>
													</tr>
													<tr>
													    <td class="FieldCell">Requisition by Unit&nbsp;</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sUnitName%>&nbsp;</span></td>
													    <td class="FieldCell">Requisition Date</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sDate%>&nbsp;</span></td>
													</tr>
													<tr>
													    <td class="FieldCell">Requisition Type</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sType%>&nbsp;</span></td>
													    <!--td class="FieldCell">For Department</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sDept%>&nbsp;</span></td-->
													    <td class="FieldCell">Usage of Item</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sUsage%>&nbsp;</span></td>
													</tr>
													<tr>
													    <td class="FieldCell">Item Type</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sItmTypeName%>&nbsp;</span></td>
													    <!--td class="FieldCell">Usage of Item</td>
													    <td class="FieldCellSub"><span class="DataOnly"><%=sUsage%>&nbsp;</span></td-->
													</tr>
													<tr>
														<td class="FieldCell" valign="top" rowspan="4">Item Classification</td>
														<td class="FieldCellSub" rowspan="4">
															<select size="6" name="selClass" class="FormElem" multiple disabled>
														<%	populateClassification() %>
															</select>
														</td>
													    <!--td class="FieldCell">Ship to Location</td>
													    <td class="FieldCellSub">
															<select size="1" name="selShip" class="FormElem">
																<option value="select">Select</option>
																<option value="S">Single</option>
																<option value="M">Multiple</option>
														    </select>
														</td-->
													</tr>
													<!--tr>
													    <td class="FieldCell">Cost Center</td>
													    <td class="FieldCellSub">
														</td>
													</tr>
													<tr>
														<td class="FieldCell">Additional Details</td>
														<td class="FieldCell">
															<input type="radio" value="1" name="radAdd" class="FormElem"> Yes&nbsp;
															<input type="radio" value="0" name="radAdd" class="FormElem" checked> No
														</td>
													</tr-->
													<tr>
														<td class="FieldCell">Remarks</td>
														<td class="FieldCell"><textarea rows="3" name="txtRemarks" value="<%=sRemarks%>" maxlength=50 cols="35" class="FormElem" readonly></textarea></td>
													</tr>
													<tr>
														<td class="FieldCell" valign="top"></td>
														<td class="FieldCell"></td>
													</tr>
												<%
													' Function call to insert the display part in case of WO
													'DisplayWODetail sMRNo,"INV"
												%>
												<%
													' Function call to insert the display part in case of MRP
													'DisplayMRPDetail sMRNo,"INV"
												%>
                                                </table>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Next" name="B7" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Cancel" name="B8" class="ActionButton" onClick="window.location.href='../welcome_Inventory.asp'">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
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
<%
	' Function to populate Classification
	Function populateClassification()
		' Declaration of variables
		Dim dcrs,sDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT GROUPNAME FROM VWMRSITEMDETAILS WHERE MRSNUMBER = " & sMRNo & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sDesc = dcrs(0)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sDesc)&""">"&trim(sDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>

