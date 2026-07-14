<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasCategoryAddEditPop.asp
	'Module Name				:	Inventory (Master)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 16, 2003
	'Modified By                :   Ragavendran R
	'Modified On				:   Jul 22,2011
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Category</TITLE>
<base target="_self"></base>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT>
function trimValue(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function checkSubmit() {
	var form = document.formname;
	form.hFlag.value = form.selCategory.selectedIndex === 0 ? "N" : "A";

	if (trimValue(form.txtCatName.value) === "") {
		alert("Enter Category Name");
		form.txtCatName.select();
		return false;
	}
	if (trimValue(form.txtCatShName.value) === "") {
		alert("Enter Category Short Name");
		form.txtCatShName.select();
		return false;
	}

	form.action = "MasCategoryAddEditPopInsert.asp";
	form.submit();
	return true;
}

function GetDetails(obj) {
	var form = document.formname;
	if (form.selCategory.selectedIndex === 0) {
		form.hCatCode.value = "";
		form.txtCatShName.value = "";
		form.txtCatName.value = "";
		form.chkEligible.checked = false;
		return false;
	}

	var arrTemp = obj.value.split("|");
	form.hCatCode.value = trimValue(arrTemp[0]);
	form.txtCatShName.value = trimValue(arrTemp[1]);
	form.chkEligible.checked = trimValue(arrTemp[2]) === "1";
	form.txtCatName.value = trimValue(form.selCategory.options[form.selCategory.selectedIndex].text);
	return true;
}

function Delete() {
	var form = document.formname;
	if (form.selCategory.selectedIndex === 0) {
		alert("Select Category");
		form.selCategory.focus();
		return false;
	}

	form.hCatName.value = trimValue(form.selCategory.options[form.selCategory.selectedIndex].text);
	form.action = "CategoryDeletionUpdate.asp";
	form.submit();
	return true;
}
</SCRIPT>

</HEAD>
 
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hCatName" value="">
<input type=hidden name="hFlag" value="">
<input type=hidden name="hCatVal" value="">
<input type="hidden" name="hCatCode" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Master Amendment</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td width="100%">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="FieldCellSub">Select Category</td>
															<td class="FieldCellSub">																	
																<select size="1" name="selCategory" class="FormElem" onChange="GetDetails(this)">
																	<option value="0">NEW CATEGORY</option> 
																	<%	'Calling the Function which populates the Category list
																		populateCategory
																	%>
																</select>
														    </td>
														</tr>
														<!--<tr>
															<td class="FieldCellSub" width="95"> Category Code</td>
															<td class="FieldCellSub"><input type="text" name="txtCatCode" size="5" maxlength=3 class="Formelem"></td>
														</tr>-->
														<tr>
															<td class="FieldCellSub" width="95"> Name</td>
															<td class="FieldCellSub"><input type="text" name="txtCatName" size="55" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class="FieldCellSub" width="95"> Short Name</td>
															<td class="FieldCellSub"><input type="text" name="txtCatShName" size="12" maxlength=10 class="Formelem"></td>
														</tr>
														<tr>
														    <td class="FieldCellSub" ></td>
															<td class="FieldCellSub" >
															    <input type="checkbox" name="chkEligible" value="1">&nbsp;Eligible For Web-Store
															</td>
														</tr>
												</center>
													</table>
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
                                                                <input type="button" value="Save" name="B1" class="ActionButton" onClick="checkSubmit()">
                                                                <input type="button" value="Delete" class="ActionButton" onClick="Delete()">
																<input type="button" value="Close" name="B1" class="ActionButton" onClick="window.close()">
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

<%
	' Function to populate Category
	Function populateCategory()
		' Declaration of variables
		dim dcrs,sCatID,sCatName,sCatShName,sEligible
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CATEGORYCODE,CATEGORYNAME,CATEGORYSHORTNAME,IsNull(EligibleForWebStore,'0') FROM INV_M_CLASSIFICATIONCATEGORY ORDER BY CATEGORYCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		
		set sCatID = dcrs(0)
		set sCatName = dcrs(1)
		set sCatShName = dcrs(2)
		set sEligible = dcrs(3)
			
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sCatID)&"|"&trim(sCatShName)&"|"&trim(sEligible)&""">"&trim(sCatName)&"</OPTION>" &vbcrlf)
		dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>
