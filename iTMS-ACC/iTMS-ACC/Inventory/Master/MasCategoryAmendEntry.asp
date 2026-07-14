<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasCategoryAmendEntry.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 16, 2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	masCategoryUpdate.asp
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
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/masMasterAmendCompat.js"></SCRIPT>
</HEAD>
 
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hCatName" value="">
<input type=hidden name="hFlag" value="">
<input type=hidden name="hCatVal" value="">
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
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" bordercolor="#000000">
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="75">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Category
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="50">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)" height="13">
										<tr><a href="MasUOMAmendEntry.asp">
											<td align="center">UoM
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="100">
								  <span style="cursor: hand">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="MasClassificationAmendEntry.asp">
									  <td align="center">Classification</td></a>
									</tr>
								  </table>
								  </span>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
									<table border="0" cellpadding="0" cellspacing="0" width="20" class="TabTableEnd">
										<tr>
											<td width="100%" valign="bottom">
												<p align="center"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font></p>
											</td>
										</tr>
									</table>
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
															<td class='FieldCell'>Select Category</td>
															<td class='FieldCellSub'>																	
																<select size="1" name="selCategory" class="FormElem" onChange="GetDetails(this)">
																	<option value="0">NEW CATEGORY</option> 
																	<%	'Calling the Function which populates the Category list
																		populateCategory
																	%>
																</select>
														    </td>
														</tr>
														<tr>
															<td class=FieldCell width="95"> Category Code</td>
															<td class='FieldCellSub'><input type="text" name="txtCatCode" size="5" maxlength=3 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="95"> Name</td>
															<td class='FieldCellSub'><input type="text" name="txtCatName" size="55" maxlength=50 class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="95"> Short Name</td>
															<td class='FieldCellSub'><input type="text" name="txtCatShName" size="12" maxlength=10 class="Formelem"></td>
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
																<input type="reset" value="Reset" name="B1" class="ActionButton">
																<input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
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
		dim dcrs,sCatID,sCatName,sCatShName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CATEGORYCODE,CATEGORYNAME,CATEGORYSHORTNAME FROM INV_M_CLASSIFICATIONCATEGORY ORDER BY CATEGORYCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		
		set sCatID = dcrs(0)
		set sCatName = dcrs(1)
		set sCatShName = dcrs(2)
			
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sCatID)&"|"&trim(sCatShName)&""">"&trim(sCatName)&"</OPTION>" &vbcrlf)
		dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>
