<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasUOMAmendEntry.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 16, 2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	masUOMUpdate.asp
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
<HTML><HEAD><TITLE>Unit Of Measurement</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/masMasterAmendCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hUoMName" value="">
<input type=hidden name="hFlag" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Master Amendment
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
								<td class="TabCell" valign="bottom" width="75">
								  <span style="cursor: hand">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="MasCategoryAmendEntry.asp">
											<td align="center">Category
											</td></a>
										</tr>
									</table>
								  </span>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="50">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">UoM
											</td>
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
											<td>
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td class='FieldCell'>Select UoM</td>
														<td class='FieldCellSub'>																	
															<select size="1" name="selUoM" class="FormElem" onChange="GetDetails(this)">
																<option value="0">NEW UOM</option>
																<%	'Calling the Function which populates the UoM list
																	populateUoMDetails
																%>
															</select>
													    </td>
													</tr>
													<tr>
														<td class=FieldCell> UoM Short Description</td>
														<td class='FieldCellSub'><input type="text" name="txtUOMCode" size="5" maxlength=3 class="Formelem"></td>
													</tr>
													<tr>
														<td class=FieldCell> UoM Description</td>
														<td class='FieldCellSub'><input type="text" name="txtUOMName" size="50" maxlength=40 class="Formelem"></td>
													</tr>
													<tr>
														<td class=FieldCell> Decimals Allowed</td>
														<td class='FieldCellSub'>
															<input type="radio" name="radDecimal" value="Y" class="Formelem"> Yes 
															<input type="radio" name="radDecimal" value="N" class="Formelem"> No 
														</td>
													</tr>
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
																<input type="reset" value="Reset" name="B2" class="ActionButton">
																<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
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
	Function populateUoMDetails()
		' Declaration of variables
		dim dcrs,sUoMID,sUoMName,sUoMShName,sDecimalAllowed
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION,DECIMALALLOWED FROM MS_UNITOFMEASUREMENT ORDER BY UOMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		
		set sUoMID = dcrs(0)
		set sUoMName = dcrs(1)
		set sUoMShName = dcrs(2)
		set sDecimalAllowed = dcrs(3)
			
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUoMID)&"|"&trim(sUoMShName)&"|"&trim(sDecimalAllowed)&""">"&trim(sUoMName)&"</OPTION>" &vbcrlf)
		dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>
