<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	UpdateMasterXMLEntry.asp
	'Module Name				:	Admin (Update Master XML Display)
	'Author Name				:	TAJUDEEN
	'Created On					:	May 04, 2004
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	UpdateMasterXML.asp
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
<HTML><HEAD><TITLE>Update XML</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
	function CheckSubmit()
	{
		if (!((document.formname.ChkDivision.checked) || (document.formname.ChkOrganization.checked) || (document.formname.ChkUnit.checked)))
		{
			alert ("Select any one");
			document.formname.ChkDivision.focus();
			return false;
		}
		else
			document.formname.submit();
	}

</Script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname" action="UpdateMasterXML.asp">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Update XML</p>
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
											<td width="100%" align="center">
													<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width=50%>
														<tr>
															<td class="ExcelHeaderCell" align="Center" width="10">Select</td>
															<td class="ExcelHeaderCell" align="Center" width="200">Update For</td>
														</tr>
														<tr>
															<td class='ExcelDisplayCell' align="center"><input type="Checkbox" name="ChkDivision" class="Formelem" value="1"></td>
															<td class='ExcelDisplayCell'> Division</td>
														</tr>
														<tr>
															<td class='ExcelDisplayCell' align="center"><input type="Checkbox" name="ChkOrganization" class="Formelem" value="1"></td>
															<td class='ExcelDisplayCell'> Organization</td>
														</tr>
														<tr>
															<td class='ExcelDisplayCell' align="center"><input type="Checkbox" name="ChkUnit" class="Formelem" value="1"></td>
															<td class='ExcelDisplayCell'> Unit</td>
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
                                                                <input type="button" value="Update" name="B2" class="ActionButton" onClick="javascript:CheckSubmit()">
																<input type="reset" value="Cancel" name="B1" class="ActionButton">
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

