<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TransferClosingEntry.asp
	'Module Name				:	Transfer Closing Values
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 12, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	TransferClosingDetailsEntry.asp
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
<HTML><HEAD><TITLE>Transfer Closing</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../Scripts/AdminTransferClosingCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
window.ITMSAdminTransferClosingCompat.installCloseEntry();
</SCRIPT>
</HEAD>
<%
	dim sWho,sFor
	sWho = Request.QueryString("sWho")
	sFor = Request.QueryString("Frm")
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgName" value="">
<input type=hidden name="hCFinStartDate" value="">
<input type=hidden name="hCFinEndDate" value="">
<input type=hidden name="hWho" value="<%=sWho%>">
<input type=hidden name="hFor" value="<%=sFor%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td align="center" class=PageTitle height="20">
			<p align="center">Transfer Closing
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
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%" align="left">
                                    <table BORDER="0" CELLSPACING="1" CELLPADDING="0">
                                        <tr>
											<td class="FieldCell" valign="top">For Unit</td>
											<td class="FieldCellSub" valign="top">
												<select size="1" name="selUnit" class="FormElem">
													<option value="select">Select</option>
														<%	'Calling the Function which populates the Organization Units list
															populateUnit
														%>
												</select>
											</td>
                                        </tr>
                                    </table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class='GroupTitleLeft' width="10">&nbsp;</td>
														<td class='GroupTitle' width="147"><p align="center">Transfer Closing Values</td>
														<td class='GroupTitleRight'><p align="left">&nbsp;</td>
													</tr>
												</table>
                                            </td>
										</tr>
										<tr>
											<td class=GroupTable>
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td class=MiddlePack colspan="7"> </td>
													</tr>
													<tr>
														<td class=FieldCellSub valign="top"> Previous Financial Year</td>
														<td class="FieldCellSub" valign="top">Start Date</td>
														<td class='FieldCellSub'>
															<select size="1" name="selPFinStartDate" class="FormElem" onChange="SetDates(this)">
																<option value="select">Select</option>
															<%	'Calling the Function which populates the Previous Financial Year Start Date list
																populateFinDate
															%>
															</select>
                                                        </td>
														<td class='FieldCellSub'></td>
														<td class='FieldCellSub'>End Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly" id="idPFinEndDate">&nbsp;</span>
                                                        </td>
														<td class='FieldCellSub'>
                                                        </td>
													</tr>
													<tr>
														<td class=MiddlePack valign="top" colspan="7"> 
														</td>
													</tr>
													<tr>
														<td class=FieldCellSub valign="top">Current Financial Year</td>
														<td class="FieldCellSub" valign="top">Start Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly" id="idCFinStartDate">&nbsp;</span>
                                                        </td>
														<td class='FieldCellSub'></td>
														<td class='FieldCellSub'>End Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly" id="idCFinEndDate">&nbsp;</span>
														</td>
														<td class='FieldCellSub'></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
							</tr>

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
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center">
												<input type="button" value="Proceed" name="Proceed" class="ActionButton" onclick="CheckSubmit()">
                                                <input type="reset" value="Reset" name="Reset" class="ActionButton">
                                                <input type="button" value="Cancel" name="cancel" class="ActionButton" OnClick="window.location.href='../welcome_admin.asp'">
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
	' Function which populates the Previous Financial Year Start Date list
	Function populateFinDate()
		' Declaration of variables
		Dim dcrs,sStartDate,sEndDate
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CONVERT(CHAR,FROMPERIOD,103),CONVERT(CHAR,TOPERIOD,103) FROM MS_FINANCIALPERIOD WHERE ACTIVE = 'N' ORDER BY 1 DESC"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sStartDate = dcrs(0)
		set sEndDate = dcrs(1)
		
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sStartDate)&":"&trim(sEndDate)&""">"&trim(sStartDate)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>

