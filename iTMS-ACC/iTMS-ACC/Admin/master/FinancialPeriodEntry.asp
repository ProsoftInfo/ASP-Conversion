<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	FinancialPeriodEntry.asp
	'Module Name				:	Financial Period
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 11, 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	FinancialPeriodInsert.asp
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
<HTML><HEAD><TITLE>Financial Period</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT>
function checkNumbers(value) {
	return /^[0-9]+$/.test(String(value || ""));
}

function SetDates(obj) {
	var yearField = document.formname.txtYear;
	var parts;
	yearField.value = "";
	yearField.readOnly = true;
	if (obj.selectedIndex === 0) {
		yearField.readOnly = false;
		return;
	}
	parts = String(obj.value || "").split(":")[0].split("/");
	yearField.value = parts[2] || "";
}

function CheckSubmit() {
	var yearField = document.formname.txtYear;
	var value = String(yearField.value || "").replace(/^\s+|\s+$/g, "");
	if (value === "") {
		alert("Enter Year");
		yearField.focus();
		return false;
	}
	if (value.length !== 4) {
		alert("Invalid Year");
		yearField.select();
		return false;
	}
	if (!checkNumbers(value)) {
		alert("Enter Numerals Only");
		yearField.select();
		return false;
	}
	document.formname.action = "FinancialPeriodInsert.asp";
	document.formname.submit();
	return true;
}
</SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgName" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td>&nbsp;</td>
		<td align="center" class=PageTitle height="20">
			<p align="center">Financial Period
		</td>
    </tr>
	<tr>
		<td>&nbsp;</td>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">

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
											<td class="FieldCell" valign="top">Financial Year</td>
											<td class="FieldCellSub" valign="top">
												<select size="1" name="selPFinStartDate" class="FormElem" onChange="SetDates(this)">
													<option value="new">< New ></option>
												<%	'Calling the Function which populates the Previous Financial Year Start Date list
													populateFinDate
												%>
												</select>
											</td>
                                        </tr>
                                        <tr>
											<td class="FieldCell" valign="top"></td>
											<td class="FieldCellSub" valign="top">
												<input type="text" name="txtYear" size="5" maxlength=4 class="FormElem">
											</td>
                                        </tr>
                                        <tr>
											<td class="FieldCell" valign="top" colspan="2"><b>*Note Enter the Starting Finiancial Year. E.G if Finaicial is 2008 To 2009 Enter 2008</b></td>

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
			.Source = "SELECT CONVERT(CHAR,FROMPERIOD,103),CONVERT(CHAR,TOPERIOD,103) FROM MS_FINANCIALPERIOD	ORDER BY 1 DESC"
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

