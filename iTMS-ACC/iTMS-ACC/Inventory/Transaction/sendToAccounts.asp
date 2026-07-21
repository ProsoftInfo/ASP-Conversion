<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	sendToAccounts.asp
	'Module Name				:	Inventory (Send Closing Stock to Accounts)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	September 24,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	sendToAccountsDetails.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include File="../../include/IncludeDatePicker.asp" -->
<%
Dim sFinPeriod,sArrFin,sFromDate,sToDate
sFinPeriod = Session("FinPeriod")
sArrFin = Split(sFinPeriod,":")
sFromDate = "01/04/"&sArrFin(0)
sToDate = "31/03/"&sArrFin(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Closing Stock to Accounts</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	Function SetDate()
	    Dim sFromDate,sToDate
	    sFromDate = document.formname.hFromDate.value
	    sToDate = document.formname.hToDate.value
	    if DateDiff("d",sToDate,date)>0 then
	        document.formname.ctlClosingDate.setMinDate = sFromDate
		    document.formname.ctlClosingDate.SetMaxDate=sToDate
		    document.formname.ctlClosingDate.setDate = sToDate
		else
		    document.formname.ctlClosingDate.setMinDate = sFromDate
		    document.formname.ctlClosingDate.SetMaxDate=Date
		    document.formname.ctlClosingDate.setDate = date
		end if

	end Function

	Function CheckSubmit(todaysdate)
		if document.formname.selUnit.selectedIndex = "0" then
			alert("Select Unit")
			document.formname.selUnit.focus
			exit function
		elseif DateDiff("d",todaysdate,document.formname.ctlClosingDate.GetDate) > 0  then
			alert("Closing as on should be less than or equal to Today's date")
			exit function
		elseif document.formname.selFor.selectedIndex = "0" then
			alert("Select Closing Stock For")
			document.formname.selFor.focus
			exit function
		else
			document.formname.hClosingDate.value = document.formname.ctlClosingDate.GetDate
			document.formname.hUnitName.value = document.formname.selUnit(document.formname.selUnit.selectedIndex).text
			document.formname.hForName.value = document.formname.selFor(document.formname.selFor.selectedIndex).text
			document.formname.action = "sendToAccountsDetails.asp"
			document.formname.submit()
		end if
	end Function

</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate()">
<form method="POST" name="formname">
<input type=hidden name="hClosingDate" value="">
<input type=hidden name="hUnitName" value="">
<input type=hidden name="hForName" value="">
<input type="hidden" name="hFromDate" value="<%=sFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sToDate%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
			Closing Stock to Accounts
		</td>
	</tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
				<tr>
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onmouseover="tabrollover(this)" onmouseout="tabrollout(this)">
										<tr>
											<td width="100%" align="center">Details</td>
										</tr>
									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center" width="5"></td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Select Unit</td>
											<td class="FieldCellSub">
												<select size="1" name="selUnit" class="FormElem">
													<option value="select">Select</option>
													<%	'Calling the Function which populates Organization Unit list
														populateUnit
													%>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Closing as on</td>
											<td class="FieldCellSub">
												<%
													' Function Call to Insert Date Picker
													Response.Write InsertDatePicker("ctlClosingDate")
												%>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Closing Stock For</td>
											<td class="FieldCellSub">
												<select size="1" name="selFor" class="FormElem">
													<option value="select">Select</option>
													<option value="SOH">Stock On Hand</option>
												</select>
											</td>
										</tr>
									</table>
								</td>
								<td align="center" width="5"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<input type="button" value="Next" name="B2" class="ActionButton" onclick="CheckSubmit('<%=FormatDate(date())%>')">
												<input type="reset" value="Reset" name="B1" class="ActionButton">
 												<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="10" colspan="3" class="BottomPack"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</body>
</html>