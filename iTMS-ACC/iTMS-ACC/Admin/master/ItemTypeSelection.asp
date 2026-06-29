<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItemTypeSelection.asp
	'Module Name				:	Admin (Master)
	'Author Name				:	UmaMaheswari S
	'Created On					:	January 2011, 20
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Role Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"><Root ItemTypeID="" Done=""/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
	function itemTypeSelectionRoot() {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		return window.OutData && window.OutData.documentElement;
	}

	function CheckSubmit() {
		var root = itemTypeSelectionRoot();
		if (root) {
			root.setAttribute("ItemTypeID", document.formname.selIType.value);
			root.setAttribute("Done", "Y");
			if (window.ITMSModernCompat) {
				window.ITMSModernCompat.returnModalValue(root);
			} else {
				window.returnValue = root;
				window.returnvalue = root;
			}
		}
		window.close();
	}

	window.addEventListener("beforeunload", function () {
		var root = itemTypeSelectionRoot();
		if (root && window.ITMSModernCompat) {
			window.ITMSModernCompat.returnModalValue(root);
		} else if (root) {
			window.returnValue = root;
			window.returnvalue = root;
		}
	});
</SCRIPT>
<%
	Dim sTemp,sArr,sPassType,sSql

	Dim dcrs
	Set dcrs = Server.CreateObject("ADODB.RecordSet")


%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >

	<form method="POST" name="formname" action="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Item Type
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td class="TabCell" valign="bottom" align="center" width="95">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">

										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>

					<tr>
						<td class="TabBody">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<div align="left">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td class="FieldCellSub">
														<select size="8" name="selIType" class="FormElem">
															<option value="" Selected>Not Applicable</option>
															<%	'Calling the Function which populates the Item Type list
																'populateItemType
															%>
														</select>
													</td>
												</tr>

											</table>
										</div>
									</td>
									<td align="center">
									</td>
								</tr>


								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="Button" value="Done" name="B1" class="ActionButton" onclick="CheckSubmit()">
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
</body>
</html>
<%
Function populateItemType()
	' Declaration of variables
	Dim dcrs,stypID,stypName
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.Source = "SELECT ITEMTYPEID,ITEMTYPENAME FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
		.ActiveConnection = con
		.Open
	end with
	set stypID = dcrs(0)
	set stypName = dcrs(1)
	If not dcrs.EOF then
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
	end if
	dcrs.Close
	set dcrs.ActiveConnection = nothing

End Function
	%>
