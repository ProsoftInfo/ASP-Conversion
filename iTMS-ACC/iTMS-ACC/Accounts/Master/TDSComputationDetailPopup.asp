<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/ReportsBody.css" type="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"><Root/></script>

</head>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" Onload="Loadvalues()">
<%
	Dim GroupCode,iHeadID,OutValue,splt
	GroupCode = Request("GroupCode")
	'splt = Split(OutValue,":")
	'GroupCode = CInt(splt(0))
	'HeadID = CInt(splt(1))  
	iHeadID = Request("HeadID")
	'Response.Write iHeadID
	'HeadID = 3	
	
Response.Write OutValue
%>
<script>
window.__itmsPopupCompat = { type: "tdsComputation" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>



	<form method="POST" name="formname" action onload="Loadvalues()">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="PageTitle" height="20">Computation Detail
			</td>
		</tr>
		<input type="hidden" name="GroupCode" value="<%=GroupCode%>">
		<input type="hidden" name="HeadID" value="<%=iHeadID%>">
		<input type="hidden" name="iRowCount" value="1">
		
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td class="FieldCell">TDS Group Name
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtGname" size="30" maxlength="13" class="FormElem" ReadOnly>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Computation for
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtcomputationfor" size="30" maxlength="13" class="FormElem" ReadOnly>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Percentage
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtpercentage" size="4" maxlength="3" class="FormElem">
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td>
										<div class="frmBody" id="td" style="width: 355; height:120;">
											<table border="0" cellspacing="1" class="ExcelTable" width="345" id="tbltds">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" width="10">
													</td>
													<td class="ExcelHeaderCell" align="center">TDS Head Name
													</td>
												</tr>

												

												<center>
												</table>
											</div>
										</center>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell" align="center">
													<input type="button" value="Save" name="B1" class="ActionButton" onClick="UpdateXML()">
 													<input type="button" value="Close" name="B2" class="ActionButton" onclick="window.close();">
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
