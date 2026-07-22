<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	popTaxDetails.asp
	'Module Name				:	Purchase (Transactions-Invoice)
	'Author Name				:
	'Created On					:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	invPurInvoiceHeaderEntry.asp
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
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/PurchaseTermsConditions.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS-Invoice Form Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/calcAlternateUoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/popTaxDetails.js"></SCRIPT>

<script type="application/xml" id="TaxFormData" data-itms-xml-island="1"><Root/></script>

</HEAD>
<%
Dim sPurTypeName

Dim nPtype

Dim rsTemp

nPtype =  Request.QueryString("PurType")

set rsTemp = Server.CreateObject("ADODB.RecordSet")

with rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT PURTYPESHORTNAME,PURCHASETYPENAME FROM APP_M_PURCHASETYPES  where upper(isNull(Active,'Y')) = 'Y'  ORDER BY PURCHASETYPE"
	.ActiveConnection = con
	.Open
end with
set rsTemp.ActiveConnection = nothing

sPurTypeName = ""
if not rsTemp.EOF then
	sPurTypeName = rsTemp(1)
end if
rsTemp.Close
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload=init()>

<form method="POST" name="formname" action="">
<input type=hidden name="hPurType" value="<%=nPtype%>">
<input type=hidden name="hPurTypeName" value="<%=sPurTypeName%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
          Purchase Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
<TABLE id="Table16" cellSpacing=0 cellPadding=0 border=0 width="100%"  >
<TR>
<TD class="TabBodyWithTopLine">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
					<td align="center" colspan="3" class="MiddlePack" height="7" width="600">
					</td>
                </tr>
                <tr>
					<td align="center" width="5" class="ClearPixel" height="2">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
					<td align="center" class="ClearPixel" width="6" height="2">
                        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
                </tr>
                <tr>
					<td align="center" colspan="3" class="MiddlePack" height="7" width="600">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
                </tr>
                <tr>
					<td align="center" width="5" class="ClearPixel" height="2">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
					<td valign="top">
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <!---------------------------------------------->
                <tr>
					<td  class="FieldCell">Item Name</td>
					<td  class="FieldCellSub">
					<span id="spnItemName" class="Dataonly"></span></td>
				 </tr>
				<tr>
					<td  class="FieldCell">Purchase Type</td>
					<td  class="FieldCellSub">
					<span id="spnPurType" class="Dataonly"></span></td>
				 </tr>
				 <tr>
					<td  class="FieldCell">Basic Value</td>
					<td  class="FieldCellSub">
					<span id="spnBasicValue" class="Dataonly"></span></td>
				 </tr>
				 <tr>
					<td  class="FieldCell">Nett Basic Value</td>
					<td  class="FieldCellSub">
					<span id="spnAmount" class="Dataonly"></span></td>
				 </tr>
				<!------------------------------------------------>
                  <tr>
                    <td valign="top" colspan="2">
                      <div class="frmbody" id="frm1" style="width:450;">
                        <table border="0" cellspacing="1" class="ExcelTable" id="tblTaxDet" width="100%">
                          <tr>
                            <td class="ExcelHeaderCell" align="center" width=5 >
                              <p align="center">S.No</td>
                            <td class="ExcelHeaderCell"  align="center" >Tax<br>Name</td>
                            <td class="ExcelHeaderCell" align="center" >Tax<br>Percentage
                            </td>
                            <td class="ExcelHeaderCell"  align="center" >Tax<br>Value
                            </td>
						</tr>

				       </table>
                      </div>
                    </td>
                  </tr>
                </table>
								</td>
								<td align="center" class="ClearPixel" width="6" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<input type="button" value="Done" name="B8" onClick="window.close()"  class="ActionButton" tabindex="3" >
															<input type="button" value="Cancel" name="B10"  onClick="window.close()" class="ActionButton" tabindex="3" >
															<input type="reset" value="Reset" name="B9" class="ActionButton" tabindex="3" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="6">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
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

