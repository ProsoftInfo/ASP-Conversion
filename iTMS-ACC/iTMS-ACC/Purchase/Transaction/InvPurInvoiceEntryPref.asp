<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/PurchaseTermsConditions.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/PurChkItemSpecPack.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/invPurInvoiceEntryPref.js"></SCRIPT>

<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Dim InvData
Set InvData = window.dialogArguments

'-------------------------------------------------------------------------------------------

function Done_Clk()
dim TempNode
	if InvData.HasChildNodes then
		for each TempNode in InvData.ChildNodes
					TempNode.setAttribute "DespatchMode",document.formname.cmbDespatch.value
					TempNode.setAttribute "PaymentMode",document.formname.cmbPayment.value
					TempNode.setAttribute "PayTerms",document.formname.cmbPayTerms.value
					TempNode.setAttribute "IssueBank",document.formname.cmbIssueBank.value
					TempNode.setAttribute "BenificiaryBank",document.formname.cmbBenificiaryBank.value
					TempNode.setAttribute "PricingBasis",document.formname.cmbPricing.value
					TempNode.setAttribute "Transporter",document.formname.cmbTransporter.value
					TempNode.setAttribute "LoadingPort",document.formname.cmbLoadPort.value
					TempNode.setAttribute "DestPort",document.formname.cmbDestPort.value
		next
	end if
	'msgbox InvData.xml
	window.close()
end function

'-------------------------------------------------------------------------------------------
Function window_onunload()
	Set window.returnvalue= InvData
End Function
'-------------------------------------------------------------------------------------------

</script>
<%
'Declaring Variables
Dim Curr,Mod1,Mop,IssueBank,PayTerm,Bop,Transporter
Dim nBenefitBank,nLoadPort,nDestPort

Mod1=Request.QueryString("Mod1")
Mop=Request.QueryString("Mop")
IssueBank=Request.QueryString("IssueBank")
PayTerm=Request.QueryString("PayTerm")
Bop=Request.QueryString("Bop")
Transporter=Request.QueryString("Transporter")
nBenefitBank = Request.QueryString("BenefitBank")
nLoadPort = Request.QueryString("LoadPort")
nDestPort	= Request.QueryString("DestPort")


if trim(Mod1) = "" then Mod1=0
if trim(Mop) = "" then Mop=0
if trim(IssueBank) = "" then IssueBank=0
if trim(PayTerm) = "" then PayTerm=0
if trim(Bop) = "" then Bop=0
if trim(Transporter) = "" then Transporter=0
if trim(nBenefitBank) = "" then nBenefitBank = 0
if trim(nLoadPort) = "" then nLoadPort = 0
if trim(nDestPort) = "" then nDestPort	= 0

'Response.Write "<p><p> " & Request.QueryString

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<form method="POST" name="formname" action>
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Purchase Invoice Entry - Preferences
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
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td class="FieldCell">Payment Terms
												</td>
												<td class="FieldCellSub">
												<select size="1" name="cmbPayTerms" class="FormElem">
													<option value = "0" selected>Select</option>
								  					<%
								  					''To populate Payment Terms
								  					popPaymentTerms(PayTerm) %>

								              </select>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Mode of Despatch
												</td>
												<td class="FieldCellSub">
												<select size="1" name="cmbDespatch" class="FormElem">
												    <option value = "0" selected>Select</option>
								  					<%
								  					popModeDespatch(Mod1)
								  					%>
												 </select>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Basis of Pricing
												</td>
												<td class="FieldCellSub">
												<select size="1" name="cmbPricing" class="FormElem">
								                <option  value = "0" selected>Select</option>
								                <%
								                popPricingBasis(Bop)
								                %>
								              </select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Mode of Payment
												</td>
												<td class="FieldCellSub"><select size="1" name="cmbPayment" class="FormElem">
								                <option value = "0" selected>Select</option>
								                <%
								                popModePayment(Mop)
								                %>
								              </select>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Transporter
												</td>
												<td class="FieldCellSub"><select size="1" name="cmbTransporter" class="FormElem">
								            <option value = "0" selected >Select</option>
								  			<%
								  			popTransporter(Transporter)
								  			%>
								          </select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Issuing Bank
												</td>
												<td class="FieldCellSub"><select size="1" name="cmbIssueBank" class="FormElem">
											<option value = "0" selected> Select</option>
											<%
											popIssueBank(IssueBank)
											%>
								            </select>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Loading Port
												</td>
												<td class="FieldCellSub"><select size="1" name="cmbLoadPort" class="FormElem">
								                <option value = "0" selected>Select</option>
								  			<%
								  			popLoadPlaces(nLoadPort)
								  			%>
								              </select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Beneficiary Bank
												</td>
												<td class="FieldCellSub">
												<select size="1" name="cmbBenificiaryBank" class="FormElem">
								  				<option value = "0" selected >Select</option>
								  				<%
								  				popBenificiaryBank(nBenefitBank)
								  				%>
												</select>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Destination Port
												</td>
												<td class="FieldCellSub"><select size="1" name="cmbDestPort" class="FormElem">
								                <option value = "0" selected>Select</option>
								  			<%
								  			popDestination(nDestPort)
								  			%>
								              </select>
												</td>
											</tr>

										</table>
									</td>
									<td>
									</td>
								</tr>

								<tr>
									<td colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td>
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<input type="button" value="Done" name="B4" class="ActionButton" onclick="Done_Clk()">
													<input type="button" value="Close" name="B4" class="ActionButton" onclick="window.close()">
												</td>
											</tr>

										</table>
									</td>
									<td>
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td colspan="3" class="BottomPack">
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
