<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParPerferencePopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 15,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/sessionVerify.asp"-->
<%
dim objRs,sQuery,oDOM,iPartyCode,Root,nodHeader
Dim sTrans,sMod,sCurr,sMop,sBop,sPayTerm
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iPartyCode = Request.QueryString("PartyCode")


	IF CStr(iPartyCode) = "" Then
		iPartyCode = 0
	End IF

	sQuery = "SELECT isNull(PrefTransporterCode,0), isNull(PrefDespatchMode,0), isNull(PrefCurrencyCode,0), isNull(PrefPaymentMode,0), "&_
			 "isNull(PrefBasisOfPricing,0), isNull(PrefPaymentTerms,0) FROM APP_R_OrgParty "&_
		 "WHERE PartyCode = "&iPartyCode&" "

	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With

	Set objRs.ActiveConnection = Nothing
	IF Not objRs.EOF Then
		sTrans = objRs(0)
		sMod = objRs(1)
		sCurr = objRs(2)
		sMop = objRs(3)
		sBop = objRs(4)
		sPayTerm = objRs(5)

	End IF
	objRs.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self">
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "partyPreference" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname">
<input type="hidden" name="hPartyCode" value="<%=iPartyCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Party Preference</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
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
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCell width="100"> Payment Terms</td>
															<td class='FieldCell'><select size="1" name="selPayTerms" class="FormElem">
															<option value="0">Select Payment Terms</option>
<%
	sQuery="SELECT PaymentTermsNo, PaymentTermsDesc, PymtTermsShortDesc FROM APP_M_PaymentTermsHeader"
	with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
	end with
	set objRs.ActiveConnection=nothing

	do while not objRs.EOF
		IF CStr(sPayTerm) = CStr(objRs(0)) Then
			Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(1)&"</option>"
		Else
			Response.Write "<option value="""&objRs(0)&""">"&objRs(1)&"</option>"
		End IF
		objRs.MoveNext
	loop
	objRs.Close
%>
                                                            </select></td>
														</tr>
														<tr>
															<td class=FieldCell width="100"> Basis of Pricing</td>
															<td class='FieldCell'><select size="1" name="selBop" class="FormElem">
                                                        <option value="0">Select  Basis of Pricing</option>
<%
	sQuery="SELECT BasisOfPricingNo, BasisOfPricing, ShortBasisofPricing FROM APP_M_BasisOfPricing"
	with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
	end with
	set objRs.ActiveConnection=nothing

	do while not objRs.EOF
		IF CStr(sBop) = CStr(objRs(0)) Then
			Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(1)&"</option>"
		Else
			Response.Write "<option value="""&objRs(0)&""">"&objRs(1)&"</option>"
		End IF
		objRs.MoveNext
	loop
	objRs.Close
%>
                                                            </select></td>
														</tr>
														<tr>
															<td class=FieldCell width="100"> Despatch Mode</td>
															<td class='FieldCell'><select size="1" name="selDespatch" class="FormElem">
                                                        <option value="0">Select Despatch Mode</option>
<%
	sQuery="SELECT DespatchModeNo, DespatchModeDesc, ShortDespatchMode FROM APP_M_ModeOfDespatch"
	with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
	end with
	set objRs.ActiveConnection=nothing

	do while not objRs.EOF
		IF CStr(sMod) = CStr(objRs(0)) Then
			Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(1)&"</option>"
		Else
			Response.Write "<option value="""&objRs(0)&""">"&objRs(1)&"</option>"
		End IF
		objRs.MoveNext
	loop
	objRs.Close
%>
                                                            </select></td>
														</tr>
														<tr>
															<td class=FieldCell width="100"> Payment Mode</td>
															<td class='FieldCell'><select size="1" name="selPayMode" class="FormElem">
                                                        <option value="0">Select Payment Mode</option>
<%
	sQuery="SELECT PaymentModeNo, PaymentMode, ShortPaymentMode FROM APP_M_ModeOfPayment"
	with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
	end with
	set objRs.ActiveConnection=nothing

	do while not objRs.EOF
		IF CStr(sMop) = CStr(objRs(0)) Then
			Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(1)&"</option>"
		Else
			Response.Write "<option value="""&objRs(0)&""">"&objRs(1)&"</option>"
		End IF
		objRs.MoveNext
	loop
	objRs.Close
%>
                                                            </select></td>
														</tr>
														<tr>
															<td class=FieldCell width="100"> Transporter</td>
															<td class='FieldCell'><select size="1" name="selTransport" class="FormElem">
                                                        <option value="0">Select Transporter</option>
<%
	sQuery="SELECT TransporterCode, TransporterName, TransportShortName FROM APP_M_Transporter"

	with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
	end with
	set objRs.ActiveConnection=nothing

	do while not objRs.EOF
		IF CStr(sTrans) = CStr(objRs(0)) Then
			Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(1)&"</option>"
		Else
			Response.Write "<option value="""&objRs(0)&""">"&objRs(1)&"</option>"
		End IF
		objRs.MoveNext
	loop
	objRs.Close
%>
                                                            </select></td>
														</tr>
														<tr>
															<td class=FieldCell width="100"> Currency</td>
															<td class='FieldCell'><select size="1" name="selCurrency" class="FormElem">
                                                        <option value="0">Select Currency</option>
<%
	sQuery="SELECT CurrencyCode, CurrencyName, CurrencyShortName FROM Ms_CurrencyMaster"
	with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
	end with
	set objRs.ActiveConnection=nothing

	do while not objRs.EOF
		IF CStr(sCurr) = CStr(objRs(0)) Then
			Response.Write "<option value="""&objRs(0)&""" Selected>"&objRs(1)&"</option>"
		Else
			Response.Write "<option value="""&objRs(0)&""">"&objRs(1)&"</option>"
		End IF
		objRs.MoveNext
	loop
	objRs.Close
%>
                                                            </select></td>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Save" name="B2" class="ActionButton" onClick="PageSubmit()">
                                                                <input type="button" value="Close" name="B3" onClick="window.close()"  class="ActionButton" >
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton" >

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
