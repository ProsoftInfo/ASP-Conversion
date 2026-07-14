
<%
'FUNCTION TO POPULATE PAYMENT TERMS
Function popPaymentTerms(iPayTermNoSel)
	Dim rsPop,sPayTermsShDesc,iPayTermNo

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT PaymentTermsNo, PymtTermsShortDesc,PaymentTermsDesc FROM APP_M_PaymentTermsHeader order by PaymentTermsNo"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iPayTermNo = rsPop(0)
	set sPayTermsShDesc = rsPop(1)
	Do While Not rsPop.EOF
		If cstr(iPayTermNoSel) = cstr(iPayTermNo) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iPayTermNo))&""" Selected>"&trim(sPayTermsShDesc)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iPayTermNo))&""">"&trim(sPayTermsShDesc)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE MODE OF PAYMENT
Function popModePayment(iPayModeNoSel)
	Dim rsPop,sShPayMode,iPayModeNo

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT PaymentModeNo, ShortPaymentMode,PaymentMode FROM APP_M_ModeOfPayment order by PaymentModeNo"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iPayModeNo = rsPop(0)
	set sShPayMode = rsPop(1)
	Do While Not rsPop.EOF
		If cstr(iPayModeNoSel) = cstr(iPayModeNo) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iPayModeNo))&""" Selected>"&trim(sShPayMode)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iPayModeNo))&""">"&trim(sShPayMode)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE MODE OF DESPATCH
Function popModeDespatch(iDesModeNoSel)
	Dim rsPop,sShDesMode,iDesModeNo

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DespatchModeNo, ShortDespatchMode,DespatchModeDesc FROM APP_M_ModeOfDespatch order by DespatchModeNo"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iDesModeNo = rsPop(0)
	set sShDesMode = rsPop(1)
	Do While Not rsPop.EOF
		If cstr(iDesModeNoSel) = cstr(iDesModeNo) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iDesModeNo))&""" Selected>"&trim(sShDesMode)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iDesModeNo))&""">"&trim(sShDesMode)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE DESTINATION PLACES
Function popDestination(iDestCodeSel)
	Dim rsPop,sDestShName,iDestCode

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DestinationPlaceCode,DestPlaceShortName,DestinationPlaceName FROM PUR_M_DestinationPlaces order by DestinationPlaceCode"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iDestCode = rsPop(0)
	set sDestShName = rsPop(2)
	Do While Not rsPop.EOF
		If cint(iDestCodeSel) = cint(iDestCode) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iDestCode))&""" Selected>"&trim(sDestShName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iDestCode))&""">"&trim(sDestShName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE TRANSPORTER
Function popTransporter(iTransCodeSel)
	Dim rsPop,sTransShName,iTransCode

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT TransporterCode,TransportShortName,TransporterName FROM APP_M_Transporter order by TransporterCode"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iTransCode = rsPop(0)
	set sTransShName = rsPop(1)
	Do While Not rsPop.EOF
		If cstr(iTransCodeSel) = cstr(iTransCode) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iTransCode))&""" Selected>"&trim(sTransShName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iTransCode))&""">"&trim(sTransShName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE BASIS OF PRICING
Function popPricingBasis(iPricingNoSel)
	Dim rsPop,sPricingShName,iPricingNo

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT BasisOfPricingNo,ShortBasisofPricing,BasisOfPricing FROM PUR_M_BasisOfPricing order by BasisOfPricingNo"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iPricingNo = rsPop(0)
	set sPricingShName = rsPop(1)
	Do While Not rsPop.EOF
		If cstr(trim(iPricingNoSel)) = cstr(trim(iPricingNo)) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iPricingNo))&""" Selected>"&trim(sPricingShName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iPricingNo))&""">"&trim(sPricingShName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE LOADING PLACES
Function popLoadPlaces(iLdPlaceCodeSel)
	Dim rsPop,sLdPlaceShName,iLdPlaceCode

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT LoadingPlaceCode,LoadPlaceShortName,LoadingPlaceName FROM PUR_M_LoadingPlaces order by LoadingPlaceCode"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iLdPlaceCode = rsPop(0)
	set sLdPlaceShName = rsPop(2)
	Do While Not rsPop.EOF

		If cint(iLdPlaceCodeSel) = cint(iLdPlaceCode) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iLdPlaceCode))&""" Selected>"&trim(sLdPlaceShName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iLdPlaceCode))&""">"&trim(sLdPlaceShName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE BENEFICIARY BANK
Function popBenificiaryBank(iBankCodeSel)
	Dim rsPop,sBankName,iBankCode

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT BankNumber,BankName FROM PUR_M_SupplierBankers order by BankNumber"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iBankCode = rsPop(0)
	set sBankName = rsPop(1)
	Do While Not rsPop.EOF
		If cint(iBankCodeSel) = cint(iBankCode) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iBankCode))&""" Selected>"&trim(sBankName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iBankCode))&""">"&trim(sBankName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE ISSUE BANK
Function popIssueBank(iBankCodeSel)
	Dim rsPop,sBankName,iBankCode

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT BookNumber,BookName FROM Acc_R_ApplicableAccountHeads where bookcode='02' order by BookNumber"

		.Source = "SELECT Distinct A.BookNumber,B.BankName FROM Acc_R_ApplicableAccountHeads A, " &_
					" Acc_M_BankDetails B where A.bookcode='02' and A.bookcode=B.BookCode " &_
					" and A.BookNumber=B.BookNumber	order by A.BookNumber "
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iBankCode = rsPop(0)
	set sBankName = rsPop(1)
	Do While Not rsPop.EOF
		If cint(iBankCodeSel) = cint(iBankCode) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iBankCode))&""" Selected>"&trim(sBankName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iBankCode))&""">"&trim(sBankName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>

<%
'FUNCTION TO POPULATE BENEFICIARY BANK For the Party

' Added By Subbiah.S on 29.07.03 for the Purpose of Selecting the Party bank based on the Partycode

Function popPartyBank(iBankCodeSel,iPartycode)
	Dim rsPop,sBankName,iBankCode

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT BankNumber,BankName FROM PUR_M_SupplierBankers where Partycode = "&iPartycode&" order by BankNumber "
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iBankCode = rsPop(0)
	set sBankName = rsPop(1)
	Do While Not rsPop.EOF
		If cint(iBankCodeSel) = cint(iBankCode) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iBankCode))&""" Selected>"&trim(sBankName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iBankCode))&""">"&trim(sBankName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>
<%
'FUNCTION TO POPULATE CURRENCY
Function popCurr(iCurr)
	Dim rsPop,sLdPlaceShName,iLdPlaceCode

	Set rsPop = Server.CreateObject("ADODB.RecordSet")
	with rsPop
		.CursorLocation = 3
		.CursorType = 3
		.Source = "Select CurrencyCode,CurrencyShortName from Ms_CurrencyMaster"
		.ActiveConnection = con
		.Open
	end with
	set rsPop.ActiveConnection = nothing

	set iLdPlaceCode = rsPop(0)
	set sLdPlaceShName = rsPop(1)
	Do While Not rsPop.EOF

		If cint(iCurr) = cint(iLdPlaceCode) Then
			Response.Write("<OPTION VALUE="""&trim(cstr(iLdPlaceCode))&""" Selected>"&trim(sLdPlaceShName)&"</OPTION>" &vbcrlf)
		Else
			Response.Write("<OPTION VALUE="""&trim(cstr(iLdPlaceCode))&""">"&trim(sLdPlaceShName)&"</OPTION>" &vbcrlf)
		End If
		rsPop.MoveNext
	Loop
	rsPop.Close
End Function
%>