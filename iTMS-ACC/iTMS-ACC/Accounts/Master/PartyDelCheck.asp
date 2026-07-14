<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PartyDelCheck.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Jun 22 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	On Clicking the Save Button it calls the SalPartyBankrel.asp
	'Procedures/Functions Used	:
	'Internal Variables			:

	'Database					:	SITMS
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
	Dim objrs,sQuery,sRet,sCallTy,Temparr,sTemp,sPartycode,sBankname,sCount
	Dim sCode,sValue
	Set Objrs = Server.CreateObject("ADODB.Recordset")

	'Response.Write sCallTy & sPartycode & sBankName
	sTemp = Request.QueryString("sCallType")
	Temparr = Split(sTemp,"?")
	sCallTy = Temparr(0)

	IF sCallTy = "P" Then
		CheckParty()
	End IF
	'Response.Write "AA"
%>

<%

	Function CheckParty()
		sTemp = Request.QueryString("sCallType")
		Temparr = Split(sTemp,"?")
		'Response.Write "Temparr = "& UBound(Temparr)
		sTemp = Temparr(4)
		sTemp = Trim(sTemp)
		'Response.Write sTemp

		sQuery = "Select Count(1) From Acc_T_VoucherHeader Where isNull(PartyCode,0) = "&sTemp
		With Objrs
		 	.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End with
		Set Objrs.Activeconnection = nothing
		IF Not objrs.EOF Then
			sCount = Objrs(0)
		End IF
		objrs.Close
		
		sQuery = "Select Count(1) From Acc_T_VoucherDetails Where isNull(AccUnitPartyCode,0) = "&sTemp
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End with
		Set Objrs.Activeconnection = nothing
		IF Not objrs.EOF Then
			sCount = Objrs(0)
		End IF
		objrs.Close
		
'Response.Write "Voucher Count = "& sCount 
		IF CStr(sCount) = "0" Then
			sQuery = "Select Count(1) From Acc_T_CreatedVoucherHeader Where PartyCode = "&sTemp
			With Objrs
		 		.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			End with
			Set Objrs.Activeconnection = nothing
			IF Not objrs.EOF Then
				sCount = Objrs(0)
			End IF
			objrs.Close
		End IF
		
		IF CStr(sCount) = "0" then
			sQuery = "Select Count(1) From Acc_T_CreatedVoucherDetails Where AccUnitPartyCode = "& sTemp
			With Objrs
		 		.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			End with
			Set Objrs.Activeconnection = nothing
			IF Not objrs.EOF Then
				sCount = Objrs(0)
			End IF
			objrs.Close
		End if
		

		IF CStr(sCount) = "0" Then
			sQuery = "Select isNull(Sum(OpeningAmount),0),isNull(Sum(ClosingAmount),0) From Acc_T_PartyOpeningAmt "&_
					 "Where PartyCode = "&sTemp

			With Objrs
		 		.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			End with
			Set Objrs.Activeconnection = nothing
			IF Not objrs.EOF Then
				IF CStr(objrs(0)) = "0" and CStr(objrs(1)) = "0" Then
					sCount = 0
				Else
					sCount = 1
				End IF
			End IF
			objrs.Close
		End IF

		IF Cstr(sCount) = "0" Then

			'================== Check For the Party Avilable in Sales =========================
			sQuery = "Select Count(1) From Sal_T_EnquiryHeader Where isNull(PartyCode,0) = "&sTemp&" "
			objrs.Open sQuery,Con
			IF not objrs.EOF Then
				sCount = objrs(0)
			End IF
			objrs.Close

			IF CStr(sCount) = "0" Then
				sQuery = "Select Count(1) From Sal_T_InvoiceHeader Where isNull(PartyCode,0) = "&sTemp&" "
				objrs.Open sQuery,Con
				IF not objrs.EOF Then
					sCount = objrs(0)
				End IF
				objrs.Close
			End IF

			IF CStr(sCount) = "0" Then
				sQuery = "Select Count(1) From Sal_T_QuotationHeader Where isNull(PartyCode,0) = "&sTemp&" "
				objrs.Open sQuery,Con
				IF not objrs.EOF Then
					sCount = objrs(0)
				End IF
				objrs.Close
			End IF

			IF CStr(sCount) = "0" Then
				sQuery = "Select Count(1) From Sal_T_OrdersHeader Where isNull(PartyCode,0) = "&sTemp&" "
				objrs.Open sQuery,Con
				IF not objrs.EOF Then
					sCount = objrs(0)
				End IF
				objrs.Close
			End IF

			'======================== End of Sales Check =======================================
		End IF


		IF cstr(sCount) = "0" Then
			Response.Write "F"
		else
			 Response.Write "T"
		End IF
	End Function

%>

