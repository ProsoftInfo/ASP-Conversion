<%@ Language="VBScript" %>
<% option explicit %>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CheckVouAdjDet.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu R
	'Created On					:	Feb 15, 2006	
%>
<!--#include virtual="/include/Databaseconnection.asp"-->
<%
	Dim Objrs,iTransNo,sBookCode,sQuery,sRetVal
	sBookCode = Request("BookCode")
	iTransNo = Request("TransNo")
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	'sRetVal = "FF"
	
	IF CStr(sBookCode) = "02" or CStr(sBookCode) = "01" or CStr(sBookCode) = "08"   Then 'IF It is an Bank 
		sQuery = "Select C.Narration From ACC_T_CreatedAdvanceAdj C,Acc_T_CreatedAdvances T "&_
				 "Where C.CreatedAdvanceNo = T.CreatedAdvanceNo and T.CreatedTransNo = "&iTransNo&" "&_
				 "And C.AmountAdjusted > 0 "
				
		'Response.Write sQuery 
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not Objrs.EOF Then
			'sRetVal = "This Voucher has been adjusted for the Following Bills \n "
			Do While Not objRs.EOF
				sRetVal = sRetVal &","& Objrs(0)
				'sRetVal = sRetVal &","& Mid(Objrs(0),12)
				objRs.MoveNext
			Loop
			'sRetVal = Mid(sRetVal,2)
		End IF
		objRs.Close	
		
		sQuery = "Select CreatedVoucherNo+' Dt:'+Convert(Varchar,VoucherDate,103)+' Amount:'+ "&_
				 "RTRIM(Cast(VoucherAmount As Char)),BookCode From  CrPayAdvAdjDet Where CreatedTransNo = "&iTransNo
		
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not Objrs.EOF Then
			'sRetVal = "This Voucher has been adjusted for the Following Bills \n "
			Do While Not objRs.EOF
				IF CStr(Objrs(1)) = "02" Then
					sRetVal = sRetVal &","&"Bank Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "01" Then
					sRetVal = sRetVal &","&"Cash Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "04" Then
					sRetVal = sRetVal &","&"Purchase Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "05" Then
					sRetVal = sRetVal &","&"Sales Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "07" Then
					sRetVal = sRetVal &","&"Credit Note Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "06" Then
					sRetVal = sRetVal &","&"Debit Note Voucher No: " & Objrs(0)
				End IF
				objRs.MoveNext
			Loop
			'sRetVal = Mid(sRetVal,2)
		End IF
		objRs.Close	
		
		sQuery = "Select Distinct C.CreatedVoucherNo+' Dt:'+ Convert(Varchar(10),C.VoucherDate,103) + ' Amount:'+RTRim(Cast(C.VoucherAmount AS Char)) "&_
				 ",C.BookCode From CrRecAdjDetails C, Acc_T_CreatedAdvances P Where  "&_
				 "C.ReceivableNumber = P.CreatedAdvanceNo and P.CreatedTransno = "&iTransNo
		
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not Objrs.EOF Then
			'sRetVal = "This Voucher has been adjusted for the Following Bills \n "
			Do While Not objRs.EOF
				IF CStr(Objrs(1)) = "02" Then
					sRetVal = sRetVal &","&"Bank Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "01" Then
					sRetVal = sRetVal &","&"Cash Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "04" Then
					sRetVal = sRetVal &","&"Purchase Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "05" Then
					sRetVal = sRetVal &","&"Sales Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "07" Then
					sRetVal = sRetVal &","&"Credit Note Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "06" Then
					sRetVal = sRetVal &","&"Debit Note Voucher No: " & Objrs(0)
				End IF
				objRs.MoveNext
			Loop
			
		End IF
		objRs.Close	
		
		sQuery = "Select Distinct C.CreatedVoucherNo+' Dt:'+ Convert(Varchar(10),C.VoucherDate,103) + ' Amount:'+RTRim(Cast(C.VoucherAmount AS Char)) "&_
				 ",C.BookCode From CrPayAdjDetails C, Acc_T_CreatedAdvances P Where  "&_
				 "C.PayablesNumber = P.CreatedAdvanceNo and P.CreatedTransno = "&iTransNo
		
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not Objrs.EOF Then
			'sRetVal = "This Voucher has been adjusted for the Following Bills \n "
			Do While Not objRs.EOF
				IF CStr(Objrs(1)) = "02" Then
					sRetVal = sRetVal &","&"Bank Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "01" Then
					sRetVal = sRetVal &","&"Cash Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "04" Then
					sRetVal = sRetVal &","&"Purchase Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "05" Then
					sRetVal = sRetVal &","&"Sales Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "07" Then
					sRetVal = sRetVal &","&"Credit Note Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "06" Then
					sRetVal = sRetVal &","&"Debit Note Voucher No: " & Objrs(0)
				End IF
				objRs.MoveNext
			Loop
			'sRetVal = Mid(sRetVal,2)
		End IF
		objRs.Close	
		sRetVal = Mid(sRetVal,2)
		
	ElseIF CStr(sBookCode) = "04" Then
		sQuery = "Select H.CreatedVoucherNo+' Dt:'+Convert(Varchar,H.VoucherDate,103) +' Amount:'+ "&_
				 "RTRIM(Cast(H.VoucherAmount As Char)),H.BookCode From  "&_
				 "Acc_T_CreatedVoucherHeader H,CrPayAdjDetails C "&_
				 "Where H.CreatedTransNo = C.AdjustedTransNo And C.CreatedTransNo = "&iTransNo&" "&_
				 "And C.AdjustType Is Null  And H.BookCode <> '04' "
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not Objrs.EOF Then
			'sRetVal = "This Voucher has been adjusted for the Following Bills \n "
			Do While Not objRs.EOF
				IF CStr(Objrs(1)) = "02" Then
					sRetVal = sRetVal &","&"Bank Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "01" Then
					sRetVal = sRetVal &","&"Cash Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "04" Then
					sRetVal = sRetVal &","&"Purchase Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "05" Then
					sRetVal = sRetVal &","&"Sales Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "07" Then
					sRetVal = sRetVal &","&"Credit Note Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "06" Then
					sRetVal = sRetVal &","&"Debit Note Voucher No: " & Objrs(0)
				End IF
				objRs.MoveNext
			Loop
			'sRetVal = Mid(sRetVal,2)
		End IF
		objRs.Close	
		sRetVal = Mid(sRetVal,2)
		
	ElseIF CStr(sBookCode) = "05" Then
		sQuery = "Select CreatedVoucherNo+' Dt:'+Convert(Varchar(10),VoucherDate,103)+' Amount:'+RTRIM(Cast(VoucherAmount As Char)) "&_
				 ",BookCode From CrRecAdjDetails Where CreatedTransno = "&iTransNo&" and AdjustType is NULL "
		
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not Objrs.EOF Then
			'sRetVal = "This Voucher has been adjusted for the Following Bills \n "
			Do While Not objRs.EOF
				IF CStr(Objrs(1)) = "02" Then
					sRetVal = sRetVal &","&"Bank Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "01" Then
					sRetVal = sRetVal &","&"Cash Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "04" Then
					sRetVal = sRetVal &","&"Purchase Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "05" Then
					sRetVal = sRetVal &","&"Sales Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "07" Then
					sRetVal = sRetVal &","&"Credit Note Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "06" Then
					sRetVal = sRetVal &","&"Debit Note Voucher No: " & Objrs(0)
				End IF
				objRs.MoveNext
			Loop
			'sRetVal = Mid(sRetVal,2)
		End IF
		objRs.Close	
		sRetVal = Mid(sRetVal,2)
	Elseif CStr(sBookCode) = "06" Then
		sQuery = "Select CreatedVoucherNo+' Dt:'+Convert(Varchar(10),VoucherDate,103)+' Amount:'+RTRIM(Cast(AmountReceivable As Char)) "&_
				 ",BookCode From CrDNBillAdjDet Where CreatedTransno = "&iTransNo
		
		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs.ActiveConnection = Nothing
		IF Not Objrs.EOF Then
			'sRetVal = "This Voucher has been adjusted for the Following Bills \n "
			Do While Not objRs.EOF
				IF CStr(Objrs(1)) = "02" Then
					sRetVal = sRetVal &","&"Bank Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "01" Then
					sRetVal = sRetVal &","&"Cash Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "04" Then
					sRetVal = sRetVal &","&"Purchase Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "05" Then
					sRetVal = sRetVal &","&"Sales Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "07" Then
					sRetVal = sRetVal &","&"Credit Note Voucher No: " & Objrs(0)
				Elseif CStr(Objrs(1)) = "06" Then
					sRetVal = sRetVal &","&"Debit Note Voucher No: " & Objrs(0)
				End IF
				objRs.MoveNext
			Loop
			'sRetVal = Mid(sRetVal,2)
		End IF
		objRs.Close	
		sRetVal = Mid(sRetVal,2)
	End IF
	
	Response.Write sRetVal
	
	
	
%>