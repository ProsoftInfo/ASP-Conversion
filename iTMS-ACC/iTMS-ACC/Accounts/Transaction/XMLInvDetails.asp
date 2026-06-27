<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLInvDetails.asp	
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 15,2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	VouCNBookSelection.asp
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

<!-- #include File="../../include/DatabaseConnection.asp" -->

<%
	dim OutData,Root,newElem,sQuery,objRs,Objrs2,dCrValue,sQuery2
	dim sPartyCode,sTemp,sOrgId,sBookCode,sCallTy,sFromApp

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs2 = Server.CreateObject("ADODB.Recordset")
	
	sPartyCode = Request("PartyCode")
	sOrgId = Request("OrgId")
	sBookCode=Request("BookCode")
	sCallTy = Request("sCallTy")
	sTemp=split(sPartyCode,"?")
	'IF CStr(sBookCode) = "05" Then
	'	sFromApp = "3"
	'Else
	'	sFromApp = "2"
	'End IF
	
	sTemp = Split(sPartyCode,"?")
	
	
	
	
	sFromApp = 3
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	
	IF CStr(Trim(sCallTy)) <> "O" Then
	
		'sQuery = "select CreatedTransNo,VoucherNumber,convert(char,VoucherDate,103),VoucherAmount,PayToRecdFrom from Acc_T_VoucherHeader where "&_
		'		 "OUDefinitionID='"&sOrgId&"' and  BookCode='"&sBookCode&"' and "&_
		'		 " PartyType='"&trim(sTemp(0))&"' and PartySubType= "&trim(sTemp(1))&" and  PartyCode="&trim(sTemp(3))  
		'==============================================================================================
		'Added On			:	22/11/2004
		'Reason				:	This is to filter only the invoices that the credit note is not 
		'						full raised to the invoice quantity.
		'==============================================================================================
		
		IF CStr(sBookCode) = "05" Then 'Purchase Invoice
			sQuery = "select H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
					 "H.PayToRecdFrom,P.ReceivableNumber from Acc_T_VoucherHeader H, Acc_T_CreatedReceivables P "&_
					 "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='"&sBookCode&"' and H.PartyType='"&trim(sTemp(0))&"' "&_
					 "and H.CreatedTransNo = P.CreatedTransNo "&_
					 "and H.CreatedTransNo Not In (Select Distinct isNull(AccTransactionNo,0) From Sal_T_AdditionalAgents) "&_
					 "and isNull(BankInstrumentNo,'0') <> '1' "&_
					 "and H.PartyCode in (Select PartyCode From App_R_AgentOrgParty Where AgentCode = "&Trim(sTemp(3))&") "&_
					 "Order By 1 "
					 
					 '"and P.AmountReceivable > P.AmountReceived Order By 1 "
		Elseif CStr(sBookCode) = "04" Then 'Sales Invoice
			sQuery = "select H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
					 "H.PayToRecdFrom,P.PayablesNumber from Acc_T_VoucherHeader H, Acc_T_Payables P "&_
					 "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='"&sBookCode&"' and H.PartyType='"&trim(sTemp(0))&"' "&_
					 "and H.TransactionNumber = P.TransactionNumber "&_
					 "and H.CreatedTransNo Not In (Select Distinct isNull(AccTransactionNo,0) From Sal_T_AdditionalAgents) "&_
					 "and isNull(BankInstrumentNo,'0') <> '1' "&_
					 "and H.PartyCode in (Select PartyCode From App_R_AgentOrgParty Where AgentCode = "&Trim(sTemp(3))&") "&_
					 "Order By 1 "
					 '"and P.AmountPayable > P.AmountPaid Order By 1 " 
					 
					 
		End IF
		
		

		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		Do While Not objRs.EOF 
			dCrValue = CDbl(dCrValue)
			Set newElem = OutData.createElement("SalInv")
			newElem.setAttribute "TransNo", trim(objRs(0))
			newElem.setAttribute "Amount", trim(objRs(3))
			newElem.setAttribute "InvDetails", trim(objRs(1))&"-"&trim(objRs(2))
			newElem.setAttribute "ReferenceNo", trim(objRs(4))
			newElem.setAttribute "TotalCrDrValue", dCrValue
			IF CStr(sBookCode) = "05" Then
				newElem.setAttribute "FromValue", "P"
			Else
				newElem.setAttribute "FromValue", "S"
			End IF
			Root.appendChild newElem
			objRs.MoveNext
		Loop
		objRs.Close
 	    
	Else ' Check For the sCallTy
		
		sQuery = "select H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
					 "H.PayToRecdFrom,P.ReceivableNumber from Acc_T_VoucherHeader H, Acc_T_CreatedReceivables P "&_
					 "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='05' and H.PartyType='"&trim(sTemp(0))&"' "&_
					 "and H.PartySubType = "&trim(sTemp(1))&" and  H.PartyCode = "&trim(sTemp(3))&" and H.CreatedTransNo = P.CreatedTransNo "&_
					 "Order By 1 "
					 '"and P.AmountReceivable > P.AmountReceived Order By 1 "
					 
					 
		sQuery2 = "select H.CreatedTransNo,H.VoucherNumber,convert(char,H.VoucherDate,103),H.VoucherAmount, "&_
				  "H.PayToRecdFrom,P.PayablesNumber from Acc_T_VoucherHeader H, Acc_T_Payables P "&_
				  "where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='04' and H.PartyType='"&trim(sTemp(0))&"' "&_
				  "and H.PartySubType = "&trim(sTemp(1))&" and  H.PartyCode = "&trim(sTemp(3))&" and H.TransactionNumber = P.TransactionNumber "&_
				  "Order By 1 "
				  '"and P.AmountPayable > P.AmountPaid Order By 1 " 		
				  
				  
		
			
	
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing

		if not objRs.EOF then
			do while not objRs.EOF
				Set newElem = OutData.createElement("SalInv")
				newElem.setAttribute "TransNo", trim(objRs(0))
				newElem.setAttribute "Amount", trim(objRs(3))
				newElem.setAttribute "InvDetails", trim(objRs(1))&"-"&trim(objRs(2))
				newElem.setAttribute "ReferenceNo", trim(objRs(4))
				newElem.setAttribute "TotalCrDrValue",0
				newElem.setAttribute "FromValue", "S"
				Root.appendChild newElem
			objRs.MoveNext
			loop
		end if
		objRs.Close
		
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery2
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing

		if not objRs.EOF then
			do while not objRs.EOF
				Set newElem = OutData.createElement("SalInv")
				newElem.setAttribute "TransNo", trim(objRs(0))
				newElem.setAttribute "Amount", trim(objRs(3))
				newElem.setAttribute "InvDetails", trim(objRs(1))&"-"&trim(objRs(2))
				newElem.setAttribute "ReferenceNo", trim(objRs(4))
				newElem.setAttribute "TotalCrDrValue",0
				newElem.setAttribute "FromValue", "P"
				Root.appendChild newElem
			objRs.MoveNext
			loop
		end if
		objRs.Close
		
	End IF
	
	Response.ContentType="text/xml"
	Response.Write OutData.xml
	
%>
