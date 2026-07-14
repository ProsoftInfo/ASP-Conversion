<!--#include virtual="/Accounts/Transaction/GetCBGVoucherxml.asp"-->
<!--#include virtual="/Accounts/Transaction/GetPSVoucherxml.asp"-->
<!--#include virtual="/Accounts/Transaction/GetCDNVoucherxml.asp"-->
<!--#include virtual="/Accounts/Transaction/GetDNVoucherxml.asp"-->
<!--#include virtual="/Accounts/Transaction/GetSalesReturnVouXml.asp"-->

<%
	Dim tempFinYear
	' Function to populate the Application list
	Function populateApplication()
		' Declaration of variables
		Dim dcrs,sAppID,sAppLName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT APPLICATIONCODE,APPLICATIONNAME,APPLNSHORTNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sAppID = dcrs(0)
		set sAppLName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="&trim(sAppID)&">"&trim(sAppLName)&"</OPTION>")
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Application list
	Function populateDayBooks()
		' Declaration of variables
		Dim dcrs,sBookID,sBookLName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT BookCode,BookName,BookshortName FROM Acc_M_DayBooks ORDER BY BookName"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sBookID = dcrs(0)
		set sBookLName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="&trim(sBookID)&">"&trim(sBookLName)&"</OPTION>")
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
	End Function
' Function to populate the PartyType list
Function populatePartyType(sorgID)
	' Declaration of variables
	Dim objRs,objRs1,iParSubType,sParType,sParSubTypeName,sQuery
	'Declaration of Objects
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")

	sQuery="select distinct PartyType,PartySubType,SubTypeName from vwOrgPartyType  where OUDefinitionID='"&sOrgId&"'"

	with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
	end with

	set objRs.ActiveConnection=nothing

	set sParType=objRs(0)
	set iParSubType = objRs(1)
	set sParSubTypeName = objRs(2)

	If not objRs.EOF then

		Do While Not objRs.EOF
			sQuery="select count(1) from APP_R_OrgParty where PartyType='"&sParType&"' and PartySubType="&iParSubType&" and OUDefinitionID='"&sOrgId&"'"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with

			set objRs1.ActiveConnection=nothing

			if CDbl(objRs1(0)) then
				Response.Write("<OPTION VALUE="&trim(sParType)&"?"&trim(iParSubType)&">"&trim(sParType)&"-"&trim(sParSubTypeName)&"</OPTION>")
			end if

			objRs1.Close

			objRs.MoveNext
		Loop
	end if
	objRs.Close
	set objRs=nothing
	set objRs1=nothing
End Function

' Function to populate the Frequently Used List
Function popFrequentHead(sorgID,sBookCode,iBookNo)
dim sCode,sValue,bCostcenter,bAnalytical,sTranFalg,iHeadCount,iTdsElgi

sQuery = "Select AccountHead,AccountDescription,CostCenterExists,AnalyticalHeadExists,"&_
"AllowTransactions,EligibleForTds from VwBooksFrequentlyUsedGL where OUDefinitionID='"&sOrgId&"'"&_
" and BookCode='"&sBookCode&"' and BookNumber="&iBookNo&" Order By AccountDescription "
	With objRs
  		.CursorLocation = 3
  		.CursorType = 3
  		.Source = sQuery
  		.ActiveConnection = con
  		.Open
  	End with
  	Set objRs.Activeconnection = nothing
  	iHeadCount=objRs.RecordCount

  	Set sCode = objRs(0)
  	Set sValue = objRs(1)
  	set bCostcenter= objRs(2)
  	set bAnalytical= objRs(3)
  	set sTranFalg= objRs(4)
  	Set iTdsElgi = objRs(5)

  	Do while not objRs.EOF
		Response.Write("<OPTION VALUE="""&sCode&"?"&bCostcenter&"?"&bAnalytical&"?"&sTranFalg&"?"&iTdsElgi&""">"&sValue&"</OPTION>")
		objRs.MoveNext
	Loop
	objRs.Close
	popFrequentHead=iHeadCount
End Function

Function getUserid()
		'getUserid=1234
		getUserid=session("userid")
End Function

Function getFinancialYear()
'Response.Write Session("FinPeriod")
		'getFinancialYear="042004"
		'getFinancialYear=session("finacialYear")
		'Added on 01/04/2005 to get the current financial year
		tempFinYear = split(trim(Session("FinPeriod")),":")
		getFinancialYear = tempFinYear(0)
End Function

Function getFromFinYear()
		 'getFromFinYear="042004"
		'getFinancialYear=session("FromfinYear")
		'Added on 01/04/2005 to get the current financial year
		tempFinYear = split(trim(Session("FinPeriod")),":")
	'Response.Write tempFinYear(0)
		getFromFinYear = "04"&tempFinYear(0)
End Function

Function getToFinYear()
		 'getToFinYear="032005"
		'getFinancialYear=session("TofinYear")
		'Added on 01/04/2005 to get the current financial year
		tempFinYear = split(trim(Session("FinPeriod")),":")
		getToFinYear = "03"&tempFinYear(1)
End Function

Function getFromFinDate()
		 'getFromFinDate="01/04/2004"
			'getFromFinDate=session("FromfinYear")
			'Added on 01/04/2005 to get the current financial start date
			tempFinYear = split(trim(Session("FinPeriod")),":")
			getFromFinDate = "01/04/"&tempFinYear(0)
End Function

Function getToFinDate()
			'getToFinDate="31/03/2005"
			'getToFinDate=session("TofinYear")
			'Added on 01/04/2005 to get the current financial end date
			tempFinYear = split(trim(Session("FinPeriod")),":")
			getFromFinDate = "31/03/"&tempFinYear(1)
End Function

Function CurrentDate(Date1)
dim Date2,Da,Mo,Ye,Dastr

	 Date2 = Date1
	 Da = Day(Date2)
	 Mo = Month(Date2)
	 Ye = Year(Date2)
	 If Da < 10 then
	 Da = "0"&Da
	 End if
	 If Mo < 10 then
	 Mo = "0"&Mo
	 End if
	 Dastr = Da&"/"&Mo&"/"&Ye
	 CurrentDate = Dastr
End Function

Function IIf(AE)
		If AE = 1 then
			IIf = " "
		else
			IIf = "s "
		End if
End Function

Function AmountWords(Amount)

    Dim Paise
    ReDim Ones(20)
    ReDim Tens(10)
    Dim Hundred
    ReDim Ws(5)

    Amount = Trim(Amount)
    Amount = CDbl(Amount)

    Dim A1, S1, Crt, T1
    Dim ReturnStr
	Amount=cdbl(Amount)
    Paise = Amount - Int(Amount)
    Ones(1) = "One "
    Ones(2) = "Two "
    Ones(3) = "Three "
    Ones(4) = "Four "
    Ones(5) = "Five "
    Ones(6) = "Six "
    Ones(7) = "Seven "
    Ones(8) = "Eight "
    Ones(9) = "Nine "
    Ones(10) = "Ten "
    Ones(11) = "Eleven "
    Ones(12) = "Twelve "
    Ones(13) = "Thirteen "
    Ones(14) = "Fourteen "
    Ones(15) = "Fifteen "
    Ones(16) = "Sixteen "
    Ones(17) = "Seventeen "
    Ones(18) = "Eighteen "
    Ones(19) = "Nineteen "
    Tens(1) = "Ten "
    Tens(2) = "Twenty "
    Tens(3) = "Thirty "
    Tens(4) = "Fourty "
    Tens(5) = "Fifty "
    Tens(6) = "Sixty "
    Tens(7) = "Seventy "
    Tens(8) = "Eighty "
    Tens(9) = "Ninety "
    Hundred = "Hundred"
    Ws(1) = "Crore"
    Ws(2) = "Lakh"
    Ws(3) = "Thousand"

    A1 = Int(Amount)
    Crt = 9999999
    S1 = 1
    'ReturnStr = "Rupees "

    Do While A1 > 999
        If A1 > Crt Then
            T1 = Int(A1 / (Crt + 1))
            ReturnStr = ReturnStr + Some_Pro(T1, Ones, Tens)
            ReturnStr = ReturnStr + Ws(S1) + IIf(T1)
            A1 = Int(A1 Mod (Crt + 1))
        End If
        Crt = Int((Crt Mod (Crt + 1)) / 100)
        S1 = S1 + 1
    Loop

    If A1 > 99 Then
        T1 = Int(A1 / 100)
        ReturnStr = ReturnStr + Ones(T1) + Hundred + IIf(T1)
        A1 = A1 Mod 100
    End If

    If A1 > 0 Then
        ReturnStr = ReturnStr + Some_Pro(A1, Ones, Tens)
        'DO SOME_PRO WITH A1, returnStr
    End If

    If Int(Amount) > 0 Then  ReturnStr = "Rupees "+ReturnStr



    If Paise <> 0 Then
    	If Int(Amount) > 0 Then ReturnStr = ReturnStr + "And "
		ReturnStr = ReturnStr + "Paise "
		'Response.Write Returnstr
        ReturnStr = ReturnStr + Some_Pro(Round(Paise * 100,2), Ones, Tens)
    End If
    ReturnStr = ReturnStr + "Only"
    AmountWords = ReturnStr
End Function

Function Some_Pro(TT1, Ons, Tes)
    Dim SReturnStr
    'Response.Write TT1
    'Response.Write "<BR>"
    If TT1 < 20 Then
        SReturnStr = SReturnStr + Ons(TT1)
    Else
        SReturnStr = SReturnStr + Tes(Int(TT1 / 10))
        If TT1 Mod 10 <> 0 Then
            SReturnStr = SReturnStr + Ons(TT1 Mod 10)
        End If
    End If
    'Response.Write sReturnstr
    Some_Pro = SReturnStr

End Function
%>

<%
	' Function to format the date in dd/mm/yyyy
	Function FormatDate(Date1)
		dim dDate2,sDa,sMo,sYe,sDastr
		dDate2 = Date1
		sDa = Day(dDate2)
		sMo = Month(dDate2)
		sYe = Year(dDate2)

		If sDa < 10 then
			sDa = "0"&sDa
		End if

		If sMo < 10 then
			sMo = "0"&sMo
		End if

		sDastr = sDa&"/"&sMo&"/"&sYe
		FormatDate = sDastr
	End Function
%>

<%
	Function GetDayOpening(sOrgId,iAccHead,sFromdate)
	dim sQuery,saFromDate,sFinMonYear,dOpeningAmt,sMonthDay,dcrs
	dOpeningAmt=0
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
		saFromDate=split(sFromdate,"/")
		sFinMonYear=GetFromFinYear

		sQuery ="SELECT OPENINGAMOUNT,OPENINGCDINDICATION FROM ACC_T_GLACCOPENINGAMT"&_
				" WHERE ACCOUNTHEAD=" & iAccHead & " AND OPENINGMONTHYEAR="& sFinMonYear &_
				" and OUDefinitionID='"&sOrgId&"'"


		'Response.Write "<p> sQuery = "& sQuery &"<br>"

		WITH dcrs
			.cursorLocation=3
			.cursortype=3
			.activeconnection=con
			.source=sQuery
			.open
		End with
		Set dcrs.ActiveConnection=nothing

		if not dcrs.EOF then
			if dcrs(1)="C" then
				dOpeningAmt=CDbl(dcrs(0))*-1
			else
				dOpeningAmt=CDbl(dcrs(0))
			end if
		End if

		dcrs.close
		sFinMonYear=Mid(GetFromFinYear,3,4)&Mid(GetFromFinYear,1,2)

		if CDbl (saFromDate(2)&saFromDate(1)) > CDbl(sFinMonYear) then
			sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
				"WHERE  OUDEFINITIONID='"&sOrgId& "' AND ACCOUNTHEAD="&iAccHead&" and "&_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))


			 'Response.Write sQuery &"<BR>"


			with dcrs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing
			While not dcrs.EOF

			dOpeningAmt=CDbl(dOpeningAmt)+ cdbl(dcrs(0))
			' Response.Write  "<BR>dOpeningAmt="&dOpeningAmt&"<BR><BR>"
			dcrs.MoveNext
			wend
			dcrs.Close
		end if
		'Response.Write saFromDate(0)
		if  CInt(saFromDate(0))> 1 then
			sMonthDay="01/"&saFromDate(1)&"/"&saFromDate(2)

			'sQuery="select voucheramount,crdrindication from Acc_T_VoucherHeader where " &_
			'"OUDEFINITIONID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
			'"convert(datetime,voucherdate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
			'" and convert(datetime,voucherdate,103)<= convert(datetime,'"& sFromDate&"',103)" '= condition newly added by Maheswari on Feb 5th 2009 bec its is considering one day before last date ie(for april-- upto 29th ly fetching)

			'Added on May 26,2011 By UmaMaheswari S
			sQuery="select Amount,Transcrdrindication from Acc_T_GLTransactions where " &_
			"OUDEFINITIONID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
			"convert(datetime,voucherdate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
			" and convert(datetime,voucherdate,103)< convert(datetime,'"& sFromDate&"',103)"

			with dcrs
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing
		'	 Response.Write sQuery&"<br>"
		'	 Response.Write "AAAAAA="& dcrs.EOF

			While not dcrs.EOF

				if dcrs(1)="C" then
					dOpeningAmt=CDbl(dOpeningAmt)+(CDbl(dcrs(0))*-1)
				else
					dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dcrs(0))
				end if
			'	Response.Write "<BR>"& dcrs(1) &"---"& dOpeningAmt &"<BR><BR>"
				dcrs.MoveNext
			wend
			dcrs.Close
		end if
		' Response.Write dOpeningAmt
		GetDayOpening = dOpeningAmt
	End Function
%>

<%
	Function GetAccOpening(sOrgId,iAccHead,sFromdate,bSummary)

	'Response.Write "<p> sFromdate = " & sFromdate
	dim sQuery,saFromDate,sFinMonYear,dOpeningAmt,sMonthDay,dcrs
	dOpeningAmt=0
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
		saFromDate=split(sFromdate,"/")
		sFinMonYear=GetFromFinYear
		sQuery ="SELECT OPENINGAMOUNT,OPENINGCDINDICATION FROM ACC_T_GLACCOPENINGAMT"&_
		" WHERE ACCOUNTHEAD=" & iAccHead & " AND OPENINGMONTHYEAR="& sFinMonYear &_
		" and OUDefinitionID='"&sOrgId&"'"
		'Response.Write sQuery
		WITH dcrs
			.cursorLocation=3
			.cursortype=3
			.activeconnection=con
			.source=sQuery
			.open
		End with
		Set dcrs.ActiveConnection=nothing

		if not dcrs.EOF then
			if dcrs(1)="C" then
				dOpeningAmt=CDbl(dcrs(0))*-1
			else
				dOpeningAmt=CDbl(dcrs(0))
			end if
		End if

		dcrs.close
		sFinMonYear=Mid(GetFromFinYear,3,4)&Mid(GetFromFinYear,1,2)

		if CDbl (saFromDate(2)&saFromDate(1)) > CDbl(sFinMonYear) then

			if bSummary=0 and CInt(saFromDate(0))> 1 then
				sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
					"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))
			else
					sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
					"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) <= "&Trim(saFromDate(2))&Trim(saFromDate(1))
			end if
		 'Response.Write sQuery
			with dcrs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing

			While not dcrs.EOF
			dOpeningAmt=CDbl(dOpeningAmt)+ cdbl(dcrs(0))
			dcrs.MoveNext
			wend
			dcrs.Close
		end if

		if  CInt(saFromDate(0))> 1 and bSummary=0 then
			sMonthDay="01/"&saFromDate(1)&"/"&saFromDate(2)
			sQuery="select Amount,TransCrDrIndication from Acc_T_GLTransactions where " &_
			"OUDefinitionID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
			"convert(datetime,VoucherDate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
			" and convert(datetime,VoucherDate,103)< convert(datetime,'"& sFromDate&"',103)"
	'Response.Write sQuery
			with dcrs
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.Open
			End with

			Set dcrs.ActiveConnection=nothing

			While not dcrs.EOF
				if dcrs(1)="C" then
					dOpeningAmt=CDbl(dOpeningAmt)+(CDbl(dcrs(0))*-1)
				else
					dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dcrs(0))
				end if
				dcrs.MoveNext
			wend
			dcrs.Close
		end if
		GetAccOpening = dOpeningAmt
	End Function

%>

<%
Function GetDayOpeningCreated(sOrgId,iAccHead,sFromdate)
	dim sQuery,saFromDate,sFinMonYear,dOpeningAmt,sMonthDay,dcrs
	dOpeningAmt=0
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
		saFromDate=split(sFromdate,"/")
		sFinMonYear=GetFromFinYear
		sQuery ="SELECT OPENINGAMOUNT,OPENINGCDINDICATION FROM ACC_T_GLACCOPENINGAMT"&_
		" WHERE ACCOUNTHEAD=" & iAccHead & " AND OPENINGMONTHYEAR="& sFinMonYear &_
		" and OUDefinitionID='"&sOrgId&"'"

		'Response.Write sQuery
		WITH dcrs
			.cursorLocation=3
			.cursortype=3
			.activeconnection=con
			.source=sQuery
			.open
		End with

		Set dcrs.ActiveConnection=nothing

		if not dcrs.EOF then
			if dcrs(1)="C" then
				dOpeningAmt=CDbl(dcrs(0))*-1
			else
				dOpeningAmt=CDbl(dcrs(0))
			end if
		End if

		dcrs.close
		sFinMonYear=Mid(GetFromFinYear,3,4)&Mid(GetFromFinYear,1,2)

		if CDbl (saFromDate(2)&saFromDate(1)) > CDbl(sFinMonYear) then
			sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
				"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))

			'Response.Write sQuery
			with dcrs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing
			While not dcrs.EOF
			dOpeningAmt=CDbl(dOpeningAmt)+ cdbl(dcrs(0))
			dcrs.MoveNext
			wend
			dcrs.Close
		end if

		if  CInt(saFromDate(0))> 1 then
			sMonthDay="01/"&saFromDate(1)&"/"&saFromDate(2)
			sQuery="select voucheramount,crdrindication from Acc_T_CreatedVoucherHeader where " &_
			"OUDEFINITIONID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
			"convert(datetime,voucherdate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
			" and convert(datetime,voucherdate,103)< convert(datetime,'"& sFromDate&"',103)"

			'Response.Write sQuery
			with dcrs
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing

			While not dcrs.EOF
				if dcrs(1)="C" then
					dOpeningAmt=CDbl(dOpeningAmt)+(CDbl(dcrs(0))*-1)
				else
					dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dcrs(0))
				end if
				dcrs.MoveNext
			wend
			dcrs.Close
		end if
		GetDayOpeningCreated = dOpeningAmt
	End Function
%>

<%
	Function GetPartyDayOpening(sOrgId,sPartyType,iPartySubType,iPartyCode,sFromdate)
	dim sQuery,saFromDate,sFinMonYear,dOpeningAmt,sMonthDay,dcrs
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

		saFromDate=split(sFromdate,"/")
		sFinMonYear=GetFromFinYear

		sQuery ="SELECT OPENINGAMOUNT,OPENINGCDINDICATION FROM Acc_T_PartyOpeningAmt"&_
			" WHERE Partycode='"&iPartycode & "' and " &_
			"Partytype='"&sPartyType & "' and Partysubtype=" & iPartySubType &_
			" and OPENINGMONTHYEAR='"& sFinMonYear &"' and OUDefinitionID='"&sOrgId&"'"



		WITH dcrs
			.cursorLocation=3
			.cursortype=3
			.activeconnection=con
			.source=sQuery
			.open
		End with

		Set dcrs.ActiveConnection=nothing

		if not dcrs.EOF then
			if dcrs(1)="C" then
				dOpeningAmt=CDbl(dcrs(0))*-1
			else
				dOpeningAmt=CDbl(dcrs(0))
			end if
		End if
		dcrs.Close

		sFinMonYear=Mid(GetFromFinYear,3,4)&Mid(GetFromFinYear,1,2)

		'Response.Write CDbl (saFromDate(2)&saFromDate(1)) &" " & CDbl(sFinMonYear) &"<br>"

		if CDbl (saFromDate(2)&saFromDate(1)) > CDbl(sFinMonYear) then

			sQuery ="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM Acc_T_PartyTransactAmt "&_
				" WHERE OUDEFINITIONID='"&sOrgId&"' AND Partycode='"&iPartycode & "' and " &_
				" Partytype='"&sPartyType & "' and Partysubtype=" & iPartySubType &_
				" and substring(MonthYear,3,4)+substring(MonthYear,1,2) > ="&sFinMonYear& " and " &_
				" substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))


			'Response.Write sQuery &"<br>"

			with dcrs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with

			Set dcrs.ActiveConnection=nothing

			While not dcrs.EOF
				dOpeningAmt=CDbl(dOpeningAmt)+ cdbl(dcrs(0))
				dcrs.MoveNext
			wend
			dcrs.Close
		end if

		sMonthDay="01/"&saFromDate(1)&"/"&saFromDate(2)

		if CInt(saFromDate(0))> 1 then
			sQuery="select voucheramount,crdrindication from Acc_T_VoucherHeader where " &_
			" OUDEFINITIONID='"&sOrgId& "' and Partycode='"&iPartycode & "' and " &_
			" Partytype='"&sPartyType & "' and Partysubtype=" & iPartySubType &_
			" and convert(datetime,voucherdate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
			" and convert(datetime,voucherdate,103)< convert(datetime,'"& sFromDate&"',103)"


			with dcrs
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing

			While not dcrs.EOF
				if dcrs(1)="C" then
					dOpeningAmt=CDbl(dOpeningAmt)+(CDbl(dcrs(0))*-1)
				else
					dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dcrs(0))
				end if
				dcrs.MoveNext
			wend
			dcrs.Close
		end if

		'Response.Write dOpeningAmt

		GetPartyDayOpening = dOpeningAmt
	End Function
%>

<%
	Function LastDayOfMonth(DateIn)
    	Dim TempDate, sMonth
    	sMonth = Month(DateIn)
    	If Len(Month(DateIn)) = 1 Then sMonth = "0"&sMonth
    	TempDate = Year(dateIn) & "-" & sMonth & "-"
    	If IsDate(TempDate & "31") Then
    		LastDayOfMonth = 31
    	ElseIf IsDate(TempDate & "30") Then
    		LastDayOfMonth = 30
    	ElseIf IsDate(TempDate & "29") Then
    		LastDayOfMonth = 29
    	ElseIf IsDate(TempDate & "28") Then
    		LastDayOfMonth = 28
    	End If
    End function
%>

<%
	Function popIUTUnits(sOrgId)
	dim sToOrgName,sToOrgId
		sQuery = "Select ToOUDefinitionID,OrgUnitShortDescription  from VwIUTDetails where FromOUDefinitionID='"&sOrgId&"' and ToOUDefinitionID in ("&_
			"Select FromOUDefinitionID from VwIUTDetails where ToOUDefinitionID='"&sOrgId&"')"

		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End with

		Set objRs.Activeconnection = nothing

		Set sToOrgId = objRs(0)
		Set sToOrgName = objRs(1)

		while not objRs.EOF
			Response.Write "<option value="""&sToOrgId&""">"&sToOrgName&"</options>"
			objRs.MoveNext
		wend
		objRs.Close
    End function
%>

<%
function CalculateTax(oNodTaxRoot,sFormula,dBValue,dDValue,dPercentage)
dim saTemp,dTaxAmount,iCounter,iTemp
dim oNodTemp
dim saTemp1

saTemp=Split(sFormula,",")
if trim(saTemp(0))="BV" then
	dTaxAmount=dBValue
	iTemp=1
elseif trim(saTemp(0))="BD" then
	dTaxAmount=dDValue
	iTemp=1
else
	dTaxAmount=0
	iTemp=0
end if

for iCounter=iTemp to UBound(saTemp)
	saTemp1=Split(trim(saTemp(iCounter)),"#")
	For Each oNodTemp in oNodTaxRoot.childNodes
		if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
			dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(5).nodeValue)
		end if
	next
next

if trim(dPercentage)<>"" then
	CalculateTax=dTaxAmount*(cdbl(dPercentage)/100)
else
	CalculateTax=dTaxAmount
end if
End function

Function GetPartyName(sPartyCode)
	Dim sQuery,sPartyName,objrs
	sQuery = "Select PartyName From App_M_PartyMaster Where PartyCode = "&sPartyCode&" "
	Set objrs = Server.CreateObject("ADODB.RecordSet")
	With objrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objrs.ActiveConnection = nothing
	IF not objrs.EOF Then
		sPartyName = objrs(0)
	End IF
	objrs.close
	GetPartyName = sPartyName
End Function

%>

<%
	Function GetAccOpeningBal(sOrgId,iAccHead,sFromdate,bSummary)
	dim sQuery,saFromDate,sFinMonYear,dOpeningAmt,sMonthDay,dcrs
	dOpeningAmt=0
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
		saFromDate=split(sFromdate,"/")
		sFinMonYear=GetFromFinYear
		sQuery ="SELECT OPENINGAMOUNT,OPENINGCDINDICATION FROM ACC_T_GLACCOPENINGAMT"&_
		" WHERE ACCOUNTHEAD=" & iAccHead & " AND OPENINGMONTHYEAR="& sFinMonYear &_
		" and OUDefinitionID='"&sOrgId&"'"

		WITH dcrs
			.cursorLocation=3
			.cursortype=3
			.activeconnection=con
			.source=sQuery
			.open
		End with
		Set dcrs.ActiveConnection=nothing

		if not dcrs.EOF then
			if dcrs(1)="C" then
				dOpeningAmt=CDbl(dcrs(0))*-1
			else
				dOpeningAmt=CDbl(dcrs(0))
			end if
		End if

		dcrs.close
		sFinMonYear=Mid(GetFromFinYear,3,4)&Mid(GetFromFinYear,1,2)

		if CDbl (saFromDate(2)&saFromDate(1)) > CDbl(sFinMonYear) then
			if bSummary=0 and CInt(saFromDate(0))> 1 then
				sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
					"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))
			else
					sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
					"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
					"substring(MonthYear,3,4)+substring(MonthYear,1,2) <= "&Trim(saFromDate(2))&Trim(saFromDate(1))
			end if

			with dcrs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing

			While not dcrs.EOF
			dOpeningAmt=CDbl(dOpeningAmt)+ cdbl(dcrs(0))
			dcrs.MoveNext
			wend
			dcrs.Close
		end if

		'if  CInt(saFromDate(0))> 1 and bSummary=0 then
		'	sMonthDay="01/"&saFromDate(1)&"/"&saFromDate(2)
		'	sQuery="select Amount,TransCrDrIndication from Acc_T_GLTransactions where " &_
		'	"OUDefinitionID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
		'	"convert(datetime,VoucherDate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
		'	" and convert(datetime,VoucherDate,103)< convert(datetime,'"& sFromDate&"',103)"
'
'			with dcrs
'				.CursorLocation =3
'				.CursorType=3
'				.ActiveConnection =con
'				.Source=sQuery
'				.Open
'			End with
'
'			Set dcrs.ActiveConnection=nothing
'
'			While not dcrs.EOF
'				if dcrs(1)="C" then
'					dOpeningAmt=CDbl(dOpeningAmt)+(CDbl(dcrs(0))*-1)
'				else
'					dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dcrs(0))
'				end if
'				dcrs.MoveNext
'			wend
'			dcrs.Close
'		end if
		GetAccOpeningBal = dOpeningAmt
	End Function

%>

<%
	Function GetVouchXML(iCrTransNo)
		Dim sVouBkCode,Objrs,sQry,sRetPath,objfs
		Set Objrs = Server.CreateObject("ADODB.RecordSet")
		Set objfs = CreateObject("Scripting.FileSystemObject")
		IF Cstr(iCrTransNo) = "" Then
			iCrTransNo = 0
		End IF


		'if objfs.FileExists(Server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")) then
		'	sRetPath = "../xmldata/Voucher/"&iCrTransNo&".xml"
		'Else
			sQry = "Select BookCode From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iCrTransNo
			'Response.Write sQry
			Objrs.open sQry,Con
			IF Not Objrs.Eof Then
				sVouBkCode = Objrs(0)
			End IF
			Objrs.Close

			Select	Case CStr(sVouBkCode)
				Case "01"
					sRetPath = GetXmlForCBG(iCrTransNo,"S")
				Case "02"
					sRetPath = GetXmlForCBG(iCrTransNo,"S")
				Case "04"
					sRetPath = GetXmlForPur(iCrTransNo,"S")
				Case "05"
					sRetPath = GetXmlForSal(iCrTransNo,"S")
				Case "07"
					sRetPath = GetXmlForCN(iCrTransNo,"S")
				Case "08"
					sRetPath = GetXmlForCBG(iCrTransNo,"S")
				Case "06"
					sRetPath = GetXmlForDN(iCrTransNo,"S")
			End Select
		'End IF

		GetVouchXML = sRetPath

	End Function
%>

<%
Function GetDayOpeningCreatedForPLBS(sOrgId,iAccHead,sFromdate)
	dim sQuery,saFromDate,sFinMonYear,dOpeningAmt,sMonthDay,dcrs,sGp
	dOpeningAmt=0
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
		saFromDate=split(sFromdate,"/")
		sFinMonYear=GetFromFinYear
		sQuery ="SELECT OPENINGAMOUNT,OPENINGCDINDICATION FROM ACC_T_GLACCOPENINGAMT"&_
		" WHERE ACCOUNTHEAD=" & iAccHead & " AND OPENINGMONTHYEAR="& sFinMonYear &_
		" and OUDefinitionID='"&sOrgId&"'"

		'Response.Write sQuery
		WITH dcrs
			.cursorLocation=3
			.cursortype=3
			.activeconnection=con
			.source=sQuery
			.open
		End with

		Set dcrs.ActiveConnection=nothing

		if not dcrs.EOF then
			if dcrs(1)="C" then
				dOpeningAmt=CDbl(dcrs(0))*-1
			else
				dOpeningAmt=CDbl(dcrs(0))
			end if
		End if

		dcrs.close
		sFinMonYear=Mid(GetFromFinYear,3,4)&Mid(GetFromFinYear,1,2)

		if CDbl (saFromDate(2)&saFromDate(1)) > CDbl(sFinMonYear) then
			sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
				"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))

			'Response.Write sQuery
			with dcrs
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing
			While not dcrs.EOF
			dOpeningAmt=CDbl(dOpeningAmt)+ cdbl(dcrs(0))
			dcrs.MoveNext
			wend
			dcrs.Close
		end if

		if  CInt(saFromDate(0))> 1 then
			sMonthDay="01/"&saFromDate(1)&"/"&saFromDate(2)
			sQuery="select voucheramount,crdrindication from Acc_T_CreatedVoucherHeader where " &_
			"OUDEFINITIONID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
			"convert(datetime,voucherdate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
			" and convert(datetime,voucherdate,103)< convert(datetime,'"& sFromDate&"',103)"

			'Response.Write sQuery
			with dcrs
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing

			While not dcrs.EOF
				if dcrs(1)="C" then
					dOpeningAmt=CDbl(dOpeningAmt)+(CDbl(dcrs(0))*-1)
				else
					dOpeningAmt=CDbl(dOpeningAmt)+CDbl(dcrs(0))
				end if
				dcrs.MoveNext
			wend
			dcrs.Close
		end if

		sQuery = "Select GroupCategory From VworgGlHeads Where Accounthead = "&iAccHead&" "
		dcrs.open sQuery,Con
		IF Not dcrs.eof	Then
			sGp = dcrs(0)
		End IF
		dcrs.close
		IF CStr(sGp) = "01" Then
			dOpeningAmt = dOpeningAmt * -1 'For Liablities
		Elseif CStr(sGp) = "02" then
			dOpeningAmt = dOpeningAmt * 1
		End IF

		GetDayOpeningCreatedForPLBS = dOpeningAmt
	End Function
%>

<%
	Function GetDayOpeningForSel(sOrgId,iAccHead,sFromdate,sType)
		dim sQuery,saFromDate,sFinMonYear,dOpeningAmt,sMonthDay,dcrs,sGp
		Dim dDrAmt,dCrAmt,iOldParCode,dParOpenAmt,dcrs2,iCurrParCode
		Dim sParTy
		dOpeningAmt = 0
		dDrAmt = 0
		dCrAmt = 0
		iOldParCode = 0
		iCurrParCode = 0
		dParOpenAmt = 0
		'Response.End

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
		saFromDate=split(sFromdate,"/")
		sFinMonYear=GetFromFinYear

		sQuery = "Select Distinct PartyType From VwOrgPartyType Where  "&_
				 "OUDefinitionID = '"&sOrgId&"' and AccountHead = "&iAccHead


		dcrs.open sQuery,Con
		IF Not dcrs.Eof Then
			sParTy = dcrs(0)
		Else
			sParTy = ""
		End IF
		dcrs.close

		IF CStr(sParTy) <> "" Then

			sQuery = "Select isNull(Sum(OpeningAmount),0),isNull(OpeningCDIndication,0),PartyCode From Acc_T_PartyOpeningAmt "&_
					 "Where OUDefinitionID = '"&sOrgId&"' and OpeningMonthYear = '"&sFinMonYear&"' and "&_
					 "PartyType = '"&sParTy&"'  "&_
					 "Group By OpeningCDIndication,PartyCode Order By PartyCode "

			'Response.Write sQuery &"<br>"
			with dcrs
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.Open
			End with
			Set dcrs.ActiveConnection=nothing
			dDrAmt = 0
			dCrAmt = 0
			IF Not dcrs.Eof Then
				iOldParCode = dcrs(2)
			End IF

			sFinMonYear=Mid(GetFromFinYear,3,4)&Mid(GetFromFinYear,1,2)
			While not dcrs.EOF
				'Response.Write "PartyCode=== " & dcrs(2) & "<br>"
				iCurrParCode = dcrs(2)
				IF CStr(iOldParCode) <> CStr(dcrs(2)) Then
					dParOpenAmt = CDbl(dDrAmt) + CDbl(dCrAmt)
					dDrAmt = 0
					dCrAmt = 0
					sQuery = "Select isNull(Sum(MonthDrAmount)-Sum(MonthCrAmount),0) From Acc_T_PartyTransactAmt "&_
							 "Where OUDefinitionID = '"&sOrgId&"' and PartyCode = "&iOldParCode&" and "&_
							 "PartyType = '"&sParTy&"' and "&_
							 "substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
							 "substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))



					with dcrs2
						.CursorLocation =3
						.CursorType=3
						.ActiveConnection =con
						.Source=sQuery
						.Open
					End with
					Set dcrs2.ActiveConnection=nothing
					IF Not dcrs2.EOF Then
						dParOpenAmt = dParOpenAmt + CDbl(dcrs2(0))
						IF CStr(sType) = "C" Then
							IF CDbl(dParOpenAmt) < 0 Then
								'Response.Write dParOpenAmt &" --> " & iOldParCode &"<br>"
								dOpeningAmt = dOpeningAmt + CDbl(Abs(dParOpenAmt))
							Else
								dOpeningAmt = dOpeningAmt + 0
							End IF
						Else
							IF CDbl(dParOpenAmt) > 0 Then
								'Response.Write dParOpenAmt &" --> " & iOldParCode &"<br>"
								dOpeningAmt = dOpeningAmt + CDbl(Abs(dParOpenAmt))
							Else
								dOpeningAmt = dOpeningAmt + 0
							End IF
						End IF
						dParOpenAmt = 0
					End IF

					dcrs2.Close
					iOldParCode = dcrs(2)
					IF CStr(dcrs(1)) = "C" Then
						dCrAmt = dcrs(0)
						dCrAmt = Cdbl(dCrAmt) * - 1
					Else
						dDrAmt = dcrs(0)
					End IF
				Else
					IF CStr(dcrs(1)) = "C" Then
						dCrAmt = dcrs(0)
						dCrAmt = Cdbl(dCrAmt) * - 1
					Else
						dDrAmt = dcrs(0)
					End IF
				End IF
				dcrs.MoveNext
			Wend
			dcrs.close

			sQuery = "Select isNull(Sum(MonthDrAmount)-Sum(MonthCrAmount),0) From Acc_T_PartyTransactAmt "&_
					 "Where OUDefinitionID = '"&sOrgId&"' and PartyCode = "&iCurrParCode&" and "&_
					 "PartyType = '"&sParTy&"' and "&_
					 "substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
					 "substring(MonthYear,3,4)+substring(MonthYear,1,2) < "&Trim(saFromDate(2))&Trim(saFromDate(1))

			with dcrs2
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.Open
			End with
			Set dcrs2.ActiveConnection=nothing
			IF Not dcrs2.EOF Then
				dParOpenAmt = dParOpenAmt + CDbl(dcrs2(0))
				IF CStr(sType) = "C" Then
					IF CDbl(dParOpenAmt) < 0 Then
						'Response.Write dParOpenAmt &" --> " & iOldParCode &"<br>"
						dOpeningAmt = dOpeningAmt + CDbl(Abs(dParOpenAmt))
					Else
						dOpeningAmt = dOpeningAmt + 0
					End IF
				Else
					IF CDbl(dParOpenAmt) > 0 Then
						'Response.Write dParOpenAmt &" --> " & iOldParCode &"<br>"
						dOpeningAmt = dOpeningAmt + CDbl(Abs(dParOpenAmt))
					Else
						dOpeningAmt = dOpeningAmt + 0
					End IF
				End IF
				dParOpenAmt = 0
			End IF
			'Response.Write dOpeningAmt
			GetDayOpeningForSel = dOpeningAmt
		Else
			GetDayOpeningForSel = 0
		End IF
	End Function
%>

<%
	Function CheckSummAccTy(iAcc,sOrgid)
		Dim dcrs,sQuery,sSummTy
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		sQuery = "Select Distinct SummaryPosting From VwOrgGLHeads  "&_
				 "Where AccountHead = "&iAcc&" and OUDefinitionID = '"&sOrgid&"' "

		'Response.Write sQuery
		dcrs.open sQuery,Con
		IF Not dcrs.eof Then
			sSummTy = dcrs(0)
		Else
			sSummTy = "0"
		End IF
		dcrs.close

		IF CStr(sSummTy) = "1" Then
			sSummTy = "Y" 'It is an Summary Posting Entry
		Else
			sSummTy = "N" 'It is Not an Summary Posting Entry
		End IF
		CheckSummAccTy = sSummTy
	End Function
%>

<%
	Function GetCtrlAccForParty(sParTy,sParSubTy,sUnit)
		Dim sQuery,dcrs,iAccHead
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		sQuery = "Select Distinct AccountHead From VwOrgPartyType Where  "&_
				 "OUDefinitionID = '"&sUnit&"' and PartyType = '"&sParTy&"' and "&_
				 "PartySubType = "&sParSubTy

		dcrs.open sQuery,Con
		IF Not dcrs.eof Then
			iAccHead = dcrs(0)
		Else
			iAccHead = "0"
		End IF
		dcrs.close
		GetCtrlAccForParty = iAccHead
	End Function

%>

<%
	Function CheckTBGL(sOrgID)
		Dim sQuery,dcrs,dDrAmt,dCrAmt,sFinStDate,sFinEndDate,sTemp,sRetVal,sStDates,sEnDates
		sTemp = Session("FinPeriod")
		sFinStDate = Trim(Left(sTemp,4))&"04"
		sFinEndDate = Trim(Right(sTemp,4))&"03"
		sStDates = "01/04/"&Trim(Left(sTemp,4))
		sEnDates = "31/03/"&Trim(Right(sTemp,4))

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		sQuery	= " Select isNull(Sum(MonthDrAMount),0) DR,isNull(Sum(MonthCrAmount),0) CR,isNull(Sum(MonthDrAMount),0)-isNull(Sum(MonthCrAMount),0) Diff "&_
				  " From Acc_T_GLAccTransactAmt Where  "&_
				  " OUDefinitionID = '"&sOrgID&"' and CAST(RTRIM(Cast(SUBString(MonthYear,3,4) AS Char))+RTRIM(Cast(SUBString(MonthYear,1,2) AS Char)) AS NUMERIC)>= "&sFinStDate&" "&_
				  " and CAST(RTRIM(Cast(SUBString(MonthYear,3,4) AS Char))+RTRIM(Cast(SUBString(MonthYear,1,2) AS Char)) AS NUMERIC) <= "&sFinEndDate&" "

		'Response.Write sQuery
		dcrs.open sQuery,Con
		IF Not dcrs.EOF Then
			'Response.Write "===================="
			dDrAmt = dcrs(0)
			dCrAmt = dcrs(1)
		End If
		dcrs.Close

		IF CDbl(dDrAmt) = CDbl(dCrAmt) Then
			dDrAmt = 0
			dCrAmt = 0
			sQuery = "Select Sum(Amount),TransCrDrIndication From  "&_
					 "Acc_T_GLTransactions Where OUDefinitionID = '"&sOrgID&"' "&_
					 "and Convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sStDates&"',103) "&_
					 "and Convert(datetime,VoucherDate,103) <= Convert(datetime,'"&sEnDates&"',103) "&_
					 "Group By TransCrDrIndication "

			'Response.write sQuery &"<br><br><br>"
			dcrs.open sQuery,Con
			Do While Not dcrs.EOF
				IF CStr(dcrs(1)) = "D" Then
					dDrAmt = dcrs(0)
				Else
					dCrAmt = dcrs(0)
				End IF
				dcrs.MoveNext
			loop
			dcrs.Close
			sRetVal = dDrAmt&":"&dCrAmt&":L"
		Else
			sRetVal = dDrAmt&":"&dCrAmt&":T"
		End IF

		CheckTBGL = sRetVal
	End Function
%>

<%


	Function GetTBDiff(sOrgID)
		Dim sQuery,dcrs,dDrAmt,dCrAmt,sFinStDate,sFinEndDate,sTemp,sRetVal
		sTemp = Session("FinPeriod")
		sFinStDate = Trim(Left(sTemp,4))&"04"
		sFinEndDate = Trim(Right(sTemp,4))&"03"
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		sQuery	= " Select Sum(MonthDrAMount)-Sum(MonthCrAMount) Diff "&_
				  " From Acc_T_GLAccTransactAmt Where  "&_
				  " OUDefinitionID = '"&sOrgID&"' and CAST(RTRIM(Cast(SUBString(MonthYear,3,4) AS Char))+RTRIM(Cast(SUBString(MonthYear,1,2) AS Char)) AS NUMERIC)>= "&sFinStDate&" "&_
				  " and CAST(RTRIM(Cast(SUBString(MonthYear,3,4) AS Char))+RTRIM(Cast(SUBString(MonthYear,1,2) AS Char)) AS NUMERIC) <= "&sFinEndDate&" "
		dcrs.open sQuery,Con
		IF Not dcrs.EOF Then
			dDrAmt = dcrs(0)
		End If
		dcrs.Close
		sRetVal = dDrAmt
		GetTBDiff = sRetVal
	End Function
%>

<% 'Added by Maheshwari on 15 Dec 2006
'Function to Find the Finacial year Closed or not
Function GetFinYear(sFinYear)
	Dim sQuery,sTemp,sFinStDate,sFinEndDate,sRetVal,dcrs,sClosed
	Set dcrs = Server.CreateObject("ADODB.Recordset")
	sTemp = sFinYear

	sFinStDate = "01"&"/"&"04"&"/"&Trim(Left(sTemp,4))
	sFinEndDate = "31"&"/"&"03"&"/"&Trim(Right(sTemp,4))
	'Response.Write "sFinStDate="& sFinStDate &"<BR><BR>"
	'Response.Write "sFinEndDate="& sFinEndDate &"<BR><BR>"
	sQuery = "Select Closed from Ms_FinancialPeriod where FromPeriod = Convert(datetime,'"&sFinStDate&"',103) and ToPeriod = Convert(datetime,'"&sFinEndDate&"',103) "
	'Response.Write "sqr=" & sQuery &"<BR><BR>"
	dcrs.open sQuery,Con
	If Not dcrs.EOF Then
		sClosed = dcrs(0)
	End If
	dcrs.close
'	Response.Write "sClosed="&sClosed&"<BR><BR>"
	If trim(sClosed) = "Y" Then
		sRetVal = "True"
	Else
		sRetVal = "False"
	End If
	GetFinYear = sRetVal
End Function
%>

<%
	Function GetLastDayMonYr(sYrMon)
		Dim sMon,sYear,sLastDay
		sMon = Right(sYrMon,2)
		sYear = Left(sYrMon,4)
		'Response.Write "<br>"
		'Response.Write sYear &"<br><br>"

		IF Cint(sMon) = 4 or Cint(sMon) = 6 or Cint(sMon) = 9 or Cint(sMon) = 11 Then
			sLastDay = "30/"&sMon&"/"&sYear
		Elseif CInt(sMon) <> 2 Then
			sLastDay = "31/"&sMon&"/"&sYear
		Elseif Cint(sYear) Mod 4 = 0 Then
			sLastDay = "29/"&sMon&"/"&sYear
		Else
			sLastDay = "28/"&sMon&"/"&sYear
		End IF
		'Response.Write sLastDay

		GetLastDayMonYr = sLastDay
	End Function
%>


<% 'Added by Maheshwari on 3rd April 2007 for TDS Percentage Calculation

function CalculatePer(Root,sFormula,sTDSAmt)
dim saTemp,iCounter,iTemp,dPercentage,sTotAmt
dim oNodTemp,sGrpCtr,sVal,sTempGrp
dim saTemp1,iCtr

saTemp=Split(sFormula,",")
'Response.Write sFormula
'sTotAmt = 0
iTemp = 0
iCtr = 1
sTempGrp = split(saTemp(0),"#")
sGrpId = sTempGrp(0)
dPercentage = Split(trim(saTemp(0)),"-")
If  trim(sGrpId) = "0" then
	sTotAmt=sTDSAmt*(cdbl(dPercentage(1))/100)
End If
For iCounter=iTemp to UBound(saTemp)

	dPercentage = Split(trim(saTemp(iCounter)),"-")
	saTemp1=Split(trim(saTemp(iCounter)),"#")
	sGrpId = saTemp1(0)
	sGrpCtr = left(saTemp1(1),1)

	For Each oNodTemp in Root.childNodes
		If trim(sGrpCtr) = oNodTemp.getAttribute("Ctr") then
				sVal = oNodTemp.getAttribute("PayRecAmount")

				sTotAmt = sTotAmt + sVal *(dPercentage(1) / 100)

			End If
	Next
next
CalculatePer = sTotAmt

End function
%>

<%
Function BreakString(sString,iCharLen)
    Dim iLoop , iCount , iCurrLen
    Dim sDum
    Dim sRetVal()
	ReDim sRetVal(100)
    iCount = 0

    While Len(Trim(sString)) > 0
        sDum = Mid(sString, 1, iCharLen)
        If Len(sString) > iCharLen Then
            For iLoop = 1 To iCharLen
				if iCharLen - iLoop > 0 then 'added on aug 06,2009 by kalaiselvi
					If Mid(sDum, iCharLen - iLoop, 1) = " " Then
					    iCurrLen = iCharLen - iLoop
					    Exit For
					End If
				else 'added on aug 06,2009 by kalaiselvi
					iCurrLen = iCharLen
					Exit For
				end if
            Next
        Else
            iCurrLen = Len(sString)
        End If

        sRetVal(iCount)  =  Mid(sString, 1, iCurrLen)


        iCount = iCount + 1
        sString = Mid(sString, iCurrLen + 1)
    Wend
    ReDim  Preserve sRetVal(iCount + 1)

    BreakString =  sRetVal

End Function


%>


<%
	Function GetContraStatus(iCrTransNo)
		Dim dcrs,sQuery,sType
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		sQuery = "Select Dbo.GetContraStatus("&iCrTransNo&") "
		dcrs.open sQuery,Con
		If Not dcrs.EOF Then
			sType = dcrs(0)
		End If
		dcrs.close
		GetContraStatus = sType
	End Function
%>

<%
	Function GetVouchSalReturnXML(iSalRetNo)
		Dim sVouBkCode,Objrs,sQry,sRetPath,objfs
		Set Objrs = Server.CreateObject("ADODB.RecordSet")
		Set objfs = CreateObject("Scripting.FileSystemObject")
		IF Cstr(iSalRetNo) = "" Then
			iSalRetNo = 0
		End IF

		sRetPath = GetXmlForSalesReturn(iSalRetNo)

		GetVouchSalReturnXML = sRetPath
	End Function
%>