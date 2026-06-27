<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->

<%
	Dim Objrs,sQuery,sOrgId,iBookAccHead,sDate,sTemp,Temparr,dOpenBal,dBookBal,dCurrBal
	sTemp = Request("sValue")
	Temparr = Split(sTemp,":")
	sOrgId = Temparr(0)
	iBookAccHead = Temparr(1)
	sDate = Temparr(2)
	'dBookBal = GetBookBal(sOrgId,iBookAccHead,sDate)
	'dBookBal = GetBookBalCr()
	'Response.Write sOrgId &" " & iBookAccHead &" " & FormatDate(Date+1)
	dBookBal = GetDayOpening(sOrgId,iBookAccHead,FormatDate(Date+1))
	'dCurrBal = GetCurrBal(sOrgId,iBookAccHead,sDate)
	dCurrBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
	dOpenBal = dBookBal&"*"&dCurrBal
	Response.Write dOpenBal
	
%>

<%
	Function GetBookBalCr()
		GetBookBalCr = GetDayOpening(sOrgId,iBookAccHead,FormatDate(Date+1))
	End Function
%>

<%
	
	Function GetBookBal(sOrgId,iAccHead,sFromdate)
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
			sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
				"WHERE  OUDEFINITIONID='"&sOrgId& "' AND ACCOUNTHEAD="&iAccHead&" and "&_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) <= "&Trim(saFromDate(2))&Trim(saFromDate(1))

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
			sQuery="select voucheramount,crdrindication from Acc_T_VoucherHeader where " &_
			"OUDEFINITIONID='"&sOrgId& "' and AccountHead='"&iAccHead & "' and " &_
			"convert(datetime,voucherdate,103) >= convert(datetime,'"& sMonthDay & "',103)" &_
			" and convert(datetime,voucherdate,103)<= convert(datetime,'"& sFromDate&"',103)"
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
		GetBookBal = dOpeningAmt
	End Function
%>

<%
Function GetCurrBal(sOrgId,iAccHead,sFromdate)
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
			sQuery="SELECT MONTHDRAMOUNT-MONTHCRAMOUNT FROM ACC_T_GLACCTRANSACTAMT "&_
				"WHERE OUDEFINITIONID='"&sOrgId&"' AND ACCOUNTHEAD="&iAccHead&" and "&_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) >="&sFinMonYear& " and " &_
				"substring(MonthYear,3,4)+substring(MonthYear,1,2) <= "&Trim(saFromDate(2))&Trim(saFromDate(1))

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
			" and convert(datetime,voucherdate,103)<= convert(datetime,'"& sFromDate&"',103)"

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
		GetCurrBal = dOpeningAmt
	End Function
%>