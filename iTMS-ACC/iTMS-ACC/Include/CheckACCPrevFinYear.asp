
<%
	' Function to Check for Fin. Year
	Function CheckFinantialYear()
		' Declaration of variables
		Dim dcrs,rsTemp
		dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin
		dim sFinPeriodFrom, sFinPeriodTo,bActiveFinYear,bStockClosed,bAccClosed
		dim sOpMonthYr,sClMonthYr
		'Response.Write Session("FinPeriod")
		'Declaration of Objects
		sFinPeriodFrom = FormatDate("01/04/" & Mid(Session("FinPeriod"),1,4))
		sFinPeriodTo = FormatDate("31/03/" & Mid(Session("FinPeriod"),6,4))

		sOpMonthYr = "04"&Mid(Session("FinPeriod"),1,4)
		sClMonthYr = "03"&Mid(Session("FinPeriod"),6,4)

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		set rsTemp = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CONVERT(datetime,Fromperiod,103) AS fromperiod,CONVERT(datetime,ToPeriod,103) AS ToPeriod FROM Ms_FinancialPeriod WHERE Active = 'Y'"
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.source
		set dcrs.ActiveConnection = nothing
		if Not dcrs.EOF then
			sFinFrom = trim(dcrs(0))
			sFinTo = trim(dcrs(1))
		End If
		dcrs.close
		'Response.Write Formatdate(sFinFrom)& ","& sFinPeriodFrom&"<br>"
		'Response.Write DateDiff("d",sFinPeriodFrom,Formatdate(sFinFrom))&"<br>"

		If DateDiff("d",sFinPeriodFrom,sFinFrom) = 0 Then
			'CheckFinantialYear = True
			bActiveFinYear = True
		Else
			'CheckFinantialYear = False
			bActiveFinYear = False
			'Exit Function
		End If


		with rsTemp
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			'.Source = "SELECT DISTINCT ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK"
			.Source = "SELECT DISTINCT OUDEFINITIONID FROM Acc_T_GLAccOpeningAmt"
			.Open
		end with
		if rsTemp.EOF then
			bAccClosed = True
		else
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					''''''.Source = "SELECT DISTINCT ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME," & Pack(Formatdate(sFinFrom)) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(Formatdate(sFinTo)) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
					'Added on 31Mar 2009
					'.Source = "SELECT DISTINCT ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME," & Pack(FormatDate(sFinPeriodFrom)) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinPeriodTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
					.Source = "SELECT DISTINCT OUDEFINITIONID FROM ACC_T_GLACCOPENINGAMT WHERE OPENINGMONTHYEAR = '"& sOpMonthYr &"' AND CLOSINGMONTHYEAR ='"& sClMonthYr &"' "
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				'Response.Write "<p>aaa=" & dcrs.source

				if Not dcrs.EOF then
					'CheckFinantialYear = True
					bAccClosed = True
				else
					'CheckFinantialYear = False
					bAccClosed = False
				end if
				dcrs.Close
		end if ' if rsTemp.eof then

	'	Response.write bActiveFinYear &"--" & bStockClosed

		'If bActiveFinYear and bStockClosed Then
  		'	CheckFinantialYear = 1 'Active Finanacial Year and Stock closed
  		''ElseIf bActiveFinYear and Not bStockClosed Then
  		''	CheckFinantialYear = 4 'Active financial year & stock not closed
  		'ElseIf Not bActiveFinYear and Not bStockClosed Then
  		'	CheckFinantialYear = 2 'Not Active financial year & stock not closed
  		'Else'If Not bActiveFinYear and bStockClosed Then
  		'	CheckFinantialYear = 3 'Not Active Fiancial Year but Stock is closed
  		'End If

  		if bActiveFinYear then
			if bAccClosed then
				CheckFinantialYear = 1 'Active Finperiod and ACC Closed
			else
				CheckFinantialYear = 2 'Active Fin Period and ACC not close
			end if
		else
			if bAccClosed then
				CheckFinantialYear = 3 'Inactive Finperiod and Acc closed
			else
				CheckFinantialYear = 4 'Inactive Fin Period and ACC not close
			end if
		end if


		'Response.Write CheckFinantialYear

	End Function
%>
<%
If CheckFinantialYear = 2 Then%>
  <SCRIPT LANGUAGE=vbscript>
	alert("Since Accounts is not closed for the login Financial Year. Transaction Can not be Done in this Financial Year...!")
  	window.history.back(1)
  </script>

  <!--<%'ElseIf CheckFinantialYear = 3 Then%>

    <SCRIPT LANGUAGE=vbscript>
    	alert("Since Finacial Year and Accounts Closing is done. Transaction Can not be Done in this Financial Year...!")
    	window.history.back(1)
  </script>

  <%'ElseIf CheckFinantialYear = 4 Then%>

  <SCRIPT LANGUAGE=vbscript>
  	alert("Since Year End Accounts Closing is not done. Transaction Can not be Done in this Financial Year...!")
  	window.history.back(1)
  </script>-->

<%
  End If

 %>