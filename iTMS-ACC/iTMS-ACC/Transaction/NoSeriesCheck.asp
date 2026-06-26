<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	NoSeriesCheck.asp
	'Module Name				:	ACCOUNTS (Transcation - No Series Checking)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	Apr 25,2005
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
<!--#include file="../../include/NoSeries.asp"-->

<%
	Dim ObjRsSer,sCheckSql,sNoCheckVal,sTemparr,sSelVal,sResVal,sTempMonYrPs
	Dim sMonYrPs,arrFinPs,sFinFromPs,sFinToPs,sActiveYears,iSerNo,iSerCode
	Dim iGetPerInvs,sNoDates,sNumbers,sSerType
	sSelVal = Request("sValue")
	sTemparr = Split(sSelVal,":")
	
	if len(Month(date())) = 1 then
		sTempMonYrPs = "0"&Month(date())
	else
		sTempMonYrPs = Month(date())
	end if
	sMonYrPs = sTempMonYrPs&Year(date())

	arrFinPs = split(GetFinancialYearP(sMonYrPs),":")
	sFinFromPs = arrFinPs(0)
	sFinToPs = arrFinPs(1)

	if right(sFinFromPs,4)&":"&right(sFinToPs,4) = Session("FinPeriod") then
		sActiveYears = "Y"
	else
		sActiveYears = "N"
	end if
	
	'Response.Write sActiveYears

	if sActiveYears = "N" then 
		sNoDates = "01/03/"&right(sFinFromPs,4)
	Else
		sNoDates = sTemparr(5)
	End IF

	
	sResVal = CheckSeries(sTemparr(0),sTemparr(2),sTemparr(1),sTemparr(4),sTemparr(3),sTemparr(5))
	Response.Write sResVal
	
	Function CheckSeries(sUnit,sCallTy,sBookCode,sBookNo,sVouTy,sPeriod)
	
		Set ObjRsSer = Server.CreateObject("ADODB.RecordSet")
		
		IF CStr(sCallTy) = "C" Then 'For Created Voucher 
			IF CStr(sBookCode) = "01" or CStr(sBookCode) = "02" Then 'If it is of Cash/Bank Voucher
				IF CStr(sVouTy) = "D" Then 'For Receipt Voucher 
					sCheckSql = "Select B.CreatedDrSeriesNo,B.CreatedDrSeriesCode From Acc_M_BookNumberSeries B "&_
								"where B.OUDefinitionID='"&sUnit&"' and B.BookCode='"&sBookCode&"' "&_
								"and B.BookNumber= "&sBookNo&" "
				Else ' For Payment Voucher 
					sCheckSql = "Select B.CreatedCrSeriesNo,B.CreatedCrSeriesCode From Acc_M_BookNumberSeries B "&_
								"where B.OUDefinitionID='"&sUnit&"' and B.BookCode='"&sBookCode&"' "&_
								"and B.BookNumber= "&sBookNo&" "
				End IF
			Else ' if it is Other Than Cash or Bank Voucher Then
				sCheckSql = "Select B.CreatedDrSeriesNo,B.CreatedDrSeriesCode From Acc_M_BookNumberSeries B "&_
							"where B.OUDefinitionID='"&sUnit&"' and B.BookCode='"&sBookCode&"' "&_
							"and B.BookNumber= "&sBookNo&" "
								
			End IF
		Else ' For Accounted Voucher 
			IF CStr(sBookCode) = "01" or CStr(sBookCode) = "02" Then 'If it is of Cash/Bank Voucher
				IF CStr(sVouTy) = "D" Then 'For Receipt Voucher 
					sCheckSql = "Select B.DrSeriesNo,B.DrSeriesCode From Acc_M_BookNumberSeries B "&_
								"where B.OUDefinitionID='"&sUnit&"' and B.BookCode='"&sBookCode&"' "&_
								"and B.BookNumber= "&sBookNo&" "
				Else ' For Payment Voucher 
					sCheckSql = "Select B.CrSeriesNo,B.CrSeriesCode From Acc_M_BookNumberSeries B  "&_
								"where B.OUDefinitionID='"&sUnit&"' and B.BookCode='"&sBookCode&"' "&_
								"and B.BookNumber= "&sBookNo&" "
				End IF
			Else ' if it is Other Than Cash or Bank Voucher Then
				sCheckSql = "Select B.DrSeriesNo,B.DrSeriesCode From Acc_M_BookNumberSeries B  "&_
							"where B.OUDefinitionID='"&sUnit&"' and B.BookCode='"&sBookCode&"' "&_
							"and B.BookNumber= "&sBookNo&" "
								
			End IF
		End IF
		
		With ObjRsSer
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = Con
			.Source = sCheckSql
			.Open
		End With
		Set ObjRsSer.ActiveConnection = Nothing
		IF Not ObjRsSer.EOF Then
			iSerNo = ObjRsSer(0)
			iSerCode = ObjRsSer(1)
		Else
			iSerNo = "0"
			iSerCode = "0"
		End IF
		ObjRsSer.Close
		
		sCheckSql = "select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
					""&iSerNo&" and OUDefinitionID='"&sUnit&"' and SeriesCode="&iSerCode

	with ObjRsSer
		.CursorLocation = 3
		.CursorType = 3
		.Source = sCheckSql
		.ActiveConnection = con
		.Open
	end with

	sSerType =ObjRsSer(0)
	ObjRsSer.close
	
	iGetPerInvs = GetPeriodInterval(sNoDates,sSerType)
	
	sCheckSql = "SELECT Number FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
				""&iSerNo&" and OUDefinitionID='"&sUnit&"' and SeriesCode="&iSerCode&" and "&_
				"Period='"&iGetPerInvs&"'"

	'Response.Write sCheckSql&"<BR>"

	with ObjRsSer
		.CursorLocation = 3
		.CursorType = 3
		.Source = sCheckSql
		.ActiveConnection = con
		.Open
	end with
	set ObjRsSer.ActiveConnection = nothing

	if  not ObjRsSer.EOF then
		sNumbers=trim(ObjRsSer(0))
	End if
	ObjRsSer.Close
	
	'sNoCheckVal = sNumbers

	If CDbl(sNumbers)=0 Then
		sNoCheckVal = "F"
	Else
		sNoCheckVal = "T"
	End If
	
	CheckSeries = sNoCheckVal
End Function
%>