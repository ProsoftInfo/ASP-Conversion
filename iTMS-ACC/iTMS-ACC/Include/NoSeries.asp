<%
Function GenSeriesNumberItemWise(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries,sSql,rsTemp
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery


	Set objRsSeries = Server.CreateObject("ADODB.RecordSet")
	set rsTemp = Server.CreateObject("ADODB.RecordSet")


	sSql = "Select ItemValue from VwPurNoSeriesSel where MainSeriesNo="&iSeriesNo&" and MainSeriesCode="&iSeriesCode&" "
	'Response.Write sSql
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sSql
		.Open
	end with
	set rsTemp.ActiveConnection = nothing
	if not rsTemp.EOF then
		sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

		'Response.Write " <p> squery =" & sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		if not objRsSeries.EOF then
			sType =objRsSeries(0)
			iLength=objRsSeries(1)
		end if
		objRsSeries.close

		sPeriod=GetPeriodInterval(sNoDate,sType)
		'response.write sPeriod
		sQuery="SELECT Number, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
		"Period='"&sPeriod&"'"
		'Response.Write sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		if  not objRsSeries.EOF then
			iNumber=trim(objRsSeries(0))
			sPrefix=trim(objRsSeries(1))
			sSufix=trim(objRsSeries(2))
		End if
		set objRsSeries=nothing
		if CDbl(iNumber)=0 then
			iNumber=1
		end if

		sQuery="Update APP_R_NoSeriesModuleEntry set Number="&CDbl(iNumber)+1&"  where SeriesNo="&_
			""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
			"Period='"&sPeriod&"'"
		con.execute(sQuery)
		if cint(iLength)>0 then
			iTemp=iLength-Len(iNumber)
			for iCounter=1 to iTemp
				iNumber="0"&iNumber
			next
			iNumber=sPrefix&iNumber&sSufix
		else
			iNumber=sPrefix&iNumber&sSufix
		end if

		GenSeriesNumberItemWise=iNumber
	End if
End function
%>

<%
Function GenSeriesNumber(sOrgid,iSeriesNo,iSeriesCode,sNoDate)


'Response.Write "<p> inside function =" + trim(iSeriesCode)
dim objRsSeries
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
	Set objRsSeries = Server.CreateObject("ADODB.RecordSet")

	sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode
'Response.Write "sNoDate = "& sNoDate
	'Response.Write " <p> squery =" & sQuery
	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	if not objRsSeries.EOF then
		sType =objRsSeries(0)
		iLength=objRsSeries(1)
	end if
	objRsSeries.close

	'Response.Write "<p> sNoDate = "  + trim(sNoDate)
	'Response.Write "<p> sType = "  + trim(sType)
	sPeriod=GetPeriodInterval(sNoDate,sType)
	'response.write "<p> sPeriod = " & sPeriod

	sQuery="SELECT Number, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
			""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
			"Period='"&sPeriod&"'"

	'Response.Write "<br><br>" & sQuery & "<br><br>"
	'Response.End

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	if  not objRsSeries.EOF then
		iNumber=trim(objRsSeries(0))
		sPrefix=trim(objRsSeries(1))
		sSufix=trim(objRsSeries(2))
	End if
	set objRsSeries=nothing
	if CDbl(iNumber)=0 then
		iNumber=1
	end if

	sQuery="Update APP_R_NoSeriesModuleEntry set Number="&CDbl(iNumber)+1&"  where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
		"Period='"&sPeriod&"'"
	'Response.Write sQuery &"<br><br>"


	con.execute(sQuery)
	if cint(iLength)>0 then
		iTemp=iLength-Len(iNumber)
		for iCounter=1 to iTemp
			iNumber="0"&iNumber
		next
		iNumber=sPrefix&iNumber&sSufix
	else
		iNumber=sPrefix&iNumber&sSufix
	end if

GenSeriesNumber=iNumber
End function
%>
<%
Function GetSeriesNumberLotNumber(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
Set objRsSeries = Server.CreateObject("ADODB.RecordSet")

	sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	sType =objRsSeries(0)
	iLength=objRsSeries(1)

	objRsSeries.close

	sPeriod=GetPeriodInterval(sNoDate,sType)

'Response.Write sPeriod&"<BR>"

	sQuery="SELECT Number-1, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
	"Period='"&sPeriod&"'"
'Response.Write sQuery&"<BR>"
	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	if  not objRsSeries.EOF then
		iNumber=trim(objRsSeries(0))
		sPrefix=trim(objRsSeries(1))
		sSufix=trim(objRsSeries(2))
	End if
	set objRsSeries=nothing
	if CDbl(iNumber)=-1 or CDbl(iNumber)=0 then
		iNumber=0
	end if

	if cint(iLength)>0 then
		iTemp=iLength-Len(iNumber)
		for iCounter=1 to iTemp
			iNumber="0"&iNumber
		next
	'	iNumber=sPrefix&iNumber&sSufix
	else
	'	iNumber=sPrefix&iNumber&sSufix
	    iNumber = iNumber
	end if

GetSeriesNumberLotNumber=iNumber
End function
%>
<%
Function GetSeriesNumberUsedLast(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
Set objRsSeries = Server.CreateObject("ADODB.RecordSet")

	sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	sType =objRsSeries(0)
	iLength=objRsSeries(1)

	objRsSeries.close

	sPeriod=GetPeriodInterval(sNoDate,sType)

'Response.Write sPeriod&"<BR>"

	sQuery="SELECT Number-1, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
	"Period='"&sPeriod&"'"
'Response.Write sQuery&"<BR>"
	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	if  not objRsSeries.EOF then
		iNumber=trim(objRsSeries(0))
		sPrefix=trim(objRsSeries(1))
		sSufix=trim(objRsSeries(2))
	End if
	set objRsSeries=nothing
	if CDbl(iNumber)=-1 or CDbl(iNumber)=0 then
		iNumber=0
	end if

	if cint(iLength)>0 then
		iTemp=iLength-Len(iNumber)
		for iCounter=1 to iTemp
			iNumber="0"&iNumber
		next
		iNumber=sPrefix&iNumber&sSufix
	else
		iNumber=sPrefix&iNumber&sSufix
	end if

GetSeriesNumberUsedLast=iNumber
End function
%>

<%
Function GetSeriesNumber(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
Set objRsSeries = Server.CreateObject("ADODB.RecordSet")

	sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	sType =objRsSeries(0)
	iLength=objRsSeries(1)

	objRsSeries.close

	sPeriod=GetPeriodInterval(sNoDate,sType)

'Response.Write sPeriod&"<BR>"

	sQuery="SELECT Number, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
	"Period='"&sPeriod&"'"
'Response.Write sQuery&"<BR>"
	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	if  not objRsSeries.EOF then
		iNumber=trim(objRsSeries(0))
		sPrefix=trim(objRsSeries(1))
		sSufix=trim(objRsSeries(2))
	End if
	set objRsSeries=nothing
	if CDbl(iNumber)=0 then
		iNumber=1
	end if

	if cint(iLength)>0 then
		iTemp=iLength-Len(iNumber)
		for iCounter=1 to iTemp
			iNumber="0"&iNumber
		next
		iNumber=sPrefix&iNumber&sSufix
	else
		iNumber=sPrefix&iNumber&sSufix
	end if

GetSeriesNumber=iNumber
End function
%>

<%
function GetPeriodInterval(sDate,sIntervalType)
'response.write "A " & sIntervalType
dim iMonth,iYear
'Response.Write "<p>sDate="&sDate
iMonth=cint(mid(sDate,4,2))
iYear=mid(sDate,7,4)

	select Case sIntervalType
		Case "M"
				if iMonth < 10 then
					iMonth = "0"&iMonth
				end if
		Case "Q"
				if iMonth>=4 and iMonth<=6 then
					iMonth="01"
				end if
				if iMonth>=7 and iMonth<=9 then
					iMonth="02"
				end if
				if iMonth>=10 and iMonth<=12 then
					iMonth="03"
				end if
				if iMonth>=1 and iMonth<=3 then
					iMonth="04"
				end if
		Case "Y"
'				if iMonth<=12 then
'					iYear=cint(iYear)+1
'				end if
'				iMonth="03"
' CHANGED BY SRIDHARAN FOR THE MONTH/PERIOD PROBLEM
				if iMonth >= 4 and iMonth <= 12 then
					iYear = Cint(iYear) + 1
				else
					iYear = Cint(iYear)
				end if
				iMonth="03"
	end select
				GetPeriodInterval=cstr(iYear)&cstr(iMonth)
End function
%>

<%
Function GenSeriesCode(sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen)
dim iSeriesCode
dim objRsSeries,sQuery
dim iEntryNo,sPeriod,iNumber,sPrefix,sSufix

	Set objRsSeries = Server.CreateObject("ADODB.RecordSet")
	sQuery="SELECT isnull(max(SeriesCode),0)+1 FROM APP_R_NoSeriesModules where OUDefinitionID='"&sOrgid&"' and SeriesNo="&iSeriesNo
	'Response.Write sQuery &"<br>"
	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	iSeriesCode =objRsSeries(0)
	objRsSeries.close

	sQuery="INSERT INTO APP_R_NoSeriesModules(OUDefinitionID, SeriesNo, SeriesCode, "&_
			"ModuleCode, ApplicationCode, Description, LastUsedDate,CounterType,NumberLength)"&_
			"VALUES('"&sOrgid&"',"&iSeriesNo&","&iSeriesCode&","&iModuleCode&","&iAppcode&",'"&sDescription&"',"&_
			"NULL,'"&sType&"',"&iLen&")"
	'Response.Write sQuery &"<br><br>"
	con.Execute (sQuery)

	sQuery="SELECT EntryNo, Period FROM Ms_NumberSeriesEntry where SeriesNo="&iSeriesNo
	'Response.Write sQuery
	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	do while not objRsSeries.EOF
		iEntryNo=cint(trim(objRsSeries(1)))
		sPeriod=GetPeriod(getFromFinYear(),cint(objRsSeries(1))-1,sType)
		iNumber=Request.Form("txt"&sName&"StartNo"&iEntryNo)
		sPrefix=Request.Form("txt"&sName&"Prefix"&iEntryNo)
		sSufix=Request.Form("txt"&sName&"Suffix"&iEntryNo)

		sQuery="INSERT INTO APP_R_NoSeriesModuleEntry(OUDefinitionID, SeriesNo, SeriesCode, "&_
				"EntryNo,Period, Number, Prefix, Suffix)VALUES('"&iUnitNo&"',"&iSeriesNo&","&iSeriesCode&","&_
				""&iEntryNo&",'"&sPeriod&"',"&iNumber&",'"&sPrefix&"','"&sSufix&"')"

	'	Response.Write sQuery &"<br><br>"

		con.Execute (sQuery)
		objRsSeries.MoveNext
	loop
	set objRsSeries=nothing
	GenSeriesCode=iSeriesCode
End function

Function GetPeriod(sFinacialYear,iDifference,sType)
dim iMonth,iYear
	iMonth=mid(sFinacialYear,1,2)
	iYear=mid (sFinacialYear,3,4)

	select Case sType
		Case "M"
				iMonth=cint(iMonth)+cint(iDifference)
		Case "Q"
				iMonth=cint(iMonth)+cint(iDifference)*4
		Case "Y"
				iMonth=cint(iMonth)-1
				iYear=cint(iYear)+1
	end select

	if iMonth >12 then
		iMonth=iMonth -12
		iYear=cint(iYear)+1
	end if
	if iMonth<10 then
		iMonth="0"&iMonth
	end if
	GetPeriod=cstr(iYear)&cstr(iMonth)
End function
%>



<%
Function GenSeriesNumberMaintenance(sOrgid,iSeriesNo,iSeriesCode,sNoDate,sWorkGroupType)

dim objRsSeries,sSql,rsTemp
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
	Set objRsSeries = Server.CreateObject("ADODB.RecordSet")
	set rsTemp = Server.CreateObject("ADODB.RecordSet")

	sSql = "Select WorkGroupValue from VwMtnNoSeriesSel where MainSeriesNo="&iSeriesNo&" and MainSeriesCode="&iSeriesCode&" "&_
		"	and WorkGroupValue in ('0','"&sWorkGroupType&"') "
	'Response.Write sSql
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sSql
		.Open
	end with
	set rsTemp.ActiveConnection = nothing
	if not rsTemp.EOF then
		sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

		'Response.Write " <p> squery =" & sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		if not objRsSeries.EOF then
			sType =objRsSeries(0)
			iLength=objRsSeries(1)
		end if
		objRsSeries.close

		sPeriod=GetPeriodInterval(sNoDate,sType)
		'response.write sPeriod
		sQuery="SELECT Number, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
		"Period='"&sPeriod&"'"
		'Response.Write sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		if  not objRsSeries.EOF then
			iNumber=trim(objRsSeries(0))
			sPrefix=trim(objRsSeries(1))
			sSufix=trim(objRsSeries(2))
		End if
		set objRsSeries=nothing
		if CDbl(iNumber)=0 then
			iNumber=1
		end if

		sQuery="Update APP_R_NoSeriesModuleEntry set Number="&CDbl(iNumber)+1&"  where SeriesNo="&_
			""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
			"Period='"&sPeriod&"'"
		con.execute(sQuery)
		if cint(iLength)>0 then
			iTemp=iLength-Len(iNumber)
			for iCounter=1 to iTemp
				iNumber="0"&iNumber
			next
			iNumber=sPrefix&iNumber&sSufix
		else
			iNumber=sPrefix&iNumber&sSufix
		end if

		GenSeriesNumberMaintenance=iNumber
	End if
End function
%>

<%
Function checkNoSeriesEntryAndGenerateMaintenance(sOrgID,iActivityNo,dtPassDate,sPassWorkGroupType)

	Dim sRetVal
	Dim iSeriesNo,iSeriesCode
	Dim rsNumber

	set rsNumber = Server.CreateObject("ADODB.RecordSet")
	sRetVal = ""

	'' To Generate Code (Using Number Series)'---------------------------------------------
	With rsNumber
		.CursorLocation = 3
		.CursorType = 3
		.Source =  "SELECT MainSeriesNo,MainSeriesCode FROM VwMtnNoSeriesSel where " &_
			" ActivityType=" & iActivityNo & " AND OrganisationCode='" & trim(sOrgID) & "' "&_
			" and WorkGroupValue in ('0','"&sPassWorkGroupType&"') "
		.ActiveConnection = con
		.Open
	End With
	'Response.Write "<p> " & rsNumber.Source
	Set rsNumber.ActiveConnection = nothing

	If 	not rsNumber.EOF then
		iSeriesNo = rsNumber(0)
		iSeriesCode = rsNumber(1)

		sRetVal = GenSeriesNumberMaintenance(sOrgID,iSeriesNo,iSeriesCode,dtPassDate,sPassWorkGroupType)
	End If
	rsNumber.Close



	checkNoSeriesEntryAndGenerateMaintenance = sRetVal
End Function
%>

<%
Function GetPreviousSeriesNumberItemWise(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries,sSql,rsTemp
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
	Set objRsSeries = Server.CreateObject("ADODB.RecordSet")
	set rsTemp = Server.CreateObject("ADODB.RecordSet")

	sSql = "Select ItemValue from VwPurNoSeriesSel where MainSeriesNo="&iSeriesNo&" and MainSeriesCode="&iSeriesCode
	'Response.Write sSql
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sSql
		.Open
	end with
	set rsTemp.ActiveConnection = nothing
	if not rsTemp.EOF then
		sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

		'Response.Write " <p> squery =" & sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		if not objRsSeries.EOF then
			sType =objRsSeries(0)
			iLength=objRsSeries(1)
		end if
		objRsSeries.close

		sPeriod=GetPeriodInterval(sNoDate,sType)
		'response.write sPeriod
		sQuery="SELECT Number-1, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
		"Period='"&sPeriod&"'"
		'Response.Write sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		if  not objRsSeries.EOF then
			iNumber=trim(objRsSeries(0))
			sPrefix=trim(objRsSeries(1))
			sSufix=trim(objRsSeries(2))
		End if
		set objRsSeries=nothing
		if CDbl(iNumber)=0 then
			iNumber=1
		end if

		if cint(iLength)>0 then
			iTemp=iLength-Len(iNumber)
			for iCounter=1 to iTemp
				iNumber="0"&iNumber
			next
			iNumber=sPrefix&iNumber&sSufix
		else
			iNumber=sPrefix&iNumber&sSufix
		end if

		GetPreviousSeriesNumberItemWise=iNumber
	End if
End function
%>
<%
Function UpdatePreviousSeriesNumberItemWise(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries,sSql,rsTemp
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
	Set objRsSeries = Server.CreateObject("ADODB.RecordSet")
	set rsTemp = Server.CreateObject("ADODB.RecordSet")

	sSql = "Select ItemValue from VwPurNoSeriesSel where MainSeriesNo="&iSeriesNo&" and MainSeriesCode="&iSeriesCode
	'Response.Write sSql
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sSql
		.Open
	end with
	set rsTemp.ActiveConnection = nothing
	if not rsTemp.EOF then
		sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

		'Response.Write " <p> squery =" & sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		if not objRsSeries.EOF then
			sType =objRsSeries(0)
			iLength=objRsSeries(1)
		end if
		objRsSeries.close

		sPeriod=GetPeriodInterval(sNoDate,sType)
		'response.write sPeriod
		sQuery="SELECT Number-1, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
		"Period='"&sPeriod&"'"
		'Response.Write sQuery
		with objRsSeries
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		if  not objRsSeries.EOF then
			iNumber=trim(objRsSeries(0))
			sPrefix=trim(objRsSeries(1))
			sSufix=trim(objRsSeries(2))
		End if
		set objRsSeries=nothing
		if CDbl(iNumber)=0 then
			iNumber=1
		end if

		sQuery="Update APP_R_NoSeriesModuleEntry set Number="&CDbl(iNumber)&"  where SeriesNo="&_
			""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
			"Period='"&sPeriod&"'"
		con.execute(sQuery)
		if cint(iLength)>0 then
			iTemp=iLength-Len(iNumber)
			for iCounter=1 to iTemp
				iNumber="0"&iNumber
			next
			iNumber=sPrefix&iNumber&sSufix
		else
			iNumber=sPrefix&iNumber&sSufix
		end if

		UpdatePreviousSeriesNumberItemWise=iNumber
	End if
End function
%>
<%
Function GetPreviousSeriesNumber(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
Set objRsSeries = Server.CreateObject("ADODB.RecordSet")

	sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	if not objRsSeries.eof then
	sType =objRsSeries(0)
	iLength=objRsSeries(1)
	end if
	objRsSeries.close

	sPeriod=GetPeriodInterval(sNoDate,sType)

	'Response.Write "<p> sPeriod ="  & sPeriod & "<BR>"

	sQuery="SELECT Number-1, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
	"Period='"&sPeriod&"'"

	'Response.Write sQuery&"<BR>"

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	if  not objRsSeries.EOF then
		iNumber=trim(objRsSeries(0))
		sPrefix=trim(objRsSeries(1))
		sSufix=trim(objRsSeries(2))
	End if
	set objRsSeries=nothing
	if CDbl(iNumber)=0 then
		iNumber=1
	end if

	if cint(iLength)>0 then
		iTemp=iLength-Len(iNumber)
		for iCounter=1 to iTemp
			iNumber="0"&iNumber
		next
		iNumber=sPrefix&iNumber&sSufix
	else
		iNumber=sPrefix&iNumber&sSufix
	end if

GetPreviousSeriesNumber=iNumber
End function
%>

<%
Function UpdatePreviousSeriesNumber(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
Set objRsSeries = Server.CreateObject("ADODB.RecordSet")

	sQuery="select CounterType,NumberLength from APP_R_NoSeriesModules where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	if not objRsSeries.eof then
	sType =objRsSeries(0)
	iLength=objRsSeries(1)
	end if
	objRsSeries.close

	sPeriod=GetPeriodInterval(sNoDate,sType)

	'Response.Write "<p> sPeriod ="  & sPeriod & "<BR>"

	sQuery="SELECT Number-1, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
	""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
	"Period='"&sPeriod&"'"

	'Response.Write sQuery&"<BR>"

	with objRsSeries
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	if  not objRsSeries.EOF then
		iNumber=trim(objRsSeries(0))
		sPrefix=trim(objRsSeries(1))
		sSufix=trim(objRsSeries(2))
	End if
	set objRsSeries=nothing
	if CDbl(iNumber)=0 then
		iNumber=0
	end if

	sQuery="Update APP_R_NoSeriesModuleEntry set Number="&CDbl(iNumber)&"  where SeriesNo="&_
			""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
			"Period='"&sPeriod&"'"
		con.execute(sQuery)

	if cint(iLength)>0 then
		iTemp=iLength-Len(iNumber)
		for iCounter=1 to iTemp
			iNumber="0"&iNumber
		next
		iNumber=sPrefix&iNumber&sSufix
	else
		iNumber=sPrefix&iNumber&sSufix
	end if
UpdatePreviousSeriesNumber=iNumber
End function
%>