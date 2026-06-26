<%
Function GetInvNumberSeriesCodes(sActivityType,sOrgCode,iNumClassCode)
Dim rsObj,rsTemp
Dim sQuery,sCatCode
Dim iSeriesCode,iSeriesNo

if iNumClassCode = "0" or isnull(iNumClassCode) or iNumClassCode = "" then iNumClassCode = 0

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")

sQuery = "Select GroupCategory from INV_M_Classification where ParentGroup in("
sQuery = sQuery & " Select ParentGroup from INV_M_Classification where GroupCode = "& iNumClassCode &") and GroupCategory is not null"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    sCatCode = rsTemp(0)
else
	sCatCode= 0
end if
rsTemp.Close


sQuery = "SELECT SERIESNO,SERIESCODE FROM VwInvNumSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
if Trim(iNumClassCode)<>"0" then
	sQuery = sQuery  &" and ClassCode in ("& iNumClassCode &") and CatCode in ("& sCatCode &")"
end if
sQuery = sQuery &  " Order by SERIESCODE desc "
	with rsObj
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	'Response.Write "<p> No series = "&rsObj.source
  '
	set rsObj.ActiveConnection = nothing
   if not rsObj.EOF then
		iSeriesNo = trim(rsObj(0))
		iSeriesCode = trim(rsObj(1))
	else
		iSeriesCode = "0"
		iSeriesNo = "0"
	end if
   rsObj.close

   if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
			sQuery = "SELECT SERIESNO,SERIESCODE FROM VwInvNumSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
			if Trim(iNumClassCode)<>"0" then
				sQuery = sQuery  &" and ClassCode in (Select ParentGroup from INV_M_Classification where GroupCode in ("& iNumClassCode &")) and CatCode in ("& sCatCode &")"
			end if
			sQuery = sQuery &  " Order by SERIESCODE desc "
				with rsObj
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				''Response.Write "<p> No series = "&rsObj.source
			  '
				set rsObj.ActiveConnection = nothing
			   if not rsObj.EOF then
					iSeriesNo = trim(rsObj(0))
					iSeriesCode = trim(rsObj(1))
				else
					iSeriesCode = "0"
					iSeriesNo = "0"
				end if
			   rsObj.close
   end if

   if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
		sQuery = "SELECT SERIESNO,SERIESCODE FROM VwInvNumSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
		sQuery = sQuery & " and ClassCode is null and CatCode in ("& sCatCode &")"
		sQuery = sQuery &  " Order by SERIESCODE desc "

	   ' Response.Write "<p>"& sQuery
			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set rsObj.ActiveConnection = nothing
		   if not rsObj.EOF then
				iSeriesNo = trim(rsObj(0))
				iSeriesCode = trim(rsObj(1))
			else
				iSeriesCode = "0"
				iSeriesNo = "0"
		   end if
		   rsObj.close
   end if 'if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then



   if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
		sQuery = "SELECT SERIESNO,SERIESCODE FROM VwInvNumSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
		sQuery = sQuery & " and ClassCode is null and CatCode is null"
		sQuery = sQuery &  " Order by SERIESCODE desc "
			with rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set rsObj.ActiveConnection = nothing
		   if not rsObj.EOF then
				iSeriesNo = trim(rsObj(0))
				iSeriesCode = trim(rsObj(1))
			else
				iSeriesCode = "0"
				iSeriesNo = "0"
		   end if
		   rsObj.close
   end if 'if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then

	GetInvNumberSeriesCodes = iSeriesNo&":"&iSeriesCode

End Function
%>

<%
Function GetPRDNumberSeriesCodes(sOrgCode,iNumClassCode)
Dim rsObj,rsTemp
Dim sQuery
Dim iSeriesCode,iSeriesNo,sCatCode

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")

if iNumClassCode = "0" or isnull(iNumClassCode) or iNumClassCode = "" then iNumClassCode = 0


sQuery = "Select GroupCategory from INV_M_Classification where ParentGroup in("
sQuery = sQuery & " Select ParentGroup from INV_M_Classification where GroupCode = "& iNumClassCode &") and GroupCategory is not null"
'Response.write "<p>"& sQuery
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    sCatCode = rsTemp(0)
else
	sCatCode= 0
end if
rsTemp.Close


sQuery = "SELECT SERIESNO,SERIESCODE FROM VwPackNoSeries WHERE ORGANISATIONCODE = " & Pack(sOrgCode)
			    if Trim(iNumClassCode)<>"0" then
			        sQuery = sQuery  &" and ClassCode in ("& iNumClassCode &") and CatCode in ("& sCatCode &")"
			    end if
			    sQuery = sQuery &  " Order by SERIESCODE "
			       'Response.Write "<p> No series = "&rsObj.source
			      ' Response.write "<p>"& sQuery
				    with rsObj
					    .CursorLocation = 3
					    .CursorType = 3
					    .Source = sQuery
					    .ActiveConnection = con
					    .Open
				    end with

				  '
				    set rsObj.ActiveConnection = nothing
				   if not rsObj.EOF then
					    iSeriesNo = trim(rsObj(0))
					    iSeriesCode = trim(rsObj(1))
				    else
				        iSeriesCode = "0"
				        iSeriesNo = "0"
				    end if
				   rsObj.close

				   if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
				            sQuery = "SELECT SERIESNO,SERIESCODE FROM VwPackNoSeries WHERE ORGANISATIONCODE = " & Pack(sOrgCode)
			                if Trim(iNumClassCode)<>"0" then
			                    sQuery = sQuery  &" and ClassCode in (Select ParentGroup from INV_M_Classification where GroupCode in ("& iNumClassCode &")) and CatCode in ("& sCatCode &")"
			                end if
			                sQuery = sQuery &  " Order by SERIESCODE desc "
			            '    Response.write "<p>"& sQuery
				                with rsObj
					                .CursorLocation = 3
					                .CursorType = 3
					                .Source = sQuery
					                .ActiveConnection = con
					                .Open
				                end with
				                ''Response.Write "<p> No series = "&rsObj.source
				              '
				                set rsObj.ActiveConnection = nothing
				               if not rsObj.EOF then
					                iSeriesNo = trim(rsObj(0))
					                iSeriesCode = trim(rsObj(1))
				                else
				                    iSeriesCode = "0"
				                    iSeriesNo = "0"
				                end if
				               rsObj.close
				   end if


				    if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
				        sQuery = "SELECT SERIESNO,SERIESCODE FROM VwPackNoSeries WHERE  ORGANISATIONCODE = " & Pack(sOrgCode)
				        sQuery = sQuery & " and ClassCode is null and CatCode in ("& sCatCode &")"
			            sQuery = sQuery &  " Order by SERIESCODE desc "
			           ' Response.write "<p>"& sQuery
				            with rsObj
					            .CursorLocation = 3
					            .CursorType = 3
					            .Source = sQuery
					            .ActiveConnection = con
					            .Open
				            end with
				            set rsObj.ActiveConnection = nothing
				           if not rsObj.EOF then
					            iSeriesNo = trim(rsObj(0))
					            iSeriesCode = trim(rsObj(1))
					        else
					            iSeriesCode = "0"
					            iSeriesNo = "0"
				           end if
				           rsObj.close
				   end if 'if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then


				   if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
				        sQuery = "SELECT SERIESNO,SERIESCODE FROM VwPackNoSeries WHERE  ORGANISATIONCODE = " & Pack(sOrgCode)
				        sQuery = sQuery & " and ClassCode is null and CatCode is null"
			            sQuery = sQuery &  " Order by SERIESCODE desc "
			            'Response.write "<p>"& sQuery
				            with rsObj
					            .CursorLocation = 3
					            .CursorType = 3
					            .Source = sQuery
					            .ActiveConnection = con
					            .Open
				            end with
				            set rsObj.ActiveConnection = nothing
				           if not rsObj.EOF then
					            iSeriesNo = trim(rsObj(0))
					            iSeriesCode = trim(rsObj(1))
					        else
					            iSeriesCode = "0"
					            iSeriesNo = "0"
				           end if
				           rsObj.close
				   end if 'if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then

	GetPRDNumberSeriesCodes = iSeriesNo&":"&iSeriesCode

End Function
%>

<%
Function GetSalNumberSeriesCodes(sActivityType,sOrgCode,iNumClassCode)
Dim rsObj,rsTemp
Dim sQuery
Dim iSeriesCode,iSeriesNo,sCatCode

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")

if iNumClassCode = "0" or isnull(iNumClassCode) or iNumClassCode = "" then iNumClassCode = 0

sQuery = "Select GroupCategory from INV_M_Classification where ParentGroup in("
sQuery = sQuery & " Select ParentGroup from INV_M_Classification where GroupCode = "& iNumClassCode &") and GroupCategory is not null"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    sCatCode = rsTemp(0)
else
	sCatCode= 0
end if
rsTemp.Close


sQuery = "SELECT SERIESNO,SERIESCODE FROM VWSalNoSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
			    if Trim(iNumClassCode)<>"0" then
			        sQuery = sQuery  &" and ClassCode in ("& iNumClassCode &") and CatCode in ("& sCatCode &")"
			    end if
			    sQuery = sQuery &  " Order by SERIESCODE desc "
				    with rsObj
					    .CursorLocation = 3
					    .CursorType = 3
					    .Source = sQuery
					    .ActiveConnection = con
					    .Open
				    end with
				    ''Response.Write "<p> No series = "&rsObj.source
				  '
				    set rsObj.ActiveConnection = nothing
				   if not rsObj.EOF then
					    iSeriesNo = trim(rsObj(0))
					    iSeriesCode = trim(rsObj(1))
				    else
				        iSeriesCode = "0"
				        iSeriesNo = "0"
				    end if
				   rsObj.close

				   if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
				            sQuery = "SELECT SERIESNO,SERIESCODE FROM VWSalNoSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
			                if Trim(iNumClassCode)<>"0" then
			                    sQuery = sQuery  &" and ClassCode in (Select ParentGroup from INV_M_Classification where GroupCode in ("& iNumClassCode &")) and CatCode in ("& sCatCode &")"
			                end if
			                sQuery = sQuery &  " Order by SERIESCODE desc "
				                with rsObj
					                .CursorLocation = 3
					                .CursorType = 3
					                .Source = sQuery
					                .ActiveConnection = con
					                .Open
				                end with
				                ''Response.Write "<p> No series = "&rsObj.source
				              '
				                set rsObj.ActiveConnection = nothing
				               if not rsObj.EOF then
					                iSeriesNo = trim(rsObj(0))
					                iSeriesCode = trim(rsObj(1))
				                else
				                    iSeriesCode = "0"
				                    iSeriesNo = "0"
				                end if
				               rsObj.close
				   end if


				    if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
				        sQuery = "SELECT SERIESNO,SERIESCODE FROM VWSalNoSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
				        sQuery = sQuery & " and ClassCode is null and CatCode in ("& sCatCode &")"
			            sQuery = sQuery &  " Order by SERIESCODE desc "
				            with rsObj
					            .CursorLocation = 3
					            .CursorType = 3
					            .Source = sQuery
					            .ActiveConnection = con
					            .Open
				            end with
				            set rsObj.ActiveConnection = nothing
				           if not rsObj.EOF then
					            iSeriesNo = trim(rsObj(0))
					            iSeriesCode = trim(rsObj(1))
					        else
					            iSeriesCode = "0"
					            iSeriesNo = "0"
				           end if
				           rsObj.close
				   end if 'if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then


				   if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then
				        sQuery = "SELECT SERIESNO,SERIESCODE FROM VWSalNoSeries WHERE ACTIVITYTYPE = '"& sActivityType &"' AND ORGANISATIONCODE = " & Pack(sOrgCode)
				        sQuery = sQuery & " and ClassCode is null and CatCode is null"
			            sQuery = sQuery &  " Order by SERIESCODE desc "
				            with rsObj
					            .CursorLocation = 3
					            .CursorType = 3
					            .Source = sQuery
					            .ActiveConnection = con
					            .Open
				            end with
				            set rsObj.ActiveConnection = nothing
				           if not rsObj.EOF then
					            iSeriesNo = trim(rsObj(0))
					            iSeriesCode = trim(rsObj(1))
					        else
					            iSeriesCode = "0"
					            iSeriesNo = "0"
				           end if
				           rsObj.close
				   end if 'if Trim(iSeriesNo)="0" and Trim(iSeriesCode)="0" then

	GetSalNumberSeriesCodes = iSeriesNo&":"&iSeriesCode

End Function
%>

<%

Function GetPurNumberSeriesCodes(sActivityNo,sOrgCode,iNoSerClassCode)

Dim rsObj,rsTemp
Dim sQuery
Dim iSeriesCode,iSeriesNo,sCatCode

set rsObj = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")

Response.Write "<p>iNoSerClassCode = "& iNoSerClassCode

if iNoSerClassCode = "0" or isnull(iNoSerClassCode) or iNoSerClassCode = "" then iNoSerClassCode = 0

sQuery = "Select GroupCategory from INV_M_Classification where ParentGroup in("
sQuery = sQuery & " Select ParentGroup from INV_M_Classification where GroupCode = "& iNoSerClassCode &") and GroupCategory is not null"
Response.Write "<p>"& sQuery
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    sCatCode = rsTemp(0)
else
	sCatCode= 0
end if
rsTemp.Close


			sQuery = "SELECT MainSeriesNo,MainSeriesCode FROM VwPurNoSeriesSel where " &_
						" ActivityType=" & sActivityNo & " AND OrganisationCode='" & trim(sOrgCode) & "' "

			if Trim(iNoSerClassCode)<>"0" then
			    sQuery  = sQuery & " and ClassCode in ("& iNoSerClassCode  &") and CatCode in ("& sCatCode &")"
			end if
			sQuery= sQuery &" Order by MainSeriesCode desc"
			Response.Write "<p>"& sQuery
			With rsObj
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			End With
			Set rsObj.ActiveConnection = nothing
			If 	not rsObj.EOF then
				iSeriesNo = rsObj(0)
				iSeriesCode = rsObj(1)
			else
			    iSeriesCode = 0
			    iSeriesNo = 0
			End If
			rsObj.Close

			if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
			    sQuery = "SELECT MainSeriesNo,MainSeriesCode FROM VwPurNoSeriesSel where " &_
						" ActivityType=" & sActivityNo & " AND OrganisationCode='" & trim(sOrgCode) & "' "

						if Trim(iNoSerClassCode)<>"0" then
							sQuery = sQuery & " and ClassCode in (Select ParentGroup from INV_M_Classification where GroupCode in ("& iNoSerClassCode &")) and CatCode in ("& sCatCode &")"
						end if
						sQuery= sQuery &" Order by MainSeriesCode desc"
            Response.Write "<p>"& sQuery
			    With rsObj
				    .CursorLocation = 3
				    .CursorType = 3
				    .Source = sQuery
				    .ActiveConnection = con
				    .Open
			    End With
			    Set rsObj.ActiveConnection = nothing
			    If 	not rsObj.EOF then
				    iSeriesNo = rsObj(0)
				    iSeriesCode = rsObj(1)
			    else
			        iSeriesCode = 0
			        iSeriesNo = 0
			    End If
			    rsObj.Close
			end if


			if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
			    sQuery= "SELECT MainSeriesNo,MainSeriesCode FROM VwPurNoSeriesSel where " &_
						" ActivityType=" & sActivityNo & " AND OrganisationCode='" & trim(sOrgCode) & "' "&_
						" and ClassCode is Null and CatCode in ("& sCatCode &")"
						sQuery= sQuery &" Order by MainSeriesCode desc"
                Response.Write "<p>"& sQuery
			    With rsObj
				    .CursorLocation = 3
				    .CursorType = 3
				    .Source = sQuery
				    .ActiveConnection = con
				    .Open
			    End With
			    Set rsObj.ActiveConnection = nothing
			    If 	not rsObj.EOF then
				    iSeriesNo = rsObj(0)
				    iSeriesCode = rsObj(1)
			    else
			        iSeriesCode = 0
			        iSeriesNo = 0
			    End If
			    rsObj.Close
			end if


			if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
			    sQuery= "SELECT MainSeriesNo,MainSeriesCode FROM VwPurNoSeriesSel where " &_
						" ActivityType=" & sActivityNo & " AND OrganisationCode='" & trim(sOrgCode) & "' "&_
						" and ClassCode is Null and CatCode is null "
						sQuery= sQuery &" Order by MainSeriesCode desc"
                Response.Write "<p>"& sQuery
			    With rsObj
				    .CursorLocation = 3
				    .CursorType = 3
				    .Source = sQuery
				    .ActiveConnection = con
				    .Open
			    End With
			    Set rsObj.ActiveConnection = nothing
			    If 	not rsObj.EOF then
				    iSeriesNo = rsObj(0)
				    iSeriesCode = rsObj(1)
			    else
			        iSeriesCode = 0
			        iSeriesNo = 0
			    End If
			    rsObj.Close
			end if

			GetPurNumberSeriesCodes = iSeriesNo &":"& iSeriesCode

End Function

%>

<%
Function CheckNoSerAvilForThisYear(sOrgid,iSeriesNo,iSeriesCode,sNoDate)

dim objRsSeries,sSql,rsTemp
dim iLength,sType,iCounter,iTemp
dim sPeriod,iNumber,sPrefix,sSufix,sQuery
dim bEligible

Set objRsSeries = Server.CreateObject("ADODB.RecordSet")
set rsTemp = Server.CreateObject("ADODB.RecordSet")

bEligible = false

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
		if trim(sType)<>"" then
		    sPeriod=GetPeriodInterval(sNoDate,sType)
		end if

		sQuery="SELECT Number, Prefix, Suffix FROM APP_R_NoSeriesModuleEntry where SeriesNo="&_
		""&iSeriesNo&" and OUDefinitionID='"&sOrgid&"' and SeriesCode="&iSeriesCode&" and "&_
		"Period='"&sPeriod&"'"
		objRsSeries.open sQuery,con
		if not objRsSeries.eof then
		    bEligible = true
		else
		    bEligible = false
		end if

		CheckNoSerAvilForThisYear=bEligible

End Function

%>