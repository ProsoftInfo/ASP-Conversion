
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>

<%
	'Program Name				:	View_Schedule.asp
	'Module Name				:	Accounts (Masters- Configure Balance Sheet - Schedule Setup)
	'Author Name				:	KalaiSelvi R
	'Created On					:	Jan 19,2012
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/PurPopulate.asp" -->
<%
Const iPageSize = 7

Dim sQuery,Objrs1,Objrs2,Objrs3,rsSch,iCtr,sOrgID,sCatCode,sFinYr,iAccHead
Dim dClosing,iSchID,iSchSubID,iSchSubSubID,sLastDay,dPrvVal,sSchName
Dim iShId,iShSub,iShSubSubId,sForTheDate,sPreDate,dPreYrVal
Dim sAccDesc,sAccCheck,iEntNo,dHeadVal,sCallFrom
Dim nTempCtr,dSchAddTotal,dSchSubTotal
Dim nTot1,nTot2


Dim iCurrentPage,iTotPage,iPageCtr,lnPage,CurrPage
Dim iLineCtr,nMaxLinesPrint

nMaxLinesPrint = 20 '67

Set rsSch = Server.CreateObject("ADODB.RecordSet")
Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
Set Objrs2 = Server.CreateObject("ADODB.RecordSet")
Set Objrs3 = Server.CreateObject("ADODB.RecordSet")
	

sOrgID = trim(request("OrgID"))
sForTheDate = Trim(Request("ForTheDate"))
sCallFrom = trim(Request("CallFrom"))
sFinYr = Session("FinPeriod")

IF CStr(sOrgID) = "" Then
	sOrgID = "0"
End IF

sLastDay = sForTheDate

IF CStr(sLastDay) = "" Then
	sLastDay = "30/04/"&Trim(Left(sFinYr,4))
End IF

sPreDate = "31/03/"&Left(sFinYr,4)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - View Schedule</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 >

<form method="POST" name="formname" action="">

<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<input type="hidden" name="hPageSelection" value="<%=iCurrentPage%>">


<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" bordercolor="#000000">
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
                                <td class="PageTitle" align="center" colspan="3">
                                    <%=getUnitName(sOrgID)%>
                                </td>
                            </tr>
                            <tr>
		                        <td align="center" class=PageTitle colspan="3" height="20">
		                        </td>
                            </tr>
                            
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td>
                                <div class="frmBody" id="frm2" style="width: 750; height:430;">
							           <table border="0" cellspacing="1" class="ExcelTable" width="730">

<%


sQuery = "Select Distinct ScheduleID,Cast(ScheduleNumber as Char) ScheduleNo, "&_
		 "ScheduleHeading,Hierarchy From Vw_Acc_SchSetup Where ApplicableFor='" & sCallFrom & "' "&_
		 " and OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' Order By Hierarchy "
With rsSch
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End WIth
Set rsSch.ActiveConnection = Nothing

set sCatCode = rsSch(0)

Do While Not rsSch.EOF

	nTot1 = 0
	nTot2 = 0
	
	%>
	<tr>
		<td align="left" class="ExcelDisplayCell" colspan="6"><b><%=rsSch("ScheduleNo")%>.<%=rsSch("ScheduleHeading")%></b></td>
	</tr>
	<tr>
		<td align="center" class="ExcelHeaderCell"><p>Heading Name</td>
		<td align="center" class="ExcelHeaderCell">As on <%=Replace(sForTheDate,"/",".")%></td>
		<td align="center" class="ExcelHeaderCell">As on <%=Replace(sPreDate,"/",".")%></td>
	</tr>
	<%	
	
	'Taking The SchSubID and Sch SubSUbID,Name and EntryType for the Selected Schedule
	sQuery = "Select Distinct ScheduleSubID,ScheduleSubSubID,isNull(SubHeadingName,'')  "&_
			 "SubHeadingName,EntryType,DisplayACHeadDescr,Hierachy From Vw_Acc_SchSetup "&_
			 "Where ScheduleID = "&sCatCode&" and ScheduleSubSubID = 0 and "&_
			 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' Order By Hierachy "
	'Response.Write sQuery &"<BR><BR>"
	With Objrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End With
	Set Objrs1.ActiveConnection = Nothing
	nTempCtr = 0
	Do WHile Not Objrs1.EOF
	
		nTempCtr = nTempCtr + 1
		iSchID = Objrs1("ScheduleSubID")

			'From The Particular SchSub ID Take the Values SchSubSub ID and Name and EntryType
			sQuery = "Select Distinct ScheduleSubSubID,isNull(SubHeadingName,'') SubHeadingName,EntryType,Hierarchy From Vw_Acc_SchSetup  "&_
					 "Where ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&" and ScheduleSubSubID <> 0 and "&_
					 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' Order By Hierarchy "
			'Response.Write sQuery &"<BR><BR>"
			With Objrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With

			Set Objrs3.ActiveConnection = Nothing
			IF Not Objrs3.EOF Then 'IF There is SubSubIt For the Each SchSub ID then Proceed

				dPreYrVal = 0
				sQuery = "Select ScheduleSubHeadValue  "&_
						 "From Acc_T_ScheduleACDetail Where ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&"  "&_
						 "and ScheduleSubSubID = 0 and Convert(Varchar,AsOnDate,103) = '"&sPreDate&"' "
				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dPreYrVal = Objrs2(0)
				End IF
				Objrs2.Close

				IF CStr(Objrs1("EntryType")) = "N" Then

	%>
					<tr>
						<td align="left" class="ExcelDisplayCell" colspan="6"><b>
	<%						Response.Write(Objrs1("SubHeadingName"))%></b>
						</td>
					</tr>
	<%			ElseIF CStr(Objrs1("EntryType")) = "D" Then %>
					<tr>
						<td align="left" class="ExcelDisplayCell"><b>
	<%						Response.Write(Objrs1("SubHeadingName"))%></b>
						</td>
	<%
					sQuery = "Select isNull(ApplicableACHeadCode,0),ScheduleSubHeadValue,EntryNumber  "&_
							 "From Acc_T_ScheduleACDetail Where ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&"  "&_
							 "and ScheduleSubSubID = 0 and Convert(Varchar,AsOnDate,103) = '"&sForTheDate&"' "


					'Response.Write sQuery &"<br><br>"
					Objrs2.Open sQuery,Con
					iSchSubSubID = 0
					iAccHead = 0
					IF Not Objrs2.EOF Then
						iAccHead = Objrs2(0)
						dPrvVal = Objrs2(1)
						iEntNo = Objrs2(2)
					Else
						iAccHead = 0
						dPrvVal = 0
						iEntNo = 0
					End IF
					Objrs2.Close
					dHeadVal = dPrvVal



					nTot1 = CDbl(nTot1) + CDbl(dHeadVal)
					nTot2 = CDbl(nTot2) + CDbl(dPreYrVal)
	%>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dHeadVal,2,,,0)%></td>
					
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
					
				</tr>
	<%			ElseIF CStr(Objrs1("EntryType")) = "A" Then %>
					<tr>
						<td align="left" class="ExcelDisplayCell"><b>
	<%						Response.Write(Objrs1("SubHeadingName"))%></b>
						</td>
	<%
					sQuery = "Select isNull(ApplicableACHeadCode,0),ScheduleSubHeadValue,EntryNumber From Acc_T_ScheduleACDetail  "&_
							 "Where ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&" and  "&_
							 "ScheduleSubSubID = 0 and Convert(Varchar,AsOnDate,103) = '"&sForTheDate&"' "


					'Response.Write sQuery &"<br><br>"
					Objrs2.Open sQuery,Con
					iSchSubSubID = 0
					iAccHead = 0
					IF Not Objrs2.EOF Then
						iAccHead = Objrs2(0)
						dPrvVal = Objrs2(1)
						iEntNo = Objrs2(2)
					Else
						iAccHead = 0
						dPrvVal = 0
						iEntNo = 0
					End IF
					Objrs2.Close
					IF CStr(iAccHead) <> "0" Then
						'Response.Write sLastDay
						dHeadVal = GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay)
						sAccDesc = GetAccHeadName(iAccHead)
					End IF
					
					nTot1 = CDbl(nTot1) + CDbl(dHeadVal)
					nTot2 = CDbl(nTot2) + CDbl(dPreYrVal)
					
	%>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dHeadVal,2,,,0)%></td>
					
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
				</tr>
	<%			End IF
			iAccHead = 0
		 	Do WHile Not Objrs3.EOF
				iSchSubSubID  = Objrs3("ScheduleSubSubID")
				'Response.Write "subid="& iSchSubSubID &"<BR><BR>"
				iAccHead = 0
				'Taking The Previous Value For the SchSub and SchSubSubID
				sQuery = "Select ScheduleSubHeadValue,isNull(ApplicableACHeadCode,0) From Acc_T_ScheduleACDetail Where "&_
						 "ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&" and ScheduleSubSubID = "&iSchSubSubID&" and  "&_
						 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' "&_
						 "and Convert(Varchar,AsOnDate,103) = '"&sForTheDate&"' "

				'Response.Write sQuery
				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dPrvVal = Objrs2(0)
					iAccHead = Objrs2(1)
				Else
					dPrvVal = 0
					iAccHead = 0
				End IF
				Objrs2.Close
				IF CStr(iAccHead) <> "0" Then
								'Response.Write sLastDay
					dHeadVal = GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay)
				End IF

				'Checking That Entry Type is of "D" Then it is "Data Entry"
				'S Means The Values has to Take from Breakup Table
				'A Accounthead Take the ACcountHead From Acc_T_ScheduleACDetail for SchID,SchSubID and SchSubSubID

				%>

				<tr>
					<td class="ExcelDisplayCell">
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<%  Response.Write(Objrs3("SubHeadingName"))
						sQuery = "Select ScheduleSubHeadValue  "&_
								 "From Acc_T_ScheduleACDetail Where ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&"  "&_
								 "and ScheduleSubSubID = "&iSchSubSubID&" and Convert(Varchar,AsOnDate,103) = '"&sPreDate&"' "


						'Response.Write sQuery &"<br><br>"
						Objrs2.Open sQuery,Con
						IF Not Objrs2.EOF Then
							dPreYrVal = Objrs2(0)
						Else
							dPreYrVal = 0
						End IF
						Objrs2.Close
						
						
						nTot2 = CDbl(nTot2) + CDbl(dPreYrVal)
					

					%></td>
					<% IF CStr(Objrs3("EntryType")) = "D" Then 'IF it is an Data Entry the User Will enter the Value 
						nTot1 = CDbl(nTot1) + CDbl(dPrvVal)
					%>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
					<% ElseIF CStr(Objrs3("EntryType")) = "S" Then 'If is is and Scheduled Type

						sQuery = "Select Sum(T.ScheduleSubHeadValue),V.ScheduleID,V.ScheduleSubID,V.ScheduleSubSubID From Acc_T_SchdBreakupAcDetail T,Vw_Acc_SchBreakSetup V  "&_
								 "Where V.ScheduleID = "&sCatCode&" and V.ScheduleSubID = "&iSchID&" and V.ScheduleSubSubID = "&iSchSubSubID&" and  "&_
								 "V.BreakupID = T.BreakUpID and V.BreakupSubID = T.BreakupSubID and  "&_
								 "V.BreakupSubSubID = T.BreakupSubSubID and V.FinYear = '"&sFinYr&"' and  "&_
								 "T.OrganisationCode = '"&sOrgID&"' and V.FinYear = T.FinYear  "&_
								 "and Convert(Varchar,T.AsOnDate,103)= '"&sForTheDate&"'  "&_
								 "Group by V.ScheduleID,V.ScheduleSubID,V.ScheduleSubSubID "
						'Response.Write sQuery & "<BR><BR>"
						Objrs2.Open sQuery,Con
						IF Not Objrs2.EOF Then
							iShId = Objrs2(1)
							iShSubSubId =Objrs2(3)
							
							nTot1 = CDbl(nTot1) + CDbl(Objrs2(0))
					%>
						 	<td class="ExcelDisplayCell" align="right"><%=FormatNumber(Objrs2(0),2,,,0)%></td>
					<% 	Else %>
							<td class="ExcelDisplayCell" align="right">0.00</td>
					<% 	End IF
						Objrs2.Close
					  ElseIF CStr(Objrs3("EntryType")) = "A" Then 'If it is an AccountHead

						sQuery = "Select isNull(ApplicableAcHeadCode,0),ScheduleSubHeadValue,EntryNumber From Acc_T_ScheduleACDetail Where ScheduleID = "&sCatCode&" "&_
								 "and ScheduleSubID = "&iSchID&" and ScheduleSubSubID = "&iSchSubSubID&" and "&_
								 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' and Convert(Varchar,AsOnDate,103) = '"&sForTheDate&"' "
						'Response.Write sQuery
						Objrs2.Open sQuery,Con
						IF Not Objrs2.EOF Then
							iAccHead = Objrs2(0)
							dPrvVal = Objrs2(1)
							iEntNo = Objrs2(2)
						Else
							iAccHead = 0
							dPrvVal = 0
							iEntNo = 0
						End IF
						Objrs2.Close

						IF CStr(iAccHead) <> "0" Then
							sAccDesc = GetAccHeadName(iAccHead)
						Else
							sAccDesc = ""
						End IF

						IF CStr(iAccHead) <> "0" Then
							'Response.Write sLastDay
							dClosing = GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay)
						Else
							dClosing = 0
						End IF

						nTot1 = CDbl(nTot1) + CDbl(dClosing)


					%>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dClosing,2,,,0)%></td>
					<%ElseIF CStr(Objrs3("EntryType")) = "N" Then 'If it is an AccountHead%>
						<td class="ExcelDisplayCell" align="right">&nbsp;</td>
					<%ENd IF
					
					
					
					  IF CStr(Objrs3("EntryType")) = "N" Then
					%>
						<td class="ExcelDisplayCell" align="right" >&nbsp;</td>
					<%Else%>
						
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
						
					<%End If %>

					
				</tr>
<% 				Objrs3.MoveNext
			loop
 %>

<%		 Else 'IF There is No SubSubID For the SchSubID then Check Wheather the "SubHead Name is NULL
			 'IF the SubHead Name is NULL/ or Acc Desc Is Yes then Taken the AccountHead Name
			 'and AccountHead for the SchSubID

			iSchSubSubID = 0
			sAccCheck = "Y"
			iEntNo = 0

			sQuery = "Select ScheduleSubHeadValue  "&_
					 "From Acc_T_ScheduleACDetail Where ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&"  "&_
					 "and ScheduleSubSubID = "&iSchSubSubID&" and Convert(Varchar,AsOnDate,103) = '"&sPreDate&"' "
			Objrs2.Open sQuery,Con
			IF Not Objrs2.EOF Then
				dPreYrVal = Objrs2(0)
			Else
				dPreYrVal = 0
			End IF
			Objrs2.Close


			'IF Cstr(Objrs1("SubHeadingName")) = "" Then
			IF CStr(Objrs1("EntryType")) = "A" Then

				'sQuery = "Select isNull(T.ApplicableACHeadCode,0),G.AccountDescription,T.ScheduleSubHeadValue,T.EntryNumber"&_
				'		 " From Acc_T_ScheduleACDetail T,Acc_M_GLAccountHead G Where "&_
				'		 "T.ScheduleID = "&sCatCode&" and T.ScheduleSubID = "&iSchID&" and T.ScheduleSubSubID = 0 and  "&_
				'		 "T.OrganisationCode = '"&sOrgID&"' and T.FinYear = '"&sFinYr&"' and "&_
				'		 "T.ApplicableACHeadCode = G.AccountHead "

				sQuery = "Select isNull(T.ApplicableACHeadCode,0) "&_
						 " From Acc_T_ScheduleACDetail T Where "&_
						 "T.ScheduleID = "&sCatCode&" and T.ScheduleSubID = "&iSchID&" and T.ScheduleSubSubID = 0 and  "&_
						 "T.OrganisationCode = '"&sOrgID&"' and T.FinYear = '"&sFinYr&"' "

				'Response.Write sQuery  &"<br><br>"
				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					iAccHead = Objrs2(0)
					sSchName = Objrs1("SubHeadingName")
					dClosing = GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay)
				Else
					iAccHead = 0
					sSchName = Objrs1("SubHeadingName")
				End IF
				'Response.Write "EntNo = "& iEntNo
				Objrs2.Close

				dPrvVal = 0
				iEntNo = 0
				
				sQuery = "Select T.ScheduleSubHeadValue,T.EntryNumber,isNull(ComputeMode,'')"&_
						 " From Acc_T_ScheduleACDetail T Where "&_
						 "T.ScheduleID = "&sCatCode&" and T.ScheduleSubID = "&iSchID&" and T.ScheduleSubSubID = 0 and  "&_
						 "T.OrganisationCode = '"&sOrgID&"' and T.FinYear = '"&sFinYr&"' "&_
						 "and Convert(Varchar,T.AsOnDate,103) = '"&sForTheDate&"' "

				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dPrvVal = Objrs2(0)
					iEntNo = Objrs2(1)
					
					if trim(Objrs2(2)) = "+" then
						nTot1 = CDbl(nTot1) + CDbl(Objrs2(0))
					elseif trim(Objrs2(2)) = "-" then
						nTot1 = CDbl(nTot1) - CDbl(Objrs2(0))
					end if 
				End IF
				'Response.Write "EntNo = "& iEntNo
				Objrs2.Close

			ElseIF CStr(Objrs1("EntryType")) = "D" Then
				sQuery = "Select T.ScheduleSubHeadValue,T.EntryNumber,isNull(ComputeMode,'')"&_
						 " From Acc_T_ScheduleACDetail T Where "&_
						 "T.ScheduleID = "&sCatCode&" and T.ScheduleSubID = "&iSchID&" and T.ScheduleSubSubID = 0 and  "&_
						 "T.OrganisationCode = '"&sOrgID&"' and T.FinYear = '"&sFinYr&"' "&_
						 "and Convert(Varchar,T.AsOnDate,103) = '"&sLastDay&"' "
				'Response.Write "EntryNum = "& sQuery &"<BR>"

				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					sAccCheck = "N"
					iAccHead = 0
					sSchName = Objrs1("SubHeadingName")
					dPrvVal = Objrs2(0)
					dClosing = 0
					iEntNo = Objrs2(1)
					
					
					if trim(Objrs2(2)) = "+" then
						nTot1 = CDbl(nTot1) + CDbl(Objrs2(0))
					elseif trim(Objrs2(2)) = "-" then
						nTot1 = CDbl(nTot1) - CDbl(Objrs2(0))
					end if 
					
				Else
					sAccCheck = "N"
					iAccHead = 0
					sSchName = Objrs1("SubHeadingName")
					dClosing = 0
					dPrvVal = 0
				End IF
				Objrs2.Close
			ElseIF CStr(Objrs1("EntryType")) = "S" Then
				dClosing = 0
				sAccCheck = "N"
				iAccHead = 0
				dSchAddTotal = 0
				dSchSubTotal = 0
				sSchName = Objrs1("SubHeadingName")

				dPrvVal = 0
				sQuery = "Select T.ScheduleSubHeadValue,isNull(ComputeMode,'')  "&_
						 "From Acc_T_ScheduleACDetail T Where "&_
						 "T.ScheduleID = "&sCatCode&" and T.ScheduleSubID = "&iSchID&" and T.ScheduleSubSubID = 0 and  "&_
						 "T.OrganisationCode = '"&sOrgID&"' and T.FinYear = '"&sFinYr&"' "

				'Response.Write sQuery &"<br><br>"
				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dPrvVal = Objrs2(0)
					if trim(Objrs2(1)) = "+" then
						nTot1 = CDbl(nTot1) + CDbl(dPrvVal)
					elseif trim(Objrs2(1)) = "-" then
						nTot1 = CDbl(nTot1) - CDbl(dPrvVal)
					end if 	
				End IF
				Objrs2.Close

				sQuery = "Select isNull(Sum(T.ScheduleSubHeadValue),0),isNull(V.ComputeMode,'') From Acc_T_SchdBreakupACDetail T,Vw_Acc_SchBreakSetup V "&_
						 "Where T.BreakupID = V.BreakUpID and V.ScheduleID = "&sCatCode&" and V.ScheduleSubID = "&iSchID&" "&_
						 "and V.ScheduleSubSubID = "&iSchSubSubID&" and T.FinYear = '"&sFinYr&"' "&_
						 "and T.BreakupSubID = V.BreakupSubID and T.BreakupSubSubID = V.BreakupSubSubID "&_
						 "Group By V.ComputeMode "
				'Response.Write sQuery

				Objrs2.Open sQuery,Con
				Do While Not Objrs2.EOF
					IF CStr(Objrs2(1)) = "+" Then
						dSchAddTotal = Objrs2(0)
					Else
						dSchSubTotal = Objrs2(0)
					End IF
					Objrs2.MoveNext
				Loop
				Objrs2.Close
				dClosing = CDbl(dSchAddTotal) - CDbl(dSchSubTotal)

				'dPrvVal = 0
			Else
				dClosing = 0
				sAccCheck = "N"
				iAccHead = 0
				dPrvVal = 0
			End IF

			IF CStr(iAccHead) <> "0" Then
				dClosing = 0
				sAccDesc = GetAccHeadName(iAccHead)
				'Response.Write sLastDay
				dClosing = GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay)
				nTot1 = CDbl(nTot1) + CDbl(dClosing)
				'Response.Write dClosing &"<br>"
			ElseIF CStr(Objrs1("EntryType")) = "S" Then
				sAccDesc = ""
				'dClosing = 0
				sSchName = Objrs1("SubHeadingName")
			Else
				sAccDesc = ""
				dClosing = 0
				sSchName = Objrs1("SubHeadingName")
			End IF

			IF CStr(Objrs1("DisplayACHeadDescr")) = "N" Then
				sSchName = Objrs1("SubHeadingName")
			End IF

			

	%>
			<tr>
				<%IF CStr(Objrs1("EntryType")) = "N" Then %>
					<td align="left" class="ExcelDisplayCell" colspan="6"><b>
					<% Response.Write(sSchName)%>
				<%Else%>
					<td align="left" class="ExcelDisplayCell"><b>
						<%Response.Write(sSchName)%></b>
				</td>
				<%IF CStr(Objrs1("EntryType")) = "D" Then%>				
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
				<%Else%>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dClosing,2,,,0)%></td>
				<%End IF %>
				
				<%nTot2 = CDbl(nTot2) + CDbl(dPreYrVal)%>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
				
				
				<%End IF %>
			</tr>

<%		End IF
		Objrs3.Close

		Objrs1.MoveNext
	loop
	Objrs1.Close
	%>
	<tr>
		<td class="ExcelDisplayCell" align="right"><B>Total</B></td>
		<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(nTot1,2,,,0)%></B></td>
		<td class="ExcelDisplayCell" align="right"><B><%=FormatNumber(nTot2,2,,,0)%></B></td>
	</tr>
	<%
	for nTempCtr = 1 to 5
	%>
	<tr>
		<td align="left" class="ExcelDisplayCell" colspan="3">&nbsp;</td>
	</tr>
	<%
	next
	rsSch.MoveNext 	
Loop 'Do While Not rsSch.EOF	
rsSch.Close
%>
							                        
										</table>
									</div> 
								</td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center">
                                                <%
												If iTotPage >= 2 Then 
													
													if iCurrentPage = 1 then
													%>
													<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
													<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
													<%else	%>			
													<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
													<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
    												<%end if	%>
    												<SELECT class="FormElem" onChange="Paginate(this.options[this.selectedIndex].value)" id=select1 name=select1>
    												<%	
														For lnPage = 1 To iTotPage
													%>
														<OPTION value="<%=lnPage%>" <%If lnPage = iCurrentPage Then Response.Write "Selected" %> >Page <%=lnPage%> of <%=iTotPage%></OPTION>
    												<%
    													next			
    												%>
    												</SELECT>
    												<%
    														if iCurrentPage = iTotPage then
    												%>
													<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
													<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>
    											
    												<%		else	
    													%>
													<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
													<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotPage%>')" id=button8 name=button8>
    												<%		end if
												End If 
													%>
													<%'="iLineCtr= " + trim(iLineCtr) %>
                                                
                                                <input type="button" value="Close" name="B1" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
</table>
</form>
</BODY>
<%
	Function GetAccHeadName(iAccHead)
		Dim sAccDesc,ObjAcc
		Set ObjAcc = Server.CreateObject("ADODB.RecordSet")
		IF iAccHead = "" Then
			iAccHead = 0
		End IF
		'sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where ACcountHead = "&iAccHead
		sQuery = "Select Distinct AccHeadAlias From VwOrgGLHeads WHere AccountHead = "&iAccHead&" and OUDefinitionID = '"&sOrgID&"' "

		'Response.Write sQuery
		ObjAcc.Open sQuery,Con
		IF Not ObjAcc.EOF Then
			sAccDesc = ObjAcc(0)
		Else
			sAccDesc = ""
		End IF
		ObjAcc.Close
		GetAccHeadName = sAccDesc
	End Function
%>

<%
	Function GetSchName(iNSchId,iNSchSId,sRetTy)
		Dim sSelSchName,sSelSchID,ObjAcc,sRetval,sTemp1,sTemp2
		sSelSchName = ""
		sTemp1 = ""
		sTemp2 = ""
		sSelSchID = ""
		Set ObjAcc = Server.CreateObject("ADODB.RecordSet")
		sQuery = "Select Distinct RTRIM(Cast(ScheduleID as VarChar)+'-'+Cast(ScheduleSubID as VarChar)+'-'+Cast(ScheduleSubSubID As VarChar)+'-'+Cast(BreakupID As VarChar)), "&_
				 "BreakupHeading From Vw_Acc_SchBreakSetup WHere ScheduleID = "&iNSchId&" and ScheduleSubID = "&iNSchSId&" "

		'Response.Write sQuery &"<br>"
		ObjAcc.Open sQuery,COn
		Do While Not ObjAcc.Eof
			sTemp1 = Replace(ObjAcc(1),":"," ")
			sTemp2 = ObjAcc(0)
			sSelSchName = sSelSchName&":"&sTemp1
			sSelSchID = sSelSchID&","&sTemp2
			ObjAcc.MoveNext
		Loop
		ObjAcc.Close
		sSelSchName = Mid(sSelSchName,2)
		sSelSchID = Mid(sSelSchID,2)
		IF CStr(sRetTy) = "F" Then
			sRetval = sSelSchName&"~~"&sSelSchID
		Else
			'sSelSchName = Replace(sSelSchName,":",",")
			sRetval = sSelSchName
		End IF
		'sRetval = iNSchId&"~~"&iNSchSId
		GetSchName = sRetval
	End Function
%>

