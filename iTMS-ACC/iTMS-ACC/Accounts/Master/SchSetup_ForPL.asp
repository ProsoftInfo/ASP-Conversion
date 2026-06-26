<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	SchSetup_ForPL.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Manohar Prabhu .R
	'Created On					:	Nov 27 2003
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
	Dim sQuery,Objrs1,Objrs2,Objrs3,iCtr,sOrgID,sCatCode,sFinYr,iAccHead
	Dim dClosing,iSchID,iSchSubID,iSchSubSubID,sLastDay,dPrvVal,sSchName
	Dim iShId,iShSub,iShSubSubId,sForTheDate,sPreDate,dPreYrVal
	Dim sAccDesc,sAccCheck,iEntNo,dHeadVal
	Dim nTempCtr,dSchAddTotal,dSchSubTotal

	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs2 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs3 = Server.CreateObject("ADODB.RecordSet")


	sOrgID = session("organizationcode") 'Request("selUnitId")
	sCatCode = Request("selSch")
	sFinYr = Session("FinPeriod")
	sForTheDate = Trim(Request.Form("selForMonth"))



	IF CStr(sCatCode) = "" Then
		sCatCode = "0"
	End IF

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
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=VBScript>
Function ViewSchedule(sCallFrom)

	Dim sForTheDate,sOptions,sOrgID
	
	sOptions = "height=540,width=795, toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0"
    	
	sForTheDate = trim(document.formname.selForMonth.value)
	sOrgID = document.formname.hOrgID.value
	    
    sTempValues = "ForTheDate=" & sForTheDate & "&OrgID=" & sOrgID & "&CallFrom=" & sCallFrom
    
    win = open ( "View_Schedule.asp" & "?" & sTempValues, "",sOptions)
End Function
        
	Function CheckVal(sObj)
		IF Len(sObj.Value) > 100 Then
			MsgBox "Account Head Name Should be Less than 100 Characters "
			sObj.focus()
			Exit Function
		End IF

		IF Len(Trim(sObj.Value)) = 0 Then
			MsgBox "Account Head Name Should be blank "
			sObj.focus()
			Exit Function
		End IF
	End Function

	Function DisplayVal()
		'IF document.formname.selUnitId.selectedIndex <> 0 Then
			document.formname.action = "SchSetup_ForPL.asp"
			document.formname.submit()
		'End IF
	End Function

	Function CheckSubmit()

		'alert(document.formname.hAccCode.Value)
		'exit function

		'IF document.formname.selUnitId.selectedIndex = 0 Then
		'	MsgBox "Select Organization"
		'	document.formname.selUnitId.focus()
		'	Exit Function
		'End IF

		IF document.formname.selSch.selectedIndex = 0 Then
			MsgBox "Select Schedule"
			document.formname.selSch.focus()
			Exit Function
		End IF

		document.formname.action ="SchUpdateOld.asp"
		document.formname.submit()
	End Function
'**************************Added By Maheshwari on 03 Oct 2006*******************************
Function popAccList(nPassCtr,nEntNo,nEntType)

dim iUnitNo,saTemp,iGlHead,sRetVal,OutValue,sAccHeadName,sNewAcc

document.formname.hEntNo.value = nEntNo
document.formname.htype.value = nEntType
'if document.formname.selUnitId.selectedIndex >0 then
	sOrgId= document.formname.hOrgID.value

	OutValue= showModalDialog("ChgAccHeadName.asp?orgId="+sOrgId,"","dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")

	'arrTemp = split(OutValue,":")
	while UBound(arrTemp) = 0
		OutValue = showModalDialog("ChgAccHeadName.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
	sRetVal = OutValue
	'alert(sRetVal)

	If UBound(arrTemp)<=1 then exit function
	Set sNewAcc = Eval("document.formname.hAccHead"&nEntNo)
	sNewAcc.value = arrTemp(0)
	'alert(sNewAcc.value)
	document.formname.hAccHeadNo.value= sNewAcc.value
	sAccHeadName = arrTemp(1)
	eval("spAccHead" & nPassCtr & nEntNo &"").innerHTML=sAccHeadName
	document.formname.hAccHeadName.value=sAccHeadName
'Else
'	MsgBox "Select Unit"
'	document.formname.selUnitId.focus
'End If
End Function
'****************************************************************************************
Function SelAccList(nPassCtr,nCatCode,nSchID,nSchSubSubID,nEntType,shdType)
	Dim sOrgId,sRetVal,OutValue,sAccHeadName,sNewAcc,sAccList,iNewSchSubId
	Dim sAccCode,sAccName,arrTemp,iCtr,iBkPara
	document.formname.hEntNo.value = nEntNo
	document.formname.htype.value = nEntType
	iBkPara = nCatCode &":"&nSchID&":"&nSchSubSubID
	'MsgBox iBkPara
	document.formname.hSchdid.Value = iBkPara
	'alert(document.formname.hSchdid.Value)
	'If document.formname.selUnitId.selectedIndex > 0 Then
		sOrgId= document.formname.hOrgID.value
		'MsgBox Eval("document.formname.hSchHead"&nPassCtr & nSchID & nEntNo).Value
		OutValue= showModalDialog("SelAccHeadName.asp?sTemp="&iBkPara,"","dialogHeight:480px;dialogWidth:700px;center:Yes;help:No;resizable:No;status:No")
		'OutValue = window.open("SelAccHeadName.asp?sTemp="&iBkPara,"NewWin","Width=700,height=480")
		sRetVal = OutValue
		arrTemp = split(sRetVal,"~~")
		sAccName = arrTemp(0)
		sAccCode = arrTemp(1)
		iCtr = arrTemp(2)
		iNewSchSubId = arrTemp(3)
		IF CStr(shdType) <> "A" Then
			eval("spSchHead" & nPassCtr & nSchID & nEntNo &"").innerHTML = sAccName
			Eval("document.formname.hSchHead"&nPassCtr & nSchID & nEntNo).Value = sRetVal
			'MsgBox sRetVal
			document.formname.hAccCode.Value = document.formname.hAccCode.Value&"*"&sAccCode&":"&iCtr&":"&iNewSchSubId
		Else
			eval("spSchHead" & nPassCtr & nSchID & nSchSubSubID).innerHTML = sAccName
			Eval("document.formname.hSchHead"&nPassCtr & nSchID & nSchSubSubID).Value = sRetVal
		End IF
		'alert(document.formname.hAccCode.Value)
	'End if
End Function
'****************************************************************************************
Function SchdSetupPopUp(sCallFrom)
Dim RetVal
Dim sOrgId,sSchID,sInsDate
	sSchID = 0
	sOrgId = 0
	sSchID = document.formname.selSch.value
	sOrgId = document.formname.hOrgID.value
	sInsDate = document.formname.selForMonth.value
	If sOrgId = 0 Then
		Msgbox "Select Organization And Continue...!",vbOKOnly
		Exit Function
	End If
	RetVal = ShowModalDialog("BalScheduleSetUp.asp?sUnit="&sOrgId&"&sSchID="&sSchID&"&InsDate="&sInsDate&"&CallFrom="&sCallFrom,"A","dialogHeight:250px;dialogWidth:620px;center:Yes;help:No;resizable:No;status:No")
	IF CStr(RetVal) = "Y" Then
		document.formname.action = "SchSetup_ForPL.asp"
		document.formname.submit()
	End IF

End Function
'****************************************************************************************
Function AddPopUp()
Dim sOrgId,sSchedName,sInsDate

	sOrgId = document.formname.hOrgID.value
	sSchedName = document.formname.selSch.value
	sCatCode = document.formname.hCatCode.value
	sInsDate = document.formname.selForMonth.value
	Dim sVal

	'ShowModalDialog "AddSchedSubHeads.asp?sUnit="&sOrgId&"&sSchName="&sSchedName&"&sCatCode="&sCatCode,"dialogHeight:680px;dialogWidth:900px;center:Yes;help:No;resizable:No;status:No"
	'window.open "AddSchedSubHeads.asp?sUnit="&sOrgId&"&sSchName="&sSchedName&"&sCatCode="&sCatCode,"A","Height:680px;Width:900px;center:Yes"
	'window.open "AddSchedSubHeads.asp?sUnit="&sOrgId&"&sSchName="&sSchedName&"&sCatCode="&sCatCode,"A",""
	sVal = showModalDialog("AddSchedSubHeads.asp?sUnit="&sOrgId&"&sSchName="&sSchedName&"&sCatCode="&sCatCode&"&InsDate="&sInsDate,"A","dialogHeight:330px;dialogWidth:760px;center:Yes;help:No;resizable:No;status:No")
	IF CStr(sVal) = "Y" Then
		document.formname.action = "SchSetup_ForPL.asp"
		document.formname.submit()
	End IF
End Function
Function Checkreturn()

End Function
'****************************************************************************************
</Script>
<script language="javascript">
window.__itmsPopupCompat = {
	type: "scheduleSetupCaller",
	page: "SchSetup_ForPL.asp",
	reloadPage: "SchSetup_ForPL.asp",
	updateAction: "SchUpdateOld.asp",
	addPopup: "AddSchedSubHeads.asp",
	addFeatures: "dialogHeight:330px;dialogWidth:760px;center:Yes;help:No;resizable:No;status:No"
};
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" onload ="Checkreturn()">
<Input type="hidden" name="hAccHeadNo" value="">
<Input type="hidden" name="hAccHeadName" value="">
<Input type="hidden" name="hEntNo" value="">
<Input type="hidden" name="hAccCode" value="">
<Input type="hidden" name="hType" value="">
<Input type="hidden" name="hSchdid" value="">
<input type="hidden" name="hCatCode" value="<%=sCatCode%>">
<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Schedule
          Setup
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<a href="ConfigureProfitLossAcc.ASP"><td align="center">GL A/C Alias</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
										<tr>
											<td align="center">Schedule Setup</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="165">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<a href="SchBreakupSetup_forPL.asp"><td align="center">Schedule Breakup Setup</td></a>
									</tr>
								  </table>
								</td>
								
								<td class="TabCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<a href="PLSetup.asp"><td align="center">PL Setup</td></a>
									</tr>
								  </table>
								</td>
								<!--
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<a href="BSSetup.asp"><td align="center">BS Setup</td></a>
								  	</tr>
								  </table>
								</td>
								-->
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                             <!--<tr>
								<td class="FieldCellSub" width="50">Organization</td>
                                <td class="FieldCell" width="110">
									<select size="1" name="selUnitId" class="FormElem" onChange="DisplayVal()">
										<OPTION value="0">Select</option>
											<% populateOrganizationListDBWithVal(sOrgID)%>
									</select>
								</td>
							  </tr>-->
							  <tr>
								<td class="FieldCellSub" width="100">Upto The Month</td>
                                <td class="FieldCell" width="110">
									<select size="1" name="selForMonth" class="FormElem" onChange="DisplayVal()">
										<%
											Dim sFromYear,sToYear,iCounter
											sFromYear=CDbl(Right(getFromFinYear(),4)&Left(getFromFinYear(),2))
											sToYear=CDbl(Right(getToFinYear(),4)&Left(getToFinYear(),2))
											iCounter=sFromYear
											do while iCounter<=sToYear
												IF CStr(sLastDay) = CStr(GetLastDayMonYr(iCounter)) Then
													Response.Write "<option Value="""&GetLastDayMonYr(iCounter)&""" Selected>"&MonthName(Right(iCounter,2))&"-"&Left(iCounter,4)&"</option>"
												Else
													Response.Write "<option Value="""&GetLastDayMonYr(iCounter)&""">"&MonthName(Right(iCounter,2))&"-"&Left(iCounter,4)&"</option>"
												End IF

												iCounter=CDbl(iCounter)+1
												if CDbl(Right(iCounter,2))>12 then
													iCounter=CDbl(CDbl(Left(iCounter,4))+1&"01")
												end if
											loop
											%>

									</select>
								</td>
							  </tr>
							  <tr>
								<td class="FieldCellSub" width="50">Schedule</td>
										<td class="FieldCell" colspan="2">
											<select size="1" name="selSch" class="FormElem" onChange="DisplayVal()">
											<Option Value="0" Selected>Select</Option>
											<%
												sQuery = "Select Distinct ScheduleID,RTRIM('Schedule-'+Cast(ScheduleNumber as Char)) ScheduleNo, "&_
														 "ScheduleHeading,Hierarchy From Vw_Acc_SchSetup Where ApplicableFor='P' "&_
														 " and OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' Order By Hierarchy "

												With Objrs1
													.CursorLocation = 3
													.CursorType = 3
													.Source = sQuery
													.ActiveConnection = Con
													.Open
												End WIth
												Set Objrs1.ActiveConnection = Nothing
												Do While Not Objrs1.EOF
													IF CStr(sCatCode) = CStr(Objrs1(0)) Then
											%>
														<Option Value="<%=Objrs1(0)%>" Selected><%=Objrs1(1)%> - <%=Objrs1(2)%></Option>
												<% Else %>
														<Option Value="<%=Objrs1(0)%>"><%=Objrs1(1)%> - <%=Objrs1(2)%></Option>
											<%
													End IF

												Objrs1.MoveNext
												Loop
												Objrs1.Close
											%>
											</select>
										</td>
										</tr>
										<tr>
											<td>&nbsp&nbsp&nbsp&nbsp </td>
											<td><Input type="Button" value="Schedule Setup" class="ActionButtonX" id=button4 name=ButAddSchedHeads onclick="SchdSetupPopUp('PL')"></td>
											<td><Input type="Button" value="Schedule Head Setup" class="ActionButtonX" id=button3 name=ButAdd onclick="AddPopUp()"></td>



							</tr>
							</table>
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
   <% IF CStr(sOrgID) <> "" Then %>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
<DIV class=frmBody id=frm4 style="width: 100%; height:240;">
<table BORDER="0" CELLSPACING="1"  CELLPADDING="0" class="ExcelTable" width="750">
<tr>
	<td align="center" class="ExcelHeaderCell"><p>Heading Name</td>
	<td align="center" class="ExcelHeaderCell">As on <%=Replace(sForTheDate,"/",".")%></td>
	<td align="center" class="ExcelHeaderCell">Previous Value</td>
	<td align="center" class="ExcelHeaderCell">As on 31.03.<%=Left(sFinYr,4)%></td>
	<td align="center" class="ExcelHeaderCell">Acc / Sch Det</td>
	<td align="center" class="ExcelHeaderCell">&nbsp;</td>
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

				sQuery = "Select ScheduleSubHeadValue  "&_
						 "From Acc_T_ScheduleACDetail Where ScheduleID = "&sCatCode&" and ScheduleSubID = "&iSchID&"  "&_
						 "and ScheduleSubSubID = 0 and Convert(Varchar,AsOnDate,103) = '"&sPreDate&"' "
				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dPreYrVal = Objrs2(0)
				Else
					dPreYrVal = 0
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




	%>
					<td class="ExcelInputCell" align="right">
						<input type="text" name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" value="<%=FormatNumber(dHeadVal,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElem">
					</td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="left">
	<%				IF CStr(Objrs1("EntryType")) = "A" Then %>
						<SPan id="spAccHead<%=nTempctr%><%=iEntNo%>" class="DataOnly"><%=sAccDesc%></Span>
						<Input type="hidden" name="hAccHead<%=iEntNo%>" Value="<%=iAccHead%>">
						</td>
						<td class="ExcelDisplayCell" align="right">
						<Input type="button" value="Chg4" class="ActionButtonX" id=button1 name=button1 onclick = "popAccList(<%=nTempctr%>,<%=iEntNo%>,'A')">
	<%				Elseif CStr(Objrs1("EntryType")) = "S" Then %>
						<SPan id="spSchHead<%=nTempctr%><%=iSchID%><%=iSchSubSubID%>" class="DataOnly">&nbsp;</Span>
						<Input type="hidden" name="hSchHead<%=nTempctr%><%=iSchID%><%=iSchSubSubID%>" Value="<%=GetSchName(sCatCode,iSchID,"F")%>~~<%=iSchID%>~~<%=sCatCode%>:<%=iSchID%>:<%=iSchSubSubID%>">
						<td class="ExcelDisplayCell" align="right">
						<Input type="button" value="Sch" class="ActionButtonX" id=button2 name=button2 onclick = "SelAccList(<%=nTempctr%>,<%=sCatCode%>,<%=iSchID%>,<%=iSchSubSubID%>,'S','A')">
	<%				Else%>
						<td class="ExcelDisplayCell" align="right">
							<SPan id="spAccHead" class="DataOnly"></Span></td>
	<%				End IF %>
					</td>
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
	%>
					<td class="ExcelDisplayCell" align="right">
						<input type="text" name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" value="<%=FormatNumber(dHeadVal,2,,,0)%>" style="text-align:right" readonly maxlength="13" class="FormElemRead">
					</td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="left">
	<%				IF CStr(Objrs1("EntryType")) = "A" Then %>
						<SPan id="spAccHead<%=nTempctr%><%=iEntNo%>" class="DataOnly"><%=sAccDesc%></Span>
						<Input type="hidden" name="hAccHead<%=iEntNo%>" Value="<%=iAccHead%>">
						</td>
						<td class="ExcelDisplayCell" align="right">
						<Input type="button" value="Chg" class="ActionButtonX" id=button1 name=button1 onclick = "popAccList(<%=nTempctr%>,<%=iEntNo%>,'A')">
	<%				Elseif CStr(Objrs1("EntryType")) = "S" Then %>
						<SPan id="spSchHead<%=nTempctr%><%=iSchID%><%=iSchSubSubID%>" class="DataOnly">&nbsp;</Span>
						<Input type="hidden" name="hSchHead<%=nTempctr%><%=iSchID%><%=iSchSubSubID%>" Value="<%=GetSchName(sCatCode,iSchID,"F")%>~~<%=iSchID%>~~<%=sCatCode%>:<%=iSchID%>:<%=iSchSubSubID%>">
						<td class="ExcelDisplayCell" align="right">
						<Input type="button" value="Sch" class="ActionButtonX" id=button2 name=button2 onclick = "SelAccList(<%=nTempctr%>,<%=sCatCode%>,<%=iSchID%>,<%=iSchSubSubID%>,'S','A')">
	<%				Else%>
						<td class="ExcelDisplayCell" align="right">
							<SPan id="spAccHead" class="DataOnly"></Span></td>
	<%				End IF %>
					</td>
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

					%></td>
					<% IF CStr(Objrs3("EntryType")) = "D" Then 'IF it is an Data Entry the User Will enter the Value %>
						<td class="ExcelInputCell" align="right"><input type="text" name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" value="<%=FormatNumber(dPrvVal,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElem"></td>
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
					%>
						 	<td class="ExcelDisplayCell" align="right"><input type="text" name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" readonly value="<%=FormatNumber(Objrs2(0),2,,,0)%>" style="text-align:right" maxlength="13" class="FormElemRead"></td>
					<% 	Else %>
							<td class="ExcelDisplayCell" align="right"><input type="text" name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" readonly value="0.00" style="text-align:right" maxlength="13" class="FormElemRead"></td>
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




					%>
						<td class="ExcelDisplayCell" align="right"><input type="text" name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" readonly value="<%=FormatNumber(dClosing,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElemRead"></td>
					<%ElseIF CStr(Objrs3("EntryType")) = "N" Then 'If it is an AccountHead%>
						<td class="ExcelDisplayCell" align="right">&nbsp;</td>
					<%ENd IF
					  IF CStr(Objrs3("EntryType")) = "N" Then
					%>
					<td class="ExcelDisplayCell" align="right" colspan="3">&nbsp;</td>
					<%Else%>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="left">
					<%ENd IF %>

					<% IF CStr(Objrs3("EntryType")) = "A" Then %>
						<SPan id="spAccHead<%=nTempctr%><%=iEntNo%>" class="DataOnly"><%=sAccDesc%></Span>
						<Input type="hidden" name="hAccHead<%=iEntNo%>" Value="<%=iAccHead%>">
						</td>
					<%Elseif CStr(Objrs3("EntryType")) = "S" Then %>
						<SPan id="spSchHead<%=nTempctr%><%=iSchID%><%=iSchSubSubID%>" class="DataOnly">&nbsp;</Span>
						<Input type="hidden" name="hSchHead<%=nTempctr%><%=iSchID%><%=iSchSubSubID%>" Value="<%=GetSchName(sCatCode,iSchID,"F")%>~~<%=iSchID%>~~<%=sCatCode%>:<%=iSchID%>:<%=iSchSubSubID%>">
					<%Else%>
						<SPan id="spAccHead" class="DataOnly"></Span></td>
					<%End IF %>
					<td class="ExcelDisplayCell" align="right">
					<% IF CStr(Objrs3("EntryType")) = "A" Then %>
						<Input type="button" value="Chg" class="ActionButtonX" id=button1 name=button1 onclick = "popAccList(<%=nTempctr%>,<%=iEntNo%>,'A')">
					<%Elseif CStr(Objrs3("EntryType")) = "S" Then %>
						<Input type="button" value="Sch" class="ActionButtonX" id=button2 name=button2 onclick = "SelAccList(<%=nTempctr%>,<%=sCatCode%>,<%=iSchID%>,<%=iSchSubSubID%>,'S','A')">

					<%Else%>
						&nbsp;
					<%ENd IF %>
					</td>
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

				sQuery = "Select T.ScheduleSubHeadValue,T.EntryNumber"&_
						 " From Acc_T_ScheduleACDetail T Where "&_
						 "T.ScheduleID = "&sCatCode&" and T.ScheduleSubID = "&iSchID&" and T.ScheduleSubSubID = 0 and  "&_
						 "T.OrganisationCode = '"&sOrgID&"' and T.FinYear = '"&sFinYr&"' "&_
						 "and Convert(Varchar,T.AsOnDate,103) = '"&sForTheDate&"' "

				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dPrvVal = Objrs2(0)
					iEntNo = Objrs2(1)
				Else
					dPrvVal = 0
					iEntNo = 0
				End IF
				'Response.Write "EntNo = "& iEntNo
				Objrs2.Close

			ElseIF CStr(Objrs1("EntryType")) = "D" Then
				sQuery = "Select T.ScheduleSubHeadValue,T.EntryNumber"&_
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

				sQuery = "Select T.ScheduleSubHeadValue  "&_
						 "From Acc_T_ScheduleACDetail T Where "&_
						 "T.ScheduleID = "&sCatCode&" and T.ScheduleSubID = "&iSchID&" and T.ScheduleSubSubID = 0 and  "&_
						 "T.OrganisationCode = '"&sOrgID&"' and T.FinYear = '"&sFinYr&"' "

				'Response.Write sQuery &"<br><br>"
				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dPrvVal = Objrs2(0)
				Else
					dPrvVal = 0
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
				<%IF CStr(Objrs1("EntryType")) = "D" Then %>
					<td class="ExcelInputCell" align="right">
						<input type="text" name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" value="<%=FormatNumber(dPrvVal,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElem"></td>
				<%Else%>
					<td class="ExcelDisplayCell" align="right">
						<input type="text" readonly name="txtCurrVal<%=sCatCode%>?<%=iSchID%>?<%=iSchSubSubID%>?<%=iAccHead%>" size="15" value="<%=FormatNumber(dClosing,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElemRead"></td>
				<%End IF %>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="left">
				<%IF CStr(Objrs1("EntryType")) = "A" Then

				%>
					<SPan id="spAccHead<%=nTempctr%> <%=iEntNo%>" class="DataOnly"><%=sAccDesc%></Span>
					<Input type="hidden" name="hAccHead<%=iEntNo%>" Value="<%=iAccHead%>">
				<%Else%>
					<SPan id="spSchHead<%=nTempctr%> <%=iSchID%>" class="DataOnly"><%=GetSchName(sCatCode,iSchID,"N")%></Span>
					<Input type="hidden" name="hSchHead<%=nTempctr%><%=iSchID%>" Value="<%=GetSchName(sCatCode,iSchID,"F")%>~~<%=iSchID%>~~<%=sCatCode%>:<%=iSchID%>:<%=iSchSubSubID%>">

				<%End IF %>

				</td>
				<td class="ExcelDisplayCell" align="right">
				<%IF CStr(Objrs1("EntryType")) = "A" Then %>
					<Input type="button" value="Chg" class="ActionButtonX" id=button1 name=button1 onclick = "popAccList(<%=nTempctr%>,<%=iEntNo%>,'A')">
				<%Elseif CStr(Objrs1("EntryType")) = "S" Then %>
					<Input type="button" value="Sch" class="ActionButtonX" id=button2 name=button2 onclick = "SelAccList(<%=nTempctr%>,<%=sCatCode%>,<%=iSchID%>,<%=iSchSubSubID%>,'S')">

				<%End IF %>
				</td>
				<%End IF %>
			</tr>

<%		End IF
		Objrs3.Close

		Objrs1.MoveNext
	loop
	Objrs1.Close
%>
</table>
</DIV>
								</td>
								<td align="center">
								</td>
                            </tr>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
     <% End IF %>

                 <tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center">
                                                <input type="button" value="Save" name="B4" class="ActionButton" onClick="CheckSubmit()">
                                                <!--input type="button" value="Sch Breakup Setup" name="B6" class="ActionButtonX" onClick="SchBrk()" -->
                                                <input type="Reset" value="Reset" name="B5" class="ActionButton">
                                                <input type="button" value="Schedule View" name="But_View" class="ActionButtonX" onClick="ViewSchedule('P')">
												<!--input type="button" value="Cancel" name="B5" class="ActionButton"-->

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
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
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

