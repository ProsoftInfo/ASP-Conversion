<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SchBreakupSetup_forPL.asp
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
	Dim sQuery,Objrs1,Objrs2,Objrs3,Objrs4,iCtr,sOrgID,sCatCode,sFinYr,iAccHead
	Dim dClosing,iBkID,iBkSubID,iBkSubSubID,sLastDay,dPrvVal,sAccHeadName,iEntNo

	Dim nTempCtr,sTempVal,sTempVal2,dPreYrVal,dPreDate

	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs2 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs3 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs4 = Server.CreateObject("ADODB.RecordSet")

	'sOrgID = Request.Form("selUnitId")
	sOrgID = session("organizationcode")
	sCatCode = Request.Form("selSch")
	sFinYr = Session("FinPeriod")

	IF CStr(sCatCode) = "" Then
		sCatCode = "0"
	End IF

	IF CStr(sOrgID) = "" Then
		sOrgID = "0"
	End IF

	sLastDay = Request.Form("selForMonth")
	IF CStr(sLastDay) = "" Then
		sLastDay = "30/04/"&Trim(Left(sFinYr,4))
	End IF

	dPreDate = "31/03/"&Left(sFinYr,4)

	'Response.Write sLastDay

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
window.__itmsPopupCompat = {
	type: "scheduleSetupCaller",
	page: "SchBreakupSetup_forPL.asp",
	reloadPage: "SchBreakupSetup_forPL.asp",
	updateAction: "SchBreakupUpdate.asp",
	addPopup: "AddSchedBrkSubHeads.asp",
	addCatCode: "0",
	addFeatures: "dialogHeight:320px;dialogWidth:710px;center:Yes;help:No;resizable:No;status:No"
};
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" >
<Input type="hidden" name="hAccHeadNo" value="">
<Input type="hidden" name="hAccHeadName" value="">
<Input type="hidden" name="hEntNo" value="">
<Input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Schedule
          Breakup Setup
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
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<a href="SchSetup_ForPL.ASP"><td align="center">Schedule Setup</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="165">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								  	<tr>
									  	<td align="center">Schedule Breakup Setup</td>
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
								<td class="FieldCellSub" width="100">Organization</td>
                                <td class="FieldCell" width="110">
									<select size="1" name="selUnitId" class="FormElem" onChange="DisplayVal()">
										<OPTION value="0">Select</option>
											<%populateOrganizationListDBWithVal(sOrgID)%>
									</select>
								</td>
							  </tr>-->
							  <tr>
								<td class="FieldCellSub" width="100">Schedule Breakup</td>
										<td class="FieldCell" colspan="2">
											<select size="1" name="selSch" class="FormElem" onChange="DisplayVal()">
											<Option Value="0" Selected>Select</Option>
											<%
												sQuery = "Select Distinct ScheduleID,ScheduleHeading From Vw_Acc_SchSetup Where "&_
														 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' and "&_
														 "EntryType = 'S'  Order By ScheduleHeading "


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
														<Option Value="<%=Objrs1(0)%>" Selected><%=Objrs1(1)%></Option>
											<% Else %>
														<Option Value="<%=Objrs1(0)%>"><%=Objrs1(1)%></Option>
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
								<td class="FieldCellSub" width="120"></td>
								<td class="FieldCell">
								<Input type="Button" name="ButAdd" value="Schedule Breakup Setup" class="ActionButtonX" OnClick="AddPopUp()" >
                                &nbsp;</td>
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
<DIV class=frmBody id=frm4 style="width: 100%; height:230;">


<table BORDER="0" CELLSPACING="1"  CELLPADDING="0" class="ExcelTable" width="740">

<tr>

<td align="center" class="ExcelHeaderCell"><p>Heading Name</td>
<td align="center" class="ExcelHeaderCell">As on 31.03.<% =Right(sFinYr,4)%></td>
<td align="center" class="ExcelHeaderCell">Previous Value</td>
<td align="center" class="ExcelHeaderCell">As on 31.03.<% =Left(sFinYr,4)%></td>
<td align="center" class="ExcelHeaderCell">Acc Head Name</td>
<td align="center" class="ExcelHeaderCell">&nbsp;</td>
</tr>
<%
	sQuery = "Select Distinct ScheduleSubID,ScheduleSubSubID,SubHeadingName,Hierachy From Vw_Acc_SchSetup "&_
			 "Where ScheduleID = "&sCatCode&" and OrganisationCode = '"&sOrgID&"' "&_
			 "and FinYear = '"&sFinYr&"' and EntryType = 'S' Order By ScheduleSubID,ScheduleSubSubID,Hierachy "

	'Response.Write sQuery &"<br><br>"
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
%>
		<tr>
			<td align="left" class="ExcelDisplayCell" colspan="6"><b>
			<%Response.Write(Objrs1("SubHeadingName"))%></b></td>
		</tr>
<%
		sQuery = "Select Distinct BreakupID,BreakupHeading,Hierarchy From Vw_Acc_SchBreakSetup Where  "&_
				 "ScheduleSubID = "&Objrs1(0)&" and ScheduleSubSubID = "&Objrs1(1)&" and ScheduleID = "&sCatCode&" and "&_
				 "FinYear = '"&sFinYr&"' Order By Hierarchy "

		'Response.Write sQuery &"<BR>"
		With Objrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set Objrs2.ActiveConnection = Nothing
		Do WHile Not Objrs2.EOF
			iBkID = Objrs2("BreakupID")
			'Response.Write iBkID &"<BR><BR>"
%>
			<tr>
				<td align="left" class="ExcelDisplayCell" colspan="6">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>
					<%Response.Write(Objrs2("BreakupHeading"))%>
				</b></td>
			</tr>
<%
			sQuery = "Select Distinct BreakupSubID,BreakupSubSubID,SubBreakupName,DataEntry,Hierachy From Vw_Acc_SchBreakSetup  "&_
					 "Where ScheduleSubID = "&Objrs1(0)&" and ScheduleSubSubID = "&Objrs1(1)&" and ScheduleID = "&sCatCode&" and "&_
					 "BreakupID = "&Objrs2(0)&" and FinYear = '"&sFinYr&"' and Useable = 'Y' and BreakupSubID is NOT NULL Order By Hierachy "
			'Response.Write sQuery &"<BR><BR>"
				With Objrs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = Con
					.Open
				End With
				Set Objrs3.ActiveConnection = Nothing
				iAccHead = 0
				'nTempCtr = 0
				iEntNo = 0
				Do WHile Not Objrs3.EOF

					nTempCtr = nTempCtr + 1

					iBkSubID = Objrs3("BreakupSubID")
					iBkSubSubID  = Objrs3("BreakupSubSubID")

					'Talking Account Head for the BreakupID,BreakupSub and BreakupSUbSubID
					sQuery = "Select isNull(ApplicableACHeadCode,0) ApplicableACHeadCode,isNull(ScheduleSubHeadValue,0) ScheduleSubHeadValue,EntryNumber From Acc_T_SchdBreakupACDetail Where  "&_
							 "OrganisationCode = '"&sOrgID&"' and BreakupID = "&Objrs2(0)&" and BreakupSubID = "&Objrs3(0)&" "&_
							 "and BreakUpSubSubID = "&Objrs3(1)&" and FinYear = '"&sFinYr&"' "&_
							 "and Convert(Varchar,AsOnDate,103) = '"&sLastDay&"' "
					'Response.Write sQuery &"<br><br>"

					Objrs4.Open sQuery,con
					IF Not Objrs4.EOF Then
						iAccHead = Objrs4("ApplicableACHeadCode")
						dPrvVal = Objrs4("ScheduleSubHeadValue")
						iEntNo = Objrs4("EntryNumber")
					Else
						iAccHead = 0
						dPrvVal = 0
						iEntNo = 0
					End IF
					Objrs4.Close

					IF CStr(iEntNo) = "0" Then
						sTempVal2 = CreateEntry(Objrs2(0),Objrs3(0),Objrs3(1),sLastDay)
						sTempVal = Split(sTempVal2,"+")
						'Response.Write sTempVal2
						iEntNo = sTempVal(0)
						iAccHead = sTempVal(1)
					End IF

					IF CStr(iEntNo) <> "0" and CStr(iAccHead) = "0" Then
						iAccHead = GetOldAccHead(Objrs2(0),Objrs3(0),Objrs3(1))
					End IF

					IF CStr(iAccHead) <> 0 Then
						dClosing = GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay)
					Else
						dClosing = 0
					End IF

					'IF CStr(iBkID) = "13" and CStr(iBkSubID) = "10" Then
					'	dClosing = GetDayOpeningForSel(sOrgID,iAccHead,sLastDay,"C")
					'ENd IF

					'IF CStr(iBkID) = "13" and CStr(iBkSubID) = "11" Then
					'	dClosing = GetDayOpeningForSel(sOrgID,iAccHead,sLastDay,"C")
					'ENd IF

					'IF CStr(iBkID) = "11" and CStr(iBkSubID) = "18" Then
						'Response.Write iAccHead
					'	dClosing = GetDayOpeningForSel(sOrgID,iAccHead,sLastDay,"D")
						'Response.Write dClosing
						'Response.End
					'ENd IF

					'Test Break up Should be as BreakupID,BreakupSubID,BreakupSubSubID,iAccHead

					sQuery = "Select isNull(ScheduleSubHeadValue,0) From Acc_T_SchdBreakupACDetail Where  "&_
							 "OrganisationCode = '"&sOrgID&"' and BreakupID = "&Objrs2(0)&" and BreakupSubID = "&Objrs3(0)&" "&_
							 "and BreakUpSubSubID = "&Objrs3(1)&" "&_
							 "and Convert(Varchar,AsOnDate,103) = '"&dPreDate&"' "
					'Response.Write sQuery &"<br><br>"

					Objrs4.Open sQuery,con
					IF Not Objrs4.EOF Then
						dPreYrVal = Objrs4(0)
					Else
						dPreYrVal = 0
					End IF
					Objrs4.Close

			%>


					<tr>
						<td class="ExcelDisplayCell">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<%Response.Write(Objrs3("SubBreakupName"))%></td>
						<%IF CStr(Objrs3("DataEntry")) = "Y" Then %>
							<td class="ExcelInputCell" align="right"><input type="text" name="txtCurrVal<%=iBkID%>?<%=iBkSubID%>?<%=iBkSubSubID%>?<%=iAccHead%>" size="15" value="<%=FormatNumber(dPrvVal,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElem"></td>
						<%Else%>
							<td class="ExcelDisplayCell" align="right"><input type="text" Readonly name="txtCurrVal<%=iBkID%>?<%=iBkSubID%>?<%=iBkSubSubID%>?<%=iAccHead%>" size="15" value="<%=FormatNumber(dClosing,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElemRead"></td>
						<%ENd IF %>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="Left"><Span id="spAccHead<%=nTempCtr%> <%=iEntNo%>" class"DataOnly"><%=GetAccHeadName(iAccHead)%></Span></td>
						<Input type="hidden" name="hAccHead<%=iEntNo%>" Value="<%=iAccHead%>">
						<td class="ExcelDisplayCell" align="Left">
						<%IF CStr(Objrs3("DataEntry")) = "N" Then %>
							<Input type="button" value="Chg" class="ActionButtonX" onclick = "popAccList(<%=nTempCtr%>,<%=iEntNo%>)">
						<%Else%>
							&nbsp;
						<%ENd IF %>
						</td>
					</tr>
		<%			Objrs3.MoveNext
				loop
				Objrs3.Close
			Objrs2.MoveNext
			Loop
			Objrs2.Close

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
                                                 <!--input type="button" value="Delete" name="B4" class="ActionButton" onClick="Del()"-->

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
		Dim sAccDesc
		'sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where ACcountHead = "&iAccHead
		sQuery = "Select Distinct AccHeadAlias From VwOrgGLHeads WHere AccountHead = "&iAccHead&" and OUDefinitionID = '"&sOrgID&"' "
		Objrs4.Open sQuery,Con
		IF Not Objrs4.EOF Then
			sAccDesc = Objrs4(0)
		Else
			sAccDesc = ""
		End IF
		Objrs4.Close
		GetAccHeadName = sAccDesc
	End Function

	Function CreateEntry(sBkId,sBkSubId,sBkSubSubID,sLastDate)
		Dim iNewEntNo,sRetVal
		sQuery = "Select isNull(Max(EntryNumber),0)+1 From Acc_T_SchdBreakupACDetail "

		With Objrs4
			.ActiveConnection = Con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		End With
		Set Objrs4.ActiveConnection = Nothing
		IF Not Objrs4.EOF Then
			iNewEntNo = Objrs4(0)
		End IF
		Objrs4.Close

		sQuery = "Select Hierarchy,isNull(ApplicableACGroupCode,'0'),ComputeMode, "&_
				 "isNull(AddnDescription,'NULL'),ApplicableACHeadCode,FinYear "&_
				 "From ACC_T_SchdBreakupACDetail Where BreakUpID = "&sBkId&" "&_
				 "and BreakupSubID = "&sBkSubId&" and BreakupSubSubID = "&sBkSubSubID&" and "&_
				 "OrganisationCode = '"&sOrgID&"' "

		With Objrs4
			.ActiveConnection = Con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		End With
		Set Objrs4.ActiveConnection = Nothing
		IF Not Objrs4.EOF Then
			sQuery = "INSERT INTO ACC_T_SchdBreakupACDetail(EntryNumber, BreakupID, BreakupSubID, BreakupSubSubID, "&_
					 "Hierarchy, OrganisationCode, ApplicableACGroupCode, ApplicableACHeadCode, FinYear, "&_
					 "ScheduleSubHeadValue, ComputeMode, AsOnDate, AddnDescription) "&_
					 "VALUES ("&iNewEntNo&", "&iBkID&", "&iBkSubID&", "&iBkSubSubID&", "&Objrs4(0)&",  "&_
					 "'"&sOrgID&"', '"&Objrs4(1)&"', "&Objrs4(4)&", '"&Objrs4(5)&"', 0, "&_
					 "'"&Objrs4(2)&"', Convert(datetime,'"&sLastDate&"',103), '"&Objrs4(3)&"') "

			'Response.Write sQuery &"<br><br>"
			Con.Execute sQuery
			sRetVal = iNewEntNo&"+"&Objrs4(4)
		End IF
		Objrs4.Close
		CreateEntry = sRetVal
	End Function

	Function GetOldAccHead(sBkId,sBkSubId,sBkSubSubID)
		Dim iOldAccHead
		sQuery = "Select isNull(ApplicableACHeadCode,0) ApplicableACHeadCode From "&_
				 "Acc_T_SchdBreakupACDetail Where OrganisationCode = '"&sOrgID&"' and  "&_
				 "BreakupID = "&sBkId&" and BreakupSubID = "&sBkSubId&" and "&_
				 "BreakUpSubSubID = "&sBkSubSubID&" and ApplicableACHeadCode <> 0  "
		With Objrs4
			.ActiveConnection = Con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		End With
		Set Objrs4.ActiveConnection = Nothing
		IF Not Objrs4.EOF Then
			iOldAccHead = Objrs4(0)
		Else
			iOldAccHead = 0
		End IF
		Objrs4.Close
		GetOldAccHead = iOldAccHead
	End Function
%>
