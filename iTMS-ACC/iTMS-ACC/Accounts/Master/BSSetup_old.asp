<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BSSetup.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Manohar Prabhu .R
	'Created On					:	Dec 27 2006
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<%
	Dim sQuery,Objrs1,Objrs2,Objrs3,iCtr,sOrgID,sCatCode,sFinYr,iAccHead
	Dim dClosing,iBSID,iBSSubID,iBSSubSubID,sLastDay,dPrvVal,sSchName,sEntTy
	Dim dAddTot,dSubTot,dSchTot,dPreYrVal,sPreDate

	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs2 = Server.CreateObject("ADODB.RecordSet")

	'sOrgID = Request.Form("selUnitId")
	sOrgID = session("organizationcode")
	sCatCode = Request.Form("selSch")
	sFinYr = Session("FinPeriod")

	sLastDay = Request.Form("selForMonth")
	sPreDate = "31/03/"&Trim(Left(sFinYr,4))

	IF CStr(sCatCode) = "" Then
		sCatCode = "0"
	End IF

	IF CStr(sOrgID) = "" Then
		sOrgID = "0"
	End IF

	IF CStr(sLastDay) = "" Then
		sLastDay = "30/04/"&Trim(Left(sFinYr,4))
	End IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" ID="TempData">
	<Root/>
</script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = {
	type: "scheduleSetupCaller",
	page: "BSSetup.asp",
	reloadPage: "BSSetup.asp",
	addPopup: "AddSchedBSSubHeads.asp",
	addIncludesCat: false,
	addFeatures: "dialogHeight:300px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No",
	deleteScheduleId: "7",
	deleteReloadPage: "BSSetup.asp"
};
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="BSUpdate.asp">
	<Input type="Hidden" name="hOrgID" value="<%=sOrgID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Balance Sheet
          Setup
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<a href="ConfigureBalanceSheet.ASP"><td align="center">GL A/C Alias</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<a href="SchSetup.ASP"><td align="center">Schedule Setup</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="165">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<a href="SchBreakupSetup.ASP"><td align="center">Schedule Breakup Setup</td></a>
									</tr>
								  </table>
								</td>
								<!--
								<td class="TabCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<a href="PLSetup.asp"><td align="center">PL Setup</td></a>
									</tr>
								  </table>
								</td>
								-->
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
								  	<tr>
								  		<td align="center">BS Setup</td>
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
								<td class="FieldCellSub" width="50">Organization</td>
                                <td class="FieldCell" width="110">
									<select size="1" name="selUnitId" class="FormElem" onChange="DisplayVal()">
										<OPTION value="0">Select</option>
											<%populateOrganizationListDBWithVal(sOrgID)%>
									</select>
								</td>
							  </tr>-->
							  <tr>
								<td class="FieldCellSub" width="50">Schedule</td>
										<td class="FieldCell" colspan="2">
											<select size="1" name="selSch" class="FormElem" onChange="DisplayVal()">
											<Option Value="0" Selected>Select</Option>
											<%
												sQuery = "Select BSHeadID,BSHeading From ACC_M_BSSetupHeads  Where FinYear='"&sFinYr&"'  "

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
									<td class="FieldCellSub" width="50"></td>
									<td class="FieldCell">
									<Input type="Button" name="ButSubmit" value="B/S Setup" class="ActionButtonX" OnClick="AddPopUp()" >
									<Input type="Button" name="ButDelete" value="Delete" class="ActionButtonX" OnClick="Del()" >
									&nbsp;</td>
								</tr>


							</tr>
							</table>
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
   <%IF CStr(sOrgID) <> "" Then %>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
<DIV class=frmBody id=frm4 style="width: 100%; height:230;">
<table BORDER="0" CELLSPACING="1"  CELLPADDING="0" class="ExcelTable">

<tr>

<td align="center" class="ExcelHeaderCell"><p>Heading Name</td>
<td align="center" class="ExcelHeaderCell">As on 31.03.<%=Right(sFinYr,4)%></td>
<td align="center" class="ExcelHeaderCell">Previous Value</td>
<td align="center" class="ExcelHeaderCell">As on 31.03.<%=Left(sFinYr,4)%></td>
</tr>
<%

	'Taking The SchSubID and Sch SubSUbID,Name and EntryType for the Selected Schedule
	sQuery = "Select BSSubID,BSSubSubID,BSSubHeadingName,isNull(EntryType,'N') EntryType From  "&_
			 "ACC_M_BSSetupSubHeads Where BSHeadID = "&sCatCode&"  and FinYear = '"&sFinYr&"' "&_
			 "Order By BSSubID,BSSubSubID,Hierachy "

	'Response.Write sQuery &"<br><br>"
	With Objrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End With
	Set Objrs1.ActiveConnection = Nothing

	Do WHile Not Objrs1.EOF
		iBSSubID = Objrs1("BSSubID")
		iBSSubSubID = Objrs1("BSSubSubID")
		sEntTy = Objrs1("EntryType")

		sQuery = "Select isNull(ApplicableACHeadCode,0) From ACC_T_BSACDetail Where BSHeadID = "&sCatCode&" "&_
				 "and BSSubID = "&iBSSubID&" and BSSubSUbID = "&iBSSubSubID&" and "&_
				 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' "

		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			iAccHead = Objrs2(0)
		Else
			iAccHead = 0
		End IF
		Objrs2.Close

		sQuery = "Select BSSubHeadValue From ACC_T_BSACDetail Where BSHeadID = "&sCatCode&" "&_
				 "and BSSubID = "&iBSSubID&" and BSSubSUbID = "&iBSSubSubID&" and "&_
				 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' "&_
				 "and Convert(Varchar,AsOnDate,103) = '"&sLastDay&"' "

		'Response.Write sQuery &"<br><br>"
		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dPrvVal = Objrs2(0)
		Else
			dPrvVal = 0
		End IF

		Objrs2.Close

		sQuery = "Select isNull(BSSubHeadValue,0) From ACC_T_BSACDetail Where BSHeadID = "&sCatCode&" "&_
				 "and BSSubID = "&iBSSubID&" and BSSubSUbID = "&iBSSubSubID&" and "&_
				 "OrganisationCode = '"&sOrgID&"'  "&_
				 "and Convert(Varchar,AsOnDate,103) = '"&sPreDate&"' "


		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dPreYrVal = Objrs2(0)
		Else
			dPreYrVal = 0
		End IF
		Objrs2.Close

%>
				<tr>
					<td class="ExcelDisplayCell">
					<%IF CStr(iBSSubSubID) <> "0" Then
						Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
					End IF
					Response.Write(Objrs1("BSSubHeadingName"))
					%>

					</td>
					<%IF CStr(sEntTy) = "D" Then 'IF it is an Data Entry the User Will enter the Value %>
						<td class="ExcelInputCell" align="right"><input type="text" name="txtCurrVal<%=sCatCode%>?<%=iBSSubID%>?<%=iBSSubSubID%>" size="15" value="<%=FormatNumber(dPrvVal,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElem">
						
						</td>
					<%ElseIF CStr(sEntTy) = "S" Then 'If is is and Scheduled Type
						sQuery = "Select isNull(SUM(S.ScheduleSubHeadValue),0),S.ComputeMode From Acc_T_ScheduleACDetail S, "&_
								 "ACC_T_BSACDetail P Where P.BSHeadID = "&sCatCode&" and P.BSSubID = "&iBSSubID&" "&_
								 " and  P.BSSubSubID = "&iBSSubSubID&" and P.ScheduleID = S.ScheduleID "&_
								 " and P.ScheduleSubSubID = S.ScheduleSubSubID " & _
								 " and P.ScheduleSubID = S.ScheduleSubID  "&_
								 " and S.OrganisationCode = P.OrganisationCode and S.FinYear = P.FinYear "&_
								 " and P.OrganisationCode = '"&sOrgID&"' and P.FinYear = '"&sFinYr&"' "&_
								 "and Convert(Varchar,P.AsOnDate,103) = '"&sLastDay&"'" &_
								 " Group By S.ComputeMode "
						'Response.Write sQuery &"<br><br>"												
						
						Objrs2.Open sQuery,Con
						IF Not Objrs2.EOF Then
							dSchTot = 0
							dAddTot = 0
							dSubTot = 0
							Do While Not Objrs2.EOF
								IF CStr(Objrs2(1)) = "+" Then
									dAddTot = cdbl(dAddTot) + cdbl(Objrs2(0))
								Else
									dSubTot = cdbl(dSubTot) + cdbl(Objrs2(0))
								End IF
								Objrs2.MoveNext

							Loop
							
							dSchTot = CDbl(dAddTot) - CDbl(dSubTot)						
							
					%>
							<td class="ExcelDisplayCell" align="right"><input type="text" readonly name="txtCurrVal<%=sCatCode%>?<%=iBSSubID%>?<%=iBSSubSubID%>" size="15" value="<%=FormatNumber(dSchTot,2,,,0)%>" style="text-align:right" maxlength="13" class="FormElemRead"></td>
					<%	Else %>
							<td class="ExcelDisplayCell" align="right">&nbsp;
							<textarea id=textarea1 name=textarea1><%=sQuery%></textarea>
							</td>
					<%	End IF
						Objrs2.Close
						
					Elseif CStr(sEntTy) = "A" and CStr(iAccHead) <> "0" Then 'If is is and Account Head Type
					%>
						<td class="ExcelDisplayCell" align="right">
							<input type="text" readonly name="txtCurrVal<%=sCatCode%>?<%=iBSSubID%>?<%=iBSSubSubID%>" size="15" value="<%=FormatNumber(GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay),2,,,0)%>" style="text-align:right" maxlength="13" class="FormElemRead"></td>
					<%
					Else
					%>
						<td class="ExcelDisplayCell" align="right">&nbsp;</td>
					<%
					End IF %>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right">0.00</td>
				</tr>

<%
		Objrs1.MoveNext
		Loop
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
                                                <input type="button" value="BalanceSheet View" name="But_View" class="ActionButtonX" onClick="ViewBalanceSheet()">
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
