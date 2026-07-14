<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>

<%
	'Program Name				:	View_ProfitAndLoss_Acc.asp
	'Module Name				:	Accounts (Masters- Configure P&L Sheet - PL Setup)
	'Author Name				:	KalaiSelvi R
	'Created On					:	Jan 20,2012
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
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/PurPopulate.asp"-->
<%

Const iPageSize = 7

Dim rsSch,Objrs1,Objrs2,Objrs3
Dim sQuery,iCtr,sOrgID,sCatCode,sFinYr,iAccHead,nTempCtr
Dim dClosing,iPLID,iPLSubID,iPLSubSubID,sLastDay,dPrvVal,sSchName,sEntTy
Dim dAddTot,dSubTot,dSchTot,dPreYrVal,sPreDate,nData,nTot1,nTot2

Dim iCurrentPage,iTotPage,iPageCtr,lnPage,CurrPage
Dim iLineCtr,nMaxLinesPrint

nMaxLinesPrint = 20 '67

Set rsSch = Server.CreateObject("ADODB.RecordSet")
Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
Set Objrs2 = Server.CreateObject("ADODB.RecordSet")

sOrgID = trim(request("OrgID"))
sLastDay = Trim(Request("ForTheDate"))

sFinYr = Session("FinPeriod")



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
<HTML><HEAD><TITLE>iTMS - View Profit and Loss A/C</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/PrintWindow.js"></SCRIPT>
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


sQuery = "Select PLHeadID,PLHeading From ACC_M_PLSetupHeads Where FinYear='"&sFinYr&"'  "

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
		<td align="left" class="ExcelDisplayCell" colspan="4"><b><%=rsSch("PLHeading")%></b></td>
	</tr>
	<tr>
	<td align="center" class="ExcelHeaderCell"><p>Heading Name</td>
	<td align="center" class="ExcelHeaderCell">As on 31.03.<%=Right(sFinYr,4)%></td>
	<td align="center" class="ExcelHeaderCell">As on 31.03.<%=Left(sFinYr,4)%></td>
	</tr>
<%
	'Taking The SchSubID and Sch SubSUbID,Name and EntryType for the Selected Schedule
	sQuery = "Select PLSubID,PLSubSubID,PLSubHeadingName,isNull(EntryType,'N') EntryType From  "&_
			 "ACC_M_PLSetupSubHeads Where PLHeadID = "&sCatCode&" and FinYear = '"&sFinYr&"' Order By Hierachy "

	With Objrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End With
	Set Objrs1.ActiveConnection = Nothing

	Do WHile Not Objrs1.EOF
		iPLSubID = Objrs1("PLSubID")
		iPLSubSubID = Objrs1("PLSubSubID")
		sEntTy = Objrs1("EntryType")

		sQuery = "Select isNull(ApplicableACHeadCode,0) From ACC_T_PLACDetail Where PLHeadID = "&sCatCode&" "&_
				 " and PLSubID = "&iPLSubID&" and PLSubSUbID = "&iPLSubSubID&"  "&_
				 " and OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' "


		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			iAccHead = Objrs2(0)
		Else
			iAccHead = 0
		End IF
		Objrs2.Close

		sQuery = "Select isNull(PLSubHeadValue,0) From ACC_T_PLACDetail Where PLHeadID = "&sCatCode&" "&_
				 "and PLSubID = "&iPLSubID&" and PLSubSUbID = "&iPLSubSubID&" and "&_
				 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' "&_
				 "and Convert(Varchar,AsOnDate,103) = '"&sLastDay&"' "

		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dPrvVal = Objrs2(0)
		Else
			dPrvVal = 0
		End IF
		Objrs2.Close

		sQuery = "Select isNull(PLSubHeadValue,0) From ACC_T_PLACDetail Where PLHeadID = "&sCatCode&" "&_
				 "and PLSubID = "&iPLSubID&" and PLSubSUbID = "&iPLSubSubID&" and "&_
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
					<%IF CStr(iPLSubSubID) <> "0" Then
						Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
					End IF
					Response.Write(Objrs1("PLSubHeadingName"))%></td>
					<%IF CStr(sEntTy) = "D" Then 'IF it is an Data Entry the User Will enter the Value 
						nTot1 = CDbl(nTot1) + CDbl(dPrvVal)
					%>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
					<%ElseIF CStr(sEntTy) = "S" Then 'If is is and Scheduled Type
						sQuery = "Select isNull(SUM(S.ScheduleSubHeadValue),0),S.ComputeMode From Acc_T_ScheduleACDetail S, "&_
								 "ACC_T_PLACDetail P Where P.PLHeadID = "&sCatCode&" and P.PLSubID = "&iPLSubID&" and  "&_
								 "P.PLSubSubID = "&iPLSubSubID&" and P.ScheduleID = S.ScheduleID and  "&_
								 "P.ScheduleSubID = S.ScheduleSubID and P.ScheduleSubSubID = S.ScheduleSubSubID "&_
								 "and S.OrganisationCode = P.OrganisationCode and S.FinYear = P.FinYear "&_
								 "and P.OrganisationCode = '"&sOrgID&"' and P.FinYear = '"&sFinYr&"'  "&_
								 "and Convert(Varchar,P.AsOnDate,103) = '"&sLastDay&"' Group By S.ComputeMode "



						'Response.Write sQuery
						Objrs2.Open sQuery,Con
						IF Not Objrs2.EOF Then
							dSchTot = 0
							dAddTot = 0
							dSubTot = 0
							Do While Not Objrs2.EOF
								IF CStr(Objrs2(1)) = "+" Then
									dAddTot = Objrs2(0)
								Else
									dSubTot = Objrs2(0)
								End IF
								Objrs2.MoveNext

							Loop
							dSchTot = CDbl(dAddTot) - CDbl(dSubTot)
							
							nTot1 = CDbl(nTot1) + CDbl(dSchTot)
					%>
							<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dSchTot,2,,,0)%></td>
					<%	Else %>
							<td class="ExcelDisplayCell" align="right">&nbsp;</td>
					<%	End IF
						Objrs2.Close

					Elseif CStr(sEntTy) = "A" and CStr(iAccHead) <> "0" Then 'If is is and Account Head Type
						nData = GetDayOpeningCreatedForPLBS(sOrgID,iAccHead,sLastDay)
						nTot1 = CDbl(nTot1) + CDbl(nData)
					%>
						<td class="ExcelDisplayCell" align="right">
							<%=FormatNumber(nData,2,,,0)%>
						</td>
					<%
					Else
					%>
						<td class="ExcelDisplayCell" align="right">&nbsp;</td>
					<%
					End IF 
					
					nTot2 = CDbl(nTot2) + CDbl(dPreYrVal)
					%>
					
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPreYrVal,2,,,0)%></td>
				</tr>

<%
		Objrs1.MoveNext
		Loop
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
		<td align="left" class="ExcelDisplayCell" colspan="6">&nbsp;</td>
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
