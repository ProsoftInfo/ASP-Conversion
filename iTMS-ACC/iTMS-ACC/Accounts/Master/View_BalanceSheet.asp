
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>

<%
	'Program Name				:	View_BalanceSheet.asp.asp
	'Module Name				:	Accounts (Masters- Configure Balance Sheet - BS Setup)
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
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/PurPopulate.asp" -->
<%
Const iPageSize = 7

Dim iCurrentPage,iTotPage,iPageCtr,lnPage,CurrPage
Dim iLineCtr,nMaxLinesPrint
Dim rsSch,Objrs1,Objrs2,Objrs3
Dim sQuery,iCtr,sOrgID,sCatCode,sFinYr,iAccHead,nTempCtr
Dim dClosing,iBSID,iBSSubID,iBSSubSubID,sLastDay,dPrvVal,sSchName,sEntTy
Dim dAddTot,dSubTot,dSchTot,dPreYrVal,sPreDate,nTot1,nTot2,nData

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
<HTML><HEAD><TITLE>iTMS - View BalanceSheet</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<script language="vbscript">

</script>
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

sQuery = "Select BSHeadID,BSHeading,Hierarchy From ACC_M_BSSetupHeads  Where FinYear='"&sFinYr&"' Order by Hierarchy "		 
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
		<td align="left" class="ExcelDisplayCell" colspan="6"><b><%=rsSch("BSHeading")%></b></td>
	</tr>
	<tr>
	<td align="center" class="ExcelHeaderCell"><p>Heading Name</td>
	<td align="center" class="ExcelHeaderCell">As on <%=Replace(sLastDay,"/",".")%></td>
	<td align="center" class="ExcelHeaderCell">As on <%=Replace(sPreDate,"/",".")%></td>
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

		dPrvVal = 0
		sQuery = "Select BSSubHeadValue From ACC_T_BSACDetail Where BSHeadID = "&sCatCode&" "&_
				 "and BSSubID = "&iBSSubID&" and BSSubSUbID = "&iBSSubSubID&" and "&_
				 "OrganisationCode = '"&sOrgID&"' and FinYear = '"&sFinYr&"' "&_
				 "and Convert(Varchar,AsOnDate,103) = '"&sLastDay&"' "

		'Response.Write sQuery &"<br><br>"
		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dPrvVal = Objrs2(0)
		End IF
		Objrs2.Close

		dPreYrVal = 0
		
		sQuery = "Select isNull(BSSubHeadValue,0) From ACC_T_BSACDetail Where BSHeadID = "&sCatCode&" "&_
				 "and BSSubID = "&iBSSubID&" and BSSubSUbID = "&iBSSubSubID&" and "&_
				 "OrganisationCode = '"&sOrgID&"'  "&_
				 "and Convert(Varchar,AsOnDate,103) = '"&sPreDate&"' "


		Objrs2.Open sQuery,Con
		IF Not Objrs2.EOF Then
			dPreYrVal = Objrs2(0)
		End IF
		Objrs2.Close
		
		

%>
				<tr>
					<td class="ExcelDisplayCell">
					<%IF CStr(iBSSubSubID) <> "0" Then
						Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
					End IF
					Response.Write(Objrs1("BSSubHeadingName"))%></td>
					<%IF CStr(sEntTy) = "D" Then 'IF it is an Data Entry the User Will enter the Value 
						nTot1 = CDbl(nTot1) + CDbl(dPrvVal)
					%>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dPrvVal,2,,,0)%></td>
					<%ElseIF CStr(sEntTy) = "S" Then 'If is is and Scheduled Type
						sQuery = "Select isNull(S.ApplicableACHeadCode,0),S.ComputeMode,isNull(S.ScheduleSubHeadValue,0) "&_
								 " From Acc_T_ScheduleACDetail S, ACC_T_BSACDetail P " &_
								 " Where P.BSHeadID = "&sCatCode&" and P.BSSubID = "&iBSSubID&" "&_
								 " and P.BSSubSubID = "&iBSSubSubID&" and P.ScheduleID = S.ScheduleID "&_
								 " and P.ScheduleSubID = S.ScheduleSubID  "&_
								 " and S.OrganisationCode = P.OrganisationCode and S.FinYear = P.FinYear "&_
								 " and P.OrganisationCode = '"&sOrgID&"' and P.FinYear = '"&sFinYr&"' "&_
								 " and Convert(Varchar,S.AsOnDate,103) = '"&sLastDay&"' " &_
								 " and P.ScheduleSubSubID = S.ScheduleSubSubID " & _
								 " Order By S.ComputeMode "
								 
								 'blocked on Jan 30,2012
								 '" and Convert(Varchar,P.AsOnDate,103) = '"&sLastDay&"'" &_
								 
						'Response.Write sQuery &"<br><br>"
						Objrs2.Open sQuery,Con
						IF Not Objrs2.EOF Then
							dSchTot = 0
							dAddTot = 0
							dSubTot = 0
							Do While Not Objrs2.EOF
								if trim(objrs2(0)) = "0" then
									IF CStr(Objrs2(1)) = "+" Then
										dAddTot = cdbl(dAddTot) + cdbl(Objrs2(2))
									Else
										dSubTot = cdbl(dSubTot) + cdbl(Objrs2(2))
									End IF
								else
									'Response.Write "<p>iAccHead =  " & iAccHead
									'Response.Write "<p>sLastDay =  " & sLastDay
									dClosing = GetDayOpeningCreatedForPLBS(sOrgID,objrs2(0),sLastDay)
									'Response.Write "<p>dClosing =  " & dClosing
								
									IF CStr(Objrs2(1)) = "+" Then
										dAddTot = cdbl(dAddTot) + cdbl(dClosing)
									Else
										dSubTot = cdbl(dSubTot) + cdbl(dClosing)
									End IF
								end if 'if trim(objrs2(0)) = "0" then
								
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
					End IF %>
					
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
    												<SELECT class="FormElem" onChange="Paginate(this(this.selectedIndex).value)" id=select1 name=select1>
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
