<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLCreate_Edit_AccHeadDet.asp
	'Module Name				:	ACCOUNTS (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 24,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	Code
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
<!--#include virtual="/include/accpopulate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
	Dim iAccHead,sAccHeadName,sOrgId,oDom,Root,dCreditLimit,sAccGroupCode,sAccGroupName
	dim objRs,sQuery,objrs2,sUnitarr(),iCtr,sAmenType,sDisArr(),sCheckArr,sDisUnits
	dim bIUT,bCostCenter,bAnalytical,bContra,bMemorandum,sTransLimit,bSubLedger
	dim bTDS,bSummary,sAccHeadShortName,sAppUsed,sFrqBooks,sAppUsedCode,sFrqBooksCode
	Dim iOpenAmt,iCloseAmt,iTransDr,iTransCr,sSelGLSummBook,sAction,sCDIndication
	Dim sCostCenter,sAnalyticalCode,objDOM,objrs1,sSummPostBook,objFSO,dDRTotal,dCRTotal

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
	Set Root = oDOM.documentElement
	dCreditLimit=Root.childNodes.item(0).text

Response.Write "<font color=red>"
	iAccHead = trim(Request("hHeadValue"))
	sAccHeadName = trim(Request("hHeadName"))
'	sOrgId= trim(Request("selUnitId"))
'	Response.Write "iAccHead="& iAccHead
	
		'Response.Write "iAccHead = "& iAccHead
	set objRs = Server.CreateObject("ADODB.Recordset")
	Set objrs1 = Server.CreateObject("ADODB.Recordset")
	set objRs2 = Server.CreateObject("ADODB.Recordset")

	

	if trim(iAccHead)="" or Trim(iAccHead)="0" then
		sAction = "Create"
	else
		sAction = "Edit"
		sAccGroupCode = trim(Request("GCode"))
	    sAccGroupName = trim(Request("GName"))
	    if Trim(sAccGroupCode)="" or IsNull(sAccGroupCode) then 
	        sQuery = "Select AccountsGroupCode,AccountsGroupName from Acc_M_AccountGroups "&_
	                 " where AccountsGroupCode in( Select AccountsGroupCode from VwOrgGLHeads "&_
	                 " where AccountHead = "& iAccHead &")"
	        objRs.Open sQuery,con
	        if not objRs.EOF then
	            sAccGroupCode = objRs(0)
	            sAccGroupName = objRs(1)
	        end if
	        objRs.Close 
	    end if
	end if

	if iAccHead<>"" then
		sQuery = "Select Top 1 M.AccountsGroupName,V.AccountDescription,V.AccountHeadCode From  "&_
				 "Acc_M_AccountGroups M,VwOrgGLHeads V Where V.AccountsGroupCode = "&_
				 "M.AccountsGroupCode and V.AccountHead = "&iAccHead

		'Response.write sQuery
		'sQuery ="SELECT  AccountHeadCode FROM Acc_M_GLAccountHead where  AccountHead="&iAccHead
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing

		if not objRs.EOF then
			sAccGroupName = objRs(0)
			sAccHeadName = objRs(1)
			sAccHeadShortName = objRs(2)
		end if
		objRs.Close
		'Response.write sAccGroupName &"  " & sAccGroupCode


		sQuery= "Select ApplicationName,ApplicationCode from Ms_Applications where ApplicationCode in "&_
				"(Select AvailableInAppln from Acc_R_GLAccApplications where "&_
				"AccountHead="&iAccHead&") order by ApplicationName"
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		Do While Not objRs.EOF
			sAppUsed = sAppUsed &","&objRs(0)
			sAppUsedCode = sAppUsedCode&":"&objRs(1)
			objRs.MoveNext
		Loop
		objRs.Close
		IF CStr(sAppUsed) <> "" Then
			sAppUsed = Right(sAppUsed,Len(sAppUsed)-1)
		End IF

		sQuery = "Select Distinct M.BookName,M.BookCode From Acc_M_DayBooks M,Acc_R_GLAccFrequentlyUsed R "&_
				 "Where M.BookCode = R.BookCode and "&_
				 "R.AccountHead = "&iAccHead&" "
'Response.Write sQuery
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection = nothing
		Do While Not objRs.EOF
			sFrqBooks = sFrqBooks &","&objRs(0)
			sFrqBooksCode = sFrqBooksCode &":"&objRs(1)
			objRs.MoveNext
		Loop
		objRs.Close
		IF CStr(sFrqBooks) <> "" Then
			sFrqBooks = Right(sFrqBooks,Len(sFrqBooks)-1)
		End IF
	end if


	if iAccHead<>"" then
		'**************************************
		' XML Creation
		Dim ndRoot,ndChild,newElem,newElem1

		Set ndRoot = objDOM.createElement("Root")
		objDOM.appendChild ndRoot

		Set ndChild = objDOM.createElement("OpeningMonthYear")
		ndChild.text = getFromFinYear
		ndRoot.appendChild ndChild

		Set ndChild = objDOM.createElement("ClosingMonthYear")
		ndChild.text = getToFinYear
		ndRoot.appendChild ndChild

		Set ndChild = objDOM.createElement("GroupCode")
		ndChild.setAttribute "Name", sAccGroupName
		ndChild.Text= sAccGroupCode
		ndRoot.appendChild ndChild

		Set ndChild = objDOM.createElement("Description")
		ndChild.Text= sAccHeadName
		ndRoot.appendChild ndChild

		Set ndChild = objDOM.createElement("ShortName")
		ndChild.Text= sAccHeadShortName
		ndRoot.appendChild ndChild

		Set ndChild = objDOM.createElement("AccHeadNo")
		ndChild.Text= iAccHead
		ndRoot.appendChild ndChild

		Set ndChild = objDOM.createElement("Units")
		ndRoot.appendChild ndChild

		sQuery = "Select OUDefinitionID,(Select OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID =OP.OUDefinitionID ) OrgName,OpeningAmount,ClosingAmount,OpeningCDIndication From Acc_T_GLAccOpeningAmt OP Where AccountHead = "& iAccHead
		objrs2.Open sQuery,con
		if not objrs2.EOF then
			do while not objrs2.EOF
				Set newElem = objDOM.createElement("UN")
					newElem.setAttribute "Code", objrs2(0)
					newElem.setAttribute "Name", objrs2(1)
					newElem.setAttribute "OpBalance",objrs2(2)
					newElem.setAttribute "OpCRDR",objrs2(4)
					ndChild.appendChild newElem

					sQuery = "Select OUDefinitionID,InterUnitTransact,SummaryPosting,SubLedger,CostCenterExists,"&_
							 "AnalyticalHeadExists,EligibleForContras,EligibleForTDS,MemorandumAccount,AllowTransactions"&_
							 " from Acc_R_OrgGLAccountHead where OUDefinitionID = "& objrs2(0) &" and AccountHead = "& iAccHead
					objRs.Open sQuery,con
					if not objRs.EOF then
						Set newElem1 = objDOM.createElement("IUT")
						newElem1.setAttribute "Flag",objrs(1)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("SummaryPosting")
						newElem1.setAttribute "Flag",objrs(2)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("SubLedger")
						newElem1.setAttribute "Flag",objrs(3)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("CostCenter")
						newElem1.setAttribute "Flag",objrs(4)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("Analytical")
						newElem1.setAttribute "Flag",objrs(5)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("Contra")
						newElem1.setAttribute "Flag",objrs(6)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("TDS")
						newElem1.setAttribute "Flag",objrs(7)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("Memorandum")
						newElem1.setAttribute "Flag",objrs(8)
						newElem.appendChild newElem1

						Set newElem1 = objDOM.createElement("CashTrans")
						newElem1.setAttribute "Flag",objrs(9)
						newElem.appendChild newElem1

						sQuery = "Select BookCode from Acc_M_GLSummaryApp where AccountHead = "& iAccHead &" and OUDefinitionID = '"& objRs2(0)  &"'"
						objrs1.Open sQuery,con
						if not objrs1.EOF then
							do while not objrs1.EOF
								sSummPostBook = sSummPostBook &":"& objrs1(0)
								objrs1.MoveNext
							loop
						end if
						objrs1.Close

						Set newElem1 = objDOM.createElement("SummaryPostBook")
						if trim(sSummPostBook)<>"" then
							newElem1.setAttribute "BookCodes", mid(sSummPostBook,2)
						else
							newElem1.setAttribute "BookCodes", ""
						end if
						newElem.appendChild newElem1
					end if
					objRs.Close
				objrs2.MoveNext
			loop
		end if
		objrs2.Close

		sQuery = "Select ApplicationCode,ApplicationName from MS_Applications where ApplicationCode in "&_
		" (Select AvailableInAppln from Acc_R_GLAccApplications where AccountHead = "& iAccHead &" Group by AvailableInAppln)"
		objrs2.Open sQuery,con
		if not objrs2.EOF then
			Set ndChild = objDOM.createElement("Applications")
				ndRoot.appendChild ndChild
				do while not objrs2.EOF
					Set newElem1 = objDOM.createElement("APP")
						newElem1.setAttribute "Code", objrs2(0)
						newElem1.text=objrs2(1)
						ndChild.appendChild newElem1
					objrs2.MoveNext
				loop
		end if
		objrs2.Close


'			IF CStr(sAppCheck) = "Y" Then
'
'				for iCounter=0 to UBound(arrApplication)
'				next
'			End IF
'
'			IF CStr(sBookCheck) = "Y" Then
'				Set newElem = GLHeadData.createElement("Books")
'				if trim(document.formname.hBookCode.value)="" then
'					newElem.setAttribute "Count", 0
'					iBookCount=0
'				elseif 	UBound(arrBooks)=0 then
''					newElem.setAttribute "Count", 1
'					iBookCount=1
'				else
'					newElem.setAttribute "Count", UBound(arrBooks)
'					iBookCount=UBound(arrBooks)
'				end if
'				Root.appendChild newElem
'			End IF

'			IF CStr(sBookCheck) = "Y" Then
'				for iCounter=0 to UBound(arrBooks)
'					Set newElem1 = GLHeadData.createElement("BK")
'					newElem1.setAttribute "Code", trim(arrBooks(iCounter))
'					newElem1.setAttribute "Name", trim(arrBooksName(iCounter))
'					newElem.appendChild newElem1
'				next
'			End IF

			sQuery = "Select OUDefinitionID,(Select CCGroupCode from VwOrgCostCenter where CostCenterHead = C.CostCenterHead) GroupCode ,CostCenterHead from Acc_R_OrgGLCostCentre C where AccountHead = "& iAccHead
			objrs2.Open sQuery,con
			if not objrs2.EOF then
				Set ndChild = objDOM.createElement("CostCenter")
					ndRoot.appendChild ndChild
					do while not objrs2.EOF
						Set newElem1 = objDOM.createElement("CC")
						newElem1.setAttribute "UNCode",objrs2(0)
						newElem1.setAttribute "GRCode",objrs2(1)
						newElem1.setAttribute "CCCode",objrs2(2)
						ndChild.appendChild newElem1
						objrs2.MoveNext
					loop
			end if
			objrs2.Close

			sQuery = "Select OUDefinitionID,Cast(AnalyticalCode as varchar),AHGroupCode from Acc_R_OrgGLAnalytical where AccountHead = "& iAccHead
			objrs2.Open sQuery,con
			if not objrs2.EOF then
					set ndChild = objDOM.createElement("Analytical")
					ndRoot.appendChild ndChild
				do while not objrs2.EOF
					Set newElem1 = objDOM.createElement("AN")
						newElem1.setAttribute "UNCode", objrs2(0)
						newElem1.setAttribute "Code", objrs2(1)
						newElem1.setAttribute "GRCode", objrs2(2)
						ndChild.appendChild newElem1
					objrs2.MoveNext
				loop
			end if
			objrs2.Close

		objDOM.Save Server.MapPath("../temp/master/GLAccount_New_Head_"&Session.SessionID&".xml")

	''XML Finish
	''********************
	end if 'if iAccHead<>"" then

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<% if objFSO.FileExists(Server.MapPath("../temp/master/GLAccount_New_Head_"&Session.SessionID&".xml")) then %>
	<script type="application/xml" data-itms-xml-island="1" id="GLHeadData" data-src="<%="../temp/master/GLAccount_New_Head_"&Session.SessionID&".xml"%>" ><Root/></script>
<%else%>
	<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root/></script>
<%end if%>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root /></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "glAccountHeadDetails" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="init()">


	<form method="POST" name="formname">
	<input type="Hidden" name="hAccCode" value="<%=iAccHead%>" >
	<input type="Hidden" name="hAccName" value="<%=sAccHeadName%>" >
	<input type="Hidden" name="hGCode" value="<%=sAccGroupCode%>" >
	<input type="Hidden" name="hGName" value="<%=sAccGroupName%>" >
	<input type="Hidden" name="hAppCode" value="<%=sAppUsedCode%>" >
	<input type="Hidden" name="hBookCode" value="<%=sFrqBooksCode%>" >
	<input type="Hidden" name="hAction" value="<%=sAction%>">
	<input type="Hidden" name="hOpenYear" value="<%=getFromFinYear%>">
	<input type="Hidden" name="hCloseYear" value="<%=getToFinYear%>">
	<input type="Hidden" name="hCostCenter" value="<%=sCostCenter%>">
	<input type="Hidden" name="hAnalyticalCode" value="<%=sAnalyticalCode%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">
				<%
					if trim(sAccGroupName)<>"" then
						Response.Write "GL Account Head Amendment"
					else
						Response.Write "GL Account Head"
					end if 'if trim(sAccGroupName)<>"" then
				%>
				</p>
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<!--<td class="TabCell" valign="bottom" width="70">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onmouseover="tabrollover(this)" onmouseout="tabrollout(this)">
											<tr>
												<td width="100%" align="center">Header
												</td>

											</tr>

										</table>
									</td>-->
									<td class="TabCurrentCell" valign="bottom" align="center" width="70">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
											<tr>
												<td width="100%" align="center">Details
												</td>
											</tr>

										</table>
									</td>
									<!--<td class="TabCell" valign="bottom" align="center" width="95">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Cost Center
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="120">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Analytical Head</td>
										</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Books</td>
										</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
									<tr>
										<td align="center">Unit
										</td>
									</tr>
								  </table>
								</td>-->
									<td class="TabCellEnd" valign="bottom" align="left">&nbsp;
									</td>
								</tr>

							</table>
						</td>
					</tr>

					<tr>
						<td class="TabBody">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<!--tr>
									<td align="center">
									</td>
									<td valign="top" width="100%">
										<table border="0" width="100%" cellspacing="0" cellpadding="0" class="ToolBarTable">
											<tr>
												<td width="40" align="center" valign="middle" class="ToolBarCell" onclick="toolClick(this)" onmouseover="toolrollover(this)" onmouseout="toolrollout(this)">
													<span style="cursor: pointer" title="New">
													<p align="center"><font face="Wingdings" size="5">2</font></p>
													</span>
												</td>
												<td align="center" class="ToolBarCell">&nbsp;
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr-->

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td class="FieldCell">Account Group Name
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly" id="AccGroupName" ><%Response.Write sAccGroupName%> </span>&nbsp;&nbsp;&nbsp
													<%
														if trim(sAccGroupName)="" then
															Response.Write "<img border='0' src='../../assets/images/iTMS Icons/EntryIcon.gif' alt='Select Account Group' onClick='SelectAccHead()'>"
														end if 'if trim(sAccGroupName)<>"" then
													%>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Account Description
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtAccname" size="50" class="FormElem" value="<%=sAccHeadName%>" maxlength="100">
 													&nbsp;
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Account Short Name
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtAccShortName" size="30" class="FormElem" value="<%=sAccHeadShortName%>" maxlength="10">
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td valign="top" width="100%">
										<div class="frmBody" id="frm3" style="width: 585; height:300;">
											<table border="0" cellspacing="1" class="ExcelTable" width="550">
												<tr>
													<td class="ExcelHeaderCell" align="center" >Applicable <br>in Units
													</td>
													<%
														Dim sTemp1,sTemp2,iTmp
														iCtr = 0
														with objRs
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
															.ActiveConnection = con
															.Open
														end with
														'Response.Write objRs.Source
														redim sUnitarr(objRs.RecordCount)
														Redim sDisArr(objRs.RecordCount)

														set objRs.ActiveConnection = nothing

														Do while Not objRs.EOF
															 sUnitarr(iCtr) = objRs(0)
															 'Response.Write " <p> Unit = "&  sUnitarr(iCtr)
															 sOrgId = sOrgId &","&objRs(0)
															 sAmenType = ""

															 Response.Write "<td class=ExcelHeaderCell align=center >"


														if iAccHead<>"" then

															 sQuery = "Select Top 1 isNull(C.AccountHead,0) From Acc_T_CreatedVoucherHeader C "&_
																	  "Where C.OUDefinitionID = '"&objRs(0)&"'  "&_
																	  "and isNull(C.AccountHead,0) = "&iAccHead&" "



															With objrs2
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															 End With
															 Set objrs2.ActiveConnection = Nothing
															 IF Not objrs2.EOF Then
																sAmenType = "disabled"
															 Else
																sAmenType = ""
															 End If
															 objrs2.Close

														end if 'if iAccHead<>"" then


															 '====================== IF The Selected Account Head is Not been used in the Transaction Then Checking
															 '====================== For Opening and Closing Account Head is been for 0
															 IF Len(sAmenType) = 0 Then
																if iAccHead<>"" then
																	sQuery = "Select OpeningAmount,ClosingAmount From Acc_T_GLAccOpeningAmt "&_
																			 "Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&objRs(0)&"' "

																'	Response.Write "sQuery = "  & sQuery

																	objrs2.Open sQuery,Con
																	IF Not objrs2.EOF Then
																		iOpenAmt = objrs2(0)
																		iCloseAmt = objrs2(1)
																	Else
																		iOpenAmt = 0
																		iCloseAmt = 0
																	End IF
																	objrs2.Close

																	sQuery = "Select MonthDrAmount,MonthCrAmount From Acc_T_GLAccTransactAmt "&_
																			 "Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&objRs(0)&"' "

																	objrs2.Open sQuery,Con
																	IF Not objrs2.EOF Then
																		iTransDr = objrs2(0)
																		iTransCr = objrs2(1)
																	Else
																		iTransCr = 0
																		iTransDr = 0
																	End IF
																	objrs2.Close

																	'Response.Write iOpenAmt &" "& iCloseAmt &" " & iTransCr &" " & iTransDr &"<br>"
																	IF CStr(iOpenAmt) = "0" and CStr(iCloseAmt) = "0" and CStr(iTransCr) = "0" and CStr(iTransDr) = "0" Then
																		sAmenType = ""
																	Else
																		sAmenType = "disabled"
																	End IF
																end if 'if iAccHead<>"" then
															 End IF

															 IF Len(sAmenType) = 0 Then
																sDisArr(iCtr) = objRs(0)&"B"
															 else
																sDisArr(iCtr) = objRs(0)&"D"
															 End IF

															 '=========== Check For Opening,Closing, Trans DR, Trans Cr Checking is Over ============
														if iAccHead<>"" then

																	 sQuery = "Select OUDefinitionID From Acc_R_OrgGLAccountHead Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&objRs(0)&"' and AmendmentExists = '0' "
																	'Response.Write sQuery

																	 With objrs2
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	 End With
																	 Set objrs2.ActiveConnection = Nothing

																IF Not objrs2.EOF Then
																	sTemp1 = objRs(1)
																	'Redim sTemp2(1)
																	sTemp2 = Split(objRs(1),"-")
																	IF Len(sAmenType) <> 0 Then
																		sDisUnits = sDisUnits&":"&objRs(0)&"?"&objRs(1)
																	End IF
																%>
																<Input type="checkbox" class="FormElem" name="chkUnit<%=objRs(0)%>" value="<%=objRs(0)%>:<%=Replace(objRs(1),","," ")%>" checked <%=sAmenType%> onClick="CheckDetails('<%=objRs(0)%>',this,'<%=iAccHead%>')">
																<%else%>
																<Input type="checkbox" class="FormElem" name="chkUnit<%=objRs(0)%>" value="<%=objRs(0)%>:<%=Replace(objRs(1),","," ")%>" onClick="CheckDetails('<%=objRs(0)%>',this,'<%=iAccHead%>')" >
																<%end if
																objrs2.Close
														else
															%>
																<Input type="checkbox" class="FormElem" name="chkUnit<%=objRs(0)%>" value="<%=objRs(0)%>:<%=Replace(objRs(1),","," ")%>" onClick="CheckDetails('<%=objRs(0)%>',this,'<%=iAccHead%>')" >
															<%
														end if 'if iAccHead<>"" then
														'For iTmp = 0 to UBound(sTemp2)
														'	Response.Write sTemp2(iTmp) & "<br>"

														'Next
														Response.Write Replace(Trim(objRs(1)),","," ")

													%>
													    <input type="hidden" name="hUnitNameZ<%=trim(objrs(0))%>" value="<%=Replace(trim(objRs(1)),","," ") %>">

													</td>
													<%
														objRs.MoveNext
														iCtr = iCtr + 1
														loop
														objRs.Close
													'	Response.Write "sOrgId = "& sOrgId
														sOrgId = Right(sOrgId,Len(sOrgId)-1)
													'	Response.Write "sOrgId = "& sOrgId
														'Response.Write sAmenType
													%>
													</tr>
													<input type="Hidden" name="hOrgid" value="<%=sOrgId%>" >

												<tr>
													<td class="ExcelDisplayCell">Opening Balance</td>
													<%
													For iCtr = 0 to UBound(sUnitarr)-1

														if iAccHead<>"" then
																'	sQuery = "Select OpeningAmount,ClosingAmount,OpeningCDIndication From Acc_T_GLAccOpeningAmt "&_
																'			 "Where AccountHead = "&iAccHead&" and OUDefinitionID = '"& sUnitarr(iCtr) &"' "

																    sQuery = "Select SUM(T.OpeningAmount),T.OpeningCDIndication From Acc_T_PartyOpeningAmt T, "&_
	                                                                         "Acc_R_OrgPartyType R,VWOrgParty M Where R.OUDefinitionID = T.OUDefinitionID and "&_
	                                                                         "R.PartyType = T.PartyType and R.PartySubType = T.PartySubType and "&_
	                                                                         "R.AccountHead = "&iAcchead&" and R.OUDefinitionID = '"&sUnitarr(iCtr)&"' and  "&_
	                                                                         "T.OpeningMonthYear = '"&Trim(getFromFinYear)&"' and "&_
	                                                                         "M.PartyCode = T.PartyCode and M.OUDEFINITIONID = '"&sUnitarr(iCtr)&"' and "&_
	                                                                         "R.PartyType = M.PartyType and R.PartySubType = M.PartySubType "&_
	                                                                         "Group By T.OpeningCDIndication "

																	'Response.Write "sQuery = "  & sQuery

																	objrs2.Open sQuery,Con
																	dDRTotal= 0
																	dCRTotal= 0
																	IF Not objrs2.EOF Then
																	    do while not objrs2.EOF
																	        if trim(objrs2(1))="C" then
																	            dCRTotal = objrs2(0)
																	        else
																	            dDRTotal = objrs2(0)
																	        end if
																	       ' Response.Write "dCRTotal = "& dCRTotal
																	        'Response.Write "dDRTotal = "& dDRTotal
																	        objrs2.MoveNext
																	    loop
																	End IF
																	objrs2.Close
														end if ' if iAccHead <>"" then

														if CDbl(dCRTotal)> cdbl(dDRTotal) then
														    iOpenAmt = CDbl(dCRTotal)- cdbl(dDRTotal)
														    sCDIndication  = "C"
														else
														    iOpenAmt = cdbl(dDRTotal) - CDbl(dCRTotal)
														    sCDIndication  = "D"
														end if

													%>
													<td class="ExcelFieldCell">MonthYear
													<input type=text name=txtOpenYear value="<%=getFromFinYear%>" class="FormElem" size=7><br>
														Rs.&nbsp;<input type="text" name="txtOpenBal<%=Trim(sUnitarr(iCtr))%>" maxlength="12" size="15" class="FormElem" style="text-align: Right" value="<%=iOpenAmt%>" >
														<%if sCDIndication = "D" then %>
															<input type="radio" value="D" name="optOpenCD<%=Trim(sUnitarr(iCtr))%>" class="FormElem" checked>Dr <input type="radio" value="C" name="optOpenCD<%=Trim(sUnitarr(iCtr))%>" class="FormElem">Cr
														<%else%>
															<input type="radio" value="D" name="optOpenCD<%=Trim(sUnitarr(iCtr))%>" class="FormElem">Dr <input type="radio" value="C" name="optOpenCD<%=Trim(sUnitarr(iCtr))%>" class="FormElem" checked >Cr
														<%end if %>
													</td>
													</td>
													<%
													Next
													%>
												</tr>
												<tr>
													<td class="ExcelDisplayCell">IUT Account
													</td>
													<%

													Dim iRecCount,sDisType
													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = ""

														if iAccHead<>"" then
															sQuery ="SELECT  InterUnitTransact FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bIUT=objRs(0)
															else
																bIUT = "0"
															end if
															objRs.Close
														end if 'if iAccHead<>"" then
											%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(bIUT) = "1" Then %>
 														<input type="radio" value="1" name="optIUT<%=sUnitarr(iCtr)%>" class="FormElem" Checked  <%=sDisType%> <%=sAmenType%>>
 														Yes
														<input type="radio" value="0" name="optIUT<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														No
 														<%else%>
 														<input type="radio" value="1" name="optIUT<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														Yes
														<!--input type="radio" value="0" name="optIUT<%=sUnitarr(iCtr)%>" class="FormElem"  Checked onClick="IUTAccCheck('<%=sUnitarr(iCtr)%>',this)" <%=sDisType%> <%=sAmenType%>-->
														<input type="radio" value="0" name="optIUT<%=sUnitarr(iCtr)%>" class="FormElem"  Checked <%=sDisType%> <%=sAmenType%>>
 														No
 														<%end if %>
													</td>
													<%Next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Summary posting
													</td>
													<%
													'Response.Write "UBound(sUnitarr) = "& UBound(sUnitarr)

													For iCtr = 0 To UBound(sUnitarr)-1
															sAmenType = ""
														if iAccHead<>"" then

															sQuery ="SELECT SummaryPosting FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead
																'	Response.Write sQuery


															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bSummary=objRs(0)
															else
																bSummary = "0"
															end if
															objRs.Close

																sQuery = "Select AccountHead From Acc_T_CreatedVoucherHeader Where AccountHead = "&iAccHead&" "&_
																		 "and OUDefinitionID = '"&Trim(sUnitarr(iCtr))&"' "

																'Response.Write sQuery &"<br>"
																objRs.Open sQuery,Con
																IF Not objRs.EOF Then
																	sAmenType = "disabled"
																Else
																	sAmenType = ""
																End IF
																objRs.Close
														end if 'if iAccHead<>"" then

														IF Len(Trim(sAmenType)) = 0 Then
															if iAccHead<>"" then
																sQuery = "Select AccUnitAccountHead From Acc_T_CreatedVoucherDetails Where  "&_
																		 "AccUnitAccountHead = "&iAccHead&"  and AccountingUnit = '"&Trim(sUnitarr(iCtr))&"' "
																objRs.Open sQuery,Con
																IF Not objRs.EOF Then
																	sAmenType = "disabled"
																Else
																	sAmenType = ""
																End IF
																objRs.Close
															end if 'if iAccHead<>"" then
														End IF

														sSelGLSummBook = ""
														IF CStr(bSummary) = "1" Then
															sQuery = "Select BookCode From Acc_M_GLSummaryApp Where "&_
																	 "AccountHead = "&iAccHead&" and OUDefinitionID = '"&Trim(sUnitarr(iCtr))&"' "
															With objRs
																.CursorLocation = 3
																.CursorType = 3
																.ActiveConnection = con
																.Source = sQuery
																.Open
															End With
															Set objRs.ActiveConnection = Nothing
															Do While Not objRs.EOF
																sSelGLSummBook = sSelGLSummBook&","&objRs(0)
																objRs.MoveNext
															Loop
															objRs.Close
															sSelGLSummBook = Mid(sSelGLSummBook,2)
														End IF

													%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(bSummary) = "1" Then

														%>
 														<input type="radio" value="1" name="optSummary<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%>>
 														Yes
														<input type="radio" value="0" name="optSummary<%=sUnitarr(iCtr)%>" class="FormElem"   <%=sDisType%> <%=sAmenType%>>
 														No
 														<%
 															else
 														%>
 														<input type="radio" value="1" name="optSummary<%=sUnitarr(iCtr)%>" class="FormElem"  <%=sDisType%> <%=sAmenType%>>
 														Yes
														<input type="radio" value="0" name="optSummary<%=sUnitarr(iCtr)%>" class="FormElem" Checked  <%=sDisType%> <%=sAmenType%>>
 														No
 														<%end if %>
 														&nbsp; <a href="#" onclick="SelVouType(document.formname.optSummary<%=sUnitarr(iCtr)%>,'<%=sUnitarr(iCtr)%>','<%=iAccHead%>'); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Voucher Type"></a>
 														<input type="hidden" name="hSummAppSel<%=sUnitarr(iCtr)%>" value="<%=sSelGLSummBook%>">
 														<input type="hidden" name="hSummAppVal<%=sUnitarr(iCtr)%>" value="">
													</td>
													<%Next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Party Control Account
													</td>
													<%


													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = ""


														if iAccHead<>"" then

															sQuery ="SELECT SubLedger FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bSubLedger=objRs(0)
															else
																bSubLedger="0"
															end if
															objRs.Close
														end if 'if iAccHead<>"" then

													%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(bSubLedger) = "1" Then %>
 														<input type="radio" value="1" name="optSubledger<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%>>
 														Yes
														<input type="radio" value="0" name="optSubledger<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														No
 														<%else%>
 														<input type="radio" value="1" name="optSubledger<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														Yes
														<!--input type="radio" value="0" name="optSubledger<%=sUnitarr(iCtr)%>" class="FormElem" Checked onClick="DisSumm('<%=sUnitarr(iCtr)%>', this)" <%=sDisType%> <%=sAmenType%>-->
														<input type="radio" value="0" name="optSubledger<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%>>

 														No
 														<%end if %>

														&nbsp;&nbsp;<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif"  style="cursor: pointer" alt="Select Party Sub Types" onclick="ControlAccount('<%=sUnitarr(iCtr)%>')">

													</td>
													<%Next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Cost Center
													</td>
													<%

													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = ""

														if iAccHead<>"" then

															sQuery ="SELECT  CostCenterExists FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead

														'Response.Write sQuery
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bCostCenter=objRs(0)
															else
																bCostCenter="0"
															end if
															objRs.Close
														end if
													%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(bCostCenter) = "1" Then %>
 														<input type="radio" value="1" name="optCCZ<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%> onClick="ValidateCostCenterHead(this)" >
 														Yes
														<input type="radio" value="0" name="optCCZ<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%> onClick="ValidateCostCenterHead(this)">
 														No
 														<%else%>
 														<input type="radio" value="1" name="optCCZ<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%> onClick="ValidateCostCenterHead(this)">
 														Yes
														<input type="radio" value="0" name="optCCZ<%=sUnitarr(iCtr)%>" class="FormElem"  Checked <%=sDisType%> <%=sAmenType%> onClick="ValidateCostCenterHead(this)">
 														No
 														<%end if %>
														&nbsp;&nbsp;<img border='0' style="cursor: pointer" id="imgCostCenterZ<%=sUnitarr(iCtr)%>"  src='../../assets/images/iTMS Icons/EntryIcon.gif' alt='Select Cost Center' onClick="PopCostCenter('<%=sUnitarr(iCtr)%>')" disabled >
 														<input type="hidden" name="hSelCostCodeZ<%=sUnitarr(iCtr)%>" value="">
													</td>
													<%next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Analytical Code
													</td>
													<%

													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = ""
														if iAccHead<>"" then

															sQuery ="SELECT  AnalyticalHeadExists FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bAnalytical=objRs(0)
															else
																bAnalytical = "0"
															end if
															objRs.Close
														end if 'if iAccHead<>"" then
													%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(bAnalytical) = "1" Then %>
 														<input type="radio" value="1" name="optAnalZ<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%> onClick="ValidatePopAnalHead(this)" >
 														Yes
														<input type="radio" value="0" name="optAnalZ<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%> onClick="ValidatePopAnalHead(this)" >
 														No
 														<%else%>
 														<input type="radio" value="1" name="optAnalZ<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%> onClick="ValidatePopAnalHead(this)" >
 														Yes
														<input type="radio" value="0" name="optAnalZ<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%> onClick="ValidatePopAnalHead(this)" >
 														No
 														<%end if %>
														&nbsp;&nbsp;<img id="imgAnalyticalEntryZ<%=sUnitarr(iCtr)%>" style="cursor: pointer"  border='0' src='../../assets/images/iTMS Icons/EntryIcon.gif' alt='Select Analytical Code' onClick="PopAnalyticalHead('<%=sUnitarr(iCtr)%>')" disabled >
 														<input type="hidden" name="hSelAnayCodeZ<%=sUnitarr(iCtr)%>" value="">
													</td>
													<%Next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Elgible For <br>Contra Entry
													</td>
													<%

													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = ""

														if iAccHead<>"" then

															sQuery ="SELECT EligibleForContras FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead

															'Response.Write sQuery
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bContra=objRs(0)
															else
																bContra = "0"
															end if
															objRs.Close

															sQuery = "Select AccountHead From Acc_T_CreatedVoucherHeader Where AccountHead = "&iAccHead&" "&_
																	 "and OUDefinitionID = '"&Trim(sUnitarr(iCtr))&"' "
															objRs.Open sQuery,Con
															IF Not objRs.EOF Then
																sAmenType = "disabled"
															Else
																sAmenType = ""
															End IF
															objRs.Close


															IF Len(Trim(sAmenType)) = 0 Then
																sQuery = "Select AccUnitAccountHead From Acc_T_CreatedVoucherDetails Where  "&_
																		 "AccUnitAccountHead = "&iAccHead&"  and AccountingUnit = '"&Trim(sUnitarr(iCtr))&"' "
																objRs.Open sQuery,Con
																IF Not objRs.EOF Then
																	sAmenType = "disabled"
																Else
																	sAmenType = ""
																End IF
																objRs.Close
															End IF
														end if 'if iAccHead<>"" then



													%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(Trim(bContra)) = "1" Then %>
 														<input type="radio" value="1" name="optContra<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%>>
 														Yes
														<input type="radio" value="0" name="optContra<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														No
 														<%else%>
 														<input type="radio" value="1" name="optContra<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														Yes
														<!--input type="radio" value="0" name="optContra<%=sUnitarr(iCtr)%>" class="FormElem" Checked onClick="ContraEnt('<%=sUnitarr(iCtr)%>',this)" <%=sDisType%> <%=sAmenType%>-->
														<input type="radio" value="0" name="optContra<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%>>
 														No
 														<%end if %>
													</td>
													<%Next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Eligible for TDS
													</td>
													<%

													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = ""

														if iAccHead<>"" then

															sQuery ="SELECT EligibleForTDS FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bTDS=objRs(0)
															else
																bTDS = "0"
															end if
															objRs.Close
														end if 'if iAccHead<>"" then

													%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(bTDS) = "1" Then %>
 														<input type="radio" value="1" name="optTDS<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%>>
 														Yes
														<input type="radio" value="0" name="optTDS<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														No
 														<%else%>
 														<input type="radio" value="1" name="optTDS<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														Yes
														<!--input type="radio" value="0" name="optTDS<%=sUnitarr(iCtr)%>" class="FormElem" Checked onClick="TDSCheck('<%=sUnitarr(iCtr)%>', this)" <%=sDisType%> <%=sAmenType%>-->
														<input type="radio" value="0" name="optTDS<%=sUnitarr(iCtr)%>" class="FormElem" Checked  <%=sDisType%> <%=sAmenType%>>
 														No
 														<%end if %>
													</td>
													<%Next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Memorandum A/c
													</td>
													<%

													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = ""

														if iAccHead<>"" then

															sQuery ="SELECT MemorandumAccount FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																bMemorandum=objRs(0)
															else
																bMemorandum = "0"
															end if
															objRs.Close
														end if 'if iAccHead<>"" then

													%>
													<td class="ExcelFieldCell">&nbsp;
														<% IF CStr(bMemorandum) = "1" Then %>
 														<input type="radio" value="1" name="optMemo<%=sUnitarr(iCtr)%>" class="FormElem" Checked <%=sDisType%> <%=sAmenType%>>
 														Yes
														<input type="radio" value="0" name="optMemo<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														No
 														<%else%>
 														<input type="radio" value="1" name="optMemo<%=sUnitarr(iCtr)%>" class="FormElem" <%=sDisType%> <%=sAmenType%>>
 														Yes
														<!--input type="radio" value="0" name="optMemo<%=sUnitarr(iCtr)%>" class="FormElem" Checked onClick="TDSCheck('<%=sUnitarr(iCtr)%>', this)" <%=sDisType%> <%=sAmenType%>-->
														<input type="radio" value="0" name="optMemo<%=sUnitarr(iCtr)%>" class="FormElem" Checked  <%=sDisType%> <%=sAmenType%>>
 														No
 														<%end if %>
													</td>
													<%Next%>
												</tr>

												<tr>
													<td class="ExcelDisplayCell">Allow Transactions &gt; <br>Rs. <%=dCreditLimit%>&nbsp;
													</td>
													<%

													For iCtr = 0 To UBound(sUnitarr)-1
														sAmenType = sDisArr(iCtr)
														sAmenType = Right(sAmenType,1)
															sAmenType = ""

														if iAccHead<>"" then
															sQuery ="SELECT AllowTransactions FROM Acc_R_OrgGLAccountHead where "&_
																	"OUDefinitionID='"&sUnitarr(iCtr)&"' and AccountHead="&iAccHead
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing

															if not objRs.EOF then
																sTransLimit=objRs(0)
															else
																sTransLimit = "W"
															end if
															objRs.Close
														end if 'if iAccHead<>"" then

													%>
													<td class="ExcelFieldCell">&nbsp;
														<%IF CStr(sTransLimit) = "A" Then %>
 														<input type="radio" value="A" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem" Checked>
 														Allow &nbsp;&nbsp;
														<input type="radio" value="R" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem">
 														Restrict &nbsp;&nbsp;
														<input type="radio" value="W" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem">
 														Warn
 														<%elseif CStr(sTransLimit) = "R" Then%>
 														<input type="radio" value="A" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem">
 														Allow &nbsp;&nbsp;
														<input type="radio" value="R" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem" Checked>
 														Restrict &nbsp;&nbsp;
														<input type="radio" value="W" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem">
														Warn
														<%else%>
														<input type="radio" value="A" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem">
 														Allow &nbsp;&nbsp;
														<input type="radio" value="R" name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem">
 														Restrict &nbsp;&nbsp;
														<input type="radio" value="W" checked name="optTrans<%=sUnitarr(iCtr)%>" class="FormElem">
 														Warn
 														<%end if %>
													</td>
													<%Next%>
												</tr>


											</table>
										</div>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td class="FieldCell">Used in Application
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtAppUsed" size="50" class="FormElemRead" readonly value="<%=sAppUsed%>">
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<img border='0' src='../../assets/images/iTMS Icons/EntryIcon.gif' onClick="PopUsed('A')" alt='Select Used in Application'>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Frequently Used Books
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtBooks" size="50" class="FormElemRead" readonly value="<%=sFrqBooks%>">
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<img border='0' src='../../assets/images/iTMS Icons/EntryIcon.gif' alt='Select Frequently Used Books' onClick="PopUsed('B')">
												</td>
											</tr>
										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
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
													<input type="button" value="Save" name="B9" class="ActionButton" onClick="Finaldone()">
 													<input type="reset" value="Reset" name="B1" class="ActionButton" >
 													<Input type="hidden" name="hDisUnits" value="<%=sDisUnits%>">
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
</body>
</html>
