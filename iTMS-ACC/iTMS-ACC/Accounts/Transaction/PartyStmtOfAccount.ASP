<%@ Language="VBScript" %>
<% option explicit %>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PartyStmtOfAccount.ASP
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 16, 2011
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/getCurrentDate.asp"-->
<%

	Dim sPartyType,sFormVal,sTemparr,sOrgId,sSql,sPartyCode,iCnt
	Dim sLedgCrDrId,sFromMonth,sToMonth,sToYear,sFromYear,sPArtySubTypeName
	Dim iPartyCodeParam,iPartyNameParam,iCRDRArr,sTempArrN,iCount,sTemArrCDDR

	Dim iVouNo,sPartyName
	dim iTransNo,sTranType,iVouEntryNo,sVouDate,sVouNarration,dVouAmount,sTransCrDr,iPartyCode,dOpeningAmt
	dim sOpenCrDr,sCloseCrDr,dCloseBal,sPartySubType,sFromDate,bFlag,sCurrentDate,dPayTotal,dRecTotal
	dim dAmount,sOrgName,sOrgnPartyCode,sCurrMonth,iPartySubTypeParam,sDispType,sSubTypeName,sFlag
	dim iCurrentParty,oFlag,iopCurrentParty,iSno,sTempVouNo,sNarration,iCrTransNo,sCNDNType
	dim sAllParType

	dim sCompFromDt1,sCompToDt1,sLastToDay,sCompFromDt,sCompToDt
	Dim sInsTy,sInsNo,sFinPeriod,sFromMonthCheck,sToMonthCheck,sBookCode

	Dim Objrs,Objrs1,objRs2

	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")
	set objRs2=server.CreateObject("Adodb.Recordset")

	sOrgId = session("organizationcode")
	sLedgCrDrId = Request("PARTYTYPE")	'CR or DR
	sFormVal = Request("hFormVal")
	sFinPeriod=session("finperiod")

	'Response.Write "<p><font color=red>hFormVal="&sFormVal
	Dim sPartySubTypeCheck,sPartySubTypeCR,sPartySubTypeDR,sPartyTypeCR,sPartyTypeDR

	If sFormVal <> "" Then
		sLedgCrDrId = Split(sFormVal,":")(0)	'PartyType
		sPartySubTypeCheck = Split(sFormVal,":")(1)	'PartySubType check Specific or All
		If sPartySubTypeCheck <> "A" Then
			sPartySubType = Split(sFormVal,":")(2)
		Else
			sPartySubTypeCR = Split(sFormVal,":")(6)	'Sub Type Value
			sPartyTypeCR = Split(sFormVal,":")(7)	'CR Value
			sPartySubTypeDR = Split(sFormVal,":")(8)	'Sub Type Value
			sPartyTypeDR = Split(sFormVal,":")(9)	'DR Value
		End IF
		iPartyCodeParam = Split(sFormVal,":")(3)	'PartyCode
		sFromMonth=Right(Split(sFormVal,":")(4),2)
		sToMonth=Right(Split(sFormVal,":")(5),2)
		sToYear = Left(Split(sFormVal,":")(5),4)
		sFromYear = Left(Split(sFormVal,":")(4),4)
		sFromMonthCheck = Split(sFormVal,":")(4)
		sToMonthCheck = Split(sFormVal,":")(5)
	Else
	    ''blocked and added by ragav on july 26,2014
	    ''begin
		   ' sFromMonth	= "04"
		   ' sToMonth	= "03"
		  'sFromYear	= Split(sFinPeriod,":")(0)
	       'sToYear		= Split(sFinPeriod,":")(1)

	       sToMonth = Month(getCurrentDate())
	       if cdbl(sToMonth) = 1 then
		        sFromMonth = "12"
		   else
		        sFromMonth = right("0"&cdbl(sToMonth)-1,2)
		   end if

	        sToMonth = Right("0"&cdbl(sToMonth),2)
		    sFromYear	= Split(sFinPeriod,":")(0)
		    if CDbl(sToMonth)>12 then
	            sToYear		= Split(sFinPeriod,":")(1)
	        else
	            sToYear = sFromYear
	        end if
		''end

	    sFromMonthCheck = ""
	    sToMonthCheck = ""
	End IF
	sFromDate = "01/"&sFromMonth&"/"&sFromYear


	IF CStr(sToMonth) = "04" or CStr(sToMonth) = "06" or CStr(sToMonth) = "09" or CStr(sToMonth) = "11" Then
		sLastToDay = "30"
	else
		sLastToDay = "31"
	End IF

	if CStr(sToMonth) = "02" Then
		IF CDbl(sToYear) Mod 4 = 0 Then
			sLastToDay = "29"
		Else
			sLastToDay = "28"
		End IF
	End IF

	sCompFromDt = "01/"&sFromMonth&"/"&sFromYear
	sCompToDt = sLastToDay&"/"&sToMonth&"/"&sToYear

	If sPartySubType <> "" Then
		sSql = "Select Distinct SubTypeName From VwOrgParty Where PartyType='"& sLedgCrDrId &"' and PartySubType="& sPartySubType &" "
		objrs.Open sSql,con
		Do While Not objrs.EOF
			sPArtySubTypeName = sPArtySubTypeName & "," & objrs(0)
			objrs.MoveNext
		Loop
		objrs.Close
		If sPArtySubTypeName <> "" Then
			sPArtySubTypeName = Mid(sPArtySubTypeName,2)
		End IF
	End IF
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<!-- XML Data Island -->
<XML ID="PartyData"><Root/></XML>
<XML ID="PartySubTypeData"><Root/></XML>
<xml id="PartyAllTypes"><Root></Root></xml>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/PartyStmtOfAccountCompat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0">
<%
	Const iPageSize=15
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,hCnt

	iCurrentPage=CInt(Request.Form("hPageSelection"))

	'iCnt=Request.Form("hCnt")
%>
	<!--<form method="POST" name="formname" action="ACCPURVOUCHERS.ASP">-->
	<form method="POST" name="formname" action="PartyStmtOfAccount.ASP">

		<input type=hidden name="hOrgID" value="<%=sOrgId%>">
		<input type=hidden name="hPartyCode" value="<%=iPartyCodeParam%>">
		<input type=hidden name="hPartyType" value="<%=sLedgCrDrId%>">
		<Input type="Hidden" Name="hLedCrOrDr" value="<%=sLedgCrDrId%>">
		<Input type="Hidden" Name="hPArtySubType" value="<%=sPartySubType%>">
		<Input type="Hidden" Name="hFormVal" value="">
		<Input type="hidden" name="hCRCtr" value="0">
		<Input type="hidden" name="hDRCtr" value="0">
		<Input type="hidden" name="hPartySubTypeCheck" value="0">
		<Input type="hidden" name="hPartyTypeCheck" value="0">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				<%If sLedgCrDrId <> "S" Then%>
					<%If sPartySubTypeCheck <> "A"  and sPArtySubTypeName <> "" Then%>
						<%If sLedgCrDrId = "CR" Then Response.Write "Creditors" Else Response.Write "Debtors" End IF%> [ <%=sPArtySubTypeName%> ]
					<%Else%>
						<%If sLedgCrDrId = "CR" Then Response.Write "Creditors" Else Response.Write "Debtors" End IF%>
					<%End IF%>
					Party Statement Of Account From <%=MonthName(sFromMonth)%>&nbsp;<%=sFromYear%> To <%=MonthName(sToMonth)%>&nbsp;<%=sToYear%>
				<%Else%>
					Party Statement Of Account From <%=MonthName(sFromMonth)%>&nbsp;<%=sFromYear%> To <%=MonthName(sToMonth)%>&nbsp;<%=sToYear%>
				<%End IF%>
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="middle"><a style="width: 1em; height: 1em;" title="" onclick="return Div_OnClick(idUnprocessed,'',event)">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="middle" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
	<td width="100%">
	<div id="idUnprocessed" style="display: none">
		<Table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
                <tr>
	                <td class="FieldCellSub">Party Type</td>
	                <td class="FieldCell" colspan="2">
		                <Select size="1" class="FormElem" name="SelPartyType">
							<option value="S" Selected>Select Party Type</option>
							<option value="CR" <%If sLedgCrDrId = "CR" Then Response.Write "selected"%>>Creditors</option>
			                <option value="DR" <%If sLedgCrDrId = "DR" Then Response.Write "selected"%>>Debtors</option>
		                </select>
	                </td>
                </tr>
                <tr>
	                <td class="FieldCellSub">Party SubType</td>
	                <td class="FieldCell" colspan="2">
		                <SELECT class="FormElem" size=1 name="SelPartySubType" onChange="SelectPartyHead()">
		                <OPTION value="A" <%If sPartySubTypeCheck = "A" Then Response.Write "selected"%>>All Party Sub Types</OPTION>
		                <OPTION value="S" <%If sPartySubTypeCheck = "S" Then Response.Write "selected"%>>Specific Party Sub Type</OPTION></SELECT>&nbsp;&nbsp;
			            <span id="spPartyHead" class="DataOnly"><%=sPArtySubTypeName%></span></TD>
	                </td>
                </tr>
<!--<tr>
	<td class="FieldCellSub">Party SubType</td>
	<td class="FieldCell" colspan="2">
		<textarea name="txtPartySubType" class="FormElemRead" readonly rows="1" cols=30></textarea>
		<img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="select Specific Party Type" width="11" height="11" Style="cursor: pointer" onclick="PartySubTypeSel()">
	</td>
</tr>-->

                    <tr>
	                    <td class="FieldCellSub" style="height: 34px">Party</td>
	                    <td class="FieldCell" colspan="2" style="height: 34px">
		                    <textarea name="txtPartyName" class="FormElemRead" readonly rows=2 cols=40></textarea>
		                    <img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="select Specific Party Name" width="11" height="11" Style="cursor: pointer" onclick="PartySelection()">
	                    </td>
                    </tr>

                    <Tr>
	                    <td colspan="3">
	                        <div id="divCRDRSubType" style="display:none" >
		                        <Table border="0" cellspacing="0" cellpadding="0" width="100%">
		                        <Tr>
		                            <Td width="1%"></td>
			                        <TD width="48%"  style="height: 35px" valign="top">
					                        <Table  id="CRTable" border="0" cellspacing="1px" cellpadding="0" class="ExcelTable" width="100%">
					                            <tr>
					                                <td class="ExcelHeaderCell" colspan=3 align=center>Creditor Party SubType</td>
					                            </tr>
						                        <Tr>
							                        <td class="ExcelHeaderCell" width="10px"></td>
							                        <td class="ExcelHeaderCell">Sub Type</td>
							                        <td class="ExcelHeaderCell">Status</td>
						                        </tr>
					                        </Table>

			                        </Td>
			                        <Td  width="2%" style="height: 35px"></td>
			                        <TD  width="48%" style="height: 35px" valign="Top">

					                        <Table id="DRTable" border="0" cellspacing="1px" cellpadding="0" class="ExcelTable" width="100%">
					                            <tr>
					                                <td class="ExcelHeaderCell" colspan=3 align=center>Debtor Party SubType</td>
					                            </tr>
						                        <Tr>
							                        <td class="ExcelHeaderCell" width="10px"></td>
							                        <td class="ExcelHeaderCell">Sub Type</td>
							                        <td class="ExcelHeaderCell">Status</td>
						                        </tr>
					                        </Table>

			                        </Td>
			                        <td  width="1%"></td>
			                        </tr>
		                        </table>
		                    </div>
	                    </td>
                    </Tr>

                    <tr>
	                    <TD class="FieldCellSub">Period</TD>
	                    <TD colspan="2">
		                    <table border="0" cellspacing="0" cellpadding="0">
		                    <tr>
			                    <td class="FieldCellSub">&nbsp; From </td>
			                    <td class="FieldCellSub">
			                    <select size="1" name="SelFromMonth" class="FormElem">
				                    <%

				                    Dim sFromYearN,sToYearN,iCounter

				                    sFromYearN=CDbl(Right(getFromFinYear(),4)&Left(getFromFinYear(),2))
				                    sToYearN=CDbl(Right(getToFinYear(),4)&Left(getToFinYear(),2))

				                    iCounter=sFromYearN

				                    Do While iCounter<=sToYearN
										If Trim(iCounter) = sFromMonthCheck and sFromMonthCheck <> "" Then
											Response.Write "<option Value="""&iCounter&""" selected>"&MonthName(Right(iCounter,2))&"-"&Left(iCounter,4)&"</option>"
										Else
											Response.Write "<option Value="""&iCounter&""">"&MonthName(Right(iCounter,2))&"-"&Left(iCounter,4)&"</option>"
										End IF

										iCounter=CDbl(iCounter)+1
										IF CDbl(Right(iCounter,2))>12 Then
				 						   iCounter=CDbl(CDbl(Left(iCounter,4))+1&"01")
										End IF
				                    Loop
				                    %>

			                    </select></td>
			                    <td class="FieldCellSub">To </td>
			                    <td class="FieldCellSub">
			                    <select size="1" name="SelToMonth" class="FormElem">
			                    <%
			                    sFromYearN=CDbl(Right(getFromFinYear(),4)&Left(getFromFinYear(),2))
			                    sToYearN=CDbl(Right(getToFinYear(),4)&Left(getToFinYear(),2))
			                    iCounter=sFromYearN

			                    Do While iCounter<=sToYearN
									If Trim(iCounter) = Trim(sToMonthCheck) and sToMonthCheck <> "" Then
										Response.Write "<option Value="""&iCounter&""" selected>"&MonthName(Right(iCounter,2))&"-"&Left(iCounter,4)&"</option>"
				                    Else
										Response.Write "<option Value="""&iCounter&""">"&MonthName(Right(iCounter,2))&"-"&Left(iCounter,4)&"</option>"
				                    End IF

				                    iCounter=CDbl(iCounter)+1
				                    IF CDbl(Right(iCounter,2))>12 Then
					                    iCounter=CDbl(CDbl(Left(iCounter,4))+1&"01")
				                    End IF
			                    Loop
			                    %>
		                    </select></td>
		                    <td class="FieldCell" >
			                    <input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
		                    </td>
	                    </tr>
	                    </Table>
                    </td>
                    </tr>

                    </table>
                    </div>
                    </td>
                    </tr>
                    </table>
                    </td>
                    </tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top">
<!--<div class="frmBody" id="frm4" style="width: 585; height:235;">-->
<table border="0" cellspacing="1" class="ExcelTable" width="100%">
<tr>
	<td class="ExcelHeaderCell" Colspan="9">Party</td>
</tr>
<tr>
	<td class="ExcelHeaderCell" rowspan="2">S.No.</td>
	<td class="ExcelHeaderCell" colspan="2"> Voucher</td>
	<td class="ExcelHeaderCell" rowspan="2">Narration</td>
	<td class="ExcelHeaderCell" colspan="5">Amount</td>
</tr>
<tr>
	<td class="ExcelHeaderCell">No</td>
	<td class="ExcelHeaderCell">Entry</td>
	<td class="ExcelHeaderCell">Date</td>
	<td class="ExcelHeaderCell">Type</td>
	<td class="ExcelHeaderCell">Debit</td>
	<td class="ExcelHeaderCell">Credit</td>
	<td class="ExcelHeaderCell">Opening/Closing</td>
</tr>
<%
	Dim iRecCnt,Flag,sQuery
	Dim iCrVal,iDrVal,iOpenVal,iOldParCode,sCheckPartySelection

	Flag = True
	sCheckPartySelection = True
	'Response.Write iPartyCodeParam &" " & iPartySubTypeParam &"<br>"
	'Response.Write "<p><font color=red>sLedgCrDrId="&sLedgCrDrId

	sQuery = " Select Distinct(partycode),orgnpartycode,partyName,PartySubType,PartyType from vwOrgparty where "&_
		     " OUDefinitionID='" & sOrgId & "' "

	sQuery = sQuery &" and PartyCode in (Select PartyCode from Acc_T_PartyTransactions where Convert(datetime,VoucherDate,103) "&_
	        " between Convert(Datetime,'"& sCompFromDt &"',103) and Convert(datetime,'"& sCompToDt &"',103) Group By PartyCode)"

	If sLedgCrDrId <> "" and sLedgCrDrId <> "S" Then
		sQuery =sQuery& " and partytype='" & sLedgCrDrId  & "' "
	Else
		sCheckPartySelection = False
	End IF

	If Trim(iPartyCodeParam)<>"0" and trim(iPartyCodeParam)<>"" then
		sQuery =sQuery& " and partycode in ("& iPartyCodeParam&")"
		iPartyCode = iPartyCodeParam
	Else
		Flag = False
	End if

	If sPartySubTypeCheck = "S" and sPartySubType <> "" Then
		sQuery =sQuery & " and PartySubType IN ("& sPartySubType&") "
	End If

	If sLedgCrDrId = "S" and sPartySubTypeCheck = "S" Then
		Flag = False
		sCheckPartySelection = False
	End IF

	'If sPartySubTypeCR <> "" Then
	'	sQuery =sQuery & " and PartySubType IN ("& sPartySubTypeCR&") "
	'End IF

	'If sPartySubTypeDR <> "" Then
	'	sQuery =sQuery & " and PartySubType IN ("& sPartySubTypeDR&") "
	'End IF

	'If sPartyTypeCR <> "" Then
	'	sQuery =sQuery& " and partytype='" & sPartyTypeCR  & "' "
	'End IF
	'If sPartyTypeDR <> "" Then
	'	sQuery =sQuery& " and partytype='" & sPartyTypeDR  & "' "
	'End IF

	sQuery = sQuery & " Order By PartyName "
	'Response.write "<p><font color=red>Sql="&sQuery

	with objRs
		.CursorLocation =3
		.CursorType =3
		.ActiveConnection=con
		.Source =sQuery
		.Open
	End with

	If not objRs.EOF Then
		iOldParCode = 0

	'------------- Displaying Opening Balance For the First Time -----------
	'----------------- Displaying Opening Amount ---------------------------
	%><!--<textarea><%=sQuery%></textarea>--><%
	dOpeningAmt = 0
	oFlag=true

	while not objRs.EOF
		iPartyCode = objRs(0)
		sOrgnPartyCode = objRs(1)
		sPartyName = objRs(2)
		sPartySubType = objRs(3)
		sLedgCrDrId = objrs(4)
		'Response.Write sPartySubType
		sAllParType = sAllParType&","& objRs(3)

		dOpeningAmt = dOpeningAmt + GetPartyDayOpening(sOrgId,sLedgCrDrId,sPartySubType,iPartyCode,sFromDate)
		dCloseBal=dOpeningAmt
		If Flag = False Then iPartyCodeParam = iPartyCodeParam & "," & iPartyCode
		'If Flag = True Then iCRDRArr = iCRDRArr & "," & sLedgCrDrId
		If sCheckPartySelection = False Then iCRDRArr = iCRDRArr & "," & sLedgCrDrId
		objRs.MoveNext
	Wend
	objRs.Close

	sAllParType = Mid(sAllParType,2)
	If Flag = False Then  iPartyCodeParam = Mid(iPartyCodeParam,2)

	'If Flag = True Then iCRDRArr = Mid(iCRDRArr,2)
	If sCheckPartySelection = False Then iCRDRArr = Mid(iCRDRArr,2)

	'Response.Write "<p>iPartyCodeParam="&iPartyCodeParam
	'Response.Write "<p>iCRDRArr="&iCRDRArr

	if dOpeningAmt <0 then
		sOpenCrDr="CR"
		sCloseCrDr="CR"
		dOpeningAmt=dOpeningAmt*-1
	else
		sOpenCrDr="DR"
		sCloseCrDr="DR"
	end if
	dOpeningAmt=FormatNumber(dOpeningAmt,2,,,0)
	'Response.Write "<p>Amt="&dOpeningAmt

	sTempArrN = Split(iPartyCodeParam,",")
	sTemArrCDDR = Split(iCRDRArr,",")
	'Response.Write "<p>iCRDRArr="&iCRDRArr
	If iCRDRArr = "" Then
		sCheckPartySelection = True
	End IF

	For iCount = LBound(sTempArrN) to UBound(sTempArrN)
		iPartyCode = sTempArrN(iCount)
		If sCheckPartySelection = False Then
			If UBound(sTemArrCDDR) >= iCount Then
				sLedgCrDrId = sTemArrCDDR(iCount)
			End IF
		End IF
		dOpeningAmt = 0

		sQuery = " select distinct(partycode),orgnpartycode,partyName,PartySubType from vwOrgparty where "&_
		         " OUDefinitionID='" & sOrgId & "'  "

		If iPartyCode <> "" Then
		      sQuery = sQuery & " and partycode = "& iPartyCode &" "
		End IF

		'If sLedgCrDrId <> "S" and sLedgCrDrId <> "" Then
		If sLedgCrDrId <> "" Then
		      sQuery = sQuery & " and partytype='" & sLedgCrDrId  & "'  "
		End IF

		sQuery = sQuery & " Order By PartyName "
		'Response.write "<p><font color=red>Sql="&sQuery

		with objRs
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sQuery
			.Open
		End with

		If not objRs.EOF Then
			iOldParCode = 0

			iPartyCode = objRs(0)
			sOrgnPartyCode = objRs(1)
			sPartyName = objRs(2)
			sPartySubType = objRs(3)
			sAllParType = sAllParType&","& objRs(3)

			dOpeningAmt = dOpeningAmt + GetPartyDayOpening(sOrgId,sLedgCrDrId,sPartySubType,iPartyCode,sFromDate)
			dCloseBal=dOpeningAmt

		End if
		objRs.Close

		sAllParType = Mid(sAllParType,2)

		if dOpeningAmt <0 then
			sOpenCrDr="CR"
			sCloseCrDr="CR"
			dOpeningAmt=dOpeningAmt*-1
		else
			sOpenCrDr="DR"
			sCloseCrDr="DR"
		end if
		dOpeningAmt=FormatNumber(dOpeningAmt,2,,,0)


	'To Find any transactions avialble for current party

	sQuery = "Select Count(1) from Acc_T_PartyTransactions Where PartyType='"&sLedgCrDrId&"' and  "&_
			 "partycode in ("&iPartyCode&") and OUDefinitionID='"&sOrgId&"' "&_
			 "and Convert(datetime,Voucherdate,103) >= Convert(datetime,'"&sCompFromDt&"',103)  "&_
			 "and Convert(datetime,Voucherdate,103) <= Convert(datetime,'"&sCompToDt&"',103)  "
			 'Response.Write sQuery
	objRs2.Open sQuery,Con
	IF Not objRs2.EOF Then
		iRecCnt = objRs2(0)
	End IF
	objRs2.Close

	'Response.Write "<p><font color=red>"&iRecCnt

	IF CStr(iRecCnt) <> "0" Then
		bFlag=true
		sFlag=true

		'sQuery="Select TransactionNumber,TransactionType,VoucherEntryNumber,"&_
		'	"convert(char,VoucherDate,103),VoucherNarration,Amount,TransCrDrIndication,"&_
		'	"partysubtype from Acc_T_PartyTransactions Where PartyType='" & sLedgCrDrId &_
		'	"'and PartySubType=" & sPartySubType &_
		'	" and partycode=" & iPartyCode &_
		'	" and OUDefinitionID='" & sOrgId & "'"&_
		'	" and month(VoucherDate)>=" & sFromMonth &_
		'	" and month(VoucherDate)<=" &sToMonth & " Order by voucherdate"

	sQuery = "Select TransactionNumber,TransactionType,VoucherEntryNumber,"&_
			 "convert(char,VoucherDate,103),VoucherNarration,Amount,TransCrDrIndication,"&_
			 "partysubtype from Acc_T_PartyTransactions Where PartyType='" & sLedgCrDrId &_
			 "'and PartySubType IN ("&sAllParType&") "&_
			 " and partycode in(" & iPartyCode &")"&_
			 " and OUDefinitionID='" & sOrgId & "'"&_
			 " and Convert(datetime,Voucherdate,103) >= Convert(datetime,'"&sCompFromDt&"',103) "&_
			 " and Convert(datetime,Voucherdate,103) <= Convert(datetime,'"&sCompToDt&"',103) "&_
			 " Order by voucherdate"

	sQuery = "Select TransactionNumber,TransactionType,VoucherEntryNumber,"&_
			 "convert(char,VoucherDate,103),VoucherNarration,Amount,TransCrDrIndication,"&_
			 "partysubtype from Acc_T_PartyTransactions Where PartyType='" & sLedgCrDrId &"' "&_
			 " and partycode in(" & iPartyCode &")"&_
			 " and OUDefinitionID='" & sOrgId & "'"&_
			 " and Convert(datetime,Voucherdate,103) >= Convert(datetime,'"&sCompFromDt&"',103) "&_
			 " and Convert(datetime,Voucherdate,103) <= Convert(datetime,'"&sCompToDt&"',103) "&_
			 " Order by voucherdate"
	%><!--<textarea id=textarea1 name=textarea1><%=sQuery%></textarea>--><%

	'Response.Write "<p><textarea>"& sQuery&"</textarea>"
	'Recordset To display the Party Transactions

	with objRs1
		.CursorLocation =3
		.CursorType =3
		.ActiveConnection=con
		.Source =sQuery
		.Open
	End with
	if not objRs1.EOF then

	sQuery="select partyName from vwOrgparty where "&_
			" OUDefinitionID='" & sOrgId & "' and partytype='" & sLedgCrDrId  & "'" &_
			" and partycode=" & iPartyCode & "and PartySubType =" & sPartySubType

	sQuery="select partyName from vwOrgparty where "&_
			" OUDefinitionID='" & sOrgId & "' and partytype='" & sLedgCrDrId  & "'" &_
			" and partycode=" & iPartyCode

	with objRs2
		.CursorLocation =3
		.CursorType =3
		.ActiveConnection =con
		.Source =sQuery
		.Open
	End with
	set objRs2.ActiveConnection =nothing
	If not objRs2.EOF then
		sPartyName =objRs2(0)
	End if
	objRs2.Close


	set objRs1.ActiveConnection=nothing
	set iTransNo=objRs1(0)
	set sTranType=objRs1(1)
	set iVouEntryNo=objRs1(2)
	set sVouDate=objRs1(3)

	set dAmount=objRs1(5)
	set sTransCrDr=objRs1(6)
	sCurrMonth=MonthName(Mid(sVouDate,4,2))
	End if
	If objRs1.RecordCount>0 then
	while not objRs1.EOF
		sVouNarration = objRs1(4)
		if oFlag then%>
		<tr>
				<TD COLSPAN=9 class="ExcelDisplayCell" align="Left"><b><%=sPartyName%></b></Td>
		</tr>
		<tr>
				<TD COLSPAN=6 class="ExcelDisplayCell" align="Right">
					<b>Opening As On&nbsp;<%=MonthName(sFromMonth)%>&nbsp;<%=sFromYear%></b></td>
				<TD class="ExcelDisplayCell" align="right"></td>
				<TD class="ExcelDisplayCell" align="Right">&nbsp;</TD>
				<TD class="ExcelDisplayCell" align="Right"><b><%=FormatNumber(Abs(dCloseBal),2,,,0)%>&nbsp;<%=sOpenCrDr%></b></td>
				<P align="right"></TD>
		</tr>
		<%
			iopCurrentParty=iPartyCode
			oFlag=false
		end if
	'----------- Function Call For Getting  the Opening Balance -----------
		if strcomp(iopCurrentParty,iPartyCode)<>0 then

			iopCurrentParty=iPartyCode
			dOpeningAmt=GetPartyDayOpening(sOrgId,sLedgCrDrId,sPartySubType,iPartyCode,sFromDate)
			dCloseBal=dOpeningAmt
			if dOpeningAmt <0 then
				sOpenCrDr="CR"
				sCloseCrDr="CR"
				dOpeningAmt=dOpeningAmt*-1
			else
				sOpenCrDr="DR"
				sCloseCrDr="DR"
			end if
			dOpeningAmt=FormatNumber(dOpeningAmt,2,,,0)

	'----------------- Displaying Opening Amount ---------------------------

	%>
				<tr>
				<TD COLSPAN=9 class="ExcelDisplayCell" align="Left"><b><%=sPartyName%></Td>
				</tr>
				<tr><TD COLSPAN=6 class="ExcelDisplayCell" align="Right">
					<b>Opening As On&nbsp;<%=sCurrMonth%>&nbsp;<%=sFromYear%></b></td>
				<TD class="ExcelDisplayCell" align="right"></td>
				<TD class="ExcelDisplayCell" align="Right">&nbsp;</TD>
				<TD class="ExcelDisplayCell" align="Right"><b><%=FormatNumber(Abs(dCloseBal),2,,,0)%>&nbsp;<%=sOpenCrDr%></td>
				<P align="right"></TD>
		</tr>
	<%
		End if

		'--------------------------TO GET THE VOUCHER NUMBER -----------------

		    sQuery = "select VoucherNumber,isNull(BankInstrumentType,''),isNull(BankInstrumentNo,'0'),BookCode,CreatedTransNo "&_
					 " from Acc_T_VoucherHeader where TransactionNumber=" & iTransNo

		   ' Response.Write "<p>"&sQuery

			with objRs2
				.CursorLocation =3
				.CursorType=3
				.ActiveConnection =con
				.Source=sQuery
				.open
			End with
			objRs2.ActiveConnection =nothing
			If not objRs2.EOF then
				iVouNo = objRs2(0)
				sBookCode = Trim(objRs2(3))
				iCrTransNo = trim(objrs2(4))
				if sBookCode = "02" then
				    sInsTy = Trim(objRs2(1))
    				sInsNo = Trim(objRs2(2))
    		    else
    		        sInsTy = ""
				    sInsNo = ""
				end if
			else
				iVouNo = " "
				sInsTy = ""
				sInsNo = ""
				iCrTransNo=""
			End if
			objRs2.Close

			sQuery = "Select IsNull(BankInstrumentType,'OT') from Acc_T_CreatedVoucherHeader where CreatedTransNo in ("
            sQuery = sQuery & " Select CreatedTransNo from Acc_T_VoucherHeader where TransactionNumber in ("& iTransNo &"))"
			objRs2.open sQuery,con
			if not objRs2.eof then
			    sCNDNType = trim(objrs2(0))
			end if
			objRs2.close

			Select Case Cstr(sInsTy)
				Case "C" sInsTy = "CH.No "
				Case "D" sInsTy = "DD No "
				Case "B" sInsTy = "CH.No "
				Case "T" sInsTy = "TT "
				Case "W" sInsTy = "CH.No "
			End Select

			IF Cstr(sInsNo) = "0" Then
				sInsTy = ""
				sInsNo = ""
			Else
				sInsTy = sInsTy&sInsNo&" "
			End IF

'---- Setting Flag For Checking Current And Previous voucher Dates ----
'------------- Displaying All the Transaction Details -----------------

		if bFlag then
			sCurrentDate=sVouDate
			bFlag=false
		end if
		if sFlag then
			iCurrentParty=iPartyCode
			sFlag=false
		end if

		if (Mid(sCurrentDate,4,2)<>Mid(sVouDate,4,2)) or  (cdbl(iCurrentParty)<>cdbl(iPartyCode)) Then

			'Response.write dCloseBal &" "& dPayTotal &" " & dRecTotal &"<br>"

			IF Cdbl(dCloseBal) < 0 Then
				dCloseBal = (CDbl(dCloseBal)- CDbl(dRecTotal))+CDbl(dPayTotal)

			ElseIF Cdbl(dCloseBal) > 0 Then
				dCloseBal = (CDbl(dCloseBal)+ CDbl(dPayTotal))-CDbl(dRecTotal)
			Else
				dCloseBal = (CDbl(dPayTotal)-CDbl(dRecTotal))
			End IF
	If 1 =2 Then
	%>
	<tr>
	<%IF Cdbl(Mid(sCurrentDate,4,2)) >= 4 and Cdbl(Mid(sCurrentDate,4,2)) <= 12 Then %>
		<TD colspan=6 class="ExcelDisplayCell" align="Right"><b>Closing for the Month Of  <%=MonthName(Mid(sCurrentDate,4,2))%>&nbsp;<%=sFromYear%></b></TD>
	<%Else%>
		<TD colspan=6 class="ExcelDisplayCell" align="Right"><b>Closing for the Month Of  <%=MonthName(Mid(sCurrentDate,4,2))%>&nbsp;<%=sToYear%></b></TD>
	<%End IF
	End IF	'If 1 =2 Then
	%>

	<!--<TD class="ExcelDisplayCell" align="right"><b><%'=FormatNumber(dPayTotal,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><b><%'=FormatNumber(dRecTotal,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"><P>-->
	<%		sCurrentDate=sVouDate
			iCurrentParty=iPartyCode
			dRecTotal=0
			dPayTotal=0
			if dCloseBal<0 then
				sCloseCrDr="CR"
			    'Response.Write "<b>"&FormatNumber(dCloseBal*-1,2,,,0) &"&nbsp;"&sCloseCrDr
			else
				sCloseCrDr="DR"
				'Response.Write "<b>"&FormatNumber(dCloseBal,2,,,0) &"&nbsp;"&sCloseCrDr
			end if

		'if sCloseCrDr="CR" then
		'	dCloseBal=CDbl(dCloseBal)*-1
		'end if

	End if
	iSno =iSno +1

	'===========================================================================================
	IF CStr(sTranType) = "CNR" or  CStr(sTranType) = "DNR" Then
		'sQuery = "Select CreatedVoucherNo +' '+ Convert(Char,VoucherDate,103) From Acc_T_CreatedVoucherHeader Where "&_
		'		 "CreatedTransNo IN (Select C.BankInstrumentNo From Acc_T_VoucherHeader A, "&_
		'		 "Acc_T_CreatedVoucherHeader C Where A.TransactionNumber = "&iTransNo&" and "&_
		'		 "A.BankInstrumentNo = C.CreatedTransNo ) "



		'With objRs2
		'	.CursorLocation = 3
		'	.CursorType = 3
		'	.Source = sQuery
		'	.ActiveConnection = Con
		'	.Open
		'End With
		'Set objRs2.ActiveConnection = Nothing
		'IF Not objRs2.EOF Then
		'	sTempVouNo = objRs2(0)
		'Else
		'	sTempVouNo = ""
		'End IF
		'objRs2.Close

		'IF CStr(sTempVouNo) <> "" Then
		'	sVouNarration = "Reference No - Sales Returns No - "&sTempVouNo&" "
		'End IF

		sQuery = "Select VoucherNarration From Acc_T_GLTransactions Where TransactionNumber = "&iTransNo&" and VoucherNarration is Not NULL"
		With objRs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set objRs2.ActiveConnection = Nothing
		IF Not objRs2.EOF Then
			sVouNarration = objRs2(0)
		Else
			sVouNarration = ""
		End IF
		objRs2.Close

		IF CStr(sLedgCrDrId) = "CR" and CStr(sPartySubType) = "1" and CStr(iVouEntryNo) = "0" Then
			sVouNarration = objRs1(4)
		End IF



	End IF

	'===========================================================================================

	%>
	<tr>
	<TD class="ExcelDisplayCell"><%=iSno%></TD>
	<TD class="ExcelDisplayCell" align="Left">
	<%if trim(sLedgCrDrId)="CR" and trim(iCrTransNo)<>"" then %>
	    <a href="#" class="ExcelDisplayLink" onclick="return ShowVouch('<%=iCrTransNo%>','<%=sTranType%>','<%=sCNDNType%>')"><%=iVouNo%></a>
	<%else%>
	    <%=iVouNo%>
	<%end if %>
	</TD>
	<TD class="ExcelDisplayCell" align="center">
			<%IF CStr(iVouEntryNo) = "0" Then
				Response.Write "<p>1</TD>"
			Else%>
			<P><%=iVouEntryNo%></TD>
			<%End IF %>
	<TD class="ExcelDisplayCell"><%=sInsTy%><%=sVouNarration%></TD>
	<TD class="ExcelDisplayCell" align="center"><%=sVouDate%></TD>
	<TD class="ExcelDisplayCell" align="center"><%=sTranType%></TD>

	<%
	if sTransCrDr="D" Then
		dPayTotal=CDbl(dPayTotal)+CDbl(dAmount)
	%><TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmount,2,,,0)%></td>
	<TD class="ExcelDisplayCell" align="right"></TD>
	<TD class="ExcelDisplayCell" align="right"></TD>
	<% Else
		dRecTotal=CDbl(dRecTotal)+CDbl(dAmount)
	%>
	<TD class="ExcelDisplayCell" align="right"></TD>
	<TD class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmount,2,,,0)%></TD>
	<TD class="ExcelDisplayCell" align="right"></TD>
	<%end if%>
	</tr>
<%
	objRs1.MoveNext
wend

%>
	<tr>
	<%IF Cdbl(Mid(sCurrentDate,4,2)) >= 4 and Cdbl(Mid(sCurrentDate,4,2)) <= 12 Then %>
		<TD colspan=6 class="ExcelDisplayCell" align="Right"><b>Closing for the Month Of  <%=MonthName(Mid(sCurrentDate,4,2))%>&nbsp;<%=sFromYear%></b></TD>
	<%Else%>
		<TD colspan=6 class="ExcelDisplayCell" align="Right"><b>Closing for the Month Of  <%=MonthName(Mid(sCurrentDate,4,2))%>&nbsp;<%=sToYear%></b></TD>
	<%End IF %>
	<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dPayTotal,2,,,0)%></b></TD>
	<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dRecTotal,2,,,0)%></b></TD>
	<TD class="ExcelDisplayCell" align="right">
	<%
	'if sCloseCrDr="CR" then
	'	dCloseBal=CDbl(dCloseBal)*-1
	'end if

	'Response.write dCloseBal &" "& dPayTotal &" " & dRecTotal &"<br>"


	IF Cdbl(dCloseBal) < 0 Then
		dCloseBal = (CDbl(dCloseBal)- CDbl(dRecTotal))+CDbl(dPayTotal)

	ElseIF Cdbl(dCloseBal) > 0 Then
		dCloseBal = (CDbl(dCloseBal)+ CDbl(dPayTotal))-CDbl(dRecTotal)
	Else
		dCloseBal = (CDbl(dPayTotal)-CDbl(dRecTotal))
	End IF

'--------------------- Displaying Closing Balance ---------------------
		'dCloseBal=CDbl(dCloseBal)- CDbl(dRecTotal)+CDbl(dPayTotal)
		if dCloseBal<0 then
			sCloseCrDr="CR"
		    Response.Write "<b>"&FormatNumber(dCloseBal*-1,2,,,0) &"&nbsp;"&sCloseCrDr
		else
			sCloseCrDr="DR"
			Response.Write "<b>"&FormatNumber(dCloseBal,2,,,0) &"&nbsp;"&sCloseCrDr
		end if
		%></td>
		</tr>
	<%
	End if
	dRecTotal=0
	dPayTotal=0
	objRs1.Close

Else
	IF CStr(iPartyCode) <> CStr(iOldParCode) Then
		sQuery = "Select Sum(OpeningAmount),OpeningCDIndication From Acc_T_PartyOpeningAmt Where  "&_
				 "PartyCode = "&iPartyCode&" and OpeningMonthYear = '042005' and PartyType='"&sLedgCrDrId&"' Group By OpeningCDIndication "


		iDrVal = 0
		iCrVal = 0

		objRs2.Open sQuery,Con
		IF Not objRs2.EOF Then
			Do While Not objRs2.EOF
				IF CStr(objRs2(1)) = "C" Then
					iCrVal = objRs2(0)
				Else
					iDrVal = objRs2(0)
				End IF
				objRs2.MoveNext
			Loop
		Else
			iDrVal = 0
			iCrVal = 0
		End IF
		objRs2.Close



		iOpenVal = CDbl(iDrVal) - CDbl(iCrVal)

		IF CDbl(iOpenVal) <> 0 Then



%>
			<tr>
				<TD colspan=4 class="ExcelDisplayCell" align="left"><b><%=sPartyName%></b></TD>
				<TD class="ExcelDisplayCell" align="right"><b>Opening Balance</b></TD>
				<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
				<!--TD class="ExcelDisplayCell" align="right">&nbsp;</TD-->
			<%IF CDbl(iOpenVal) < 0 then %>
					<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
					<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(abs(iOpenVal),2,,,0)%></b></TD>

			<%Elseif CDbl(iOpenVal) > 0 then %>

					<TD class="ExcelDisplayCell" align="right"><b><%=FormatNumber(abs(iOpenVal),2,,,0)%></b></TD>
					<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
			<%End IF %>
			<TD class="ExcelDisplayCell" align="right">&nbsp;</TD>
			</tr>

<%
			iOpenVal = 0
			End IF
		End IF
	End IF
	'objRs.MoveNext
	iOldParCode = iPartyCode
	'Wend
	'objRs.Close
Next
End IF

%>

</table>
<!--/div-->

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top" align="right">
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hCnt" value=<%=iCnt%>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="PaginateAcc('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage - 1%>')" id=button4 name=button4>
<%		end if	%>
<SELECT class="FormElem" onChange="PaginateAcc(this(this.selectedIndex).value)" id=select1 name=select1>
<%
For lnPage = 1 To iTotalPage
If lnPage = iCurrentPage Then
%>
<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
<%		else	%>
<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
<%		end if
next
%>
</SELECT>
<%
if iCurrentPage = iTotalPage then
%>
<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

<%		else	%>
<input type="button" value=" >> " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage + 1%>')" id=button7 name=button7>
<input type="button" value=" >| " class="ActionButtonX" onclick="PaginateAcc('<%=iTotalPage%>')" id=button8 name=button8>
<%		end if
End If
%>
</td>
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td class="ActionCell">
	<input type="button" value="Print Statement" name="B9" class="ActionButtonX" tabindex="3" onclick="">
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="BottomPack" colspan="3">
</td>
</tr>

</form>
</body>
</html>
