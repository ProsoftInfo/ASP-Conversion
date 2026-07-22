<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassSelection.asp
	'Module Name				:	Inventory (Transcation)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 03,2010
	'Modified On				:	Jan 06,2011
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
<%
	Dim sUnitid, sInvoiceType, sUnitName, iCtr, sInvoiceNo, Salrs, Salrs1
	Dim iGatePassNo, iForInvoiceNo, iPartyCode, sPartyName, sStatus,sSql
	Dim dFromDate,dToDate,sDCNo,sDCDate
	Dim sSentType, sInvTypeForDisp, sInvTypeName
	Set Salrs = Server.CreateObject ("ADODB.Recordset")
	Set Salrs1 = Server.CreateObject ("ADODB.Recordset")


	'sFinPeriod=session("finperiod")
	'sUnitid=Request("selUnitId")


	sUnitid=Session("Organizationcode")
	with Salrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sUnitid
		.Open
	end with
	if not Salrs.EOF then
		sUnitName = trim(Salrs(0))
	end if
	salrs.Close

	sInvoiceType=Request("hInvType")
	sSentType = Request("hSentType")
	dFromDate = Request.Form("hFromDate")
	dToDate = Request.Form("hToDate")
	if trim(sInvoiceType) ="" then
		sInvoiceType = Request.QueryString("InvoiceType")
		sSentType=Request.QueryString("SelSent")
	end if
	'if Trim(sInvoiceType )="" then sInvoiceType=0
	If sInvoiceType = "P" Then
		sInvTypeForDisp = "P,PR"
		sInvoiceType = "'P','PR'"
	Else
		sInvTypeForDisp = sInvoiceType
		sInvoiceType = Pack(sInvoiceType)
	End If

	if Trim(sInvTypeForDisp)="" then sInvTypeForDisp = "0"
'	Response.Write sInvoiceType
'	Response.Write sSentType

	'If Trim(sInvoiceType) = "'0'" Then sInvoiceType = "''"
	If sInvoiceType = "'0'" Then sInvoiceType = "''"
	If sSentType = "" Then sSentType = "N"

	if trim(dFromDate)="" and trim(dToDate)="" then
		dFromDate = "01/04/"&split(Session("FinPeriod"),":")(0)
		dToDate = FormatDate(date)
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">

<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/printwindow.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/gatePassSelection.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDefault()" >
<%
	Const iPageSize=15
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iPageNo,hCnt,iCnt, iTotPage

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	'iCnt=Request.Form("hCnt")
	'Response.Write Request("hInvNo")
	'Response.Write Request("hInvoiceNo")
%>
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
	<input type=hidden name="hInvNo" value="<%=Request("hInvNo")%>">
	<input type=hidden name="hUnitNo" value="<%=Request("hUnitNo")%>">
	<input type=hidden name="hInvoiceNo" value="<%=Request("hInvoiceNo")%>">
	<input type=hidden name="hInvoiceType" value="<%=sInvoiceType%>">
	<input type=hidden name="hSentType" value="<%=sSentType%>">
	<input type=hidden name="hFromDate" value="<%=dFromDate%>">
	<input type=hidden name="hToDate" value="<%=dToDate%>">
	<input type=hidden name="hInvType" value="<%=sInvTypeForDisp%>">
	<input type=hidden name="hOrgID" value="<%=sUnitid%>">
	<input type=hidden name="hOrgName" value="<%=sUnitName%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Gate Pass
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
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
    	<input type="radio" name="rStatus" value="N" <%if ucase(sSentType)="N" then  Response.Write "Checked"%> onclick="Status()">To Be Sent&nbsp;
			<input type="radio" name="rStatus" value="Y" <%if ucase(sSentType)="Y" then  Response.Write "Checked"%> onclick="Status()">Sent&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
<tr>
<td width="100%">
<div id="idUnprocessed" style="width: 575; display: none">
<table cellpadding="0" cellspacing="0">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<!--<tr>
<td class="FieldCellSub">&nbsp;&nbsp;</td>
<td class="FieldCellSub">Select Unit</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selUnitId" class="FormElem" >

	<%
		sSql ="SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		'Response.Write sSql
		with Salrs
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source=sSql
			.Open
		end with
		set Salrs.ActiveConnection=nothing
		If not Salrs.EOF then
			Do While Not Salrs.EOF
		%>
		   <OPTION VALUE="<%=Salrs(0)%>" ><%=Salrs(1)%></Option>
		<%
			Salrs.MoveNext
			Loop
		end if
		Salrs.Close
	%>
</select>
</td>
</tr>-->

<tr>
<td class="FieldCellSub">
</td>
<td class="FieldCellSub">Select Invoice Type
</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selInvType" class="FormElem" >
		<option value="0">Select InvoiceType</option>
	<!--	<option value="A">MILL SALES </option>
		<option value="T">TRANSFERS TO DEPOT</option>
		<option value="U">TRANSFER TO GROUP UNITS / COMPANIES</option>
		<option value="C">TRANSFER TO CONVERTORS FOR CONVERSION</option>
		<option value="S">INVOICE FOR SAMPLE ITEMS</option>
		<option value="I">STOCK ISSUED FOR SALES</option>
		<option value="D">UNPROCESSED ORDERS / DIRECT</option>
		<option value="N">INTER UNIT TRANSFERS</option>
		<option value="J">STOCK ISSUED FOR JOBWORK</option>
		<option value="P">PURCHASE RETURNS</option>-->

		<option value="X">Cash</option>
		<option value="Y">Non Exeise</option>
		<option value="Z">Exeise</option>
		<option value="V">SERVICES</option>
</select>
</td>
</tr>
<tr>
	<td></td>
	<td class="FieldCellSub">
	From Date
	</td>
	<td class="FieldCellSub">
	<%Response.Write insertDatePicker("ctlFromDate")%>
	</td>
	<td class="FieldCellSub">
	To Date
	</td>
	<td class="FieldCellSub">
	<%Response.Write insertDatePicker("ctlToDate")%>
	</td>
</tr>
<tr>
	<td></td>
	<td></td>
		<td class="FieldCell" colspan="4">

		</td>
</tr>

<tr>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell"></td>
<td class="FieldCell" >
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
</td>
<td class="FieldCell" >
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
</td>
</table>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
<table border="0" cellspacing="1" class="ExcelTable" width="100%" >
<tr>
<td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
<td class="ExcelHeaderCell" align="center" width="10" ></td>
<td class="ExcelHeaderCell" align="center" >DC No - Date</td>
<td class="ExcelHeaderCell" align="center" >Party / Unit Name</td>
<td class="ExcelHeaderCell" align="center" >Type</td>
<td class="ExcelHeaderCell" align="center" >Gate Pass Status</td>
<!--td class="ExcelHeaderCell" align="center" >Accounting</td-->
</tr>
	<%


			iCtr = 0
			'Response.Write sInvoiceType
				'Response.Write sInvoiceType
				If sInvoiceType <> "''" and sSentType = "N" Then
				'Response.Write sInvoiceType & "," & sSentType
					if Trim(sInvoiceType) = "R" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER WHERE OrganisationCode ='" & sUnitid & "' AND INVOICETYPE In (" & sInvoiceType & ") AND STATUS = 'N'"
					elseif Trim(sInvoiceType) = "U" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(TOUNIT,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER WHERE OrganisationCode ='" & sUnitid & "' AND INVOICETYPE  IN(" & sInvoiceType & ") AND STATUS = 'N'"
					else
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER WHERE OrganisationCode ='" & sUnitid & "' AND INVOICETYPE IN(" & sInvoiceType & ") AND STATUS = 'N'"
					end if
				ElseIf sInvoiceType = "''" and sSentType = "N" Then

					if Trim(sInvoiceType) = "'R'" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER where OrganisationCode ='" & sUnitid & "' AND STATUS = 'N'"
					elseif Trim(sInvoiceType) = "'U'" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(TOUNIT,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER where OrganisationCode ='" & sUnitid & "' AND STATUS = 'N'"
					else
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER where OrganisationCode ='" & sUnitid & "'  AND STATUS = 'N'"
					end if
				ElseIf sInvoiceType <> "''" and sSentType <> "N" Then
					if Trim(sInvoiceType) = "'R'" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER WHERE OrganisationCode ='" & sUnitid & "' AND INVOICETYPE In (" & sInvoiceType & ") and status ='Y'"
					elseif Trim(sInvoiceType) = "'U'" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(TOUNIT,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER WHERE OrganisationCode ='" & sUnitid & "' AND INVOICETYPE  IN(" & sInvoiceType & ") and status ='Y'"
					else
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER WHERE OrganisationCode ='" & sUnitid & "' AND INVOICETYPE IN(" & sInvoiceType & ") and status ='Y'"
					end if
				ElseIf sInvoiceType = "''" and sSentType <> "N" Then
					if Trim(sInvoiceType) = "'R'" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER where OrganisationCode ='" & sUnitid & "'"
					elseif Trim(sInvoiceType) = "'U'" then
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(TOUNIT,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER  where OrganisationCode ='" & sUnitid & "'"
					else
						sSql = "SELECT GATEPASSNO, ISNULL(ReferenceNo,0), ISNULL(PARTYCODE,0),InvoiceType,ISNULL(DCCODE,''),status,Convert(varchar,MarkedOn,103) FROM FORGATEPASSHEADER where OrganisationCode ='" & sUnitid & "'"
					end if
				End If

				sSql = sSql & " and Convert(datetime,MarkedOn,103) >= Convert(datetime,'"& dFromDate &"',103) and  Convert(datetime,MarkedOn,103) <= Convert(datetime,'"& dToDate&"',103) "
				'Response.Write sSql
			with salrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source  = sSql
				.Open
			end with
			set salrs.ActiveConnection = nothing

			If Not salrs.EOF then


		'''''''''''''''''''''''''''''''''''''''''''''''''''''''
   			Salrs.PageSize = iPageSize
			If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
			Salrs.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
			iTotPage = Salrs.PageCount					'stores total no. of pages
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''
			For iPageCtr = 1 to Salrs.PageSize



'			do while not Salrs.EOF
				iCtr = iCtr + 1
				iGatePassNo = Salrs(0)
				iForInvoiceNo = Salrs(1)
				iPartyCode = Salrs(2)
				sInvTypeForDisp = Salrs("InvoiceType")
				sDCNo = Salrs(4)
				sDCDate = Salrs(6)
				'Response.Write sInvTypeForDisp

				with salrs1
					.CursorLocation = 3
					.CursorType = 3
					if Trim(sInvoiceType) = "'R'" then
						.Source = "SELECT INVOICENUMBER, CONVERT(VARCHAR,INVOICEDATE,103) FROM SAL_T_INVOICEHEADER WHERE SALETRANSACTIONNO = " & iForInvoiceNo
					else
						.Source = "SELECT INVOICENUMBER, CONVERT(VARCHAR,INVOICEDATE,103) FROM SAL_T_INVOICEHEADER WHERE SALETRANSACTIONNO = (SELECT ISNULL(SALETRANSACTIONNO,0) FROM FORINVOICE_HEADER WHERE FORINVOICENO = " & iForInvoiceNo & ")"
						'.Source = "SELECT INVOICENUMBER, CONVERT(VARCHAR,INVOICEDATE,103) FROM SAL_T_INVOICEHEADER WHERE SALETRANSACTIONNO = (SELECT ISNULL(SALETRANSACTIONNO,0) FROM FORINVOICE_HEADER WHERE FORINVOICENO = " & iGatePassNo & ")"
					end if
					.ActiveConnection = con
					'Response.Write Salrs1.Source
					.Open
				end with
				'Response.Write Salrs1.Source
				set salrs1.ActiveConnection = nothing

				if Not Salrs1.EOF then
					if Trim(Salrs1(0)) <> "" then
						sInvoiceNo = Trim(Salrs1(0))&" - "&Trim(salrs1(1))
					else
						sInvoiceNo = Trim(salrs1(1))
					end if
					sStatus = "Generated"
				else
					sInvoiceNo = "N/A"
					sStatus = "Pending"
				end if
				Salrs1.Close
				If Salrs(5) = "N" Then
					sStatus = "To Be Created"
				Else
					sStatus ="Created"
				End If

				with salrs1
					.CursorLocation = 3
					.CursorType = 3
					if Trim(sInvoiceType) = "'U'" then
						.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = " & Pack(iPartyCode)
					else
						.Source = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE = " & iPartyCode
					end if
					.ActiveConnection = con
					.Open
				end with
				set salrs1.ActiveConnection = nothing

				if Not Salrs1.EOF then
					sPartyName = Salrs1(0)
				else
					sPartyName = ""
				end if
				Salrs1.Close
				If uCase(Trim(sInvTypeForDisp)) = "A" Then
					sInvTypeForDisp = "MILL SALES"
				ElseIf  uCase(Trim(sInvTypeForDisp)) = "T" Then
					sInvTypeForDisp = "TRANSFERS TO DEPOT"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "U" Then
					sInvTypeForDisp = "TRANSFER TO GROUP UNITS / COMPANIES"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "C" Then
					sInvTypeForDisp = "TRANSFER TO CONVERTORS FOR CONVERSION"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "S" Then
					sInvTypeForDisp ="INVOICE FOR SAMPLE ITEMS"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "I" Then
					sInvTypeForDisp = "STOCK ISSUED FOR SALES"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "D" Then
					sInvTypeForDisp = "UNPROCESSED ORDERS / DIRECT"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "N" Then
					sInvTypeForDisp = "INTER UNIT TRANSFERS"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "J" Then
					sInvTypeForDisp = "STOCK ISSUED FOR JOBWORK"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "P" Then
					sInvTypeForDisp ="PURCHASE RETURNS"
				ElseIf uCase(Trim(sInvTypeForDisp)) = "V" Then
					sInvTypeForDisp = "SERVICES"
				End If
				'Response.Write sInvTypeForDisp
%>



					<tr>
						<td class="ExcelSerial" align="center"><%=iCtr%></td>
						<td class="ExcelDisplayCell" width="20" align="center">
							<input type="radio" value="<%=iGatePassNo%>"  name="R1" class="FormElem" Onclick= "OptionClick(this)">
						</td>
						<td class="ExcelDisplayCell" align ="left"><%=sDCNo%> - <%=sDCDate%></td>
						<td class="ExcelDisplayCell"><%=sPartyName%></td>
						<td class="ExcelDisplayCell" align ="center"><%=sInvTypeForDisp%></td>
						<td class="ExcelDisplayCell" align ="center"><%=sStatus%></td>
						<% 'If sSentType = "N" and sInvoiceType = "'V'" Then %>
						<!--td class="ExcelDisplayCell" align ="center"><input type="Button" name="BtnA<%=iCtr%>" class="AddButtonX" value="Yes" onClick="Check(<%=iGatePassNo%>)">
						</td>
						<%' Else %>
						<td class="ExcelDisplayCell" align ="center"></td-->
						<%' End If %>

					</tr>
		<%
					Salrs.MoveNext
					If Salrs.EOF Then Exit For
					next
'					Salrs.MoveNext
'				loop
				End If
				Salrs.Close

		%>
</table>
<!--/div-->
</td>
<td align="center" class="ClearPixel" width="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top" align="right">
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hCnt" value=<%=iCtr  %>>
<input type=hidden name="hPageSelection" value="0">

<%	'Response.Write iTotPage
If iTotPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
<%		end if	%>
<SELECT class="FormElem" onChange="Paginate(this(this.selectedIndex).value)" id=select1 name=select1>
<%
For lnPage = 1 To iTotPage
If lnPage = iCurrentPage Then
%>
<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotPage%></OPTION>
<%		else	%>
<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
<%		end if
next
%>
</SELECT>
<%
if iCurrentPage = iTotPage then
%>
<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

<%		else	%>
<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotPage%>')" id=button8 name=button8>
<%		end if
End If
%>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td valign="middle" class="ActionCell">

<%if trim(sInvoiceType) ="'V'" then%>
	<input type="button" value="Edit" name="B11" class="ActionButton" tabindex="3" onclick="CheckEdit()">
<%end if%>
<input type="button" value="Create" name="B11" class="ActionButton" tabindex="4" onclick="CreateNew()">
<input type="button" value="Next" name="B9" class="ActionButton" tabindex="4" onclick="CheckSubmit()">
<input type="reset" value="Reset" name="B10" class="ActionButton" tabindex="5" >
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="BottomPack" colspan="3">
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
