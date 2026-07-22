<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassEntry.asp
	'Module Name				:	Sales - Gate Pass
	'Author Name				:	TAJUDEEN S
	'Created On					:	22 December, 2004
	'Modified By				:	ragavendran
	'modified on				:	27/05/2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	GatePassInsert.asp
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
<!--#include File="../../include/populate.asp" -->
<!--#include File="../../include/NoSeries.asp" -->
<!--#include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!--#include File="../../include/NoSeriesCommonFunctions.asp" -->
<HTML><HEAD><TITLE>iTMS - Gate Pass</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></script>


<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/gatePassEntry.js"></SCRIPT>

<%
	'XML DOM Variables
	Dim oDOM,newElem,RootNode,HeaderNode,DetailsNode
	Dim dGatePassDate,iAgentCode,sSaleTransNo,sInvoiceDate
	dim dcrs,dcrs1,iGatePassNo,arrTemp,arrTemp1,iClass,iItemCode,sSql,iCtr, iPartyCode
	dim sOrgID,sUnitDesc,sUnitAddr,sUnitAddr1,sUnitCity,sPostCode,sUnitState
	dim sDCNo,sItemType,iSeriesNo,iSeriesCode,sItemName,iQty,sReference,sInvoiceType
	Dim sSentType,sUoM,sForSubConNo,sApplFormJJ,sReturnToPage,sCallFrom
	Dim sFinPeriod, sFinPeriodFrom, sFinPeriodTo, sFinancialYearTo, ChkStr	
	Dim sPackNos, sSourceRefNo, sDCCode,sAttID,sOptName,sQuery,sArrID
	Dim iNumClassCode,sTempSeries,sArrSeries,sNumClassName
	
	set dcrs = Server.CreateObject("ADODB.RecordSet")
	set dcrs1 = Server.CreateObject("ADODB.RecordSet")
		
	iGatePassNo = Request.Querystring("GatePassNo")
	sInvoiceType = Request.Querystring("InvoiceType")
	sForSubConNo = Request.Querystring("ForSubConNo")
	sReturnToPage = Request.QueryString("ReturnToPage")
	sCallFrom	 = Request.QueryString("CallFrom")
	
	
	if trim(sForSubConNo) = "" then
		
	end if 
	
	sSentType = Request("hSentType")
	if trim(sSentType)="" then
	    sSql = "Select isNull(Status,'N') from FORGATEPASSHEADER where GatePassNo = "& iGatePassNo
	    dcrs.open sSql,con
	    if not dcrs.eof then
	        if trim(dcrs(0))="Y" then
	            sSentType = "Y"
	        else
	            sSentType = "N"
	        end if
	    end if
	    dcrs.close
	end if
	
	Dim sFinFrom, sFinTo, sTempMonYr, sMonYr, arrFin
	sFinPeriod = Session("FinPeriod")
	sFinPeriodFrom = FormatDate("04/01/" & Mid(sFinPeriod,1,4))
	sFinPeriodTo = FormatDate("03/31/" & Mid(sFinPeriod,6,4))
	sFinFrom = FormatDate("04/01/" & Mid(sFinPeriod,1,4))
	sFinTo = FormatDate("03/31/" & Mid(sFinPeriod,6,4))
'	If DateDiff("d",FormatDate(Date()),FormatDate(sFinTo)) < 0 Then
'		If len(Month(sFinTo)) = 1 Then
'			sTempMonYr = "0"&Month(sFinTo)
'		Else
'			sTempMonYr = Month(sFinTo)
'		End If 
'		'Response.Write sTempMonYr
'		sMonYr = sTempMonYr&Year(sFinTo)
'	Else
'		if len(Month(date())) = 1 then
'			sTempMonYr = "0"&Month(date())
'		else
'			sTempMonYr = Month(date())
'		end if
'		sMonYr = sTempMonYr&Year(date())
'	End If
'
''	Response.Write sFinPeriodFrom 
''	Response.Write sFinPeriodTo 	
'	arrFin = split(GetFinancialYear(sMonYr),":")
'	sFinFrom = arrFin(0)
'	sFinTo = arrFin(1)
dGatePassDate =""
if trim(iGatePassNo) <> "" then
	with dcrs
		.cursorlocation = 3
		.cursortype = 3
		.source = "select InvoiceType,ReferenceNo,DCCode From FORGATEPASSHEADER WHERE GATEPASSNO = " & iGatePassNo 
		.ActiveConnection = con
		.open
	End With
	set dcrs.activeConnection = Nothing
	'Response.Write "<p> " & dcrs.source
	If not dcrs.eof Then
		sInvoiceType =  trim(dcrs(0))
		sSaleTransNo = dcrs(1)
	End If	
	dcrs.close
End If 'if trim(iGatePassNo) <> "" then

if Trim(iGatePassNo)<>"" then
    sQuery = "SELECT IsNull(ITEMCODE,0), isNull(CLASSIFICATIONCODE,0), QUANTITY,Description,EntryNo,MaterialRcvd,InvoicedUoM,NoofPacks,PackingType,isNull(FormJJ,'N') as FormJJ,isNull(ItemAttributes,0) FROM FORGATEPASSDETAILS WHERE GATEPASSNO = " & iGatePassNo
    dcrs.open sQuery,con
    if not dcrs.eof then
        iNumClassCode = dcrs(1)
    end if
    dcrs.close
end if


If sInvoiceType = "C" Then
'	with dcrs
'		.cursorlocation = 3
'		.cursortype = 3
'		.source = "select ForSubContractNo From ForSubcontractHeader  where InvoiceNumber =  " & sSaleTransNo  & ""
'		.ActiveConnection = con
'		.open
'	End With
'	set dcrs.activeConnection = Nothing
'	'Response.Write "<p> " & dcrs.source
'	If not dcrs.eof Then
'		sForSubConNo = dcrs(0)
'	End If	
'	dcrs.close
End If 'If sInvoiceType = "C" Then


sInvoiceType = "'"& sInvoiceType &"'"
	
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	if sInvoiceType = "'U'" then
		.Source = "SELECT OrganisationCode, TYPEOFITEMS, TOUNIT, Convert(datetime,GENERATEDON,103) as markedon,ReferenceNo,ISNULL(DCCODE,'-'),isNull(Remarks,'') FROM FORGATEPASSHEADER WHERE GATEPASSNO = " & iGatePassNo
	else
		.Source = "SELECT OrganisationCode, TYPEOFITEMS, PARTYCODE, Convert(datetime,GENERATEDON,103) as markedon,ReferenceNo,ISNULL(DCCODE,'-'),isNull(Remarks,'') FROM FORGATEPASSHEADER WHERE GATEPASSNO = " & iGatePassNo
	end if
'	response.write dcrs.source
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sOrgID = trim(dcrs(0))
	sItemType = trim(dcrs(1))
	iPartyCode = trim(dcrs(2))
	dGatePassDate = trim(dcrs(3))
	sSourceRefNo = dcrs(4)
	sDCCode = dcrs(5)
	sRemarks = dcrs(6)
end if
dcrs.Close

if sSaleTransNo<>"" then
	with dcrs
		.cursorlocation = 3
		.cursortype = 3
		.source = "select Convert(datetime,InvoiceDate,103) From Sal_T_InvoiceHeader where SaleTransactionNo =  " & sSaleTransNo  & ""
		.ActiveConnection = con
		.open
	End With
	if not dcrs.eof then
		sInvoiceDate = dcrs(0)
	end if
	dcrs.close
end if 'if sSaleTransNo<>"" then

if trim(dGatePassDate)="" or IsNull(dGatePassDate) then
	'blocked on March 06,2010
	dGatePassDate = GetTransDate()
end if



if trim(sInvoiceDate)="" or isnull(sInvoiceDate) then
	sInvoiceDate = 	dGatePassDate
end if
'added on feb 2009
	'Response.Write sDCCode 
	dGatePassDate = FormatDate(dGatePassDate)
	sInvoiceDate = FormatDate(sInvoiceDate)
	
if trim(iPartyCode) = "" or isNull(iPartyCode) then 

	if trim(sSaleTransNo)<> "" then
		with dcrs
			.cursorlocation = 3
			.cursortype = 3
			.source = "select AgentCode from Sal_T_InvoiceHeader where SaleTransactionNo  = " & sSaleTransNo & "" 
			.ActiveConnection = con
			.open
		End With
		set dcrs.activeConnection = Nothing
		'Response.Write "<p>" & dcrs.source
		If not dcrs.eof Then
			iAgentCode = dcrs(0)
		End If	
		dcrs.close
		
	end if 'if trim(sSaleTransNo)<> "" then
end if 'if trim(iPartyCode) = "" then 
Response.Write "<font color=red>"
If sDCCode = "-" Then

    	sTempSeries = GetInvNumberSeriesCodes("DC",sOrgID,iNumClassCode)
	    sArrSeries = Split(sTempSeries,":")
	    iSeriesNo = sArrSeries(0)
	    iSeriesCode = sArrSeries(1)


	if Trim(iSeriesCode)="0" then
	
	    sQuery = "Select GroupName from INV_M_Classification where GroupCode = "& iNumClassCode
	    dcrs1.Open sQuery,con
	    if not dcrs1.EOF then
	        sNumClassName = Trim(dcrs1(0))
	    end if
	    dcrs1.Close 
	    
	    sDCNo = "NULL"
	    Response.Clear 
	    Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Gate Pass - "& sNumClassName &"  Classification </H2></p>"
	    Response.End 
	else
	    sDCNo = GetSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,dGatePassDate)
	    sDCCode = sDCNo 
	end if
	
Else
	sDCNo = sDCCode 
End If	
'Response.Write "<p color=red>"&sDCNo 
		'Response.Write sDCNo
'	sSql = "UPDATE INV_T_DIRECTISSUE SET GENERATEDFROM = 4,DICODE = " & sDCNo & " WHERE " & _
'		"DINUMBER = " & iGatePassNo & " AND DIFORUNIT = " & Pack(sOrgID) & ""
'	'con.Execute sSql
'	'Response.Write sSql
	set oDOM = nothing
%>
<% 


ChkStr = CheckFinYr(dGatePassDate)

 if ChkStr  = "3" then 	
%>
<SCRIPT LANGUAGE=javascript>
	alert("This transaction cannot be performed for this current Financial Year.");
	window.history.back(1);
</SCRIPT>
<%	
	elseif ChkStr = "2" then
	%>
<%
	elseif ChkStr = "1"  then
		'dGDate = ReqDate 
%>
<SCRIPT LANGUAGE=javascript>
	if (!confirm("Since Year End closing has been done and transaction date is in last FY this transaction will be accounted in current financial year. Do you want to proceed?")) {
		window.history.back(1);
	}
</SCRIPT>		
<%
	end if 
	
%>
</head>

<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" OnLoad="Init()">
<form method="POST" name="formname" action="">
<input type=hidden name="hGatePassNo" value="<%=iGatePassNo%>">
<input type=hidden name="hForSubConNo" value="<%=sForSubConNo%>">
<input type=hidden name="hDCNo" value="<%=sDCNo%>">
<input type=hidden name="hOrg" value="<%=sOrgID%>">
<input type=hidden name="hInvoiceType" value="<%=sInvoiceType%>">
<input type=hidden name="hSentType" value="<%=sSentType%>">
<input type=hidden name="hFinFrom" value="<%=sFinFrom%>">
<input type=hidden name="hFinTo" value="<%=sFinTo%>">
<input type=hidden name="hItemType" value="<%=sItemType%>">
<input type=hidden name="hGatePassDate" value="<%=dGatePassDate%>">
<input type=hidden name="hInvDate" value="<%=sInvoiceDate%>">
<input type=hidden name="hReturnToPage" value="<%=sReturnToPage%>">
<input type=hidden name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hNumClassCode" value="<%=iNumClassCode%>">



	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Gate Pass cum Delivery Challan
			</td>
		</tr>
		<tr>
			<td align="center" class="TopPack"></td>
		</tr>
		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack"></td>
								</tr>
								<tr>
									<td align="center"></td>
									<td width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td class="FieldCell" valign="bottom">To<br>
												<%
													with dcrs
														.CursorLocation = 3
														.CursorType = 3
														if sInvoiceType = "'U'" then
															.Source = "SELECT ORGUNITDESCRIPTION,ADDRESS1,ADDRESS2,CITY,POSTCODE,STATE FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = " & Pack(iPartyCode) & ""
														else
															if trim(iPartyCode) <> "" then
																.Source="SELECT PARTYNAME, ADDRESSLINE1, ADDRESSLINE2, CITY, PINCODE, STATE FROM APP_M_PARTYMASTER WHERE PARTYCODE =" & iPartyCode
															else
																.source= "SELECT PARTYNAME, ADDRESSLINE1, ADDRESSLINE2, CITY, PINCODE, STATE FROM APP_M_PARTYMASTER WHERE PartyCode= " & iAgentCode & ""
															end if 	
														end if
														.ActiveConnection = con
														.Open
													end with
													set dcrs.ActiveConnection=nothing
													if not dcrs.EOF then
														sUnitDesc = trim(dcrs(0))
														sUnitAddr = trim(dcrs(1))
														sUnitAddr1 = trim(dcrs(2))
														sUnitCity = trim(dcrs(3))
														sPostCode = trim(dcrs(4))
														sUnitState = trim(dcrs(5))
													end if
													dcrs.Close
												%>
													<span class="DataOnly">
														<%=sUnitDesc%><br>
														<%=sUnitAddr %><br>
														<%=sUnitAddr1 %><br>
														<%=sUnitCity%><br>
														<%=sUnitState%>
													</span>
												</td>
												<td valign="top">
													<table border="0" cellspacing="1" cellpadding="0" class="TableOutlineOnly">
														<tr>
															<td class="TableHeader"  align="center">Delivery Challan cum Gate Pass</td>
														</tr>
														<tr>
															<!--<td class="TableHeader" align="center">Number - Date</td>-->
															<td class="TableHeader" align="center">Number</td>
														</tr>
														<tr>
															<!--td class="ExcelDisplayCell" align="center"><%'=mid(sDCNo,2,(len(sDCNo)-2))%> - <%'=dGatePassDate%></td-->
															<!--<td class="ExcelDisplayCell" align="center"><%'=sDCNo%> - <%'=dGatePassDate%></td>-->
															<td class="ExcelDisplayCell" align="center"><%=sDCNo%></td>
														</tr>
													</table>
												</td>
												<td>
													<table>
														<tr>
														<td class="FieldCell">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Date :&nbsp;</td>
														<td class="FieldCellSub"><%Response.Write InsertDatePicker("ctlGatePassDate") %></td>
														</tr>
													</table>
												</td>
											</tr>
										</table>
									</td>
									<td align="center"></td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="MiddlePack"></td>
								</tr>
								<tr>
									<td align="center"></td>
									<td width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td class="FieldCell" width="55">Reference</td>
												<td class="FieldCellSub"><%=sReference%></td>
											</tr>
										</table>
									</td>
									<td align="center"></td>
								</tr>
								<tr>
									<td align="center"></td>
									<td width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td class="FieldCell" valign="top" width="55">Remarks</td>
												<td class="FieldCellSub" valign="top">
													<p align="left" >
													<textarea class="formelem" name="txtRemarks"  rows=3 cols=50 maxlength=250><%=sRemarks%></textarea>
												</td>
											</tr>
										</table>
									</td>
									<td align="center"></td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="MiddlePack"></td>
								</tr>
								<tr>
									<td align="center"></td>
									<td width="100%">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Item Description</td>
												<td class="ExcelHeaderCell" align="center" width="100">Quantity</td>
												<td class="ExcelHeaderCell" align="center" width="100">UoM</td>
												<% If sSentType = "Y" and sInvoiceType = "'V'" Then %>												
												<td class="ExcelHeaderCell" align="center" width="100">Return</td>
												<% End If %>
												
												
												
											</tr>
										<% Dim PackType, PackCode,cnt,sRemarks
										
											sApplFormJJ = "N"
											
											cnt = 0
											sPackNos = ""
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT IsNull(ITEMCODE,0), isNull(CLASSIFICATIONCODE,0), QUANTITY,Description,EntryNo,MaterialRcvd,InvoicedUoM,NoofPacks,PackingType,isNull(FormJJ,'N') as FormJJ,isNull(ItemAttributes,0) FROM FORGATEPASSDETAILS WHERE GATEPASSNO = " & iGatePassNo
												.ActiveConnection = con
												.Open
											end with
											'Response.Write "<p>" & dcrs.Source 
											set dcrs.ActiveConnection = nothing
											do while not dcrs.EOF
											
												if trim(dcrs("FormJJ")) = "Y" then
													sApplFormJJ = "Y"
												end if 
													
												iCtr = iCtr + 1
												If cDbl(dcrs(0)) <> 0 Then
													sItemName = ItemDisplay(trim(dcrs(0)),trim(dcrs(1))) & "-" & dcrs(3)
												Else
													sItemName = dcrs(3)																									
												End If
												sAttID = dcrs(10)
												'Response.Write "sAttID  = "& sAttID
												if trim(sAttID)<>"" then
												    sAttID = split(sAttID,":")(0)
												    sArrID = Split(sAttID,"#")
												    if UBound(sArrID)=1 then
													    sAttID = sArrID(1)
												    else
													    sAttID = sArrID(0)
												    end if
												 end if
												'response.write "sAttID = "& sAttID	
												
												if trim(sAttID)<>"" and trim(sAttID)<>"0" then
													sQuery = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "& sAttID
													dcrs1.open sQuery,con
													if not dcrs1.eof then
														sOptName = trim(dcrs1(0))
													end if
													dcrs1.close
												end if
												
												If sItemType = "STO" Then
														With dcrs1
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT CompanyItemCode FROM inv_M_ItemMaster where itemcode = "& Trim(dcrs(0))
															.ActiveConnection = con
															.Open
														end with	
														If Not dcrs1.EOF Then
															sItemName = dcrs1(0) & "-" &  sItemName
														End If
														dcrs1.Close 	
												End If												
												'Response.Write dcrs("NoofPacks")]
												'Response.Write "<p>"& sInvoiceType & "<p>"
												If sInvoiceType = "'A'" or sInvoiceType = "A" Then
														With dcrs1
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT distinct Count(*),PackingCode  FROM Sal_t_invPackDetails GROUP BY SaleTransactionNo,ItemCode,ClassificationCode,PackingCode HAVING Sal_t_invPackDetails.SaleTransactionNo = " & sSourceRefNo
															.ActiveConnection = con
															.Open
														end with	
														'Response.Write dcrs1.Source 											
														If Not dcrs1.EOF Then
															sItemName = sItemName & "-" & UCase(dcrs1(0))
															PackCode = dcrs1(1)
														End If
														dcrs1.Close 
														If PackCode <>"" Then
															with dcrs1
																.CursorLocation = 3
																.CursorType = 3
																.Source = "SELECT PackingShortName FROM App_M_PackingType WHERE PackingCode = " & PackCode 
																.ActiveConnection = con
																.Open
															end with												
															If Not dcrs1.EOF Then
																sItemName = sItemName & "-" & UCase(dcrs1("PackingShortName"))
															End If
															dcrs1.Close 
														End If
														'Response.Write sItemType & "<p>"											
														If sItemType = "FIB" Then
															With dcrs1
																.CursorLocation = 3
																.CursorType = 3
																.Source = "SELECT distinct packNettWeight, PackNumber FROM Sal_t_invPackDetails where Sal_t_invPackDetails.SaleTransactionNo = " & sSourceRefNo & " order by packnumber"
																.ActiveConnection = con
																.Open
															end with	
															If Not dcrs1.EOF Then
																While Not dcrs1.EOF 
																If sPackNos = "" Then 
																	sPackNos = formatnumber(dcrs1(0),2,,0)
																	cnt = cnt + 1
																	
																Else
																	If cnt = 13 Then
																		sPackNos = sPackNos & "," & formatnumber(dcrs1(0),2,,0) & "<BR>"
																	Else
																		sPackNos = sPackNos & "," & formatnumber(dcrs1(0),2,,0)
																	End If	
																	cnt = cnt + 1																	
																End If	
																If cnt = 14 Then cnt = 0
																dcrs1.MoveNext 
																Wend
															End If
															dcrs1.Close 
															sItemName = sItemName & "<BR>" & sPackNos 															
														End If														
														
														
														
												ElseIf sInvoiceType = "'T'" or sInvoiceType = "T" Then
														With dcrs1
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT distinct Count(*),PackingCode  FROM Sal_T_DepotInvoicePacking GROUP BY DepotTransferNo,ItemCode,ClassificationCode,PackingCode HAVING Sal_T_DepotInvoicePacking.DepotTransferNo = " & iGatePassNo
															.ActiveConnection = con
															.Open
														end with												
														If Not dcrs1.EOF Then
															sItemName = sItemName & "-" & UCase(dcrs1(0))
															PackCode = dcrs1(1)
														End If
														dcrs1.Close 
														If PackCode <>"" Then
														with dcrs1
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT PackingName FROM App_M_PackingType WHERE PackingCode = " & PackCode 
															.ActiveConnection = con
															.Open
														end with												
														If Not dcrs1.EOF Then
															sItemName = sItemName & "-" & UCase(dcrs1("PackingName"))
														End If
														dcrs1.Close 	
														End If											
												ElseIf sInvoiceType = "'V'" Then												
													If dcrs("NoofPacks")<>"" Then
														sItemName = sItemName &"-"&dcrs("NoofPacks")
													End if
													If dcrs("packingtype")<>"" Then
														with dcrs1
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT PackingName FROM App_M_PackingType WHERE PackingCode = " & dcrs("PackingType")
															.ActiveConnection = con
															.Open
														end with												
														If Not dcrs1.EOF Then
															sItemName = sItemName & "-" & UCase(dcrs1("PackingName"))
														End If
														dcrs1.Close 
													End if
												End If
													iQty = trim(dcrs(2))
													sUoM = trim(dcrs(6))												
												
												if trim(sOptName)<>"" then 
													sItemName = sItemName &" ["& sOptName&"]"
												end if
												
										%>
										
											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td	class="ExcelDisplayCell"><%=sItemName%></td>
												<td class="ExcelDisplayCell" align="right"><%=iQty%></td>
												<td class="ExcelDisplayCell" align="Center"><%=sUoM%></td>
												<% If sSentType = "Y" and sInvoiceType = "'V'" Then 
														If dcrs(5) = "Y" Then	%>
															<td class="ExcelDisplayCell" align ="center"><input type="Button" name="BtnA<%=iCtr%>" class="AddButtonX" value="Yes" onClick="Check(<%=iGatePassNo%>,<%=CDbl(dcrs(4))%>)" disabled></td>
														<% Else %>
															<td class="ExcelDisplayCell" align ="center"><input type="Button" name="BtnA<%=iCtr%>" class="AddButtonX" value="Yes" onClick="Check(<%=iGatePassNo%>,<%=CDbl(dcrs(4))%>)"></td>
														<% End If%>													
												<% End If %>																								
											</tr>											
										<%
												dcrs.MoveNext
											loop
											dcrs.Close
										%>
										</table>
									</td>
									<td align="center"></td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td>
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell" align=center>
												<% if sSentType = "Y" Then %>
													<input type="button" value="Modifiy" class="ActionButton" onclick="CheckModifiy()" id=button1 name=button1>
												<%Else %>
													<%if trim(sDCCode) ="-" then %>
													<input type="button" value="Save" class="ActionButton" onclick="CheckSubmit()" id=button1 name=button1 disabled>
													<%else%>
													<input type="button" value="Save" class="ActionButton" onclick="CheckSubmit()" id=button1 name=button1 >
													<%end if%>
												<% End If %>	
													<input type="button" value="Print" class="ActionButton" onClick="Print()" id=button2 name=button2>
												
												
												<%if sApplFormJJ = "Y" then %>	
													<input type="button" value="Print Form JJ" class="ActionButtonX" onClick="FormJJPrint()" id=buttonFormJJ name=buttonFormJJ >
												<%end if %>	
												<input type="button" value="Cancel" class="ActionButton" onclick="Cancel('GatePassSelection.asp')" id=button3 name=button3>
												</td>
											</tr>
										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="BottomPack"></td>
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
</HTML>


<%
	' Function to Check for Fin. Year
	Function CheckFinYr(dDate)
		' Declaration of variables
		Dim dcrs
		dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sMonYrNew
		dim sCurYear, sCurYearFrom, sCurYearTo
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		'Response.Write dDate & "        "
		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(dDate)
		else
			sTempMonYr = Month(dDate)
		end if
		sMonYr = sTempMonYr&Year(dDate)
		'Response.Write dDate &"," & sFinPeriodFrom 
		arrFin = split(GetFinancialYear(sMonYr),":")
		sCurYearFrom  = arrFin(0)
		sCurYearTo = arrFin(1)
		'Response.Write sFinPeriodFrom
		'Response.Write "dDate="& dDate  
		dDate = FormatDateTime(dDate) 
		'Response.Write DateDiff("d",FormatDate(sFinPeriodFrom),dDate) 
		If (DateDiff("d",FormatDate(sFinPeriodFrom),dDate) >= 0) and (DateDiff("d",FormatDate(sFinPeriodTo),dDate)<= 0) Then
			CheckFinYr = "2"
		ElseIf (DateDiff("d",FormatDate(sFinPeriodFrom),dDate)<=0) Then
			CheckFinYr = "1"
		Else	 	
			CheckFinYr = "3"	
		End If  
		If CheckFinYr = "1" Then 
			If (DateDiff("d",FormatDate(sFinPeriodTo),dDate))<=0 Then
				dGatePassDate = right("0" & trim(day(dDate)),2) & "/" & right("0"& trim(month(dDate)),2) &"/"&Year(dDate)
			Else	
				dGatePassDate = sFinancialYearTo 
			End If
		Else	
		End If
		sFinancialYearTo = arrFin(1)
		'If CheckFinYr = 2 Then dGatePassDate = Date()
	End Function
%>

