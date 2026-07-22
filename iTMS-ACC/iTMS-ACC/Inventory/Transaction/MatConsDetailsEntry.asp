<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MatConsDetailsEntry.asp
	'Module Name				:	Inventory (Material Consumption)
	'Author Name				:	Ragavendran
	'Created On					:	Jun 17,2011
	'Modified By				:	
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	MacConDetInsert.asp
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<%
	dim iCtr,rsTemp,arrTemp,sTemp,arrValue,sOrgID,sDept,sIssueCode, sRefDet, iIssueEntryNo
	dim arrUoM,sUoMDesc,sUoMCode,sItemName,sAddFlag, sSql, sRefType
	dim sIssueDate,sIssuedToStr,sIssueType,rs2,sMRDRcode,sMRDRDate
	dim sIssQty, sConQty, sIssuedToCode, iIssueNo, iSourceRef,sIssuedToType,sIssuedToSubCode
	Dim iMRS, iDI, newElem1,sMCFlag,sWorkCenterCode,newElem2,newElem3
	Dim sRetQty,sBalQty,iItemEntNo,sAttribute,sIssueNo,sLotNo,iRowsPan,iOutputQty

	Dim oDom,Root,newElem,rs, rs1,dcrs,dcrs1
	Dim oDom1,Root1, sFinPeriodFrom, sFinPeriodTo, sFinancialYearTo,ChkStr
	dim sFinPeriod, sFinFrom, sFinTo, sTempMonYr, sMonYr, arrFin, IssDate
	Dim dFrmDate,dToDate,iItemCode,iClassCode
	Set rs = Server.CreateObject("ADODB.RecordSet")		
	Set rs1 = Server.CreateObject("ADODB.RecordSet")		
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set rsTemp = Server.CreateObject("ADODB.RecordSet")		
	
	dFrmDate = Request("FromDate")
	dToDate = Request("ToDate")
	iItemCode = Request("ItemCode")
	iClassCode = Request("ClassCode")
	sOrgID = Session("organizationcode")
	sIssueNo = trim(Request("IssueNo"))
	
	if trim(iItemCode)<>"" then
	    sItemName = GetItemName(iItemCode,iClassCode)
    end if
	
	if trim(sIssueNo)<>"" then
	    sSql = "Select IssueEntryCode,Convert(varchar,IssueDate,103),IssuedToType,IssuedToCode,IssuedToSubCode from INV_T_MaterialIssueHeader where IssueEntryNo = "& sIssueNo
	    rs.Open sSql,con
	    if not rs.eof then
	        sIssueCode =  rs(0)
	        sIssueDate = rs(1)
	        sIssuedToType = rs(2)
	        sIssuedToCode = rs(3)
	        sIssuedToSubCode = rs(4)
	    end if
	    rs.close
	    sIssuedToStr=IssuedToString(sIssuedToType,sIssuedToCode,sIssuedToSubCode)
	end if
	
	
	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Material Consumption - Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
    <Root OrgID="<%=sOrgID%>"></Root>
</script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<% ' Response.Write Formatdate(IssDate)
ChkStr = CheckFinYr(Formatdate(IssDate))
'Response.Write ChkStr 
 if ChkStr  = "3" then 	
%>
<%	
	elseif ChkStr = "2" then
	%>
<%
	elseif ChkStr = "1"  then
		'dGDate = ReqDate 
%>
<%
	end if 
	
'	oDOM.Save server.MapPath("../temp/transaction/RECEIPTEX"&Session.SessionID&".xml")
'	Response.Write sIssueDate  
%>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/matConsDetailsEntry.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0>
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hFromDate" value="<%=dFrmDate%>">
<input type=hidden name="hToDate" value="<%=dToDate%>">
<input type=hidden name="hItemCode" value="<%=iItemCode%>">
<input type=hidden name="hClassCode" value="<%=iClassCode%>">
<input type="hidden" name="hIssueNo" value="<%=sIssueNo%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Consumption
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
                                <td align="center"></td>
                        		<td width="100%">
									<div align="left" >
										<table border="0" cellspacing="0" cellpadding="0" width="100%">
										    <%if trim(sIssueNo)<>"" then %>
											<tr>
												 <td class="FieldCell" style="width:100px;">&nbsp; Issue No - Date&nbsp;</td>
												 <td class="FieldCellSub"><span class="DataOnly"><%=sIssueCode%>&nbsp;</span>&nbsp;<span class="DataOnly"><%=sIssueDate%>&nbsp;</span></td>
											</tr>
											<tr>
												<td class="FieldCell" style="width:100px;">&nbsp Issued To&nbsp;</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=sIssuedToStr%>&nbsp;</span></td>																		
											</tr>	
											<%end if'if trim(sIssueNo)<>"" then %>
											
											<%if trim(sItemName)<>"" then %>
											<tr>
												<td class="FieldCell" style="width:100px;">&nbsp;&nbsp;Item Description</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=sItemName%>&nbsp;</span></td>
											</tr>	
                                            <%end if'if trim(sItemName)<>"" then %>
                                            <tr>
											    <td class="FieldCell" style="width:120px;">&nbsp;&nbsp;Consumption Date</td>
											    <td class="FieldCellSub">
											    <%
												    ' Function Call to Insert Date Picker
												    Response.Write InsertDatePicker("ctlDDate")
											    %>
											    </td>	
											</tr>
										</table>
									</div>
								</td>
								<td align="center"></td>
					        </tr>
                            <tr>
								<td align="center"></td>
								<td>
	                                <div class="frmBody" id="frm2" style="height:380;">
					                    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
												<%if trim(sIssueNo)<>"" then%>
												    <td class="ExcelHeaderCell" align="center" rowspan="2"><p align="center">Item Description</td>
												<%else%>
												    <td class="ExcelHeaderCell" align="center" rowspan="2"><p align="center">Issue No-Date</td>
												<%end if%>
												<td class="ExcelHeaderCell" align="center" Colspan="4"><p align="center">Quantity</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2"><p align="center">Output</p></td>
												<td class="ExcelHeaderCell" align="center" rowspan="2"><p align="center">By Product</p></td>
												<td class="ExcelHeaderCell" align="center" rowspan="2"><p align="center">Remarks</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2"><p align="center">Add.<br />Details</td>												
											</tr>

											<tr>
												<td class="ExcelHeaderCell" align="center">Issued</td>
												<td class="ExcelHeaderCell" align="center">Consumed</td>
												<td class="ExcelHeaderCell" align="center">Returned</td>
												<td class="ExcelHeaderCell" align="center">Consumption</td>												
											</tr>
											<%
											    'Response.write "<p>sIssueNo="& sIssueNo
											    if trim(sIssueNo)<>"" then
											    
											        sSql = "Select count(*) from INV_T_MaterialIssueDetails where IssueEntryNo = "& sIssueNo
											        rs.Open sSql,con
											        if not rs.eof then
											            iRowsPan = rs(0)
											        else
											            iRowsPan = 0
											        end if
											        rs.Close
											    
										            sSql = " Select D.ItemCode,D.Classificationcode,D.OrganisationCode,SUM(D.QuantityIssued),"&_
						                                 " SUM(D.QuantityConsumed),SUM(D.QuantityReturned),D.QuantityUOM,D.IssueEntryNo "&_
						                                 "  from INV_T_MaterialIssueDetails D  where D.IssueEntryNo = "& sIssueNo
						                            sSql = sSql &" Group By D.ItemCode,D.ClassificationCode,D.OrganisationCode,D.QuantityUOM,D.IssueEntryNo"
'						                            response.write "<textarea>"& sSql &"</textarea>"
							                        rs.Open sSql,con
						                            if not rs.EOF then
						                                iCtr = 0
						                                do while not rs.EOF 
						                                    iCtr = iCtr + 1
						                                    iItemCode = rs(0)
						                                    iClassCode = rs(1)
						                                    sIssQty = rs(3)
						                                    sConQty = rs(4)
						                                    sRetQty = rs(5)
						                                    iIssueEntryNo = rs(7)
						                                    
						                                    sBalQty = cdbl(sIssQty) - cdbl(sConQty) - cdbl(sRetQty)
						                                    sSql= "Select IsNull(SUM(OutputQuantity),0) from INV_T_MaterialConsumptionOutput where IssueEntryNo ="& iIssueEntryNo
							                                    rsTemp.open sSql,con
							                                    if not rsTemp.eof then
							                                        iOutputQty = rsTemp(0)
							                                    end if
							                                    rsTemp.close
							                                    
						                                    %>
						                                        <tr>
						                                            <td class="ExcelSerial" align="center"><%=iCtr%></td>
						                                            <td class="ExcelDisplayCell" align="Left"><%=GetItemName(iItemCode,iClassCode) %>
						                                            </td>
											                        <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sIssQty,2)%></td>
											                        <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sConQty,2)%></td>
											                        <td class="ExcelDisplayCell" align="Right"><a href="#" class="ExcelDisplayLink" onclick="ReceiptSelectionIss('<%=iIssueEntryNo%>','<%=sOrgID%>')"><%=FormatNumber(sRetQty,2)%></a></td>
											                        <%if iCtr = 1 then %>
											                        <td class="ExcelDisplayCell" align="Right" rowspan="<%=iRowsPan%>">
											                            <input type=text name="txtConsumeQtyZ<%=iIssueEntryNo%>" value="0" size=5 class="FormElemRead">
											                            <%if sBalQty > 0 then %>
											                                <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" onClick="PackSelectionIss('<%=iIssueEntryNo%>','<%=sOrgID%>')" >
											                            <%else %>
											                                <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" >
											                            <%end if %>
											                        </td>
											                        <td class="ExcelDisplayCell" align="Right" rowspan="<%=iRowsPan%>">
											                         <%
										                                if cdbl(iOutputQty)>0 then
										                                    response.write iOutputQty
										                                end if
										                             %>
										                             &nbsp;
											                            <input type=text name="txtOPQtyZ<%=iIssueEntryNo%>" value="0" size=5 class="FormElemRead">
											                                <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" onClick="ReceiptSelectionIss('<%=iIssueEntryNo%>','<%=sOrgID%>')" >
											                            <%if sBalQty > 0 then %>
											                                
											                            <%else %>
											                                <!--<img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" >-->
											                            <%end if %>
											                        </td>
											                        <td class="ExcelDisplayCell" align="Right" rowspan="<%=iRowsPan%>">
											                            <input type=text name="txtBPQtyZ<%=iIssueEntryNo%>" value="0" size=5 class="FormElemRead">
											                            <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" onClick="ReceiptSelectionIss('<%=iIssueEntryNo%>','<%=sOrgID%>')" >
											                            <%if sBalQty > 0 then %>
											                                
											                            <%else %>
											                                <!--<img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" >-->
											                            <%end if %>
											                        </td>
											                        <td class="ExcelDisplayCell" align="center" rowspan="<%=iRowsPan%>">
											                            <input type=text name="txtRemarksZ<%=iIssueEntryNo%>" value="" size=25 class="FormElem">
											                        </td>
											                        <td class="ExcelDisplayCell" align="center" rowspan="<%=iRowsPan%>">
											                            <input type=button name="btnAddDetZ<%=iIssueEntryNo%>" class="ActionButtonX" value="Yes" onclick="AddDetailsIss('<%=iIssueEntryNo%>','<%=sOrgID%>','<%=sIssuedToCode%>')">
											                        </td>
											                        <%end if 'if iCtr = 1 then %>
											                    </tr>
						                                    <%
						                                    rs.MoveNext 
						                                loop
						                            end if
											    else ' if trim(sIssueNo)="" then
											      
											          sSql = " Select D.ItemCode,D.Classificationcode,D.OrganisationCode,SUM(D.QuantityIssued),"&_
							                                 " SUM(D.QuantityConsumed),SUM(D.QuantityReturned),D.QuantityUOM,H.IssueEntryNo,"&_
							                                 " IsNull(H.IssueEntryCode,H.IssueEntryNo),Convert(varchar,IssueDate,103),H.IssuedToCode,H.IssuedToType,H.IssuedToSubCode from "&_
							                                 " INV_T_MaterialIssueHeader H,INV_T_MaterialIssueDetails D  where H.IssueEntryNo = "&_
							                                 " D.IssueEntryNo and (H.IssueDate >= Convert(datetime,'"& dFrmDate &"',103))"&_
							                                 " and (H.IssueDate <=Convert(datetime,'"& dToDate &"',103)) and D.ItemCode IN "&_
							                                 " (" & iItemCode & ") and D.ClassificationCode IN(" & iClassCode & ")"&_
							                                 " Group by D.ItemCode,D.ClassificationCode,D.OrganisationCode,D.QuantityUOM,H.IssueEntryNo,"&_
							                                 " H.IssueEntryCode,H.IssueDate,H.IssuedToCode,H.IssuedToType,H.IssuedToSubCode Order by H.IssueEntryNo "
    							                      'response.write "<textarea>"& sSql &"</textarea>"
							                          rs.Open sSql,con
							                          if not rs.EOF then
							                            iCtr = 0
							                                do while not rs.EOF 
							                                    iCtr = iCtr + 1
							                                    sIssQty = rs(3)
							                                    sConQty = rs(4)
							                                    sRetQty = rs(5)
							                                    iIssueEntryNo = rs(7)
							                                    sIssueCode = rs(8)
							                                    IssDate = rs(9)
							                                    sIssuedToCode = rs(10)
							                                    sIssuedToType = rs(11)
							                                    sIssuedToSubCode= rs(12)
							                                    if trim(sIssueCode)="" then
							                                        sIssueCode = iIssueEntryNo 
							                                    end if
							                                    sBalQty = cdbl(sIssQty) - cdbl(sConQty) - cdbl(sRetQty)
							                                    
							                                    sSql= "Select IsNull(SUM(OutputQuantity),0) from INV_T_MaterialConsumptionOutput where IssueEntryNo ="& iIssueEntryNo&" and ItemCode = "& iItemCode
							                                    rsTemp.open sSql,con
							                                    if not rsTemp.eof then
							                                        iOutputQty = rsTemp(0)
							                                    end if
							                                    rsTemp.close
							                                    %>
							                                        <tr>
							                                            <td class="ExcelSerial" align="center"><%=iCtr%></td>
							                                            <td class="ExcelDisplayCell" align="Left"><%=sIssueCode%>-<%=IssDate%>
							                                                <input type=hidden name="txtIssNoZ<%=iCtr%>" value="0" size="5" class="FormElemRead">
							                                            </td>
												                        <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sIssQty,2)%></td>
												                        <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sConQty,2)%></td>
												                        <td class="ExcelDisplayCell" align="Right"><a href="#" class="ExcelDisplayLink" onclick="ReceiptSelection('<%=iIssueEntryNo%>','<%=iItemCode%>','<%=iClassCode%>','<%=sOrgID%>','<%=iCtr%>')"><%=FormatNumber(sRetQty,2)%></a></td>
												                        <td class="ExcelDisplayCell" align="Right">
												                            <input type=text name="txtConsumeQtyZ<%=iCtr%>" value="0" size=5 class="FormElemRead">
												                            <%if sBalQty > 0 then %>
												                                <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" onClick="PackSelection('<%=iIssueEntryNo%>','<%=iItemCode%>','<%=iClassCode%>','<%=sOrgID%>','<%=iCtr%>')" >
												                            <%else %>
												                                <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" >
												                            <%end if %>
												                        </td>
												                         <td class="ExcelDisplayCell" align="Right">
												                         <%
												                            if cdbl(iOutputQty)>0 then
												                                response.write iOutputQty
												                            end if
												                         %>
												                         &nbsp;
											                            <input type=text name="txtOPQtyZ<%=iCtr%>" value="0" size=5 class="FormElemRead">
											                            <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" onClick="ReceiptSelection('<%=iIssueEntryNo%>','<%=iItemCode%>','<%=iClassCode%>','<%=sOrgID%>','<%=iCtr%>')" >
											                            <%if sBalQty > 0 then %>
											                                
											                            <%else %>
											                                <!--<img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" >-->
											                            <%end if %>
											                        </td>
											                        <td class="ExcelDisplayCell" align="Right">
											                            <input type=text name="txtBPQtyZ<%=iCtr%>" value="0" size=5 class="FormElemRead">
											                            <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" onClick="ReceiptSelection('<%=iIssueEntryNo%>','<%=iItemCode%>','<%=iClassCode%>','<%=sOrgID%>','<%=iCtr%>')" >
											                            <%if sBalQty > 0 then %>
											                                
											                            <%else %>
											                                <!--<img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" >-->
											                            <%end if %>
											                        </td>
												                        <td class="ExcelDisplayCell" align="center">
												                            <input type=text name="txtRemarksZ<%=iItemCode%>Z<%=iClassCode%>Z<%=iIssueEntryNo%>" value="" size=25 class="FormElem">
												                        </td>
												                        <td class="ExcelDisplayCell" align="center">
												                            <input type=button name="btnAddDetZ<%=iCtr%>" class="ActionButtonX" value="Yes" onclick="AddDetails('<%=iIssueEntryNo%>','<%=iItemCode%>','<%=iClassCode%>','<%=sOrgID%>','<%=iCtr%>','<%=sIssuedToCode%>')">
												                        </td>
												                    </tr>
							                                    <%
							                                    rs.MoveNext 
							                                loop
							                            end if
							                    end if'if trim(sIssueNo)<>"" then
											%>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="reset" value="Reset" name="B1" class="ActionButton">
                                                    <!--<input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('MatConsEntry.asp')">-->
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
	' Function to populate Store
	Function DisplayUoM(sOrgID,iClass,iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE IN (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & ")"
			'Response.Write dcrs.source
			.ActiveConnection = con
			.Open
		end with
		'Response.Write  dcrs.source
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>


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
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())
		'Response.Write dDate &"," & FormatDate(sFinPeriodFrom) 
		arrFin = split(GetFinancialYear(sMonYr),":")
		sCurYearFrom  = arrFin(0)
		sCurYearTo = arrFin(1)
		'Response.Write DateDiff("d",FormatDate(sFinPeriodFrom),dGDate) 
		If (DateDiff("d",FormatDate(sFinPeriodFrom),dDate) >= 0) and (DateDiff("d",FormatDate(sFinPeriodTo),dDate)<= 0) Then
			CheckFinYr = "2"
		ElseIf (DateDiff("d",FormatDate(sFinPeriodFrom),dDate)<=0) Then
			CheckFinYr = "1"
		Else	 	
			CheckFinYr = "3"	
		End If  
		If CheckFinYr = "1" Then 
			If (DateDiff("d",FormatDate(sFinPeriodTo),date()))<=0 Then
				IssDate = FormatDate(date())
			Else	
				IssDate = sFinancialYearTo 
			End If
		End If
		sFinancialYearTo = arrFin(1)
		'Response.Write dGDate 
	End Function
%>
