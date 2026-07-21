<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	IssReturnEntry.asp
	'Module Name				:	Inventory 
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 24,2011
	'Modified By				:	
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<%
	dim iCtr,arrTemp,sTemp,arrValue,sSource,sOrgID,sDept,sItemName,sRefDet,iIssueEntryNo
	Dim oDom,Root,newElem,rs, rs1, rs2, dcrs,dcrs1,sIssQty,sConQty
	Dim sIssueEntryCode,sIssuedForDescription,sIssueDate, sMRDRcode,sMRDRDate
	Dim sIssueType, sRetQty,newElem1, newElem2,newElem3, sWorkCenterCode, sRefType
	Dim iSourceRef, iIssueNo,sIssuedForCode,iMRS,iDI,sQuery,sBalQty,sItemType
	Dim sAppRefNo,sAppRefDate,sAppRefType,sAppRefName,sAppRefNoDate,sAttribute,sRcptNumbering
	Dim iItemCode,iClassCode,iItemEntNo
	Dim sIssuedToCode,sIssuedToType,sIssueSubCode
	
	Set rs = Server.CreateObject("ADODB.RecordSet")		
	Set rs1 = Server.CreateObject("ADODB.RecordSet")
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")	
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set Root = oDOM.createElement("root")
	
	Dim sFinPeriodFrom, sFinPeriodTo, sFinancialYearTo,ChkStr
	dim sFinPeriod, sFinFrom, sFinTo, sTempMonYr, sMonYr, arrFin, IssDate
	sFinPeriod = Session("FinPeriod")
	sFinPeriodFrom = FormatDate("04/01/" & Mid(sFinPeriod,1,4))
	sFinPeriodTo = FormatDate("03/31/" & Mid(sFinPeriod,6,4))
	sFinFrom = FormatDate("04/01/" & Mid(sFinPeriod,1,4))
	sFinTo = FormatDate("03/31/" & Mid(sFinPeriod,6,4))
	If DateDiff("d",FormatDate(Date()),FormatDate(sFinTo)) < 0 Then
		If len(Month(sFinTo)) = 1 Then
			sTempMonYr = "0"&Month(sFinTo)
		Else
			sTempMonYr = Month(sFinTo)
		End If 
		'Response.Write sTempMonYr
		sMonYr = sTempMonYr&Year(sFinTo)
	Else
		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())
	End If

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)			
	
	sIssQty = 0 : sConQty = 0
	if Trim(Request("RefDet"))<>"" then
	  sRefDet = split(trim(Request("RefDet")),":")
				'sOrgID = trim(sRefDet(0))
				sTemp = ""
				iIssueEntryNo = trim(sRefDet(1))
				IssDate = trim(sRefDet(2))
	end if
	sOrgID = Session("organizationcode")
	
	if trim(iIssueEntryNo)="" or IsNull(iIssueEntryNo) then
	%>
	    <script>
	        alert("Please select the Issue from List Tab and Press Issue Return button");
	        window.history.back(1)
	    </script>
	<%
	end if
	
	sQuery = "Select isNull(IssueEntryCode,IssueEntryNo),Convert(varchar,IssueDate,103),isNull(AppRefType,0),isNull(AppRefNo,0),Convert(varchar,AppRefDate ,103),IssuedToType,IssuedToCode,IssuedToSubCode,ItemTypeID from INV_T_MaterialIssueHeader where IssueEntryNo = "&iIssueEntryNo
	rs.Open sQuery,con
	if not rs.EOF then
	    sIssueEntryCode = Trim(rs(0))
	    sIssueDate = trim(rs(1))
	    sAppRefType = trim(rs(2))
	    sAppRefNo = trim(rs(3))
	    sAppRefDate = trim(rs(4))
	    sIssuedToType = trim(rs(5))
	    sIssuedToCode = trim(rs(6))
	    sIssueSubCode = trim(rs(7))
	    sItemType = trim(rs(8))
	end if
	rs.Close 
	
	sQuery = "Select ReceiptNumbering from VWITEM where ItemCode in (Select ItemCode from INV_T_MaterialIssueDetails where IssueEntryNo = "& iIssueEntryNo &")"
	rs.Open sQuery,con
	if not rs.EOF then
	    sRcptNumbering = Trim(rs(0))
	end if
	rs.Close 
	
	if sAppRefType<>"" and sAppRefType <>"0" then
	    sQuery = "Select ReferenceName from VW_ReferenceTypes where ReferenceEntryNo = "& sAppRefType 
	    rs.Open sQuery,con
	    if not rs.EOF then
	        sAppRefName = Trim(rs(0))
	    end if
	    rs.Close 
	    sAppRefNoDate = sAppRefNo &" - "& sAppRefDate 
	else
	    sAppRefName ="N/A"
	    sAppRefNoDate = "N/A"
	end if

   
    sIssuedForDescription = IssuedToString(sIssuedToType,sIssuedToCode,sIssueSubCode)
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receipt New - Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
    <ROOT DEPT="<%=sIssuedForCode%>" SOURCE="N" ORGCODE="<%=sOrgID%>" STYPE="N" ITEMTYPE="<%=sItemType%>" PACKNUM="" SRCREFTYPE="N" SRCREFNO="" RCPTNUMBERINV="<%=sRcptNumbering%>"></ROOT>
</script>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/IssReturnEntry.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0>
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hIssueEntryNo" value="<%=iIssueEntryNo%>">
<input type=hidden name="hIssuedForCode" value="<%=sIssuedForCode%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Return
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>	
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
		        <tr>
				    <td height="20" valign="bottom">
					    <table border="0" cellpadding="0" cellspacing="0" >
						    <tr>
						        <td class="TabCell" valign="bottom" width="90">
								    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
									    <tr><a href="IssueMGMT.asp">
										    <td align="center">List
										    </td></a>
									    </tr>
								    </table>
							    </td>
						   	   <td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="MaterialIssueEntry.asp">
												<td align="center">Basic
												</td></a>
											</tr>
										</table>
									</td>
							    <td class="TabCurrentCell" valign="bottom" align="center" width="50">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr>
												<td align="center">Return
												</td>
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
				<TR>				
					<TD class="TabBody">
					
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
                            <% If trim(Request("RefDet")) <> "" Then   
                            %>
                            
                            		<td width="100%" colspan="4">
										<div align="left">
											<table border="0" cellspacing="0" cellpadding="0" width="560">
												<tr>
													 <td class="FieldCell">&nbsp; Issue Number&nbsp;</td>
													 <td class="FieldCellSub"><span class="DataOnly"><%=sIssueEntryCode%>&nbsp;</span></td>
													 <td class="FieldCell">&nbsp Issue Date&nbsp;</td>
													 <td class="FieldCellSub"><span class="DataOnly"><%=sIssueDate%>&nbsp;</span></td>
												</tr>
												<tr>
													 <td class="FieldCell">&nbsp; Reference Name&nbsp;</td>
													 <td class="FieldCellSub"><span class="DataOnly"><%=sAppRefName%>&nbsp;</span></td>
													 <td class="FieldCell">&nbsp Reference No - Date&nbsp;</td>
													 <td class="FieldCellSub"><span class="DataOnly"><%=sAppRefNoDate%></span></td>
												</tr>													
												<tr>
													<td class="FieldCell">&nbsp Issued To&nbsp;</td>
													<td class="FieldCellSub"><span class="DataOnly"><%=sIssuedForDescription%>&nbsp;</span></td>																		
												</tr>																						
											</table>
											</div>
									</td>
					
                            <%
								End If
                            %>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td>
	                                <div class="frmBody" id="frm2" style="width: 100%; height:350;">
					                    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" rowspan = "2" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" rowspan = "2" align="center">Item Description</td>												
												<td class="ExcelHeaderCell" colspan = "4" align="center">Quantity</td>
												<td class="ExcelHeaderCell" rowspan = "2" align="center">Remarks</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2"><p align="center">Add.Details</td>													
									        </tr>
									        <tr>
												<td class="ExcelHeaderCell" align="center" width="10">Issued</td>
												<td class="ExcelHeaderCell" align="center">Consumed</td>
												<td class="ExcelHeaderCell" align="center">Returned</td>												
												<td class="ExcelHeaderCell" align="center">Return</td>
									        </tr>
									        <% 
									        sQuery = " Select D.ItemCode,D.Classificationcode,D.OrganisationCode,SUM(D.QuantityIssued),"&_
							                             " SUM(D.QuantityConsumed),SUM(D.QuantityReturned),D.QuantityUOM,H.IssueEntryNo,"&_
							                             " IsNull(H.IssueEntryCode,H.IssueEntryNo),Convert(varchar,IssueDate,103),D.ItemEntryNo,D.ItemAttributes,H.IssuedToType,H.IssuedToCode,H.IssuedToSubCode from "&_
							                             " INV_T_MaterialIssueHeader H,INV_T_MaterialIssueDetails D  where H.IssueEntryNo = "&_
							                             " D.IssueEntryNo and (H.IssueDate >= Convert(datetime,'"& sFinPeriodFrom &"',103))"&_
							                             " and (H.IssueDate <=Convert(datetime,'"& sFinPeriodTo &"',103)) and D.IssueEntryNo = "& iIssueEntryNo &" "&_
							                             " Group by D.ItemCode,D.ClassificationCode,D.OrganisationCode,D.QuantityUOM,H.IssueEntryNo,"&_
							                             " H.IssueEntryCode,H.IssueDate,D.ItemEntryNo,D.ItemAttributes,H.IssuedToType,H.IssuedToCode,H.IssuedToSubCode Order by H.IssueEntryNo "
							                      'Response.Write sQuery
							                      rs.Open sQuery,con
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
							                                sIssueEntryCode = rs(8)
							                                IssDate = rs(9)
							                                iItemEntNo = rs(10)
							                                sAttribute = rs(11)
							                                sIssuedForCode = rs(12)
							                                if trim(sIssueEntryCode)="" then
							                                    sIssueEntryCode = iIssueEntryNo 
							                                end if
							                                sBalQty = cdbl(sIssQty) - cdbl(sConQty) - cdbl(sRetQty)
							                                if trim(sAttribute)="" or IsNull(sAttribute) then sAttribute = "NULL"
							                                %>
							                                    <tr>
							                                        <td class="ExcelSerial" align="center"><%=iCtr%></td>
							                                        <td class="ExcelDisplayCell" align="Left"><%'=sIssueEntryCode%><%'=IssDate%>
							                                        <%=GetItemName(iItemCode,iClassCode) %>
							                                            <input type=hidden name="txtIssNoZ<%=iCtr%>" value="0" size="5" class="FormElemRead">
							                                        </td>
												                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sIssQty,2)%></td>
												                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sConQty,2)%></td>
												                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sRetQty,2)%></td>
												                    <td class="ExcelDisplayCell" align="Right">
												                        <input type=text name="txtConsumeQtyZ<%=iCtr%>" value="0" size=5 class="FormElemRead">
												                        <%if sBalQty > 0 then %>
												                            <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" onClick="PackSelection('<%=iIssueEntryNo%>','<%=iItemCode%>','<%=iClassCode%>','<%=sOrgID%>','<%=iCtr%>','<%=sAttribute%>','<%=iItemEntNo%>')" >
												                        <%else %>
												                            <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" >
												                        <%end if %>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center">
												                        <input type=text name="txtRemarksZ<%=iCtr%>" value="" size=25 class="FormElem">
												                    </td>
												                    <td class="ExcelDisplayCell" align="center">
												                        <input type=button name="btnAddDetZ<%=iCtr%>" class="ActionButtonX" value="Yes" onclick="AddDetails('<%=iIssueEntryNo%>','<%=iItemCode%>','<%=iClassCode%>','<%=sOrgID%>','<%=iCtr%>','<%=iItemEntNo%>','<%=sAttribute%>','<%=sIssuedForCode%>')">
												                    </td>
												                </tr>
							                                <%
							                                rs.MoveNext 
							                            loop
							                      end if
											%>
										</table>
									</div>
								</td>
								<td align="center">
								    <input type=hidden name="hCtr" value="<%=iCtr%>">
								</td>
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
                                                    <input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('receiptNewEntry.asp')">
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
	' Function to populate Classification and Item
	Function populateClassItem(iClass,iItem)
		' Declaration of variables
		Dim dcrs,sItemDesc,sItemShDesc,sClassDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMDESCRIPTION,SHORTDESCRIPTION,GROUPNAME FROM INV_M_ITEMORGMASTER IM,INV_M_CLASSIFICATION IC WHERE IM.CLASSIFICATIONCODE = IC.GROUPCODE AND IM.CLASSIFICATIONCODE = " & iClass & " AND IM.ITEMCODE = " & iItem & " ORDER BY IM.ITEMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sItemDesc = dcrs(0)
		set sItemShDesc = dcrs(1)
		set sClassDesc = dcrs(2)
		
		if Not dcrs.EOF then
			populateClassItem = trim(sClassDesc)&" -- "&trim(sItemDesc)
		else
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
