
<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	ItemStockOpening.asp
	'Module Name				:	INVENTORY (Item List)
	'Author Name				:
	'Created On					:
	'Modified By				:	Ragavendran R
	'Modified On				:	Oct 19,2011
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
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!-- #include File="../../include/ItemDisplay.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Grid</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemDetails"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="StockData"><Root></Root></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/printwindow.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function Search()
    if document.formname.ChkPur.checked = true then
        sEligible = sEligible &","& document.formname.ChkPur.value
    end if
    if document.formname.ChkSales.checked = true then
        sEligible = sEligible &","& document.formname.ChkSales.value
    end if
    if document.formname.ChkInv.checked =true then
        sEligible = sEligible &","& document.formname.ChkInv.value
    end if
    if document.formname.ChkManu.checked =true then
        sEligible = sEligible &","& document.formname.ChkManu.value
    end if
    sEligible = mid(sEligible,2)
    document.formname.hEligibleFor.value = sEligible
	document.formname.submit()
End Function


Function Sort(nFieldNo,sOrderByField,sOrder)

	eval("document.formname.hField" +  trim(nFieldNo)).value = trim(sOrderByField) & ":" & trim(sOrder)
	
	document.formname.hFieldSelected.value = nFieldNo
	
	document.formname.submit
End Function
'**************************************
Function CalculateTotalStock(iItemCtr,iRowCtr,iLocBinCtr)
Dim iIssQty,iOpRate,iStkQty
   iIssQty = eval("document.formname.hIssQtyZ"&iItemCtr&"Z"&iRowCtr&"Z"&iLocBinCtr).value
   iOpRate = eval("document.formname.hOpRateZ"&iItemCtr&"Z"&iRowCtr&"Z"&iLocBinCtr).value
   iStkQty = eval("document.formname.txtStkQtyZ"&iItemCtr&"Z"&iRowCtr&"Z"&iLocBinCtr).value
   if Trim(iOpRate)<>"0" then
        eval("document.formname.txtStkValueZ"&iItemCtr&"Z"&iRowCtr&"Z"&iLocBinCtr).value = CDbl(iStkQty)*CDbl(iOpRate)
   end if 'if Trim(iOpRate)<>"0" then
   eval("document.formname.txtTotQtyZ"&iItemCtr&"Z"&iRowCtr&"Z"&iLocBinCtr).value = CDbl(iIssQty)+CDbl(iStkQty)
   
End Function
'***************************************
Function EnableStock(iItemCtr)
    Dim nLocBinCtr,iCnt,nTotCnt
    nTotCnt  = Eval("document.formname.hTotalRow"&iItemCtr).value
    if Eval("document.formname.ChkBox"&iItemCtr).checked = true then
        For iCnt = 1 to nTotCnt
            nLocBinCtr = eval("document.formname.hLocBinCtrZ"&iItemCtr&"Z"&iCnt).value
            eval("document.formname.txtStkQtyZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).className="FormElem"
            eval("document.formname.txtStkValueZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).className="FormElem"
        Next
    else
        For iCnt = 1 to nTotCnt
            nLocBinCtr = eval("document.formname.hLocBinCtrZ"&iItemCtr&"Z"&iCnt).value
            eval("document.formname.txtStkQtyZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).className="FormElemRead"
            eval("document.formname.txtStkValueZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).className="FormElemRead"
        Next
    end if
End Function
'**************************************************
Function CheckSubmit()
    Dim nLocBinCtr,iCnt,nTotCnt,nTotItem,iItemCtr,iSelectItem,nLocBinVal
    Dim ndRoot,ndLocStock,ndItem,sArrChkVal,sArrLocBin
    nTotItem = document.formname.hTotItem.value 
    iSelectItem = 0
    set ndRoot = StockData.documentElement
    
    For iItemCtr = 1 to nTotItem 
        nTotCnt  = Eval("document.formname.hTotalRow"&iItemCtr).value
        if Eval("document.formname.ChkBox"&iItemCtr).checked = true then
            iSelectItem = iSelectItem + 1   
            sArrChkVal = Split(Eval("document.formname.ChkBox"&iItemCtr).value,":")
            set ndItem = StockData.createElement("Item")
                ndItem.setAttribute "ItemCode",sArrChkVal(0)
                ndItem.setAttribute "ClassCode",sArrChkVal(1)
                ndItem.setAttribute "OrgCode",sArrChkVal(2)
                ndItem.setAttribute "CompItemCode",sArrChkVal(3)
                ndItem.setAttribute "Desc",sArrChkVal(4)
                ndRoot.appendChild ndItem 
            For iCnt = 1 to nTotCnt
                nLocBinCtr = eval("document.formname.hLocBinCtrZ"&iItemCtr&"Z"&iCnt).value
                nLocBinVal = eval("document.formname.hLocBinValZ"&iItemCtr&"Z"&iCnt).value
                sArrLocBin = split(nLocBinVal,":")
                set ndLocStock = StockData.createElement("Loc")
                    ndLocStock.setAttribute "Loc",sArrLocBin(0)
                    ndLocStock.setAttribute "Bin",sArrLocBin(1)
                    ndLocStock.setAttribute "TotChangeQty",eval("document.formname.txtTotQtyZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).value
                    ndLocStock.setAttribute "Rate",eval("document.formname.hOpRateZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).value
                    ndLocStock.setAttribute "StkChange",eval("document.formname.txtStkQtyZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).value
                    ndLocStock.setAttribute "StkValue",eval("document.formname.txtStkValueZ"&iItemCtr&"Z"&iCnt&"Z"&nLocBinCtr).value
                    ndItem.appendChild ndLocStock
            Next
        end if
    Next 'For iItemCtr = 1 to nTotItem 
    
    if iSelectItem> 0 then
        set objhttp = CreateObject("Microsoft.XMLHTTP")
        objhttp.open "POST","XMLSave.asp?SessionFlag=true&Name=ItemOpenStockChange_",false
        objhttp.send StockData.XMLDocument
        
        document.formname.action ="ItemStockOpenInsert.asp"
        document.formname.submit 
    end if
End Function
</script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<%
	'Declaring Variables
	Dim sUnitID,sItemType,sSql,iCnt,rsStock,sField1,sField2,sField3,sSortBy
	Dim sCat,sSearchBy,sFilter,sClass,rsTemp,sOrgName,sGroupBy,sQuery
	Dim sEligibleFor,Arr1,nFieldSelected,nPrevItemCode,nRowSpan,nLocBinCtr
	dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,sItemDesc,sSqlBasic,sCondition
	Dim sListOfItemCode,sListOfClassCode,sListOfUnit,sBinName,sDisplayData
	Dim iOpenStok,iOpenValue,iIssueStock,iBalanceStock,iOpenRate,iBalanceValue,iItemCnt

	set rsStock=server.CreateObject("ADODB.Recordset")
	set rsTemp=server.CreateObject("ADODB.Recordset")
	
	
	sListOfItemCode		= Request.Form("hItemCode")
	sListOfClassCode	= Request.Form("hClassCode")
	sListOfUnit			= Request.Form("hUnitID")
	
	sUnitID = Session("organizationcode")
	sSql = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sUnitID
	rsStock.Open sSql,con
	if not rsStock.EOF then
		sOrgName = trim(rsStock(0))
	end if
	rsStock.Close

	'sUnitID = Request("selUnitId")
	sItemType = Request("selItemType")
		
	'sField1  = Request("hField1")
	'sField2  = Request("hField2")
	'sField3  = Request("hField3")
	
	'if trim(sField1) = "" then sField1 = "I:A"
	'if trim(sField2) = "" then sField2 = "D:A" 
	'if trim(sField3) = "" then sField3 = "C:A"  

	sField1  = ""
	sField2  = ""
	sField3  = ""
	
	nFieldSelected = trim(Request.Form("hFieldSelected"))
	if trim(nFieldSelected) = "" then nFieldSelected = 0
	
	if nFieldSelected = "1"  then
		sField1  = Request("hField1")
	end if
	
	if nFieldSelected = "2" then
		sField2  = Request("hField2")	
	end if
	
	if nFieldSelected = "3" then
		sField3  = Request("hField3")
	end if
	
	'--------------------
	if nFieldSelected = "0" then
		sField1  = "I:A"
	end if
	
	sCat = Request("selCategory")
	sSearchBy = Request("chkSearch")
	sFilter = trim(Request("txtSearch"))
	
	sEligibleFor = trim(Request("hEligibleFor"))
	if trim(sEligibleFor)="" then sEligibleFor = "I"

	if sCat = "select" then sCat = ""

	sTempMonYr = mid(FormatDate(date()),4,2)
	sMonYr = sTempMonYr&Year(FormatDate(date()))

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)

	if sUnitID = "" then sUnitID = "010101"
	if sItemType = "" then sItemType = "STO"
	

	''''''''''''''''''''' Paging Declaration ''''''''''''''''''''''''''''''''''''''''
    Const iPageSize=15	'How many records to show
    Dim iCurrentPage	'Current Page No.
    Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
    Dim iPageCtr		'Counter
	Dim lnPage

    iCurrentPage = CInt(Request.Form("hPageSelection"))
    'if iCurrentPage = "" or iCurrentPage = "0" then iCurrentPage = "1"
    'iCtr = (Cint(iPageSize) * (iCurrentPage - 1))

    con.CursorLocation = 3

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	sSortBy = ""
	
	if trim(sField1) <> ""  then
		if instr(1,sField1,":") > 0 then
			Arr1 = Split(sField1,":")
			
			if Arr1(1) = "A" then
				sSortBy = "COMPANYITEMCODE"
			else
				sSortBy = "COMPANYITEMCODE desc "
			end if
		end if 
	end if
	
	
	if trim(sField2) <> ""  then
		if instr(1,sField2,":") > 0 then
			Arr1 = Split(sField2,":")
			
			if Arr1(1) = "A" then
				'sSortBy = sSortBy  & ",ITEMDESCRIPTION"
				sSortBy = sSortBy  & "ITEMDESCRIPTION"
			else
				'sSortBy = sSortBy  & ",ITEMDESCRIPTION desc "
				sSortBy = sSortBy  & "ITEMDESCRIPTION desc "
			end if
		end if 
	end if
	
	if trim(sField3) <> ""  then
		if instr(1,sField3,":") > 0 then
			Arr1 = Split(sField3,":")
			
			if Arr1(1) = "A" then
				'sSortBy = sSortBy  & ",CategoryCode"
				sSortBy = sSortBy  & "CategoryCode"
			else
				'sSortBy = sSortBy  & ",CategoryCode desc "
				sSortBy = sSortBy  & "CategoryCode desc "
			end if
		end if 
	end if 
	 
	'Response.Write "<p>data="&Request.ServerVariables("SCRIPTNAME")

%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  >
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
		<input type="hidden" name="hOrgId" value="<%=sUnitID%>">
		<input type="hidden" name="hItemTypeCode" value="<%=sItemType%>">
		<input type="hidden" name="hClassCode" value ="">
		
		<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
		<input type="hidden" name="hItemTypeName" value="">
		<input type="hidden" name="hItemCode" value="">
		<input type="hidden" name="hItemName" value="">
		<input type="hidden" name="hUnitID" value="">
		<input type=hidden name="hFromDate" value="<%=sFinFrom%>">
		<input type=hidden name="hToDate" value="<%=sFinTo%>">
        <input type=Hidden name="hEligibleFor" value="<%=sEligibleFor%>">
		<input type="hidden" name="selClass" value="">
		<input type="hidden" name="hSelectedValue" value="">
		
		
		<input type="hidden" name="hField1" value="<%=sField1%>">
		<input type="hidden" name="hField2" value="<%=sField2%>">
		<input type="hidden" name="hField3" value="<%=sField3%>">
		
		<input type="hidden" name="hFieldSelected" value="<%=nFieldSelected%>">
		

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" class="PageTitle" height="20">
					<p align="center">Item Opening Stock
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
																	<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')">
																		<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
																		</a>
																	</td>
																	<td valign="center" class="SubTitle">&nbsp;&nbsp;
																	</td>
																</tr>
															</table>
															<table border="0" cellpadding="0" cellspacing="0" width=100%>
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="width: 100%; display: none">
																		<table border=0 cellpadding="0" cellspacing="0" width="100%" class=BodyTable>
																			<tr>
																				<td class="MiddlePack" width="16"></td>
																				<td class="MiddlePack" colspan="6" width="510"></td>
																			</tr>
																			<!--<tr>
																				<td class="FieldCellSub" width="6"></td>
																				<!--<td class="FieldCellSub" width="82">Unit Name</td>
																				<td class="FieldCellSub" colspan="2" width="169">
																					<select size="1" name="selUnitId" class="FormElem">
																						<%populateUnitSelected sUnitID%>
																					</select>
																				</td>
																				<td class="FieldCellSub" width="20"></td>-->
																			<!--	<td class="FieldCellSub" width="90">Item Type</td>
																				<td class="FieldCellSub" width="175">
																					<select size="1" name="selItemType" class="FormElem">
																						<%'populateItemTypeSelected sItemType%>
																					</select>
																				</td>
																			</tr>-->
																			<tr>
																				<td class="FieldCellSub" width="11"></td>
																				<td class="FieldCellSub" width="87">Category</td>
																				<td class="FieldCellSub" colspan="2" width="169">
																					<select size="1" name="selCategory" class="Formelem">
																						<option value="select">Select</option>
																						<%populateCategorySelected sCat%>
																			        </select>
																				</td>
																				<td class="FieldCellSub" width="24"></td>
																				<td class="FieldCellSub" width="95"></td>
																				<td class="FieldCellSub" width="163">
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="11"></td>
																				<td class="FieldCellSub" width="87">Eligible For</td>
																				<td class="FieldCellSub" colspan="5">
																				    <input type="Checkbox" value="P" name="ChkPur" <%if instr(1,sEligibleFor,"P") > 0 then Response.Write "checked" %>>Purchase &nbsp;
																					<input type="Checkbox" value="S" name="ChkSales" <%if instr(1,sEligibleFor,"S") > 0 then Response.Write "checked" %>>Sales &nbsp;
																					<input type="Checkbox" value="I" name="ChkInv" <%if instr(1,sEligibleFor,"I") > 0 then Response.Write "checked" %>>Inventory &nbsp;
																					<input type="Checkbox" value="M" name="ChkManu" <%if instr(1,sEligibleFor,"M") > 0 then Response.Write "checked" %>>Manufacrue&nbsp;
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="11"></td>
																				<td class="FieldCellSub">Search By</td>
																				<td class="FieldCellSub" colspan="5">
																					<input type="radio" value="I" name="chkSearch" <%if instr(1,sSearchBy,"I") > 0 then Response.Write "checked" %>>Item Code &nbsp;
																					<input type="radio" value="N" name="chkSearch" <%if instr(1,sSearchBy,"N") > 0 then Response.Write "checked" %>>Item Name &nbsp;
																					<input type="radio" value="D" name="chkSearch" <%if instr(1,sSearchBy,"D") > 0 then Response.Write "checked" %>>Drawing No.&nbsp;
																					<input type="radio" value="C" name="chkSearch" <%if instr(1,sSearchBy,"C") > 0 then Response.Write "checked" %>>Catalog No.&nbsp;
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="11"></td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" colspan="5">
																					<input type="text" name=txtSearch value="<%=sFilter%>" class="FormElem">
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCell" width="11"></td>
																				<td class="FieldCell" width="87">

																			        &nbsp;

																				</td>
																				<td class="FieldCellSub" width="169">
																			        <p align="right"><input type="button" value="Go" name="Cmdgo" class="ActionButton" onClick="Search()">
																			        </p>
																			    </td>
																				<td class="FieldCellSub" colspan="2" width="20"></td>

																				<td class="FieldCellSub" colspan="2" width="224">
																				<input type="button" value="Reset" name="Cmdreset" class="ActionButtonX">
																				</td>
																			</tr>
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
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
									<tr>
										<td align="center" width="5" class="ClearPixel"></td>
										<td valign="top">
											<div>
											<table border="0" cellspacing="1" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2" width="10" ></td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
													<%
													if trim(sField1) <> ""  then
														if instr(1,sField1,":") > 0 then
															Arr1 = Split(sField1,":")
															
															if Arr1(1) = "A" then ' if ascending order is exist, give option to descending order
																%>
																<span style="cursor:hand" onclick="Sort(1,'I','D')">Item Code</span>
																<%
															else
																%>
																<span style="cursor:hand" onclick="Sort(1,'I','A')">Item Code</span>
																<%
															end if
														end if 
													else
														%>
														<span style="cursor:hand" onclick="Sort(1,'I','A')">Item Code</span>
														<%
													end if
													%>
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
													
													<%
													if trim(sField2) <> ""  then
														if instr(1,sField2,":") > 0 then
															Arr1 = Split(sField2,":")
															
															if Arr1(1) = "A" then ' if ascending order is exist, give option to descending order
																%>
																<span style="cursor:hand" onclick="Sort(2,'D','D')">Description</span>
																<%
															else
																%>
																<span style="cursor:hand" onclick="Sort(2,'D','A')">Description</span>
																<%
															end if
														end if 
													else
														%>
														<span style="cursor:hand" onclick="Sort(2,'D','A')">Description</span>
														<%
													end if
													%>
													
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
														Store - Bin													
													</td>
													<td class="ExcelHeaderCell" align="center" colspan="4">Opening Stock</td>
													<td class="ExcelHeaderCell" align="center" colspan="2">Change Stock</td>
													<td class="ExcelHeaderCell" align="center">Total</td>
												</tr>
												<tr>
													<td class="ExcelHeaderCell" align="center" >Quantity</td>
													<td class="ExcelHeaderCell" align="center" >Value</td>

													<td class="ExcelHeaderCell" align="center" >Issue</td>
													<td class="ExcelHeaderCell" align="center" >Balance</td>
													
													<td class="ExcelHeaderCell" align="center" >Quantity</td>
													<td class="ExcelHeaderCell" align="center" >Value</td>
													
													<td class="ExcelHeaderCell" align="center" >Quantity</td>
												</tr>
												<%
													
													
												    Response.Write "<font color=red>"
												    iItemCnt = 0
												    
												   ' sSqlBasic = "SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD,isNull(LocationNumber,0),isNull(BinNumber,0),LocationName,SUM(YearIssueQuantity),SUM(YearIssueValue)  FROM VIEWWALLITEMS " &_
													'			" WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & ""
																
												    sSqlBasic = "SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD,isNull(LocationNumber,0),isNull(BinNumber,0),LocationName,SUM(YearIssueQuantity),SUM(YearIssueValue)  FROM VIEWWALLITEMS " &_
																" WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " "
																
																
												    sCondition =  " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) "
												    
												    if sCat <> "" then
														sCondition = sCondition &  " AND CATEGORYCODE = " & Pack(sCat)
													end if													
												    
														
													if sSearchBy = "I" then
													
														sFilter = "%"&sFilter&"%"
														
														sCondition = sCondition & " AND COMPANYITEMCODE LIKE " & Pack(sFilter)
														
													elseif sSearchBy = "D" then
														
														sFilter = "%"&sFilter&"%"
														
														sCondition = sCondition  & " AND DRAWINGNUMBER LIKE " & Pack(sFilter)
														
													elseif sSearchBy = "C" then
														
														sFilter = "%"&sFilter&"%"
														
														
														sCondition = sCondition  & " AND CATALOGUENO LIKE " & Pack(sFilter)
														
													elseif sSearchBy = "N" then
														
														sFilter = "%"&sFilter&"%"
														
														sCondition = sCondition  & " AND ITEMDESCRIPTION LIKE " & Pack(sFilter)
														
													end if
													
													if instr(1,sEligibleFor,"P") > 0 then
														sCondition = sCondition  & " and PurchaseEligible = '1'"
													end if
													
													if instr(1,sEligibleFor,"S") > 0 then
														sCondition = sCondition  & " and SalesEligible = '1'"
													end if 													 
													
													if instr(1,sEligibleFor,"I") > 0 then
														sCondition = sCondition  & " and InventoryEligible = '1'"
													end if 
														
													if instr(1,sEligibleFor,"M") > 0 then
														sCondition = sCondition  & " and ManufactureEligible = '1'"
													end if 
													
													
													if trim(sListOfItemCode) <> "" then
														sCondition = sCondition  & " and ITEMCODE in (" & sListOfItemCode & ") and CLASSIFICATIONCODE in (" & sListOfClassCode & ")"
													end if
													
																										
													sGroupBy = " GROUP BY COMPANYITEMCODE,ITEMDESCRIPTION,CategoryCode,ITEMCODE,CLASSIFICATIONCODE,ITEMACTIVE,ITEMONHOLD,isNull(LocationNumber,0),isNull(BinNumber,0),LocationName  "
													
													sSql = sSqlBasic & sCondition & sGroupBy
													
													if trim(sSortBy) <> "" then
														sSql = sSql & " Order By " & sSortBy
													end if 	
																																							
													'Response.Write "<P style='color:red' >" & sSql 
													'Response.Write "<P style='color:red' >" & sSortBy
													with rsStock
														.ActiveConnection=con
														.CursorLocation=3
														.CursorType=3
														.Source=sSql
														.Open
													end with
													
													nPrevItemCode = 0
													
													if not rsStock.EOF then
														'''''''''''''''''''''''''''''''''''''''''''''''''''''''
   															rsStock.PageSize = iPageSize
															If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
															rsStock.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
															iTotPage = rsStock.PageCount					'stores total no. of pages
														'''''''''''''''''''''''''''''''''''''''''''''''''''''''
														iPageCtr = 1
														do while iPageCtr <= rsStock.PageSize and NOT rsStock.EOF 
														    if rsStock(8) = 1 or rsStock(7) = "N" then
																sClass = "ExcelDisplayCellcolor1"
															else
																sClass = "ExcelDisplayCell"
															end if

															
															sBinName = ""
															sSql = "Select BinName from Inv_M_StoreBinDetails where OUDefinitionID = '" & sUnitID & "' and LocationNumber = "& rsStock(9) & " and  BinNumber = " & rsStock(10) & ""
															'Response.Write "<p> " & sSql
															rsTemp.Open sSql,con
															if not rsTemp.EOF then
																sBinName = trim(rsTemp(0))
															end if
															rsTemp.Close
															
															nRowSpan = 1
															if trim(nPrevItemCode) <> trim(rsStock(0)) then
															
																nPrevItemCode = rsStock(0)
																sDisplayData = "Y"
																
																nLocBinCtr = 1
																
																sSql = "Select ItemCode from VIEWWALLITEMS where ORGANISATIONCODE = '" & sUnitID & "'" &_
																		" and ItemCode = "& rsStock(0) & " and  ClassificationCode = " & rsStock(1) & "" &_
																		sCondition
																	
																'Response.Write "<p> " & sSql
																rsTemp.Open sSql,con
																if not rsTemp.EOF then
																	nRowSpan = rsTemp.RecordCount 
																end if
																rsTemp.Close
															end if 	
	
															

															sItemDesc = Replace(rsStock(5),"'","~~")
															sItemDesc = Replace(rsStock(5),Chr(34),"``")
															
															iOpenStok = 0
															iOpenRate = 0
															iOpenValue = 0
															iIssueStock = 0 
															iBalanceStock = 0
															
														    sQuery = "Select isNull(SUM(LotQuantityNett),0),isNull(SUM(QuantityIssued),0),isNull(SUM(Rate),0) from INV_T_LocationLot where ItemCode in ("& rsStock(0) &")"&_
														             " and SrcType = 'RO' and isNull(StorageLocationNo,0) = "& rsStock(9) &" and isNull(StorageBinNumber,0) = "& rsStock(10)
														    rsTemp.Open sQuery,con
														    if not rsTemp.EOF then
														        iOpenStok = rsTemp(0)
														        iIssueStock = rsTemp(1)
														        iOpenRate = rsTemp(2)
														        if trim(iOpenRate)<>"0" then
														            iOpenValue = CDbl(iOpenStok)*cdbl(iOpenRate)
														        else
														            iOpenValue = 0 
														        end if
														        iBalanceStock = CDbl(iOpenStok)- CDbl(iIssueStock)
														        if trim(iOpenRate)<>"0" then
														            iBalanceValue = CDbl(iBalanceStock) * CDbl(iOpenRate)
														        else
														            iBalanceValue = 0
														        end if
														    end if
														    rsTemp.Close 
														    iCnt = iCnt + 1
															%>
															<tr>
																<%if sDisplayData = "Y" then
																    iItemCnt = iItemCnt + 1
																    iCnt = 1
																	sDisplayData = "N"
																%>
																	<Input Type=Hidden name="hTotalRow<%=iItemCnt%>" Value="<%=nRowSpan%>" >
																	
																	<td class="ExcelSerial" align="center" rowspan="<%=nRowSpan%>" ><%=iItemCnt%></td>
																	<td class="<%=sClass%>" align="center" rowspan="<%=nRowSpan%>"  width="10">
																		<input type="checkbox" name="Chkbox<%=iItemCnt%>" value="<%=trim(rsStock(0))%>:<%=trim(rsStock(1))%>:<%=sUnitID%>:<%=trim(rsStock(2))%>:<%=sItemDesc%>:<%=trim(rsStock(7))%>:<%=trim(rsStock(8))%>:<%=sItemType%>:<%=rsStock(9)%>:<%=rsStock(10)%>" onclick="EnableStock('<%=iItemCnt%>')">
																	</td>
																	<td class="<%=sClass%>" align="Left"  rowspan="<%=nRowSpan%>" ><%=trim(rsStock(2))%></td>
																	<td class="<%=sClass%>" align="left"  rowspan="<%=nRowSpan%>" ><%=trim(rsStock(5))%></td>
																<%else
																	nLocBinCtr = nLocBinCtr + 1 
																end if %>
																
																
																<td class="<%=sClass%>" align="left" ><%=trim(rsStock(11)) & "-" & trim(sBinName) %></td>
																<td class="<%=sClass%>" align=right><%=trim(iOpenStok)%></td>
																<td class="<%=sClass%>" align=right><%=trim(iOpenValue)%></td>
																<td class="<%=sClass%>" align=right><%=trim(iIssueStock)%></td>
																<td class="<%=sClass%>" align=right><%=trim(iBalanceStock)%></td>
																<td class="<%=sClass%>" align=right>
																    <input type="hidden" name="hLocBinValZ<%=iItemCnt%>Z<%=iCnt%>" value="<%=rsStock(9)%>:<%=rsStock(10)%>">
																    <input type="hidden" name="hLocBinCtrZ<%=iItemCnt%>Z<%=iCnt%>" value="<%=nLocBinCtr%>">
																    <input type="hidden" name="hIssQtyZ<%=iItemCnt%>Z<%=iCnt%>Z<%=nLocBinCtr%>" value="<%=trim(iIssueStock)%>">
																    <input type="hidden" name="hOpRateZ<%=iItemCnt%>Z<%=iCnt%>Z<%=nLocBinCtr%>" value="<%=trim(iOpenRate)%>">
																	<input type="text" name="txtStkQtyZ<%=iItemCnt%>Z<%=iCnt%>Z<%=nLocBinCtr%>" value="<%=iBalanceStock%>" class="FormElemRead" size="6" onBlur="CalculateTotalStock('<%=iItemCnt%>','<%=iCnt%>','<%=nLocBinCtr%>')">
																</td>
																<td class="<%=sClass%>" align=right>
																	<input type="text" name="txtStkValueZ<%=iItemCnt%>Z<%=iCnt%>Z<%=nLocBinCtr%>" value="<%=iBalanceValue%>" class="FormElemRead" size="10">
																</td>
																<td class="<%=sClass%>" align=right>
																	<input type="text" name="txtTotQtyZ<%=iItemCnt%>Z<%=iCnt%>Z<%=nLocBinCtr%>" value="<%=cdbl(iIssueStock)+cdbl(iBalanceStock)%>" class="FormElemRead" size="6">
																</td>
															</tr>

															<%
															rsStock.MoveNext
														loop
													end if 'if not rsStock.EOF then
													rsStock.Close
												%>
											</table>
											</div>
										</td>
										<td align="center" class="ClearPixel" width="5">
										<input type="hidden" name="hTotItem" value="<%=iItemCnt%>">
										</td>
									</tr>
									<tr>
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
									<tr>
									    <td colspan="2" align=right>
													<p align="Right">
								                        <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
								                        <Input Type=Hidden name="hCtr" Value="<%=iCnt%>" >
								                        <Input Type=Hidden name="hPageSelection" Value="1" >
														<%	If iTotPage >= 2 Then
																if iCurrentPage = 1 then
														%>
														<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
														<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
														<%		else	%>
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
												<td></td>

									</tr>
									<tr>
										<td align="center" width="5" class="ClearPixel">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
													<p align="center">
													
													<Input type="button" value="Save" name="ButSave" class="ActionButton" tabindex="3" onclick="CheckSubmit()" >
													
												</td>
											</tr>
										</table>
									</td>
									<td align="center" class="ClearPixel" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<tr>
									<td align="center" class="BottomPack" colspan="3"></td>
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
