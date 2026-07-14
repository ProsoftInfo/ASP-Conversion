<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	ItemListEntry.asp
	'Module Name				:	INVENTORY (Item List)
	'Author Name				:
	'Created On					:
	'Modified By				:	Ragavendran R
	'Modified On				:	Jan 07,2011
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
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/CheckPrevFinYear.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Grid</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemDetails"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemListOld.js"></SCRIPT>

</head>
<%
	'Declaring Variables
	Dim sUnitID,sItemType,sSql,iCnt,dcrs,sSortBy,sCat,sSearchBy,sFilter,sClass,sCatName,dcrs1,sOrgName,sACTN
	Dim sEligibleFor, sFSNFlag

	set dcrs=server.CreateObject("ADODB.Recordset")
	set dcrs1=server.CreateObject("ADODB.Recordset")
	dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,sItemDesc

	sUnitID = Session("organizationcode")
	sSql = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sUnitID
	dcrs.Open sSql,con
	if not dcrs.EOF then
		sOrgName = trim(dcrs(0))
	end if
	dcrs.Close

	'sUnitID = Request("selUnitId")
	sItemType = Request("selItemType")
	sSortBy = Request("hSortBy")
	sCat = Request("selCategory")
	sSearchBy = Request("chkSearch")
	sFilter = trim(Request("txtSearch"))
	sACTN = trim(Request("ACTN"))
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
	if sSortBy = "" then sSortBy = "D"

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
	if sSortBy = "D" then
		sSortBy = "ITEMDESCRIPTION,COMPANYITEMCODE,CategoryCode "
	elseif sSortBy = "I" then
		sSortBy = "COMPANYITEMCODE,ITEMDESCRIPTION,CategoryCode "
	else
		sSortBy = "CategoryCode,ITEMDESCRIPTION,COMPANYITEMCODE "
	end if

	'Response.Write "<p>data="&Request.ServerVariables("SCRIPTNAME")

%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  >
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
		<input type="hidden" name="hOrgId" value="<%=sUnitID%>">
		<input type="hidden" name="hItemTypeCode" value="<%=sItemType%>">
		<input type="hidden" name="hClassCode" value ="">
		<input type="hidden" name="hSortBy" value="D">
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

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" class="PageTitle" height="20">
					<%If sACTN = "PA" Then%>
						<p align="center">Physical Adjustment
					<%ElseIf sACTN = "MI" Then%>
						<p align="center">Move Item
					<%ElseIf sACTN = "M" Then%>
						<p align="center">Manage Item
					<%Else%>
						<p align="center">Item List
					<%end If%>
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
																			<tr>
																				<td class="FieldCellSub" width="6"></td>
																				<!--<td class="FieldCellSub" width="82">Unit Name</td>
																				<td class="FieldCellSub" colspan="2" width="169">
																					<select size="1" name="selUnitId" class="FormElem">
																						<%populateUnitSelected sUnitID%>
																					</select>
																				</td>
																				<td class="FieldCellSub" width="20"></td>-->
																				<td class="FieldCellSub" width="90">Item Type</td>
																				<td class="FieldCellSub" width="175">
																					<select size="1" name="selItemType" class="FormElem">
																						<%populateItemTypeSelected sItemType%>
																					</select>
																				</td>
																			</tr>
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
													<td class="ExcelHeaderCell" align="center" rowspan="2"><span style="cursor:hand" onclick="Sort('I')">Item Code</span></td>
													<td class="ExcelHeaderCell" align="center" rowspan="2"><span style="cursor:hand" onclick="Sort('D')">Description</span></td>
													<td class="ExcelHeaderCell" align="center" rowspan="2"><span style="cursor:hand" onclick="Sort('C')">Category</span></td>
													<td class="ExcelHeaderCell" align="center" colspan="3">Stock</td>
												</tr>
												<tr>
													<td class="ExcelHeaderCell" align="center" >Quantity</td>
													<td class="ExcelHeaderCell" align="center" >Value</td>
													<td class="ExcelHeaderCell" align="center" >Type</td>
												</tr>
												<%
												    Response.Write "<font color=#000000>"
													if sSearchBy = "" then
														if sCat = "" then
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType)
														else
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND CATEGORYCODE = " & Pack(sCat)
														end if
													elseif sSearchBy = "I" then
														sFilter = "%"&sFilter&"%"
														if sCat = "" then
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND COMPANYITEMCODE LIKE " & Pack(sFilter)
														else
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND CATEGORYCODE = " & Pack(sCat) & " AND COMPANYITEMCODE LIKE " & Pack(sFilter)
														end if
													elseif sSearchBy = "D" then
														sFilter = "%"&sFilter&"%"
														if sCat = "" then
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND DRAWINGNUMBER LIKE " & Pack(sFilter)
														else
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND CATEGORYCODE = " & Pack(sCat) & " AND DRAWINGNUMBER LIKE " & Pack(sFilter)
														end if
													elseif sSearchBy = "C" then
														sFilter = "%"&sFilter&"%"
														if sCat = "" then
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND CATALOGUENO LIKE " & Pack(sFilter)
														else
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND CATEGORYCODE = " & Pack(sCat) & " AND CATALOGUENO LIKE " & Pack(sFilter)
														end if
													elseif sSearchBy = "N" then
														sFilter = "%"&sFilter&"%"
														if sCat = "" then
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND ITEMDESCRIPTION LIKE " & Pack(sFilter)
														else
															sSql="SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,SUM(YEARCLOSINGSTOCK),SUM(YEARCLOSINGVALUE),ITEMDESCRIPTION,CategoryCode,ITEMACTIVE,ITEMONHOLD FROM VIEWWALLITEMS WHERE ORGANISATIONCODE = " & Pack(sUnitID) & " AND ITEMTYPEID = " & Pack(sItemType) & " AND CATEGORYCODE = " & Pack(sCat) & " AND ITEMDESCRIPTION LIKE " & Pack(sFilter)
														end if
													end if

													if InStr(1,sEligibleFor,"P,S,I,M") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1' and SalesEligible ='1' and ManufactureEligible = '1' and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"P,S,I") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1' and SalesEligible ='1' and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"P,S,M") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1' and SalesEligible ='1' and ManufactureEligible = '1'"
													elseif InStr(1,sEligibleFor,"P,I,M") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1' and ManufactureEligible = '1' and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"S,I,M") > 0 then
													    sSql = sSql &" and SalesEligible ='1' and ManufactureEligible = '1' and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"P,S") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1' and SalesEligible ='1'"
													elseif InStr(1,sEligibleFor,"P,M") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1' and ManufactureEligible = '1'"
													elseif InStr(1,sEligibleFor,"P,I") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1' and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"I,M") > 0 then
													    sSql = sSql &" and ManufactureEligible = '1' and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"S,I") > 0 then
													    sSql = sSql &" and SalesEligible ='1' and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"S,M") > 0 then
													    sSql = sSql &" and SalesEligible ='1' and ManufactureEligible = '1'"
													elseif InStr(1,sEligibleFor,"P") > 0 then
													    sSql = sSql &" and PurchaseEligible = '1'"
													elseif InStr(1,sEligibleFor,"S") > 0 then
													    sSql = sSql &" and SalesEligible ='1'"
													elseif InStr(1,sEligibleFor,"I") > 0 then
													    sSql = sSql &" and InventoryEligible = '1'"
													elseif InStr(1,sEligibleFor,"M") > 0 then
													    sSql = sSql &" and ManufactureEligible = '1'"
													end if


													sSql = sSql & " GROUP BY " & sSortBy & ",ITEMCODE,CLASSIFICATIONCODE,ITEMACTIVE,ITEMONHOLD"
													'Response.Write sSql
													with dcrs
														.ActiveConnection=con
														.CursorLocation=3
														.CursorType=3
														.Source=sSql
														.Open
													end with
													if not dcrs.EOF then
													'''''''''''''''''''''''''''''''''''''''''''''''''''''''
   														dcrs.PageSize = iPageSize
														If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
														dcrs.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
														iTotPage = dcrs.PageCount					'stores total no. of pages
													'''''''''''''''''''''''''''''''''''''''''''''''''''''''
														For iPageCtr = 1 to dcrs.PageSize
															iCnt = iCnt + 1
															if dcrs(8) = 1 or dcrs(7) = "N" then
																sClass = "ExcelDisplayCellcolor1"
															else
																sClass = "ExcelDisplayCell"
															end if

															with dcrs1
																.cursorLocation = 3
																.Cursortype= 3
																.activeconnection = con
																.source = "Select CategoryName from INV_M_CLASSIFICATIONCATEGORY where CategoryCode = '"& dcrs(6) &"'"
																.Open
															end with
															if not dcrs1.EOF then
																sCatName = dcrs1(0)
															end if
															dcrs1.close

															sItemDesc = Replace(dcrs(5),"'","~~")
															sItemDesc = Replace(dcrs(5),Chr(34),"``")

															with dcrs1
																.cursorLocation = 3
																.Cursortype= 3
																.activeconnection = con
																.source = "Select FSNCategory = CASE FSNCategory When 'F' Then 'Fast Moving' When 'S' Then 'Slow Moving' When 'N' Then 'Non-moving' Else '-' End from Inv_M_ItemOrgInventory where OrganisationCode = '"& sUnitID &"' and ItemCode = " & trim(dcrs(0)) & " and ClassificationCode = " & trim(dcrs(1))
																.Open
															end with
															if not dcrs1.EOF then
																sFSNFlag = dcrs1(0)
															Else
																sFSNFlag = "-"
															end if
															dcrs1.close
												%>
												<tr>
													<td class="ExcelSerial" align="center" ><%=iCnt%></td>
													<td class="<%=sClass%>" align="center" width="10">
													<input type="checkbox" name="Chkbox<%=iCnt%>" value="<%=trim(dcrs(0))%>:<%=trim(dcrs(1))%>:<%=sUnitID%>:<%=trim(dcrs(2))%>:<%=sItemDesc%>:<%=trim(dcrs(7))%>:<%=trim(dcrs(8))%>:<%=sItemType%>">
													</td>
													<td class="<%=sClass%>" align="Left" ><%=trim(dcrs(2))%></td>
													<td class="<%=sClass%>" align="left" ><%=trim(dcrs(5))%></td>
													<td class="<%=sClass%>" align="left" ><%=sCatName%></td>
													<td class="<%=sClass%>" align=right><%=trim(dcrs(3))%></td>
													<td class="<%=sClass%>" align=right><%=trim(dcrs(4))%></td>
													<td class="<%=sClass%>" align=right><%=sFSNFlag%></td>
												</tr>

												<%
														dcrs.MoveNext
														If dcrs.EOF Then Exit For
														next
													end if
													dcrs.Close
												%>
											</table>
											</div>
										</td>
										<td align="center" class="ClearPixel" width="5"></td>
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
    													<SELECT class="FormElem" onChange="Paginate(this.options[this.selectedIndex].value)" id=select1 name=select1>
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
													<%If sACTN = "PA" Then%>
														<select size="1" name="Choice" class="FormElem">
															<option value ="PAD" selected>Physical Adjusment</option>
														</select>
														<Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
													<%ElseIf sACTN = "MI" Then%>
														<select size="1" name="Choice" class="FormElem">
															<option value ="STR" selected>Move Item</option> <!--Stock Transfer-->
														</select>
														<Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
													<%ElseIf sACTN = "M" Then%>
														<select size="1" name="Choice" class="FormElem">
															<option value ="HOL" selected>Hold / UnHold</option>
															<option value ="ACT">Active / InActive</option>
															<option value ="STM">Stock Management</option>
															<option value ="MRG">Merge</option>
															<option value ="ABS">Arrange Bin Stock</option>
															<option value ="AWS">Attributes</option>
														</select>
														<Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
													<%ElseIf sACTN = "A" Then%>
														<select size="1" name="Choice" class="FormElem">
															<option value ="CAP" selected>Capitalise</option>
															<option value ="NCAP">Non-Capitalise</option>
														</select>
												    <%Elseif sACTN = "SC" then %>
												        <select size="1" name="Choice" class="FormElem">
															<option value ="USC" selected>Update Stock Closing</option>
														</select>
														<Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
												    <%Elseif sACTN = "SO" then %>
												        <select size="1" name="Choice" class="FormElem">
															<option value ="VSO" selected>View Stock Opening</option>
														</select>
														<Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
													<%Else%>
													<select size="1" name="Choice" class="FormElem">
														<option value = "SEL">Select</option>
														<option value ="DEL">Delete Item</option>
														<!--option value ="ADD">Add new Item</option-->
														<option value ="VEW" selected>View Item Details</option>
														<option value ="EDT">Edit Item Details</option>
														<option value ="REC">View Receipt History</option>
														<option value ="SAL">View Sales History</option>
														<option value ="ISS">View Issue History</option>
														<option value ="CON">View Consumption History</option>
													</select>
													<Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
													<%End If%>
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

