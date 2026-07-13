<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	TEMPORARYITEMS.asp
	'Module Name				:	INVENTORY (Item List)
	'Author Name				:	KalaiSelvi R
	'Created On					:	04 Sep 2011
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/getCurrentDate.asp" -->

<%
	'Declaring Variables
	Dim sUnitID,sOrgName,sClass,sSql,sFilterBy
	Dim sCompanyItemCode,sItemDescription,sStoresUOM,sItemType
	
	Dim iCnt

	Dim dcrs,rsTemp
	
	set dcrs	= server.CreateObject("ADODB.Recordset")
	set rsTemp	= server.CreateObject("ADODB.Recordset")
		
	dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr,sTempItemDesc
	
	sFilterBy = Request.Form("hFilterBy")
	sItemType = Request.Form("selItemType")

	sUnitID = Session("organizationcode")
	sSql = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sUnitID
	dcrs.Open sSql,con
	if not dcrs.EOF then
		sOrgName = trim(dcrs(0))
	end if
	dcrs.Close

	
	sTempMonYr = mid(getCurrentDate(),4,2)
	sMonYr = sTempMonYr&Year(getCurrentDate())

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	
	if Trim(sFilterBy)="" then sFilterBy = "ALL"


	''''''''''''''''''''' Paging Declaration ''''''''''''''''''''''''''''''''''''''''
    Const iPageSize=15	'How many records to show
    Dim iCurrentPage	'Current Page No.
    Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
    Dim iPageCtr		'Counter
	Dim lnPage

    iCurrentPage = CInt(Request.Form("hPageSelection"))
    
    

    con.CursorLocation = 3
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Grid</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">

<script type="application/xml" data-itms-xml-island="1" id="ItemDetails"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemSelectData"><Root/></script>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/TempItem.js"></SCRIPT>
<script language="javascript" src="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/temporaryItems.js"></SCRIPT>

</head>

<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  >
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
		<input type="hidden" name="hOrgId" value="<%=sUnitID%>">
		<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
		
		<input type="hidden" name="hTempItemCode" value="">
		<input type="hidden" name="hItemCode" value="">
		<input type="hidden" name="hClassCode" value="">
		
		<input type="hidden" name="hItemType" value="">
		
		<input type="hidden" name="hFilterBy" value="<%=sFilterBy%>">
		
		
		
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" class="PageTitle" height="20">
					Temporary Items
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
																		<input type="checkbox" class="Formelem" name="ChkALL"		value="ALL"		OnClick="ChangeStatus(this)" <%if trim(sFilterBy)="ALL" then response.write "checked"%> >All&nbsp;&nbsp;
																		<input type="checkbox" class="Formelem" name="ChkMAPPED"	value="MAP"		OnClick="ChangeStatus(this)" <%if trim(sFilterBy)="MAP" then response.write "checked"%> >Mapped&nbsp;&nbsp;
																		<input type="checkbox" class="Formelem" name="ChkNOTMAPPED" value="NOTMAP"	OnClick="ChangeStatus(this)" <%if trim(sFilterBy)="NOTMAP" then response.write "checked"%> >Not Mapped
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
																				
																				<td class="FieldCellSub" width="90">Item Type</td>
																				<td class="FieldCellSub" width="175">
																					<select size="1" name="selItemType" class="FormElem">
																						<%'populateItemTypeSelected sItemType%>
																					</select>
																				</td>
																			</tr>-->
																			
																			
																			
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
													<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
													<td class="ExcelHeaderCell" align="center" width="10" ></td>
													<td class="ExcelHeaderCell" align="center">Temporary Item</td>
													<td class="ExcelHeaderCell" align="center">Item Code</td>
													<td class="ExcelHeaderCell" align="center">Item Description</td>
													<td class="ExcelHeaderCell" align="center">Store UOM</td>
												</tr>
												<%
													sSql = "Select M.TempItemCode,M.ItemDescription from MS_TEMPORARYITEMMASTER M where 1 = 1 "
													
													if trim(sItemType) <> ""  then
														sSql = sSql & " and M.ItemTypeId in('" & sItemType & "')"
													end if 
													
													if sFilterBy = "" then
													elseif sFilterBy = "MAP" then
														sSql = sSql & " and M.FinalStatus = 'Y'"
													elseif sFilterBy = "NOTMAP" then
														sSql = sSql & " and M.FinalStatus = 'N'"
													end if 
													sSql = sSql & "Order By M.ItemDescription"
													
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
													'For iCnt = 1 to dcrs.PageSize 
														do while not dcrs.EOF  and iCnt < dcrs.PageSize
															iCnt = iCnt + 1
															
															'sClass = "ExcelDisplayCellcolor1"
															
															sClass = "ExcelDisplayCell"
															
															sTempItemDesc = Replace(dcrs(1),"'","~~")
															sTempItemDesc = Replace(dcrs(1),Chr(34),"``")
															
															
															sCompanyItemCode	= "-"
															sItemDescription	= "-"
															sStoresUOM			= "-"
																
															sSql = "Select V.CompanyItemCode,V.ItemDescription,V.StoresUOM " & _
																	" from Ms_TempFinalItemDetail D, VIEWWALLITEMS V where V.OrganisationCode= D.OrganisationCode " & _
																	" and V.ItemCode = D.ItemCode and V.ClassificationCode = D.ClassificationCode " &_
																	" and D.TempItemCode = " & dcrs(0) & ""
															
															with rsTemp
																.ActiveConnection=con
																.CursorLocation=3
																.CursorType=3
																.Source=sSql
																.Open
															end with
															
															set rsTemp.ActiveConnection = nothing
															
															if not rsTemp.EOF then
																sCompanyItemCode	= rsTemp(0)
																sItemDescription	= rsTemp(1)
																sStoresUOM			= rsTemp(2)
															end if 
															rsTemp.Close 

															%>
															<tr>
																<td class="ExcelSerial" align="center" ><%=iCnt%></td>
																<td class="<%=sClass%>" align="center" width="10">
																	<input type="checkbox" name="Chkbox<%=iCnt%>" value="<%=trim(dcrs(0))%>">
																</td>
																<td class="<%=sClass%>" align="Left" ><%=sTempItemDesc%></td>
																<td class="<%=sClass%>" align="Left" ><%=sCompanyItemCode%></td>
																<td class="<%=sClass%>" align="Left" ><%=sItemDescription%></td>
																<td class="<%=sClass%>" align="Left" ><%=sStoresUOM%></td>
															</tr>

															<%
															dcrs.MoveNext
														'	if dcrs.EOF then exit for
														'Next
														loop
													end if 'if not dcrs.EOF then
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
													<Input type="button" value="Delete Temp.Item"	name="ButDelete"	class="ActionButtonX" onclick="DeleteItem()" >
													<Input type="button" value="Map to Actual Item" name="ButMap"		class="ActionButtonX" onclick="MapItem()" >
													<Input type="button" value="Create New Item"	name="ButCreate"	class="ActionButtonX" onclick="CreateNewItem()" >
													<Input type="button" value="New Temp.Item"		name="ButNewItem"	class="ActionButtonX" onclick="NewTempItem()" >
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

