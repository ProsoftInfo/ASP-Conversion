<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	ItemCycleCountGrid.asp
	'Module Name				:	INVENTORY (Item Cycle Count)
	'Author Name				:
	'Created On					:
	'Modified By				:	Ragavendran R
	'Modified On				:	Aug 28,2013
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
<!-- #include File="../../include/CommonFunctions.asp" -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Cycle Count Grid</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemDetails"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemCycleCount.js"></SCRIPT>
</head>
<%
    Response.write "<font color=red>"
	'Declaring Variables
	Dim dcrs,dcrs1
	Dim sOrgCode,sOrgName,sTempMonYr,sMonYr,arrFin,sFinFrom,sFinTo,sFromDate,sToDate,sSql,sClass,sUserName
	Dim iCnt
	

	set dcrs=server.CreateObject("ADODB.Recordset")
	set dcrs1=server.CreateObject("ADODB.Recordset")

	sOrgCode = Session("organizationcode")
	sOrgName = session("OrgName")
	
	sTempMonYr = mid(FormatDate(date()),4,2)
	sMonYr = sTempMonYr&Year(FormatDate(date()))

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	
	sFromDate = Request("hFromDate")
	sToDate = Request("hToDate")
'	Response.write "<font color=red>"
'	Response.write "<p>sFromDate = "& sFromDate
	if trim(sFromDate)="" or IsNull(sFromDate) then
	    sFromDate = sFinFrom
	    sToDate = sFinTo
	end if

	if sOrgCode = "" then sOrgCode = "010101"
	

	''''''''''''''''''''' Paging Declaration ''''''''''''''''''''''''''''''''''''''''
    Const iPageSize=15	'How many records to show
    Dim iCurrentPage	'Current Page No.
    Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
    Dim iPageCtr		'Counter
	Dim lnPage

    iCurrentPage = Request.Form("hPageSelection")
    if iCurrentPage = "" or iCurrentPage = "0" then iCurrentPage = "1"
    iCurrentPage = CInt(iCurrentPage)

    con.CursorLocation = 3

	

%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  onload="Init()">
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
		<input type="hidden" name="hOrgCode" value="<%=sOrgCode%>">
		<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
		<input type=hidden name="hFinFromDate" value="<%=sFinFrom%>">
		<input type=hidden name="hFinToDate" value="<%=sFinTo%>">
		<input type=hidden name="hFromDate" value="<%=sFromDate%>">
		<input type=hidden name="hToDate" value="<%=sToDate%>">

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" height="20">
				     <table>
			            <tr>
			                <td class="PageTitle" >
			                    <p align="center">
			                    Cycle Counting
			                    </p>
			                </td>
			                <td class="PageTitle" >
			                    <a style="text-decoration:none;font:color:black" href="#" onclick="Help()">Help</a>
			                </td>
			            </tr>
			        </table>
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
																			    <td class="FieldCellSub" align="right">From</td>
																			    <td class="FieldCellSub">
																			        <%
																                        ' Function Call to Insert Date Picker
																                        Response.Write InsertDatePicker("ctlFromDate")
															                        %>
																			    </td>
																			    <td class="FieldCellSub">To</td>
																			    <td class="FieldCellSub">
																			        <%
																                        ' Function Call to Insert Date Picker
																                        Response.Write InsertDatePicker("ctlToDate")
															                        %>
																			    </td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" colspan="4">
																			        <p align="center"><input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Search()">
																			        <input type="button" value="Reset" name="Cmdreset" class="ActionButtonX">
																			        </p>
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
											<div style="height:400px;">
											<table id="tblItem" border="0" cellspacing="1" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
													<td class="ExcelHeaderCell" align="center">
													    Cycle Count No
													</td>
													<td class="ExcelHeaderCell" align="center">
													    Cycle Count Date
													</td>
													<td class="ExcelHeaderCell" align="center">Created By</td>
												</tr>
												<%
												
												    
												    Response.Write "<font color=#000000>"
												    
												    sSql = "Select CycleCountEntryNo,Convert(Varchar,CycleCountDate,103),CycleCountDoneBy,Convert(varchar,CycleCountDoneOn,103) from Inv_T_ItemCycleCount where Convert(datetime,FinancialYearFrom,103)=Convert(datetime,'"& sFinFrom &"',103) and Convert(datetime,FinancialYearTo,103)=Convert(datetime,'"& sFinTo &"',103) and OrganisationCode="&sOrgCode
												    sSql = sSql & " and Convert(datetime,CycleCountDate,103) between Convert(datetime,'"& sFromDate &"',103) and Convert(datetime,'"& sToDate &"',103)"
												  	'Response.Write "<P style='color:red' >" & sSql
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
														    sClass="ExcelDisplayCell"
														    
														    sUserName = split(GetUserInfo(dcrs(2)),":")(2)
														    %>
												            <tr>
													            <td class="ExcelSerial" align="center" ><%=iCnt%>
													            <input type="hidden" name="Chkbox<%=iCnt%>" value="" onclick="EnableCycleCount('<%=iCnt%>')">
													            </td>
													            <td class="<%=sClass%>" align="Left" >
													                <a href="#" onclick="ViewCycleCount('<%=trim(dcrs(0))%>','<%=trim(dcrs(1))%>','<%=sOrgCode%>')" class="ExcelDisplayLink" style="cursor:hand;"><%=trim(dcrs(0))%></a>
													            
													            </td>
													            <td class="<%=sClass%>" align="left" ><%=trim(dcrs(1))%></td>
													            <td class="<%=sClass%>" align="left" ><%=sUserName%></td>
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
    													<SELECT class="FormElem" onChange="Paginate(this.options[this.selectedIndex].value)" id="selPage">
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
													    <input type="button" name="btnSave" value="New" class="ActionButton" onclick="PopulateItem()" />
													    <!--<input type="button" name="btnView" value="View" class="ActionButton" onclick="ShowCycleCountDet()"  />-->
													</p>
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

