<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	ItemListEntryForFA.asp
	'Module Name				:	INVENTORY (Item List)
	'Author Name				:
	'Created On					:
	'Modified By				:	UmaMaheswari S
	'Modified On				:	Dec 24,2010
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
<!-- #include File="../../include/ItemDisplay.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Grid</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" id="ItemDetails" data-itms-xml-island="1"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<Script>
function Search() {
	document.formname.submit();
}

function GotoAction() {
	var form = document.formname;
	var count = Number(form.hCtr.value || 0);
	var selected = 0;
	var passValue = "INV";
	var radio;
	for (var i = 1; i <= count; i += 1) {
		radio = form.elements["Radio" + i];
		if (radio && radio.checked) {
			selected += 1;
			if (String(radio.value || "").replace(/^\s+|\s+$/g, "") !== "") {
				passValue += ":" + radio.value;
				form.action = "../../Fixedassets/TRANSACTION/ASTCreationDetail.asp?sPassVal=" + encodeURIComponent(passValue);
				form.submit();
				return true;
			}
		}
	}
	if (selected === 0) {
		alert("Select an Item");
	}
	return false;
}

function Sort(sBy) {
	document.formname.hSortBy.value = sBy;
	document.formname.submit();
}

function ChangeStatus(sPassData) {
	var form = document.formname;
	form.hchoice.value = sPassData;
	if (form.ChkStatus && form.ChkStatus.length > 1) {
		form.ChkStatus[0].checked = String(sPassData).replace(/^\s+|\s+$/g, "") === "D";
		form.ChkStatus[1].checked = String(sPassData).replace(/^\s+|\s+$/g, "") === "T";
	}
	form.submit();
}

function Paginate(pageNo) {
	document.formname.hPageSelection.value = pageNo;
	document.formname.submit();
}
</script>
</head>
<%
	'Declaring Variables
	Dim sUnitID,sItemType,sSql,iCnt,dcrs,sSortBy,sStatus,sPassValue
	set dcrs=server.CreateObject("ADODB.Recordset")
	dim sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr

	sUnitID = Request("selUnitId")
	sSortBy = Request("hSortBy")

	arrFin = split(Session("FinPeriod"),":")
	sFinFrom = "01/04/"&arrFin(0)
	sFinTo = "31/03/"&arrFin(1)

	if sUnitID = "" then sUnitID = "010101"
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
	'Response.Write sSortBy
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if sSortBy = "D" then
		'sSortBy = "ITEMDESCRIPTION,COMPANYITEMCODE "
		sSortBy = "V.ITEMDESCRIPTION "
	elseif sSortBy = "I" then
		'sSortBy = "COMPANYITEMCODE,ITEMDESCRIPTION "
		sSortBy = "V.COMPANYITEMCODE "
	else
		'sSortBy = "COMPANYITEMCODE "
		sSortBy = "V.CLASSIFICATIONCODE "
	end if
%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  >
	<form method="POST" name="formname" action="">
		<input type="hidden" name="hOrgId" value="<%=sUnitID%>">
		<input type="hidden" name="hItemType" value="<%=sItemType%>">
		<input type="hidden" name="hSortBy" value="D">
		<input type="hidden" name="hOrgName" value="">
		<input type="hidden" name="hClassSelectedCode" value="">
		<input type="hidden" name="hItemSelectedCode" value="">
		<input type="hidden" name="selOrgUnit" value="">
		<input type=hidden name="hFromDate" value="<%=sFinFrom%>">
		<input type=hidden name="hToDate" value="<%=sFinTo%>">
		<input type="hidden" name="hchoice" value="<%=sStatus%>">
		<input type="hidden" name="hPara" value="">
		<input type="hidden" name="hItemTypeName" value="">

		<input type="hidden" name="selUnit" value="">
		<input type="hidden" name="selClass" value="">
		<input type="hidden" name="hSelectedValue" value="">
		<input type="hidden" name="hCallFrom" value="ItemList">

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" class="PageTitle" height="20">
					<p align="center">Item List
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
																	<td valign="center" class="SubTitle">&nbsp;</td>
																</tr>
															</table>
															<table border="0" cellpadding="0" cellspacing="0">
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="width: 575; display: none">
																		<table cellpadding="0" cellspacing="0" width="613">
																			<tr>
																				<td class="MiddlePack" width="16"></td>
																				<td class="MiddlePack" colspan="6" width="510"></td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="6"></td>
																				<td class="FieldCellSub" width="82">Unit Name</td>
																				<td class="FieldCellSub" colspan="2" width="169">
																					<select size="1" name="selUnitId" class="FormElem">
																						<%populateUnitSelected sUnitID%>
																					</select>
																				</td>
																				<!--<td class="FieldCellSub" width="20"></td>
																				<td class="FieldCellSub" width="90">Item Type</td>
																				<td class="FieldCellSub" width="175">
																					<select size="1" name="selItemType" class="FormElem">
																						<%populateItemTypeSelected sItemType%>
																					</select>
																				</td>-->
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
													<td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
													<td class="ExcelHeaderCell" align="center" width="10" ></td>
													<td class="ExcelHeaderCell" align="center" ><span style="cursor:hand" onclick="Sort('I')">Item Code</span></td>
													<td class="ExcelHeaderCell" align="center" ><span style="cursor:hand" onclick="Sort('D')">Description</span></td>
													<td class="ExcelHeaderCell" align="center" ><span style="cursor:hand" onclick="Sort('C')">Category</span></td>
													<td class="ExcelHeaderCell" align="center" >Stock Qty</td>
													<td class="ExcelHeaderCell" align="center" >Stock Value</td>
												</tr>
												<%

													'sSql = " SELECT  DISTINCT ITEMCODE,CLASSIFICATIONCODE,COMPANYITEMCODE,ITEMDESCRIPTION,"&_
													'	   " YEARCLOSINGSTOCK,YEARCLOSINGVALUE,STORESUOM FROM VWALLITEMS WHERE YEARCLOSINGSTOCK > 0 "&_
													'	   " and ITEMTYPEID = 'STO' AND ORGANISATIONCODE = " & Pack(sUnitID) & " AND "&_
													'	   " CONVERT(DATETIME,FINANCIALYEARFROM,103) = CONVERT(DATETIME,'"& sFinFrom &"',103) AND "&_
													'	   " CONVERT(DATETIME,FINANCIALYEARTO,103) = CONVERT(DATETIME,'"& sFinTo &"',103) ORDER BY "& sSortBy

													sSql = " SELECT  DISTINCT V.ITEMCODE,V.CLASSIFICATIONCODE,V.COMPANYITEMCODE,V.ITEMDESCRIPTION,"&_
														   " V.YEARCLOSINGSTOCK,V.YEARCLOSINGVALUE,V.STORESUOM FROM VWALLITEMS V,Inv_M_ItemOrgMaster M  WHERE V.YEARCLOSINGSTOCK > 0 "&_
														   " AND V.ITEMTYPEID = 'STO' AND V.ORGANISATIONCODE = " & Pack(sUnitID) & " AND V.ITEMCODE = M.ITEMCODE "&_
														   " AND V.CLASSIFICATIONCODE = M.CLASSIFICATIONCODE and M.STOCKNONSTOCK = 'S' AND "&_
														   " CONVERT(DATETIME,V.FINANCIALYEARFROM,103) = CONVERT(DATETIME,'"& sFinFrom &"',103) AND "&_
														   " CONVERT(DATETIME,V.FINANCIALYEARTO,103) = CONVERT(DATETIME,'"& sFinTo &"',103) ORDER BY "& sSortBy


													'Response.Write sSql

													with dcrs
														.ActiveConnection=con
														.CursorLocation=3
														.CursorType=3
														.Source=sSql
														.PageSize = iPageSize
														.Open
													end with

												'''''''''''''''''''''''''''''''''''''''''''''''''''''''

													If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
													iCnt = 0
													sPassValue = ""

													if not dcrs.EOF then
														dcrs.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
														iTotPage = dcrs.PageCount					'stores total no. of pages

														''''''''''''''''''''''''''''''''''''''''''''''''''
														For iPageCtr = 1 to dcrs.PageSize
															iCnt = iCnt + 1
															sPassValue = sUnitID & ":" & dcrs(0) & ":" & dcrs(1) & ":" & dcrs(4) & ":" & dcrs(5) & ":" & dcrs(6) & ":" & dcrs(3)
														%>
															<tr>
																<td class="ExcelSerial" align="center" ><%=iCnt%></td>
																<td class="ExcelDisplayCell" align="center" width="10">
																	<input type="Radio" name="Radio<%=iCnt%>" value="<%=Trim(sPassValue)%>">
																</td>
																<td class="ExcelDisplayCell" align="Left" ><%=trim(dcrs(2))%></td>
																<td class="ExcelDisplayCell" align="left" ><%=trim(dcrs(3))%></td>

																<td class="ExcelDisplayCell" align="left" >-</td>
																<td class=ExcelDisplayCell align=right><%=trim(dcrs(4))%></td>
																<td class=ExcelDisplayCell align=right><%=trim(dcrs(5))%></td>
															</tr>
														<%
															dcrs.MoveNext
															If dcrs.EOF Then Exit For

														next

													end if 	'if not dcrs.EOF then
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
										<td align="center" width="5" class="ClearPixel">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
										<td valign="top">
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<!--<td valign="middle" class="ActionCell">
													<p align="center">
													<select size="1" name="Choice" class="FormElem">
														<option value = "SEL">Select</option>
														<%' SelFun()	%>

													</select>
													<Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
												</td>-->
												<td valign="middle" class="ActionCell">
													<p align="center">
								                        <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
								                        <Input Type=Hidden name="hCtr" Value="<%=iCnt%>" >
								                        <Input Type=Hidden name="hPageSelection" Value="1" >
								                        <Input type="button" value="Go" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >
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
