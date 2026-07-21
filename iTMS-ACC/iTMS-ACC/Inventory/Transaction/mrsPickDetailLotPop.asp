<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsPickDetailLotPop.asp
	'Module Name				:	Inventory(Transaction)
	'Author Name				:	Ragavendran R
	'Created On					:	April 12,2011
	'Modified By				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsPickDetailPoP.asp
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/GetSerialDetail.asp" -->

<%
dim RootNode,HeaderNode,objfs
dim oDOM

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objfs = Server.CreateObject("Scripting.FileSystemObject")

' Declaration of variables
Dim dcrs,dcrs1,dcrs2,iCtr
iCtr = 0
'Declaration of Objects

Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

dim iItem,iClass,sItemName,iLot,iLineNo,iPickNo,iPackNo,iStQty,iPageSize
Dim iCurrentPage,iTotPage,iRowCount,lnPage,iPageCtr
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,iQty,iInvRecNo
dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore
dim arrUoM,sUoMDesc,sUoMCode,sSubCon,iEntNo,sOptName,sAttrList,sQuery,sAttID
Dim sFilter,sSearchSeparator,sSearchBy,arrTempFilter,sFrom,sTo
Response.Write "<font color=#000000>"
'Response.Write Request.QueryString
arrTemp = split(trim(Request.QueryString("sTemp")),":")
iLineNo = arrTemp(1)
iLot = arrTemp(2)
sStoreCode = arrTemp(3)
sBinCode = arrTemp(4)
iStQty = arrTemp(5)
iClass = arrTemp(6)
iItem = arrTemp(7)
sOrgID = arrTemp(8)
iEntNo = arrTemp(9)
sOptName= arrTemp(11)
sAttID = arrTemp(12)
sAttrList = arrTemp(13)

'Response.Write "<p><font color=red>AttValue="&sAttrList

with dcrs2
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT GROUPNAME,SHORTDESCRIPTION,ORGUNITSHORTDESCRIPTION,ITEMDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs2.ActiveConnection = nothing

if not dcrs2.EOF then
	'sItemName = trim(dcrs2(0)) & " -- " & trim(dcrs2(1))
	sItemName = dcrs2(3)
	sOrgName = trim(dcrs2(2))
end if
dcrs2.Close

sItemName = sItemName & sOptName

arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoMDesc = arrUoM(1)

sSubCon = "NO"

Dim iAttrbList,sAttrbName,sTempAttrb,iOptVal,i

	'iPageSize = 5
	iPageSize = 20

	sSearchBy = Request.QueryString("SearchBy")
	sFilter=Request.QueryString("SearchFor")
	sSearchSeparator=Request.QueryString("SearchType")
	iRowCount = Request.QueryString("hRowCount")

	arrTempFilter = Split(sFilter,",")
	'Response.Write "<p>sFilter="&UBound(arrTempFilter)
	If UBound(arrTempFilter) = 0 Then
		sFrom = sFilter
	End IF
	if UBound(arrTempFilter)>0 then
		sFrom = arrTempFilter(0)
		sTo = arrTempFilter(1)
	end if
	'Response.Write "<p>sFrom="&sFrom
	if sSearchSeparator="E" then
		sFrom = Pack(sFilter)
		sFrom = replace(sFrom,",","','")
	end if
	'Response.Write "<p><font color=red>sFrom="&sFrom

	'sFrom = "214"
	iCurrentPage = cdbl(Request.QueryString("hCurrentPage"))
	iRowCount = CInt(Request.QueryString("hRowCount"))

	Select Case Request("hSubmit")
    	Case "Previous"								'if prev button clicked
    		iCurrentPage = Cint(iCurrentPage) - 1	'decrease current page
    	Case "Next"									'if next button clicked
    		iCurrentPage = Cint(iCurrentPage) + 1	'increase page count
    	Case "First"
    		iCurrentPage = 1
    	Case "Last"
    		iCurrentPage = iLastPage
    End Select
	iCurrentPage = cint(Request.QueryString("hSubmit"))

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/IssuePickLot.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</head>
<script type="application/xml" data-itms-xml-island="1" id="PackFromData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="PackToData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="PackNumberData"><Root/></script>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>');CheckSerialNo()">
<form method="POST" name="formname" action="">
<input type=hidden name="hSUBC" value="<%=sSubCon%>">
<input type=hidden name="hRowSelect" value="0">
<input type=hidden name="hTemp" value="<%=Request.QueryString("sTemp")%>">
<input type=hidden name="hXMLSaveMod" value="SerNo">

<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Serial Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item</td>
                                            <td class="FieldCellSub" colspan="4">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Lot Number</td>
                                            <td class="FieldCellSub" >
												<span class="DataOnly"><%=iLot%>&nbsp;</span>
											</td>
											<td class="FieldCellSub">Lot Qty</td>
                                            <td class="FieldCellSub" colspan="2">
												<span class="DataOnly"><%=iStQty%>&nbsp;</span>&nbsp;
												<span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>
											</td>
                                        </tr>
                                        <!--<tr>
                                            <td class="FieldCell">Lot Qty</td>
                                            <td class="FieldCellSub" colspan="2">
												<span class="DataOnly"><%=iStQty%>&nbsp;</span>&nbsp;
												<span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>
											</td>-->
											<!--<td class="FieldCell" align=right>No.of Pack</td>
											<td><span class="DataOnly" id="NoofPack"></span>
											</td>-->
                                        <!--/tr-->
                                        <tr>
											<!--<td class="FieldCell">Search By</td>
											<td class="FieldCellSub">
												<Select name="selSearchBy" class=FormElem>
													<option value="PN">Pack Number</option>
													<option value="PQ">Pack Quantity</option>
													<option value="PT">Pack Type</option>
													<option value="PD">Pack Date</option>
												</Select>
											</td>-->
											<td class="FieldCell">Selected Quantity[Packs]</td>
											<td class="FieldCellSub">
												<span id="spaNoofPackSelected" class="DataOnly">0.000[0]</span>
											</td>
                                        </tr>
                                        <!--<tr>
											<td class="FieldCell">Search For</td>
											<td class="FieldCellSub">
												<input type=text name=txtSearchFor value="<%'=sFilter%>" size=20 class="FormElem">
											</td>
											<td class="FieldCellSub">Use comma(,) to seperate for</td>
											<td class="FieldCellSub">
												<Select name="selSearchType" class="FormElem">
													<option value="R">Range</option>
													<option value="E">Exact</option>
												</Select>&nbsp;&nbsp;
												<input type=button name=btnGO value=GO class="ActionButtonX" onClick="NextSelection('1')">
											</td>
                                        </tr>-->
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td>
									<div class="frmBody" id="frm2" style="width: 100%; height:200px;">
									<!--<div class="frmBody" id="frm2" style="width: 100%; height:250px;">-->
										<table id="tblPackNumber"  border="0" cellspacing="1" class="ExcelTable" width="100%" >
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="10">
													<input type="CheckBox" name="chkAll" onClick="CheckAll()">
												</td>
												<td class="ExcelHeaderCell" align="center">Packing No</td>
												<td class="ExcelHeaderCell" align="center">Attributes</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
											</tr>
										<%
										if iLot = "N/A" or iLot = "" then
											sQuery = "SELECT DISTINCT SERIALNUMBER,ISNULL(AVAILABLENETSTOCK,0),ISNULL(PACKINGNUMBER,SERIALNUMBER),ISNULL(ATTRIBUTELIST,0),InventoryReceiptNo FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND ( LOTNUMBER IS NULL OR LotNumber='' or LOTNUMBER = '0' or LOTNUMBER = 'N/A') AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND AVAILABLENETSTOCK > 0 AND SERIALNUMBER IS NOT NULL "
										else
											sQuery = "SELECT DISTINCT SERIALNUMBER,ISNULL(AVAILABLENETSTOCK,0),ISNULL(PACKINGNUMBER,SERIALNUMBER),ISNULL(ATTRIBUTELIST,0),InventoryReceiptNo FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND LOTNUMBER  = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND AVAILABLENETSTOCK > 0 AND SERIALNUMBER IS NOT NULL "
										end if

									'	if trim(sAttrList)<>"0" and trim(sAttrList)<>"" then
									'		sQuery = sQuery & " and ATTRIBUTELIST='"& sAttrList &"'"
									'	end if ' if trim(sAttrList)<>"" then
									    if Trim(sAttID)<>"" and (not isNull(sAttID)) then
									        sQuery = sQuery & " and AttributeList like '%" & sAttID &"%'"
									    end if

										'Response.Write "<font color=#000000>"& sQuery

										if sFilter<>"" then
											if trim(sSearchBy)="PN" then
											'	if trim(sFrom)<>"" and trim(sTo)<>"" and sSearchSeparator ="R" then
											'		sQuery = sQuery & " and ( cast(SERIALNUMBER as integer) >= "& sFrom &" and cast(SERIALNUMBER as integer) <="& sTo &")"
											'	else
											'		sQuery = sQuery & " and SERIALNUMBER in ("& lcase(sFrom)&")"
											'	end if 'if trim(sSearchType)="R" then

												if trim(sFrom)<>"" and trim(sTo)<>"" and sSearchSeparator ="R" then
													sQuery = sQuery & " and ( cast(PACKINGNUMBER as integer) >= "& sFrom &" and cast(PACKINGNUMBER as integer) <="& sTo &")"
												else
													sQuery = sQuery & " and PACKINGNUMBER in ("& lcase(sFrom)&")"
												end if 'if trim(sSearchType)="R" then
											elseif trim(sSearchBy)="PQ" then
												if trim(sFrom)<>"" and trim(sTo)<>"" and sSearchSeparator ="R" then
													sQuery = sQuery & " and ( AVAILABLENETSTOCK >= "& sFrom &" and AVAILABLENETSTOCK <="& sTo &")"
												else
													sQuery = sQuery & " and AVAILABLENETSTOCK in ("& sFrom &")"
												end if 'if trim(sSearchType)="R" then
											elseif trim(sSearchBy)="PT" then
											'	if trim(sFrom)<>"" and trim(sTo)<>"" and sSearchSeparator ="R"  then
											'		sQuery = sQuery & " and ( PACKINGCODE >= "& sFrom &" and PACKINGCODE <="& sTo &")"
											'	else
											'		sQuery = sQuery & " and PACKINGCODE in ("& sFrom &")"
											'	end if 'if trim(sSearchType)="R" then
											elseif trim(sSearchBy)="PD" then
												sQuery = sQuery & " and Convert(varchar,DateofReceipt,103) in ("& sFilter &")"
											end if
										end if' if  sFilter<>"" then

											sQuery = sQuery & " ORDER BY 1"

										'Response.Write sQuery
										'Response.Write "<p><font color=red>"&sQuery
										%><!--<textarea><%'=sQuery%></textarea>--><%

											with dcrs2
												.CursorLocation = 3
												.CursorType = 3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with

											set dcrs2.ActiveConnection = nothing

											if not dcrs2.EOF then

												dcrs2.PageSize = iPageSize
												if iCurrentPage =0 or iCurrentPage="" then iCurrentPage = 1
												dcrs2.AbsolutePage = iCurrentPage
												iTotPage = dcrs2.PageCount

												For iPageCtr = 1 to dcrs2.PageSize
													iPackNo = dcrs2(2)
													iAttrbList = dcrs2(3)
													'Response.Write "<font color=#000000>"& iAttrbList
													sAttrbName=""
													If trim(iAttrbList) <> "0" and trim(iAttrbList)<>"" then
													sTempAttrb = split(iAttrbList,",")
													sAttrbName	= ""
														For i = LBOUND(sTempAttrb) to UBOUND(sTempAttrb)
															iOptVal = sTempAttrb(i)
															iOptVal = split(iOptVal,":")(0)
															if iOptVal<>"0" then
															    with dcrs1
																    .CursorLocation = 3
																    .CursorType = 3
																    .Source = "SELECT OPTIONNAME FROM INV_M_ITEMTYPEOPTIONS WHERE OPTIONVALUE = "& iOptVal &" "
																   ' Response.Write dcrs1.Source &"<BR><BR>"
																    .ActiveConnection = con
																    .Open
															    end with
    															Do while not dcrs1.EOF
																    sAttrbName	= sAttrbName &"/"& dcrs1(0)
																    dcrs1.MoveNext
															    loop
															    dcrs1.Close
															end if 'if iOptVal<>"0" then
														Next
													End If

													sAttrbName = Mid(sAttrbName,2)
													'Response.Write "sAttrbName="&sAttrbName&"<BR><BR>"


													if not cdbl(trim(dcrs2(1))) = 0 then
														iCtr = iCtr + 1
														iRowCount = iRowCount + 1

										%>

											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelSerial" align="center" >
												<!--<p align="center"><input type=checkbox name="chkSer<%=iCtr%>" onClick="CheckSer()" onchange="CheckSer()">-->
												<p align="center"><input type=checkbox name="chkSer<%=iCtr%>" onClick="CheckSer()">
												</td>

												<td class="ExcelDisplayCell" >
													<input type="text" name="txtSerial<%=iCtr%>"  value="<%=FindPackNumber(trim(dcrs2(0)))%>" maxlength=10 class="FormElemDisp" READONLY>
													<input type="hidden" name="hSerial<%=iCtr%>" value="<%=trim(dcrs2(0))%>">
													<input type="hidden" name="hInvNoZ<%=iCtr%>" value="<%=trim(dcrs2(4))%>">
												</td>
												<td class="ExcelDisplayCell" >
													<input type="text" name="txtAttrb<%=iCtr%>" size="25" maxlength=20  value="<%=sAttrbName%>" class="FormElemDisp" READONLY style="text-align:left">
												</td>
												<td class="ExcelDisplayCell">
													<input type="text" name="txtStQty<%=iCtr%>" size="16" maxlength=10 value="<%=cdbl(trim(dcrs2(1)))%>" class="FormElemDisp" READONLY style="text-align:right">
												</td>
												<td class="ExcelInputCell"  align=center>
												<%	if lcase(sSubCon) = "no" then %>
													<input type="text" name="txtQty<%=iCtr%>" size="16" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align:right">
												<%	else %>
													<input type="checkbox" name="chkQty<%=iCtr%>" value="<%=trim(dcrs2(0))%>" class="FormElem">
												<%	end if %>
												</td>
											</tr>
										<%
													end if
												dcrs2.MoveNext
													if dcrs2.EOF then exit for
												Next
											end if
											dcrs2.Close
										%>

										<Input type = "hidden" name="hRowcount" value ="<%=iRowcount%>" >
										<Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
										<Input Type=Hidden name="hPageSelection" Value="" >

										</table>
										<input type=hidden name="hiCtr" value="<%=iCtr%>">
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align=center colspan=3>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td valign="middle" class="FieldCell"><p align="center">
							            <%	If iTotPage >= 2 Then
												if iCurrentPage = 1 then
							            %>
							            <input type="button" value=" |< " name="Submit" class="ActionButtonX">
							            <input type="button" value=" << " name="Submit" class="ActionButtonX">
										<%		else	%>
							            <input type="button" value=" |< " name="Submit" class="ActionButtonX" onclick="NextSelection('1')">
							            <input type="button" value=" << " name="Submit" class="ActionButtonX" onclick="NextSelection('<%=iCurrentPage - 1%>')">
    									<%		end if	%>
    									<SELECT class="FormElem" onChange="NextSelection(this(this.selectedIndex).value)" id=select1 name=select1>
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
							            <input type="button" value=" >> " name="Submit" class="ActionButtonX">
							            <input type="button" value=" >| " name="Submit" class="ActionButtonX">

    									<%		else	%>
							            <input type="button" value=" >> " name="Submit" class="ActionButtonX" onclick="NextSelection('<%=iCurrentPage + 1%>')">
							            <input type="button" value=" >| " name="Submit" class="ActionButtonX" onclick="NextSelection('<%=iTotPage%>')">
    									<%		end if
											End If
										%>
									</td>
									</tr>
								</table>
								</td>
							</tr>

							<tr>
								<td colspan="3">
									<Table border="0" class="PopupTable" cellspacing = "0" width="100%">
										<tr>
											<td class="TopPack">
												<Table class="BodyTable" width="100%" cellspacing="1">

													<tr class="ExcelHeaderCell">
														<td align="center" colspan="4">
															<input type=button name=btnAddToList value="Add To List" class="ActionButtonX" onClick="btnAddToList_Click()">
															<input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
														</td>
													</tr>
													<tr>
														<td class="ExcelHeaderCell" >Search By
															<Select name="selSearchBy" class=FormElem>
																<option value="PN">Pack Number</option>
																<option value="PQ">Pack Quantity</option>
																<option value="PT">Pack Type</option>
																<option value="PD">Pack Date</option>
															</Select>
														</td>
														<td class="ExcelHeaderCell">Use comma(,) to seperate for
															<Select name="selSearchType" class="FormElem">
																<option value="R">Range</option>
																<option value="E">Exact</option>
															</Select>&nbsp;&nbsp;
														</td>
													</tr>
													<tr>
														<td class="ExcelHeaderCell" colspan="4">Search For
															<input type=text name=txtSearchFor value="<%=sFilter%>" size=20 class="FormElem">&nbsp;&nbsp;
															<input type=button name=btnGO value="Search" class="ActionButton" onClick="NextSelection('1')">
														</td>
													</tr>
													<tr>
														<td class="ExcelHeaderCell" colspan="4">Selected Entries
														</td>
													</tr>
													<tr>
														<td colspan="3">
															<div class="frmBody" id="frm2" style="height:100px;">
															<Table class="ExcelTable" cellspacing="1" width="100%" Id="tblSerDet">
																<tr>
																	<td class="ExcelHeaderCell" align="center"></td>
																	<td class="ExcelHeaderCell" align="center">Packing Number</td>
																	<td class="ExcelHeaderCell" align="center">Stock Qty</td>
																	<td class="ExcelHeaderCell" align="center">Qty</td>
																</tr>
															</Table>
															</div>
														</td>
													</tr>

												</Table>
											</td>
										</tr>
									</Table>
								</td>
							</tR>

							<!--<tr>
								<td align=center colspan=3>
									<div style="width:100%;height:160px">
										<Table id="TableSelectedPacks" width=100% class="ExcelTable" cellspacing=1>
											<tr>
												<td class="ExcelHeaderCell" align=center colspan=2>Selected Packing Numbers</td>
											</tr>
											<tr>
												<td align=center>
													<Select name="selFrombox" class=FormElem size=10 multiple>

													</Select>
												</td>
												<td align=center>
													<Select name="selTobox" class=FormElem size=10 multiple>

													</Select>
												</td>
											</tr>
											<tr>
												<td align=center>
													<input type="button" name="add"  value ="Add >>" class="AddButtonX" onClick="addclick('selTobox','selFrombox','remove');btnAdd_Click()" >
												</td>
												<td align=center>
													<input type="button" name="remove" value="<< Remove" class="AddButtonX" onClick="removeclick('selTobox','selFrombox','remove');btnRemove_Click()" >
												</td>
											</tr>
										</Table>
									</div>
								</td>
							</tr>-->
							<!--<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton">
                                                    <input type="button" value="Close" name="btnClose" class="ActionButton" onClick="btnClose_Click()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>-->
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
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
	Function FindPackNumber(iSerialNumber)
	    Dim dcrs,iPackNo,sQuery
	    Set dcrs = Server.CreateObject("ADODB.Recordset")

	    sQuery= "Select IsNull(PackingNumber,SerialNumber) from Inv_T_LocationLot where SerialNumber = "& iSerialNumber
	    dcrs.Open sQuery,con
	    if not dcrs.eof then
	        iPackNo = dcrs(0)
	    else
	        iPackNo = iSerialNumber
	    end if
	    dcrs.close
	    FindPackNumber =iPackNo
	End Function
%>
