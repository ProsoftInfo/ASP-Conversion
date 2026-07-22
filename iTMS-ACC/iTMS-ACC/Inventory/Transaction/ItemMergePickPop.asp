<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItemMergePickPop.asp
	'Module Name				:	Inventory
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 03,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
' Declaration of variables
Dim dcrs,dcrs1,dcrs2,dcrs3,iCtr,bexists
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
Set dcrs3 = Server.CreateObject("ADODB.RecordSet")

dim oDom,Root,PageNode,HeaderNode,PGNode,objfs,newElem

dim sql,sItemTypeName,sUnitName,sUsageName,iItem,iClass,iPickEntryNoCheck
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,bFlag,iQty,bChkFlag,sLot,iLotQtyReserved
dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore,iTotLotQty,iTempCtr
dim arrUoM,sUoMDesc,sUoMCode,sItemName,sUsage,sInspDet,sIssueCode,sTempLot
dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sType,bSerialFlag,iSer,sSrcType,nChkFlag
Dim sRecNumFlag,sAttList,sQuery,sAttID,sTempArrAttribute
dim iEntNo,sOptName,iLocNo,iBinNo
sInspDet = "-"
iTotLotQty = 0
bChkFlag = true

if len(Month(date())) = 1 then
	sTempMonYr = "0"&Month(date())
else
	sTempMonYr = Month(date())
end if
sMonYr = sTempMonYr&Year(date())

'arrFin = split(GetFinancialYear(sMonYr),":")
'sFinFrom = arrFin(0)
'sFinTo = arrFin(1)
arrFin = split(session("Finperiod"),":")
sFinFrom = "01/04/"&arrFin(0)
sFinTo = "31/03/"&arrFin(1)


set oDom = server.CreateObject("Microsoft.xmlDom")
Set objfs = CreateObject("Scripting.FileSystemObject")

arrTemp = split(trim(Request.QueryString("sTemp")),":")
sOrgID = Session("organizationcode")
Response.Write "<font color=#000000>"
'Response.Write Request.QueryString("sTemp")

iItem	= arrTemp(0)
iClass	= arrTemp(1)
sItemName = arrTemp(2)
iLocNo = arrTemp(3)
iBinNo = arrTemp(4)
sOptName = arrTemp(5)
sAttID	= arrTemp(6)
sAttList = arrTemp(7)
iEntNo = arrTemp(8)

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
set oDom = nothing

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

Set Root = oDOM.createElement("Pick")
Root.setAttribute "TOT",""
oDOM.appendChild Root

iTotLotQty = 0
with dcrs1
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0),OrganisationCode FROM VWItemStockStatus WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " and LOCATIONNUMBER = "& iLocNo &" and (BINNUMBER is Null or BINNUMBER = "& iBinNo &") AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
	.ActiveConnection = con
	.Open
end with
'Response.Write  dcrs1.Source
set dcrs1.ActiveConnection = nothing

Do While Not dcrs1.EOF
	iTotLotQty = 0
	sStoreCode = trim(dcrs1(0))
	sBinCode = trim(dcrs1(1))
	sOrgID = 	trim(dcrs1(2))
	with dcrs2
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID= '"& sOrgID &"' "
		.ActiveConnection = con
		.Open
	end with
'	Response.Write dcrs2.Source
	if not dcrs2.EOF then
		sOrgName = dcrs2(0)
	end if
	dcrs2.Close
	with dcrs2
		.CursorLocation = 3
		.CursorType = 3
		if Trim(sUsage) = "JWK" then
			.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,ISNULL(IL.LOTNUMBER,0) FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sStoreCode & " AND (IL.STORAGEBINNUMBER = " & sBinCode & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) AND IR.RECEIPTNUMBER = (SELECT ACTUALRECEIPTNO FROM SUB_T_SALEARHEADER WHERE SOURCEREFNO =(SELECT DISTINCT SOURCECONFIRMNO FROM PRD_T_PRODUCTIONORDERHEADER WHERE PRODUCTIONORDERNO =(SELECT DISTINCT PRODUCTIONORDERNO FROM PRD_T_PRODUCTIONSCHEDULEDETAILS WHERE PRODNSCHNO =(SELECT DISTINCT SOURCESCHNO FROM PRD_T_PRODUCTIONMRPHEADER WHERE MRPNO =(SELECT DISTINCT SOURCEREFNO FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRSNo & "))))) GROUP BY IL.STORAGELOCATIONNO,IL.STORAGEBINNUMBER,IR.INVENTORYRECEIPTNO,IL.LOTNUMBER ORDER BY 5"
		else
			sQuery = " SELECT SUM(AVAILABLENETSTOCK),ISNULL(LOTNUMBER,0),ISNULL(SRCTYPE,''),isNull(SERIALNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND "&_
					  " ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND "&_
					  " (LotQuantityNett - QuantityIssued) > 0"
			if trim(sAttList)<>"0" and trim(sAttList)<>"" then
				sQuery = sQuery & " and AttributeList in ('"& sAttList &"')"
			end if
			sQuery = sQuery & " GROUP BY AVAILABLENETSTOCK,LOTNUMBER,SRCTYPE,SERIALNUMBER  "
		.Source = sQuery
		end if
		.ActiveConnection = con
		.Open
	end with
	 'Response.Write dcrs2.Source
	set dcrs2.ActiveConnection = nothing
	sTempLot = ""
	if not dcrs2.EOF then
	sTempLot = ""
		Do While Not dcrs2.EOF
			iLotQtyReserved  = dcrs2(0)
			sLot = trim(dcrs2(1))
			iSer = dcrs2(3)

				if trim(sLot) = "0" then sLot = "N/A"
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			if not iSer = "NULL" and not iSer = "0" then '*****
				iTotLotQty = cdbl(iTotLotQty) + cdbl(trim(dcrs2(0)))
				if  cdbl(iLotQtyReserved) > 0 then

					if sLot <> sTempLot then
					    iPickEntryNoCheck = iPickEntryNoCheck + 1
						Set newElem = oDOM.createElement("PICK")
						newElem.setAttribute "LOC", trim(sStoreCode)
						newElem.setAttribute "BIN", trim(sBinCode)
						newElem.setAttribute "LOTNO", sLot
						newElem.setAttribute "INVRECNO",""' trim(dcrs2(1))
						newElem.setAttribute "QTYISS", ""
						'newElem.setAttribute "SERIALNO",iSer
						newElem.setAttribute "ENTNOCHECK",iPickEntryNoCheck
						Root.appendChild newElem
					end if 'if sLot <> sTempLot then
				end if
			end if'****
			sTempLot = sLot
		dcrs2.MoveNext
		Loop
	else
		iTotLotQty = "0"
		iLotQtyReserved = "0"
		if Trim(sUsage) <> "JWK" then
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				if sLot = "NULL" then
					.Source = "SELECT ISNULL(SUM(QuantityForPick),0) - ISNULL(SUM(QuantityPicked),0) FROM INV_T_MaterialIssuedForPick WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNumber = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND LotNo IS NULL"
				else
					.Source = "SELECT ISNULL(SUM(QuantityForPick),0) - ISNULL(SUM(QuantityPicked),0) FROM INV_T_MaterialIssuedForPick WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sorgID) & " AND LOCATIONNumber = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND LotNo = " & Pack(sLot) & ""
				end if
		'		Response.Write dcrs.source
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iLotQtyReserved = cdbl(dcrs(0))
			end if
			dcrs.close

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
'				Response.Write dcrs.source
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			do while Not dcrs.EOF
				if not (cdbl(trim(dcrs(0))) - cdbl(iTotLotQty)) <= 0 then
				    iPickEntryNoCheck = iPickEntryNoCheck + 1
					Set newElem = oDOM.createElement("PICK")
					newElem.setAttribute "LOC", trim(sStoreCode)
					newElem.setAttribute "BIN", trim(sBinCode)
					newElem.setAttribute "LOTNO", "N/A"
					newElem.setAttribute "INVRECNO", ""
					newElem.setAttribute "QTYISS", ""
					'newElem.setAttribute "SERIALNO",""
					newElem.setAttribute "ENTNOCHECK",iPickEntryNoCheck
					Root.appendChild newElem
				end if
			dcrs.MoveNext
			Loop
			dcrs.close

		end if
	end if
	dcrs2.Close

dcrs1.MoveNext
loop
dcrs1.Close
'Response.End
sLot = ""
iLotQtyReserved = 0
oDOM.Save server.MapPath("../temp/transaction/MRSPICKISSUE"&Session.SessionID&".xml")

arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoMDesc = arrUoM(1)

'sItemName = ItemDisplay(iItem,iClass)
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ITEMDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sItemName = trim(dcrs(0))
end if
dcrs.close
if trim(sOptName)<>"" then  sItemName = sItemName &" ["& sOptName &"]"
sTempArrAttribute = sAttID &"#"& sAttList &":"& sOptName
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : Sales Invoice - Bag Selection</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/MRSPICKISSUE"&Session.SessionID&".xml"%>"></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemMergePickPopModern.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0  onLoad="fnInit('<%=iItem%>','<%=iClass%>','<%=iEntNo%>','<%=sTempArrAttribute%>')">

<form method="POST" name="formname" action="">
<input type="hidden" name="hAttrList" value="<%=sAttList%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"> Mark Details
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
                                            <td class="FieldCell">Unit Name</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly"><%=sOrgname%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Item Description&nbsp;</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="idItemName"><%=sItemName%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                        <% If iQty <> "" then %>
                                            <td class="FieldCell">Quantity Pending&nbsp;</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idQty"><%=iQty%></span>
												<span class="DataOnly"><%=sUoMDesc%></span>
                                            </td>
                                         <% end if %>
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
								<td>
									<div class="frmBody" id="frm2" style="width: 100%; height:140;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Store -- Bin</td>
												<td class="ExcelHeaderCell" align="center">Lot Number</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Quantity Issue</td>
											<%' if sType = "F" then 'Firm Based%>
												<td class="ExcelHeaderCell" align="center" width="40">Serial Number</td>
											<%' end if %>
												<td class="ExcelHeaderCell" align="center" width="10">No.of Pack</td>
											</tr>
										<%Dim sTemp,iTotStock,rsTemp,sSql,iNewTotStock
										Set rsTemp = Server.CreateObject("ADODB.recordset")
										iTotLotQty = 0
										iLotQtyReserved = 0
										iCtr	 = 0
										iTempCtr = 0
										Response.Write "<font color=#000000>"
										with dcrs1
											.CursorLocation = 3
											.CursorType = 3
											.Source = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM VWItemStockStatus WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
											.ActiveConnection = con
											.Open
										end with
									'	Response.Write dcrs1.Source
										set dcrs1.ActiveConnection = nothing
										iTotStock = 0
										Do While Not dcrs1.EOF
											iTotLotQty = 0
											iLotQtyReserved = 0
											sStoreCode = trim(dcrs1(0))
											sBinCode = trim(dcrs1(1))


											with dcrs2
												.CursorLocation = 3
												.CursorType = 3
												if Trim(sUsage) = "JWK" then
													.Source = "SELECT DISTINCT ISNULL(SUM(LOTQUANTITYNETT),0) - ISNULL(SUM(QUANTITYISSUED),0),ISNULL(IL.LOTNUMBER,0),SRCTYPE FROM INV_T_RECEIPTDETAILS IR,INV_T_RECEIPTLOTDETAILS IL WHERE IR.ITEMCODE = " & iItem & " AND IR.ITEMCODE = IL.ITEMCODE AND IR.CLASSIFICATIONCODE = " & iClass & " AND IR.CLASSIFICATIONCODE = IL.CLASSIFICATIONCODE AND IR.ORGANISATIONCODE = " & Pack(sOrgID) & " AND IR.ORGANISATIONCODE = IL.ORGANISATIONCODE AND IR.INVENTORYRECEIPTNO = IL.INVENTORYRECEIPTNO AND (IL.STORAGELOCATIONNO = " & sStoreCode & " AND (IL.STORAGEBINNUMBER = " & sBinCode & " OR IL.STORAGEBINNUMBER IS NULL)) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND CONVERT(DATETIME,ACCOUNTEDON,103) BETWEEN CONVERT(DATETIME," & Pack(sFinFrom) & ",103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) AND IR.RECEIPTNUMBER = (SELECT ACTUALRECEIPTNO FROM SUB_T_SALEARHEADER WHERE SOURCEREFNO =(SELECT DISTINCT SOURCECONFIRMNO FROM PRD_T_PRODUCTIONORDERHEADER WHERE PRODUCTIONORDERNO =(SELECT DISTINCT PRODUCTIONORDERNO FROM PRD_T_PRODUCTIONSCHEDULEDETAILS WHERE PRODNSCHNO =(SELECT DISTINCT SOURCESCHNO FROM PRD_T_PRODUCTIONMRPHEADER WHERE MRPNO =(SELECT DISTINCT SOURCEREFNO FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRSNo & "))))) GROUP BY IL.LOTNUMBER,SRCTYPE ORDER BY 2"
												else
													sQuery = " SELECT AVAILABLENETSTOCK,INVENTORYRECEIPTNO,ISNULL(LOTNUMBER,0),ISNULL(SRCTYPE,''),ISNULL(SERIALNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND  "&_
															  " ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (LotQuantityNett - QuantityIssued > 0) "

													if trim(sAttList)<>"0" and trim(sAttList)<>"" then
														sQuery = sQuery & " and AttributeList in ('"& sAttList &"')"
													end if

													sQuery = sQuery &" GROUP BY AVAILABLENETSTOCK,INVENTORYRECEIPTNO,LOTNUMBER,SRCTYPE,SERIALNUMBER "

													Response.Write "<font color=#000000>"
													'Response.Write "<textarea>"& sQuery &"</textarea>"
													
													.Source = sQuery

												end if
												.ActiveConnection = con
												.Open
											end with
											set dcrs2.ActiveConnection = nothing
											sTemp = ""


											if  dcrs2.EOF then
												bChkFlag = True
											end if

											if not dcrs2.EOF then
												bFlag = true
												bChkFlag = false
												iCtr = iCtr + 1
												Do While Not dcrs2.EOF
'						
													sLot = trim(dcrs2(2))
													sSrcType = trim(dcrs2(3))
													iTotStock =  cdbl(dcrs2(0))

													if sTemp <> sLot then 'iTotStock = 0
														If sIssueCode <> "" then
															sSql = " SELECT AVAILABLENETSTOCK,INVENTORYRECEIPTNO,ISNULL(LOTNUMBER,0),ISNULL(SRCTYPE,''),ISNULL(SERIALNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND  "&_
															  "  ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND INVENTORYRECEIPTNO =  "& sIssueCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (LotQuantityNett - QuantityIssued > 0) "

															    if trim(sAttList)<>"0" and trim(sAttList)<>"" then
																	sSql = sSql & " and AttributeList in ('"& sAttList &"')"
																end if

															 sSql = " GROUP BY AVAILABLENETSTOCK,INVENTORYRECEIPTNO,LOTNUMBER,SRCTYPE,SERIALNUMBER "
															' Response.Write sSql
														else
														
															sSql = " SELECT AVAILABLENETSTOCK,INVENTORYRECEIPTNO,ISNULL(LOTNUMBER,0),ISNULL(SRCTYPE,''),ISNULL(SERIALNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND  "&_
															  " ORGANISATIONCODE = " & Pack(sOrgID) & " AND STORAGELOCATIONNO = " & sStoreCode & " AND (STORAGEBINNUMBER = " & sBinCode & " OR STORAGEBINNUMBER IS NULL) AND (LotQuantityNett - QuantityIssued > 0) "

															    if trim(sAttList)<>"0" and trim(sAttList)<>"" then
																	sSql = sSql & " and AttributeList in ('"& sAttList &"')"
																end if

															 sSql = sSql &" GROUP BY AVAILABLENETSTOCK,INVENTORYRECEIPTNO,LOTNUMBER,SRCTYPE,SERIALNUMBER "
															' Response.Write sSql
														end if
														 'Response.Write sSql
														rsTemp.Open sSql,con
														IF not rsTemp.EOF then
															iTotStock  = rsTemp(0)
														End IF
														rsTemp.Close

													end if 'if sTemp <> sLot then 'iTotStock = 0

													if sLot = "0" or sLot="NULL" then sLot = "N/A"

													bSerialFlag = false
													if sLot <> "N/A" then
														if trim(dcrs2(4)) <>  "0" then
															bSerialFlag = true
														end if
													elseif sLot = "N/A" then
														bSerialFlag = false
														bSerialFlag = true
													end if 'if sLot <> "N/A" then

												if bSerialFlag then

													if cdbl(dcrs2(0)) > 0 then
														iTempCtr = iTempCtr + 1
														if bFlag then
															'iCtr = iCtr + 1
															bFlag = false
															sStoreName = DisplayStore(sStoreCode,sBinCode)
															if trim(sLot)<>"N/A" then
															    sSql = " SELECT SUM(AVAILABLENETSTOCK) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND (LOTNUMBER = '"&sLot&"') "
															else
															    sSql = " SELECT SUM(AVAILABLENETSTOCK) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " AND (LOTNUMBER = '"&sLot&"' OR LOTNUMBER is Null) "
															end if
															if trim(sAttList)<>"0" and trim(sAttList)<>"" then
																sSql = sSql & " and AttributeList in ('"& sAttList &"')"
															end if

															'Response.Write sSql
															rsTemp.Open sSql,con

															IF not rsTemp.EOF then
																iTotStock  =  rsTemp(0)
																'Response.Write "rsTemp="&iTotStock
															End IF
															rsTemp.Close
										%>
														<tr>
															<td class="ExcelSerial" align="center" rowspan="<%=dcrs2.RecordCount%>"><%=iCtr%></td>
															<td class="ExcelDisplayCell" rowspan="<%=dcrs2.RecordCount%>"><%=sStoreName%></td>
															<td class="ExcelDisplayCell"><p align="left"><span class=dataonly onMouseOver="ShowDet('<%=sInspDet%>')" onMouseOut="ShowDet('-')"><%=sLot%></span></td>
															<td class="ExcelDisplayCell" width="10">
																<input type="text" name="txtQty<%=iCtr%>" value="<%=iTotStock%>" size="11" class="FormElemRead" READONLY style="text-align=right">
															</td>

														<%
														'Response.Write "sType="& sType
																if sType = "F" then
																	iCtr = iCtr + 1
																	if bSerialFlag then

																		if sSrcType <> "RW" then
														%>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" READONLY  style="text-align:right">
																			</td>
																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" onClick="CheckLot(this,'<%=iEntNo%>','<%=iClass%>','<%=iItem%>','<%=sOrgID%>','<%=sOptName%>','<%=sAttID%>','<%=sAttList%>')">
																			</td>
														<%				else %>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" READONLY  style="text-align=right">
																			</td>
																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																			</td>
														<%
																		end if
																	else

																		if sSrcType <> "RW" then

														%>
																	<td class="ExcelInputCell" width="10">
																		<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align:right">
																	</td>
																	<td class="ExcelFieldCell" align="center">
																		<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																	</td>
														<%				else %>
																	<td class="ExcelInputCell" width="10">
																		<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" READONLY  style="text-align=right">
																	</td>
																	<td class="ExcelFieldCell" align="center">
																		<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																	</td>
														<%

																		end if
																	end if'if bSerialFlag then

																else
																'Response.Write bSerialFlag
																	sRecNumFlag = ""
																	sSql ="Select ReceiptNumbering from INV_M_ItemMaster where ItemCode ="& iItem &"  and OrganisationCode = "& Pack(sOrgID)&""
																	with dcrs3
																		.CursorLocation = 3
																		.CursorType = 3
																		.ActiveConnection = con
																		.Source = sSql
																		.Open
																	end with
																	if not dcrs3.EOF then
																		sRecNumFlag = dcrs3(0)
																	end if
																	dcrs3.Close
																	'Response.Write sSql

																		if sTemp <> sLot then
																		'iTotStock = 0
																			'iTempCtr = iTempCtr + 1
																			'iCtr = iCtr + 1

																					if sSrcType <> "RW" then
																						'if 	Not sLot = "N/A" then
																						if trim(sRecNumFlag)="S" or trim(sRecNumFlag)="LS" then
																							%>
																								<td class="ExcelInputCell" width="10">
																									<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" READONLY>
																								</td>
																							<%else %>
																							<td class="ExcelInputCell" width="10">
																								<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" >
																							</td>
																					  <%end if
																					else
																					'if 	Not sLot = "N/A" then
																						if trim(sRecNumFlag)="S" or trim(sRecNumFlag)="LS" then %>
																								<td class="ExcelInputCell" width="10">
																									<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" READONLY>
																								</td>
																						<%else %>
																								<td class="ExcelInputCell" width="10">
																									<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" >
																								</td>

																						<%end if %>

																	<%				end if 'if sSrcType <> "RW" then

																					if bSerialFlag then

																						'if sLot <> "N/A" then
																						if trim(sRecNumFlag)="S" or trim(sRecNumFlag)="LS" then %>
																							<td class="ExcelFieldCell" align="center">
																								<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" onClick="CheckLot(this,'<%=iEntNo%>','<%=iClass%>','<%=iItem%>','<%=sOrgID%>','<%=sOptName%>','<%=sAttID%>','<%=sAttList%>')">
																							</td>
																						<%else %>
																							<td class="ExcelFieldCell" align="center">
																								<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																							</td>

																						<%end if
																					end if  'if bSerialFlag then
																		end if ' if sTemp <> sLot then
																		sTemp = sLot
																end if  'if sType = "F" then

														%>
														<td class="ExcelDisplayCell" align="left" width="10">
															<input type=text name='txtTotPackZ<%=iCtr%>' value="0" size=5 class="FormElemRead" ReadOnly>
														</td>

														</tr>
												<%else 'if bFlag then
													'iCtr = iCtr + 1
														sStoreName = DisplayStore(sStoreCode,sBinCode)
														if trim(sLot) <>"N/A" then
														    sSql = " SELECT SUM(AVAILABLENETSTOCK) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " and ClassificationCode ="& iClass &" AND (LOTNUMBER = '"&sLot&"') "
													    else
													        sSql = " SELECT SUM(AVAILABLENETSTOCK) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItem & " and ClassificationCode ="& iClass &" AND (LOTNUMBER = '"&sLot&"' OR LOTNUMBER IS NULL) "
													    end if

														if trim(sAttList)<>"0" and trim(sAttList)<>"" then
															sSql = sSql & " and AttributeList in ('"& sAttList &"')"
														end if


													'	Response.Write sSql
														rsTemp.Open sSql,con

														IF not rsTemp.EOF then
															iTotStock  =  rsTemp(0)
														'	Response.Write "rsTemp="&iTotStock
														End IF
														rsTemp.Close



																sRecNumFlag = ""
																sSql ="Select ReceiptNumbering from INV_M_ItemMaster where ItemCode ="& iItem &"  and OrganisationCode = "& Pack(sOrgID)&""
																with dcrs3
																	.CursorLocation = 3
																	.CursorType = 3
																	.ActiveConnection = con
																	.Source = sSql
																	.Open
																end with
																if not dcrs3.EOF then
																	sRecNumFlag = dcrs3(0)
																end if
																dcrs3.Close

																if sTemp <> sLot then
																'iTotStock = 0
															'	iTempCtr = iTempCtr + 1
																iCtr = iCtr + 1
																'Response.Write "iTotStock="&iTotStock
										%>
														<tr>
															<td class="ExcelDisplayCell"><p align="left"><span class=dataonly onMouseOver="ShowDet('<%=sInspDet%>')" onMouseOut="ShowDet('-')"><%=sLot%></span></td>
															<td class="ExcelDisplayCell" width="10">
																<input type="text" name="txtQty<%=iCtr%>" value="<%=iTotStock%>" size="11" class="FormElemRead" READONLY style="text-align:right">
															</td>
														<%		if sType = "F" then
																	if bSerialFlag then
																		if sSrcType <> "RW" then
														%>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" READONLY  style="text-align=right">
																			</td>
																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" onClick="CheckLot(this,'<%=iEntNo%>','<%=iClass%>','<%=iItem%>','<%=sOrgID%>','<%=sOptName%>','<%=sAttID%>','<%=sAttList%>')">
																			</td>
														<%				else %>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" READONLY  style="text-align=right">
																			</td>
																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																			</td>

														<%				end if
																	else


																		if sSrcType <> "RW" then
														%>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right">
																			</td>
																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																			</td>
														<%				else %>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" READONLY  style="text-align=right">
																			</td>
																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																			</td>

														<%				end if
																	end if
															else
																 'Response.Write sSrcType & bSerialFlag

																	if sSrcType <> "RW" then
																		if 	bSerialFlag then
																			'if sLot <> "N/A" then
																			if trim(sRecNumFlag)="S" or trim(sRecNumFlag)="LS" then%>
																					<td class="ExcelInputCell" width="10">
																						<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" READONLY>
																					</td>
																			<%else %>
																					<td class="ExcelInputCell" width="10">
																						<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" >
																					</td>
																			<%end if
																		else %>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" >
																			</td>

																		<%end if
																	else
																		if 	bSerialFlag then%>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" READONLY>
																			</td>
																		<%else %>
																			<td class="ExcelInputCell" width="10">
																				<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right" >
																			</td>

																		<%end if
																	end if'if sSrcType <> "RW" then

																	if bSerialFlag then
																		'if sLot <> "N/A" then
																		if trim(sRecNumFlag)="S" or trim(sRecNumFlag)="LS" then%>
																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" onClick="CheckLot(this,'<%=iEntNo%>','<%=iClass%>','<%=iItem%>','<%=sOrgID%>','<%=sOptName%>','<%=sAttID%>','<%=sAttList%>')">
																			</td>
																		<%else %>

																			<td class="ExcelFieldCell" align="center">
																				<input type="button" name="BtnSerial:<%=iTempCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>" value="Pick" class="AddButtonX" DISABLED>
																			</td>
																		<%end if 'if trim(sRecNumFlag)="S" or trim(sRecNumFlag)="LS" then
																	end if  'if bSerialFlag then

																end if

														%>
															<td class="ExcelDisplayCell" align="left" width="10">
															<input type=text name='txtTotPackZ<%=iCtr%>' value="0" size=5 class="FormElemRead" ReadOnly>
														</td>
														</tr>
										<%					end if'	if sTemp <> sLot then
															sTemp = sLot

														end if
														end if
													end if
												dcrs2.MoveNext
												Loop

											else

												sLot = "N/A"
												if Not bChkFlag then
													with dcrs
														.CursorLocation = 3
														.CursorType = 3
														'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM Inv_T_ItemYearlyStock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sStoreCode & " AND (BINNUMBER = " & sBinCode & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
														.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM Inv_T_ItemYearlyStock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) GROUP BY ORGANISATIONCODE"
														.ActiveConnection = con
														.Open
													end with
													set dcrs.ActiveConnection = nothing
													if Not dcrs.EOF then
														bChkFlag = false

														bSerialFlag = false
														if not cdbl(dcrs(0)) <= 0 then
															sStoreName = DisplayStore(sStoreCode,sBinCode)
												%>
													<tr>
														<td class="ExcelSerial" align="center"><%=iCtr%></td>
														<td class="ExcelDisplayCell"><%=sStoreName%></td>
														<td class="ExcelDisplayCell"><p align="left">N/A</td>
														<td class="ExcelDisplayCell" width="10">
															<input type="text" name="txtQty<%=iCtr%>" value="<%=iTotStock%>" size="11" class="FormElemRead" READONLY style="text-align=right">
														</td>

													<% 	if sTemp <>  sLot then
															if sType = "F" then
																if bSerialFlag then
													%>
																<td class="ExcelInputCell" width="10">
																	<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" READONLY style="text-align=right">
																</td>
																<td class="ExcelFieldCell" align="center">
																	<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>" value="Pick" class="AddButtonX" onClick="CheckLot(this,'<%=iEntNo%>','<%=iClass%>','<%=iItem%>','<%=sOrgID%>','<%=sOptName%>','<%=sAttID%>','<%=sAttList%>')">
																</td>
													<%			else %>
																<td class="ExcelInputCell" width="10">
																	<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right">
																</td>
																<td class="ExcelFieldCell" align="center">
																	<input type="button" name="BtnSerial:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>" value="Pick" class="AddButtonX" DISABLED>
																</td>
													<%
																end if
															else
													%>
																<td class="ExcelInputCell" width="10">
																	<input type="text" name="txtIss<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="FormElem" style="text-align=right">
																</td>
													<% 		end if
														end if ' if sTemp <> sLot then%>
													</tr>
												<%
														end if
													end if
													sTemp = sLot

													dcrs.Close
												end if 'if not bChkFlag then
											end if
											dcrs2.Close

										dcrs1.MoveNext
										loop
										dcrs1.Close

											if bChkFlag then
										%>
											<tr>
												<td colspan=7 class="ExcelDisplayCell" align="center"><B>No Stock Available</B></td>
											</tr>
										<%	end if%>
										<input type="hidden" name="hCtr" value="<%=iCtr%>">
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
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton">
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
	' Function to get Store
	Function DisplayStore(sLoc,sBin)
		' Declaration of variables
		Dim dcrs,sBinName,sLocName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT LOCATIONNAME,LOCATIONCODE FROM Inv_M_Storage WHERE LOCATIONNUMBER = " & sLoc & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			sLocName = trim(dcrs(0))
		else
			sLocName = "-"
		end if
		dcrs.close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY BINNUMBER"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			DisplayStore = trim(sLocName)&" -- "&trim(dcrs(0))
		else
			DisplayStore = trim(sLocName)
		end if
		dcrs.Close
	End Function
%>

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
'			Response.Write dcrs.source
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
%>

<%
	' Function to get the Pass Count from the Inspection table of Purchase
	Function GetInspDetails(iInvRecNo,sLot)
		if sLot = "N/A" then
			GetInspDetails = "-"
			exit function
		end if

		' Declaration of variables
		Dim dcrs,sFrom,sTo
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT PASSEDFORCOUNTFROM,PASSEDFORCOUNTTO FROM RCV_T_PURCHINSPECTIONHEADER WHERE RECEIPTNUMBER = (SELECT RECEIPTNUMBER FROM INV_T_RECEIPTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ") AND PASSEDFORCOUNTFROM IS NOT NULL AND PASSEDFORCOUNTTO IS NOT NULL"
			.Source = "SELECT PASSEDFORCOUNTFROM,PASSEDFORCOUNTTO FROM RCV_T_PURCHINSPECTIONHEADER WHERE RECEIPTNUMBER = (SELECT RECEIPTNUMBER FROM RCV_T_ActualReceiptHeader WHERE INVENTORYRECNO = " & iInvRecNo & ") "

			.ActiveConnection = con
			.Open
		end with
		'Response.Write DCRS.SOURCE

		set dcrs.ActiveConnection = nothing
		if Not dcrs.EOF then
			sFrom = trim(dcrs(0))
			sTo = trim(dcrs(1))
			if sFrom <> "" or sTo <> "" then
				GetInspDetails = sFrom&"|"&sTo
			else
				GetInspDetails = "-"
			end if
		else
			GetInspDetails = "-"
		end if
		dcrs.Close
	End Function
%>
