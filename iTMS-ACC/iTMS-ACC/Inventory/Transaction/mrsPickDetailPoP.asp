<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsPickDetailPoP.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 17,2012
	'Modified By				:
	'Modified On				:
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
Dim dcrs,dcrs1,dcrs2,dcrs3,iCtr,bexists,rsStore,rsStock,rsTemp
Dim sSql


'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
Set dcrs3 = Server.CreateObject("ADODB.RecordSet")
set rsStore = Server.CreateObject("ADODB.Recordset")
set rsStock = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")

dim oDom,Root,PageNode,HeaderNode,PGNode,objfs,newElem

dim sql,sItemTypeName,sUnitName,sUsageName,iItem,iClass
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,bFlag,iQty,bChkFlag,sLot,iLotQtyReserved
dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore,iTotLotQty,iTempCtr
dim arrUoM,sUoMDesc,sUoMCode,sItemName,sUsage,sInspDet,sIssueCode,sTempLot
dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sType,bSerialFlag,iSer,nChkFlag
Dim sAttList,sQuery,sAttID,sTempArrAttribute,sStoreFlag
dim iEntNo,sOptName,sAttributeList,sPickPackFlag,sIssueType,iTotStock,sRcptNumbering
Dim sDisableButton,sReadOnlyText,iCtrXML
Dim sLotAttrList,sLotAttrListName,sArrAtt
sInspDet = "-"
iTotLotQty = 0
bChkFlag = true

if len(Month(date())) = 1 then
	sTempMonYr = "0"&Month(date())
else
	sTempMonYr = Month(date())
end if
sMonYr = sTempMonYr&Year(date())
arrFin = split(session("Finperiod"),":")
sFinFrom = "01/04/"&arrFin(0)
sFinTo = "31/03/"&arrFin(1)


set oDom = server.CreateObject("Microsoft.xmlDom")
Set objfs = CreateObject("Scripting.FileSystemObject")

sAttributeList = Request.QueryString("AttributeList")
sPickPackFlag = Request("PickPackFlag")
arrTemp = split(trim(Request.QueryString("sTemp")),":")



iClass	= arrTemp(1)
iItem	= arrTemp(0)
'Response.Write iItem
iMRSNo = arrTemp(2)
iEntNo = arrTemp(3)
sOptName = arrTemp(4)
sStoreCode = arrTemp(5)
sBinCode = arrTemp(6)
iQty = arrTemp(7)
sType = arrTemp(8)
sUsage = arrTemp(9)
sOrgID = arrTemp(10)
Response.Write "<font color=#000000> "

if UBound(arrTemp)>12 then
    sAttList = arrTemp(12)
    sAttID	= arrTemp(11)
end if
if UBound(arrTemp) > 10 then
    sAttID	= arrTemp(11)
end if


if trim(sAttributeList)<>"" then
    sArrAtt = split(sAttributeList,"$")
    if UBound(sArrAtt)>0 then
    sAttID = split(sArrAtt(1),"@")(0)
    else
    sAttID = sArrAtt(0)
    end if
else
    sAttID = ""
end if 'if trim(sAttributeList)<>"" then


if sBinCode = "N" then sBinCode = "0"

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
if iMRSNo <> "" then
	with dcrs2
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT MRSFORUNIT,ORGUNITSHORTDESCRIPTION,CONVERT(CHAR,MRSDATE,103),ISNULL(MRSCODE,MRSNUMBER) FROM VWMRSLIST WHERE MRSNUMBER = " & iMRSNo & " "'AND ISNULL(ICOUNTER,0) = "& iEntNo &" "
		.ActiveConnection = con
		.Open
	end with
	set dcrs2.ActiveConnection = nothing

	if not dcrs2.EOF then
		sOrgID = trim(dcrs2(0))
		sOrgName = trim(dcrs2(1))
		dMRSDate = trim(dcrs2(2))
		sIssueCode = trim(dcrs2(3))
	end if
	dcrs2.Close
end if


sRcptNumbering = ""

sSql = "Select ReceiptNumbering from VWItem where Itemcode ="& iItem
dcrs.open sSql,con
if not dcrs.eof then
    sRcptNumbering = trim(dcrs(0))
end if
dcrs.close


set oDom = nothing

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

Set Root = oDOM.createElement("Pick")
Root.setAttribute "TOT",""
oDOM.appendChild Root

iTotLotQty = 0


iCtrXML = 0

sSql = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM VWItemStockStatus WHERE ITEMCODE = "& iItem
sSql = sSql & " AND CLASSIFICATIONCODE="& iClass &" AND ORGANISATIONCODE = "& Pack(sOrgID) &" AND APPLICABLEFOR = 'IN' "
sSql = sSql & " AND CONVERT(DATETIME,"& pack(sFinFrom) &",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) "
sSql = sSql & " AND CONVERT(DATETIME,"& pack(sFinTo) &",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) "
rsStore.open sSql,con
if not rsStore.eof then
    do while not rsStore.eof

        sStoreCode = rsStore(0)
        sBinCode = rsStore(1)
        sStoreName = DisplayStore(sStoreCode,sBinCode)

    	sSql = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID= '"& sOrgID &"' "
    	dcrs2.open sSql,con
		if not dcrs2.EOF then
		    sOrgName = dcrs2(0)
	    end if
	    dcrs2.Close

        if trim(sPickPackFlag)="N" or trim(sType)="F" then

            sSql = " SELECT SUM(AVAILABLENETSTOCK),ISNULL(LOTNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK "
            sSql = sSql & " WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND  ORGANISATIONCODE = "& pack(sOrgID) &" AND (LotQuantityNett - QuantityIssued > 0) "

            if Trim(sAttID)<>"" and (not IsNull(sAttID)) then
                sSql = sSql & " and AttributeList = '"& sAttID &"'"
            end if

            if trim(sStoreCode)<>"" then
                 sSql = sSql & " AND STORAGELOCATIONNO = " & sStoreCode
            end if

            sSql = sSql & " GROUP BY LOTNUMBER "
            'Response.Write "<textarea>"& sSql &"</textarea>"
            rsStock.open sSql,con
            if rsStock.eof then
                bChkFlag =  true
            end if
            if not rsStock.eof then
                bChkFlag =  false
                do while not rsStock.eof
                    iCtrXML = iCtrXML + 1
                    iTotStock = rsStock(0)
                    sLot = rsStock(1)
                    if sLot = "0" or sLot="NULL" then sLot = "N/A"
                    if trim(sRcptNumbering)="N" then
                            Set newElem = oDOM.createElement("STORE")
					        newElem.setAttribute "LOC", trim(sStoreCode)
					        newElem.setAttribute "BIN", trim(sBinCode)
					        newElem.setAttribute "LOTNO", sLot
					        newElem.setAttribute "INVRECNO",""
					        newElem.setAttribute "QTYISS", ""
					        newElem.setAttribute "Count",iCtrXML
					        Root.appendChild newElem
                    else
                            Set newElem = oDOM.createElement("PICK")
						    newElem.setAttribute "LOC", trim(sStoreCode)
						    newElem.setAttribute "BIN", trim(sBinCode)
						    newElem.setAttribute "LOTNO", sLot
						    newElem.setAttribute "INVRECNO",""
						    newElem.setAttribute "QTYISS", ""
						    newElem.setAttribute "Count",iCtrXML
						    Root.appendChild newElem
					end if

                    rsStock.movenext
                loop
            end if
            rsStock.close
        else 'if trim(sPickPackFlag)="L" then
            sSql = " SELECT SUM(AVAILABLENETSTOCK) FROM VW_ITEMLOCATIONLOT_STOCK "
            sSql = sSql & " WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND  ORGANISATIONCODE = "& pack(sOrgID) &" AND (LotQuantityNett - QuantityIssued > 0) "

           if Trim(sAttID)<>"" and (not IsNull(sAttID)) then
                sSql = sSql & " and AttributeList = '"& sAttID &"'"
            end if

            if trim(sStoreCode)<>"" then
                 sSql = sSql & " AND STORAGELOCATIONNO = " & sStoreCode
            end if

            rsStock.open sSql,con
            if rsStock.eof then
                bChkFlag =  true
            end if
            if not rsStock.eof then
                bChkFlag =  false
                do while not rsStock.eof
                    iCtrXML = iCtrXML + 1
                    iTotStock = rsStock(0)
                    if sLot = "0" or trim(sLot)="" or sLot="NULL" then sLot = "N/A"
                    if trim(sRcptNumbering)="N" then
                            Set newElem = oDOM.createElement("STORE")
					            newElem.setAttribute "LOC", trim(sStoreCode)
					            newElem.setAttribute "BIN", trim(sBinCode)
					            newElem.setAttribute "LOTNO", sLot
					            newElem.setAttribute "INVRECNO",""
					            newElem.setAttribute "QTYISS", ""
					            newElem.setAttribute "Count",iCtrXML
					            Root.appendChild newElem
                    else
                            Set newElem = oDOM.createElement("PICK")
						        newElem.setAttribute "LOC", trim(sStoreCode)
						        newElem.setAttribute "BIN", trim(sBinCode)
						        newElem.setAttribute "LOTNO", sLot
						        newElem.setAttribute "INVRECNO",""
						        newElem.setAttribute "QTYISS", ""
						        newElem.setAttribute "Count",iCtrXML
						        Root.appendChild newElem
					end if
                    rsStock.movenext
                loop
            end if
            rsStock.close
        end if 'if trim(sPickPackFlag)="N" or trim(sType)="F" then
        rsStore.movenext
    loop
end if
rsStore.close
sLot = ""
iLotQtyReserved = 0
oDOM.Save server.MapPath("../temp/transaction/MRSPICKISSUE"&Session.SessionID&".xml")

arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoMDesc = arrUoM(1)

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
sItemName = sItemName & sOptName



if trim(sRcptNumbering)="N" then
    sReadOnlyText =""
    sDisablebutton = "Disabled"
elseif trim(sRcptNumbering)="S" then
    if trim(sType)="F" then
        sReadOnlyText ="readonly"
        sDisablebutton = ""
    else
        if trim(sPickPackFlag)="L" then
            sReadOnlyText =""
            sDisablebutton = "Disabled"
        else
            sReadOnlyText ="readonly"
            sDisablebutton = ""
        end if
    end if
else
    if trim(sType)="F" then
        sReadOnlyText ="readonly"
        sDisablebutton = ""
    else
        sReadOnlyText =""
        sDisablebutton = "Disabled"
    end if
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : MR Issue - Mark Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/MRSPICKISSUE"&Session.SessionID&".xml"%>"></script>

<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsIssuePick.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=iItem%>','<%=iClass%>','<%=iEntNo%>','<%=sAttributeList%>');EnableLot()">

<form method="POST" name="formname" action="">

<input type="hidden" name="hAttrList" value="<%=sAttList%>">
<input type="hidden" name="hType" value="<%=sType%>">
<input type="hidden" name="hPickPackFlag" value="<%=sPickPackFlag%>">
<input type="hidden" name="hRcptNumbering" value="<%=sRcptNumbering%>">
<input type="hidden" name="hXMLCtr" value="<%=iCtrXML%>">
<input type="hidden" name="hTemp" value="<%=Request.QueryString("sTemp")%>" />
<input type="hidden" name="hOptName" value="<%=sOptName%>" />
<input type="hidden" name="hAttID" value="<%=sAttID%>" />
<input type="hidden" name="hAttList" value="<%=sAttList%>" />
<input type="hidden" name="hOrgID" value="<%=sOrgID%>" />

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Mark Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%"  >
				<TR>
					<TD class="TabBodyWithTopLine">
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
                                            <td class="FieldCell">MR No. - Date</td>
                                            <td width="200" class="FieldCellSub">
                                            <%IF sIssueCode <> "" then %>
                                                <span class="DataOnly"><%=sIssueCode%>&nbsp;</span> -
                                                <span class="DataOnly"><%=dMRSDate%>&nbsp;</span>
											<%Else %>
												<span class="DataOnly">N/A</span>

                                            <%End IF %>
                                            </td>
                                        </tr>
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
                                        <%if trim(sType)="M" and trim(sPickPackFlag)="N" and trim(sRcptNumbering)="LS" then %>
                                        <tr>
                                            <td class="FieldCell">
                                                Select Data
                                            </td>
                                            <td class="FieldCellSub">
                                                <input type="radio" name="radLotPack" value="L" checked onclick="EnableLot()" onfocus="EnableLot()">
                                                Only Lot&nbsp
                                                <input type="radio" name="radLotPack" value="P" onclick="EnableLot()"  onfocus="EnableLot()">
                                                Lot With Pack
                                            </td>
                                        </tr>
                                        <%end if'if trim(sType)="M"  and trim(sPickPackFlag)="N" and trim(sRcptNumbering)="LS"  then %>
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
									<div class="frmbody" id="frm2" style="width: 100%; height:175;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Store -- Bin</td>
												<td class="ExcelHeaderCell" align="center">Lot No</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Quantity Issue</td>
												<td class="ExcelHeaderCell" align="center" width="40">Serial Number</td>
												<td class="ExcelHeaderCell" align="center" width="10">No.of Pack</td>
											</tr>
										    <%
										        Response.write "<font color=red>"


										        iCtr = 0

										        sSql = "SELECT DISTINCT LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM VWItemStockStatus WHERE ITEMCODE = "& iItem
										        sSql = sSql & " AND CLASSIFICATIONCODE="& iClass &" AND ORGANISATIONCODE = "& Pack(sOrgID) &" AND APPLICABLEFOR = 'IN' "
										        sSql = sSql & " AND CONVERT(DATETIME,"& pack(sFinFrom) &",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) "
										        sSql = sSql & " AND CONVERT(DATETIME,"& pack(sFinTo) &",103) = CONVERT(DATETIME,FINANCIALYEARTO,103) "

										        rsStore.open sSql,con
										        if not rsStore.eof then
										            do while not rsStore.eof

										                sStoreCode = rsStore(0)
										                sBinCode = rsStore(1)
										                sStoreName = DisplayStore(sStoreCode,sBinCode)

										                if trim(sPickPackFlag)="N" or trim(sType)="F" then

										                    sSql = " SELECT SUM(AVAILABLENETSTOCK),ISNULL(LOTNUMBER,0) FROM VW_ITEMLOCATIONLOT_STOCK "
										                    sSql = sSql & " WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND  ORGANISATIONCODE = "& pack(sOrgID) &" AND (LotQuantityNett - QuantityIssued > 0) "

                                                            if Trim(sAttID)<>"" and (not IsNull(sAttID)) then
                                                                sSql = sSql & " and AttributeList = '"& sAttID &"'"
                                                            end if

                                                            if trim(sStoreCode)<>"" then
                                                                 sSql = sSql & " AND STORAGELOCATIONNO = " & sStoreCode
                                                            end if

                                                            sSql = sSql & "GROUP BY LOTNUMBER "
                                                            'Response.write "<textarea>"& sSql &"</textarea>"
                                                            rsStock.open sSql,con
                                                            if rsStock.eof then
                                                                bChkFlag =  true
                                                            end if
                                                            if not rsStock.eof then
                                                                bChkFlag =  false
                                                                do while not rsStock.eof
                                                                    iCtr = iCtr + 1
                                                                    iTotStock = rsStock(0)
                                                                    sLot = rsStock(1)
                                                                    if sLot = "0" or sLot="NULL" then sLot = "N/A"


                                                                    if trim(sLot)<>"N/A" then
                                                                        sSql = "Select distinct IsNull(AttributeList,'') from INV_T_LocationLot where LotNumber= '"& sLot &"'"
                                                                        rsTemp.open sSql,con
                                                                        if not rsTemp.eof then
                                                                            sLotAttrList =  rsTemp(0)
                                                                        end if
                                                                        rsTemp.close
                                                                        if trim(sLotAttrList)<>"" then
                                                                            sSql = "Select OptionName from INV_M_ItemTypeOptions where OptionValue = "& sLotAttrList
                                                                            rsTemp.open sSql,con
                                                                            if not rsTemp.eof then
                                                                                sLotAttrListName = trim(rsTemp(0))
                                                                            end if
                                                                            rsTemp.close
                                                                        end if
                                                                    end if'if trim(sLot)<>"N/A" then

                                                                    %>
										                                <tr>
										                                    <td class="ExcelSerial" align="center"><%=iCtr%></td>
											                                <td class="ExcelDisplayCell" align="center"><%=sStoreName%></td>
											                                <td class="ExcelDisplayCell" align="left">
											                                    <span class="Dataonly"><%=sLot%>
											                                    <%if trim(sLotAttrListName)<>"" then%>
											                                        -<%=sLotAttrListName%>
											                                    <%end if %>
											                                    </span></td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="text" name="txtQtyZ<%=iCtr%>" value="<%=iTotStock%>" class="FormElemRead" size="11" style="text-align:right" readonly>

											                                </td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="text" name="txtIssZ<%=iCtr%>" value="0" class="FormElem" size="11" style="text-align:right" <%=sReadOnlyText%>>
											                                </td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="button" name="btnSerialZ<%=iCtr%>" value="Pick" class="AddButtonX" <%=sDisableButton%> onClick="CheckLot('<%=iCtr%>','<%=iEntNo%>','<%=iClass%>','<%=iItem%>','<%=sOrgID%>','<%=sIssueCode%>','<%=sOptName%>','<%=sAttID%>','<%=sAttList%>')">
											                                    <input type="hidden" name="hValueZ<%=iCtr%>" value="hValueZ:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>">
											                                </td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="text" name="txtTotPackZ<%=iCtr%>" value="0" class="FormElemRead" readonly size="5" style="text-align:right">
											                                    <input type="hidden" name="hLotNo<%=iCtr%>" value="<%=sLot%>" />
											                                </td>
												                        </tr>
										                            <%
                                                                    rsStock.movenext
                                                                loop
                                                            end if
                                                            rsStock.close
                                                        elseif trim(sPickPackFlag)="L" then
                                                            sSql = " SELECT SUM(AVAILABLENETSTOCK) FROM VW_ITEMLOCATIONLOT_STOCK "
										                    sSql = sSql & " WHERE ITEMCODE = "& iItem &" AND CLASSIFICATIONCODE = "& iClass &" AND  ORGANISATIONCODE = "& pack(sOrgID) &" AND (LotQuantityNett - QuantityIssued > 0) "

                                                            if Trim(sAttID)<>"" and (not IsNull(sAttID)) then
                                                                sSql = sSql & " and AttributeList = '"& sAttID &"'"
                                                            end if

                                                            if trim(sStoreCode)<>"" then
                                                                 sSql = sSql & " AND STORAGELOCATIONNO = " & sStoreCode
                                                            end if

                                                            rsStock.open sSql,con
                                                            if rsStock.eof then
                                                                bChkFlag =  true
                                                            end if
                                                            if not rsStock.eof then
                                                                bChkFlag =  false
                                                                do while not rsStock.eof
                                                                    iCtr = iCtr + 1
                                                                    iTotStock = rsStock(0)
                                                                    if sLot = "0" or trim(sLot)="" or sLot="NULL" then sLot = "N/A"
                                                                    %>
										                                <tr>
										                                    <td class="ExcelSerial" align="center"><%=iCtr%></td>
											                                <td class="ExcelDisplayCell" align="center"><%=sStoreName%></td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <span class="Dataonly"><%=sLot%></span></td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="text" name="txtQtyZ<%=iCtr%>" value="<%=iTotStock%>" class="FormElemRead" size="11" style="text-align:right" readonly>

											                                </td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="text" name="txtIssZ<%=iCtr%>" value="0" class="FormElem" size="11" style="text-align:right" <%=sReadOnlyText%>>
											                                </td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="button" name="btnSerialZ<%=iCtr%>" value="Pick" class="AddButtonX" <%=sDisableButton%> onClick="CheckLot('<%=iCtr%>','<%=iEntNo%>','<%=iClass%>','<%=iItem%>','<%=sOrgID%>','<%=sIssueCode%>','<%=sOptName%>','<%=sAttID%>','<%=sAttList%>')">
											                                    <input type="hidden" name="hValueZ<%=iCtr%>" value="hValueZ:<%=iCtr%>:<%=sLot%>:<%=sStoreCode%>:<%=sBinCode%>:<%=iTotStock%>">
											                                </td>
											                                <td class="ExcelDisplayCell" align="center">
											                                    <input type="text" name="txtTotPackZ<%=iCtr%>" value="0" class="FormElemRead" readonly size="5" style="text-align:right">
											                                    <input type="hidden" name="hLotNo<%=iCtr%>" value="<%=sLot%>" />
											                                </td>
												                        </tr>
										                            <%
                                                                    rsStock.movenext
                                                                loop
                                                            end if
                                                            rsStock.close
										                end if 'if trim(sPickPackFlag)="N" or trim(sType)="F" then
										                rsStore.movenext
										            loop
										        end if
										        rsStore.close
										        if bChkFlag then
									                %>
									                    <tr>
									                        <td class="ExcelDisplayCell" align="center" colspan="6">No Stock Available</td>
											            </tr>
									                <%
										        end if
										    %>
										</table>
										<input type="hidden" name="hCtr" value="<%=iCtr%>">
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5" alt="">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
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
