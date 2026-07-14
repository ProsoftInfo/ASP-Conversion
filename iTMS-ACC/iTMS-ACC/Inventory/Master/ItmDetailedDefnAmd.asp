<%@ Language=VBScript %>
<%  option explicit	%>
<%
	'Program Name				:	ItmDetailedDefn.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:
	'Created On					:
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->
<!--#include virtual="/include/CommonFunctions.asp"-->
<%
    Dim rsTemp,rsParty,ObjDOM,ndRoot,ndPur,ndSal,ndAlt,ndAltEntry,ndOptUOM,ndOptEntry
    Dim ndVendor,ndBasic,ndDisEntry,ndValEntry
    Dim iItmCode,iClassCode
    Dim sItemName,sClassName,sQuery,sUnder,sOver,sUnOrder,sUoMDesc,sOrgCode,sOrgName,sSalesUOM
    'Alternate Item
    Dim sAltItemCode,sAltClassCode,sAltPeriority,sAltItemName
    'Option UOM
    Dim sUCode,sBRate,sOperator,sOperatorText,sUName
    ''Vendor
    Dim sWarrenty,sTransitLeadTime,sPurLeadTime,sSuppItemNo,sSuppLeadTime
    Dim sMarketPrice,sPreOrdLeadTime,sMarketDate,sPreMinOrdQty,sPreMaxOrdQty
    Dim sSuppCode,sSuppName,sSuppType,sSuppSubType,sSuppItemDesc,sSuppDrawingNo,sSuppUOM,sBuyerName
    ''Pur Basic
    Dim sBuyer,sModVat,sInvMatch,sSubCont,sSubReceipts,sEnforceShipTo,sRecDateAction
    Dim sRecDaysEarly,sRecDaysLate,sUnRecLow,sUnRecHigh,sOverRecLow,sOverRecHigh,sUnOrdRecLow,sUnOrdRecHigh
    ''Sal Basic
    Dim sMarketRate,sWarrentyPeriod,sMinSalQty,sActual,sUnitSize,sUnitSizeUOM,sMinimum,sVolume,sVolumeUOM,sPreferred,sCommodity,sTaxType,sTaxTypeOverride
    'Dis Entry
    Dim sQtyFrom,sQtyTo,sQtyDis,sQtyUoM,sQtyVal,sApplicableIn
    'Val Entry
    Dim sValFrom,sValTo,sValDis
    
    Dim sSql,dcrs,iCnt,nItemRate,sFinPeriod
    Dim sSalItemRate,sSalRatePer,sSalMarPer,sSalMarVal,sSalOthPer,sSalOthVal,sSalItmPrice,sEffFrom

    set dcrs = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    set rsParty = Server.CreateObject("ADODB.Recordset")
    set ObjDOM = Server.CreateObject("Microsoft.XMLDOM")

    Response.Write "<font color=red>"

    iItmCode = Request("ItemCode")
    iClassCode = Request("ClassCode")
    sOrgCode = Session("organizationcode")
    sOrgName = Session("OrgShortName")
    
    if trim(iItmCode)="" then
	%>
	    <script>
	        alert("Please Select the Item in List Tab")
	        window.history.back(-1)
	    </script>
	<%
	end if
	
    
    sFinPeriod = session("Finperiod")
    
    sUnder = 1
    sOver = 1
    sUnOrder = 1

    sQuery = "Select ItemDescription,(Select GroupName from INV_M_Classification where GroupCode = V.ClassificationCode),SalesUOM from VWITEM V where ItemCode = "& iItmCode  &" and ClassificationCode = "& iClassCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sItemName = rsTemp(0)
        sClassName = rsTemp(1)
        sSalesUOM = rsTemp(2)
    end if
    rsTemp.Close

    set ndRoot = ObjDOM.createElement("Root")
    ObjDOM.appendChild ndRoot
    ndRoot.setAttribute "ItemName",sItemName
    ndRoot.setAttribute "ClassName",sClassName
    ndRoot.setAttribute "OrgCode",sOrgCode
    ndRoot.setAttribute "ItemCode",iItmCode
    ndRoot.setAttribute "ClassCode",iClassCode
    ndRoot.setAttribute "OrgName",sOrgName

    set ndPur = ObjDOM.createElement("Purchase")
    ndRoot.appendChild ndPur

    set ndSal = ObjDOM.createElement("Sales")
    ndRoot.appendChild ndSal

    set ndAlt = ObjDOM.createElement("Alternate")
    ndPur.appendChild ndAlt

    sQuery = "Select ItemCode,ClassificationCode,OrganisationCode,AlternateItemCode,AlternateClassification,Prority from INV_M_ItemOrgAlternate where ItemCode = "& iItmCode &" and ClassificationCode = "& iClassCode &" and OrganisationCode = "& sOrgCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        do while not rsTemp.EOF
            sAltItemCode = rsTemp(3)
            sAltClassCode = rsTemp(4)
            sAltPeriority = rsTemp(5)
            sAltItemName = GetItemName(sAltItemCode,sAltClassCode)
            set ndAltEntry = ObjDOM.createElement("Entry")
                ndAltEntry.setAttribute "ITEMCODE",sAltItemCode
                ndAltEntry.setAttribute "CLASSCODE",sAltItemCode
                ndAltEntry.setAttribute "PRIORITY",sAltPeriority
                ndAltEntry.setAttribute "ITEMNAME",sAltItemName
                ndAlt.appendChild ndAltEntry
            rsTemp.MoveNext
        loop
    end if
    rsTemp.Close

    set ndOptUOM = ObjDOM.createElement("OptionalUOM")
    ndPur.appendChild ndOptUOM

    sQuery ="Select ItemCode,ClassificationCode,OrganisationCode,OptionalUoMFor,UoMCode,OptionToBaseRate,OptionToBaseOperator from INV_M_ItemOptionalUOM where ItemCode ="& iItmCode &" and ClassificationCode = "& iClassCode &" and OrganisationCode = "& sOrgCode  &" and OptionalUOMFor='P'"
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        do while not rsTemp.EOF
            sUCode = rsTemp(4)
            sBRate = rsTemp(5)
            sOperator = rsTemp(6)
            sUName = sUCode
            if sOperator = 0 then
                sOperatorText ="*"
            else
                sOperatorText = "/"
            end if

            set ndOptEntry = ObjDOM.createElement("OpUoMEntry")
                ndOptUOM.appendChild ndOptEntry
                ndOptEntry.setAttribute "UCODE",sUCode
                ndOptEntry.setAttribute "BRATE",sBRate
                ndOptEntry.setAttribute "OPERATOR", sOperator
                ndOptEntry.setAttribute "UNAME",sUName
                ndOptEntry.setAttribute "OPERATORTEXT",sOperatorText
                ndOptUOM.appendChild ndOptEntry

            rsTemp.MoveNext
        loop
    end if
    rsTemp.Close

    sQuery = "Select PartyType,PartySubType,PartyCode,SuppItemCode,SuppItemDescription,SupplierDrawingNo,SupplierUOM,"&_
             " MinOrderQuantity,MaxOrderQuantity,PreOrderLeadTime,SuppLeadTime,SuppTransitTime,"&_
             " PurchaseLeadTime,SuppWarrantyPeriod,SuppMarketPrice,Convert(varchar,MarketDate,103) from Inv_R_ItemSupplier where ItemCode ="& iItmCode &" and ClassificationCode = "& iClassCode &" and OrganisationCode = "& sOrgCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
           set ndVendor = ObjDOM.createElement("Vendor")
               ndPur.appendChild ndVendor
               sSuppType = rsTemp(0)
               sSuppSubType = rsTemp(1)
               sSuppCode = rsTemp(2)
               sSuppItemNo = rsTemp(3)
               sSuppItemDesc = rsTemp(4)
               sSuppDrawingNo = rsTemp(5)
               sSuppUOM = rsTemp(6)
               sPreMinOrdQty = rsTemp(7)
               sPreMaxOrdQty = rsTemp(8)
               sPreOrdLeadTime = rsTemp(9)
               sSuppLeadTime = rsTemp(10)
               sTransitLeadTime = rsTemp(11)
               sPurLeadTime  = rsTemp(12)
               sWarrenty = rsTemp(13)
               sMarketPrice = rsTemp(14)
               sMarketDate = rsTemp(15)

               sQuery = "Select PartyName from APP_M_PartyMaster where PartyCode ="& sSuppCode
                rsParty.Open sQuery,con
                if not rsParty.EOF then
                    sSuppName = trim(rsParty(0))
                end if
                rsParty.Close

               ndVendor.setAttribute "Warrenty",sWarrenty
               ndVendor.setAttribute "TransitLeadTime",sTransitLeadTime
               ndVendor.setAttribute "PurLeadTime",sPurLeadTime
               ndVendor.setAttribute "SuppItemNo",sSuppItemNo
               ndVendor.setAttribute "SuppLeadTime",sSuppLeadTime
               ndVendor.setAttribute "MarketPrice",sMarketPrice
               ndVendor.setAttribute "PreOrdLeadTime",sPreOrdLeadTime
               ndVendor.setAttribute "MarketDate",sMarketDate
               ndVendor.setAttribute "PreMinOrdQty",sPreMinOrdQty
               ndVendor.setAttribute "PreMaxOrdQty",sPreMaxOrdQty
               ndVendor.setAttribute "SuppCode",sSuppCode
               ndVendor.setAttribute "SuppName",sSuppName
               ndVendor.setAttribute "SuppType",sSuppType
               ndVendor.setAttribute "SuppSubType",sSuppSubType
               ndVendor.setAttribute "SuppItemDesc",sSuppItemDesc
               ndVendor.setAttribute "SuppDrawingNo",sSuppDrawingNo
               ndVendor.setAttribute "SuppUOM",sSuppUOM
    end if' if not rsTemp.EOF then
    rsTemp.Close

    sQuery ="Select AllowUnderReceipts,UnderRcptLowLimit,UnderRcptHighLimit,AllowOverReceipts,OverRcptLowLimit,"&_
            " OverRcptHighLimit,0,SubContractEligiility,IsNull(InvoiceMatching,0),"&_
            " AllowSubstitutes,AllowUnorderedRcpt,UnOrdRcptLowLimit,UnOrdRcptHighLimit,"&_
            " IsNull(PreferredBuyer,0),AllowShipTo,ItemDefinedBy,ItemDefinedOn,EarlyDays,LateDays from INV_M_ItemOrgPurchase"&_
            " where ItemCode = "& iItmCode &" and ClassificationCode = "& iClassCode &" and OrganisationCode = "& sOrgCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sUnder = rsTemp(0)
        sUnRecLow = rsTemp(1)
        sUnRecHigh = rsTemp(2)
        sOver = rsTemp(3)
        sOverRecLow = rsTemp(4)
        sOverRecHigh = rsTemp(5)
        'sModVat = rsTemp(6)
        sSubCont = rsTemp(7)
        sInvMatch = rsTemp(8)
        sSubReceipts = rsTemp(9)
        sUnOrder = rsTemp(10)
        sUnOrdRecLow =  rsTemp(11)
        sUnOrdRecHigh = rsTemp(12)
        sBuyer = rsTemp(13)
        sEnforceShipTo = rsTemp(14)
        sRecDaysEarly = rsTemp(17)
        sRecDaysLate = rsTemp(18)

        set ndBasic = ObjDOM.createElement("Basic")
        ndPur.appendChild ndBasic
        ndBasic.setAttribute "Buyer",sBuyer
        'ndBasic.setAttribute "ModVat",sModVat
        ndBasic.setAttribute "InvMatch",sInvMatch
        ndBasic.setAttribute "SubCont",sSubCont
        ndBasic.setAttribute "SubReceipts",sSubReceipts
        ndBasic.setAttribute "EnforceShipTo",sEnforceShipTo
        ndBasic.setAttribute "RecDateAction",sRecDateAction
        ndBasic.setAttribute "RecDaysEarly",sRecDaysEarly
        ndBasic.setAttribute "RecDaysLate",sRecDaysLate
        ndBasic.setAttribute "UnRecLow",sUnRecLow
        ndBasic.setAttribute "UnRecHigh",sUnRecHigh
        ndBasic.setAttribute "OverRecLow",sOverRecLow
        ndBasic.setAttribute "OverRecHigh",sOverRecHigh
        ndBasic.setAttribute "UnOrdRecLow",sUnOrdRecLow
        ndBasic.setAttribute "UnOrdRecHigh",sUnOrdRecHigh
        ndPur.appendChild ndBasic
    end if
    rsTemp.Close

    ''' Sales Details

    set ndOptUOM = ObjDOM.createElement("OptionalUOM")
    ndSal.appendChild ndOptUOM

    sQuery ="Select ItemCode,ClassificationCode,OrganisationCode,OptionalUoMFor,UoMCode,OptionToBaseRate,OptionToBaseOperator from INV_M_ItemOptionalUOM where ItemCode ="& iItmCode &" and ClassificationCode = "& iClassCode &" and OrganisationCode = "& sOrgCode  &" and OptionalUOMFor='S'"
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        do while not rsTemp.EOF
            sUCode = rsTemp(4)
            sBRate = rsTemp(5)
            sOperator = rsTemp(6)
            sUName = sUCode
            if sOperator = 0 then
                sOperatorText ="*"
            else
                sOperatorText = "/"
            end if

            set ndOptEntry = ObjDOM.createElement("OpUoMEntry")
                ndOptUOM.appendChild ndOptEntry
                ndOptEntry.setAttribute "UCODE",sUCode
                ndOptEntry.setAttribute "BRATE",sBRate
                ndOptEntry.setAttribute "OPERATOR", sOperator
                ndOptEntry.setAttribute "UNAME",sUName
                ndOptEntry.setAttribute "OPERATORTEXT",sOperatorText
                ndOptUOM.appendChild ndOptEntry

            rsTemp.MoveNext
        loop
    end if
    rsTemp.Close

    sQuery = "Select WarrantyPeriod,PreferredMinQty from INV_M_ItemOrgSales where ItemCode ="& iItmCode &" and ClassificationCode="& iClassCode &" and OrganisationCode = "& sOrgCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sWarrentyPeriod = rsTemp(0)
        sMinSalQty = rsTemp(1)
    end if
    rsTemp.Close
    
    
    
    sQuery = "Select ItemRate,RatePer,MarginPercent,MarginValue,OtherPercent,OtherValue,ItemPrice,EffectiveFrom from Sal_M_UnitPriceDet where ItemCode="& iItmCode &" and ClassificationCode = "& iClassCode &" and OUDefinitionID = " & sOrgCode
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        sSalItemRate = rsTemp(0)
        sSalRatePer = rsTemp(1)
        sSalMarPer  = rsTemp(2)
        sSalMarVal  = rsTemp(3)
        sSalOthPer  = rsTemp(4)
        sSalOthVal  = rsTemp(5)
        sSalItmPrice= rsTemp(6)
        sEffFrom  = rsTemp(7)
    end if
    rsTemp.close
    
     set ndBasic = ObjDOM.createElement("Basic")
        ndSal.appendChild ndBasic
        ndBasic.setAttribute "WarrPeriod",sWarrentyPeriod
        ndBasic.setAttribute "MinSalQty",sMinSalQty
        ndBasic.setAttribute "PurRate",sSalItemRate
        ndBasic.setAttribute "PurRatePer",sSalRatePer
        ndBasic.setAttribute "CharPer",sSalOthPer
        ndBasic.setAttribute "CharValue",sSalOthVal
        ndBasic.setAttribute "MarPer",sSalMarPer
        ndBasic.setAttribute "MarValue",sSalMarVal
        ndBasic.setAttribute "TotPrice",sSalItmPrice
        ndBasic.setAttribute "EffectiveFrom",sEffFrom


    sQuery = "Select QtyDiscountOffered,QuantityFrom,QuantityTo,UoM,Precedence,DiscApplicableOn from INV_M_ItemOrgSaleDiscount where "&_
             "ItemCode="& iItmCode &" and ClassificationCode="& iClassCode &" and OrganisationCode="& sOrgCode &" and QuantityFrom <>0"
    'Response.Write sQuery
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        do while not rsTemp.EOF
            set ndDisEntry = ObjDOM.createElement("DisEntry")
            ndRoot.appendChild ndDisEntry
            sQtyDis = rsTemp(0)
            sQtyFrom =rsTemp(1)
            sQtyTo = rsTemp(2)
            sQtyUoM =rsTemp(3)
            sQtyVal = rsTemp(4)
            sApplicableIn = rsTemp(5)

            ndDisEntry.setAttribute "QTYFROM",sQtyFrom
            ndDisEntry.setAttribute "QTYTO",sQtyTo
            ndDisEntry.setAttribute "QTYDIS",sQtyDis
            ndDisEntry.setAttribute "QTYUOM",sQtyUoM
            ndDisEntry.setAttribute "QTYVAL",sQtyVal
            ndDisEntry.setAttribute "APPIN",sApplicableIn

            rsTemp.MoveNext
        loop
    end if
    rsTemp.Close

    sQuery = "Select ValueDiscountOffered,ValueFrom,ValueTo,Precedence,DiscApplicableOn from INV_M_ItemOrgSaleDiscount Where "&_
             "ItemCode="& iItmCode &" and ClassificationCode="& iClassCode &" and OrganisationCode="& sOrgCode &" and ValueFrom<>0"

    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        do while not rsTemp.EOF
            set ndValEntry = ObjDOM.createElement("ValEntry")
            ndRoot.appendChild ndValEntry
            sValDis  = rsTemp(0)
            sValFrom =rsTemp(1)
            sValTo = rsTemp(2)
            sQtyVal = rsTemp(3)
            sApplicableIn = rsTemp(4)

            ndValEntry.setAttribute "VALFROM",sValFrom
            ndValEntry.setAttribute "VALTO",sValTo
            ndValEntry.setAttribute "VALDIS",sValDis
            ndValEntry.setAttribute "QTYVAL",sQtyVal
            ndValEntry.setAttribute "APPIN",sApplicableIn

            rsTemp.MoveNext
        loop
    end if
    rsTemp.Close

ObjDOM.save(Server.MapPath("../temp/Master/ItemDetailsDefine"&Session.SessionID&".xml"))

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Control Definition - Purchase</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="storageData"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script>
window.ITMS_DETAIL_AMEND = true;
window.ITMS_DETAIL_PUR_SAVE_URL = "XMLSave.asp?SessionFlag=true&Value=PurchaseDetUpdate&Folder=Master";
window.ITMS_DETAIL_PUR_POST_URL = "ItmDetailedPurRecUpdate.asp";
window.ITMS_DETAIL_SAL_SAVE_URL = "XMLSave.asp?SessionFlag=true&Value=SalesDetUpdate&Folder=Master";
window.ITMS_DETAIL_SAL_POST_URL = "ItmDetailedSalesUpdate.asp";
</script>
<script src="../scripts/itemDetailedDefinition.js"></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><root></root></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/Master/ItemDetailsDefine"&Session.SessionID&".xml"%>"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 onload="DisplayData()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hItmCode" value="<%=iItmCode%>">
<input type="hidden" name="hClassCode" value="<%=iClassCode%>">
<input type="hidden" name="hOrgCode" value="<%=sOrgCode%>">
<input type="hidden" name="hSuppName" value="">
<input type="hidden" name="hSuppCode" value="">
<input type="hidden" name="hSuppType" value="">
<input type="hidden" name="hSuppSubType" value="">
<input type="hidden" name="hItemRate" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Detail Definition
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing=0 cellPadding=0 border=0 width="100%">
				<tr>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
							    <td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="ItemListEntryForEdit.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmEditEntry.asp?hItemCode=<%=iItmCode%>">
											<td align="center">Basic
											</td>
											</a>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" width="90">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Purch. & Sales
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmInvDetAmd.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=iClassCode%>">
											<td align="center">Inventory
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmManufactureAmd.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=iClassCode%>">
											<td align="center">Manufacturing
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="ITEMIMPORTEXPORT.ASP">
											    <td align="center">Import/Export Item
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
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellspacing="0" cellpadding="0" class="BodyTable">
									    <tr>
											<td class="FieldCellSub" width="80">Item Name</td>
											<td>
											<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
											<td class="FieldCell" width="15"></td>
											<td class="FieldCell" width="82">Classification</td>
											<td>
											<span class="DataOnly"><%=sClassName%>&nbsp;</span>
											&nbsp;</td>
											<td></td>
									    </tr>
                                    </table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <!----------------->
                            <tr>
                            <td align="center" width="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							<td valign="top">
								<table id="Table" cellspacing="0" cellpadding="0" border="0" width="100%">
									<tr>
										<td>
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td align="center" colspan="3" class="MiddlePack">
														<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
													</td>
												</tr>

                            <tr>
							<td align="center" width="5" class="ClearPixel">
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							<td valign="top" width="100%">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" class=BodyTable>

                            <!----------------->
                            <tr>
                                <td ></td>
								<td align="center">
								<div>
									<table class="CollapseBand" cellspacing="0" cellpadding="0">
										<tr>
											<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(divPurRec,'')" itms_state="0">
												<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
												</a>
											</td>
											<td valign="center" class="SubTitle">&nbsp;&nbsp;
											    Purchasing and Receiving
											</td>
										</tr>
									</table>
									<table border="0" cellpadding="0" cellspacing="0" width=100% class=BodyTable>
									<tr>
										<td width="100%">
                            <div id=divPurRec style="width:100%;display:none;">
                            <table>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
   							       <div align="left">
   								<table cellpadding="0" cellspacing="0" width="100%">
   									<tr>
   										<td>
   								            <table cellpadding="0" cellspacing="0" width="100%">
   									            <tr>
   										            <td class='GroupTitleLeft' width="10">&nbsp;
                                                       </td>
   										            <td class='GroupTitle' width="80"><p align="center">
                                                       Purchasing
                                                       </td>

   										            <td class='GroupTitleRight'><p align="left">&nbsp;
                                                       </td>
   									            </tr>
   								            </table>
                                       </td>
   									</tr>
   									<tr>
   										<td class=GroupTable>
                                   <div align="left">
   								<table cellpadding="0" cellspacing="0" width="100%">
   									<tr>
   										<td class=MiddlePack> </td>
   									</tr>
   									<tr>
   										<td>
                                           <table border="0" cellspacing="0" cellpadding="0">
                                       <tr>
										<td class="FieldCellSub" width="105">Buyer</td>
										<td class="FieldCell" colspan="3">
											<select size="1" name="selBuyer" class="FormElem">
												<option value="select">Select</option>
												<%	'Calling the Function which populates the Employee list
													populateEmployee
												%>
											</select>
										</td>
                                       </tr>
                                       <tr>
										<td class="FieldCellSub">Alternate Item</td>
										<td class="FieldCell" width="110">
										    <input type="button" value="View" name="btnCheck" class="AddButton" onClick="OpenAlter()">
										</td>
										<td class="FieldCellSub" width="85">Optional UoM</td>
										<td class="FieldCell">
											<input type="button" value="View" name="btnUoMPur" class="AddButton" onClick="OpenUoM('Pur')">
										</td>
                                       </tr>


                                       <!--<tr>
										<td class="FieldCellSub" width="110">Modvat</td>
										<td class="FieldCell" width="200">
										    <input type="radio" value="1" name="radMod" class="FormElem">   Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										    <input type="radio" value="0" name="radMod" class="FormElem" CHECKED>   No
										</td>
									   </tr>-->
                                       <tr>
										<td class="FieldCellSub" width="110">Invoice Matching</td>
										<td class="FieldCell">
                                           <select size="1" name="selInvMat" class="FormElem">
												<option value="select">Select</option>
												<option value="2" >2 Way</option>
												<option value="3" >3 Way</option>
												<option value="4" >4 Way</option>
									       </select>
										</td>
										<td class="FieldCellSub" width="110">Sub-Contracting</td>
										<td class="FieldCell">
										    <input type="radio" value="1" name="radSub" class="FormElem">   Yes&nbsp;&nbsp;
										    <input type="radio" value="0" name="radSub" class="FormElem" checked>   No
										</td>
                                       </tr>
                                       <tr>
								            <td class=FieldCellSub width="130">Substitute Receipts</td>
								            <td class='FieldCell' width="200">
                                               <input type="radio" value="1" name="radSubRec" class="FormElem"> Yes&nbsp;&nbsp;
                                               <input type="radio" value="0" name="radSubRec" class="FormElem" checked> No
                                           </td>
                                            <td class=FieldCellSub width="130">Enforce Ship To</td>
								            <td class='FieldCell'>
                                               <input type="radio" value="1" name="radShip" class="FormElem"> Yes&nbsp;&nbsp;
                                               <input type="radio" value="0" name="radShip" class="FormElem" checked> No
                                           </td>
                                       </tr>

                                           </table>
                                           </td>
   									</tr>
   									<tr>
   										<td class=MiddlePack>
                                           </td>
   									</tr>
   									<tr>
   										<td class=FieldCellSub>
                                           <table border="0" cellspacing="0" cellpadding="0" width="100%">

                                           <tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p></td>
															<td class="GroupTitle" width="100">
																<p align="center">
																<!--<input type="checkbox" value="Y" name="chkVendor" class="FormElem" onclick="resetVendor(this)"> Preferred Vendor-->
																Item Supplier
	                                                        </td>
															<td class="GroupTitleRight"><p align="left">&nbsp;</td>
														</tr>
													</table>
												</td>
											</tr>
                                       <tr>
                                       <td class="GroupTable">
                                       <table cellpadding="0" cellspacing="0" border="0" width="100%">
                                        <tr>
   										    <td class=MiddlePack colspan="3">
                                            </td>
   									    </tr>
                                          <!--<tr>
										    <td class="FieldCellSub">Vendor</td>
										    <td class="FieldCellSub">
										        <span id="txtParty" class="DataOnly"><%=sSuppName%>&nbsp;</span>
    										    <img src="../../assets/images/iTMS%20icons/Entryicon.gif" onclick="PopulateSupplier()">
										    </td>
										    <td class="FieldCellSub" width="5"></td>
										    <td class="FieldCellSub" width="147">Supplier Item Description</td>
										    <td class="FieldCellSub" width="94"> <input type="text" name="txtSupItmDesc" size="15" maxlength=100 class="Formelem"> </td>
                                           </tr>
                                           <tr>
										    <td class="FieldCellSub" width="147">Supplier Item Code</td>
										    <td class="FieldCellSub" width="94"> <input type="text" name="txtSupItmNo" size="15" maxlength=10 class="Formelem"> </td>
										    <td class="FieldCellSub" width="5"></td>
										    <td class="FieldCellSub" width="147">Supplier Drawing No</td>
										    <td class="FieldCellSub" width="94"> <input type="text" name="txtSuppDrawingNo" size="15" maxlength=10 class="Formelem"></td>
                                           </tr>
                                           <tr>
										    <td class="FieldCellSub" width="147">Supplier UOM</td>
										    <td class="FieldCellSub" width="94">
										        <select name="selSuppUOM" class=FormElem>
										            <%
										                PopulateUOM
										            %>
										        </select>
										    </td>
										    <td class="FieldCellSub" width="5"></td>
										    <td class="FieldCellSub" width="147">Transit Lead Time</td>
										    <td class="FieldCellSub" width="94"> <input type="text" name="txtTrLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
                                           </tr>
                                           <tr>
										    <td class="FieldCellSub" width="152">Purchase Lead Time</td>
										    <td class="FieldCellSub" width="89"> <input type="text" name="txtPuLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
										    <td class="FieldCellSub" width="5"></td>
										    <td class="FieldCellSub" width="152">Purchase Warranty Period</td>
										    <td class="FieldCellSub" width="89"> <input type="text" name="txtPurWarranty" size="4" maxlength=3 class="Formelem"> (in days)</td>

                                           </tr>
                                           <tr>
										    <td class="FieldCellSub" width="152">Supplier Lead Time</td>
										    <td class="FieldCellSub" width="89"> <input type="text" name="txtSuLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
										    <td class="FieldCellSub" width="5"></td>
										    <td class="FieldCellSub" width="147">Market Price</td>
										    <td class="FieldCellSub" width="94"> <input type="text" name="txtMarketPrice" size="12" maxlength=10 class="Formelem"> </td>
                                           </tr>
                                           <tr>
										    <td class="FieldCellSub" width="152">Preorder Lead Time</td>
										    <td class="FieldCellSub" width="89"> <input type="text" name="txtPrLTime" size="4" maxlength=3 class="Formelem"> (in days)</td>
										    <td class="FieldCellSub" width="5"></td>
										    <td class="FieldCellSub" width="147">Market Date</td>
										    <td class="FieldCellSub" width="94"> <input type="text" name="txtMarketDate" size="12" maxlength=10 class="Formelem"> </td>
                                           </tr>
                                           <tr>
										    <td class="FieldCellSub" width="152">Preferred Min. Order Qty</td>
										    <td class="FieldCellSub" width="89"> <input type="text" name="txtPreMinQty" size="12" maxlength=10 class="Formelem"> </td>
										    <td class="FieldCellSub" width="5"></td>
										    <td class="FieldCellSub" width="147">Preferred Max. Order Qty</td>
										    <td class="FieldCellSub" width="94"> <input type="text" name="txtPreMaxQty" size="12" maxlength=10 class="Formelem"> </td>
                                           </tr>-->

                                           <tr>
                                                <td class="clearpixel" style="width:5px">
                                                </td>
                                                <td>
                                                    <table class="exceltable" cellpadding="0" cellspacing="1" border="0" width="100%">
                                                        <tr>
													        <td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
													        <td class="ExcelHeaderCell" align="center" width="10" ></td>
													        <td class="ExcelHeaderCell" align="center" >Code - Name</td>
													        <td class="ExcelHeaderCell" align="center" >Item Code</td>
													        <td class="ExcelHeaderCell" align="center" >Supp. Item Code</td>
													        <td class="ExcelHeaderCell" align="center" >Item Description</td>
													        <td class="ExcelHeaderCell" align="center" >Rate</td>
												        </tr>
												            <%
												                sSql = " select V.PartyCode,V.PartyName,S.SuppItemCode,S.SuppItemDescription,S.PartyType,"&_
														               " S.PartySubType,S.Itemcode,S.ClassificationCode,M.CompanyItemCode,S.SuppMarketPrice,V.OrgnPartyCode "&_
														               " From VWOrgParty V,Inv_R_ItemSupplier S,Inv_M_ItemMaster M  "&_
														               " Where V.PartyCode = S.PartyCode and V.PartySubType = S.PartySubType "&_
														               " and V.PartyType = S.PartyType and S.Organisationcode='"& sOrgCode &"' "&_
														               " and S.ItemCode=M.Itemcode and S.ClassificationCode = M.ClassificationCode"
            														   
														                sSql = sSql & " and V.PartyType ='CR'"
            													
													            If iItmCode <> "" and iClassCode <> "" Then
														              sSql = sSql & " and S.ItemCode="& iItmCode &" and S.Classificationcode="& iClassCode &" "
													            End If
            													

													            with dcrs
														            .ActiveConnection=con
														            .CursorLocation=3
														            .CursorType=3
														            .Source=sSql
														            .Open
													            end with
													            iCnt = "0"

													            if not dcrs.EOF then
														            do while not dcrs.eof 

															            iCnt = iCnt + 1
															            nItemRate = dcrs(9)


														            %>
															            <tr>
																            <td class="ExcelSerial" align="center" ><%=iCnt%></td>
																            <td class="ExcelDisplayCell" align="center" width="10">
																	            <input type="checkbox" name="Chkbox<%=iCnt%>" value="<%=dcrs(0)%>|<%=dcrs(4)%>|<%=dcrs(5)%>">
																	            <Input type="hidden" name="hData<%=iCnt%>" value="<%=dcrs(6)%>:<%=dcrs(7)%>:<%=dcrs(8)%>:">
																            </td>
																            <td class="ExcelDisplayCell" align="Left" ><%=Trim(dcrs(10))%>-<%=Trim(dcrs(1))%></td>
																            <td class="ExcelDisplayCell" align="Center" ><%=trim(dcrs(8))%></td>
																            <td class="ExcelDisplayCell" align="Center" ><%=trim(dcrs(2))%></td>
																            <td class="ExcelDisplayCell" align="Left"><%=trim(dcrs(3))%></td>
																            <td class="ExcelDisplayCell" align="right" ><%=FormatNumber(nItemRate,2,,,0)%></td>
															            </tr>
														            <%
														                dcrs.MoveNext
														            loop
													            end if
													            dcrs.Close
												            %>
                                                    </table>
                                                    </td>
                                                    <td class="clearpixel" style="width:5px">
                                                    </td>
                                                </tr>
   									            
												<tr>
   										            <td class=MiddlePack colspan="3">
                                                    </td>
   									            </tr>
											</table>
                                           </td>
   									    </tr>
   									</table>
   									</td>
   									</tr>
   									<tr>
   										<td class=MiddlePack>
                                           </td>
   									</tr>
   								</table>
                               </div>
                                           </td>
   									</tr>
   								</table>
                                   </div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
   								 	   <div align="left">
   								 		<table cellpadding="0" cellspacing="0" width="100%">
   								 			<tr>
   								 				<td>
   								 		            <table cellpadding="0" cellspacing="0" width="100%" height="14">
   								 			            <tr>
   								 				            <td class='GroupTitleLeft' width="10" height="14">&nbsp;
                                                                </td>
   								 				            <td class='GroupTitle' width="132" height="14"><p align="center">
                                                                Additional Receiving
                                                                </td>
   								 				            <td class='GroupTitleRight' height="14"><p align="left">&nbsp;
                                                                </td>
   								 			            </tr>
   								 		            </table>
                                                </td>
   								 			</tr>
   								 			<tr>
   								 				<td class=GroupTable>
   								                 	<center>
                                            <div align="left">
                                        <table border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                <td valign="top" colspan="3" class="MiddlePack"></td>
                                    </tr>
                                    <tr>
                                <td valign="top"><div align="left">
   								 		<table cellpadding="0" cellspacing="0">
                                            <tr>
   								 				<td class=FieldCellSub width="117">Receipt Date Action</td>
   								 				<td class='FieldCell' width="215">
                                                    <input type="radio" name="radReDate" value="R" class="FormElem" > Reject&nbsp;
                                                    <input type="radio" name="radReDate" value="W" class="FormElem" > Warning&nbsp;
                                                    <input type="radio" name="radReDate" value="N" class="FormElem" checked > None
		                                        </td>
                                            </tr>
                                            <tr>
   								 				<td class=FieldCellSub width="117">Receipt Days Early (in days)</td>
   								 				<td class='FieldCellSub' width="215">
													<input type="text" name="txtRecDaysE" size="4" maxlength=3 value="0" class="Formelem" >
												</td>
                                            </tr>
                                            <tr>
   								 				<td class=FieldCellSub width="117">Receipt Days Late (in days)</td>
   								 				<td class='FieldCellSub' width="215">
													<input type="text" name="txtRecDaysL" size="4" maxlength=3 value="0" class="Formelem" >
		                                        </td>
                                            </tr>
   								 		</table>
                                </div></td>
                                <td width="5" valign="top"></td>
                                <td valign="top" align="right">
                            <div align="left">
   						<table cellpadding="0" cellspacing="0">
   							<tr>
   								<td>
   						<table cellpadding="0" cellspacing="0" width="100%" height="14">
   							<tr>
   								<td class='GroupTitleLeft' width="10" height="14">&nbsp;
                                    </td>
   								<td class='GroupTitle' width="100" height="14"><p align="center">Tolerance in %
                                    </td>
   								<td class='GroupTitleRight' height="14"><p align="left">&nbsp;
                                    </td>
   							</tr>
   						</table>
                                </td>
   							</tr>
   							<tr>
   								<td class=GroupTable>
                        <div align="left">
   						<table cellpadding="0" cellspacing="0">
   							<tr>
   								<td class=MiddlePack colspan="3">  </td>
   							</tr>
                            <tr>
   								<td class=FieldCellSub>
                                    </td>
   								<td class='FieldCellSub' align="center">Low</td>
   								<td class='FieldCellSub' align="center">High</td>
                            </tr>
                            <tr>
   								<td class=FieldCellSub>Under Receipts</td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnLow" size="3" maxlength=3 class="Formelem">
                        </td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnHigh" size="3" maxlength=3 class="Formelem">
                        </td>
                            </tr>
                            <tr>
   								<td class=FieldCellSub>Over Receipts</td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtOvLow" size="3" maxlength=3 class="Formelem">
                        </td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtOvHigh" size="3" maxlength=3 class="Formelem">
                        </td>
                            </tr>
                            <tr>
   								<td class=FieldCellSub>Unordered Receipts</td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnOrLow" size="3" maxlength=3 class="Formelem">
                        </td>
   								<td class='FieldCellSub' align="center">
								<input type="text" name="txtUnOrHigh" size="3" maxlength=3 class="Formelem">
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
                            <tr>
                        <td valign="top" colspan="3" class="MiddlePack"></td>
                            </tr>
                                </table>
                                    </div>
								</center>
                                                </td>
										</tr>
									</table>
                                    </div>
								</td>
								<td align="center">
								</td>
								</tr>
								<tr>
								    <td align="center"></td>
								    <td align="center" class="ActionCell">
									    <input type=button name="btnPurSave" value="Update" class="ActionButtonX" onclick="PurRecSubmit()">
								    </td>
								    <td align="center"></td>
                                </tr>
								</table>
								</div><!-- id=divPurRec-->
								</td>
								</tr>
								</table>
								</div><!--div-->
								</td>
								</tr>
								<!-------->
								</table>
                            </td>
                            </tr>
                            </TABLE>
                            </td>
                            </tr>
                            </table>
                            </td>
                            </tr>
								<!-------->
								                            <!----------------->
                            <tr>
                            <td align="center" width="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							<td valign="top">
								<table id="Table1" cellspacing="0" cellpadding="0" border="0" width="100%">
									<tr>
										<td >
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td align="center" colspan="3" class="MiddlePack">
														<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
													</td>
												</tr>

                            <tr>
							<td align="center" width="5" class="ClearPixel">
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							<td valign="top" width="100%">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" class=BodyTable>

                            <!----------------->
								<tr>
								<td align="center">
								<div>
									<table class="CollapseBand" cellspacing="0" cellpadding="0" >
										<tr>
											<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(divSales,'')" itms_state="0">
												<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
												</a>
											</td>
											<td valign="center" class="SubTitle">&nbsp;&nbsp;
										 	    Sales
											</td>
										</tr>
									</table>
									<table border="0" cellpadding="0" cellspacing="0" width=100% class=BodyTable>
									<tr>
										<td width="100%">
                            <div id=divSales style="width:100%;display:none;">
                            <table>

								                            <tr>
																<td align="center" width="5">
																</td>
																<td valign="top" width="100%">
								                                    <div align="left">
								                                        <table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="75%">
																		    <tr>
																				<td class="FieldCellSub" width="75">Basic UoM</td>
																				<td><span class="DataOnly"><%=sSalesUOM%>&nbsp;</span></td>
																				<td></td>
																				<td></td>
																				<td class="FieldCellSub" width="140">Optional UoM</td>
																				<td colspan="2">
																				<%	if sSalesUOM = "N/A" then %>
																				    <input type="button" value="View" name="btnUoMSal" class="AddButton" DISABLED>
																				<%	else %>
																				    <input type="button" value="View" name="btnUoMSal" class="AddButton" onClick="OpenUoM('Sal')">
																				<%	end if %>
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="140">Sales Warranty Period</td>
																				<td class="FieldCell" colspan="2">
																					<input type="text" name="txtSalWarranty" size="3" maxlength="3" class="Formelem"> (in Days)
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Minimum Sale Quantity</td>
																				<td class="FieldCell">
																					<input type="text" name="txtMinSale" size="15" maxlength="10" class="Formelem">
																				</td>
																				<td><span class="DataOnly"><%=sSalesUOM%>&nbsp;</span></td>
																			</tr>
																			<!--Added by Ragav-->
																			<%
																			    Dim sRate,nMarPer,nMarVal,nOthPer,nOthVal,nPrice,iSellNo
																			    sRate   = 0
								                                                nMarPer = 0
								                                                nMarVal = 0
								                                                nOthPer = 0
								                                                nOthVal = 0
								                                                nPrice  = 0
								                                                
								                                                sRate = GetItemRate(sOrgCode,sFinPeriod,iClassCode,iItmCode,"WA")
								                                                
								                                                sQuery = "Select SellingPriceno,Convert(Char,AsonDate,103) From Sal_M_UnitPriceHdr "
								                                                rsTemp.Open sQuery,Con
								                                                IF Not rsTemp.EOF Then
								                                                    iSellNo = rsTemp(0)
								                                                End If
								                                                rsTemp.close
								                                                
								                                                sQuery = " Select isNull(ItemRate,0),isNull(MarginPercent,0),isNull(MarginValue,0),isNull(OtherPercent,0),isNull(OtherValue,0),isNull(ItemPrice,0),convert(Varchar,EffectiveFrom,103)"&_
										                                                " From Sal_M_UnitPriceDet Where Itemcode = "&iItmCode&" "&_
										                                                " and Classificationcode = "&iClassCode&" and OudefinitionID = '"&sOrgCode&"' "
								                                                'Response.Write"<textarea>"&sQuery&"</textarea>"
								                                                rsTemp.Open sQuery,Con
								                                                IF Not rsTemp.EOF Then
								                                                    if cdbl(sRate)=0 then
									                                                    sRate = rsTemp(0)
									                                                end if
									                                                nMarPer=rsTemp(1)
									                                                nMarVal=rsTemp(2)
									                                                nOthPer=rsTemp(3)
									                                                nOthVal=rsTemp(4)
									                                                nPrice=rsTemp(5)
									                                                sEffFrom	=rsTemp(6)
								                                                End IF
								                                                rsTemp.Close
																			
																			%>
																			<tr>
																			    <td class="FieldCellSub" width="140">Sales Rate</td>
																				<td class="FieldCell" colspan="2">
																					<input type="text" name="txtPurRate" size="15" maxlength="15" class="Formelem" value="<%=FormatNumber(sRate)%>" onBlur="AssaignValue()">
																				</td>
																				<td class="FieldCellSub">Per</td>
																				<td class="FieldCellsub" colspan="2">
																					<input type="text" name="txtPurRatePer" size="3" maxlength="3" class="Formelem" value="1">&nbsp;<span class="DataOnly"><%=sSalesUOM%></span>
																				</td>
																				<td></td>
																				
																				<td class="FieldCellSub"></td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="140">Charges %</td>
																				<td class="FieldCell" colspan="2">
																					<input type="text" name="txtCharPer" size="3" maxlength="3" class="Formelem" value="<%=nOthPer%>" onblur="CalcValue('OP')">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Charges Value</td>
																				<td class="FieldCell">
																					<input type="text" name="txtCharValue" size="15" maxlength="15" class="Formelem" value="<%=FormatNumber(nOthVal,2,,,0)%>"  onblur="CalcValue('OV')">
																				</td>
																				<td></td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="140">Margin %</td>
																				<td class="FieldCell" colspan="2">
																					<input type="text" name="txtMarPer" size="3" maxlength="3" class="Formelem" value="<%=nMarPer%>" onblur="CalcValue('MP')">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Margin Value</td>
																				<td class="FieldCell">
																					<input type="text" name="txtMarValue" size="15" maxlength="15" class="Formelem" value="<%=FormatNumber(nMarVal,2,,,0)%>"  onblur="CalcValue('MV')">
																				</td>
																				<td></td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="140">Selling Price</td>
																				<td class="FieldCell" colspan="2">
																					<input type="text" name="txtTotalPrice" size="15" maxlength="15" class="Formelem" value="<%=FormatNumber(nPrice,2,,,0)%>">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Effective from</td>
																				<td class="FieldCell" colspan="2">
																					<input type="text" id="ctlEffFrom" name="ctlEffFrom" class="formelem itms-date-picker" data-itms-datepicker="1" size="10">
																				</td>
																				<td class="FieldCellSub"></td>
																			</tr>
																			<!--Added above content and blocked the blow content on Apr 24,2013-->
																			<!--<tr>
																				<td class="FieldCellSub" width="75">Market Rate</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtMarketRate" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" colspan="3" align="center"><p align="left">Selling Price:</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="75">Actual</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtActual" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Unit Size</td>
																				<td class="FieldCell">
																					<input type="text" name="txtUnitSize" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCell">
																					<select size="1" name="selUoMUnit" class="FormElem">
																						<option value="select">Select</option>
																						<%	'Calling the Function which populates the UoM list
																							populateUoM
																						%>
								                                                    </select>
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="75">Minimum</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtMin" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																				<td class="FieldCellSub" width="140">Volume</td>
																				<td class="FieldCell">
																					<input type="text" name="txtVolume" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCell">
																					<select size="1" name="selUoMVolume" class="FormElem">
																						<option value="select">Select</option>
																						<%	'Calling the Function which populates the UoM list
																							populateUoM
																						%>
								                                                    </select>
																				</td>
																			</tr>
																			<tr>
																				<td class="FieldCellSub" width="75">Preferred</td>
																				<td><span class="DataOnly">Rs&nbsp;</span></td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtPreffered" size="15" maxlength=10 class="Formelem">
																				</td>
																				<td class="FieldCellSub"></td>
																			</tr>-->
																			<!--End-->
								                                        </table>
								                                    </div>
																</td>
																<td align="center">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>

								                            <tr>
																<td align="center" width="5">
																</td>
																<td valign="top" width="100%">
															        <div align="left">
															        <table border="0" cellspacing="0" cellpadding="0">
																	    <tr>
																			<td class="FieldCell" width="105" valign="top">Commodity</td>
																			<td class="FieldCellSub">
																		    <select size="1" name="selCommodity" class="FormElem">
																			<option value="select">Select</option>
																			<%	'Calling the Function which populates the Commodity list
																				'populateCommodity(iCommodity)
																			%>
																			</select>
																			</td>
																	    </tr>
															        </table>
															        </div>
																</td>
																<td align="center">
																</td>
								                            </tr>

								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" width="5"></td>
																<td valign="top" width="100%">
								                                    <table cellpadding="0" cellspacing="0">
																		<tr>
																			<td>
																				<table cellpadding="0" cellspacing="0" width="100%">
																					<tr>
																						<td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p></td>
																						<td class="GroupTitle" width="120">
																							<p align="center"><input type="checkbox" value="Y" name="chkDiscount" class="FormElem" onclick="resetDiscount(this)"> Sales Discount
								                                                        </td>
																						<td class="GroupTitleRight"><p align="left">&nbsp;</td>
																					</tr>
																				</table>
																			</td>
																		</tr>
																		<tr>
																			<td class=GroupTable>
																				<div align="left">
																					        <table cellpadding="0" cellspacing="0" width="100%">
																						        <tr>
																							        <td class=MiddlePack colspan="2"> <p align="left"> </td>
																						        </tr>
																						        <tr>
																							        <td class=FieldCellSub colspan="2"> <p align="left">&nbsp;Which Discount should take precendence <input type=radio name=radQV value=Q checked>Quantity&nbsp;&nbsp;<input type=radio name=radQV value=V>Value&nbsp; </td>
																						        </tr>
																						        <tr>
																							        <td class=FieldCellSub colspan="2"> <p align="left">&nbsp;Discount applicable in <input type=radio name=radApplicable value=B checked>Basic Value&nbsp;&nbsp;<input type=radio name=radApplicable value=T>Total Value&nbsp;&nbsp;<input type=radio name=radApplicable value=P>Purchase Value</td>
																						        </tr>

																						        <tr>
																							        <td class=FieldCellSub>
																								        <div align="left">
        																								<table border=0 width=100%>
        																								<tr>
        																								<td valign=top>
																								        <table border="0" cellspacing="1" id="tblDataQty" class="ExcelTable" width=100%>
																				                            <tr>
																						                        <td class="ExcelHeaderCell" align="center" width="30" rowspan="2">S.No.</td>
																						                        <td class="ExcelHeaderCell" align="center" colspan="2">Quantity</td>
																						                        <td class="ExcelHeaderCell" align="center" width="75" rowspan="2">Discount %</td>
																						                        <td class="ExcelHeaderCell" align="center" width="75" rowspan="2">UoM</td>
																						                        <td class="ExcelHeaderCell" align="center" rowspan="2"></td>
																				                            </tr>
																				                            <tr>
																						                        <td class="ExcelHeaderCell" align="center" width="75">From</td>
																						                        <td class="ExcelHeaderCell" align="center" width="75">To</td>
																				                            </tr>
																				                            <tr>
																				                                <td class="ExcelSerial">
																				                                <td class="ExcelDisplayCell">
																							                        <input type="text" name="txtQtyFrom" size="11" maxlength=10 class="Formelem">
																					                            </td>
																					                            <td class="ExcelDisplayCell">
																							                        <input type="text" name="txtQtyTo" size="11" maxlength=10 class="Formelem">
																						                        </td>
																						                        <td class="ExcelDisplayCell">
																							                        <input type="text" name="txtQtyDis" size="4" maxlength=3 class="Formelem">
																						                        </td>
																						                        <td class="ExcelDisplayCell">
																							                        <select size="1" name="selUoMQty" class="FormElem">
																								                        <!--<option value="select">Select</option>-->
																								                        <%	'Calling the Function which populates the UoM list
																									                        populateUOMForItem iItmCode,iClassCode,sOrgCode
																								                        %>
																							                        </select>
																						                        </td>
																						                        
																						                        <td class="ExcelDisplayCell" >
																			                                      <p align="center"><input type="button" value=" Add " name="B5" class="AddButtonX" onClick="CheckEntryDis()">
																						                        </td>
																				                            </tr>
																				                        </table></td><td valign=top>
																				                        <table border="0" cellspacing="1" id="tblDataVal" class="ExcelTable" width=100%>
																				                        <tr>
																						                    <td class="ExcelHeaderCell" align="center" width="30" rowspan="2">S.No.</td>
																						                    <td class="ExcelHeaderCell" align="center" colspan="2">Value</td>
																						                    <td class="ExcelHeaderCell" align="center" width="75" rowspan="2">Discount %</td>
																						                    <td class="ExcelHeaderCell" align="center" rowspan="2"></td>
																				                        </tr>
																				                        <tr>
																						                    <td class="ExcelHeaderCell" align="center" width="75">From</td>
																						                    <td class="ExcelHeaderCell" align="center" width="75">To</td>
																				                        </tr>
																				                        <tr>
																				                            <td class=ExcelSerial></td>
																				                            <td class="ExcelDisplayCell">
																							                    <input type="text" name="txtValFrom" size="11" maxlength=10 class="Formelem">
																						                    </td>
																						                    <td class="ExcelDisplayCell">
																							                    <input type="text" name="txtValTo" size="11" maxlength=10 class="Formelem">
																						                    </td>
																						                    <td class="ExcelDisplayCell">
																							                    <input type="text" name="txtValDis" size="4" maxlength=3 class="Formelem">
																						                    </td>
																						                    <td class="ExcelDisplayCell">
																			                                  <p align="center"><input type="button" value=" Add " name="B6" class="AddButtonX" onClick="CheckEntryVal()">
																						                    </td>
																				                        </tr>
																				                    </table>
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
																<td align="center"></td>
								                            </tr>
								                            <tr>
																<td align="center" colspan="3" class="MiddlePack">
																	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																</td>
								                            </tr>
								                            <tr>
																<td align="center" width="5">
																</td>
																<td valign="top" width="100%">
								                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
																	    <tr>
																			<td>
																			<!--	<DIV class=frmBody id=frm3 style="width: 290; height:75;">-->

																			<!--	</div>-->
																			</td>
																			<td>
																				<!--<DIV class=frmBody id=frm31 style="width: 290; height:75;">-->

																				<!--</div>-->
																			</td>
																	    </tr>
								                                    </table>
																</td>
																<td align="center">
																</td>
								                            </tr>
                                                        <tr>
							                                <td align="center"></td>
							                                <td align="center" class="ActionCell">
								                                <input type=button name="btnSalSave" value="Update" class="ActionButtonX" onclick="SalSubmit()">
							                                </td>
							                                <td align="center"></td>
                                                        </tr>

                            </table>
                            </div><!-- id=divSales-->
                            </td>
                            </tr>
                            </table>
                            </div><!--div-->
                            </td>
                            </tr>
                            								<!-------->
								</table>
                            </td>
                            </tr>
                            </TABLE>
                            </td>
                            </tr>
                            </table>
                            </td>
                            </tr>
								<!-------->

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
								<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Back" name="B5" class="ActionButton" onClick="window.location.href='ItmCreationDefinitionEntry.asp?Flag=O&iItmCode=<%=iItmCode%>'">
                                                    <input type="button" value="Cancel" name="B1" class="ActionButton" onclick="window.location.href='ITEMLISTENTRY.ASP?ACTN=L'">
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
    Function populateUOMForItem(ItemCode,ClassCode,OrgCode)
        sQuery = "Select SalesUOM from INV_M_ITEMMASTER where ItemCode="& ItemCode &" and ClassificationCode = "& ClassCode &" and organisationcode = "& OrgCode
        rsTemp.Open sQuery,con
        if not rsTemp.EOF then
            Response.Write "<option value="& Trim(rsTemp(0)) &">"& Trim(rsTemp(0)) &"</option>"
        end if
        rsTemp.Close
    End Function
%>
<%
    Function GetItemName(ItemCode,ClassCode)
        Dim sQuery,rsItem
        set rsItem = Server.CreateObject("ADODB.Recordset")
        sQuery = "Select ItemDescription from VWITEM where ItemCode ="& ItemCode &" and ClassificationCode = "& ClassCode
        rsItem.Open sQuery,con
        if not rsItem.EOF then
            GetItemName = trim(rsItem(0))
        end if
        rsItem.Close
    End Function
%>
<%
    Function PopulateTaxType()
        Dim sQuery,rsTax
        set rsTax = Server.CreateObject("ADODB.Recordset")
        sQuery = "Select InvoiceType,InvoiceTypeName from Sal_M_InvoiceTypes Order By InvoiceType"
        rsTax.Open sQuery,con
        if not rsTax.eof then
            do while not rsTax.eof
                response.write "<option value="& trim(rsTax(0))&">"&trim(rsTax(1))&"</option>"
                rsTax.MoveNext
            loop
        end if
        rsTax.close
    End function
%>

