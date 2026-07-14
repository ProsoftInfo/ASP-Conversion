<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmDetailedSalesInsert.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran R
	'Created On					:	Jul 13,2011
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

<%
    Dim oDOM,rsHead,rsTemp,objFSO
    Dim ndRoot,ndSal,ndChild,ndEntry
    Dim sItemName,sClassName,sOrgName,sItemCode,sClassCode,sOrgCode,sExp,sQuery
    Dim sCreatedBy,iCnt
    
    Dim sBasicUOM,sMarketRate,sWarrentyPeriod,sSellingPrice,sMinSalQty,sActual
    Dim sUnitSize,sUnitSizeUOM,sMinimum,sVolume,sVolumeUOM,sPreferred,sCommodity
    
    Dim sQtyFrom,sQtyTo,sQtyDis,sQtyUoM,sValFrom,sValTo,sValDis,sQtyVal,sApplicableIn
    Dim sUCode,sBRate,sOperator,sUName,sOperatorText
    
    sCreatedBy = getUserID()
    
    con.begintrans
 
     set oDOM = Server.CreateObject("Microsoft.XMLDOM")
     set rsHead = Server.CreateObject("ADODB.Recordset")
     set rsTemp = Server.CreateObject("ADODB.Recordset")
     set objFSO = Server.CreateObject("Scripting.FileSystemObject")
     if objFSO.FileExists(Server.MapPath("../temp/Master/SalesDet"&Session.SessionID&".xml")) then
        oDOM.load Server.MapPath("../temp/Master/SalesDet"&Session.SessionID&".xml")
            set ndRoot = oDOM.documentElement
            if ndRoot.hasChildNodes() then
                sItemCode = ndRoot.getAttribute("ItemCode")
                sClassCode = ndRoot.getAttribute("ClassCode")
                sOrgCode = ndRoot.getAttribute("OrgCode")
                sItemName = ndRoot.getAttribute("ItemName")
                sClassName = ndRoot.getAttribute("ClassName")
                sOrgName = ndRoot.getAttribute("OrgName")
            end if
            
            sExp = "//Sales/Basic"
            set ndChild = ndRoot.selectNodes(sExp)
            if ndChild.length>0 then
                sMarketRate = ndChild.Item(0).Attributes.getNamedItem("MarketRate").value
                sWarrentyPeriod = ndChild.Item(0).Attributes.getNamedItem("WarrPeriod").value
                sMinSalQty = ndChild.Item(0).Attributes.getNamedItem("MinSalQty").value
                sActual = ndChild.Item(0).Attributes.getNamedItem("Actual").value
                sUnitSize = ndChild.Item(0).Attributes.getNamedItem("UnitSize").value
                sUnitSizeUOM = ndChild.Item(0).Attributes.getNamedItem("UnitUOM").value
                sMinimum = ndChild.Item(0).Attributes.getNamedItem("Minimum").value
                sVolume = ndChild.Item(0).Attributes.getNamedItem("Volume").value
                sVolumeUOM = ndChild.Item(0).Attributes.getNamedItem("VolumeUOM").value
                sPreferred = ndChild.Item(0).Attributes.getNamedItem("Preferred").value
                sCommodity = ndChild.Item(0).Attributes.getNamedItem("Commodity").value
                if sCommodity = "select" then sCommodity = "NULL"
                
                sQuery = "Insert into INV_M_ItemOrgSales (ItemCode,ClassificationCode,OrganisationCode,"&_
                " PreferredSellingRate,MinimumSellingRate,ActualSellingRate,ExistingMarketRate,"&_
                " WarrantyPeriod,PreferredMinQty,UnitSize,UnitVolume,UnitSizeUoM,UnitVolumeUoM,"&_
                " ItemDefinedBy,ItemDefinedOn) values ("& sItemCode  &","& sClassCode &","& Pack(sOrgCode) &","&_
                " "& sPreferred &","& sMinimum &","& sActual &","& sMarketRate &","&_
                " "& sWarrentyPeriod &","& sMinSalQty&","& sUnitSize &","& sVolume &","& Pack(sUnitSizeUOM) &","& Pack(sVolumeUOM) &","&_
                " "& sCreatedBy &",Convert(datetime,getDate(),103))"
                Response.Write "<p>"& sQuery
                con.execute sQuery
                
            end if
            
            sExp = "//DisEntry"
            set ndChild = ndRoot.selectNodes(sExp)
            if ndChild.length>0 then
                For iCnt = 0 to ndChild.length-1
                    sQtyFrom = ndChild.Item(iCnt).Attributes.getNamedItem("QTYFROM").Value
                    sQtyTo = ndChild.Item(iCnt).Attributes.getNamedItem("QTYTO").Value
                    sQtyDis = ndChild.Item(iCnt).Attributes.getNamedItem("QTYDIS").Value
                    sQtyUoM = ndChild.Item(iCnt).Attributes.getNamedItem("QTYUOM").Value
                    sQtyVal = ndChild.Item(iCnt).Attributes.getNamedItem("QTYVAL").Value
                    sApplicableIn =ndChild.Item(iCnt).Attributes.getNamedItem("APPIN").Value
                
                    sQuery = "Insert into INV_M_ItemOrgSaleDiscount (ItemCode,ClassificationCode,OrganisationCode, "&_
                             " QtyDiscountOffered,QuantityFrom,QuantityTo,UoM,Precedence,DiscApplicableOn)"&_
                             " values("& sItemCode &","& sClassCode &","& Pack(sOrgCode) &","& sQtyDis &","& sQtyFrom &","& sQtyTo &","& Pack(sQtyUoM) &","& Pack(sQtyVal) &","& Pack(sApplicableIn) &")"
                    Response.Write "<p>"& sQuery
                    con.execute sQuery
                Next 'For iCnt = 0 to ndChild.length-1
            end if
            
            
            sExp = "//ValEntry"
            set ndChild = ndRoot.selectNodes(sExp)
            
            if ndChild.length>0 then
                For iCnt = 0 to ndChild.length-1
                    sValFrom = ndChild.Item(iCnt).Attributes.getNamedItem("VALFROM").Value
                    sValTo = ndChild.Item(iCnt).Attributes.getNamedItem("VALTO").Value
                    sValDis = ndChild.Item(iCnt).Attributes.getNamedItem("VALDIS").Value
                    sQtyVal = ndChild.Item(iCnt).Attributes.getNamedItem("QTYVAL").Value
                    sApplicableIn =ndChild.Item(iCnt).Attributes.getNamedItem("APPIN").Value
                    
                    sQuery = "Insert into INV_M_ItemOrgSaleDiscount (ItemCode,ClassificationCode,OrganisationCode, "&_
                             " ValueDiscountOffered,ValueFrom,ValueTo,Precedence,DiscApplicableOn)"&_
                             " values("& sItemCode &","& sClassCode &","& Pack(sOrgCode) &","& sValDis &","& sValFrom &","& sValTo &","& Pack(sQtyVal) &","& Pack(sApplicableIn) &")"
                    Response.Write "<p>"& sQuery
                    con.execute sQuery
                Next 'For iCnt = 0 to ndChild.length-1
            end if
            
            sExp = "//OptionalUOM/OpUoMEntry"
            set ndEntry = ndRoot.selectNodes(sExp)
            if ndEntry.length>0 then
                For iCnt = 0 to ndEntry.length-1
                    sUCode = ndEntry.Item(iCnt).Attributes.getNamedItem("UCODE").value
                    sBRate = ndEntry.Item(iCnt).Attributes.getNamedItem("BRATE").value
                    sOperator = ndEntry.Item(iCnt).Attributes.getNamedItem("OPERATOR").value
                    sUName  = ndEntry.Item(iCnt).Attributes.getNamedItem("UNAME").value
                    sOperatorText = ndEntry.Item(iCnt).Attributes.getNamedItem("OPERATORTEXT").value
                    
                    sQuery = "Insert into INV_M_ItemOptionalUOM (ItemCode,ClassificationCode,OrganisationCode,"&_
                             " OptionalUoMFor,UoMCode,OptionToBaseRate,OptionToBaseOperator) values"&_
                             "("& sItemCode &","& sClassCode &","& Pack(sOrgCode) &",'S',"& Pack(sUCode) &","& sBRate &","& sOperator &")"
                    Response.Write "<p>"& sQuery
                    con.execute sQuery
                Next
            end if
            
     end if'if objFSO.FileExists(Server.MapPath("../temp/Master/PurchaseDet"&Session.SessionID&".xml")) then
    if con.Errors.count <> 0 then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.End 
		Response.Clear 
		con.CommitTrans
	end if

	con.close
	set con = nothing
%>
