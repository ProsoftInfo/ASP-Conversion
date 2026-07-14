<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmDetailedPurRecInsert.asp
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
    Dim ndRoot,ndPur,ndChild,ndEntry
    Dim sQuery,sExp,iCnt,sCreatedBy
 
    Dim sBuyer,sWarrenty,sTransitLeadTime,sPurLeadTime,sSuppItemNo,sSuppLeadTime
    Dim sMarketPrice,sPreOrdLeadTime,sMarketDate,sPreMinOrdQty,sPreMaxOrdQty
    Dim sSubReceipts,sEnforceShipTo,sRecDateAction,sRecDaysEarly,sRecDaysLate
    Dim sUnRecLow,sUnRecHigh,sOverRecLow,sOverRecHigh,sUnOrdRecLow,sUnOrdRecHigh
    Dim sSuppCode,sSuppName,sSuppSubType,sSuppType,sSuppItemDesc,sItemCode,sClassCode,sOrgCode
    Dim sItemName,sClassName,sOrgName,sUnder,sOver,sUnOrder,sModVat,sInvMatch,sSubCont
    Dim sSuppDrawingNo,sSuppUOM
    
    Dim sUCode,sBRate,sOperator,sUName,sOperatorText
    Dim sAltItemCode,sAltClassCode,sAltPeriority,sAltItemName
    con.begintrans
    
    sCreatedBy = getuserId()
 
     set oDOM = Server.CreateObject("Microsoft.XMLDOM")
     set rsHead = Server.CreateObject("ADODB.Recordset")
     set rsTemp = Server.CreateObject("ADODB.Recordset")
     set objFSO = Server.CreateObject("Scripting.FileSystemObject")
     if objFSO.FileExists(Server.MapPath("../temp/Master/PurchaseDet"&Session.SessionID&".xml")) then
        oDOM.load Server.MapPath("../temp/Master/PurchaseDet"&Session.SessionID&".xml")
            set ndRoot = oDOM.documentElement
            if ndRoot.hasChildNodes() then
                sItemCode = ndRoot.getAttribute("ItemCode")
                sClassCode = ndRoot.getAttribute("ClassCode")
                sOrgCode = ndRoot.getAttribute("OrgCode")
                sItemName = ndRoot.getAttribute("ItemName")
                sClassName = ndRoot.getAttribute("ClassName")
                sOrgName = ndRoot.getAttribute("OrgName")
            end if
            
            sExp = "//Alternate/Entry"
            set ndEntry = ndRoot.selectNodes(sExp)
            if ndEntry.length>0 then
                For iCnt = 0 to ndEntry.length-1
                    sAltItemCode = ndEntry.Item(iCnt).Attributes.getNamedItem("ITEMCODE").value
                    sAltClassCode = ndEntry.Item(iCnt).Attributes.getNamedItem("CLASSCODE").value
                    sAltPeriority = ndEntry.Item(iCnt).Attributes.getNamedItem("PRIORITY").value
                    sAltItemName = ndEntry.Item(iCnt).Attributes.getNamedItem("ITEMNAME").value
                    
                    sQuery = "Insert into INV_M_ItemOrgAlternate (ItemCode,ClassificationCode,OrganisationCode,"&_
                             " AlternateItemCode,AlternateClassification,Prority) values("& sItemCode &","& sClassCode &","&_
                             ""& Pack(sOrgCode) &","& sAltItemCode &","& sAltClassCode &","& sAltPeriority &")"
                    Response.Write "<p>"& sQuery
                    con.execute sQuery
                Next
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
                             "("& sItemCode &","& sClassCode &","& Pack(sOrgCode) &",'P',"& Pack(sUCode) &","& sBRate &","& sOperator &")"
                    Response.Write "<p>"& sQuery
                    con.execute sQuery
                Next
            end if
            
            sExp = "//Vendor"
            set ndEntry = ndRoot.selectNodes(sExp)
            if ndEntry.length>0 then
                sWarrenty =ndEntry.Item(0).Attributes.getNamedItem("Warrenty").value
                sTransitLeadTime = ndEntry.Item(0).Attributes.getNamedItem("TransitLeadTime").value
                sPurLeadTime = ndEntry.Item(0).Attributes.getNamedItem("PurLeadTime").value
                sSuppItemNo = ndEntry.Item(0).Attributes.getNamedItem("SuppItemNo").value
                sSuppLeadTime = ndEntry.Item(0).Attributes.getNamedItem("SuppLeadTime").value
                sMarketPrice =ndEntry.Item(0).Attributes.getNamedItem("MarketPrice").value
                sPreOrdLeadTime= ndEntry.Item(0).Attributes.getNamedItem("PreOrdLeadTime").value
                sMarketDate = ndEntry.Item(0).Attributes.getNamedItem("MarketDate").value
                sPreMinOrdQty = ndEntry.Item(0).Attributes.getNamedItem("PreMinOrdQty").value
                sPreMaxOrdQty = ndEntry.Item(0).Attributes.getNamedItem("PreMaxOrdQty").value
                sSuppCode = ndEntry.Item(0).Attributes.getNamedItem("SuppCode").value
                sSuppName = ndEntry.Item(0).Attributes.getNamedItem("SuppName").value
                sSuppType = ndEntry.Item(0).Attributes.getNamedItem("SuppType").value
                sSuppSubType = ndEntry.Item(0).Attributes.getNamedItem("SuppSubType").value
                sSuppItemDesc = ndEntry.Item(0).Attributes.getNamedItem("SuppItemDesc").value
                sSuppDrawingNo =ndEntry.Item(0).Attributes.getNamedItem("SuppDrawingNo").value
                sSuppUOM = ndEntry.Item(0).Attributes.getNamedItem("SuppUOM").value
                
                sQuery = "Insert into Inv_R_ItemSupplier (OrganisationCode,ClassificationCode,ItemCode,PartyType,"&_
                         " PartySubType,PartyCode,SuppItemCode,SuppItemDescription,SupplierDrawingNo,SupplierUOM,"&_
                         " MinOrderQuantity,MaxOrderQuantity,PreOrderLeadTime,SuppLeadTime,SuppTransitTime,"&_
                         " PurchaseLeadTime,SuppWarrantyPeriod,SuppMarketPrice,MarketDate) values("& Pack(sOrgCode) &","&_
                         " "& sClassCode  &","& sItemCode &",'"& sSuppType &"',"& sSuppSubType&","& sSuppCode&","&_
                         " "& sSuppItemNo&","& Pack(sSuppItemDesc)&","& Pack(sSuppDrawingNo) &","& Pack(sSuppUOM)&","&_
                         " "& sPreMinOrdQty &","& sPreMaxOrdQty &","& sPreOrdLeadTime &","& sSuppLeadTime &","& sTransitLeadTime &","&_
                         " "& sPurLeadTime &","& sWarrenty &","& sMarketPrice&",Convert(datetime,"&Pack(sMarketDate)&",103))"
                Response.Write "<p>"& sQuery
                con.execute sQuery
            end if
            
            sExp = "//Basic"
            set ndEntry = ndRoot.selectNodes(sExp)
            if ndEntry.length>0 then
                sBuyer = ndEntry.Item(0).Attributes.getNamedItem("Buyer").value
                'sModVat = ndEntry.Item(0).Attributes.getNamedItem("ModVat").value
                sInvMatch = ndEntry.Item(0).Attributes.getNamedItem("InvMatch").value
                sSubCont = ndEntry.Item(0).Attributes.getNamedItem("SubCont").value
                sSubReceipts = ndEntry.Item(0).Attributes.getNamedItem("SubReceipts").value
                sEnforceShipTo = ndEntry.Item(0).Attributes.getNamedItem("EnforceShipTo").value
                sRecDateAction = ndEntry.Item(0).Attributes.getNamedItem("RecDateAction").value
                sRecDaysEarly =ndEntry.Item(0).Attributes.getNamedItem("RecDaysEarly").value
                sRecDaysLate = ndEntry.Item(0).Attributes.getNamedItem("RecDaysLate").value
                sUnRecLow =  ndEntry.Item(0).Attributes.getNamedItem("UnRecLow").value
                sUnRecHigh = ndEntry.Item(0).Attributes.getNamedItem("UnRecHigh").value
                sOverRecLow =  ndEntry.Item(0).Attributes.getNamedItem("OverRecLow").value
                sOverRecHigh =  ndEntry.Item(0).Attributes.getNamedItem("OverRecHigh").value
                sUnOrdRecLow = ndEntry.Item(0).Attributes.getNamedItem("UnOrdRecLow").value
                sUnOrdRecHigh = ndEntry.Item(0).Attributes.getNamedItem("UnOrdRecHigh").value
                
                sOver = "1"
                sUnder = "1"
                sUnOrder = "1"
                if sUnRecHigh = "" then  sUnRecHigh = "0"
                if sUnRecLow ="" then sUnRecLow = "0"
                if sOverRecHigh="" then sOverRecHigh ="0"
                if sOverRecLow="" then sOverRecLow = "0"
                if sUnOrdRecLow ="" then sUnOrdRecLow="0"
                if sUnOrdRecHigh="" then sUnOrdRecHigh ="0"
                
                
                sQuery = "Insert into INV_M_ItemOrgPurchase (ItemCode,ClassificationCode,OrganisationCode,"&_
                         " AllowUnderReceipts,UnderRcptLowLimit,UnderRcptHighLimit,AllowOverReceipts,OverRcptLowLimit,"&_
                         " OverRcptHighLimit,SubContractEligiility,InvoiceMatching,"&_
                         " AllowSubstitutes,AllowUnorderedRcpt,UnOrdRcptLowLimit,UnOrdRcptHighLimit,PreferredBuyer,"&_
                         " AllowShipTo,ItemDefinedBy,ItemDefinedOn) values"&_
                         "("& sItemCode &","& sClassCode &","& Pack(sOrgCode) &","& sUnder &","& sUnRecLow &","& sUnRecHigh &","&_
                         ""&sOver &","& sOverRecLow &","& sOverRecHigh &","& sSubCont &","& sInvMatch &","&_
                         " "& sSubReceipts &","& sUnOrder &","& sUnOrdRecLow &","& sUnOrdRecHigh &","& sBuyer &","&_
                         " "& sEnforceShipTo &","& sCreatedBy &",Convert(datetime,getDate(),103))" 
                         
               Response.Write "<p>"& sQuery
               con.execute sQuery
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
