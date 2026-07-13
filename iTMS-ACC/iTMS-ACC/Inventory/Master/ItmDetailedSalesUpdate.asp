<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmDetailedSalesUpdate.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran R
	'Created On					:	Jul 20,2011
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->

<%
    Dim oDOM,rsHead,rsTemp,objFSO
    Dim ndRoot,ndSal,ndChild,ndEntry
    Dim sItemName,sClassName,sOrgName,sItemCode,sClassCode,sOrgCode,sExp,sQuery
    Dim sCreatedBy,iCnt
    
    Dim sBasicUOM,sMarketRate,sWarrentyPeriod,sSellingPrice,sMinSalQty,sActual
    Dim sUnitSize,sUnitSizeUOM,sMinimum,sVolume,sVolumeUOM,sPreferred,sCommodity,sTaxType,sTaxTypeOverride
    
    Dim sQtyFrom,sQtyTo,sQtyDis,sQtyUoM,sValFrom,sValTo,sValDis,sQtyVal,sApplicableIn
    Dim sUCode,sBRate,sOperator,sUName,sOperatorText
    
    Dim objrs,rs
    Dim sPurRate,sPurRatePer,sCharPer,sCharValue,sMarPer,sMarValue,sTotPrice
    Dim sHisNo,sQry2
    Dim iSellno,sRate,sAsonDate,sEffDate,sRecno
    
    
    
    sCreatedBy = getUserID()
    
    con.begintrans
 
     set oDOM = Server.CreateObject("Microsoft.XMLDOM")
     set rsHead = Server.CreateObject("ADODB.Recordset")
     set rsTemp = Server.CreateObject("ADODB.Recordset")
     set objrs = server.createObject("ADODB.Recordset")
     set rs = server.createObject("ADODB.Recordset")
     set objFSO = Server.CreateObject("Scripting.FileSystemObject")
     if objFSO.FileExists(Server.MapPath("../temp/Master/SalesDetUpdate"&Session.SessionID&".xml")) then
        oDOM.load Server.MapPath("../temp/Master/SalesDetUpdate"&Session.SessionID&".xml")
            set ndRoot = oDOM.documentElement
            if ndRoot.hasChildNodes() then
                sItemCode = ndRoot.getAttribute("ItemCode")
                sClassCode = ndRoot.getAttribute("ClassCode")
                sOrgCode = ndRoot.getAttribute("OrgCode")
                sItemName = ndRoot.getAttribute("ItemName")
                sClassName = ndRoot.getAttribute("ClassName")
                sOrgName = ndRoot.getAttribute("OrgName")
            end if
            
            sQuery = " Delete from INV_M_ItemOrgSales where ItemCode = "& sItemCode &" and ClassificationCode = "& sClassCode &" and OrganisationCode = "& sOrgCode 
            Response.Write "<p>"& sQuery
            con.execute sQuery
            
            sQuery = " Delete from INV_M_ItemOrgSaleDiscount where ItemCode = "& sItemCode &" and ClassificationCode = "& sClassCode &" and OrganisationCode = "& sOrgCode 
            Response.Write "<p>"& sQuery
            con.execute sQuery
            
            sQuery = " Delete from INV_M_ItemOptionalUOM where ItemCode = "& sItemCode &" and ClassificationCode = "& sClassCode &" and OrganisationCode = "& sOrgCode &" and OptionalUoMFor ='S'"
            Response.Write "<p>"& sQuery
            con.execute sQuery
            
            sExp = "//Sales/Basic"
            set ndChild = ndRoot.selectNodes(sExp)
            if ndChild.length>0 then
                'sMarketRate = ndChild.Item(0).Attributes.getNamedItem("MarketRate").value
                sWarrentyPeriod = ndChild.Item(0).Attributes.getNamedItem("WarrPeriod").value
                sMinSalQty = ndChild.Item(0).Attributes.getNamedItem("MinSalQty").value
               ' sActual = ndChild.Item(0).Attributes.getNamedItem("Actual").value
               ' sUnitSize = ndChild.Item(0).Attributes.getNamedItem("UnitSize").value
               ' sUnitSizeUOM = ndChild.Item(0).Attributes.getNamedItem("UnitUOM").value
               ' sMinimum = ndChild.Item(0).Attributes.getNamedItem("Minimum").value
               ' sVolume = ndChild.Item(0).Attributes.getNamedItem("Volume").value
               ' sVolumeUOM = ndChild.Item(0).Attributes.getNamedItem("VolumeUOM").value
               ' sPreferred = ndChild.Item(0).Attributes.getNamedItem("Preferred").value
               ' sCommodity = ndChild.Item(0).Attributes.getNamedItem("Commodity").value
                
                if sCommodity = "select" then sCommodity = "NULL"
                
                sPurRate = ndChild.Item(0).Attributes.getNamedItem("PurRate").value
                sPurRatePer = ndChild.Item(0).Attributes.getNamedItem("PurRatePer").value
                sCharPer = ndChild.Item(0).Attributes.getNamedItem("CharPer").value
                sCharValue = ndChild.Item(0).Attributes.getNamedItem("CharValue").value
                sMarPer = ndChild.Item(0).Attributes.getNamedItem("MarPer").value
                sMarValue = ndChild.Item(0).Attributes.getNamedItem("MarValue").value
                sTotPrice = ndChild.Item(0).Attributes.getNamedItem("TotPrice").value
                sEffDate = ndChild.Item(0).Attributes.getNamedItem("EffectiveFrom").value
                sAsonDate = sEffDate
                
			    If sMarPer = "" Then sMarPer = 0
                If sMarValue = "" Then sMarValue = 0
                If sCharPer = "" Then sCharPer = 0 
                If sCharValue = "" Then sCharValue = 0

                
              '  sQuery = "Insert into INV_M_ItemOrgSales (ItemCode,ClassificationCode,OrganisationCode,"&_
              '  " PreferredSellingRate,MinimumSellingRate,ActualSellingRate,ExistingMarketRate,"&_
              '  " WarrantyPeriod,PreferredMinQty,UnitSize,UnitVolume,UnitSizeUoM,UnitVolumeUoM,"&_
              '  " ItemDefinedBy,ItemDefinedOn) values ("& sItemCode  &","& sClassCode &","& Pack(sOrgCode) &","&_
              '  " "& sPreferred &","& sMinimum &","& sActual &","& sMarketRate &","&_
              '  " "& sWarrentyPeriod &","& sMinSalQty&","& sUnitSize &","& sVolume &","& Pack(sUnitSizeUOM) &","& Pack(sVolumeUOM) &","&_
              '  " "& sCreatedBy &",Convert(datetime,getDate(),103))"
              
                sQuery = "Insert into INV_M_ItemOrgSales (ItemCode,ClassificationCode,OrganisationCode,"&_
                        " WarrantyPeriod,PreferredMinQty,ItemDefinedBy,ItemDefinedOn) values ("& sItemCode  &","& sClassCode &","& Pack(sOrgCode) &","&_
                        " "& sWarrentyPeriod &","& sMinSalQty&","& sCreatedBy &",Convert(datetime,getDate(),103))"
                Response.Write "<p>"& sQuery
                con.execute sQuery
                
                
                    sQuery = "SELECT SellingPriceno FROM Sal_M_UnitPriceHdr"
	
	                Objrs.Open sQuery,Con
	                IF Not Objrs.EOF Then
		                iSellno = objrs(0)
	                Else
		                iSellno = "0"
	                End IF
	                Objrs.Close


	                'sRate = 0

	                IF CStr(iSellno) = "0" Then
                		
		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = "Select isNull(Max(SellingPriceno),0) + 1 From Sal_M_UnitPriceHdr"
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                IF Not Objrs.EOF Then
			                iSellno = Objrs(0)
		                End IF
		                Objrs.Close

		                    sQuery = " INSERT INTO Sal_M_UnitPriceHdr (SellingPriceno, AsonDate, CurrencyCode, "&_
				                     " PackingType, Price, UnitPrice,UpdatedBy,UpdatedOn) "&_
				                     " VALUES ("&iSellno&", Convert(datetime,'"&sAsonDate&"',103),1,0,'G','S',"&_
				                     " "& sCreatedBy&",Convert(DateTime,GetDate(),103)) "
		                    Response.Write sQuery
		                    con.execute sQuery

                			sQuery = "INSERT INTO Sal_M_UnitPriceDet (SellingPriceno, OudefinitionID,Itemcode, Classificationcode,ItemRate,RatePer,MarginPercent,MarginValue,OtherPercent,OtherValue,ItemPrice,EffectiveFrom) "
			                sQuery = sQuery &"VALUES ("&iSellno&",'"& sOrgCode &"', "&sItemCode&", "&sClassCode&", "&sPurRate&","& sPurRatePer &","& sMarPer&","& smarValue&","& sCharPer&" ,"& sCharValue &" ,"& sTotPrice &" ,Convert(DateTime,'"& sEffDate &"',103)) "
			                Response.Write "<p>query="&sQuery
			                Con.Execute sQuery
			            
	                '=========================== Amendment Part Starts Here ====================================
	                Else
		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = "Select isNull(Max(HistoryNo),0) + 1 From Sal_M_HistoryUnitPriceHdr "
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                IF Not Objrs.EOF Then
			                sHisno = Objrs(0)
		                End IF
		                Objrs.Close

		                sQuery = "SELECT ItemTypeID,isNull(UpdatedBy,0),Convert(Varchar,isNull(updatedOn,''),103),SellingPriceNo FROM Sal_M_UnitPriceHdr Where SellingPriceno = "&iSellno
		                'Response.write sQuery
		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = sQuery
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                IF Not Objrs.EOF Then
			                
			                sQuery = "Select HistoryEntryNo From SAL_M_UnitPriceHdrHistory Where HistoryEntryNo="&sHisno
			                rs.Open sQuery,con
			                If Not rs.EOF Then
				                sQry2 = " Update SAL_M_UnitPriceHdrHistory Set updatedOn = Convert(datetime,'"&Objrs(2)&"',103),updatedBy = "&Objrs(1)&" ,ModifiedBy="& sCreatedBy&",ModifiedOn=Convert(datetime,getdate(),103)  Where HistoryEntryNo = "&sHisno
			                Else
				                sQry2 = "INSERT INTO SAL_M_UnitPriceHdrHistory (HistoryEntryNo, ItemTypeID,updatedOn,updatedBy,ModifiedBy,ModifiedOn) "
				                sQry2 = sQry2 &"VALUES   ("&sHisno&", '"&Objrs(0)&"', Convert(datetime,'"&Objrs(2)&"',103),"&Objrs(1)&","& sCreatedBy &",Convert(datetime,getdate(),103)) "
			                End IF
			                rs.Close 
			                Response.Write "<p>Query1="&sQry2
			                Con.Execute sQry2
		                End IF
		                Objrs.Close
                		
		                'sQry2 = "SELECT Itemcode, Classificationcode, OudefinitionID, Price FROM Sal_M_UnitPriceDet "
		                'sQry2 = sQry2 &"Where SellingPriceno = "&iSellno
                		
		                sQry2 = " Select SellingPriceNo,OudefinitionID,Itemcode,Classificationcode,ItemRate,isNull(MarginPercent,0),"&_
				                " isNull(MarginValue,0),isNull(OtherPercent,0),isNull(OtherValue,0),isNull(ItemPrice,0),Convert(Datetime,"&_
				                " isNull(EffectiveFrom,''),103),isNull(RatePer,0),isNull(RateUOM,'') From Sal_M_UnitPriceDet "&_
				                " Where SellingPriceNo = "& iSellno

		                With objrs
			                .CursorLocation = 3
			                .CursorType = 3
			                .Source = sQry2
			                .ActiveConnection = con
			                .Open
		                End with
		                Set objrs.Activeconnection = nothing
		                Do While Not Objrs.EOF
			                'sQry2 = "INSERT INTO Sal_M_HistoryUnitPriceDet (HistoryNo, SellingPriceno, Itemcode, Classificationcode, OudefinitionID, Price) "
			                'sQry2 = sQry2 &"VALUES ("&sHisno&", "&iSellno&", "&Objrs(0)&", "&Objrs(1)&", '"&Objrs(2)&"', "&Objrs(3)&") "
			                sQry2 = " INSERT INTO SAL_M_UnitPriceDetHistory (HistoryEntryNo, SellingPriceno,OudefinitionID, Itemcode, Classificationcode, "&_
					                " ItemRate,MarginPercent,MarginValue,OtherPercent,OtherValue,ItemPrice,EffectiveFrom,RatePer,RateUOM) Values ( "&sHisno&", "&iSellno&","&_
					                " "&Objrs(1)&","&Objrs(2)&","&Objrs(3)&","&Objrs(4)&","&Objrs(5)&","&Objrs(6)&","&Objrs(7)&","&Objrs(8)&","&Objrs(9)&","&_
					                " "&Objrs(10)&","& objrs(11)&",'"& objrs(12)&"') "
			                Response.Write "<p>Query2="&sQry2
			                Con.Execute sQry2
			                Objrs.MoveNext
		                Loop
		                Objrs.Close

		                sQry2 = " UPDATE Sal_M_UnitPriceHdr SET AsonDate = Convert(datetime,'"&sAsonDate&"',103), "&_
				                " UpdatedBy = '"& sCreatedBy &"',UpdatedOn = Convert(datetime,getDate(),103) Where SellingPriceno = "&iSellno&" "
                		
		                Con.Execute sQry2
                		
		                		
				                sQuery = " UPDATE Sal_M_UnitPriceDet SET OudefinitionID= '"& sOrgCode &"' ,ItemRate = "&sPurRate&",MarginPercent = "& sMarPer &" ,MarginValue = "& sMarValue &",OtherPercent = "& sCharPer&" , "&_
						                 " OtherValue = "& sCharValue &" ,ItemPrice = "& sTotPrice &" , EffectiveFrom = Convert(DateTime,'"& sEffDate &"',103), "&_
						                 " RatePer = "& sPurRatePer &" "&_
						                 " WHERE SellingPriceno = "&iSellno&" AND Itemcode = "&sItemCode&" AND Classificationcode = "&sClassCode&" "
				                Response.Write "<p>Qry3="&sQuery
				                'Response.Write "<br><br><br>"
				                Con.execute sQuery,sRecno
				                IF sRecno = 0 Then
					                sQuery = "INSERT INTO Sal_M_UnitPriceDet (SellingPriceno, OudefinitionID,Itemcode, Classificationcode,ItemRate,MarginPercent,MarginValue,OtherPercent,OtherValue,ItemPrice,EffectiveFrom,RatePer) "
					                sQuery = sQuery &"VALUES ("&iSellno&",'"& sOrgCode &"', "&sItemCode&", "&sClassCode&", "&sPurRate&", "& sMarPer &","& sMarValue &","& sCharPer &" ,"& sCharValue &" ,"& sTotPrice &" ,Convert(DateTime,'"& sEffDate &"',103),"& sPurRatePer &") "
					                Response.Write "<p>Squer4="&sQuery
					                Con.Execute sQuery
				                End IF
	                End IF
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
