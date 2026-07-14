<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetNumberSeriesList.asp
	'Module Name				:	Sales (Master Creation)
	'Author Name				:	Manohar Prabhu
	'Created On					:	May 28,2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	SalesNoSeriesInsert.asp
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
	Dim sUnit,sActitity,sTemp,sTempVal,sQuery,Objrs,iCtr,sSerCode
	Dim sMon,sYear,sMonYr,sFinYear,sFinFrom,sFinTo,saTemp,sCrValue,sDrValue
	Dim oDom,newElem,Root,sSeriesCode,sSeriesNo,sRetTy,objRs1
	Dim iSeriesCode,iSeriesNo,iAgentCode,iAddSerNo,iAddSerCode,sAgentName
	Dim iEntNo,sItemTy,sInvTy,sSalTy,sAgTy,iSerNo,iSerCode,sItemVal,sSalVal,sInvVal,sNumTy
	Dim sItemDesc,sInvDesc,sSalDesc,sTempVar,sNoUsed,Objrs3
	Dim sSelClassCodes,sSelCatCodes,sSelClassCodesName,sSelCatCodesName
	
	sTempVal = Request("sVal")
	sTemp = Split(sTempVal,":")
	
	iCtr = 1
	sUnit = sTemp(0)
	sActitity = sTemp(1)
	
	Set oDom = server.CreateObject("Microsoft.xmlDom")
	Set Root = oDom.createElement("Root")
	oDom.appendChild Root
	
	sMon = Month(Date)
	sYear = Year(Date)

	IF CInt(sMon) <=9 Then
		sMon = 0&sMon
	End IF
	sMonYr = sMon&sYear
	sFinYear = GetFinancialYear(sMonYr)
	saTemp = Split(sFinYear,":")
	sYear = Right(saTemp(0),4)
	sMon = Mid(saTemp(0),4,2)
	sFinFrom = sYear&sMon

	sYear = Right(saTemp(1),4)
	sMon = Mid(saTemp(0),4,2)
	sFinTo = sYear&sMon
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs3 = Server.CreateObject("ADODB.RecordSet")
	
	sQuery = "Select H.NoSeriesTransactionNo,OrganisationCode,NumberFor,isNull(D.SeriesNo,0),isNull(D.SeriesCode,0) "&_
			 "From Sal_M_Noseries H left join Sal_M_NoSeriesDetails D on H.NoSeriesTransactionNo = D.NoSeriesTransactionNo "&_
			 "  Where H.NoSeriesStatus <> '1'  and OrganisationCode = '"&sUnit&"' and ActivityType = '"&sActitity&"' "
	
														 
	With Objrs
		.CursorType = 3
		.CursorLocation = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set Objrs.ActiveConnection = Nothing
	Do While Not Objrs.EOF
		iSeriesNo = Objrs(3)
		iSeriesCode = Objrs(4)
		
		IF CStr(iSeriesNo) <> "0" and CStr(iSeriesCode) <> "0" Then 
		
			sQuery = "Select ItemType,InvoiceType,SaleType,CommissionAgent,isNull(DetailsSeriesNo,0), "&_
					 "isNull(DetailsSeriesCode,0),IsNull(ItemValue,0),IsNull(InvoiceValue,0),IsNull(SaleValue,0),isNull(AgentCode,0), "&_
					 "isNull(AddSeriesNo,0),isNull(AddSeriesCode,0) From VwSalNumberSeriesSel "&_
					 "Where NoSeriesTransactionNo = "&Objrs(0)
					 
			With Objrs1
				.CursorType = 3
				.CursorLocation = 3
				.ActiveConnection = Con
				.Source = sQuery
				.Open
			End With
			Set Objrs1.ActiveConnection = Nothing
			IF Not Objrs1.EOF Then
				iEntNo = "1"
				sItemTy = objRs1(0)
				sInvTy = objRs1(1)
				sSalTy = objRs1(2)
				sAgTy = objRs1(3)
				iSeriesNo = objRs1(4)
				iSeriesCode = objRs1(5)
				sItemVal = objRs1(6)
				sInvVal = objRs1(7)
				sSalVal = objRs1(8)
				iAgentCode  = objRs1(9)
				iAddSerNo = objRs1(10)
				iAddSerCode = objRs1(11)
				
			End IF
			objRs1.Close
			
			sQuery = "Select OrgnPartyCode From App_M_PartyMaster Where PartyCode = "&iAgentCode&" "
			objRs1.Open sQuery,Con
			IF Not objRs1.EOF Then
				sAgentName = objRs1(0)
			End IF
			objRs1.Close
			
			IF CStr(iSeriesNo) = "0" and CStr(iSeriesCode) = "0" Then
				iSeriesNo = iAddSerNo
				iSeriesCode = iAddSerCode
			End IF
			
		Else
			sItemTy = "0"
			sInvTy = "0"
			sSalTy = "0"
			sAgTy = "1"
		End IF
		
		sNoUsed = CheckNoSerUsed(sUnit,sActitity,sItemTy,sInvTy,sSalTy,sAgTy,sItemVal,sInvVal,sSalVal,iAgentCode)
			
		IF CStr(sItemTy) = "0" Then
			sItemTy = "All"
			sItemDesc = ""
		Else
			sItemTy = "Specific"
			sTempVar = sItemVal
			sItemDesc = GetDesc(sTempVar,"I")
		End IF
			
		IF CStr(sInvTy) = "0" Then
			sInvTy = "All"
			sInvDesc = ""
		Else
			sInvTy = "Specific"
			sTempVar = sInvVal
			sInvDesc = GetDesc(sTempVar,"V")
		End IF
			
		IF CStr(sSalTy) = "0" Then
			sSalTy = "All"
			sSalDesc = ""
		Else
			sSalTy = "Specific"
			sTempVar = sSalVal
			sSalDesc = GetDesc(sTempVar,"S")
		End IF
			
		IF CStr(sAgTy) = "0" Then
			sAgTy = "Yes"
		Else
			sAgTy = "No"
		End IF
			
		IF CStr(Objrs(2)) = "B" Then
			sNumTy = "Both"
		Elseif CStr(Objrs(2)) = "D" Then
			sNumTy = "Domestic"
		Else
			sNumTy = "Export"
		End IF
		sSelClassCodes=""
		sSelCatCodes =""
		sSelClassCodesName=""
		sSelCatCodesName =""
		
		sQuery = "Select IsNull(ClassCode,0),isNull(CatCode,0) from Sal_M_NoSeriesClass where SeriesNo = "& iSeriesNo &" and SeriesCode = "& iSeriesCode 
		
		'Response.Write sQuery + vbCrLf + vbCrLf 
		
			objRs1.Open sQuery,con
			if not objRs1.EOF then
				do while not objRs1.EOF 
					if Trim(objRs1(0))<>"0" then
						sSelClassCodes = sSelClassCodes &","& objRs1(0)
					else
						sSelClassCodes = sSelClassCodes &","& "NULL"
					end if
					if Trim(objRs1(1))<>"0" then
						sSelCatCodes = sSelCatCodes &","& objRs1(1)
					else
						sSelCatCodes = sSelCatCodes &","& "NULL"
					end if
					
					objRs1.MoveNext
				loop
			end if
			objRs1.Close 
			
			
			if Trim(sSelClassCodes)<>"" then sSelClassCodes = Mid(sSelClassCodes,2)
			if Trim(sSelCatCodes)<>"" then sSelCatCodes = Mid(sSelCatCodes,2)
			
			if Trim(sSelClassCodes)<>"" then
			    sQuery = "Select GroupName from INV_M_Classification where GroupCode in ("& sSelClassCodes  &")"
			    objRs1.Open sQuery,con
			    if not objRs1.EOF then
				    do while not objRs1.EOF 
					    sSelClassCodesName =sSelClassCodesName & ","&  Trim(objRs1(0))
					    objRs1.MoveNext 
				    loop
			    end if
			    objRs1.Close 
			end if
			if Trim(sSelClassCodesName)<>"" then sSelClassCodesName =  Mid(sSelClassCodesName,2)
			
			if Trim(sSelCatCodes)<>"" then
			    sQuery = "Select CategoryName from Inv_M_ClassificationCategory where CategoryCode in ("& sSelCatCodes  &")"
			    objRs1.Open sQuery,con
			    if not objRs1.EOF then
				    do while not objRs1.EOF 
					    sSelCatCodesName =sSelCatCodesName & ","&  Trim(objRs1(0))
					    objRs1.MoveNext 
				    loop
			    end if
			    objRs1.Close 
			end if 'if Trim(sSelCatCodes)<>"" then
			
			if Trim(sSelCatCodesName)<>"" then sSelCatCodesName =  Mid(sSelCatCodesName,2)  
		
			
		
			
		 
		Set newElem = oDom.createElement("NumSeriesList")
		newElem.setAttribute "TransNo",Objrs(0)
		newElem.setAttribute "NumFor",sNumTy
		newElem.setAttribute "SeriesNo",iSeriesNo
		newElem.setAttribute "SeriesCode",iSeriesCode
		newElem.setAttribute "EntryNo",iEntNo
		newElem.setAttribute "ItemTy",sItemTy
		newElem.setAttribute "InvTy",sInvTy
		newElem.setAttribute "SaleTy",sSalTy
		newElem.setAttribute "AgentTy",sAgTy
		newElem.setAttribute "ItemValue",sItemVal
		newElem.setAttribute "InvValue",sInvVal 
		newElem.setAttribute "SaleValue",sSalVal
		newElem.setAttribute "AgentCode",iAgentCode
		newElem.setAttribute "AgentName",sAgentName
		newElem.setAttribute "ItemDesc",sItemDesc
		newElem.setAttribute "InvDesc",sInvDesc
		newElem.setAttribute "SaleDesc",sSalDesc
		newElem.setAttribute "EditCheck","N"
		newElem.setAttribute "NoUsed","N"
		newElem.setAttribute "SelClass",sSelClassCodes
		newElem.setAttribute "SelCat",sSelCatCodes
		newElem.setAttribute "SelClassName",sSelClassCodesName
		newElem.setAttribute "SelCatName",sSelCatCodesName
			
		Root.appendChild newElem
		iCtr = iCtr + 1
		Objrs.MoveNext
		sAgentName = ""
	Loop
	Objrs.Close
	
	Response.ContentType="text/xml"
	Response.Write oDom.xml											
	
%>

<%
	Function GetDesc(sVal,sCallTy)
		Dim sFullVal,arrtemp,Salrs
		Set Salrs = Server.CreateObject("ADODB.RecordSet")
		arrtemp = Split(sVal,":")
		
		IF CStr(sCallTy) = "I" Then
			IF UBound(arrtemp) = 0 Then
				sQuery = "Select ItemTypeName from Inv_M_ItemType Where ItemTypeID = '"&sVal&"' "
			Else
				'sVal = "'"&sVal&"'"
				sVal = Replace(sVal,":","','")
				sQuery = "Select ItemTypeName from Inv_M_ItemType Where ItemTypeID in ('"&sVal&"') "
			End IF
		Elseif CStr(sCallTy) = "S" Then
			sVal = Replace(sVal,":",",")
			sQuery = "Select InvoiceTypeName From Sal_M_InvoiceTypes Where InvoiceType in ("&sVal&") "
		End IF
		
		
		IF CStr(sCallTy) = "I" or CStr(sCallTy) = "S" Then
			With Salrs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = Con
				.Source = sQuery
				.Open
			End With
			Set Salrs.ActiveConnection = Nothing
			Do While Not Salrs.EOF
				sFullVal = sFullVal&", "& Salrs(0)
				Salrs.MoveNext
			Loop
			Salrs.Close
		Else
			arrtemp = Split(sVal,":")
			For iCtr = 0 To UBound(arrtemp)
				IF arrtemp(iCtr) = "A" Then
					sFullVal = sFullVal&", "&"MILL SALES "
				Elseif arrtemp(iCtr) = "T" Then
					sFullVal = sFullVal&", "&"TRANSFERS TO DEPOT "
				Elseif arrtemp(iCtr) = "C" Then
					sFullVal = sFullVal&", "&"TRANSFER TO CONVERTORS FOR CONVERSION "
				Elseif arrtemp(iCtr) = "S" Then
					sFullVal = sFullVal&", "&"INVOICE FOR SAMPLE ITEMS "
				Elseif arrtemp(iCtr) = "D" Then
					sFullVal = sFullVal&", "&"UNPROCESSED ORDERS DIRECT INVOICE "
				Elseif arrtemp(iCtr) = "U" Then
					sFullVal = sFullVal&", "&"TRANSFER TO GROUP UNITS OR COMPANIES "
				Elseif arrtemp(iCtr) = "I" Then
					sFullVal = sFullVal&", "&"STOCK ISSUED FOR SALES "
				Elseif arrtemp(iCtr) = "R" Then
					sFullVal = sFullVal&", "&"TRANSFER INVOICE - RETURN OF SUBCONTRACT ITEMS "
				End IF
			Next
		End IF
		
		sFullVal = Mid(sFullVal,2)
		'Response.Write sFullVal &"<br><br>"
		GetDesc = sFullVal
	End Function
	

	
	Function CheckNoSerUsed(sUnitid,sAct,sItem,sInv,sSales,sAg,sItVal,sInvVal,sSVal,sAgCode)
		Dim Salrs,sTempVal1,sTempVal2,sTempVal3,sCheckVal
		
		sTempVal1 = sSVal
		sTempVal2 = sItVal
		sTempVal3 = sInvVal
		
		if trim(sTempVal1)<>"" then
		    sTempVal1 = Replace(sTempVal1,":",",")
		end if 
		if trim(sTempVal2)<>"" then
		    sTempVal2 = Replace(Trim(sTempVal2),":","','")
		end if 
		if trim(sTempVal3)<>"" then
		    sTempVal3 = Replace(Trim(sTempVal3),":","','")
		end if
		
		'======================== For Order Confirmation ===========================================
		
		IF CStr(sAct) = "OCR" Then
			sQuery = "Select OrderConfirmationNo,ConfCurrency FROM Sal_T_OCHeader Where OrderConfByUnit = '"&sUnitid&"' "
			
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and TypeofItems IN ('"&sTempVal2&"') "
			End IF
			
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and TypeOfSale IN ("&sTempVal1&") "
			End IF
			
			IF CStr(sAgTy) = "0" Then
				sQuery = sQuery &"and AgentCode IN ("&iAgentCode&") "
			End IF
		End IF
		'*********************************************************************************************
		'======================== For Quotation ======================================================
		IF CStr(sAct) = "QUT" Then
			sQuery = "Select QuotationNumber,QuotedCurrency FROM Sal_T_QuotationHeader Where QuotationNumber is Not Null "
			
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and TypeofItems IN ('"&sTempVal2&"') "
			End IF
			
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and TypeOfSale IN ("&sTempVal1&") "
			End IF
			
			IF CStr(sAgTy) = "0" Then
				sQuery = sQuery &"and AgentCode IN ("&iAgentCode&") "
			End IF
		End IF
		'**********************************************************************************************
		'======================== For Order Processing ================================================
		IF CStr(sAct) = "ORP" Then
			sQuery = "Select Top 1 H.OrderConfirmationNo FROM Sal_T_ProcessedOC P, Sal_T_OCHeader H, Sal_T_AdditionalAgents A Where  "&_
					 "H.OrderConfirmationNo = P.OrderConfirmationNo and "&_
					 "H.OrderConfByUnit = '"&sUnitid&"' and P.ForeCastOrderNo is Null "&_
					 "and H.OrderConfirmationNo *=  A.OrderConfirmationNo "
			
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and H.TypeofItems IN ('"&sTempVal2&"') "
			End IF
			
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and P.Saletype IN ("&sTempVal1&") "
			End IF
			
			IF CStr(sAgTy) = "0" Then
				sQuery = sQuery &"and A.AgentCode IN ("&iAgentCode&") "
			End IF		 
			
			IF CStr(sInvTy) <> "0" Then
				sQuery = sQuery &"and P.InvoiceType IN ('"&sTempVal3&"') "
			End IF
		
		End IF
		
		With Objrs3
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		IF Not Objrs3.EOF Then
			sCheckVal = "Y"
			CheckNoSerUsed = "Y"
		Else
			sCheckVal = "N"
			CheckNoSerUsed = "N"
		End IF
		Objrs3.Close
		
		IF CStr(sAct) = "ORP" Then
			IF CStr(sCheckVal) = "N" Then
				sQuery = "Select Top 1 P.ForeCastOrderNo FROM Sal_T_ProcessedOC P, Sal_T_ForCastOrdHeader H Where  "&_
						 "H.OrderForUnit = '"&sUnitid&"' and P.ForeCastOrderNo is Not Null  "&_
						 "and H.OrderNumber =  P.ForecastOrderNo "
						 
				IF CStr(sItemTy) <> "0" Then
					sQuery = sQuery &"and H.TypeofItems IN ('"&sTempVal2&"') "
				End IF
			
				IF CStr(sSales) <> "0" Then
					sQuery = sQuery &"and P.Saletype IN ("&sTempVal1&") "
				End IF
				
				IF CStr(sInvTy) <> "0" Then
					sQuery = sQuery &"and P.InvoiceType IN ('"&sTempVal3&"') "
				End IF
				
				With Objrs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = Con
					.Open
				End With
				IF Not Objrs3.EOF Then
					CheckNoSerUsed = "Y"
				Else
					CheckNoSerUsed = "N"
				End IF
				Objrs3.Close
			Else
				CheckNoSerUsed = "Y"
			End IF
		End IF
		'**************************************************************************************************
		'======================================= PIS No Check =============================================
		IF CStr(sAct) = "PIS" Then
			sQuery = "Select Top 1 H.OrderConfirmationNo FROM Sal_T_ProcessedOC P, Sal_T_OCHeader H, "&_
					 "Sal_T_AdditionalAgents A, Sal_T_ForProcessingDet PD Where  "&_
					 "H.OrderConfirmationNo = P.OrderConfirmationNo and "&_
					 "H.OrderConfByUnit = '010101' and P.ForeCastOrderNo is Null "&_
					 "and H.OrderConfirmationNo *=  A.OrderConfirmationNo and "&_
					 "PD.ProcessingNo = P.ProcessingNo and PD.QtyForManufacturing <> 0 "
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and H.TypeofItems IN ('"&sTempVal2&"') "
			End IF
			
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and P.Saletype IN ("&sTempVal1&") "
			End IF
			
			IF CStr(sAgTy) = "0" Then
				sQuery = sQuery &"and A.AgentCode IN ("&iAgentCode&") "
			End IF		 
			
			IF CStr(sInvTy) <> "0" Then
				sQuery = sQuery &"and P.InvoiceType IN ('"&sTempVal3&"') "
			End IF
			
			With Objrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With
			IF Not Objrs3.EOF Then
				sCheckVal = "Y"
				CheckNoSerUsed = "Y"
			Else
				sCheckVal = "N"
				CheckNoSerUsed = "N"
			End IF
			Objrs3.Close
			IF CStr(sAct) = "PIS" Then
				IF CStr(sCheckVal) = "N" Then
					sQuery = "Select Top 1 P.ForeCastOrderNo FROM Sal_T_ProcessedOC P, Sal_T_ForCastOrdHeader H, "&_
							 "Sal_T_ForProcessingDet PD Where H.OrderForUnit = '010101' and "&_
							 "P.ForeCastOrderNo is Not Null and H.OrderNumber =  P.ForecastOrderNo "&_
							 "and PD.ProcessingNo = P.ProcessingNo "
							 
					IF CStr(sItemTy) <> "0" Then
						sQuery = sQuery &"and H.TypeofItems IN ('"&sTempVal2&"') "
					End IF
			
					IF CStr(sSales) <> "0" Then
						sQuery = sQuery &"and P.Saletype IN ("&sTempVal1&") "
					End IF
				
					IF CStr(sInvTy) <> "0" Then
						sQuery = sQuery &"and P.InvoiceType IN ('"&sTempVal3&"') "
					End IF
				
					With Objrs3
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQuery
						.ActiveConnection = Con
						.Open
					End With
					IF Not Objrs3.EOF Then
						CheckNoSerUsed = "Y"
					Else
						CheckNoSerUsed = "N"
					End IF
					Objrs3.Close
				Else
					CheckNoSerUsed = "Y"
				End IF
			End IF 
		End IF
		'***************************************************************************************************
		'======================================== RIS NO Check =============================================
		IF CStr(sAct) = "RIS" Then
			sQuery = "Select Top 1 H.OrderConfirmationNo FROM Sal_T_ProcessedOC P, Sal_T_OCHeader H,  "&_
					 "Sal_T_AdditionalAgents A, Sal_T_ForProcessingDet PD Where  "&_
					 "H.OrderConfirmationNo = P.OrderConfirmationNo and "&_
					 "H.OrderConfByUnit = '010101' and P.ForeCastOrderNo is Null "&_
					 "and H.OrderConfirmationNo *=  A.OrderConfirmationNo and "&_
					 "PD.ProcessingNo = P.ProcessingNo and PD.QtyForRepacking <> 0 "
			
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and H.TypeofItems IN ('"&sTempVal2&"') "
			End IF
			
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and P.Saletype IN ("&sTempVal1&") "
			End IF
			
			IF CStr(sAgTy) = "0" Then
				sQuery = sQuery &"and A.AgentCode IN ("&iAgentCode&") "
			End IF		 
			
			IF CStr(sInvTy) <> "0" Then
				sQuery = sQuery &"and P.InvoiceType IN ('"&sTempVal3&"') "
			End IF
			
			With Objrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With
			IF Not Objrs3.EOF Then
				CheckNoSerUsed = "Y"
			Else
				CheckNoSerUsed = "N"
			End IF
			Objrs3.Close
		End IF
		'**********************************************************************************************
		'========================================= PUR No Check =======================================
		IF CStr(sAct) = "PUR" Then
			sQuery = "Select Top 1 H.OrderConfirmationNo FROM Sal_T_ProcessedOC P, Sal_T_OCHeader H,  "&_
					 "Sal_T_AdditionalAgents A, Sal_T_ForProcessingDet PD Where  "&_
					 "H.OrderConfirmationNo = P.OrderConfirmationNo and "&_
					 "H.OrderConfByUnit = '010101' and P.ForeCastOrderNo is Null "&_
					 "and H.OrderConfirmationNo *=  A.OrderConfirmationNo and "&_
					 "PD.ProcessingNo = P.ProcessingNo and PD.QtyForPurchase <> 0 "
			
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and H.TypeofItems IN ('"&sTempVal2&"') "
			End IF
			
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and P.Saletype IN ("&sTempVal1&") "
			End IF
			
			IF CStr(sAgTy) = "0" Then
				sQuery = sQuery &"and A.AgentCode IN ("&iAgentCode&") "
			End IF		 
			
			IF CStr(sInvTy) <> "0" Then
				sQuery = sQuery &"and P.InvoiceType IN ('"&sTempVal3&"') "
			End IF
			
			With Objrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With
			IF Not Objrs3.EOF Then
				CheckNoSerUsed = "Y"
			Else
				CheckNoSerUsed = "N"
			End IF
			Objrs3.Close
		End IF
		'********************************************************************************************
		'=================================== DIS No Check ===========================================
		IF CStr(sAct) = "DIS" Then
			sQuery = "Select Top 1 H.OrderConfirmationNo FROM Sal_T_ProcessedOC P, Sal_T_OCHeader H,  "&_
					 "Sal_T_AdditionalAgents A, Sal_T_ForProcessingDet PD Where  "&_
					 "H.OrderConfirmationNo = P.OrderConfirmationNo and "&_
					 "H.OrderConfByUnit = '010101' and P.ForeCastOrderNo is Null "&_
					 "and H.OrderConfirmationNo *=  A.OrderConfirmationNo and "&_
					 "PD.ProcessingNo = P.ProcessingNo and PD.QtyForDespatch <> 0 "
			
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and H.TypeofItems IN ('"&sTempVal2&"') "
			End IF
			
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and P.Saletype IN ("&sTempVal1&") "
			End IF
			
			IF CStr(sAgTy) = "0" Then
				sQuery = sQuery &"and A.AgentCode IN ("&iAgentCode&") "
			End IF		 
			
			IF CStr(sInvTy) <> "0" Then
				sQuery = sQuery &"and P.InvoiceType IN ('"&sTempVal3&"') "
			End IF
			
			With Objrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With
			IF Not Objrs3.EOF Then
				CheckNoSerUsed = "Y"
			Else
				CheckNoSerUsed = "N"
			End IF
			Objrs3.Close
		End IF
		'**********************************************************************************************
		'================================== INV No Check ==============================================
		IF CStr(sAct) = "INV" Then
			sQuery = "Select SaleTransactionNo From Sal_T_InvoiceHeader Where TypeofInvoice <> 'T' "&_
					 "and InvoicedForUnit = '010101' "
					 
			IF CStr(sItemTy) <> "0" Then
				sQuery = sQuery &"and TypeofItems IN ('"&sTempVal2&"') "
			End IF
				
			IF CStr(sSales) <> "0" Then
				sQuery = sQuery &"and TypeofSale IN ("&sTempVal1&") "
			End IF
				
			IF CStr(sInvTy) <> "0" Then
				sQuery = sQuery &"and TypeofInvoice IN ('"&sTempVal3&"') "
			End IF
		
			With Objrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = Con
				.Open
			End With
			IF Not Objrs3.EOF Then
				CheckNoSerUsed = "Y"
			Else
				CheckNoSerUsed = "N"
			End IF
			Objrs3.Close
		End IF
		'**********************************************************************************************
	End Function
%>