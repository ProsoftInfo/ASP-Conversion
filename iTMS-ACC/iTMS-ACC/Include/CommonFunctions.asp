<%
Function GetTransDate()

	Dim sFinPeriod,sFinPeriodFrom, sFinPeriodTo,sCurDate
	Dim sRetVal

	sFinPeriod = Session("FinPeriod")

	sFinPeriodFrom = "01/04/" & Mid(sFinPeriod,1,4)
	sFinPeriodTo = "31/03/" & Mid(sFinPeriod,6,4)

	sCurDate = right("0" & trim(day(Date())),2) & "/" & right("0"& trim(month(Date())),2) &"/"&Year(Date())


	If datediff("d", sFinPeriodTo,sCurDate) > 0 Then
		sRetVal = sFinPeriodTo
	Else
		sRetVal = sCurDate
	End If

	GetTransDate = sRetVal
End Function

%>


<%
Function GetTransDateOLD()

	Dim sFinPeriod,sFinPeriodFrom, sFinPeriodTo,sCurDate
	Dim sRetVal

	sFinPeriod = Session("FinPeriod")
	sFinPeriodFrom = "01/04/" & Mid(sFinPeriod,1,4)
	sFinPeriodTo = "31/03/" & Mid(sFinPeriod,6,4)
	sCurDate = Month(Date())&"/"&Day(Date())&"/"&Year(Date())


	If datediff("d", formatdate(sFinPeriodTo),sCurDate) > 0 Then
		sRetVal = sFinPeriodTo
	Else
		sRetVal = FormatDate(sCurDate)
	End If

	GetTransDateOLD = sRetVal
End Function
'''
Function RefTypePop(RefCode,RefAppNo)
    'RefCode - RefCodeNo 
    'RefAppNo - RefApplicationNo
    Dim sQuery,rsRefType
    
    set rsRefType = Server.CreateObject("ADODB.Recordset")
    Response.Write "<option value='N'>None</option>"
    sQuery = "SELECT ReferenceEntryNo,ReferenceName FROM VW_ReferenceTypes WHERE RefCodeNo = "& RefCode &" and RefApplicationCode = "& RefAppNo 
	rsRefType.Open sQuery,con
	if not rsRefType.EOF then
		do while not rsRefType.EOF 
			Response.Write "<option value="& rsRefType(0) &">"& rsRefType(1) &"</option>"
			rsRefType.MoveNext
		loop
	end if
	rsRefType.Close 
End Function
'**************************
Function RefTypePopType(RefCode,RefAppNo,sRcptAs)
    'RefCode - RefCodeNo 
    'RefAppNo - RefApplicationNo
    Dim sQuery,rsRefType,TypePara
    TypePara = sRcptAs
    if trim(TypePara)="" or IsNull(TypePara) then TypePara = "NULL"
    if trim(TypePara)<>"NULL" then TypePara = pack(TypePara)
    
    set rsRefType = Server.CreateObject("ADODB.Recordset")
     Response.Write "<option value='N'>None</option>"
    sQuery = "SELECT ReferenceEntryNo,ReferenceName FROM VW_ReferenceTypes WHERE RefCodeNo = "& RefCode &" and RefApplicationCode = "& RefAppNo &" and Type = "& TypePara
	rsRefType.Open sQuery,con
	if not rsRefType.EOF then
		do while not rsRefType.EOF 
			Response.Write "<option value="& rsRefType(0) &">"& rsRefType(1) &"</option>"
			rsRefType.MoveNext
		loop
	end if
	rsRefType.Close 
	
	if trim(RefCode)="4" and trim(RefAppNo)="2" then
	    sQuery = "SELECT ReferenceEntryNo,ReferenceName FROM VW_ReferenceTypes WHERE RefCodeNo = "& RefCode &" and RefApplicationCode = "& RefAppNo &" and Type is Null"
	    rsRefType.Open sQuery,con
	    if not rsRefType.EOF then
		    do while not rsRefType.EOF 
			    Response.Write "<option value="& rsRefType(0) &">"& rsRefType(1) &"</option>"
			    rsRefType.MoveNext
		    loop
	    end if
	    rsRefType.Close 
	end if 'if trim(RefCode)="4" and trim(RefAppNo)="2" then
	
End Function
'---------------------------------------------------
Function GetRefNoDate(RefTye,RefNo)
    'RefType - ReferenceEntryNo
    'RefNo - ReferenceNoColumnName
    Dim sQuery,rsRefType,rsTemp,sReturnValue
    set rsRefType = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    
    if trim(RefTye)<>"N" and Trim(RefNo)<>"" then
        sQuery = "Select distinct ReferenceEntryNo,ReferenceName,ReferenceNoColumnName,ReferenceCodeColumnName,ReferenceDateColumnName,ReferenceSourceTableName  from VW_ReferenceTypes where ReferenceEntryNo = "& RefTye 
        rsRefType.open sQuery,con
        if not rsRefType.eof then
            sReturnValue = trim(rsRefType(1))
            sQuery = "Select "& rsRefType(2) &",isNull("& rsRefType(3) &","& rsRefType(2) &"), Convert(varchar,"& rsRefType(4)& ",103) from "& rsRefType(5) &" where "& rsRefType(2)&" = "& RefNo 
            rsTemp.open sQuery,con
            if not rsTemp.eof then
                sReturnValue =  sReturnValue &","& trim(rsTemp(1)) &" - "& trim(rsTemp(2))
            end if
            rsTemp.close 
        end if
        rsRefType.close 
        GetRefNoDate = sReturnValue
    else
        GetRefNoDate =""
    end if 'if trim(RefTye)<>"N" then
    
End Function
%>

<%
Function GetNoSeriesDate()
	GetNoSeriesDate = "01/04/"&Trim(Left(Session("FinPeriod"),4))
End Function
%>

<%
Function GetStatusFromAcc(AppNo,RefNo)
    Dim rsTemp,sQuery,sTransType
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    if AppNo = 2 then sTransType = "PJR"
    if AppNo = 3 then sTransType = "SJR"
    sQuery = " Select CreatedTransNo from Acc_T_CreatedVoucherHeader where TransactionType='"& sTransType &"' "&_
             " and CreatedVouchStatus in (010101,010102,010103) and OtherApplnTransNo = "& RefNo 
    rsTemp.Open sQuery,con
    
    if not rsTemp.EOF then
        GetStatusFromAcc = "Y"
    end if
End Function
%>

<%
Function CheckNoSeries(sActivityID,sFinFromMrYear,sFinToMrYear,sOrgID,iApplicationCode)
Dim rsTemp
Dim sSeriesCode,sSeriesNo,sQuery
set rsTemp = server.createObject("ADODB.Recordset")
	Select Case iApplicationCode
		Case 2
			sQuery =  "SELECT SeriesNo,SeriesCode FROM PUR_M_Noseries where ActivityType = "& sActivityID &" and OrganisationCode ="& sOrgID
		Case 3 
			sQuery =  "SELECT SeriesNo,SeriesCode FROM Sal_M_Noseries where ActivityType = '"& sActivityID &"' and OrganisationCode ="& sOrgID
		Case 4
			sQuery =  "SELECT SeriesNo,SeriesCode FROM Inv_M_NumberSeries where ActivityType = '"& sActivityID &"' and OrganisationCode ="& sOrgID
		Case 5
			sQuery =  "SELECT SeriesNo,SeriesCode FROM MTN_M_NoSeries where ActivityType = '"& sActivityID &"' and OrganisationCode ="& sOrgID
		Case 8
			sQuery =  "SELECT SeriesNo,SeriesCode FROM FDP_M_Noseries where ActivityType = '"& sActivityID &"' and OrganisationCode ="& sOrgID
					
	End Select
	rsTemp.open sQuery,con
	if not rsTemp.eof then
		sSeriesNo = rsTemp(0)
		sSeriesCode = rsTemp(1)
	else
		CheckNoSeries = false
	end if
	rstemp.close
	
	if trim(sSeriesNo)<>"" and trim(sSeriesCode)<>"" then
		sQuery = "Select * from APP_R_NoSeriesModuleEntry where SeriesNo = "& sSeriesNo &" and SeriesCode = "& sSeriesCode &" and Period = "&sFinToMrYear
		rsTemp.open sQuery,con
		if not rsTemp.eof then
			CheckNoSeries = true
		else
			CheckNoSeries = false
		end if
		rstemp.close
	else
		CheckNoSeries = false
	end if
End Function

Function GetItemRate(OrgID,FinPeriod,ClassCode,ItemCode,ValuationType)	
	Dim sTemp,rs,sSql,sFinPeriodFrom,sFinPeriodTo,nYrlyClosingValue,nYrlyClosingStock
	Dim nItemRate
	
	set rs = Server.CreateObject("ADODB.Recordset")
	sTemp = Split(FinPeriod,":")
	
	sFinPeriodFrom = "01/04/" & sTemp(0)
	sFinPeriodTo = "31/03/" & sTemp(1)
	
	nYrlyClosingValue = cdbl("0")
	nYrlyClosingStock = cdbl("0")
	nItemRate = cdbl("0")
	
	sSql = " Select isNull(YearClosingValue,0),isNull(YearClosingStock,0) From Inv_T_ItemYearlyStock "&_
		   " where OrganisationCode = '"& OrgID&"' and ClassificationCode = "& ClassCode &" and ItemCode = "& ItemCode &" "&_
		   " and FinancialYearFrom = Convert(DateTime,'"& sFinPeriodFrom &"',103) and FinancialYearTo = Convert(DateTime,'"& sFinPeriodTo &"',103) "
	'Response.Write "<p><font color=red>sql="&sSql
	rs.Open sSql,con
	Do While Not rs.EOF 
		
		nYrlyClosingValue = nYrlyClosingValue + cdbl(rs(0))
		nYrlyClosingStock = nYrlyClosingStock + cdbl(rs(1))
		
		rs.MoveNext 
	Loop
	rs.Close 
	
	If nYrlyClosingValue <> "0" and nYrlyClosingStock <> "0" Then
		nItemRate = cdbl(nYrlyClosingValue)/cdbl(nYrlyClosingStock)
	End IF
	GetItemRate = FormatNumber(nItemRate,2,,,0)	
End Function


Function GetItemSalePrice(OrgID,TransactionDate,ClassCode,ItemCode,PartyCode)
	Dim nSalePrice,rs,rs1,sSql,sSupplierEligible,nItemPrice
	
	set rs  = Server.CreateObject("ADODB.Recordset")
	set rs1 = Server.CreateObject("ADODB.Recordset")
	nSalePrice = cdbl("0")
	nItemPrice = cdbl("0")
	
	If PartyCode <> "" and PartyCode <> "0" Then
		sSql = " Select isNull(SUM(isNull(CustMarketPrice,0)),0) From Inv_R_ItemSupplier Where Organisationcode = '"& OrgID &"' and  Itemcode = "& ItemCode &""&_
			   " and ClassificationCode = "& ClassCode&" and PartyCode = "& PartyCode &""
		sSupplierEligible = "Y"
	Else
		sSql = " Select isNull(SUM(isNull(ItemRate,0)),0),isNull(SUM(isNull(ItemPrice,0)),0) From Sal_M_UnitPriceDet Where OudefinitionID = '"& OrgID &"' and Itemcode = "& ItemCode &"  "&_
			   " and ClassificationCode = "& ClassCode&""
			   '  and Convert(DateTime,EffectiveFrom,103) >= convert(DateTime,'"& TransactionDate &"',103)
		sSupplierEligible = "N"
	End IF
	'Response.Write "<p><font color=red>"&sSql
	rs.Open sSql,con
	If Not rs.EOF Then	
		nSalePrice = rs(0)
		
		If sSupplierEligible = "N" Then
			nSalePrice = rs(1)		'Item Price
			If nSalePrice = "0" Then
				nSalePrice = rs(0)	'Item Rate
			End IF
		End IF
		
		If sSupplierEligible = "Y" and nSalePrice = "0" Then
			sSql = " Select isNull(SUM(isNull(ItemRate,0)),0),isNull(SUM(isNull(ItemPrice,0)),0) From Sal_M_UnitPriceDet Where OudefinitionID = '"& OrgID &"' and Itemcode = "& ItemCode &"  "&_
				   " and ClassificationCode = "& ClassCode&""
		'	Response.Write "<p><font color=red>"&sSql
			rs1.Open sSql,con
			If Not rs1.EOF Then
				nSalePrice = rs1(1)		'Item Prices
				If nSalePrice = "0" Then
					nSalePrice = rs(0)	'Item Rate
				End IF
			End IF
			rs1.Close 		
		End IF
	End If
	rs.Close 
	
	GetItemSalePrice = FormatNumber(nSalePrice,2,,,0)
End Function

Function GetMarketPrice(OrgID,ClassCode,ItemCode)
	Dim rs,nMarketPrice,sSql
	
	set rs = server.CreateObject("ADODB.RecordSet")
	
	sSql = " Select isNull(SUM(isNull(ItemPrice,0)),0) From Sal_M_UnitPriceDet Where OudefinitionID = '"& OrgID &"' and Itemcode = "& ItemCode &"  "&_
		   " and ClassificationCode = "& ClassCode&""
	'Response.Write "<p><font color=red>"&sSql
	rs.open sSql,con
	If Not rs.eof Then
		nMarketPrice = rs(0)
	End IF
	rs.close
	
	GetMarketPrice = nMarketPrice
End Function

Function GetItemPurchasePrice(OrgID,TransactionDate,ClassCode,ItemCode,PartyCode)
	Dim nPurchasePrice,rs,rs1,sSql,sPurchaseEligible
	
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs1 = Server.CreateObject("ADODB.Recordset")
	nPurchasePrice  = cdbl("0")
	
	If PartyCode <> "" and PartyCode <> "0" Then
		sSql = " Select isNull(SUM(isNull(SuppMarketPrice,0)),0) From Inv_R_ItemSupplier Where Organisationcode = '"& OrgID &"' and  Itemcode = "& ItemCode &""&_
			   " and ClassificationCode = "& ClassCode&" and PartyCode = "& PartyCode &""
		sPurchaseEligible = "Y"
	Else
		sSql = " Select isNull(SUM(isNull(ItemPrice,0)),0) From Sal_M_UnitPriceDet Where OudefinitionID = '"& OrgID &"' and Itemcode = "& ItemCode &"  "&_
			   " and ClassificationCode = "& ClassCode&""
			   '  and Convert(DateTime,EffectiveFrom,103) >= convert(DateTime,'"& TransactionDate &"',103)
		sPurchaseEligible = "N"
	End IF
	
	rs.Open sSql,con
	If Not rs.EOF Then	
		nPurchasePrice = rs(0)
		If sPurchaseEligible = "Y" and nPurchasePrice = "0" Then
			sSql = " Select isNull(SUM(isNull(ItemPrice,0)),0) From Sal_M_UnitPriceDet Where OudefinitionID = '"& OrgID &"' and Itemcode = "& ItemCode &"  "&_
			       " and ClassificationCode = "& ClassCode&""
			rs1.open sSql,con
			If Not rs1.Eof Then
				nPurchasePrice = rs1(0)
			End IF
			rs1.close 
		End IF
	End If
	rs.Close 
	
	GetItemPurchasePrice = FormatNumber(nPurchasePrice,2,,,0)
End Function
%>

<%
Function FindDiscount(ItemCode,ClassCode,OrgCode,Qty,Value,PartyCode)
	Dim sReturnValue,rsCusDis,rsItemDis,sQuery,sPrecedence,sEligible
    Dim iQtyDis,iQtyFrom,iQtyTo,iValDis,iValFrom,iValTo
    
    set rsCusDis = Server.CreateObject("ADODB.Recordset")
    set rsItemDis = Server.CreateObject("ADODB.Recordset")
    
    sEligible = false 
    
	If PartyCode<>"" and PartyCode<>"0" Then 
       sQuery = "Select QtyDiscountOffered,QuantityFrom,QuantityTo,ValueDiscountOffered,ValueFrom,"&_
                " ValueTo,Precedence from INV_T_ItemSupplierDiscount where PartyCode = "& PartyCode &" and "&_
                " ItemCode = "& ItemCode &" and ClassificationCode = "& ClassCode &" and OrganisationCode = '"& OrgCode &"' "
    Else
       sQuery = "Select QtyDiscountOffered,QuantityFrom,QuantityTo,ValueDiscountOffered,ValueFrom,"&_
                " ValueTo,Precedence from Inv_M_ItemOrgSaleDiscount where ItemCode = "& ItemCode &" and "&_
                " ClassificationCode = "& ClassCode &" and OrganisationCode = '"& OrgCode &"' "
    End if
   ' Response.Write sQuery
    rsCusDis.Open sQuery,con
    If not rsCusDis.eof Then
       do while not rsCusDis.eof
           iQtyDis = rsCusDis(0)
           iQtyFrom = rsCusDis(1)
           iQtyTo = rsCusDis(2)
           iValDis = rsCusDis(3)
           iValFrom = rsCusDis(4)
           iValTo = rsCusDis(5)
           sPrecedence = rsCusDis(6)
           if sPrecedence = "Q" then
               if cdbl(Qty)>=cdbl(iQtyFrom) and cdbl(Qty) <=cdbl(iQtyTo) then
                   sEligible = true
                   FindDiscount = iQtyDis 
                   exit Do
               end if
           elseif sPrecedence = "V" then
               if cdbl(Value)>=cdbl(iValFrom) and cdbl(Value) <=cdbl(iValTo) then
                   sEligible = true
                   FindDiscount = iValDis 
                   exit Do
               end if
           end if 
           rsCusDis.movenext
       loop
    End If
    rsCusDis.close
    if (PartyCode<>"" and PartyCode<>"0") and sEligible = false then
       sQuery = "Select QtyDiscountOffered,QuantityFrom,QuantityTo,ValueDiscountOffered,ValueFrom,"&_
                " ValueTo,Precedence from Inv_M_ItemOrgSaleDiscount where ItemCode = "& ItemCode &" and "&_
                " ClassificationCode = "& ClassCode &" and OrganisationCode = "& OrgCode
       ' Response.Write "----"&sQuery
        rsCusDis.Open sQuery,con
        If not rsCusDis.eof Then
           do while not rsCusDis.eof
               iQtyDis = rsCusDis(0)
               iQtyFrom = rsCusDis(1)
               iQtyTo = rsCusDis(2)
               iValDis = rsCusDis(3)
               iValFrom = rsCusDis(4)
               iValTo = rsCusDis(5)
               sPrecedence = rsCusDis(6)
               if sPrecedence = "Q" then
                   if cdbl(Qty)>=cdbl(iQtyFrom) and cdbl(Qty) <=cdbl(iQtyTo) then
                       sEligible = true
                       FindDiscount = iQtyDis 
                       exit Do
                   end if
               elseif sPrecedence = "V" then
                   if cdbl(Value)>=cdbl(iValFrom) and cdbl(Value) <=cdbl(iValTo) then
                       sEligible = true
                       FindDiscount = iValDis 
                       exit Do
                   end if
               end if 
               rsCusDis.movenext
           loop
       End If
       rsCusDis.close
    End if ' if (PartyCode<>"" and PartyCode<>"0") and sEligible = false then
End Function    
%>
<%
Function ErrorTracking(sAppCode,sProCode,sActCode,sActAction,sProName,sErrCode,sErrDesc,sTracedOn,sTracedTime,sUserId)
    Dim rsErrTrace,iSLNo,sQuery,rsErrTemp
    set rsErrTrace = Server.CreateObject("ADODB.Recordset")
    set rsErrTemp = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select isNull(Max(ErrorTrackNo),0)+1 from APP_T_ErrorTraced"
    rsErrTemp.Open sQuery,con
    if not rsErrTemp.EOF then
        iSLNo = rsErrTemp(0)
    end if
    rsErrTemp.Close 
    
    sQuery =   " Insert into APP_T_ErrorTraced (ErrorTrackNo,ApplicationCode,"&_
               " ProcessCode,ActivityCode,ActivityAction,ProgramName,AppErrorCode,"&_
               " AppErrorDescription,ErrorTracedOn,ErrorTracedTime,LoggedInUserID,ErrorStatus)"&_
               " values("& iSLNo &","& sAppCode &","& sProCode &","&sActCode &","& Pack(sActAction) &","&_
               " "& Pack(sProName)&","& Pack(sErrCode)&","& Pack(sErrDesc)&",Convert(datetime,'"& sTracedOn &"',103),"&_
               " "& Pack(sTracedTime)&","& sUserId&",'Traced')"
        Response.Write "<p>"& sQuery
    con.execute sQuery
  
End Function
%>
<%
Function GetInfoRefType(sRefType,sAppRefNo,sOrgID)

Dim sRefCode,sRefDate,sRefName,sOthRef,sRemarks,sRefNo,sQuery,sRefTable,sRefWhereClause
Dim rsTemp
Dim sCodes,sDate,sName,sOther,sRem,sNo,sOthRefNoDate,sOthNoDate
set rsTemp = Server.CreateObject("ADODB.Recordset")


	sQuery = " Select isNull(ReferenceCodeColumnName,ReferenceNoColumnName),isNull(ReferenceDateColumnName,''),isNull(ReferenceName,''),"&_
			 " isNull(OtherReferenceColumnName,''),isNull(RemarksColumnName,''),isNull(ReferenceNoColumnName,''),"&_
			 " ReferenceSourceTableName,isNull(WhereClauseText,''),isNull(OtherRefNoColumnName,'') from APP_M_RefferenceTypes where ReferenceEntryNo = "& sRefType 
    'Response.Write "<textarea>"& sQuery &"</textarea>"
	rsTemp.Open sQuery,con
	if not rsTemp.EOF then
		sRefCode = trim(rsTemp(0))
		sRefDate = trim(rsTemp(1))
		sRefName = trim(rsTemp(2))
		sOthRef  = trim(rsTemp(3))
		sRemarks = trim(rsTemp(4))
		sRefNo   = trim(rsTemp(5))
		sRefTable= trim(rsTemp(6))
		sRefWhereClause = trim(rsTemp(7))
		sOthRefNoDate = trim(rsTemp(8))
	end if
	rsTemp.Close
	if trim(sRefCode)="" then sRefCode = Chr(39)&Chr(39)
	if trim(sRefDate)="" then sRefDate = Chr(39)&Chr(39)
	if trim(sRefName)="" then sRefName = Chr(39)&Chr(39)
	if trim(sOthRef) ="" then sOthRef  = Chr(39)&Chr(39)
	if trim(sRemarks)="" then sRemarks = Chr(39)&Chr(39)
	if trim(sRefNo)  ="" then sRefNo   = Chr(39)&Chr(39)
	if trim(sOthRefNoDate)="" then sOthRefNoDate = Chr(39)&Chr(39)
	'if trim(sRefWhereClause)="" then sRefWhereClause = Chr(39)&Chr(39)

     if trim(sRefWhereClause)<>"" then
        sQuery  = " Select isNull("& sRefCode &","&sRefNo&") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
					     ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate from "&sRefTable &" where "& sRefWhereClause &" and "&sRefNo&" in (" & sAppRefNo&")"
     else
        sQuery = " Select isNull("& sRefCode &","&sRefNo&") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
					 ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate from "&sRefTable &" where "&sRefNo&" in (" & sAppRefNo&")"
     end if
    'Response.Write "<textarea>"& sQuery &"</textarea>"
     rsTemp.Open sQuery,con
     if not rsTemp.EOF then
        do while not rsTemp.EOF 
            
            sCodes = sCodes &","& trim(rsTemp(0))
            sDate = sDate &","& trim(rsTemp(1))
            sName = sName &","& trim(rsTemp(2))
            sOther = sOther &","& trim(rsTemp(3))
            sRem = sRem &","& trim(rsTemp(4))
            sNo = sNo &","& trim(rsTemp(5))
            sOthNoDate = sOthNoDate &","& trim(rsTemp(6))
            rsTemp.MoveNext 
        loop
     end if
     rsTemp.Close
     
     if trim(sCodes)<>"" then sCodes = mid(sCodes,2)
     if trim(sDate) <>"" then sDate = mid(sDate,2)
     if trim(sName)<>"" then sName = mid(sName,2)
     if Trim(sOther) <>"" then sOther = mid(sOther,2)
     if trim(sRem)<>"" then sRem = mid(sRem,2)
     if trim(sNo)<>"" then sNo = mid(sNo,2)
     if trim(sOthNoDate)<>"" then sOthNoDate = mid(sOthNoDate,2)
     
     
     GetInfoRefType =  sName &":"& sCodes &":"& sDate
    
End Function
%>
<%
Function IssuedToString(sIssuedToType,sIssuedToCode,sIssuedToSubCode)
Dim dcrs,sQuery
set dcrs = Server.CreateObject("ADODB.RecordSet")
    if lcase(trim(sIssuedToType))=lcase("party") then
	    sQuery = "Select PartyName from APP_M_PartyMaster where PartyCode = "& sIssuedToCode &""
	    dcrs.open sQuery,con
	    if not dcrs.eof then
	        IssuedToString = trim(dcrs(0))
	    end if
	    dcrs.close
	elseif lcase(trim(sIssuedToType))=lcase("POS") then
	    sQuery = "Select POSDescription from SAL_M_PointOfSales where POSID = "& sIssuedToCode
        dcrs.Open sQuery,con
        if not dcrs.EOF then
            IssuedToString = "POS - "& trim(dcrs(0)) 
        end if
        dcrs.Close 
    elseif lcase(trim(sIssuedToType)) =lcase("Unit") then
        sQuery = "Select OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OuDefinitionID)>4 and OuDefinitionID in ('"& sIssuedToCode &"')"
        dcrs.Open sQuery,con
        if not dcrs.EOF then
            IssuedToString = "Unit - "& trim(dcrs(0)) 
        end if
        dcrs.Close 
	elseif lcase(trim(sIssuedToType))=lcase("dept") then
	    sQuery = "Select DepartmentName from APP_M_Departments where DeptShortName = '"& sIssuedToCode &"'"
	    dcrs.open sQuery,con
	    if not dcrs.eof then
	        IssuedToString = trim(dcrs(0))
	    end if
	    dcrs.close
	    
	    if trim(sIssuedToSubCode)<>"" then
	        if lcase(trim(sIssuedToCode))=lcase("PRD") then
	            sQuery = "Select WorkCenterName from PRD_M_WORKCENTER where WorkCenterCode = '"& sIssuedToSubCode &"'"
	            dcrs.open sQuery,con
	            if not dcrs.eof then
	                IssuedToString =  IssuedToString &" - "& trim(dcrs(0))
	            end if
	            dcrs.close
	        end if 'if lcase(trim(sIssuedToCode))=lcase("PRD") then
	    end if
	end if
End Function
%>
<%
Function populateIssueToSel(sOrgCode)
    Dim sQuery,rsTemp,objrs
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    set objrs = Server.CreateObject("ADODB.Recordset")
    
	    sQuery = "Select DeptShortName,DepartmentName from APP_M_Departments"
	    rsTemp.Open sQuery,con
	    if not rsTemp.Eof then
		    do while not rsTemp.EOF
			    Response.Write "<option value='Dept:"&  trim(rsTemp(0)) &"'>"&rsTemp(1)&"</option>"
			    rsTemp.MoveNext
		    loop
		end if 
	    rsTemp.Close
	    Response.Write "<option value='Party'>Party</option>"
	    Response.Write "<option value='Unit'>Other Unit</option>"
		    sQuery = "Select OuDefinitionID,OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OuDefinitionID)>4 and OuDefinitionID not in('"& sOrgCode &"')"
		    objrs.Open sQuery,con
		    if not objrs.EOF then
		        do while not objrs.EOF 
		            Response.Write "<option value='Unit:"& trim(objrs(0)) &"'>&nbsp;&nbsp;&nbsp;"&trim(objrs(1))&"</option>"
		            objrs.MoveNext 
		        loop
		    end if
		    objrs.Close 
		Response.Write "<option value='POS'>POS</option>"
		    sQuery = "Select POSID,POSDescription from SAL_M_PointOfSales where OrganisationCode in('"& sOrgCode &"')"
		    objrs.Open sQuery,con
		    if not objrs.EOF then
		        do while not objrs.EOF 
		            Response.Write "<option value='POS:"& trim(objrs(0)) &"'>&nbsp;&nbsp;&nbsp;"&trim(objrs(1))&"</option>"
		            objrs.MoveNext 
		        loop
		    end if
		    objrs.Close 
	
End Function
%>
<%
Function populateIssueToSelWithOutSubLevel(sOrgCode)
    Dim sQuery,rsTemp,objrs
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    set objrs = Server.CreateObject("ADODB.Recordset")
    
	    sQuery = "Select DeptShortName,DepartmentName from APP_M_Departments"
	    rsTemp.Open sQuery,con
	    if not rsTemp.Eof then
		    do while not rsTemp.EOF
			    Response.Write "<option value='Dept:"&  trim(rsTemp(0)) &"'>"&rsTemp(1)&"</option>"
			    rsTemp.MoveNext
		    loop
		end if 
	    rsTemp.Close
	    Response.Write "<option value='Party'>Party</option>"
	    Response.Write "<option value='Unit'>Other Unit</option>"
		Response.Write "<option value='POS'>POS</option>"
End Function
%>

<%
Function GetPartyName(sPartyCode)
    Dim sQuery,rsObj
    set rsObj = Server.CreateObject("ADODB.RecordSet")
    if trim(sPartyCode)<>"" and trim(sPartyCode)<>"0" then
        sQuery = "Select PartyName from APP_M_PartyMaster where PartyCode = "& sPartyCode
        rsObj.open sQuery,con
        if not rsObj.eof then
            GetPartyName = trim(rsObj(0))
        end if
        rsObj.close 
    else
        GetPartyName = ""
    end if
End Function
%>
<%
Function GetDepartmentName(sDeptCode)
    Dim sQuery,rsObj
    set rsObj = Server.CreateObject("ADODB.RecordSet")
    if trim(sDeptCode)<>"" then
        sQuery = "Select DepartmentName from APP_M_Departments where DeptShortName = '"& sDeptCode &"'"
        rsObj.open sQuery,con
        if not rsObj.eof then
            GetDepartmentName = trim(rsObj(0))
        end if
        rsObj.close 
    else
        GetDepartmentName = ""
    end if
End Function
%>
<%
Function GetUnitName(sUnitCode)
    Dim sQuery,rsObj
    set rsObj = Server.CreateObject("ADODB.RecordSet")
    if trim(sUnitCode)<>"" then
        sQuery = "Select OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OuDefinitionID)>4 and OuDefinitionID = '"& sUnitCode &"'"
        rsObj.open sQuery,con
        if not rsObj.eof then
            GetUnitName = "Unit - "& trim(rsObj(0))
        end if
        rsObj.close 
    else
        GetUnitName = ""
    end if
End Function
%>
<%
Function GetUnitDesc(sUnitCode)
    Dim sQuery,rsObj
    set rsObj = Server.CreateObject("ADODB.RecordSet")
    if trim(sUnitCode)<>"" then
        sQuery = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where Len(OuDefinitionID)>4 and OuDefinitionID = '"& sUnitCode &"'"
        rsObj.open sQuery,con
        if not rsObj.eof then
            GetUnitDesc = trim(rsObj(0))
        end if
        rsObj.close 
    else
        GetUnitDesc = ""
    end if
End Function
%>

<%
Function GetPOSName(sPOSID)
    Dim sQuery,rsObj
    set rsObj = Server.CreateObject("ADODB.RecordSet")
    if trim(sPOSID)<>"" then
        sQuery = "Select POSDescription from SAL_M_PointOfSales where POSID = '"& sPOSID &"'"
        rsObj.open sQuery,con
        if not rsObj.eof then
            GetPOSName = "POS - " & trim(rsObj(0))
        end if
        rsObj.close 
    else
        GetPOSName = ""
    end if
End Function
%>
<%
Function GetItemName(ItemCode,ClassCode)
Dim rsObj,sQuery
set rsObj = Server.CreateObject("ADODB.RecordSet")
  sQuery = "Select ItemDescription from VwItem where ItemCode = "& ItemCode &" and ClassificationCode = "& ClassCode
    rsObj.Open sQuery,con
    if not rsObj.EOF then
        GetItemName = trim(rsObj(0))
    end if
    rsObj.Close 
End Function
%>

<%
Function GetClassName(ClassCode)
Dim rsObj,sQuery
set rsObj = Server.CreateObject("ADODB.RecordSet")
  sQuery = "Select GroupName from INV_M_Classification where GroupCode = "& ClassCode
    rsObj.Open sQuery,con
    if not rsObj.EOF then
        GetClassName = trim(rsObj(0))
    end if
    rsObj.Close 
End Function
%>


<%

Function GetCEXReceiptType(sRceptCode)
    Dim rsObj,sQuery
    set rsObj = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select ReceiptType from Cex_M_ReceiptTypes where ReceiptCode = '"& trim(sRceptCode) &"'"
    rsObj.open sQuery,con
    if not rsObj.eof then
        GetCEXReceiptType = trim(rsObj(0)) 
    end if
    rsObj.close
End Function
%>
<%
Function populateIssueTo(sOrgCode,sIssType)
    Dim sQuery,rsTemp,objrs
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    set objrs = Server.CreateObject("ADODB.Recordset")
    
    if trim(sIssType)="GEN" or trim(sIssType)="SER" or trim(sIssType)="JWK" then
	    sQuery = "Select DeptShortName,DepartmentName from APP_M_Departments"
	    rsTemp.Open sQuery,con
	    if not rsTemp.Eof then
		    do while not rsTemp.EOF
			    Response.Write "<option value='Dept:"&  trim(rsTemp(0)) &"'>"&rsTemp(1)&"</option>"
			    rsTemp.MoveNext
		    loop
		end if 
	    rsTemp.Close
	end if 'if trim(sIssType)="GEN" or trim(sIssType)="SER" then
	
	if trim(sIssType)="SUB" or trim(sIssType)="JWK" or trim(sIssType)="SER" then
	    Response.Write "<option value='Party'>Party</option>"
	end if 'if trim(sIssType)="SUB" or trim(sIssType)="JWK" or trim(sIssType)="SER" then
	
	if trim(sIssType)="TRN" then
	    Response.Write "<option value='Unit'>Other Unit</option>"
		    sQuery = "Select OuDefinitionID,OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OuDefinitionID)>4 and OuDefinitionID not in('"& sOrgCode &"')"
		    objrs.Open sQuery,con
		    if not objrs.EOF then
		        do while not objrs.EOF 
		            Response.Write "<option value='Unit:"& trim(objrs(0)) &"'>&nbsp;&nbsp;&nbsp;"&trim(objrs(1))&"</option>"
		            objrs.MoveNext 
		        loop
		    end if
		    objrs.Close 
	end if 'if trim(sIssType)="IUT" then
	
	if trim(sIssType)="POS" then
		    
		Response.Write "<option value='POS'>POS</option>"
		    sQuery = "Select POSID,POSDescription from SAL_M_PointOfSales where OrganisationCode in('"& sOrgCode &"')"
		    objrs.Open sQuery,con
		    if not objrs.EOF then
		        do while not objrs.EOF 
		            Response.Write "<option value='POS:"& trim(objrs(0)) &"'>&nbsp;&nbsp;&nbsp;"&trim(objrs(1))&"</option>"
		            objrs.MoveNext 
		        loop
		    end if
		    objrs.Close 
	end if 'if trim(sIssType)="POS" then
	
End Function
%>
<%
Function GetItemRcptNum(iItemCode)
Dim sQuery,rsObj
set rsObj = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select ReceiptNumbering from VWItem where ItemCode = "& iItemCode
    rsObj.Open sQuery,con
    if not rsObj.EOF then
        GetItemRcptNum = trim(rsObj(0))
    end if
    rsObj.Close 
End Function
%>
<%
Function GetStoreInfo(iItemCode)

Dim objDOM,rsObj
Dim ndRoot,ndLoc
Dim sQuery
Dim iLocNo,iBinNo
set objDOM = Server.CreateObject("Microsoft.XMLDOM")
set rsObj = Server.CreateObject("ADODB.Recordset")

sQuery = "Select A.LocationNumber,LocationName,isNull(BinNumber,0),A.ApplicableFor from Inv_M_ItemStorage A,Inv_M_Storage B "&_
         " where A.LocationNumber = B.LocationNumber and ItemCode = "& iItemCode 
         rsObj.Open sQuery,con
         if not rsObj.EOF then
            iLocNo = rsObj(0)
            iBinNo = rsObj(2)
         else
            iLocNo = ""
            iBinNo = ""
         end if
         rsObj.Close 
         
         GetStoreInfo = iLocNo &":"& iBinNo
         
End Function
%>
<%
Function GetApproverList(sUserAccessMode,sUserType,iAppCode,iProcessCode,iActCode,iActTempNo)
Dim rsObj,sQuery
set rsObj = Server.CreateObject("ADODB.Recordset")

    sQuery =" Select isNull(InternalUserID,0),isNull(UserName,'') from DCS_User where UserAccessMode = '"& sUserAccessMode &"'"&_
            " and InternalUserID<>0 and UserType = '"& sUserType &"' and InternalUserID in (Select InternalUserID from MS_UserActivity"&_
            " where ApplicationCode = "& iAppCode &" and ProcessCode = "& iProcessCode &" and ActivityCode = "& iActCode &" and ActivityTemplateNo = " & iActTempNo & ")"
            'Response.write "<option>"& sQuery &"</option>"
            rsObj.open sQuery,con
            if not rsObj.eof then
                do while not rsObj.eof  
                        Response.write "<option value="&trim(rsObj(0))&">"& trim(rsObj(1)) &"</option>"
                    rsObj.movenext
                loop
            end if
            rsObj.close
End Function
%>

<%
Function GetAccBookName(sBookCode,sBookNumber)
Dim sQuery,rsObj
set rsObj = Server.CreateObject("ADODB.Recordset")

    sQuery = "Select BookName from Acc_R_ApplicableAccountHeads where BookCode = '"& sBookCode &"' and BookNumber = "& sBookNumber
    rsObj.open sQuery,con
    if not rsObj.eof then
        GetAccBookName = rsObj(0)
    end if
    rsObj.close
End Function
%>

<%
Function GetUserInfo(iUserID)
Dim sSQL,rsTemp
Dim sUserName,sUserCode,sUserAccessMode
set rsTemp = Server.CreateObject("ADODB.Recordset")
    sSQL = "Select UserAccessMode,PartyCode,UserName from Dcs_user where InternalUserID = "& iUserID
	rsTemp.open sSQL,con
	if not rsTemp.eof then
	    sUserAccessMode = rsTemp(0)
	    sUserCode = rsTemp(1)
	    sUserName = rsTemp(2)
	end if
	rsTemp.close
	GetUserInfo = sUserAccessMode&":"&sUserCode&":"&sUserName
End Function
%>

<%
Function GetAttName(sAttList)
Dim sSql,rsTemp
Dim sAttName,sTempAttList

sTempAttList = replace(sAttList,":",",")

sSql = "Select OptionName from INV_M_ItemTypeOptions where OptionValue in (" & sTempAttList & ")"
rsTemp.open sSql,con
if not rsTemp.eof then
    do while not rsTemp.eof
        sAttName = sAttName & ","& rsTemp(0)
        rsTemp.movenext
    loop    
end if
rsTemp.close
GetAttName  = mid(sAttName,2)
End Function
%>

<%
Function GetRcptIssName(sRITypeCode)
Dim sQuery,rsObj
set rsObj = Server.CreateObject("ADODB.Recordset")
sQuery = "Select ReceiptIssueTypeDesc from APP_M_ReceiptIssueTYpes where ReceiptIssueTypeCode = '"& sRITypeCode &"'"
rsObj.open sQuery,con
if not rsObj.eof then
    GetRcptIssName = rsObj(0)
end if
rsObj.close
End Function
%>


<%
Function GetClassification(ItemCode)
Dim sQuery,rsObj
set rsObj = Server.CreateObject("ADODB.Recordset")
sQuery = "Select ClassificationCode,GroupName from VWItem where ItemCode = "&ItemCode
rsObj.open sQuery,con
if not rsObj.eof then
    GetClassification = rsObj(0)&":"& rsObj(1)
end if
rsObj.close
End Function
%>


<%
Function GetLastItemCycleCount(sOrgCode,iClass,iItem)
Dim sCycleCountQty,sCycleCountDate,sQuery
Dim rsObj

set rsObj = Server.CreateObject("ADODB.Recordset")
														    
sCycleCountQty = ""
sCycleCountDate = ""

sQuery = "Select IsNull(Convert(varchar,Max(CycleCountDate),103),'dd/MM/yyyy') from Inv_T_ItemCycleCountHistory where OrganisationCode =  '"& sOrgCode &"'  and ClassificationCode= "& iClass &"  and ItemCode = "& iItem 
rsObj.open sQuery,con
if not rsObj.eof then
    sCycleCountDate = rsObj(0)
else
    sCycleCountDate = "dd/MM/yyyy"
end if
rsObj.close
if trim(sCycleCountDate) <> "dd/MM/yyyy" then
    sQuery = "Select CycleCountStock from Inv_T_ItemCycleCountHistory where OrganisationCode =  '"& sOrgCode &"'  and ClassificationCode= "& iClass &"  and ItemCode = "& iItem &" and Convert(datetime,CycleCountDate,103)=Convert(datetime,'"& sCycleCountDate &"',103)"
    rsObj.open sQuery,con
    if not rsObj.eof then
        sCycleCountQty = rsObj(0)
    else
        sCycleCountQty = "0"
    end if
    rsObj.close
else
    sCycleCountQty = "0"
end if 

'Response.write "<p>"& sQuery

GetLastItemCycleCount = sCycleCountQty &":"& sCycleCountDate
End Function

%>
<%
Function GetStockQuality(CategoryID,Value)
Dim sSql,dcrs
set dcrs = Server.CreateObject("ADODB.Recordset")

    sSql ="Select StageID,StageName from Inv_M_Stage where CategoryId = '"& CategoryID &"'"
	dcrs.Open sSql,con
	if not dcrs.EOF then
	    response.write "<option value='S' selected>Select</option>"
		do while not dcrs.EOF 
		    if trim(Value)=trim(dcrs(0)) then
		        response.write "<option value='"& dcrs(0) &"' selected>"& dcrs(1) &"</option>"
		    else
		        response.write "<option value='"& dcrs(0) &"'>"& dcrs(1) &"</option>"
		    end if
		    dcrs.movenext
		loop
	end if 'if not dcrs.EOF then
	dcrs.Close 
End Function
%>
<%
Function DispStockQuality(CategoryID,Value)
Dim sSql,dcrs
set dcrs = Server.CreateObject("ADODB.Recordset")

    sSql ="Select StageID,StageName from Inv_M_Stage where CategoryId = '"& CategoryID &"' and StageID = "& Value
	dcrs.Open sSql,con
	if not dcrs.EOF then
	    DispStockQuality = trim(dcrs(1))
	end if 'if not dcrs.EOF then
	dcrs.Close 
End Function
%>
<%
Function GetRcptToRcvdQty(sOrgId,iItemCode,iClassCode,sRcptNo,sAppRefNo,sAppRefType,sOthRefNo)
Dim rsObj
Dim iOrdQty,iRcvdQty,iBalQty,iCurrRcptQty,sQuery,sShRelReq,iItemEntNo

set rsObj = Server.CreateObject("ADODB.Recordset")

if trim(sRcptNo)="" or IsNull(sRcptNo) then sRcptNo = "0"

if trim(sAppRefType)<>"N" then
    if trim(sAppRefType)="4" or trim(sAppRefType)="20" or trim(sAppRefType)="22" then
        sQuery = "Select QuantityOrdered from PUR_T_PODetails where PurchaseOrderNo = "& sAppRefNo &" and ItemCode = "& iItemCode
    elseif trim(sAppRefType)="23" then
        sQuery = "Select QuantityForReturn from Sal_T_SalesReturnDetail where SalesReturnNo = "& sAppRefNo &" and Itemcode = "& iItemCode
    elseif trim(sAppRefType)="7" then
        sQuery = "Select QuantityReceived from RCV_T_GRNItemDetails where GRNNumber = "& sAppRefNo &" and Itemcode = "&iItemCode
    elseif trim(sAppRefType)="42" then
        sQuery = "Select QuantityOrdered from Sal_T_OrdersDetails where OrderNumber =  "& sAppRefNo &" and ItemCode = "& iItemCode
    elseif trim(sAppRefType)="36" then
	    sQuery = "Select ActionOnQty from VW_Purchase_ActionOnRcptItem V join RCV_T_ActualRcptItemDet D on V.ReceiptNumber = D.ReceiptNumber where V.EntryNo = D.EntryNo and ActionTakenNo = "& sAppRefNo &" and ItemCode = "& iItemCode
	elseif trim(sAppRefType)="12" then
	    sQuery = "Select SUM(QuantityIssued) from INV_T_MaterialIssueDetails where IssueEntryNo = "& sAppRefNo &" and ItemCode ="& iItemCode
	elseif trim(sAppRefType)="21" then
	    sQuery = "Select SUM(QuantityIssued) from INV_T_MaterialIssueDetails where IssueEntryNo in (Select IssueEntryNo from INV_T_MaterialIssueHeader where AppRefType in (21,4) and AppRefNo =  "& sAppRefNo &") and ItemCode ="& iItemCode
    end if
    'Response.write "<p>"& sQuery
    rsObj.open sQuery,con
    if not rsObj.eof then
        iOrdQty = rsObj(0)
    else
        iOrdQty = 0
    end if 
    rsObj.close
else
    iOrdQty = 0
end if

sQuery = "Select IsNull(SUM(QuantityReceived),0) from RCV_T_ActualRcptItemdet where ReceiptNumber in "&_
         " (Select ReceiptNumber from RCV_T_ActualReceiptHeader where AppRefNo ="& sAppRefNo &" and AppRefType = "& sAppRefType &") "&_
         "  and ItemCode = "& iItemCode
         
         if trim(sRcptNo)<>"0" then
            sQuery = sQuery &" and ReceiptNumber not in ("& sRcptNo &")"
         end if
'Response.write "<textarea>"&sQUery&"</textarea>"
rsObj.open sQuery,con
if not rsObj.eof then
    iRcvdQty = rsObj(0)
else
    iRcvdQty = 0
end if 
rsObj.close


if trim(sAppRefType)="4" or trim(sAppRefType)="20" or trim(sAppRefType)="21" or trim(sAppRefType)="22" then
        sQuery = "Select H.ScheduleReleaseReq,D.EntryNumber from PUR_T_POHeader H Join Pur_T_PODetails D on H.PurchaseOrderNo = D.PurchaseOrderNo where H.PurchaseOrderNo = "&  sAppRefNo  &" and D.ItemCode = "& iItemCode
        rsObj.open sQuery,con
        if not rsObj.eof then
            sShRelReq  = rsObj(0)
            iItemEntNo = rsObj(1)
        end if
        rsObj.close
        if trim(sShRelReq)="Y" then
            sQuery = "Select IsNull(SUM(ReleasedQty),0),IsNull(SUM(ReceivedQty),0) from PUR_T_POReleaseHeader H join PUR_T_POReleaseDetail D on H.ReleaseEntryNo = D.ReleaseEntryNO where H.ReferenceOrderNo = "& sAppRefNo &" and D.EntryNo = "& iItemEntNo &" and H.ReleaseEntryNo = "& sOthRefNo
            rsObj.open sQuery,con
            if not rsObj.eof then
                iOrdQty = rsObj(0)
                iRcvdQty = rsObj(1)
            end if
            rsObj.close
        end if
end if 

if trim(sRcptNo)<>"0" and trim(sRcptNo)<>"" then
    sQuery = "Select IsNull(QuantityReceived,0) from RCV_T_ActualRcptItemdet where ReceiptNumber = "& sRcptNo &" and ItemCode = "& iItemCode
    rsObj.open sQuery,con
    if not rsObj.eof then
        iCurrRcptQty = rsObj(0)
    else
        iCurrRcptQty =  0
    end if
    rsObj.close
else
    iCurrRcptQty =  0
end if

if trim(iOrdQty)="" or IsNull(iOrdQty) then iOrdQty = "0"
if trim(iCurrRcptQty) ="" or IsNull(iCurrRcptQty) then iCurrRcptQty = "0"
if trim(iRcvdQty) = "" or IsNull(iRcvdQty) then iRcvdQty = "0"


iBalQty = cdbl(iOrdQty)-cdbl(iCurrRcptQty)-cdbl(iRcvdQty)
GetRcptToRcvdQty = iOrdQty&":"&iRcvdQty&":"&iCurrRcptQty&":"&iBalQty
End Function
%>
<%
Function ShowSendAs(iApp,iPro,iAct,iActTempNo)
Dim rsObj 
Dim bFlagEmail,bFlagSMS,sQuery
set rsObj = Server.CreateObject("ADODB.Recordset")
sQuery = "Select IsNull(SendAsEmail,'N'),IsNull(SendAsSMS,'N') from Ms_ApplicationActivityTemplates WHERE ApplicationCode = "& iApp &" and ProcessCode ="& iPro &" and ActivityCode="& iAct &" and ActivityTemplateNo = "& iActTempNo 
rsObj.open sQuery,con
if not rsObj.eof then
    bFlagEmail  = rsObj(0)
    bFlagSMS = rsObj(1)
end if 
rsObj.close 
Response.Write "<td class='FieldCellSub'>Send As</td>"
Response.Write "<td class='FieldCellSub'>"
if Trim(bFlagEmail)="Y" then
    Response.Write "<input type='checkbox' name='chkEmail' value='Email'  class='FormElem'/>Email"
end if 
if Trim(bFlagSMS)="Y" then
    Response.Write "<input type='checkbox' name='chkSMS' value='SMS'  class='FormElem'/>SMS"
end if 
Response.Write "</td>"
End Function
%>