<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetNumberSeriesListPR.asp
	'Module Name				:	Production (Master Creation)
	'Author Name				:	Kalai Selvi R
	'Created On					:	July 28,2009
	'Modified On				:
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->

<%
	Dim sUnit,sNumberingType,sTempVal,sQuery,iCtr,sSerCode,sTemp
	Dim sMon,sYear,sMonYr,sFinYear,sFinFrom,sFinTo,saTemp
	Dim iSeriesCode,iSeriesNo,sItemTyDesc
	Dim sItemType,sProductWise,sPackingType,sClassCode,sItemCode,sProductCode
	Dim sItemDesc,sTempVar,sNoUsed,sTempType
	Dim sManual
	
	Dim Objrs,objRs1,Objrs3
	Dim oDom,newElem,Root
	
	sTempVal = Request("sVal")
	sTemp = Split(sTempVal,":")
	
	iCtr = 1
	sUnit = sTemp(0)
	sNumberingType = sTemp(1)
	
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
	
		
	sQuery = "Select isNull(S.ClassificationCode,0),isNull(S.ItemCode,0),isNull(S.ProductCode,''),isNull(S.PackingType,0),isNull(S.SeriesNo,0) as SeriesNo,isNull(S.SeriesCode,0) as SeriesCode,C.NumberingType as NumberingType,isNull(S.ItemTypeId,'') as ItemTypeId from PRD_M_PackingNumberSeriesCheck C,PRD_M_PackingNumberSeries S where C.OrganisationCode = S.OrganisationCode and C.OrganisationCode = '" & sUnit  & "'"
	if trim(sNumberingType) <> "" then
		sQuery = sQuery  & " and C.NumberingType ='" & sNumberingType & "'"
	end if
	
	With Objrs
		.CursorType = 3
		.CursorLocation = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set Objrs.ActiveConnection = Nothing
	Do While Not Objrs.EOF
		iSeriesNo = Objrs("SeriesNo")
		iSeriesCode = Objrs("SeriesCode")
		
		sTempType = Objrs("NumberingType")
		
		sManual = "N"
		sItemType = ""
		sProductWise = "0"
		sPackingType = "0"
		sItemTyDesc = ""
		
		IF CStr(iSeriesNo) = "0" and CStr(iSeriesCode) = "0" Then
			sManual = "Y"
		else
			sQuery = "Select isNull(S.ClassificationCode,0),isNull(S.ItemCode,0),isNull(S.ProductCode,''),isNull(S.PackingType,0),isNull(S.SeriesNo,0),isNull(S.SeriesCode,0),isNull(S.ItemTypeId,'') as ItemTypeId from PRD_M_PackingNumberSeriesCheck C,PRD_M_PackingNumberSeries S where C.OrganisationCode = S.OrganisationCode and C.OrganisationCode = '" & sUnit  & "' and C.NumberingType ='" & sTempType & "'"
			if sTempType = "I" then ' item type wise
				sQuery = sQuery & " and S.ItemTypeId='" & Objrs("ItemTypeId")  & "'"
			end if 	
			'Response.Write "<p> " & sQuery
			With Objrs1
				.CursorType = 3
				.CursorLocation = 3
				.ActiveConnection = Con
				.Source = sQuery
				.Open
			End With
			Set Objrs1.ActiveConnection = Nothing
			IF Not Objrs1.EOF Then
				sClassCode = objRs1(0)
				sItemCode =  objRs1(1)
				sProductCode = objRs1(2)
				sPackingType =  objRs1(3)
				iSeriesNo = objRs1(4)
				iSeriesCode = objRs1(5)	
				sItemType = objRs1(6)
			End IF
			objRs1.Close
			
			if trim(sProductCode) <> "" then
				sProductWise = "Y"
			end if 
			
			if trim(sItemType) <> "" then
				sQuery = "Select ItemTypeName from Inv_M_ItemType where ItemTypeId= '"& sItemType & "'"
				With Objrs1
					.CursorType = 3
					.CursorLocation = 3
					.ActiveConnection = Con
					.Source = sQuery
					.Open
				End With
				Set Objrs1.ActiveConnection = Nothing
				IF Not Objrs1.EOF Then
					sItemTyDesc = objRs1(0)
				End IF
				objRs1.Close
			end if 
		end IF 'IF CStr(iSeriesNo) = "0" and CStr(iSeriesCode) = "0" Then
		
		sNoUsed = CheckNoSerUsed(sUnit,sTempType,sItemType)
		
		 
		Set newElem = oDom.createElement("NumSeriesList")
		newElem.setAttribute "SeriesNo",iSeriesNo
		newElem.setAttribute "SeriesCode",iSeriesCode
		newElem.setAttribute "ManualNumbering",sManual
		newElem.setAttribute "ItemType",sItemType
		newElem.setAttribute "ProductWise",sProductWise
		newElem.setAttribute "PackingType",sPackingType
		newElem.setAttribute "ItemDesc",sItemDesc
		newElem.setAttribute "EditCheck","N"
		newElem.setAttribute "NoUsed",sNoUsed
		newElem.setAttribute "NumberingType",sTempType
		newElem.setAttribute "ItemTypeDesc",sItemTyDesc
		newElem.setAttribute "ClassCode",sClassCode
		Root.appendChild newElem
		iCtr = iCtr + 1
		Objrs.MoveNext
		
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
		End IF
		
		
		IF CStr(sCallTy) = "I" Then
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
		End IF
		
		sFullVal = Mid(sFullVal,2)
		'Response.Write sFullVal &"<br><br>"
		GetDesc = sFullVal
	End Function
	

	
	Function CheckNoSerUsed(sUnitid,sActionType,sItem)
		Dim Salrs,sCheckVal,sRetValue
		
		
		'======================== For  ================================================
		sQuery = "Select Top 1 DailyPackingCode from PRD_T_DailyPackingHeader where OrganisationCode = '" & sUnitid & "'"

		With Objrs3
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		IF Not Objrs3.EOF Then
			sRetValue = "Y"
		Else
			sRetValue = "N"
		End IF
		Objrs3.Close
		
		'**********************************************************************************************
		CheckNoSerUsed = sRetValue
	End Function
%>