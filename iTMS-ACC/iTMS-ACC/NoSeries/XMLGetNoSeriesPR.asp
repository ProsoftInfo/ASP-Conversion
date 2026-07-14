<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetNoSeriesPR.asp
	'Module Name				:	Sales (Master Creation)
	'Author Name				:	Kalai Selvi
	'Created On					:	July 27,2009
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
	Dim sUnit,sActitity,sTemp,sTempVal,sQuery,Objrs,iCtr,sSerCode
	Dim sMon,sYear,sMonYr,sFinYear,sFinFrom,sFinTo,saTemp,sCrValue,sDrValue
	Dim oDom,newElem,Root,sSeriesCode,sSeriesNo,sRetTy,sTotEntNo
	
	sTempVal = Request("sVal")
		
	sTemp = Split(sTempVal,":")
	
	
	
	iCtr = 1
	
	'sUnit = sTemp(0)
	'sActitity = sTemp(1)
	'sSerCode = sTemp(2)
	
	sSeriesNo = sTemp(0)
	sSeriesCode = sTemp(1)
	sUnit = sTemp(2)
	
	
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
	sMon = Mid(saTemp(1),4,2)
	sFinTo = sYear&sMon
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	
	sRetTy = "U"
	
	
		IF CStr(sSeriesNo) <> "" and CStr(sSeriesCode) <> "" Then
													
			sQuery = "Select Period,Number,isNull(Prefix,''),isNull(Suffix,''),EntryNo From APP_R_NoSeriesModuleEntry  "&_
					 "Where SeriesNo = "&sSeriesNo&" And SeriesCode = "&sSeriesCode&" and  "&_
					 "OUDefinitionID = '"&sUnit&"' and Cast(Period as Numeric) >= "&sFinFrom&"  "&_
					 "and Cast(Period as Numeric) <= "&sFinTo&" Order by cast(EntryNo as numeric) desc"
					 
			'Response.Write sQuery
																 
			With Objrs
				.CursorType = 3
				.CursorLocation = 3
				.ActiveConnection = Con
				.Source = sQuery
				.Open
			End With
														
			Set Objrs.ActiveConnection = Nothing
			Do While Not Objrs.EOF
				Set newElem = oDom.createElement("NumSeries")
				newElem.setAttribute "S.No",iCtr
				newElem.setAttribute "Period",Objrs(0)
				newElem.setAttribute "StartNo",Objrs(1)
				newElem.setAttribute "PreFix",Objrs(2)
				newElem.setAttribute "Suffix",Objrs(3)
				newElem.setAttribute "FromFin",sFinFrom
				newElem.setAttribute "ToFin",sFinTo
				newElem.setAttribute "SeriesNo",sSeriesNo
				newElem.setAttribute "SeriesCode",sSeriesCode
				newElem.setAttribute "EntryNo",Objrs(4)
				Root.appendChild newElem
				iCtr = iCtr + 1
				Objrs.MoveNext
			Loop
			Objrs.Close
		End IF
	
		Response.ContentType="text/xml"
		Response.Write oDom.xml											
	'End IF
	
	
%>