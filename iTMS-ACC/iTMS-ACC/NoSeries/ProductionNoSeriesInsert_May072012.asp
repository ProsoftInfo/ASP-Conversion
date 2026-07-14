<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ProductionNoSeriesInsert.asp
	'Module Name				:	Production (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 25,2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	
	'							:
	'Connects To				:	ProductionNoSeriesEntry.asp
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
<!--#include virtual="/include/NoSeries.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<%

Dim iUnitNo,sActivity,iCounter,sPacking,dcrs,dcrs1, sProduct, sManualNumType
Dim sSql,sActName,iSeries,iSeriesType,iLength,iSeriesCode, sItemType
Dim sTotEntNo,Objfs,objDOM,Root,sExp,TempNode,iCount

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

iUnitNo = Trim(Request.Form("selUnit"))
sActivity = Trim(Request.Form("selNumType"))
sItemType = Trim(Request.Form("selItemType"))
sProduct = Trim(Request.Form("chkProductWise"))
sPacking = Trim(Request.Form("chkPacking"))
sActName = Trim(Request.Form("hActivityName"))
iSeries = trim(Request.Form("selNoSeries"))
iSeriesType = Trim(Request.Form("hSeriesType"))
iLength = Trim(Request.Form("hSeriesLen"))
sManualNumType = Trim(Request.Form("radManual"))

sTotEntNo = Request.Form("hTotEntNo")

If sPacking = "1" and sActivity = "P" Then sActivity = "K"
If isNull(sProduct) or sProduct = "" Then
	sProduct = "0"
End If
			
'Response.Write "Manual numbering is :"&sManualNumType

con.BeginTrans

'edit case
'If objfs.FileExists(Server.MapPath("../Temp/master/NoSeries_PR_"&Session.SessionID&".xml")) then
'	objDOM.Load server.MapPath("../Temp/master/NoSeries_PR_"&Session.SessionID&".xml")
If objfs.FileExists(Server.MapPath("Temp/master/NoSeries_PR_"&Session.SessionID&".xml")) then
	objDOM.Load server.MapPath("Temp/master/NoSeries_PR_"&Session.SessionID&".xml")
	
	Set Root = objDOM.documentElement
	
	sExp = "//NumSeriesList[@EditCheck!=""N""]"
	
	Set TempNode = Root.selectNodes(sExp)
	
	For iCount = 0 To TempNode.length - 1
		IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("EditCheck").value) = "E" Then
		

			Set dcrs = Server.CreateObject("ADODB.RecordSet")
			Set dcrs1 = Server.CreateObject("ADODB.RecordSet")


	
			With dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ORGANISATIONCODE FROM PRD_M_PACKINGNUMBERSERIESCHECK WHERE ORGANISATIONCODE = '" & iUnitNo & "'"
				.ActiveConnection = con
				.Open
			End With
			Response.Write "<p>" & dcrs1.Source
			If dcrs1.EOF Then
				If sManualNumType = "N" Then
					iSeriesCode = GenSeriesCode(iUnitNo,"4","4",iSeries,iSeriesType,"",sActName,iLength)
				Else
					iSeriesCode = 0
					iSeries = 0
				End If
				
				
				sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,SERIESNO,SERIESCODE, ITEMTYPEID,PackingType) VALUES " &_
					"('" & iUnitNo & "',"& iSeries & "," & iSeriesCode & ", '"&sItemType&"','"& sPacking &"')"
				Response.Write "<p>" & sSql
				con.Execute sSql

				sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIESCHECK (ORGANISATIONCODE,NUMBERINGTYPE,SERIESNO) " &_
					"VALUES ('" & iUnitNo & "','"&sActivity& "',"& iSeries & ")"
				Response.Write "<p>" & sSql
				con.Execute sSql

				If sActivity = "I" Then
					sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIESITEMTYPE (ORGANISATIONCODE, ITEMTYPEID, PRODUCTWISE) VALUES ('"&iUnitNo&"', '"&sItemType&"', '"&sProduct&"')"
					Response.Write "<p>" & sSql
					con.Execute sSql
				End If
			ElseIf sActivity = "I" Then
				With dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ORGANISATIONCODE FROM PRD_M_PACKINGNUMBERSERIESITEMTYPE WHERE ORGANISATIONCODE = '" & iUnitNo & "' AND ITEMTYPEID = '"&sItemType&"'"
					.ActiveConnection = con
					.Open
				End With

				If dcrs.EOF Then
					If sManualNumType = "N" Then
						iSeriesCode = GenSeriesCode(iUnitNo,"4","4",iSeries,iSeriesType,"",sActName,iLength)
					Else
						iSeries = 0
						iSeriesCode = 0
					End If
					
					sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,SERIESNO,SERIESCODE, ITEMTYPEID,PackingType) VALUES " &_
						"('" & iUnitNo & "',"& iSeries & "," & iSeriesCode & ",'"&sItemType&"','"& sPacking &"')"
					Response.Write "<p>" & sSql
					con.Execute sSql

					sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIESITEMTYPE (ORGANISATIONCODE, ITEMTYPEID, PRODUCTWISE) VALUES ('"&iUnitNo&"', '"&sItemType&"', '"&sProduct&"')"
					Response.Write "<p>" & sSql
					con.Execute sSql
				End If
				dcrs.Close		
			End If
			dcrs1.Close
	
		ElseIF CStr(TempNode.Item(iCount).Attributes.getNamedItem("EditCheck").value) = "Y" Then 
			Response.Write "Inside Amendment " &"<br>"
			
			iSeries  = TempNode.Item(iCount).Attributes.getNamedItem("SeriesNo").Value
			iSeriesCode = TempNode.Item(iCount).Attributes.getNamedItem("SeriesCode").Value
			
			UpdateNoSerValue iSeries,iSeriesCode,sTotEntNo,TempNode.Item(iCount)
			
			'update query
			sSql = "Update PRD_M_PackingNumberSeries set PackingType = '"& sPacking &"' where SeriesNo ='" & iSeries & "' and SeriesCode = '" & iSeriesCode & "'"
			Response.Write "<p>" & sSql
			con.Execute sSql
			
			'update query
			sSql = "Update PRD_M_PackingNumberSeriesItemType set ProductWise = '"&sProduct&"' where OrganisationCode ='" & iUnitNo & "' and ItemTypeId = '" & sItemType & "'"
			Response.Write "<p>" & sSql
			con.Execute sSql
					
		End if 	
	Next
	
	
	

end if 'If objfs.FileExists()

If con.Errors.count <> 0 Then
	Dim iErrCounter
	con.RollbackTrans
	For iErrCounter = 0 to con.Errors.count - 1
		Response.Write con.Errors(iErrCounter) & vbCrLf
	Next
	'Redirect to Error Handling System
Else
'	con.RollbackTrans
'	Response.End
	
	Response.Clear
	con.CommitTrans
	
End If
con.close
Set con = Nothing
Response.Redirect "ProductionNoSeriesEntry.asp"
%>


<%
	Function UpdateNoSerValue(iSerNo,iSerCode,iLoop,sRoot)
		Dim sStartNo,sPreVal,sSufVal,iCtloop,sStr,NoSerNode,sQuery
		sStr = "//NumSeriesList[@SeriesNo="&iSerNo&" and @SeriesCode="&iSerCode&"]/NoList"
		Set NoSerNode = sRoot.selectNodes(sStr)
		IF NoSerNode.length <> 0 Then
			For iCtloop = 0 To NoSerNode.length - 1
				
				sStartNo = NoSerNode.Item(iCtloop).Attributes.Item(0).value
				sPreVal = NoSerNode.Item(iCtloop).Attributes.Item(1).value
				sSufVal = NoSerNode.Item(iCtloop).Attributes.Item(2).value
				
				
				sQuery = "UPDATE APP_R_NoSeriesModuleEntry SET  Number = "&sStartNo&", Prefix = '"&sPreVal&"', Suffix = '"&sSufVal&"' "&_
						 "WHERE OUDefinitionID = '"&iUnitNo&"' AND SeriesNo = "&iSerNo&" AND SeriesCode = "&iSerCode&" AND EntryNo = "&iCtloop+1&" "
						 
				Response.Write sQuery &"<br><br><br>"
						 
				Con.Execute sQuery	
			Next
		End IF
	End Function
	
%>