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
Dim sSql,sActName,iSeries,iSeriesType,iLength,iSeriesCode, sItemType,sClassCode,sArrClassCodes
Dim sTotEntNo,Objfs,objDOM,Root,sExp,TempNode,iCount,iCnt,iCatCode,sArrCateCode

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
sClassCode = Trim(Request.Form("hClassCode"))
iCatCode = Trim(Request.Form("hCatCode"))


sTotEntNo = Request.Form("hTotEntNo")

If sPacking = "1" and sActivity = "P" Then sActivity = "K"
If isNull(sProduct) or sProduct = "" Then
	sProduct = "0"
End If
			
'Response.Write "Manual numbering is :"&sManualNumType
if Trim(sItemType)="" or IsNull(sItemType) then sItemType = "NULL"
if Trim(sItemType)<>"NULL" then sItemType = pack(sItemType)

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


            if Trim(sClassCode)="" or IsNull(sClassCode) then sClassCode="NULL" 
            if Trim(iCatCode) ="" or IsNull(iCatCode) then iCatCode = "NULL"
            
            sArrClassCodes = Split(sClassCode,",")
            sArrCateCode = Split(iCatCode,",")

	
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
				
				
			'	sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,SERIESNO,SERIESCODE, ITEMTYPEID,PackingType,ClassificationCode) VALUES " &_
			'		"('" & iUnitNo & "',"& iSeries & "," & iSeriesCode & ", "&sItemType&",'"& sPacking &"',"&  sClassCode &")"
					
				sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,SERIESNO,SERIESCODE, ITEMTYPEID,PackingType) VALUES " &_
					"('" & iUnitNo & "',"& iSeries & "," & iSeriesCode & ", "&sItemType&",'"& sPacking &"')"
				Response.Write "<p>" & sSql
				con.Execute sSql
				
				if UBound(sArrCateCode)=UBound(sArrClassCodes) then
				    For iCnt = 0 to UBound(sArrClassCodes)
				        sSql = "Insert into PRD_M_PackingNoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iSeriesCode &","& sArrClassCodes(iCnt) &","& sArrCateCode(iCnt) &")"
				        Response.Write "<p>"& sSql
				        con.execute sSql 
				    Next    
			    elseif UBound(sArrCateCode)>0 then
			        For iCnt = 0 to UBound(sArrCateCode)
				        sSql = "Insert into PRD_M_PackingNoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iSeriesCode &","& sArrCateCode(iCnt) &")"
				        Response.Write "<p>"& sSql
				        con.execute sSql 
				    Next    
			    end if

				sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIESCHECK (ORGANISATIONCODE,NUMBERINGTYPE,SERIESNO) " &_
					"VALUES ('" & iUnitNo & "','"&sActivity& "',"& iSeries & ")"
				Response.Write "<p>" & sSql
				con.Execute sSql

				'If sActivity = "I" Then
			'	If sActivity = "C" Then
			'		sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIESITEMTYPE (ORGANISATIONCODE, ITEMTYPEID, PRODUCTWISE) VALUES ('"&iUnitNo&"', "&sItemType&", '"&sProduct&"')"
			'		Response.Write "<p>" & sSql
			'		con.Execute sSql
			'	End If
			ElseIf sActivity = "C" Then
				With dcrs
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT ORGANISATIONCODE FROM PRD_M_PACKINGNUMBERSERIESITEMTYPE WHERE ORGANISATIONCODE = '" & iUnitNo & "' AND ITEMTYPEID = "&sItemType&""
					.Source = "SELECT ORGANISATIONCODE FROM VWPackNoSeries WHERE ORGANISATIONCODE = '" & iUnitNo & "' AND ClassCode in ( "& sClassCode &")"
					Response.Write "<p>"& dcrs.Source
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
					
				'	sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,SERIESNO,SERIESCODE, ITEMTYPEID,PackingType,Classificationcode) VALUES " &_
				'		"('" & iUnitNo & "',"& iSeries & "," & iSeriesCode & ","&sItemType&",'"& sPacking &"',"& sClassCode &")"
						
					sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIES (ORGANISATIONCODE,SERIESNO,SERIESCODE, ITEMTYPEID,PackingType) VALUES " &_
						"('" & iUnitNo & "',"& iSeries & "," & iSeriesCode & ","&sItemType&",'"& sPacking &"')"
					Response.Write "<p>" & sSql
					con.Execute sSql
					
					if UBound(sArrCateCode)=UBound(sArrClassCodes) then
				        For iCnt = 0 to UBound(sArrClassCodes)
				            sSql = "Insert into PRD_M_PackingNoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iSeriesCode &","& sArrClassCodes(iCnt) &","& sArrCateCode(iCnt) &")"
				            Response.Write "<p>"& sSql
				            con.execute sSql 
				        Next    
			        elseif UBound(sArrCateCode)>0 then
			            For iCnt = 0 to UBound(sArrCateCode)
				            sSql = "Insert into PRD_M_PackingNoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iSeriesCode &","& sArrCateCode(iCnt) &")"
				            Response.Write "<p>"& sSql
				            con.execute sSql 
				        Next    
			        end if
					
					'	sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIESITEMTYPE (ORGANISATIONCODE, ITEMTYPEID, PRODUCTWISE) VALUES ('"&iUnitNo&"', "&sItemType&", '"&sProduct&"')"
					 '   Response.Write "<p>" & sSql
					  '  con.Execute sSql


				'	sSql = "INSERT INTO PRD_M_PACKINGNUMBERSERIESITEMTYPE (ORGANISATIONCODE, ITEMTYPEID, PRODUCTWISE,ClassificationCode) VALUES ('"&iUnitNo&"', "&sItemType&", '"&sProduct&"',"& sClassCode &")"
				'	Response.Write "<p>" & sSql
				'	con.Execute sSql
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
			if Trim(sClassCode)="0" then sClassCode = "NULL"
			'update query
			'sSql = "Update PRD_M_PackingNumberSeriesItemType set ProductWise = '"&sProduct&"' where OrganisationCode ='" & iUnitNo & "' and ItemTypeId = " & sItemType & ""
			'sSql = "Update PRD_M_PackingNumberSeriesItemType set ProductWise = '"&sProduct&"' where OrganisationCode ='" & iUnitNo & "' and classificationcode = "& sClassCode
			'Response.Write "<p>" & sSql
			'con.Execute sSql
					
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
   ' con.RollbackTrans
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